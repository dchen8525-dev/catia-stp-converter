' CATIA V5 STEP exporter driver (VBScript) -- multi-root edition, batch exporter
'
' Usage:
'   cscript //nologo convert.vbs <workDir> <outDir> <timeoutSec> <products...> -- <parts...> [-- <forcedRoots...>]
'   - workDir : clean ASCII dir holding ALL uploaded files (flat). CATIA opens from here.
'   - outDir  : dir where .stp files are written (one per root). Must be ASCII.
'   - products/parts : file names (basenames) to make available for reference resolution.
'   - forcedRoots : if given (non-empty), export exactly these products/parts instead of auto-detecting.
'
' Behavior:
'   1) Only when needed (no forced roots AND products exist), boot CATIA to SCAN
'      product references and decide which products are "roots" (not referenced
'      by any other product). Parts-only batches skip this entirely.
'   2) Export each root with the official headless converter:
'        CATSTART -run "CATBatchStarter -input <batch.xml>"
'      one single-file batch XML per root, output is "<outDir>\<basename>.stp".
'   3) Echo "ROOTS|root1|root2|..." (only the roots that actually exported) so
'      the Node wrapper can pair results.
' Exit codes: 0 = at least one root exported; 1 = failure.
'
' WHY batch exporter instead of doc.ExportData(..., "stp")?
'   On this machine the in-session STEP export API persistently returns
'   E_FAIL (0x80004005) - even for an empty part - while IGES export works and
'   interactive GUI export works. The diagnostic trail: same install, same
'   license (HD2+ST1), same files, GUI OK / API fails. The official headless
'   "Batch DXF-IGES-STEP" utility (BatchDataExchange) goes through a different
'   C++ conversion path and produces valid AP242 STEP files. So we drive that.
'   A batch XML accepts exactly ONE input file (multi-file runs are rejected
'   with RC=208), hence one separate batch invocation per root.
'
' WHY VBS not PowerShell: on this machine the .NET COM interop layer
' (New-Object / Activator) returns a dead RCW for CATIA.Application
' (properties null, setters E_FAIL), while VBScript IDispatch late binding
' works. (Same conclusion as the catia-pdf-converter project.)
'
' WHY clean path: CATIA V5 cannot open files whose path contains unusual
' Unicode characters (e.g. U+2011 non-breaking hyphen). The Node wrapper
' copies the whole batch into a clean ASCII temp dir and opens products there.
'
' Machine limits (also see the repo memory note):
'   - cscript reads scripts as ANSI: this file must stay pure ASCII.
'   - No .NET COM objects (System.Collections.ArrayList etc.) - activation
'     fails with 0x80131700; use Scripting.Dictionary only.

Option Explicit

Dim fso, sh, catia, deadline, timeout, workDir, outDir
Dim products, parts, forcedRoots, group, k, arg
Dim p, r, refs, referenced, roots, root, doc, vis
Dim anyOk, rl, needScan, rf, rc

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

If WScript.Arguments.Count < 3 Then
  WScript.StdErr.WriteLine "usage: convert.vbs <workDir> <outDir> <timeoutSec> <products...> -- <parts...> [-- <forcedRoots...>]"
  WScript.Quit 1
End If

workDir = WScript.Arguments(0)
outDir  = WScript.Arguments(1)
timeout = 300
If WScript.Arguments.Count >= 4 Then timeout = CInt(WScript.Arguments(2))

' --- parse argument groups ------------------------------------------------
Set products    = CreateObject("Scripting.Dictionary")
Set parts       = CreateObject("Scripting.Dictionary")
Set forcedRoots = CreateObject("Scripting.Dictionary")
products.CompareMode    = vbTextCompare
parts.CompareMode       = vbTextCompare
forcedRoots.CompareMode = vbTextCompare
group = "products"   ' products -> parts -> roots

For k = 3 To WScript.Arguments.Count - 1
  arg = WScript.Arguments(k)
  ' single-line "If .. Then x" does not support ElseIf; use block form
  If arg = "--" Then
    If group = "products" Then
      group = "parts"
    ElseIf group = "parts" Then
      group = "roots"
    Else
      group = "done"
    End If
  Else
    If group = "products" Then
      If Not products.Exists(arg) Then products.Add arg, True
    ElseIf group = "parts" Then
      If Not parts.Exists(arg) Then parts.Add arg, True
    ElseIf group = "roots" Then
      If Not forcedRoots.Exists(arg) Then forcedRoots.Add arg, True
    End If
  End If
Next

If Not fso.FolderExists(workDir) Then
  WScript.StdErr.WriteLine "workDir not found: " & workDir
  WScript.Quit 1
End If
If Not fso.FolderExists(outDir) Then
  On Error Resume Next
  fso.CreateFolder(outDir)
  On Error GoTo 0
End If

deadline = DateAdd("s", timeout, Now())

' --- decide whether a CATIA session is needed -------------------------------
' A session is only used to SCAN product references and deduce root products.
' Parts-only batches and explicit forced roots do not need one.
needScan = (forcedRoots.Count = 0) And (products.Count > 0)

' --- locate the batch exporter ----------------------------------------------
' env name / dirEnv come from disk; CATInstallPath is captured from the CATIA
' session when we have one, else a machine-specific fallback.
Dim dirEnv, envName, installPath, catStart, f
dirEnv = sh.ExpandEnvironmentStrings("%ProgramData%") & "\DassaultSystemes\CATEnv"
envName = ""
If fso.FolderExists(dirEnv) Then
  For Each f In fso.GetFolder(dirEnv).Files
    ' NOTE: the Like operator is broken on this machine's VBScript ("Undefined
    ' Sub or Function") - work around with InStr/Left.
    If (Left(LCase(f.Name), 5) = "catia") And (InStr(1, LCase(f.Name), ".txt") > 0) Then
      envName = fso.GetBaseName(f.Name)
      Exit For
    End If
  Next
End If
If Len(envName) = 0 Then envName = "CATIA.V5-6R2021.B31"
WScript.Echo "env=" & envName & " dirEnv=" & dirEnv

installPath = ""
' NOTE: single-line If .. Then may NOT contain a procedure call; use block form
If needScan Then
  installPath = ConnectAndScan   ' also fills module-level "roots"
End If
If Len(installPath) = 0 Then installPath = "C:\Program Files\Dassault Systemes\B31\win_b64"
catStart = installPath & "\code\bin\CATSTART.exe"
If Not fso.FileExists(catStart) Then
  WScript.StdErr.WriteLine "CATSTART.exe not found: " & catStart
  WScript.Quit 1
End If

' --- compute roots -----------------------------------------------------------
Set roots = CreateObject("Scripting.Dictionary")
roots.CompareMode = vbTextCompare
If forcedRoots.Count > 0 Then
  For Each root In forcedRoots
    If products.Exists(root) Or parts.Exists(root) Then
      roots.Add root, True
    Else
      WScript.Echo "forced root not in products/parts (skip): " & root
    End If
  Next
  If roots.Count = 0 Then
    WScript.StdErr.WriteLine "no valid forced roots provided"
    WScript.Quit 1
  End If
ElseIf products.Count > 0 Then
  ' (roots were filled by ConnectAndScan)
  If roots.Count = 0 Then
    ' Fallback: if nothing was detected as a root (edge case), treat every product as a root
    For Each p In products
      roots.Add p, True
    Next
  End If
ElseIf parts.Count > 0 Then
  ' Parts-only batch (no CATProduct): each part is its own root
  For Each p In parts
    roots.Add p, True
  Next
End If

If roots.Count = 0 Then
  WScript.StdErr.WriteLine "nothing to export (no products or parts)"
  WScript.Quit 1
End If

' --- release any scan session / leftover CNEXT ------------------------------
If needScan Then
  ReleaseSession
End If

' --- export each root via headless batch -------------------------------------
' one single-file "Batch DXF-IGES-STEP" XML per root, run through CATBatchStarter.
Dim okRoots, outs, outp, cmd, xmlPath, dst
Set okRoots = CreateObject("Scripting.Dictionary")
okRoots.CompareMode = vbTextCompare
outs = outDir
If Right(outs, 1) <> "\" Then outs = outs & "\"

For Each root In roots
  rf = workDir & "\" & root
  If Not fso.FileExists(rf) Then
    WScript.Echo "root not found (skip export): " & root
  Else
    xmlPath = outs & fso.GetBaseName(root) & "-batch.xml"
    dst     = outs & fso.GetBaseName(root) & ".stp"
    If WriteBatchXml(xmlPath, rf, outs) Then
      cmd = """" & catStart & """ -run ""CATBatchStarter -input " & xmlPath & """ -env " & envName & " -direnv """ & dirEnv & """"
      WScript.Echo "batch start: " & root
      If RunBatchAndWait(cmd, dst) Then
        okRoots.Add root, True
        WScript.Echo "exported: " & fso.GetBaseName(root) & ".stp"
      Else
        WScript.Echo "batch failed: " & root
      End If
    Else
      WScript.Echo "could not write batch xml (skip): " & root
    End If
    On Error Resume Next
    fso.DeleteFile xmlPath
    On Error GoTo 0
  End If
  If Now() > deadline Then
    WScript.Echo "deadline passed, aborting remaining roots"
    Exit For
  End If
Next

anyOk = okRoots.Count > 0
If Not anyOk Then
  WScript.StdErr.WriteLine "no root exported"
  WScript.Quit 1
End If

' --- report roots (only the successfully exported ones) ----------------------
rl = "ROOTS|"
For Each root In okRoots
  rl = rl & root & "|"
Next
WScript.Echo rl
WScript.Echo "OK"
WScript.Quit 0

' ============================================================================
'   Helpers
' ============================================================================

' Connect CATIA (retry), scan every product's references, fill module-level
' "roots" with the top-level (unreferenced) products, return CATInstallPath.
Function ConnectAndScan()
  Dim attempt, ready, n
  ConnectAndScan = ""
  Set catia = Nothing
  For attempt = 1 To 5
    On Error Resume Next
    Set catia = CreateObject("CATIA.Application")
    If Err.Number = 0 Then Exit For
    On Error GoTo 0
    WScript.Echo "create attempt " & attempt & " failed"
    If attempt < 5 Then WScript.Sleep 3000
  Next
  On Error GoTo 0
  If catia Is Nothing Then
    WScript.StdErr.WriteLine "could not create CATIA.Application (5 attempts)"
    WScript.Quit 1
  End If
  WScript.Echo "CATIA.Application created"

  ready = False
  Do While Not ready
    If Now() > deadline Then Exit Do
    On Error Resume Next
    n = catia.Documents.Count
    If Err.Number = 0 Then
      ready = True
    Else
      Err.Clear
      WScript.Sleep 2000
    End If
    On Error GoTo 0
  Loop
  If Not ready Then
    WScript.StdErr.WriteLine "CATIA did not become ready within " & timeout & "s"
    On Error Resume Next: catia.Quit: On Error GoTo 0
    WScript.Quit 1
  End If
  WScript.Echo "CATIA ready"

  ' suppress popups + grab install path from the live session
  On Error Resume Next
  catia.DisplayFileAlerts = False
  catia.Visible = True
  ConnectAndScan = catia.SystemService.Environ("CATInstallPath")
  On Error GoTo 0

  ' --- scan references of each product ----------------------------------
  Set refs = CreateObject("Scripting.Dictionary")
  For Each p In products
    rf = workDir & "\" & p
    If Not fso.FileExists(rf) Then
      WScript.Echo "product not found (skip scan): " & p
    Else
      On Error Resume Next
      Set doc = catia.Documents.Open(rf)
      If Err.Number = 0 Then
        WScript.Sleep 3000   ' allow load
        Set refs(p) = CreateObject("Scripting.Dictionary")
        Set vis = CreateObject("Scripting.Dictionary")
        CollectRefs doc.Product, refs(p), vis
        doc.Close True
        WScript.Echo "scanned: " & p
      Else
        WScript.Echo "open failed (skip scan): " & p & " : " & Err.Description
        Err.Clear
      End If
      On Error GoTo 0
    End If
  Next

  ' --- roots = products never referenced by another product --------------
  Set referenced = CreateObject("Scripting.Dictionary")
  For Each p In products
    If refs.Exists(p) Then
      For Each r In refs(p).Keys
        referenced(r) = True
      Next
    End If
  Next
  For Each p In products
    If Not referenced.Exists(p) Then roots.Add p, True
  Next
End Function

' Quit the scan session and make sure no CNEXT lingers before batch phase.
Sub ReleaseSession()
  On Error Resume Next
  catia.Quit
  On Error GoTo 0
  WScript.Sleep 3000
  On Error Resume Next
  sh.Run "taskkill /IM CNEXT.exe /F", 0, True
  On Error GoTo 0
  WScript.Sleep 2000
End Sub

' Write a single input-file batch XML for the BatchDataExchange utility.
Function WriteBatchXml(xmlPath, srcFile, outFolder)
  Dim ts, s
  WriteBatchXml = False
  On Error Resume Next
  Set ts = fso.CreateTextFile(xmlPath, True)
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  s = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf _
    & "<!DOCTYPE root SYSTEM ""Parameters.dtd"">" & vbCrLf _
    & "<root batch_name=""BatchDataExchange"" user="""" password="""" env="""" version="""" licfile="""">" & vbCrLf _
    & "<inputParameters>" & vbCrLf _
    & "<file id=""FileToProcess"" destination="""" filePath=""" & srcFile & """ type=""bin"" upLoadable=""RightNow"" automatic=""1""/>" & vbCrLf _
    & "</inputParameters>" & vbCrLf _
    & "<outputParameters>" & vbCrLf _
    & "<folder id=""OutputFolder"" destination=""" & outFolder & """ folderPath=""" & outFolder & """ type=""bin"" upLoadable=""RightNow"" extension=""*"" automatic=""1""/>" & vbCrLf _
    & "<simple_arg id=""OutputExtension1"" value=""stp""/>" & vbCrLf _
    & "</outputParameters>" & vbCrLf _
    & "<PCList>" & vbCrLf _
    & "<PC name=""HD2.slt"" />" & vbCrLf _
    & "<PC name=""ST1.prd"" />" & vbCrLf _
    & "</PCList>" & vbCrLf _
    & "</root>"
  ts.Write s
  ts.Close
  WriteBatchXml = True
  On Error GoTo 0
End Function

' Wait (inside the global deadline) until the STEP file is finished, the batch
' exits, or the deadline passes.
' NOTE: do NOT rely on the wrapper process (CATSTART.exe) exiting - it can hang
' open holding the stdout pipe long after the .stp is complete, and an
' unconditional stdout drain would block forever. Completion is detected by the
' output file's size stabilizing, not by the process.
Function RunBatchAndWait(cmd, dst)
  Dim exec, ok
  ok = False
  On Error Resume Next
  Set exec = sh.Exec(cmd)
  If Err.Number <> 0 Then
    WScript.Echo "exec failed: " & Err.Description
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  Do
    If FileDone(dst) Then
      ok = True
      Exit Do
    End If
    If exec.Status <> 0 Then
      ' the batch reported a completion/failure; give it a moment, then check
      WScript.Sleep 2000
      ok = FileDone(dst)
      Exit Do
    End If
    If Now() > deadline Then
      WScript.Echo "batch deadline - killing CNEXT"
      On Error Resume Next
      sh.Run "taskkill /IM CNEXT.exe /F", 0, True
      On Error GoTo 0
      WScript.Sleep 3000
      ok = FileDone(dst)   ' it may still have finished while we waited
      Exit Do
    End If
    WScript.Sleep 1000
  Loop

  ' cleanup: kill any leftover CNEXT so the license is free for the next root
  On Error Resume Next
  sh.Run "taskkill /IM CNEXT.exe /F", 0, True
  On Error GoTo 0
  RunBatchAndWait = ok
End Function

' True once the output file exists, is non-trivial, and its size is stable
' across a short interval (i.e. the converter finished flushing it).
Function FileDone(dst)
  Dim s1, s2, i
  FileDone = False
  If Not fso.FileExists(dst) Then Exit Function
  For i = 1 To 8
    s1 = fso.GetFile(dst).Size
    WScript.Sleep 5000
    If Not fso.FileExists(dst) Then Exit Function
    s2 = fso.GetFile(dst).Size
    If s2 > 1024 And s2 = s1 Then
      FileDone = True
      Exit Function
    End If
  Next
End Function

' collect referenced document names recursively (only used for products)
Sub CollectRefs(prod, dict, visited)
  Dim i, kid, refProd, refDoc, dn
  For i = 1 To prod.Products.Count
    On Error Resume Next
    Set kid = prod.Products.Item(i)
    Set refProd = kid.ReferenceProduct
    Set refDoc = refProd.Parent
    dn = refDoc.Name
    If Err.Number = 0 Then
      dict(dn) = True
      If Not visited.Exists(dn) Then
        visited(dn) = True
        CollectRefs refProd, dict, visited
      End If
    Else
      Err.Clear
    End If
    On Error GoTo 0
  Next
End Sub