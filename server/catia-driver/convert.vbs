' CATIA V5 automation driver for STEP export (VBScript) -- multi-root edition
'
' Usage:
'   cscript //nologo convert.vbs <workDir> <outDir> <timeoutSec> <products...> -- <parts...> [-- <forcedRoots...>]
'   - workDir : clean ASCII dir holding ALL uploaded files (flat). CATIA opens from here.
'   - outDir  : dir where .stp files are written (one per root product). Must be ASCII.
'   - products/parts : file names (basenames) to make available for reference resolution.
'   - forcedRoots : if given (non-empty), export exactly these products instead of auto-detecting.
'
' Behavior:
'   1) Connect CATIA (retry).
'   2) For each product, open it and collect referenced document names (recursively).
'   3) A product is a "root" if it is NOT referenced by any other product.
'      -> roots export independently; nested sub-products are skipped (already inside a root's tree).
'   4) Export each root to "<outDir>\<basename>.stp" via doc.ExportData(..., "stp").
'   5) Echo "ROOTS|root1|root2|..." so the Node wrapper can pair results.
' Exit codes: 0 = at least one root exported; 1 = failure.
'
' WHY VBS not PowerShell: on this machine the .NET COM interop layer
' (New-Object / Activator) returns a dead RCW for CATIA.Application
' (properties null, setters E_FAIL), while VBScript IDispatch late
' binding works. (Same conclusion as the catia-pdf-converter project.)
'
' WHY clean path: CATIA V5 cannot open files whose path contains unusual
' Unicode characters (e.g. U+2011 non-breaking hyphen). The Node wrapper
' copies the whole batch into a clean ASCII temp dir and opens products there.

Option Explicit

Dim fso, catia, deadline, timeout, workDir, outDir
Dim products, parts, forcedRoots, group, k, arg
Dim p, refs, referenced, roots, root, doc, base, file, folder
Dim anyOk, rl, vis

Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count < 3 Then
  WScript.StdErr.WriteLine "usage: convert.vbs <workDir> <outDir> <timeoutSec> <products...> -- <parts...> [-- <forcedRoots...>]"
  WScript.Quit 1
End If

workDir  = WScript.Arguments(0)
outDir   = WScript.Arguments(1)
timeout  = 300
If WScript.Arguments.Count >= 4 Then timeout = CInt(WScript.Arguments(2))

Set products = CreateObject("System.Collections.ArrayList")
Set parts = CreateObject("System.Collections.ArrayList")
Set forcedRoots = CreateObject("System.Collections.ArrayList")
group = "products"   ' products -> parts -> roots

For k = 3 To WScript.Arguments.Count - 1
  arg = WScript.Arguments(k)
  If arg = "--" Then
    If group = "products" Then group = "parts"
    ElseIf group = "parts" Then group = "roots"
    Else group = "done"
  Else
    If group = "products" Then products.Add arg
    ElseIf group = "parts" Then parts.Add arg
    ElseIf group = "roots" Then forcedRoots.Add arg
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

' --- create CATIA with retry ----------------------------------------------
Dim attempt
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

' --- wait until the automation server is ready -----------------------------
Dim ready, n
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

' --- suppress popups --------------------------------------------------------
On Error Resume Next
catia.DisplayFileAlerts = False
catia.Visible = True
On Error GoTo 0

' --- scan references of each product ---------------------------------------
Set refs = CreateObject("Scripting.Dictionary")   ' name -> Dictionary(referenced names)
For Each p In products
  Dim pf: pf = workDir & "\" & p
  If Not fso.FileExists(pf) Then
    WScript.Echo "product not found (skip scan): " & p
  Else
    On Error Resume Next
    Set doc = catia.Documents.Open(pf)
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

' --- compute roots ---------------------------------------------------------
Set referenced = CreateObject("Scripting.Dictionary")
For Each p In products
  If refs.Exists(p) Then
    For Each r In refs(p).Keys
      referenced(r) = True
    Next
  End If
Next

Set roots = CreateObject("System.Collections.ArrayList")
If forcedRoots.Count > 0 Then
  For Each root In forcedRoots
    If products.IndexOf(root) >= 0 Then
      roots.Add root
    Else
      WScript.Echo "forced root not in products (skip): " & root
    End If
  Next
  If roots.Count = 0 Then
    WScript.StdErr.WriteLine "no valid forced roots provided"
    On Error Resume Next: catia.Quit: On Error GoTo 0
    WScript.Quit 1
  End If
Else
  For Each p In products
    If Not referenced.Exists(p) Then roots.Add p
  Next
  ' 防御：一个都没被判定为根（极端情况），退化为全部产品都是根
  If roots.Count = 0 And products.Count > 0 Then
    For Each p In products: roots.Add p: Next
  End If
End If

If roots.Count = 0 Then
  WScript.StdErr.WriteLine "no products to export"
  On Error Resume Next: catia.Quit: On Error GoTo 0
  WScript.Quit 1
End If

' --- export each root -------------------------------------------------------
anyOk = False
For Each root In roots
  Dim rf: rf = workDir & "\" & root
  If Not fso.FileExists(rf) Then
    WScript.Echo "root not found (skip export): " & root
  Else
    On Error Resume Next
    Set doc = catia.Documents.Open(rf)
    If Err.Number = 0 Then
      WScript.Sleep 4000   ' allow large assembly to finish loading
      base = outDir & "\" & fso.GetBaseName(root)
      doc.ExportData base & ".stp", "stp"
      doc.Close True
      anyOk = True
      WScript.Echo "exported: " & fso.GetBaseName(root) & ".stp"
    Else
      WScript.Echo "open failed (skip export): " & root & " : " & Err.Description
      Err.Clear
    End If
    On Error GoTo 0
  End If
Next

' --- cleanup CATIA ----------------------------------------------------------
On Error Resume Next
catia.Quit
On Error GoTo 0

If Not anyOk Then
  WScript.StdErr.WriteLine "no root exported"
  WScript.Quit 1
End If

' --- report roots -----------------------------------------------------------
rl = "ROOTS|"
For Each root In roots: rl = rl & root & "|": Next
WScript.Echo rl
WScript.Echo "OK"
WScript.Quit 0

' --- helper: collect referenced document names recursively ------------------
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
