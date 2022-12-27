Option Explicit

Dim f, gf, n, pp, pt, so

Set so = CreateObject("Scripting.FileSystemObject")
Set gf = so.GetFolder(so.GetParentFolderName(WScript.ScriptFullName))
Set pp = CreateObject("PowerPoint.Application")

For Each f In gf.Files
    If LCase(so.GetExtensionName(f.Name)) = "pptx" Then
        n = so.GetBaseName(f.Name)
        Set pt = pp.Presentations.Open(gf & "\" & f.Name,,, 0)
        pt.SaveAs gf & "\" & n & ".pdf", 32
        pt.close
        Set pt = Nothing
    End If
Next

pp. Quit

Set pp = Nothing
Set gf = Nothing
Set so = Nothing
MsgBox("Finished!")