Attribute VB_Name = "Editor_RemoveComment"
'De su dung code can them thu vien Visual Basic for Application Extensibility 5.3
'Code chi xoa comment nguyen dong. Comment cuoi dong se khong xu ly duoc
Sub DeleteCommentLinesAllModules(pj As VBProject)
Dim i As Long, md As VBComponent, t As String, arr() As String
For Each md In pj.VBComponents
With md.CodeModule
endLine = .CountOfLines
If endLine > 1 Then
    t = .Lines(1, endLine)
    arr = Split(t, vbCrLf, , vbTextCompare)
    For i = 0 To UBound(arr)
    t = Trim(arr(i))
    arr(i) = t
    If Left(t, 1) = "'" Then arr(i) = ""
    Next i
    t = Join(arr, vbCrLf)
    .DeleteLines 1, endLine
    .InsertLines 1, t
End If
End With
Next md
End Sub
Sub DeleteCommentLinesAllModules2(pjName As String)
Dim i As Long, md As VBComponent, t As String, arr() As String, vb As vbe, pj As VBProject
On Error GoTo exitsub:
Set vb = Application.vbe
Set pj = vb.VBProjects(pjName)
For Each md In pj.VBComponents
With md.CodeModule
endLine = .CountOfLines
If endLine > 1 Then 'Bo qua neu codemodule trong hoac chi co 1 dong
    t = .Lines(1, endLine) 'Gan noi dung code vao t => neu code >65526 ky tu se loi
    arr = Split(t, vbCrLf, , vbTextCompare) 'Dung 1 array de xu ly theo dong se nhanh hon dung CodeModule.line de lay tung dong
    For i = 0 To UBound(arr)
    t = Trim(arr(i))
    arr(i) = t 'Xoa tab dau dong
    If Left(t, 1) = "'" Then arr(i) = "" 'Xoa dong comment
    Next i
    t = Join(arr, vbCrLf)
    .DeleteLines 1, endLine
    .InsertLines 1, t
End If
End With
Next md
exitsub:
End Sub
