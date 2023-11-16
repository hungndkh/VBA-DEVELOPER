Attribute VB_Name = "Unicode2char"
Function String2Char(ByVal MyStr As String) As String
     Dim Str As String, i As Integer, CStart As Integer, CCount As Integer, Start As Integer, t As String, k As String, dqQoute As String
     dqQoute = Chr(34)
     Start = 1 'Still text
     For i = 1 To Len(MyStr)
     t = Mid(MyStr, i, 1)
     k = AscW(t)
     If (k > 126 Or k = 13) And k <> 10 Then
     If k = 13 Then k = "vbCrLf"
     If Start = 0 Then String2Char = String2Char & dqQoute & " &": Start = 1
          If i = 1 Then
          String2Char = String2Char & "Chrw(" & k & ") & "
           ElseIf i < Len(MyStr) Then
            If Start = 1 Then String2Char = String2Char & " Chrw(" & k & ") & " Else String2Char = String2Char & " & Chrw(" & k & ") & "
          Else
            String2Char = String2Char & " & Chrw(" & k & ")"
          End If
      Else
      If Start = 1 Then String2Char = String2Char & dqQoute: Start = 0
       If i < Len(MyStr) Then String2Char = String2Char & t Else String2Char = String2Char & t & dqQoute
      End If
     Next i
     t = Replace(String2Char, "&  &", "&", , , vbTextCompare)
     t = Replace(t, "Chrw(vbCrLf)", "vbCrLf", , , vbTextCompare)
     t = Replace(t, vbLf, "", , , vbTextCompare)
     t = Replace(t, "& " & dqQoute & dqQoute & " &", "&", , , vbTextCompare)
     String2Char = t
End Function
Sub Label43Click()
Shell ("Explorer https://zalo.me/g/kmapds409")
End Sub

