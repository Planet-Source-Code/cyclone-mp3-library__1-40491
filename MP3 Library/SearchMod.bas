Attribute VB_Name = "SearchMod"
Public Sub SearchDB(SDBN As String)
    Dim SrchFld As String, XSrch As String
    If MainForm.Check2.Value = 1 Then
        XSrch = " = '" & MainForm.Text1.Text & "') "
    ElseIf MainForm.Check2.Value = 0 Then
        XSrch = " LIKE ('%" & MainForm.Text1.Text & "%') "
    End If
    SrchFld = MainForm.Combo2.Text
    RSet2.close
    RSet2.Open "SELECT * FROM List WHERE " & SrchFld & XSrch & " ORDER BY " & SrchFld
    RSet2.MoveFirst
    If RSet2.RecordCount = 0 Then Exit Sub
    Do While Not RSet2.EOF
        DoEvents
        MainForm.MSFlexGrid2.AddItem RSet2.Fields("ID") & vbTab & RSet2.Fields(SrchFld) & vbTab & RSet2.Fields("Title") & vbTab & SDBN
        RSet2.MoveNext
    Loop
    MainForm.MSFlexGrid2.RemoveItem 1
End Sub
