Attribute VB_Name = "Add_Record_Mod"
Public Sub AddRec(AllRec As Boolean)
    Dim I As Long, TempS As String, LCount As Long, TempHeader As String
    TempHeader = MainForm.Dir1.Path
    If Right(TempHeader, 1) <> "\" Then TempHeader = TempHeader & "\"
    LCount = 0
    For I = 0 To MainForm.File1.ListCount - 1
        TempS = ""
        If AllRec = True Then
            TempS = MainForm.File1.List(I)
            LCount = LCount + 1
        ElseIf AllRec = False And MainForm.File1.Selected(I) = True Then
            TempS = MainForm.File1.List(I)
            LCount = LCount + 1
        End If
        DoEvents
        WaitF.Label1.Caption = TempS
        If Len(TempS) > 0 Then ProcessSong TempHeader & TempS
    Next I
    Unload WaitF
    MsgBox LCount & " records have been successfully added to database.", vbOKOnly + vbInformation
End Sub
Private Function StripString(SStr As String) As String
    Dim I As Long
    I = InStr(1, SStr, Chr(255), vbTextCompare)
    If I = 0 Then
        StripString = Trim(SStr)
    Else
        StripString = Trim(Left(SStr, I - 1))
    End If
End Function
Private Sub ProcessSong(Songtitle As String)
    Dim FHeader As String, SYear As String
    Open Songtitle For Binary As #1
    RSet1.AddNew
    '== Analyse MP3 File Header
    FHeader = Input(4, #1)
    AddBits AnalyseBin(Asc(Mid(FHeader, 2, 1))), 2
    AddBits AnalyseBin(Asc(Mid(FHeader, 3, 1))), 3
    AddBits AnalyseBin(Asc(Mid(FHeader, 4, 1))), 4
    '== Analyse MP3 File TAG
    RSet1.Fields("FileName") = Songtitle
    RSet1.Fields("Size") = FileLen(Songtitle)
    Seek #1, LOF(1) - 127
    If Input(3, #1) = "TAG" Then
        Seek #1, LOF(1) - 124
        RSet1.Fields("Title") = StripString(Input(30, #1))
        Seek #1, LOF(1) - 94
        RSet1.Fields("Artist") = StripString(Input(30, #1))
        Seek #1, LOF(1) - 64
        RSet1.Fields("Album") = StripString(Input(30, #1))
        Seek #1, LOF(1) - 34
        SYear = Input(4, #1)
        If IsNumeric(SYear) = True Then RSet1.Fields("Year") = CLng(SYear)
        Seek #1, LOF(1) - 30
        RSet1.Fields("Comment") = StripString(Input(30, #1))
    Else
        RSet1.Fields("Title") = "Unknown Title"
        RSet1.Fields("Artist") = "Unknown Artist"
        RSet1.Fields("Album") = "Unknown Album"
        RSet1.Fields("Comment") = "Unknown Comment"
    End If
    RSet1.Update
    MainForm.MSFlexGrid1.AddItem RSet1.Fields("ID") & vbTab & RSet1.Fields("Title") & vbTab & RSet1.Fields("Artist")
    Close #1
End Sub
Private Function AnalyseBin(ChrCode As Long) As Long
    'convert to binary
    Dim BinPos(8) As Long, I As Long
    For I = 1 To 8
        BinPos(I) = 0
    Next I
    If ChrCode >= 128 Then
        BinPos(8) = 10000000
        ChrCode = ChrCode - 128
    End If
    If ChrCode >= 64 Then
        BinPos(7) = 1000000
        ChrCode = ChrCode - 64
    End If
    If ChrCode >= 32 Then
        BinPos(6) = 100000
        ChrCode = ChrCode - 32
    End If
    If ChrCode >= 16 Then
        BinPos(5) = 10000
        ChrCode = ChrCode - 16
    End If
    If ChrCode >= 8 Then
        BinPos(4) = 1000
        ChrCode = ChrCode - 8
    End If
    If ChrCode >= 4 Then
        BinPos(3) = 100
        ChrCode = ChrCode - 4
    End If
    If ChrCode >= 2 Then
        BinPos(2) = 10
        ChrCode = ChrCode - 2
    End If
    If ChrCode >= 1 Then
        BinPos(1) = 1
        ChrCode = ChrCode - 1
    End If
    For I = 1 To 8
        AnalyseBin = AnalyseBin + BinPos(I)
    Next I
End Function
Private Sub AddBits(ChrByte As String, ChrPos As Long)
    'analyse binary & add necessary values
    If Len(ChrByte) < 8 Then ChrByte = String(8 - Len(ChrByte), "0") & ChrByte
    Select Case ChrPos
    Case Is = 2
        RSet1.Fields("Version") = CLng(Mid(ChrByte, 5, 1))
        If Mid(ChrByte, 6, 2) = "00" Then
            RSet1.Fields("Layer") = 1
        ElseIf Mid(ChrByte, 6, 2) = "01" Then
            RSet1.Fields("Layer") = 2
        ElseIf Mid(ChrByte, 6, 2) = "11" Then
            RSet1.Fields("Layer") = 2
        ElseIf Mid(ChrByte, 6, 2) = "10" Then
            RSet1.Fields("Layer") = 3
        End If
        RSet1.Fields("Error") = CLng(Mid(ChrByte, 8, 1))
    Case Is = 3
        If Mid(ChrByte, 5, 2) = "00" Then
            RSet1.Fields("BitRate") = 0
        ElseIf Mid(ChrByte, 5, 2) = "01" Then
            RSet1.Fields("BitRate") = 1
        ElseIf Mid(ChrByte, 5, 2) = "10" Then
            RSet1.Fields("BitRate") = 2
        End If
        RSet1.Fields("Extension") = CLng(Mid(ChrByte, 8, 1))
    Case Is = 4
        If Mid(ChrByte, 1, 2) = "00" Then
            RSet1.Fields("Mode") = 0
        ElseIf Mid(ChrByte, 1, 2) = "01" Then
            RSet1.Fields("Mode") = 1
        ElseIf Mid(ChrByte, 1, 2) = "10" Then
            RSet1.Fields("Mode") = 2
        ElseIf Mid(ChrByte, 1, 2) = "11" Then
            RSet1.Fields("Mode") = 3
        End If
        RSet1.Fields("Copyright") = CLng(Mid(ChrByte, 5, 1))
        RSet1.Fields("Original") = CLng(Mid(ChrByte, 6, 1))
    End Select
End Sub
Public Sub EditItem(ItemID As Long)
    Dim TempV As Long
    RSet1.Filter = "ID=" & ItemID
    RSet1.MoveFirst
    For TempV = CntrlFrm.ET.LBound To CntrlFrm.ET.UBound
        CntrlFrm.ET(TempV).Text = ""
    Next TempV
    For TempV = CntrlFrm.L1.LBound To CntrlFrm.L1.UBound
        CntrlFrm.L1(TempV).Caption = ""
    Next TempV
    If IsNull(RSet1.Fields("Title")) = False Then CntrlFrm.ET(0).Text = RSet1.Fields("Title")
    If IsNull(RSet1.Fields("Artist")) = False Then CntrlFrm.ET(1).Text = RSet1.Fields("Artist")
    If IsNull(RSet1.Fields("Album")) = False Then CntrlFrm.ET(2).Text = RSet1.Fields("Album")
    If IsNull(RSet1.Fields("Year")) = False Then CntrlFrm.ET(4).Text = RSet1.Fields("Year")
    If IsNull(RSet1.Fields("Comment")) = False Then CntrlFrm.ET(3).Text = RSet1.Fields("Comment")
    
    CntrlFrm.Label31.Caption = RSet1.Fields("ID")
    If IsNull(RSet1.Fields("FileName")) = False Then
        CntrlFrm.Label33.Caption = RSet1.Fields("FileName")
        CntrlFrm.Label33.ToolTipText = RSet1.Fields("FileName")
    End If
    CntrlFrm.L1(0).Caption = RSet1.Fields("Size") & "  Bytes"
    If RSet1.Fields("Version") = 0 Then
        CntrlFrm.L1(1).Caption = "MPEG 2.0"
    ElseIf RSet1.Fields("Version") = 1 Then
        CntrlFrm.L1(1).Caption = "MPEG 1.0"
    End If

    If RSet1.Fields("Layer") = 0 Then
        CntrlFrm.L1(2).Caption = "1"
    ElseIf RSet1.Fields("Layer") = 1 Then
        CntrlFrm.L1(2).Caption = "3"
    ElseIf RSet1.Fields("Layer") = 3 Then
        CntrlFrm.L1(2).Caption = "2"
    ElseIf RSet1.Fields("Layer") = 2 Then
        CntrlFrm.L1(2).Caption = "3"
    End If
    
    If RSet1.Fields("Mode") = 0 Then
        CntrlFrm.L1(3).Caption = "Stereo"
    ElseIf RSet1.Fields("Mode") = 1 Then
        CntrlFrm.L1(3).Caption = "Joint Stereo"
    ElseIf RSet1.Fields("Mode") = 2 Then
        CntrlFrm.L1(3).Caption = "Dual Channel"
    ElseIf RSet1.Fields("Mode") = 3 Then
        CntrlFrm.L1(3).Caption = "Mono"
    End If
    
    If RSet1.Fields("Copyright") = 0 Then
        CntrlFrm.L1(4).Caption = "No"
    ElseIf RSet1.Fields("Copyright") = 1 Then
        CntrlFrm.L1(4).Caption = "Yes"
    End If
    
    Select Case RSet1.Fields("BitRate")
        Case Is = 0
            If RSet1.Fields("Version") = 0 Then
                CntrlFrm.L1(5).Caption = "22 kHz"
            ElseIf RSet1.Fields("Version") = 1 Then
                CntrlFrm.L1(5).Caption = "44.1 kHz"
            End If
        Case Is = 1
            If RSet1.Fields("Version") = 0 Then
                CntrlFrm.L1(5).Caption = "24 kHz"
            ElseIf RSet1.Fields("Version") = 1 Then
                CntrlFrm.L1(5).Caption = "48 kHz"
            End If
        Case Is = 2
            If RSet1.Fields("Version") = 0 Then
                CntrlFrm.L1(5).Caption = "16 kHz"
            ElseIf RSet1.Fields("Version") = 1 Then
                CntrlFrm.L1(5).Caption = "38 kHz"
            End If
    End Select
    
    If RSet1.Fields("Error") = 0 Then
        CntrlFrm.L1(6).Caption = "Yes"
    ElseIf RSet1.Fields("Error") = 1 Then
        CntrlFrm.L1(6).Caption = "No"
    End If
    
    If RSet1.Fields("Original") = 0 Then
        CntrlFrm.L1(7).Caption = "No"
    ElseIf RSet1.Fields("Original") = 1 Then
        CntrlFrm.L1(7).Caption = "Yes"
    End If
    
    CntrlFrm.Height = 4590
    CntrlFrm.Frame9.Visible = True
    CntrlFrm.Frame9.ZOrder 0
    CntrlFrm.Show vbModal
    RSet1.Filter = adFilterNone
End Sub
