VERSION 5.00
Begin VB.Form CntrlFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3 Libaray: Control Form"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Caption         =   "Edit Record"
      Height          =   3975
      Left            =   120
      TabIndex        =   56
      Top             =   120
      Width           =   4455
      Begin VB.TextBox ET 
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Index           =   4
         Left            =   960
         MaxLength       =   4
         TabIndex        =   84
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton EIB 
         Caption         =   "Done"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   82
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CommandButton EIB 
         Caption         =   "Update"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   81
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox ET 
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Index           =   3
         Left            =   960
         MaxLength       =   30
         TabIndex        =   72
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox ET 
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Index           =   2
         Left            =   960
         MaxLength       =   30
         TabIndex        =   71
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox ET 
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Index           =   1
         Left            =   960
         MaxLength       =   30
         TabIndex        =   70
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox ET 
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Index           =   0
         Left            =   960
         MaxLength       =   30
         TabIndex        =   69
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label ETL 
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   93
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label ETL 
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   92
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label ETL 
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   91
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label ETL 
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   90
         Top             =   720
         Width           =   375
      End
      Begin VB.Label ETL 
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   89
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   840
         TabIndex        =   88
         Top             =   2280
         Width           =   45
      End
      Begin VB.Label Label32 
         Caption         =   "Path"
         Height          =   255
         Left            =   240
         TabIndex        =   87
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3615
         TabIndex        =   86
         Top             =   1440
         Width           =   45
      End
      Begin VB.Label Label30 
         Caption         =   "ID ="
         Height          =   255
         Left            =   3240
         TabIndex        =   85
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "Year"
         Height          =   255
         Left            =   480
         TabIndex        =   83
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label L1 
         Caption         =   "a"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   840
         TabIndex        =   80
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label L1 
         Alignment       =   2  'Center
         Caption         =   "a"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   3480
         TabIndex        =   79
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label L1 
         Alignment       =   2  'Center
         Caption         =   "a"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   3120
         TabIndex        =   78
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label L1 
         Alignment       =   2  'Center
         Caption         =   "a"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   77
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label L1 
         Caption         =   "a"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   76
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label L1 
         Caption         =   "a"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   75
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label L1 
         Caption         =   "a"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   74
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label L1 
         Caption         =   "a"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   73
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "Layer"
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label27 
         Caption         =   "Original"
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "Copyright"
         Height          =   255
         Left            =   2280
         TabIndex        =   66
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "Mode"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label24 
         Caption         =   "Bit Rate"
         Height          =   255
         Left            =   2280
         TabIndex        =   64
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "CRC"
         Height          =   255
         Left            =   2280
         TabIndex        =   63
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "Version"
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "Size"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Comment"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Artist"
         Height          =   255
         Left            =   480
         TabIndex        =   59
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Album"
         Height          =   255
         Left            =   360
         TabIndex        =   58
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Title"
         Height          =   255
         Left            =   480
         TabIndex        =   57
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Rename selected File To"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command4 
         Caption         =   "&Cancel"
         Height          =   255
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Ok"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Copy Selected Files To"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command10 
         Caption         =   "&Browse"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "C&ancel"
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Copy"
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ListBox List3 
         Height          =   2010
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Move Selected Files To"
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command12 
         Caption         =   "&Cancel"
         Height          =   255
         Left            =   3000
         TabIndex        =   25
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&Move"
         Height          =   255
         Left            =   1560
         TabIndex        =   24
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Browse"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ListBox List2 
         Height          =   2010
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   4215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Select Path"
      Height          =   3255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancel"
         Height          =   255
         Left            =   2400
         TabIndex        =   21
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Select"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   2880
         Width           =   1455
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Top             =   2400
         Width           =   3975
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Select DataBase To Delete"
      Height          =   3255
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command17 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3000
         TabIndex        =   44
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton Command16 
         Caption         =   "&Delete Selected DataBase File"
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   2760
         Width           =   2655
      End
      Begin VB.FileListBox File1 
         Height          =   2430
         Left            =   240
         MultiSelect     =   2  'Extended
         Pattern         =   "*.mdb"
         TabIndex        =   42
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Options"
      Height          =   3255
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text4 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   46
         Top             =   465
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Save Changes On Exit"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   55
         Top             =   1200
         Width           =   3615
      End
      Begin VB.CommandButton Command15 
         Caption         =   "&Cancel"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   2880
         Width           =   4215
      End
      Begin VB.CommandButton Command18 
         Caption         =   "&Load Default"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   2640
         Width           =   4215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Display Hidden DataBase File"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   52
         Top             =   840
         Width           =   3615
      End
      Begin VB.CommandButton Command14 
         Caption         =   "&Save"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   2400
         Width           =   4215
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         TabIndex        =   50
         Top             =   1800
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Allow AutoSave Every"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Change The Path Of The CDROM Folder To:"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   1560
         Width           =   3210
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Minutes"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3000
         TabIndex        =   47
         Top             =   510
         Width           =   555
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Properties"
      Height          =   3255
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command13 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   1680
         TabIndex        =   30
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label14 
         Height          =   255
         Left            =   1200
         TabIndex        =   40
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label Label13 
         Height          =   255
         Left            =   1200
         TabIndex        =   39
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label12 
         Caption         =   "Album:"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Artist:"
         Height          =   195
         Left            =   360
         TabIndex        =   37
         Top             =   2040
         Width           =   390
      End
      Begin VB.Label Label9 
         Height          =   255
         Left            =   1200
         TabIndex        =   35
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   1200
         TabIndex        =   36
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label Label8 
         Height          =   495
         Left            =   1200
         TabIndex        =   34
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label7 
         Height          =   255
         Left            =   1200
         TabIndex        =   33
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label6 
         Height          =   495
         Left            =   1200
         TabIndex        =   32
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "File Type:"
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "File Path:"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "File size:"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "File Name:"
         Height          =   195
         Left            =   360
         TabIndex        =   27
         Top             =   480
         Width           =   750
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Delete Selected Files"
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command2 
         Caption         =   "&No"
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Yes"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   2880
         Width           =   1575
      End
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Are You Sure, You Want To Delete The Listed Files"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   4215
      End
   End
End
Attribute VB_Name = "CntrlFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Numb As Integer
Private Sub Command1_Click()
    On Error Resume Next
    Me.MousePointer = vbHourglass
    For I = 0 To List1.ListCount - 1
        Kill List1.List(I)
    Next I
    MainForm.File1.Path = Left(Dir1.Path, 3)
    Wait 0.1
    MainForm.File1.Path = Dir1Path
    Me.MousePointer = vbDefault
    Unload Me
End Sub
Private Sub Command10_Click()
    On Error Resume Next
    Command6.Default = True
    Numb = 3
    Frame5.Visible = True
    Frame5.ZOrder 0
End Sub
Private Sub Command11_Click()
    On Error Resume Next
    Me.MousePointer = vbHourglass
    Dim St As String, leng As Long
    For I = 0 To List2.ListCount - 1
        For leng = 1 To Len(List2.List(I))
            St = Right(List2.List(I), leng)
            If Left(St, 1) = "\" Then
                St = Right(List2.List(I), leng - 1)
                Exit For
            End If
        Next leng
        FileCopy List2.List(I), Text2Text & St
        Kill List2.List(I)
    Next I
    MainForm.File1.Path = Left(Dir1.Path, 3)
    Wait 0.1
    MainForm.File1.Path = Dir1Path
    Me.MousePointer = vbDefault
    Unload Me
End Sub
Private Sub Command12_Click()
    Unload Me
End Sub
Private Sub Command13_Click()
    Unload Me
End Sub
Private Sub Command14_Click()
    SaveSetting "MP3 Library", "CDROM", "CD", Text5.Text
    Unload Me
End Sub
Private Sub Command15_Click()
    Unload Me
End Sub
Private Sub Command16_Click()
    On Error GoTo ErrD
    Dim VYN
    VYN = MsgBox("Are You Sure You Want To Delete" & vbNewLine & "The Selected DataBase Files", vbYesNo + vbQuestion, "MP3 Library")
    If VYN = vbYes Then
        For I = 0 To File1.ListCount - 1
            If File1.Selected(I) = True Then
                Kill File1.Path & "\" & File1.List(I)
            End If
        Next I
        Unload Me
    Else
        Unload Me
    End If
    Exit Sub
ErrD:
    MsgBox Error, vbCritical + vbOKOnly, "Error"
    Unload Me
End Sub
Private Sub Command17_Click()
    Unload Me
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Command3_Click()
    On Error GoTo errordone
    Name Dir1Path & MainForm.File1.List(MainForm.File1.ListIndex) As Dir1Path & Text1.Text
    MainForm.File1.Path = Left(Dir1.Path, 3)
    Wait 0.1
    MainForm.File1.Path = Dir1Path
    Unload Me
    Exit Sub
errordone:
        MsgBox Error, vbCritical + vbOKOnly, "Error"
        Unload Me
End Sub
Private Sub Command4_Click()
    Unload Me
End Sub
Private Sub Command5_Click()
    On Error Resume Next
    Command6.Default = True
    Numb = 2
    Frame5.Visible = True
    Frame5.ZOrder 0
End Sub
Private Sub Command6_Click()
    On Error Resume Next
    Select Case Numb
        Case 3
            Text3.Text = Dir1.Path
        Case 2
            Text2.Text = Dir1.Path
    End Select
    Frame5.Visible = False
End Sub
Private Sub Command7_Click()
    On Error Resume Next
    Me.MousePointer = vbHourglass
    Dim St As String, leng As Long
    For I = 0 To List3.ListCount - 1
        For leng = 1 To Len(List3.List(I))
            St = Right(List3.List(I), leng)
            If Left(St, 1) = "\" Then
                St = Right(List3.List(I), leng - 1)
                Exit For
            End If
        Next leng
        FileCopy List3.List(I), Text3Text & St
    Next I
    Me.MousePointer = vbDefault
    Unload Me
End Sub
Private Sub Command8_Click()
    Unload Me
End Sub
Private Function Text2Text() As String
    On Error Resume Next
    If Right(Text2.Text, 1) = "\" Then
        Text2Text = Text2.Text
    Else
        Text2Text = Text2.Text & "\"
    End If
End Function
Private Sub Command9_Click()
    Frame5.Visible = False
End Sub
Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Left(Drive1.Drive, 2)
End Sub
Private Function Text3Text() As String
    On Error Resume Next
    If Right(Text3.Text, 1) = "\" Then
        Text3Text = Text3.Text
    Else
        Text3Text = Text3.Text & "\"
    End If
End Function
Private Function Text1Text() As String
    On Error Resume Next
    If Right(Text1.Text, 1) = "\" Then
        Text1Text = Text1.Text
    Else
        Text1Text = Text1.Text & "\"
    End If
End Function
Private Sub EIB_Click(Index As Integer)
    Select Case Index
    Case 0
        If Len(ET(0).Text) > 0 Then RSet1.Fields("Title") = ET(0).Text
        If Len(ET(1).Text) > 0 Then RSet1.Fields("Artist") = ET(1).Text
        If Len(ET(2).Text) > 0 Then RSet1.Fields("Album") = ET(2).Text
        If Len(ET(3).Text) > 0 Then RSet1.Fields("Comment") = ET(3).Text
        If Len(ET(4).Text) > 0 Then RSet1.Fields("Year") = ET(4).Text
        RSet1.Update
        Beep
        MainForm.MSFlexGrid1.Col = 1
        MainForm.MSFlexGrid1.Text = ET(0).Text
        MainForm.MSFlexGrid1.Col = 2
        MainForm.MSFlexGrid1.Text = ET(1).Text
        Unload Me
    Case 1
        Unload Me
    End Select
End Sub
Private Sub ET_Change(Index As Integer)
    ETL(Index).Caption = Len(ET(Index).Text)
End Sub
Private Sub Form_Load()
    Me.Height = 3900
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FC = 0
    FT = vbNullString
End Sub
