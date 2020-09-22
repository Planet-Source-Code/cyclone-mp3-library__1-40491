VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form SplashFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00808080&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Thank's"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture5 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   240
      Picture         =   "SplashFrm.frx":0000
      ScaleHeight     =   570
      ScaleWidth      =   645
      TabIndex        =   10
      Top             =   840
      Width           =   645
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   840
      Picture         =   "SplashFrm.frx":05E0
      ScaleHeight     =   570
      ScaleWidth      =   645
      TabIndex        =   2
      Top             =   360
      Width           =   645
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   0
      Picture         =   "SplashFrm.frx":0BC0
      ScaleHeight     =   570
      ScaleWidth      =   645
      TabIndex        =   0
      Top             =   120
      Width           =   645
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   840
      Picture         =   "SplashFrm.frx":11A0
      ScaleHeight     =   570
      ScaleWidth      =   645
      TabIndex        =   11
      Top             =   1320
      Width           =   645
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   960
      Top             =   2040
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   1800
      TabIndex        =   6
      Top             =   2880
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   0
      Picture         =   "SplashFrm.frx":1780
      ScaleHeight     =   570
      ScaleWidth      =   645
      TabIndex        =   1
      Top             =   1560
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "CycLonE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      TabIndex        =   4
      Top             =   3000
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "All Rights Reserved."
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1680
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080C0FF&
      Caption         =   "Registered To"
      Height          =   195
      Left            =   3540
      TabIndex        =   12
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "designed by"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   885
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "V 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Please wait while loading..."
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Library"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   3480
      TabIndex        =   7
      Top             =   240
      Width           =   1125
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3615
      Left            =   -120
      Top             =   -120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "MP3"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1245
      Left            =   1440
      TabIndex        =   9
      Top             =   480
      Width           =   2085
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   3615
      Left            =   1560
      Top             =   -240
      Width           =   3135
   End
End
Attribute VB_Name = "SplashFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Const REG_SZ = 1
Private Const REG_DWORD = 4
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_ALL_ACCESS = &H2003F
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Dim OSInfo As OSVERSIONINFO, PId As String, RootName As String
    Dim ResultName1 As String, ResultName2 As String
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    Ret& = GetVersionEx(OSInfo)
    RootName = "SOFTWARE\Microsoft\Windows\CurrentVersion"
    If OSInfo.dwPlatformId = 2 Then RootName = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    about = 0
    GetKeyValue HKEY_LOCAL_MACHINE, RootName, "RegisteredOwner", ResultName1
    GetKeyValue HKEY_LOCAL_MACHINE, RootName, "RegisteredOrganization", ResultName2
    Label8.Caption = ResultName1
    Label9.Caption = ResultName2
    Timer1.Enabled = False
    
    MainForm.File2.Path = AppPath & "Database\"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'MainForm.Show
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'me.mousepointer = vbDefault
    Label2.ForeColor = RGB(0, 0, 0)
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.ForeColor = RGB(50, 50, 255)
End Sub
Private Sub Timer1_Timer()
    If ProgressBar1.Value >= 100 Then
        MainForm.Show
        Unload Me
        Timer1.Enabled = False
        Exit Sub
    End If
    ProgressBar1.Value = ProgressBar1.Value + 2
End Sub
Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim I As Long
        Dim rc As Long
        Dim Hkey As Long
        Dim KeyValType As Long
        Dim tmpVal As String
        Dim KeyValSize As Long
        
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, Hkey)
        tmpVal = String$(1024, 0)
        KeyValSize = 1024
        
        rc = RegQueryValueEx(Hkey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)
        If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
                tmpVal = Left(tmpVal, KeyValSize - 1)
        Else
                tmpVal = Left(tmpVal, KeyValSize)
        End If
        KeyVal = tmpVal
        GetKeyValue = True
        rc = RegCloseKey(Hkey)
End Function
