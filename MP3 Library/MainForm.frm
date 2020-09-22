VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3 Library"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8175
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5835
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3148
            MinWidth        =   1587
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7806
            MinWidth        =   6245
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   820
            MinWidth        =   829
            TextSave        =   "NUM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   900
            MinWidth        =   900
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
            TextSave        =   "5:39 PM"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   5175
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   7695
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Text            =   "*.mp3"
         Top             =   2760
         Width           =   2655
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   2655
      End
      Begin VB.DirListBox Dir1 
         Height          =   2340
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2655
      End
      Begin VB.FileListBox File1 
         Height          =   3210
         Left            =   2880
         MultiSelect     =   2  'Extended
         TabIndex        =   4
         Top             =   240
         Width           =   4695
      End
      Begin VB.Frame Frame5 
         Height          =   1575
         Left            =   120
         TabIndex        =   8
         Top             =   3480
         Width           =   7455
         Begin VB.CommandButton SExp 
            Caption         =   "&Properties"
            Height          =   255
            Index           =   7
            Left            =   6120
            TabIndex        =   12
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton SExp 
            Caption         =   "&Append To DB"
            Height          =   255
            Index           =   6
            Left            =   6120
            TabIndex        =   17
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton SExp 
            Caption         =   "&Delete"
            Height          =   255
            Index           =   5
            Left            =   6120
            TabIndex        =   9
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton SExp 
            Caption         =   "&Rename"
            Height          =   255
            Index           =   4
            Left            =   6120
            TabIndex        =   13
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton SExp 
            Caption         =   "&Copy"
            Height          =   255
            Index           =   3
            Left            =   6720
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton SExp 
            Caption         =   "&Move"
            Height          =   255
            Index           =   2
            Left            =   6120
            TabIndex        =   11
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton SExp 
            Caption         =   "Clear Se&lection"
            Height          =   255
            Index           =   1
            Left            =   4440
            TabIndex        =   21
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton SExp 
            Caption         =   "&Select All"
            Height          =   255
            Index           =   0
            Left            =   2880
            TabIndex        =   20
            Top             =   240
            Width           =   1575
         End
         Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
            Height          =   735
            Left            =   2880
            TabIndex        =   14
            ToolTipText     =   "Press SpaceBar or The Play Button To Play The Selected Item."
            Top             =   720
            Width           =   3135
            AudioStream     =   -1
            AutoSize        =   0   'False
            AutoStart       =   0   'False
            AnimationAtStart=   -1  'True
            AllowScan       =   -1  'True
            AllowChangeDisplaySize=   -1  'True
            AutoRewind      =   -1  'True
            Balance         =   0
            BaseURL         =   ""
            BufferingTime   =   5
            CaptioningID    =   ""
            ClickToPlay     =   -1  'True
            CursorType      =   0
            CurrentPosition =   -1
            CurrentMarker   =   0
            DefaultFrame    =   ""
            DisplayBackColor=   0
            DisplayForeColor=   16777215
            DisplayMode     =   0
            DisplaySize     =   4
            Enabled         =   -1  'True
            EnableContextMenu=   -1  'True
            EnablePositionControls=   -1  'True
            EnableFullScreenControls=   0   'False
            EnableTracker   =   -1  'True
            Filename        =   ""
            InvokeURLs      =   -1  'True
            Language        =   -1
            Mute            =   0   'False
            PlayCount       =   1
            PreviewMode     =   0   'False
            Rate            =   1
            SAMILang        =   ""
            SAMIStyle       =   ""
            SAMIFileName    =   ""
            SelectionStart  =   -1
            SelectionEnd    =   -1
            SendOpenStateChangeEvents=   -1  'True
            SendWarningEvents=   -1  'True
            SendErrorEvents =   -1  'True
            SendKeyboardEvents=   0   'False
            SendMouseClickEvents=   0   'False
            SendMouseMoveEvents=   0   'False
            SendPlayStateChangeEvents=   -1  'True
            ShowCaptioning  =   0   'False
            ShowControls    =   -1  'True
            ShowAudioControls=   -1  'True
            ShowDisplay     =   0   'False
            ShowGotoBar     =   0   'False
            ShowPositionControls=   0   'False
            ShowStatusBar   =   0   'False
            ShowTracker     =   -1  'True
            TransparentAtStart=   0   'False
            VideoBorderWidth=   0
            VideoBorderColor=   0
            VideoBorder3D   =   -1  'True
            Volume          =   -600
            WindowlessVideo =   0   'False
         End
         Begin VB.Shape Shape8 
            Height          =   1215
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   480
            TabIndex        =   15
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Label6 
            Height          =   615
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label4 
            Height          =   255
            Left            =   1560
            TabIndex        =   19
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Files Found:"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   1260
         End
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9975
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&General"
            Key             =   "g"
            Object.Tag             =   ""
            Object.ToolTipText     =   "General Page"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Database Manager"
            Key             =   "d"
            Object.Tag             =   ""
            Object.ToolTipText     =   "DataBase Manager"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&File Manager"
            Key             =   "f"
            Object.Tag             =   ""
            Object.ToolTipText     =   "File Manager"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "F&ind"
            Key             =   "i"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Find Songs"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   240
      TabIndex        =   16
      Top             =   480
      Width           =   7695
      Begin VB.CheckBox Check1 
         Caption         =   "Enable Quick View"
         Height          =   255
         Left            =   2640
         TabIndex        =   69
         Top             =   3000
         Width           =   1695
      End
      Begin ComctlLib.ProgressBar ProgressBar2 
         Height          =   135
         Left            =   600
         TabIndex        =   68
         Top             =   4680
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.FileListBox File2 
         Height          =   2820
         Left            =   240
         Pattern         =   "*.mdb"
         TabIndex        =   39
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton MainBut 
         Caption         =   "Close and E&xit"
         Height          =   255
         Index           =   1
         Left            =   5280
         TabIndex        =   32
         Top             =   3840
         Width           =   2055
      End
      Begin VB.CommandButton MainBut 
         Caption         =   "&Load Selected Database"
         Height          =   255
         Index           =   0
         Left            =   5280
         MaskColor       =   &H00000000&
         TabIndex        =   31
         Top             =   3600
         Width           =   2055
      End
      Begin VB.ListBox List5 
         Columns         =   2
         Height          =   2400
         Left            =   2640
         TabIndex        =   29
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label Label42 
         Caption         =   "(will slow system performance)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4320
         TabIndex        =   70
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Shape Shape10 
         Height          =   735
         Left            =   5160
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "MP3 Library"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   5340
         TabIndex        =   54
         Top             =   4477
         Width           =   1965
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         FillColor       =   &H0080C0FF&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   5160
         Shape           =   4  'Rounded Rectangle
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Shape Shape6 
         Height          =   1575
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   4815
      End
      Begin VB.Label Label33 
         Height          =   255
         Left            =   1800
         TabIndex        =   41
         Top             =   4080
         Width           =   3135
      End
      Begin VB.Label Label32 
         Caption         =   "Items In DataBase:"
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label31 
         Height          =   255
         Left            =   1800
         TabIndex        =   38
         Top             =   4320
         Width           =   3135
      End
      Begin VB.Label Label30 
         Caption         =   "Total Files Number:"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label29 
         Height          =   255
         Left            =   1800
         TabIndex        =   36
         Top             =   3840
         Width           =   3135
      End
      Begin VB.Label Label28 
         Height          =   255
         Left            =   1800
         TabIndex        =   35
         Top             =   3600
         Width           =   3135
      End
      Begin VB.Label Label27 
         Caption         =   "Database Size:"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "Database Name:"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Database QuickView"
         Height          =   195
         Left            =   2640
         TabIndex        =   30
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Existing Database Files:"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1680
      End
   End
   Begin VB.Frame Frame8 
      Height          =   5175
      Left            =   240
      TabIndex        =   27
      Top             =   480
      Width           =   7695
      Begin VB.CheckBox Check2 
         Caption         =   "Exact Match"
         Height          =   255
         Left            =   3480
         TabIndex        =   81
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   600
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   1815
         Left            =   120
         TabIndex        =   78
         Top             =   3240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.CommandButton Command26 
         Caption         =   "&Search Now"
         Height          =   555
         Left            =   6360
         TabIndex        =   49
         Top             =   240
         Width           =   1215
      End
      Begin VB.FileListBox File3 
         Height          =   870
         Left            =   5160
         MultiSelect     =   1  'Simple
         TabIndex        =   46
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   48
         Top             =   240
         Width           =   5055
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Only the database files selected from the List"
         Height          =   195
         Left            =   1320
         TabIndex        =   45
         Top             =   2760
         Width           =   3735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "All Database Files"
         Height          =   195
         Left            =   1320
         TabIndex        =   44
         Top             =   2520
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Current Database"
         Height          =   255
         Left            =   1320
         TabIndex        =   43
         Top             =   2280
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.ListBox List6 
         Height          =   450
         Left            =   5880
         TabIndex        =   53
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Search Fields:"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   600
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         Height          =   1215
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   7455
      End
      Begin VB.Label Label38 
         Caption         =   "Warning:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "The Use Of  "".mp3"" At The End Of The Search Query Is Not Necessary."
         Height          =   195
         Left            =   480
         TabIndex        =   51
         Top             =   1800
         Width           =   5130
      End
      Begin VB.Label Label36 
         Caption         =   "The Use of wildcards such as (* or ?)  Is Not Allowed."
         Height          =   195
         Left            =   480
         TabIndex        =   50
         Top             =   1440
         Width           =   6885
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Search Query:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label34 
         Caption         =   "Find a Song in:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   2280
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5175
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   7695
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   6360
         Picture         =   "MainForm.frx":0442
         ScaleHeight     =   540
         ScaleWidth      =   525
         TabIndex        =   77
         Top             =   4080
         Width           =   525
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2655
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   3
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.Frame Frame6 
         Height          =   2775
         Left            =   5640
         TabIndex        =   64
         Top             =   120
         Width           =   1935
         Begin VB.CommandButton RASlct 
            Caption         =   "Select All"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   73
            Top             =   2400
            Width           =   1695
         End
         Begin VB.CommandButton RASlct 
            Caption         =   "Remove"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   74
            Top             =   2160
            Width           =   1695
         End
         Begin VB.CommandButton RASlct 
            Caption         =   "Edit"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   75
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CommandButton RASlct 
            Caption         =   "Add"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   76
            Top             =   1680
            Width           =   1695
         End
         Begin VB.CommandButton DBCommand 
            Caption         =   "&Delete"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   65
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton DBCommand 
            Caption         =   "&Save and Close"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   72
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton DBCommand 
            Caption         =   "&Load"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   66
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton DBCommand 
            Caption         =   "&New"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1695
         Left            =   120
         TabIndex        =   56
         Top             =   3360
         Width           =   5415
         Begin VB.Label Label2 
            Caption         =   "Items In List:"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label5 
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   360
            Width           =   4935
         End
         Begin VB.Label Label7 
            Height          =   255
            Left            =   1080
            TabIndex        =   60
            Top             =   1320
            Width           =   3855
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "DataBase:"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   1080
            Width           =   750
         End
         Begin VB.Label Label11 
            Caption         =   "Current Database Size:"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label12 
            Height          =   255
            Left            =   1920
            TabIndex        =   57
            Top             =   840
            Width           =   3375
         End
         Begin VB.Shape Shape5 
            Height          =   495
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   240
            Width           =   5175
         End
         Begin VB.Label Label10 
            Height          =   255
            Left            =   1080
            TabIndex        =   63
            Top             =   1080
            Width           =   4095
         End
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   3000
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.PictureBox Picture6 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Index           =   1
         Left            =   5760
         Picture         =   "MainForm.frx":0A0E
         ScaleHeight     =   540
         ScaleWidth      =   525
         TabIndex        =   26
         Top             =   4440
         Width           =   525
      End
      Begin VB.PictureBox Picture6 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Index           =   0
         Left            =   6960
         Picture         =   "MainForm.frx":0FDA
         ScaleHeight     =   540
         ScaleWidth      =   525
         TabIndex        =   24
         Top             =   4440
         Width           =   525
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   5880
         Picture         =   "MainForm.frx":15A6
         ScaleHeight     =   540
         ScaleWidth      =   525
         TabIndex        =   25
         Top             =   3600
         Width           =   525
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6000
         Top             =   3600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox Picture5 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   6840
         Picture         =   "MainForm.frx":1B72
         ScaleHeight     =   540
         ScaleWidth      =   525
         TabIndex        =   23
         Top             =   3600
         Width           =   525
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Left            =   5640
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   1935
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8520
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu MainM 
      Caption         =   "&Main"
      Begin VB.Menu new 
         Caption         =   "&New Database"
      End
      Begin VB.Menu load 
         Caption         =   "&Load Database"
      End
      Begin VB.Menu close 
         Caption         =   "&Close Database"
      End
      Begin VB.Menu delete 
         Caption         =   "&Delete Database"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu OptionM 
      Caption         =   "&Options"
      Visible         =   0   'False
   End
   Begin VB.Menu HelpM 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Here Starts The File Manager Tab Codes !!!!!
Option Explicit
Private Sub about_Click()
    SplashFrm.Label10.Visible = True
    SplashFrm.Label3.Visible = False
    SplashFrm.Timer1.Enabled = False
    SplashFrm.ProgressBar1.Visible = False
    SplashFrm.Command1.Visible = True
    SplashFrm.Show vbModal
End Sub
Private Sub close_Click()
    DBCommand(2).Value = True
End Sub
Private Sub Combo1_Change()
    File1.Pattern = Combo1.Text
    File1.Path = Dir1.List(Dir1.ListIndex)
End Sub
Private Sub Combo1_Click()
    File1.Pattern = Combo1.Text
    File1.Path = Dir1.List(Dir1.ListIndex)
End Sub
Private Sub DBCommand_Click(Index As Integer)
    Select Case Index
        Case Is = 0
            'new database file
            If CheckRS = False Then Exit Sub
NDB:
            Dim Answer As String
            Answer = InputBox("Please Enter The DataBase Name.", "MP3 Library - New DataBase")
            If InStr(1, Answer, "/", vbTextCompare) Or InStr(1, Answer, "\", vbTextCompare) Or InStr(1, Answer, ";", vbTextCompare) Or InStr(1, Answer, ",", vbTextCompare) Or InStr(1, Answer, "'", vbTextCompare) Or InStr(1, Answer, "`", vbTextCompare) Or InStr(1, Answer, "(", vbTextCompare) Or InStr(1, Answer, ")", vbTextCompare) Or InStr(1, Answer, Chr(34), vbTextCompare) Or InStr(1, Answer, ":", vbTextCompare) Then
                MsgBox "DataBase Name Cannot Contains Any Of These Characters" & vbNewLine & "/\;,'`(:)" & Chr(34), vbCritical + vbOKOnly, "MP3 Library"
                GoTo NDB
            End If
            If Answer = "" Then Exit Sub
            If LCase(Right(Answer, 4)) <> ".mdb" Then Answer = Answer & ".mdb"
            If Len(Dir(App.Path & "\DataBase\" & Answer)) > 0 Then
                MsgBox "DataBase Already Exist.", vbCritical + vbOKOnly, "MP3 Library"
                GoTo NDB
            End If
            
            StatusBar1.Panels(2).Text = "creating new database file..."
            FileCopy AppPath & "Database\Template.mdb", AppPath & "Database\" & Answer
            OpenDBFile AppPath & "Database\" & Answer, Answer, RSet1, DBase1
        Case Is = 1
            If CheckRS = False Then Exit Sub
            Me.MousePointer = vbHourglass
            CommonDialog1.CancelError = True
            On Error GoTo ErrHandler
            CommonDialog1.DialogTitle = "Load DataBase File"
            CommonDialog1.InitDir = AppPath & "\DataBase\"
            CommonDialog1.Filter = "DataBase Files (*.mdb)|*.mdb"
            CommonDialog1.ShowOpen
            OpenDBFile CommonDialog1.FileName, CommonDialog1.FileTitle, RSet1, DBase1
            ReadRS
            Me.MousePointer = vbDefault
            Exit Sub
ErrHandler:
            Me.MousePointer = vbDefault
            Exit Sub
        Case Is = 2
            If RSet1.State = adStateClosed Then Exit Sub
            If MsgBox("Are you sure you want to close " & Label10.Caption, vbQuestion + vbYesNo) = vbYes Then CloseDB
        Case Is = 3
            CntrlFrm.Frame9.Visible = False
            CntrlFrm.Height = 3870
            CntrlFrm.File1.Hidden = False
            CntrlFrm.File1.FileName = "*.mdb"
            CntrlFrm.File1.Path = App.Path & "\Database"
            CntrlFrm.Frame8.Visible = True
            CntrlFrm.Frame8.ZOrder 0
            CntrlFrm.Show vbModal
    End Select
End Sub
Private Sub DBCommand_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
    Case Is = 0
        StatusBar1.Panels(2).Text = "Create a New Database"
    Case Is = 1
        StatusBar1.Panels(2).Text = "Load an Existing Database"
    Case Is = 2
        StatusBar1.Panels(2).Text = "Close Current Database"
    Case Is = 3
        StatusBar1.Panels(2).Text = "Delete Selected Database"
    End Select
End Sub
Private Sub delete_Click()
    DBCommand(3).Value = True
End Sub
Private Sub Dir1_Change()
    On Error Resume Next
    StatusBar1.Panels(2).Text = UCase(Dir1Path)
    File1.FileName = Combo1.Text
    File1.Path = Dir1Path
    Label4.Caption = File1.ListCount
End Sub
Private Sub Dir1_Click()
    On Error Resume Next
    Dir1.Path = Dir1.List(Dir1.ListIndex)
    File1.FileName = Combo1.Text
    File1.Path = Dir1.List(Dir1.ListIndex)
    Label4.Caption = File1.ListCount
    Label6.Caption = ""
    Label3.Caption = ""
End Sub
Private Sub Drive1_Change()
    On Error Resume Next
    Label6.Caption = ""
    Label3.Caption = ""
    ChDir Left(Drive1.Drive, 2) & "\"
    File1.FileName = Combo1.Text
    Dir1.Path = Left(Drive1.Drive, 2) & "\"
    StatusBar1.Panels(2).Text = UCase(Drive1.Drive)
End Sub
Private Sub exit_Click()
    End
End Sub
Private Sub File1_Click()
    On Error Resume Next
    MediaPlayer1.AutoStart = False
    MediaPlayer1.FileName = Dir1Path & File1.List(File1.ListIndex)
    Label3.Caption = Left(FileLen(Dir1Path & File1.List(File1.ListIndex)) / 1024 ^ 2, 4) & " MBytes"
    Label6.Caption = File1.List(File1.ListIndex)
End Sub
Private Sub File1_DblClick()
    On Error Resume Next
    MediaPlayer1.AutoStart = True
    MediaPlayer1.FileName = Dir1Path & File1.List(File1.ListIndex)
End Sub
Private Sub Form_Initialize()
    If Len(GetSetting("MP3 Library", "General", "Ext")) = 0 Then
        Combo1.Text = "*.mp3"
    Else
        Combo1.Text = GetSetting("MP3 Library", "General", "Ext")
    End If
    If Len(GetSetting("MP3 Library", "General", "Dir1.Path")) = 0 Then
        Dir1.Path = "c:\"
    Else
        Drive1.Drive = Left(GetSetting("MP3 Library", "General", "Dir1.Path"), 3)
        Dir1.Path = GetSetting("MP3 Library", "General", "Dir1.Path")
    End If
    Combo1.AddItem "*.mp3"
    Combo1.AddItem "*.wav"
    Combo1.AddItem "*.ram"
    Combo1.AddItem "*.ra"
    Combo1.AddItem "*.mid"
    Combo1.AddItem "*.mp2"
    Combo1.AddItem "*.*"
    Frame1.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
    Frame8.Visible = False
    MSFlexGrid1.ColWidth(0) = 600
    MSFlexGrid1.ColWidth(1) = 2360
    MSFlexGrid1.ColWidth(2) = 2360
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Text = "ID"
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Text = "Title"
    MSFlexGrid1.Col = 2
    MSFlexGrid1.Text = "Artist"
    MSFlexGrid2.ColWidth(0) = 700
    MSFlexGrid2.ColWidth(1) = 2630
    MSFlexGrid2.ColWidth(2) = 2630
    MSFlexGrid2.ColWidth(3) = 1400
    MSFlexGrid2.Row = 0
    MSFlexGrid2.Col = 0
    MSFlexGrid2.Text = "ID"
    MSFlexGrid2.Col = 1
    MSFlexGrid2.Text = "Record Match"
    MSFlexGrid2.Col = 2
    MSFlexGrid2.Text = "Title"
    MSFlexGrid2.Col = 3
    MSFlexGrid2.Text = "Found in"
    Combo2.AddItem "Title"
    Combo2.AddItem "Artist"
    Combo2.AddItem "Album"
    Combo2.AddItem "Year"
    Combo2.AddItem "Comment"
    Combo2.AddItem "FileName"
    Combo2.Text = "Title"
End Sub
'Here Ends The File Manager Tab Codes !!!!!
'Here Starts General Code
Private Sub Form_Load()
    'On Error Resume Next
    Dim TempVar As Long
    If App.PrevInstance = True Then End
    'Me.Visible = True
    DBL = False
    File2.Path = App.Path & "\DataBase\"
    File2.FileName = "*.mdb"
    Label31.Caption = File2.ListCount
    StatusBar1.Panels(1).Text = TabStrip1.SelectedItem.ToolTipText
    TabStrip1.ZOrder 1
    File1.FileName = Combo1.Text
    Label4.Caption = File1.ListCount
    Label31.Caption = File2.ListCount
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "MP3 Library", "General", "Dir1.Path", MainForm.Dir1.Path
    SaveSetting "MP3 Library", "General", "Ext", MainForm.Combo1.Text
    End
End Sub

Private Sub load_Click()
    TabStrip1.Tabs(2).Selected = True
    DBCommand(1).Value = True
End Sub
Private Sub MainBut_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case Is = 0
        If File2.ListIndex = -1 Then Exit Sub
        If CheckRS = False Then Exit Sub
        Label5.Caption = ""
        Label7.Caption = ""
        Label10.Caption = ""
        Label12.Caption = ""
        StatusBar1.Panels(2).Text = ""
        OpenDBFile App.Path & "\DataBase\" & File2.List(File2.ListIndex), File2.List(File2.ListIndex), RSet1, DBase1
        TabStrip1.Tabs(2).Selected = True
        ReadRS
    Case Is = 1
        RSet1.close
        RSet2.close
        DBase1.close
        Unload Me
    End Select
End Sub
Private Sub MSFlexGrid1_Click()
    Dim L5Cap As String
    Label5.Caption = ""
    L5Cap = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
    L5Cap = L5Cap & " - "
    L5Cap = L5Cap & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
    Label5.Caption = L5Cap
End Sub
Private Sub MSFlexGrid1_DblClick()
    MSFlexGrid1.Col = 0
    If Len(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) > 0 Then
        EditItem (CLng(MSFlexGrid1.Text))
    End If
End Sub
Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MSFlexGrid1.ColWidth(0) = 600
    MSFlexGrid1.ColWidth(1) = 2360
    MSFlexGrid1.ColWidth(2) = 2360
End Sub
Private Sub new_Click()
    DBCommand(0).Value = True
    TabStrip1.Tabs(2).Selected = True
End Sub
Private Sub OptionM_Click()
    On Error Resume Next
    MsgBox "No options available at this time", vbOKOnly + vbInformation
    Exit Sub
    CntrlFrm.Height = 3855
    CntrlFrm.Text5.Text = GetSetting("MP3 Library", "CDROM", "CD")
    CntrlFrm.Frame7.Visible = True
    CntrlFrm.Frame7.ZOrder 0
    CntrlFrm.Command14.Default = True
    CntrlFrm.Show vbModal
End Sub
Private Sub RASlct_Click(Index As Integer)
    Dim I As Long, J As Long
    'On Error Resume Next
    Select Case Index
    Case Is = 0
        TabStrip1.Tabs(3).Selected = True
    Case Is = 1
        MSFlexGrid1.Col = 0
        EditItem (CLng(MSFlexGrid1.Text))
    Case Is = 2
        If MsgBox("Are you sure you want to delete the " & MSFlexGrid1.RowSel - MSFlexGrid1.Row + 1 & " selected records permanently.", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        StatusBar1.Panels(2).Text = "Deleting..."
        Me.MousePointer = vbHourglass
        ProgressBar1.Min = 0
        ProgressBar1.Max = MSFlexGrid1.RowSel + 1 - MSFlexGrid1.Row
        For I = MSFlexGrid1.Row To MSFlexGrid1.RowSel
            'ProgressBar1.Value = I
            RSet1.Filter = "ID=" & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
            RSet1.delete adAffectCurrent
            If MSFlexGrid1.Rows > 2 Then
                MSFlexGrid1.RemoveItem MSFlexGrid1.Row
            Else
                For J = 0 To 2
                    MSFlexGrid1.Col = J
                    MSFlexGrid1.Text = ""
                Next J
            End If
        Next I
        RSet1.Filter = adFilterNone
        StatusBar1.Panels(2).Text = "Done."
        Me.MousePointer = vbDefault
        ProgressBar1.Value = 0
    Case Is = 3
        MSFlexGrid1.Row = 1
        MSFlexGrid1.Col = 0
        MSFlexGrid1.RowSel = MSFlexGrid1.Rows - 1
        MSFlexGrid1.ColSel = 2
    End Select
End Sub
Private Sub RASlct_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
    Case Is = 0
        StatusBar1.Panels(2).Text = "Add Record to Current Database"
    Case Is = 1
        StatusBar1.Panels(2).Text = "Edit Selected Record"
    Case Is = 2
        StatusBar1.Panels(2).Text = "Remove Select Record(s)"
    Case Is = 3
        StatusBar1.Panels(2).Text = "Select All Records in Current Database"
    End Select
End Sub
Private Sub SExp_Click(Index As Integer)
    Dim I As Long
    CntrlFrm.Frame9.Visible = False
    Select Case Index
    Case Is = 0
        Me.MousePointer = vbHourglass
        For I = 0 To File1.ListCount - 1
            File1.Selected(I) = True
        Next I
        Me.MousePointer = vbDefault
    Case Is = 1
        Me.MousePointer = vbHourglass
        For I = 0 To File1.ListCount - 1
            File1.Selected(I) = False
        Next I
        Me.MousePointer = vbDefault
    Case Is = 2
        MediaPlayer1.FileName = Dir1Path
        CntrlFrm.Height = 3855
        CntrlFrm.Text2.Text = Dir1Path
        For I = 0 To File1.ListCount - 1
            If File1.Selected(I) = True Then
                CntrlFrm.List2.AddItem Dir1Path & File1.List(I)
            End If
        Next I
        CntrlFrm.Frame2.ZOrder 0
        CntrlFrm.Show vbModal
        CntrlFrm.Command5.Default = True
    Case Is = 3
        MediaPlayer1.FileName = Dir1Path
        CntrlFrm.Height = 3855
        CntrlFrm.Text3.Text = Dir1Path
        For I = 0 To File1.ListCount - 1
            If File1.Selected(I) = True Then
                CntrlFrm.List3.AddItem Dir1Path & File1.List(I)
            End If
        Next I
        CntrlFrm.Frame1.ZOrder 0
        CntrlFrm.Show vbModal
        CntrlFrm.Command7.Default = True
    Case Is = 4
        MediaPlayer1.FileName = Dir1Path
        CntrlFrm.Frame1.Visible = False
        CntrlFrm.Frame3.Visible = False
        CntrlFrm.Frame2.Visible = False
        CntrlFrm.Frame5.Visible = False
        CntrlFrm.Frame6.Visible = False
        CntrlFrm.Frame8.Visible = False
        CntrlFrm.Frame7.Visible = False
        CntrlFrm.Height = 1590
        CntrlFrm.Text1.Text = File1.List(File1.ListIndex)
        CntrlFrm.Frame4.ZOrder 0
        CntrlFrm.Show vbModal
        CntrlFrm.Command3.Default = True
    Case Is = 5
        MediaPlayer1.FileName = Dir1Path
        CntrlFrm.List1.Clear
        CntrlFrm.Height = 3855
        CntrlFrm.Frame3.Visible = True
        For I = 0 To File1.ListCount - 1
            If File1.Selected(I) = True Then CntrlFrm.List1.AddItem Dir1Path & File1.List(I)
        Next I
        CntrlFrm.Label1.Caption = "Are You Sure, You Want To Delete The " & CntrlFrm.List1.ListCount & " Listed Files"
        CntrlFrm.Command1.Default = True
        CntrlFrm.Frame3.ZOrder 0
        CntrlFrm.Show vbModal
    Case Is = 6
        FS = False
        If RSet1.State = adStateClosed Then
            MsgBox "No DataBase Found, Please Load or Create a New DataBase", vbInformation + vbOKOnly, "MP3 Library"
            Exit Sub
        End If
        If File1.ListCount = 0 Then
            MsgBox "No Available Files Found.", vbOKOnly + vbInformation, "MP3 Library"
            Exit Sub
        Else
            For I = 0 To File1.ListCount - 1
                If File1.Selected(I) = True Then
                    FS = True
                 End If
            Next I
            If FS = True Then
                WaitF.Show 'vbModal
                Me.MousePointer = vbHourglass
                AddRec (False)
            Else
                If MsgBox("No Files Are Selected, Do You Want To Append All Files.", vbYesNo + vbQuestion, "MP3 Library") = vbYes Then
                    WaitF.Show 'vbModal
                    Me.MousePointer = vbHourglass
                    AddRec (True)
                End If
            End If
        End If
        If Len(MSFlexGrid1.TextMatrix(1, 0)) = 0 And MSFlexGrid1.Rows > 2 Then MSFlexGrid1.RemoveItem 1
        Me.MousePointer = vbDefault
    Case Is = 7
        CntrlFrm.Frame1.Visible = False
        CntrlFrm.Frame3.Visible = False
        CntrlFrm.Frame2.Visible = False
        CntrlFrm.Frame5.Visible = False
        Artist = ""
        Album = ""
        FC = 0
        FS = 0
        CntrlFrm.Height = 3855
        If File1.ListCount = 0 Then
            MsgBox "No files in the list to select.", vbExclamation + vbOKOnly, "MP3 Library Error"
            Exit Sub
        End If
        For I = 0 To File1.ListCount - 1
            If File1.Selected(I) = True Then
                If InStr(1, FT, LCase(Right(File1.List(I), 3)), vbTextCompare) = 0 Then
                    FT = FT & " - " & LCase(Right(File1.List(I), 3))
                End If
                FC = FC + 1
                FS = FS + FileLen(Dir1Path & File1.List(I))
            End If
        Next I
        FT = Right(FT, Len(FT) - 2)
        If FC = 0 Then
            MsgBox "Please Select at Least One File.", vbExclamation + vbOKOnly, "MP3 Library Error"
            Exit Sub
        ElseIf FC = 1 Then
            Open Dir1Path & File1.List(File1.ListIndex) For Input As #1
            Seek #1, LOF(1) - 95
            Artist = Input(30, #1)
            Album = Input(30, #1)
            Close #1
            If Left(Artist, 1) = Chr(255) Then Artist = ""
            If Left(Album, 1) = Chr(255) Then Album = ""
            CntrlFrm.Label12.Visible = True
            CntrlFrm.Label11.Visible = True
            CntrlFrm.Label2.Caption = "File Name:"
            CntrlFrm.Label3.Caption = "File size:"
            CntrlFrm.Label6.Caption = File1.List(File1.ListIndex)
            CntrlFrm.Label7.Caption = Right(FT, Len(FT) - 1) & " File."
            CntrlFrm.Label8.Caption = UCase(Dir1.Path)
            CntrlFrm.Label9.Caption = Left((FS / 1024 ^ 2), 4) & " MBytes"
            CntrlFrm.Label10.Caption = FS
            If Trim(LCase(Right(File1.List(File1.ListIndex), 3))) = "mp3" Then
                CntrlFrm.Label13.Visible = True
                CntrlFrm.Label14.Visible = True
                CntrlFrm.Label12.Visible = True
                CntrlFrm.Label11.Visible = True
                CntrlFrm.Label13.Caption = Artist
                CntrlFrm.Label14.Caption = Album
            Else
                CntrlFrm.Label13.Visible = False
                CntrlFrm.Label14.Visible = False
                CntrlFrm.Label12.Visible = False
                CntrlFrm.Label11.Visible = False
            End If
        ElseIf FC > 1 Then
            CntrlFrm.Label12.Visible = False
            CntrlFrm.Label11.Visible = False
            CntrlFrm.Label2.Caption = "Total Files:"
            CntrlFrm.Label3.Caption = "Total Size:"
            CntrlFrm.Label6.Caption = FC
            CntrlFrm.Label7.Caption = Right(FT, Len(FT) - 1)
            CntrlFrm.Label8.Caption = UCase(Dir1.Path)
            CntrlFrm.Label9.Caption = Left((FS / 1024 ^ 2), 4) & " MBytes"
            CntrlFrm.Label10.Caption = FS
        End If
        CntrlFrm.Frame6.ZOrder 0
        CntrlFrm.Show vbModal
    End Select
End Sub
Private Sub TabStrip1_Click()
    On Error Resume Next
    StatusBar1.Panels(1).Text = TabStrip1.SelectedItem.ToolTipText
    If TabStrip1.SelectedItem.Index = 1 Then
        Frame1.Visible = True
        Frame2.Visible = False
        Frame3.Visible = False
        Frame8.Visible = False
        Label31.Caption = File2.ListCount
        File2.Refresh
        StatusBar1.Panels(2).Text = ""
    ElseIf TabStrip1.SelectedItem.Index = 2 Then
        Frame1.Visible = False
        Frame2.Visible = True
        Frame3.Visible = False
        Frame8.Visible = False
        Label7.Caption = RSet1.RecordCount
        StatusBar1.Panels(2).Text = Label10.Caption
    ElseIf TabStrip1.SelectedItem.Index = 3 Then
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = True
        Frame8.Visible = False
        StatusBar1.Panels(2).Text = UCase(Drive1.Drive)
     ElseIf TabStrip1.SelectedItem.Index = 4 Then
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = False
        Frame8.Visible = True
        If RSet1.State = adStateClosed Then
            Option1.Enabled = False
            Option2.Value = True
        Else
            Option1.Enabled = True
            Option1.Value = True
        End If
        File3.Path = Left(File1.Path, 3)
        Wait 0.1
        File3.Path = App.Path & "\DataBase"
        File3.FileName = "*.mdb"
        Command26.Default = True
        StatusBar1.Panels(2).Text = ""
    End If
End Sub
'Here Ends The General Code
'Here Starts The Database Manager Tab Code
'Here Ends The DataBase Manager Tab Codes !!!!!
'Here Starts The General Page Tab Codes !!!!
Private Sub File2_DblClick()
    If RSet1.State = adStateOpen Then
        If MsgBox("The DataBase " & Label10.Caption & " Is Still Loaded" & vbNewLine & "Do You Want To Close Current Database.", vbQuestion + vbYesNo) = vbYes Then
            CloseDB
        Else
            Exit Sub
        End If
    End If
    MainBut(0).Value = True
End Sub
Private Sub File2_Click()
    'On Error GoTo errdone
    Me.MousePointer = vbHourglass
        List5.Clear
        Label28.Caption = ""
        Label29.Caption = ""
        Label33.Caption = ""
        ProgressBar2.Value = 0
    If Label28.Caption = File2.List(File2.ListIndex) Then Exit Sub
    StatusBar1.Panels(2).Text = App.Path & "\DataBase\" & File2.List(File2.ListIndex)
    Label28.Caption = File2.List(File2.ListIndex)
    Label29.Caption = FileLen(App.Path & "\DataBase\" & File2.List(File2.ListIndex)) & " Bytes"
    
    If Check1.Value = 1 Then
        ProgressBar2.Min = 0
        Me.MousePointer = vbHourglass
        If RSet2.State = adStateOpen Then CloseDB
        OpenDBFile AppPath & "Database\" & File2.List(File2.ListIndex), File2.List(File2.ListIndex), RSet2, DBase2
        If RSet2.RecordCount = 0 Then
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        ProgressBar2.Max = RSet2.RecordCount
        RSet2.MoveFirst
        Do While Not RSet2.EOF
            DoEvents
            List5.AddItem RSet2.Fields("Title")
            ProgressBar2.Value = RSet2.AbsolutePosition
            RSet2.MoveNext
        Loop
        RSet2.close
        DBase2.close
    End If
    Me.MousePointer = vbDefault
    Exit Sub
errdone:
    MsgBox Error, vbCritical
    Me.MousePointer = vbDefault
    Close #1
    Exit Sub
End Sub
'Here Ends The General Page Tab Codes !!!!!
'Here Starts The Find Tab Codes !!!!
Private Sub Command26_Click()
    On Error Resume Next
    Dim STS As String, I As Long
    Command26.Enabled = False
    Me.MousePointer = vbHourglass
    '======================
    'clear previous results
    '======================
    MSFlexGrid2.AddItem " "
    For I = 2 To MSFlexGrid2.Rows
        MSFlexGrid2.RemoveItem 1
    Next I
    '=================
    'end clear results
    '=================
    If Option1.Value = True Then
        OpenDBFile AppPath & "Database\" & Label10.Caption, File2.List(File2.ListIndex), RSet2, DBase2
        SearchDB (Label10.Caption)
    ElseIf Option2.Value = True Then
        For I = 0 To File3.ListCount - 1
            OpenDBFile AppPath & "Database\" & File3.List(I), File2.List(File2.ListIndex), RSet2, DBase2
            SearchDB (File3.List(I))
        Next I
    ElseIf Option3.Value = True Then
        For I = 0 To File3.ListCount - 1
            If File3.Selected(I) = True Then
                OpenDBFile AppPath & "Database\" & File3.List(I), File2.List(File2.ListIndex), RSet2, DBase2
                SearchDB (File3.List(I))
            End If
        Next I
    End If
    RSet2.close
    DBase2.close
    Command26.Enabled = True
    Me.MousePointer = vbDefault
End Sub
Private Function CheckRS(Optional RSOrder As Long) As Boolean
    If RSet1.State = adStateOpen Then
        If MsgBox("The DataBase " & Label10.Caption & " Is Still Loaded" & vbNewLine & "Do You Want To Close Current Database.", vbQuestion + vbYesNo) = vbYes Then
            CloseDB
            CheckRS = True
        Else
            CheckRS = False
            Exit Function
        End If
    Else
        CheckRS = True
    End If
End Function
Sub ReadRS()
    Dim TCol1 As String, TCol2 As String, TCol3 As String, TCell
    TCell = 0
    If RSet1.RecordCount = 0 Then Exit Sub
    Me.MousePointer = vbHourglass
    ProgressBar1.Min = 0
    ProgressBar1.Max = RSet1.RecordCount
    RSet1.MoveFirst
    MSFlexGrid1.Redraw = False
    Do While Not RSet1.EOF
        DoEvents
        ProgressBar1.Value = RSet1.AbsolutePosition
        TCol1 = " "
        TCol2 = " "
        TCol3 = " "
        If IsNull(RSet1.Fields("ID")) = False Then TCol1 = RSet1.Fields("ID")
        If IsNull(RSet1.Fields("Title")) = False Then TCol2 = RSet1.Fields("Title")
        If IsNull(RSet1.Fields("Artist")) = False Then TCol3 = RSet1.Fields("Artist")
        TCell = TCell + 1
        MSFlexGrid1.AddItem TCol1 & vbTab & TCol2 & vbTab & TCol3
        RSet1.MoveNext
    Loop
    MSFlexGrid1.RemoveItem 1
    MSFlexGrid1.Redraw = True
    Me.MousePointer = vbDefault
End Sub
Private Sub CloseDB()
    Dim rc As Long
    If RSet1.State = adStateOpen Then
        RSet1.close
        DBase1.close
    End If
    MSFlexGrid1.AddItem ""
    For rc = 1 To MSFlexGrid1.Rows - 2
        MSFlexGrid1.RemoveItem 1
    Next rc
    Label5.Caption = ""
    Label7.Caption = ""
    Label10.Caption = ""
    Label12.Caption = ""
    Label28.Caption = ""
    Label29.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    ProgressBar1.Value = 0
    StatusBar1.Panels(2).Text = ""
End Sub
