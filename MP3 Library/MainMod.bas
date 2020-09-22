Attribute VB_Name = "MainMod"
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal Hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal Hkey As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public FC As Integer, FS As Long, FT As String
Public Artist As String, Album As String, DBL As Boolean, TmpStr As String
Public DBase1 As ADODB.Connection, RSet1 As ADODB.Recordset, RSet2 As ADODB.Recordset, DBase2 As ADODB.Connection
Sub Main()
    Set DBase1 = New ADODB.Connection
    Set DBase2 = New ADODB.Connection
    Set RSet1 = New ADODB.Recordset
    Set RSet2 = New ADODB.Recordset
    DBase1.Mode = 3
    DBase2.Mode = 3
    'SplashFrm.Show
    load MainForm
    MainForm.Show
End Sub
Public Function Dir1Path() As String
    If Right(MainForm.Dir1.Path, 1) = "\" Then
        Dir1Path = MainForm.Dir1.Path
    Else
        Dir1Path = MainForm.Dir1.Path & "\"
    End If
End Function
Sub Wait(TimeToWait As Variant)
Dim RetConst1 As Variant
Dim RetConst2 As Integer
Dim RetConst3 As Variant
RetConst1 = GetTickCount()
Do
    RetConst2% = DoEvents()
    RetConst3 = GetTickCount()
Loop Until RetConst3 - RetConst1 >= TimeToWait * 1000
End Sub
Public Function AppPath() As String
    AppPath = App.Path
    If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\"
End Function
Public Sub OpenDBFile(DBFPath As String, DBFName As String, RSSlct As ADODB.Recordset, DBSlct As ADODB.Connection)
    DBSlct.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & DBFPath
    RSSlct.Open "SELECT * FROM List ORDER BY Title", DBSlct, 3, 3
    MainForm.Label7.Caption = ""
    MainForm.Label33.Caption = ""
    MainForm.Label10.Caption = ""
    MainForm.Label12.Caption = ""
    MainForm.Label12.Caption = FileLen(DBFPath) & " Bytes"
    MainForm.Label7.Caption = RSSlct.RecordCount & " Records"
    MainForm.Label10.Caption = DBFName
    MainForm.Label33.Caption = RSSlct.RecordCount & " Records"
    MainForm.StatusBar1.Panels(2).Text = DBFName
End Sub
