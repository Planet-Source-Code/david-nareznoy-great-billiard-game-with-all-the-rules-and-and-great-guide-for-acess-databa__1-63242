Attribute VB_Name = "modGeneral"
'File Name   : Module1.bas
'-----------------------------------------------------

Public Type arrSetting
  LindungLayar As Integer
  PassLindungLayar As String
  MenitDelay As String
  IntervalMenit As String
End Type
Public gloSet As arrSetting

Public db As ADODB.Connection
Public m_UserID As String
Public m_Level As String
Public m_blnLogin As Boolean
Public m_blnCancel As Boolean
Public strPassword As String, Temp As String
Public rsUser As ADODB.Recordset
Public StaKonek As Boolean
Public LindungLayar As Byte
Public Awal As Date
Public Gerak As Boolean
Public Aksi As Boolean
Public IfCon As Boolean


Public Type tUser
  UserID As String
  password As String
End Type
Public tabUser() As tUser

Public Type cUser
  User As String
End Type
Public cekUser() As cUser


'MSAccess97 password protected. The password is 'masino2002'
Public Sub OpenConnection()
On Error GoTo ErrMess
  Screen.MousePointer = vbHourglass
  Set db = New Connection

  db.CursorLocation = adUseClient

  'password: masino2002
  db.Open "PROVIDER=MSDataShape;Data PROVIDER=" & _
          "Microsoft.Jet.OLEDB.4.0;Data Source=" _
          & App.Path & "\Data.mdb;Jet OLEDB:" & _
          "Database Password=masino2002;"
  StaKonek = True
  Screen.MousePointer = vbDefault
  Exit Sub
ErrMess:
  StaKonek = False
  Screen.MousePointer = vbDefault
End Sub


Public Sub OpenTableUser()
On Error GoTo MessErr
  If db Is Nothing Then OpenConnection
  Set rsUser = New ADODB.Recordset
  DoEvents

  rsUser.Open _
      "SHAPE {SELECT * FROM T_User " & _
      "Order by User_ID} AS ParentCMD APPEND " & _
      "({SELECT * FROM T_User " & vbCrLf & _
      "Order by User_ID } AS ChildCMD RELATE User_ID TO " & _
      "User_ID) AS ChildCMD", _
      db, adOpenStatic, adLockOptimistic
  DoEvents
  Exit Sub
MessErr:
  Select Case Err.Number
         Case 3709
              MsgBox "Failed to connect to database!", _
                     vbExclamation, "Failed"
              If FormLoadedByName("frmLogin") = True Then
                 Unload frmLogin
              End If
         Case Else
              MsgBox Err.Number & " - " & _
                     Err.Description, _
                     vbCritical, "Error"
  End Select
End Sub


Public Function NumOfUser() As Integer
   NumOfUser = rsUser.RecordCount
End Function

'This will encrypt/decrypt password field in database
Public Sub EncryptDecrypt()
Dim i As Integer
Dim intLocation As Integer
Dim Code As String
Code = "1234567890" 'This is key for encrypting/decrypting
  Temp$ = ""
  For i% = 1 To Len(strPassword)
      intLocation% = (i% Mod Len(Code)) + 1
      'Use XOR logic combination for encrypting/decrypting
      Temp$ = Temp$ + Chr$(Asc(Mid$(strPassword, i%, 1)) Xor _
              Asc(Mid$(Code, intLocation%, 1)))
  Next i%
End Sub


Public Function FormLoadedByName(FormName As String) As Boolean
Dim i As Integer, fnamelc As String
fnamelc = LCase$(FormName)
FormLoadedByName = False
For i = 0 To Forms.Count - 1
If LCase$(Forms(i).Name) = fnamelc Then
  FormLoadedByName = True
  Exit Function
End If
Next
End Function

'Close all forms in this project
Public Sub CloseAllForms()
Dim Form As Form
   For Each Form In Forms
       Unload Form
       Set Form = Nothing
   Next Form
   End
End Sub

'Begin to wait a process...
Public Sub StartWaiting(strMess As String)
  Screen.MousePointer = vbHourglass
  DoEvents
  With frmWait
    DoEvents
    .lblProses.Move 120, .prgBar1.Top
    .lblProses = strMess
    DoEvents
    .prgBar1.Visible = False
    .lblAngka.Visible = False
    .Show , FrmMain
  End With
End Sub


Public Sub FinishWaiting()
  Unload frmWait
  Set frmWait = Nothing
  Screen.MousePointer = vbDefault
End Sub

Public Sub UnloadAllExceptOne(FormToStay As String)
Dim oFrm As Form
For Each oFrm In Forms
    If oFrm.Name <> FormToStay And Not _
       (TypeOf oFrm Is MDIForm) Then
       Unload oFrm
       Set oFrm = Nothing
    End If
Next
End Sub

'Disable Ctrl-Alt-Del
Public Sub DisableCtrlAltDelete(bDisabled As Boolean)
Dim X As Long
    X = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Sub


