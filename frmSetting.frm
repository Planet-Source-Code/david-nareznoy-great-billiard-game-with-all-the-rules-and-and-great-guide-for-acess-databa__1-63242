VERSION 5.00
Begin VB.Form frmSetting 
   BackColor       =   &H80000001&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setting"
   ClientHeight    =   3660
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5565
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra 
      BackColor       =   &H80000001&
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5295
      Begin VB.PictureBox Picture1 
         Height          =   1155
         Left            =   120
         Picture         =   "frmSetting.frx":030A
         ScaleHeight     =   1095
         ScaleWidth      =   3900
         TabIndex        =   5
         Top             =   1560
         Width           =   3960
      End
      Begin VB.CheckBox RunProgramAtStartUp 
         BackColor       =   &H80000003&
         Caption         =   "Run program at startup. Make exe named 'BilliarD.exe'."
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   4905
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   350
      Left            =   1440
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   350
      HelpContextID   =   13
      Left            =   4080
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   2760
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
 
  Unload Me
End Sub

Private Sub cmdApply_Click()
  SaveDataToSetting
  GetDataFromSetting
  cmdApply.Enabled = False
End Sub

Private Sub cmdOK_Click()
  SaveDataToSetting
  GetDataFromSetting
  Unload Me
End Sub

Private Sub Form_Load()
  GetDataFromSetting
  DoEvents
  cmdApply.Enabled = False
End Sub



Private Sub RunProgramAtStartUp_Click()
  If RunProgramAtStartUp.Value = 1 Then
     SetRegValue HKEY_LOCAL_MACHINE, _
     "Software\Microsoft\Windows\CurrentVersion\Run", "LOGIN", App.Path & "\LOGIN.exe"
     cmdApply.Enabled = True
     Exit Sub
  End If
  If RunProgramAtStartUp.Value = 0 Then
     DeleteValue HKEY_LOCAL_MACHINE, _
     "Software\Microsoft\Windows\CurrentVersion\Run", "LOGIN"
     cmdApply.Enabled = True
     Exit Sub
  End If
End Sub


Private Sub SaveDataToSetting()
  Call SaveFromControlsToINI(frmSetting, "Setting")
End Sub

Private Sub GetDataFromSetting()
  Call ReadFromINIToControls(frmSetting, "Setting")
End Sub

