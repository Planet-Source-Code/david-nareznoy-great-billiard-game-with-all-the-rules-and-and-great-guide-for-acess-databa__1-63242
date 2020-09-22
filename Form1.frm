VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Billiards"
   ClientHeight    =   11520
   ClientLeft      =   1545
   ClientTop       =   435
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Prof2 
      BackColor       =   &H80000008&
      Caption         =   "Player2"
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   11520
      TabIndex        =   26
      Top             =   6720
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Frame Prof1 
      BackColor       =   &H80000008&
      Caption         =   "Payer1"
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   11520
      TabIndex        =   25
      Top             =   2760
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox AdminP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DataField       =   "David"
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   840
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1.938
      ScaleLeft       =   8
      ScaleMode       =   0  'User
      ScaleTop        =   8
      ScaleWidth      =   25
      TabIndex        =   23
      ToolTipText     =   "Exit"
      Top             =   9840
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Label Adm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Admin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   495
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DataField       =   "David"
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   840
      Picture         =   "Form1.frx":19C52
      ScaleHeight     =   1.938
      ScaleLeft       =   8
      ScaleMode       =   0  'User
      ScaleTop        =   8
      ScaleWidth      =   25
      TabIndex        =   21
      ToolTipText     =   "Exit"
      Top             =   7920
      Width           =   2775
      Begin VB.Label Con 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   495
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DataField       =   "David"
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   840
      Picture         =   "Form1.frx":338A4
      ScaleHeight     =   1.938
      ScaleLeft       =   8
      ScaleMode       =   0  'User
      ScaleTop        =   8
      ScaleWidth      =   25
      TabIndex        =   16
      ToolTipText     =   "Exit"
      Top             =   8880
      Width           =   2775
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   495
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DataField       =   "David"
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   840
      Picture         =   "Form1.frx":4D4F6
      ScaleHeight     =   1.938
      ScaleLeft       =   8
      ScaleMode       =   0  'User
      ScaleTop        =   8
      ScaleWidth      =   25
      TabIndex        =   15
      ToolTipText     =   "Exit"
      Top             =   7440
      Width           =   2775
      Begin VB.Label JG 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New Game"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   495
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DataField       =   "David"
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   840
      Picture         =   "Form1.frx":67148
      ScaleHeight     =   1.938
      ScaleLeft       =   8
      ScaleMode       =   0  'User
      ScaleTop        =   8
      ScaleWidth      =   25
      TabIndex        =   14
      ToolTipText     =   "Exit"
      Top             =   8400
      Width           =   2775
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   495
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DataField       =   "David"
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   840
      Picture         =   "Form1.frx":80D9A
      ScaleHeight     =   1.938
      ScaleLeft       =   8
      ScaleMode       =   0  'User
      ScaleTop        =   8
      ScaleWidth      =   25
      TabIndex        =   13
      ToolTipText     =   "Exit"
      Top             =   9360
      Width           =   2775
      Begin VB.Label ExitG 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   495
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   2775
      End
   End
   Begin MSComctlLib.ProgressBar ForceBar 
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   3000
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Min             =   10
      Max             =   250
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Status"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   3975
      Begin VB.Label lblStatus 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Timer CheckTurns 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10800
      Top             =   6360
   End
   Begin VB.Frame frmOp2 
      BackColor       =   &H8000000A&
      Caption         =   "Player 2"
      Height          =   975
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   1935
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ball type"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Shape shpOp2Ball 
         FillColor       =   &H8000000F&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Balls in"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblOp2Counter 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.Frame frmOp1 
      BackColor       =   &H00FF8080&
      Caption         =   "Player 1"
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1935
      Begin VB.Label lblOp1Counter 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Balls in"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Shape shpOp1Ball 
         FillColor       =   &H8000000F&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ball type"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Timer BallsMove 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   10800
      Top             =   5880
   End
   Begin VB.Timer cueBallMoves 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   11280
      Top             =   5880
   End
   Begin VB.CommandButton cmdShoot 
      Caption         =   "Command1"
      Height          =   195
      Left            =   720
      TabIndex        =   12
      Top             =   1560
      Width           =   75
   End
   Begin VB.Image Image28 
      Height          =   555
      Left            =   7800
      Picture         =   "Form1.frx":9A9EC
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image27 
      Height          =   555
      Left            =   6000
      Picture         =   "Form1.frx":9B9CA
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image26 
      DragMode        =   1  'Automatic
      Height          =   555
      Left            =   6600
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      Picture         =   "Form1.frx":9C9A8
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image25 
      Height          =   555
      Left            =   0
      Picture         =   "Form1.frx":9D986
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image24 
      Height          =   555
      Left            =   600
      Picture         =   "Form1.frx":9E964
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image23 
      Height          =   555
      Left            =   1200
      Picture         =   "Form1.frx":9F942
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image22 
      Height          =   555
      Left            =   4200
      Picture         =   "Form1.frx":A0920
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image21 
      Height          =   555
      Left            =   5400
      Picture         =   "Form1.frx":A18FE
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image20 
      Height          =   555
      Left            =   10200
      Picture         =   "Form1.frx":A28DC
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image19 
      Height          =   555
      Left            =   10800
      Picture         =   "Form1.frx":A38BA
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image18 
      Height          =   555
      Left            =   3000
      Picture         =   "Form1.frx":A4898
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image17 
      Height          =   555
      Left            =   9000
      Picture         =   "Form1.frx":A5876
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image15 
      Height          =   555
      Left            =   7200
      Picture         =   "Form1.frx":A6854
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image14 
      Height          =   555
      Left            =   3600
      Picture         =   "Form1.frx":A7832
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image12 
      Height          =   555
      Left            =   1800
      Picture         =   "Form1.frx":A8810
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image11 
      Height          =   555
      Left            =   8400
      Picture         =   "Form1.frx":A97EE
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image10 
      Height          =   555
      Left            =   9600
      Picture         =   "Form1.frx":AA7CC
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image6 
      Height          =   555
      Left            =   2400
      Picture         =   "Form1.frx":AB7AA
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image5 
      Height          =   555
      Left            =   600
      Picture         =   "Form1.frx":AC788
      Top             =   1800
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image4 
      Height          =   555
      Left            =   1800
      Picture         =   "Form1.frx":AD766
      Top             =   1800
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image3 
      Height          =   555
      Left            =   1200
      Picture         =   "Form1.frx":AE744
      Top             =   1800
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   4800
      Picture         =   "Form1.frx":AF722
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000004&
      BorderStyle     =   3  'Dot
      X1              =   11160
      X2              =   11160
      Y1              =   1740
      Y2              =   840
   End
   Begin VB.Image Image1 
      Height          =   2445
      Left            =   600
      Picture         =   "Form1.frx":B0700
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   3645
   End
   Begin VB.Shape Ball 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Power :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Shape cueball 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   7080
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   255
   End
   Begin VB.Line cue 
      BorderColor     =   &H00004080&
      BorderWidth     =   4
      X1              =   11160
      X2              =   11160
      Y1              =   2400
      Y2              =   4900
   End
   Begin VB.Shape Ball 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Ball 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   255
   End
   Begin VB.Shape Ball 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00800000&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   9
      Left            =   7800
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Ball 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00800000&
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   7800
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Ball 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00000000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   8
      Left            =   8400
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Ball 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Ball 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      DrawMode        =   1  'Blackness
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   10
      Left            =   7800
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Ball 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00000000&
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape MidRightHole 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   10200
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape MidLeftHole 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape Ball 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00800000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   7920
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Ball 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   255
   End
   Begin VB.Shape DownLeftHole 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   9600
      Width           =   495
   End
   Begin VB.Shape DownRightHole 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   9960
      Shape           =   3  'Circle
      Top             =   9600
      Width           =   495
   End
   Begin VB.Shape UpRightHole 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   9960
      Shape           =   3  'Circle
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape UpLeftHole 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape DownLeft 
      BorderColor     =   &H00008000&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   3975
      Left            =   5400
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape UpLeft 
      BorderColor     =   &H00008000&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   3855
      Left            =   5400
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape DownRight 
      BorderColor     =   &H00008000&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   3975
      Left            =   10200
      Top             =   5640
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   5640
      X2              =   10200
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   7200
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   4575
   End
   Begin VB.Shape UpRight 
      BorderColor     =   &H00008000&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   3855
      Left            =   10200
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Up 
      BorderColor     =   &H00008000&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   5880
      Top             =   840
      Width           =   4095
   End
   Begin VB.Shape Down 
      BorderColor     =   &H00008000&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   5640
      Top             =   9960
      Width           =   4575
   End
   Begin VB.Shape Shape18 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      Height          =   6015
      Left            =   10200
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Table 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   8895
      Left            =   5640
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   735
      Left            =   5400
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape17 
      BorderColor     =   &H00004040&
      FillColor       =   &H00004040&
      FillStyle       =   0  'Solid
      Height          =   9975
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   5775
   End
   Begin VB.Shape BallsStorage 
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   3300
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   420
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a As Integer, b As Integer
Dim NewXCueBall As Integer, NewYCueBall As Integer
Dim NewXball As Integer, NewYball As Integer
Private Type ballgroup
    angle As Double
    force As Integer
    InGame As Boolean
End Type
Dim BallP(11) As ballgroup
Dim StopCue As Boolean, StopBalls As Boolean
Dim redBallsIn As Integer
Dim bluBallsIn As Integer
Dim Turn As Boolean
Dim MoreTurns As Integer
Dim OpState As Integer
Dim BallsHit As Integer
Dim FirstHitMade As Boolean
Dim RelocateCueball As Boolean
Const Cuelen = 2500
Const Linelen = 1230
Const pi = 3.14159265358979


Private Sub Arrange()
    cueball.Left = 7440
    cueball.Top = 7540
    Ball(0).Top = 2335
    Ball(0).Left = 7920
    Ball(1).Top = 1735
    Ball(1).Left = 8340
    Ball(2).Top = 2020
    Ball(2).Left = 7260
    Ball(3).Top = 2335
    Ball(3).Left = 7560
    Ball(4).Top = 1735
    Ball(4).Left = 7545
    Ball(5).Top = 1735
    Ball(5).Left = 7140
    Ball(6).Top = 2620
    Ball(6).Left = 7740
    Ball(7).Top = 2020
    Ball(7).Left = 7740
    Ball(8).Top = 2020
    Ball(8).Left = 8100
    Ball(9).Top = 1735
    Ball(9).Left = 7980
    Ball(10).Top = 1320
    Ball(10).Left = 7740
End Sub

Private Sub NewGame()
    Dim i As Integer
    redBallsIn = 0
    bluBallsIn = 0
    Turn = 1
    OpState = 0
    MoreTurns = 1
    For i = 0 To Ball.Count - 1
        BallP(i).force = 0
        BallP(i).InGame = True
    Next
    Call Arrange
    frmOp2.BackColor = &H8000000A
    frmOp1.BackColor = &HFF8080
    shpOp1Ball.FillColor = &H8000000A
    shpOp2Ball.FillColor = &H8000000A
    lblOp1Counter.Caption = 0
    lblOp1Counter.Caption = 0
    BallsMove.Enabled = False
    cueBallMoves.Enabled = False
    CheckTurns.Enabled = False
    lblStatus.Caption = "Its New Game... Player 1 Turn Now..."
    cmdShoot.Enabled = True
    RelocateCueball = True
    cueball.Visible = True
End Sub

Private Sub EndGame(Winner As Boolean)
    cmdShoot.Enabled = False
    cueBallMoves.Enabled = False
    BallsMove.Enabled = False
    CheckTurns.Enabled = False
    lblStatus.Caption = "Player " & 2 + Winner & " won the game!"
End Sub

Private Sub BallIn(BallNum As Integer)
    Select Case BallNum
        Case 11:
            BallsHit = 3
            cueball.Visible = False
            MoreTurns = 2
            CheckTurns.Enabled = True
            lblStatus.Caption = "And cue ball is in..."
        Case 10: '8ball is in
            If Turn Then
                Select Case OpState
                    Case 0:
                        Call EndGame(0)
                    Case 1:
                        If redBallsIn = 5 Then
                            Call EndGame(1)
                        Else
                            Call EndGame(0)
                        End If
                    Case 2:
                        If bluBallsIn = 5 Then
                            Call EndGame(1)
                        Else
                            Call EndGame(0)
                        End If
                End Select
            Else
                Select Case OpState
                    Case 0:
                        Call EndGame(1)
                    Case 1:
                        If bluBallsIn = 5 Then
                            Call EndGame(0)
                        Else
                            Call EndGame(1)
                        End If
                    Case 2:
                        If bluBallsIn = 5 Then
                            Call EndGame(0)
                        Else
                            Call EndGame(1)
                        End If
                End Select
            End If
        Case Else
            BallP(BallNum).force = 0
            BallP(BallNum).InGame = False
            Ball(BallNum).Left = BallsStorage.Left + (BallsStorage.width - 255) / 2
            Ball(BallNum).Top = BallsStorage.Top + (redBallsIn + bluBallsIn) * 300
            If OpState = 0 Then 'players
                If BallNum < 5 Then
                    If Turn Then
                        OpState = 1
                        shpOp1Ball.FillColor = vbRed
                        shpOp2Ball.FillColor = vbBlue
                    Else
                        OpState = 2
                        shpOp1Ball.FillColor = vbBlue
                        shpOp2Ball.FillColor = vbRed
                    End If
                Else
                    If Turn Then
                        OpState = 2
                        shpOp1Ball.FillColor = vbBlue
                        shpOp2Ball.FillColor = vbRed
                    Else
                        OpState = 1

                        shpOp1Ball.FillColor = vbRed
                        shpOp2Ball.FillColor = vbBlue
                    End If
                End If
            End If
            If BallNum < 5 Then
                If OpState = 1 Then
                    If Turn Then
                        lblStatus.Caption = "Great shot! A red ball is in."
                        MoreTurns = 1
                    Else
                        lblStatus.Caption = "A red ball is in."
                        BallsHit = 2
                    End If
                ElseIf OpState = 2 Then
                    If Not Turn Then
                        lblStatus.Caption = "Great shot! A red ball is in."
                        MoreTurns = 1
                    Else
                        lblStatus.Caption = "A red ball is in."
                        BallsHit = 2
                    End If
                ElseIf OpState = 0 Then
                    lblStatus.Caption = "Great shot! A red ball is in."
                    MoreTurns = 1
                End If
                        
                redBallsIn = redBallsIn + 1
                If redBallsIn = 5 Then
                    If OpState = 1 Then
                        shpOp1Ball.FillColor = vbBlack
                    ElseIf OpState = 2 Then
                        shpOp2Ball.FillColor = vbBlack
                    End If
                End If
            Else
                If OpState = 1 Then
                    If Not Turn Then
                        lblStatus.Caption = "Great shot! A blue ball is in."
                        MoreTurns = 1
                    Else
                        lblStatus.Caption = "A blue ball is in."
                        BallsHit = 2
                    End If
                ElseIf OpState = 2 Then
                    If Turn Then
                        lblStatus.Caption = "Great shot! A blue ball is in."
                        MoreTurns = 1
                    Else
                        lblStatus.Caption = "A blue ball is in."
                        BallsHit = 2
                    End If
                ElseIf OpState = 0 Then
                    lblStatus.Caption = "Great shot! A blue ball is in."
                    MoreTurns = 1
                End If
                bluBallsIn = bluBallsIn + 1
                If bluBallsIn = 5 Then
                    If OpState = 1 Then
                        shpOp2Ball.FillColor = vbBlack
                    ElseIf OpState = 2 Then
                        shpOp1Ball.FillColor = vbBlack
                    End If
                End If
            End If
            If OpState = 1 Then
                lblOp1Counter.Caption = redBallsIn
                lblOp2Counter.Caption = bluBallsIn
            Else
                lblOp1Counter.Caption = bluBallsIn
                lblOp2Counter.Caption = redBallsIn
            End If
    End Select
End Sub

Private Sub SwitchTurns()
    Turn = Not Turn
    If Turn Then
        frmOp2.BackColor = &H8000000A
        frmOp1.BackColor = &HFF8080
    Else
        frmOp1.BackColor = &H8000000A
        frmOp2.BackColor = &HFF8080
    End If
    'By KenShin 'David Nareznoy
    cueBallMoves.Enabled = True
    BallsMove.Enabled = True
    CheckTurns.Enabled = False
    cmdShoot.Enabled = True
    cmdShoot.SetFocus
    cueball.Visible = True
End Sub

Private Sub isIn(BallNum As Integer, newX As Integer, newY As Integer) 'Checking if one of the balls is in
    If newX + 128 > MidRightHole.Left And newY + 128 > MidRightHole.Top And newY < MidRightHole.Top + MidRightHole.Height Then BallIn (BallNum)
    If newX + 128 < MidLeftHole.Left + MidLeftHole.width And newY + 128 > MidLeftHole.Top And newY < MidLeftHole.Top + MidLeftHole.Height Then BallIn (BallNum)
    If newX + 128 > DownRightHole.Left And newY + 128 > DownRightHole.Top Then BallIn (BallNum)
    If newX + 128 < DownLeftHole.Left + DownLeftHole.width And newY + 128 > DownLeftHole.Top Then BallIn (BallNum)
    If newX + 128 < UpLeftHole.Left + UpLeftHole.width And newY + 128 < UpLeftHole.Top + UpLeftHole.Height Then BallIn (BallNum)
    If newX + 128 > UpRightHole.Left And newY + 128 < UpRightHole.Top + UpRightHole.Height Then BallIn (BallNum)
End Sub

Private Sub isOutside(BallNum As Integer, newX As Integer, newY As Integer) 'if a ball is passing borders of table in next location
    If newX < UpLeft.Left + UpLeft.width And newY + 128 > UpLeft.Top And (newY + 128 < UpLeft.Top + UpLeft.Height) Then
        BallP(BallNum).angle = pi - BallP(BallNum).angle
    End If
    If newX < DownLeft.Left + DownLeft.width And newY + 128 > DownLeft.Top And newY + 128 < DownLeft.Top + DownLeft.Height Then
        BallP(BallNum).angle = pi - BallP(BallNum).angle
    End If
    If newX + 255 > UpRight.Left And newY + 128 > UpRight.Top And newY + 128 < UpRight.Top + UpRight.Height Then
        BallP(BallNum).angle = pi - BallP(BallNum).angle
    End If
    If newX + 255 > DownRight.Left And newY + 128 > DownRight.Top And newY + 128 < DownRight.Top + DownRight.Height Then
        BallP(BallNum).angle = pi - BallP(BallNum).angle
    End If
    If newY < Up.Top + Up.Height And newX + 128 > Up.Left And newX + 128 < Up.Left + Up.width Then
        BallP(BallNum).angle = 2 * pi - BallP(BallNum).angle
    End If
    If (newY + 255 > Down.Top) And newX + 128 > Down.Left And newX + 128 < Down.Left + Down.width Then
        BallP(BallNum).angle = 2 * pi - BallP(BallNum).angle
    End If
End Sub

Private Function findAngle(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
    If x1 > x2 And y1 < y2 Then findAngle = pi + Atn((y1 - y2) / (x1 - x2))
    If x1 > x2 And y1 > y2 Then findAngle = pi + Atn((y1 - y2) / (x1 - x2))
    If x1 < x2 And y1 > y2 Then findAngle = Atn((y1 - y2) / (x1 - x2))
    If x1 < x2 And y1 < y2 Then findAngle = Atn((y1 - y2) / (x1 - x2))
End Function


Private Sub Adm_Click()
FrmUser.Show
End Sub


Private Sub CheckTurns_Timer()
    If BallsMove.Enabled = False And cueBallMoves.Enabled = False Then
        If BallsHit = 1 Then
            If MoreTurns = 0 Then
                Call SwitchTurns
                MoreTurns = 1
                BallP(11).force = 0
                CheckTurns.Enabled = False
            End If
        Else
            Call SwitchTurns
            MoreTurns = 2
            CheckTurns.Enabled = False
        End If
        If BallsHit = 3 Then
            cueball.Left = (Line1.x1 + Line1.x2) / 2 - 128
            cueball.Top = (Line1.y1) - 128
            cueball.Visible = True
            RelocateCueball = True
        End If
        cmdShoot.Enabled = True
        cmdShoot.SetFocus
        lblStatus.Caption = "It's Player " & (2 + Turn) & "'s turn"
    End If
End Sub



Private Sub Con_Click()
If IfCon = False Then
frmLogin.Show
Else
  If m_Level = "Admin" Then AdminP.Visible = False
  'close database connection
  If Not db Is Nothing Then
     db.Close
     Set db = Nothing  'Close connection
  End If
  IfCon = False
  m_UserID = ""
  Con.Caption = "Connect"
  Prof1.Visible = False
  frmOp1.Caption = "Player1"
  frmOp1.ForeColor = &H0&
End If
End Sub

Private Sub Con_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Con.ForeColor = &HFFC0C0
End Sub

Public Sub VerifyLevel()
  If m_Level = "Admin" Then
  AdminP.Visible = True
  ElseIf m_Level = "Manager" Then
  ElseIf m_Level = "Operator" Then
  End If
End Sub

Public Sub VerifyUser()
frmOp1.Caption = m_UserID
Prof1.Caption = m_UserID
Prof1.Visible = True
frmOp1.ForeColor = &HFF&
End Sub

Private Sub Form_Load()
    Call NewGame
    ForceBar.Value = 10
    INIFileName = App.Path & "\SettingLogin.ini"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim OK As Boolean
    Dim Yahas As Double
    a = X
    b = Y
    ExitG.ForeColor = &H80000018
    Label8.ForeColor = &H80000018
    Label7.ForeColor = &H80000018
    JG.ForeColor = &H80000018
    Con.ForeColor = &H80000018
    If RelocateCueball Then
        cueball.Visible = True
        If Table.Left + 128 < X And X < Table.Left + Table.width - 128 Then
            OK = True
            For i = 0 To Ball.Count - 1
                If (BallP(i).InGame) Then
                    If (X - 128 - Ball(i).Left) ^ 2 + (cueball.Top - Ball(i).Top) ^ 2 < 65025 Then   'Pythagoras
                        OK = False
                    End If
                End If
            Next
            If OK Then cueball.Left = X - 128
        End If
    Else
        If Not cueBallMoves.Enabled And Not BallsMove.Enabled Then
        Yahas = (((a - cueball.Left - 128) ^ 2 + (b - cueball.Top - 128) ^ 2) ^ 0.5) / Cuelen
        If 0 < Yahas And Yahas <= 1 Then
            ForceBar.Value = Yahas * 240 + 10
        ElseIf Yahas > 1 Then
            ForceBar.Value = 250
        End If
    End If
        If Not cueBallMoves.Enabled And Not BallsMove.Enabled Then
            cue.x1 = cueball.Left + cueball.width / 2
            cue.y1 = cueball.Top + cueball.Height / 2
            Line2.x1 = cueball.Left + cueball.width / 2
            Line2.y1 = cueball.Top + cueball.Height / 2
            
            If X <> cueball.Left Then
                If cue.Visible = False Then cue.Visible = True
                If Line2.Visible = False Then Line2.Visible = True
                
                If X > cueball.Left + 128 Then
                    cue.x2 = cueball.Left + 128 + Cuelen * Cos(Atn((Y - cueball.Top - 128) / (X - cueball.Left - 128)))
                    cue.y2 = cueball.Top + 128 + Cuelen * Sin(Atn((Y - cueball.Top - 128) / (X - cueball.Left - 128)))
                    Line2.x2 = cueball.Left + 128 - Linelen * Cos(Atn((Y - cueball.Top - 128) / (X - cueball.Left - 128)))
                    Line2.y2 = cueball.Top + 128 - Linelen * Sin(Atn((Y - cueball.Top - 128) / (X - cueball.Left - 128)))
                ElseIf X < cueball.Left + 128 Then
                    cue.x2 = cueball.Left + 128 - Cuelen * Cos(Atn((Y - cueball.Top - 128) / (X - cueball.Left - 128)))
                    cue.y2 = cueball.Top + 128 - Cuelen * Sin(Atn((Y - cueball.Top - 128) / (X - cueball.Left - 128)))
                    Line2.x2 = cueball.Left + 128 + Linelen * Cos(Atn((Y - cueball.Top - 128) / (X - cueball.Left - 128)))
                    Line2.y2 = cueball.Top + 128 + Linelen * Sin(Atn((Y - cueball.Top - 128) / (X - cueball.Left - 128)))
                
                End If
            End If
        End If
    End If
End Sub

Private Sub cueBallMoves_Timer()
    StopCue = True
    NewXCueBall = cueball.Left + BallP(11).force * Cos(BallP(11).angle)
    NewYCueBall = cueball.Top + BallP(11).force * Sin(BallP(11).angle)
    
    Call isIn(11, NewXCueBall, NewYCueBall)
    Call collision
    Call BallsCollision
    Call isOutside(11, NewXCueBall, NewYCueBall)
    
    cueball.Left = cueball.Left + BallP(11).force * Cos(BallP(11).angle)
    cueball.Top = cueball.Top + BallP(11).force * Sin(BallP(11).angle)
    
    If BallP(11).force = 0 Then
        cueBallMoves.Enabled = False
        BallP(11).force = 0
    Else
        BallP(11).force = BallP(11).force - 1
        StopCue = False
    End If
    If StopCue Then cueBallMoves.Enabled = False
End Sub

Private Sub ballsMove_Timer()
    Dim i As Integer
    StopBalls = True
    For i = 0 To Ball.Count - 1
        If BallP(i).InGame Then
            
            NewXball = Ball(i).Left + BallP(i).force * Cos(BallP(i).angle)
            NewYball = Ball(i).Top + BallP(i).force * Sin(BallP(i).angle)
            
            Call isIn(i, NewXball, NewYball)
            Call BallsCollision
            Call collision
            Call isOutside(i, NewXball, NewYball)

            If BallP(i).force > 0 Then
                BallP(i).force = BallP(i).force - 1
                StopBalls = False
            Else
                BallP(i).force = 0
            End If
            
            Ball(i).Left = Ball(i).Left + BallP(i).force * Cos(BallP(i).angle)
            Ball(i).Top = Ball(i).Top + BallP(i).force * Sin(BallP(i).angle)
        End If
    Next
    If StopBalls Then BallsMove.Enabled = False
End Sub

Private Sub BallsCollision()
    Dim Temp As Integer, n As Integer, i As Integer
    For n = 0 To Ball.Count - 1
        If BallP(n).InGame Then
            For i = 0 To Ball.Count - 1
                If BallP(i).InGame Then
                    If (Ball(n).Left - Ball(i).Left) ^ 2 + (Ball(n).Top - Ball(i).Top) ^ 2 <= 65025 And n <> i Then 'Pythagoras
                        If BallP(i).force > 0 And BallP(n).force > 0 Then
                            BallP(i).angle = findAngle(Ball(n).Left, Ball(n).Top, Ball(i).Left, Ball(i).Top)
                            BallP(n).angle = findAngle(Ball(i).Left, Ball(i).Top, Ball(n).Left, Ball(n).Top)
            
                            Temp = BallP(i).force
                            BallP(i).force = BallP(n).force * 0.9
                            BallP(n).force = Temp * 0.9
                            
                        ElseIf BallP(i).force = 0 And BallP(n).force > 0 Then
                            
                            BallP(i).force = BallP(n).force * 0.9
                            BallP(n).force = BallP(n).force * 0.4
                            
                            If (Ball(i).Left <> Ball(n).Left) Then
                                BallP(i).angle = findAngle(Ball(n).Left, Ball(n).Top, Ball(i).Left, Ball(i).Top)
                                If (Ball(i).Left > Ball(n).Left And Ball(i).Top > Ball(n).Top) _
                                Or (Ball(i).Left < Ball(n).Left And Ball(i).Top < Ball(n).Top) Then
                                    BallP(n).angle = 3 * pi / 2 - BallP(i).angle
                                Else
                                    BallP(n).angle = pi / 2 - BallP(i).angle
                                End If
                            Else
                                BallP(i).angle = BallP(n).angle
                                BallP(n).angle = BallP(i).angle + pi
                            End If

                        ElseIf BallP(n).force = 0 And BallP(i).force > 0 Then
                            BallP(n).force = BallP(i).force * 0.9
                            BallP(i).force = BallP(i).force * 0.4
        
                            If (Ball(i).Left <> Ball(n).Left) Then
                                BallP(n).angle = findAngle(Ball(i).Left, Ball(i).Top, Ball(n).Left, Ball(n).Top)
                                If (Ball(n).Left > Ball(i).Left And Ball(n).Top > Ball(i).Top) _
                                Or (Ball(n).Left < Ball(i).Left And Ball(n).Top < Ball(i).Top) Then
                                    BallP(i).angle = 3 * pi / 2 - BallP(n).angle
                                Else
                                    BallP(i).angle = pi / 2 - BallP(n).angle
                                End If
                            Else
                                BallP(n).angle = BallP(i).angle
                                BallP(i).angle = BallP(n).angle + pi
                            End If
                            
                        End If
                        Ball(n).Left = Ball(n).Left + BallP(n).force * Cos(BallP(n).angle)
                        Ball(n).Top = Ball(n).Top + BallP(n).force * Sin(BallP(n).angle)
                                     
                        Ball(i).Left = Ball(i).Left + BallP(i).force * Cos(BallP(i).angle)
                        Ball(i).Top = Ball(i).Top + BallP(i).force * Sin(BallP(i).angle)
                                
                        BallsMove.Enabled = True
                    End If
                End If
            Next
       End If
   Next
End Sub




Private Sub mnuAbout_Click()
    frmSplash.Show
End Sub


Private Sub mnuNew_Click()
    Call NewGame
    ForceBar.Value = 10
End Sub

Private Sub Form_Click()
   If RelocateCueball Then
        RelocateCueball = False
    Else
        If (cmdShoot.Enabled) Then cmdShoot_Click
    End If
End Sub
Private Sub cmdShoot_Click()
    If ForceBar.Enabled Then
        BallP(11).force = ForceBar.Value
    End If
    BallP(11).angle = findAngle(a, b, cue.x1, cue.y1)
    cueBallMoves.Enabled = True
    cue.Visible = False
    Line2.Visible = False
    CheckTurns.Enabled = True
    MoreTurns = MoreTurns - 1
    BallsHit = 2
    FirstHitMade = False
    lblStatus.Caption = "Player " & 2 + Turn & " hit the cue ball..."
    cmdShoot.Enabled = False
End Sub

Private Sub collision()
    Dim Temp As Integer, i As Integer
    NewXCueBall = cueball.Left + Cos(BallP(11).angle) * (2 * BallP(11).force - 1)
    NewYCueBall = cueball.Top + Sin(BallP(11).angle) * (2 * BallP(11).force - 1)
    For i = 0 To Ball.Count - 1
        If (BallP(i).InGame) And (NewXCueBall - Ball(i).Left) ^ 2 + (NewYCueBall - Ball(i).Top) ^ 2 <= 65025 Then 'Pythagoras
            
            If Not FirstHitMade Then
                If OpState > 0 Then
                    If i < 5 Then
                        If (Turn And OpState = 1) Or (Turn = False And OpState = 2) Then
                            BallsHit = 1
                        Else
                            BallsHit = 2
                        End If
                    ElseIf 4 < i < 10 Then
                        If (Turn And OpState = 1) Or (Turn = False And OpState = 2) Then
                            BallsHit = 2
                        Else
                            BallsHit = 1
                        End If
                    End If
                Else
                    BallsHit = 1
                End If
                FirstHitMade = True
            End If
            
            If BallP(11).force > 0 And BallP(i).force = 0 Then
                
                BallP(i).force = BallP(11).force * 0.9
                BallP(11).force = BallP(11).force * 0.4
                
                If (Ball(i).Left <> cueball.Left) Then
                    BallP(i).angle = findAngle(cueball.Left, cueball.Top, Ball(i).Left, Ball(i).Top)
                    BallP(11).angle = BallP(i).angle + pi
                Else
                    BallP(i).angle = BallP(11).angle
                    BallP(11).angle = BallP(i).angle + pi
                End If
                
            ElseIf BallP(11).force = 0 And BallP(i).force > 0 Then
                BallP(11).force = BallP(i).force * 0.9
                BallP(i).force = BallP(i).force * 0.4
                If (Ball(i).Left <> cueball.Left) Then
                    BallP(11).angle = findAngle(Ball(i).Left, Ball(i).Top, cueball.Left, cueball.Top)
                    If (cueball.Left > Ball(i).Left And cueball.Top > Ball(i).Top) _
                    Or (cueball.Left < Ball(i).Left And cueball.Top < Ball(i).Top) Then
                        BallP(i).angle = 3 * pi / 2 - BallP(11).angle
                    Else
                        BallP(i).angle = pi / 2 - BallP(11).angle
                    End If
                Else
                    BallP(11).angle = BallP(i).angle
                    BallP(i).angle = BallP(11).angle + pi
                End If
            
            cueball.Left = cueball.Left + BallP(11).force * Cos(BallP(11).angle)
            cueball.Top = cueball.Top + BallP(11).force * Sin(BallP(11).angle)

            ElseIf BallP(11).force > 0 And BallP(i).force > 0 Then
                Temp = BallP(11).force
                BallP(11).force = BallP(i).force * 0.9
                BallP(i).force = Temp * 0.9
                
                Temp = BallP(11).angle
                BallP(11).angle = findAngle(Ball(i).Left, Ball(i).Top, cueball.Left, cueball.Top)
                BallP(i).angle = findAngle(cueball.Left, cueball.Top, Ball(i).Left, Ball(i).Top)
            End If
                

            cueBallMoves.Enabled = True
            BallsMove.Enabled = True
        End If
    Next
End Sub

Private Sub JG_Click()
Call NewGame
End Sub

Private Sub JG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
JG.ForeColor = &HFFC0C0
End Sub


Private Sub Label7_Click()
frmSetting.Show
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFFC0C0
End Sub

Private Sub Label8_Click()
frmSplash.Show
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &HFFC0C0
End Sub

Private Sub ExitG_Click()
End
End Sub

Private Sub ExitG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ExitG.ForeColor = &HFFC0C0
End Sub

