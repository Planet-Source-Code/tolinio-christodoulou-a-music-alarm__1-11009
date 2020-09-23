VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   2580
   ClientLeft      =   5100
   ClientTop       =   3405
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1780.762
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   285
      HideSelection   =   0   'False
      Left            =   1560
      Locked          =   -1  'True
      MouseIcon       =   "frmAbout.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Text            =   "tolisc@hotmail.com"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":0614
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   4080
      TabIndex        =   0
      Top             =   1800
      Width           =   1260
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmAbout.frx":091E
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   1159.566
      Y2              =   1159.566
   End
   Begin VB.Label lblDescription 
      Caption         =   " Wake Up Tunes Music Alarm. You can set it to go off playing any kind of files supported by Windows and media player. "
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   1050
      TabIndex        =   2
      Top             =   960
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Music Alarm"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   120
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5323.484
      Y1              =   1159.566
      Y2              =   1159.566
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   600
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...


Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About WakeUp Tunes "
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub









Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.ForeColor = &H80000002
End Sub

Private Sub Form_Resize()
ThreeDTunnel frmAbout, 125, 32, 212
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.ForeColor = &H80000002
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.ForeColor = &HFF0000
End Sub

Sub ThreeDTunnel(frm As Form, r%, g%, b%)
Dim X%
Dim current
   'Visit DiP's VB World At:
   'http://come.to/dipsvbworld
   frm.Scale (0, 100)-(100, 0)
   frm.BackColor = vbBlack
   frm.ForeColor = vbBlack
   For X% = 0 To 100
      frm.ForeColor = RGB(120, 23, 210)
      frm.Line (X%, 0)-(100 - X%, 100)
      frm.ForeColor = RGB(r%, 197, b%)
      frm.Line (0, X%)-(100, 100 - X%)
      current = Timer
      Do: DoEvents
      Loop Until Timer - current > 1E-99
   Next X%
   For X% = 10 To 50
      Me.Line (50 - X%, 50 + X%)-(50 + X%, 50 - X%), , BF
      current = Timer
      Do: DoEvents
      Loop Until Timer - current > 1E-99
   Next X%
   frmAbout.Cls
   frmAbout.BackColor = &H8000000F
   cmdOK.Enabled = True
   
End Sub

