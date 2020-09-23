VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmMM 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WakeUpTunes V1.0"
   ClientHeight    =   2115
   ClientLeft      =   5475
   ClientTop       =   3690
   ClientWidth     =   4590
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
   Icon            =   "frmMM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4590
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Picture         =   "frmMM.frx":030A
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   21
      Top             =   360
      Width           =   1335
      Begin VB.Label lblTime 
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "- - 88:88:88 - - "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   4080
      Top             =   1680
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Picture         =   "frmMM.frx":5CBC
      ScaleHeight     =   315
      ScaleWidth      =   2955
      TabIndex        =   15
      Top             =   360
      Width           =   3015
      Begin VB.Label lblFileName 
         BackStyle       =   0  'Transparent
         Caption         =   "Wake Up Tunes.. By Chris Christodoulou"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   2400
         TabIndex        =   20
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      Caption         =   "Music"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   840
      Picture         =   "frmMM.frx":1037E
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Select and Play Music"
      Top             =   960
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "Alarm"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      Picture         =   "frmMM.frx":107C8
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Set The Alarm Time and Music"
      Top             =   960
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4080
      Top             =   5040
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   4335
      Begin VB.Timer Timer3 
         Interval        =   1000
         Left            =   3960
         Top             =   480
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Show Controls"
         Height          =   615
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Shows the Play,Pause,Stop Controls"
         Top             =   120
         Width           =   855
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3840
         Picture         =   "frmMM.frx":10952
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   25
         Top             =   0
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Alarm Off"
         DisabledPicture =   "frmMM.frx":10C5C
         DownPicture     =   "frmMM.frx":1109E
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1440
         Picture         =   "frmMM.frx":114E0
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Set the alarm on or off"
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   4335
      Begin VB.CommandButton label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Set Time"
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmMM.frx":11922
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Music"
         Height          =   255
         Left            =   120
         MaskColor       =   &H80000012&
         MouseIcon       =   "frmMM.frx":11C2C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   2280
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Done"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   2280
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004000&
         Caption         =   "Done?"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label51 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Choose Music to Play"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.Line Line4 
         X1              =   2640
         X2              =   3840
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label41 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Set Time Here "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmMM.frx":11F36
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Line Line5 
         X1              =   2640
         X2              =   3840
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "  Click the button on the left to choose the    music file you want to play at the                 scheduled alarm time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   1320
         TabIndex        =   13
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Click The button on the left to set the alarm time. When done check below"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   19
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Line Line3 
         X1              =   3840
         X2              =   3840
         Y1              =   1320
         Y2              =   1800
      End
      Begin VB.Line Line2 
         X1              =   2640
         X2              =   2640
         Y1              =   1320
         Y2              =   1800
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "PM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "AM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   1440
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Chat Alarm"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   24
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "The Alarm is Off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Label Label7 
      Caption         =   "Click the button on the left to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   16
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Menu alarm 
      Caption         =   "Alarm"
      Begin VB.Menu setAlarm 
         Caption         =   "Set Alarm"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Play Mp3 (no alarm)"
      End
      Begin VB.Menu mnuControls 
         Caption         =   "View Controls"
      End
      Begin VB.Menu mnuEq 
         Caption         =   "Equalizer"
      End
      Begin VB.Menu dia 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu howAlarm 
         Caption         =   "How to set the Alarm"
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
   Begin VB.Menu minimize 
      Caption         =   "Minimize"
      NegotiatePosition=   3  'Right
   End
   Begin VB.Menu mnuControl 
      Caption         =   "Controls"
      Visible         =   0   'False
      Begin VB.Menu maxWin 
         Caption         =   "Maximize"
      End
      Begin VB.Menu mnuPopExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim alarmOnOff As Boolean
Dim text
Dim remConts As Integer


Const conMinimized = 1

Sub ThreeDForm(frmForm As Form)
    Const cPi = 3.1415926
    Dim intLineWidth As Integer
    intLineWidth = 5
    ' 'save scale mode
    Dim intSaveScaleMode As Integer
    intSaveScaleMode = frmForm.ScaleMode
    frmForm.ScaleMode = 3
    Dim intScaleWidth As Integer
    Dim intScaleHeight As Integer
    intScaleWidth = frmForm.ScaleWidth
    intScaleHeight = frmForm.ScaleHeight
    ' 'clear form
    frmForm.Cls
    ' 'draw white lines
    frmForm.Line (0, intScaleHeight)-(intLineWidth, 0), &HFFFFFF, BF
    frmForm.Line (0, intLineWidth)-(intScaleWidth, 0), &HFFFFFF, BF
    ' 'draw grey lines
    frmForm.Line (intScaleWidth, 0)-(intScaleWidth - intLineWidth, intScaleHeight), &H808080, BF
    frmForm.Line (intScaleWidth, intScaleHeight - intLineWidth)-(0, intScaleHeight), &H808080, BF
    ' 'draw triangles(actually circles) at corners
    Dim intCircleWidth As Integer
    intCircleWidth = Sqr(intLineWidth * intLineWidth + intLineWidth * intLineWidth)
    frmForm.FillStyle = 0
    frmForm.FillColor = QBColor(15)
    frmForm.Circle (intLineWidth, intScaleHeight - intLineWidth), intCircleWidth, QBColor(15), _
    -3.1415926, -3.90953745777778 '-180 * cPi / 180, -224 * cPi / 180
    frmForm.Circle (intScaleWidth - intLineWidth, intLineWidth), intCircleWidth, QBColor(15), _
    -0.78539815, -1.5707963 ' -45 * cPi / 180, -90 * cPi / 180
    ' 'draw black frame
    frmForm.Line (0, intScaleHeight)-(0, 0), 0
    frmForm.Line (0, 0)-(intScaleWidth - 1, 0), 0
    frmForm.Line (intScaleWidth - 1, 0)-(intScaleWidth - 1, intScaleHeight - 1), 0
    frmForm.Line (0, intScaleHeight - 1)-(intScaleWidth - 1, intScaleHeight - 1), 0
    frmForm.ScaleMode = intSaveScaleMode
End Sub



Private Sub FormCheckTime()

    alarmtime = InputBox("is the Time Ok?", "Check", alarmtime)
    Label1.Caption = "Alarm is on and set for " + alarmtime
    Label1.ForeColor = &H4000&
    Check4.Value = 1
    Check4.Enabled = True
    Check4.Caption = "Alarm On"
    If alarmtime = "" Then Exit Sub
    If Not IsDate(alarmtime) Then
        MsgBox "The time you entered was not valid."
    Else                                    ' String returned from InputBox is a valid time,
        alarmtime = CDate(alarmtime)        ' so store it as a date/time value in AlarmTime.
   End If
End Sub


Private Sub about_Click()
frmAbout.Show
 
End Sub

Private Sub Check1_Click()



alarmtime = Text1.text & text
If Check1.Value = 1 Then
Command3.Enabled = True
ElseIf Check1.Value = 0 Then
Command3.Enabled = False
End If

End Sub



Private Sub Check2_Click()
If Check2.Value = 1 Then
setAlarm_Click
Check1.Value = 0
Else: frmMM.Height = 2775
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Label5_Click
MMControl1.Command = "play"

Else: MMControl1.Command = "stop"
End If

End Sub



Private Sub Check4_Click()
If Check4.Value = 1 Then
alarmOnOff = True
Check4.Caption = "Alarm On"

ElseIf Check4.Value = 0 Then
alarmOnOff = False
Check4.Caption = "Alarm Off"

End If

End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then
MMControl1.Visible = True
Check5.Caption = "hide Controls"
Form_Resize
mnuControls.Checked = True

ElseIf Check5.Value = 0 Then
MMControl1.Visible = False
Check5.Caption = "show Controls"
mnuControls.Checked = False

Form_Resize
End If


End Sub

Private Sub Command1_Click()
frmMM.Height = 2850
Check1.Value = 0
Check2.Value = 0
Text1.text = ""
Text1.Visible = False
Option1.Visible = False
Option2.Visible = False
Label2.Visible = False
Label3.Visible = False
Line3.Visible = False
Line2.Visible = False




End Sub




Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If WindowState = conMinimized Then
PopupMenu mnuControl
End If

End Sub

Private Sub Form_Paint()
Form_Resize
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
label5.BackColor = &HE0E0E0
label4.BackColor = &HE0E0E0

End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

label4.BackColor = &HFF0000

End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
remConts = remConts + 1
If remConts Mod 2 = 0 Then
Text1.Visible = True
Option1.Visible = True
Option2.Visible = True
Label2.Visible = True
Label3.Visible = True
Label8.Visible = False
Line3.Visible = True
Line2.Visible = True
Line5.Visible = True
Line4.Visible = True

ElseIf remConts Mod 2 <> 0 Then
Text1.Visible = False
Option1.Visible = False
Option2.Visible = False
Label2.Visible = False
Label3.Visible = False
Label8.Visible = True
Line3.Visible = False
Line2.Visible = False
Line4.Visible = False
Line5.Visible = False

End If
End If
End Sub


Private Sub Label41_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
label5.BackColor = &HE0E0E0
label4.BackColor = &HE0E0E0

End Sub

Private Sub Label5_Click()

Dim strFileName As String
Dim strtemp As String
Dim i As Integer
'The  filter designates which files to look for
'All files listed after the | will be displayed.
'Text before the | is the title you give the group og files.
    CommonDialog1.Filter = "Media Files |*.mp3; *.mpg ; *.mpeg; *.wav; *.mid; *.midi; *.avi"
'Opens the standard Windows open Dialog bix for
'opening files with the criteria specified in the filter
    CommonDialog1.ShowOpen
'If blank exit sub
'blank would be no file selected or the
'cancel button was clicked
    If CommonDialog1.FileName = "" Then
        Exit Sub
    End If
' Getting the path and file name from the common dialog control
    strFileName = CommonDialog1.FileName
'Get the file extension and use the appropriate driver
    If Right$(strFileName, 3) = "wav" Then
            Wave
    ElseIf Right$(strFileName, 3) = "mp3" _
        Or Right$(strFileName, 3) = "mpg" _
        Or Right$(strFileName, 4) = "mpeg" Then
            MP
    ElseIf Right$(strFileName, 3) = "mid" _
        Or Right$(strFileName, 4) = "midi" Then
            Midi
    ElseIf Right$(strFileName, 3) = "avi" Then
            AVI
    Else
        MsgBox "Invalid File Type"
        frmMM.lblFileName = "Invalid File Type"
        Exit Sub
    End If
'Getting just the filename from the path
'and place it in the label
    For i = 1 To Len(strFileName)
        strtemp = Mid$(strFileName, Len(strFileName) - i, 1)
        If strtemp = "\" Then
            strtemp = Right$(strFileName, i)
            Exit For
        End If
    Next i
    
    frmMM.lblFileName = strtemp
    lblFileName.Width = (Len(strFileName) * 65)
    Form1.Label1.Width = (Len(strFileName) * 83)
    Label6.Caption = CommonDialog1.FileTitle & "......"
    Form1.Label1.Caption = CommonDialog1.FileTitle
    textnum = Len(strFileName)
End Sub

Private Sub Command3_Click()
If Not IsDate(alarmtime) Then
 MsgBox "The time you entered was not valid.", vbSystemModal, "Time Not Valid!"
 Check1.Value = 0
 
Else
FormCheckTime
frmMM.Height = 2048
Text1.text = ""
Check2.Value = 0
End If
End Sub

Private Sub Form_Load()
alarmOnOff = False
form1Show = 1
alarmtime = Label1.Caption
remConts = 1
Line3.Visible = False
Line2.Visible = False
Line4.Visible = False
Line5.Visible = False

End Sub

Private Sub Form_Resize()
 ThreeDForm frmMM
    If WindowState = conMinimized Then      ' If form is minimized, display the time in a caption.
        SetCaptionTime
    
    End If
    
End Sub

Private Sub SetCaptionTime()
    Caption = Format(Time, "Medium Time")   ' Display time using medium time format.
End Sub


Private Sub Form_Unload(Cancel As Integer)
MMControl1.Command = "Close"
   Unload Form1
   Unload frmAbout
    Unload Me
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
label5.BackColor = &HFF0000

End Sub





Private Sub Label51_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
label5.BackColor = &HE0E0E0
label4.BackColor = &HE0E0E0

End Sub

Private Sub lblFileName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
form1Show = form1Show + 1
If form1Show Mod 2 = 0 Then
Form1.Show
ElseIf form1Show Mod 2 <> 0 Then
Unload Form1
End If
End Sub

Private Sub maxWin_Click()
WindowState = 0
End Sub



Private Sub minimize_Click()
frmMM.WindowState = 1
End Sub

Private Sub mnuControls_Click()
If Check5.Value = 0 Then
Check5.Value = 1
ElseIf Check5.Value = 1 Then
Check5.Value = 0
End If

End Sub

Private Sub mnuEq_Click()
mnuEq.Checked = True

form1Show = form1Show + 1
If form1Show Mod 2 = 0 Then
Form1.Show
ElseIf form1Show Mod 2 <> 0 Then
Unload Form1
End If
End Sub

Private Sub mnuPlay_Click()
If Check3.Value = 0 Then
Check3.Value = 1
Else: Check3.Value = 0
Check3.Value = 1
End If
End Sub

Private Sub mnuPopExit_Click()
Unload Me
End Sub

Private Sub Option1_Click()
text = "AM"
End Sub

Private Sub Option2_Click()
text = "PM"
End Sub




Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
form1Show = form1Show + 1
If form1Show Mod 2 = 0 Then
Form1.Show
ElseIf form1Show Mod 2 <> 0 Then
Unload Form1
End If

End Sub

Private Sub setAlarm_Click()
frmMM.Height = 5580
Check2.Value = 1
End Sub






Private Sub Text1_Change()
If Text1.text <> "" Then
Check1.Enabled = True
Else: Check1.Enabled = False
End If
End Sub

Private Sub Timer1_Timer()
Static AlarmSounded As Integer
    If lblTime.Caption <> CStr(Time) Then
        ' It's now a different second than the one displayed.
        If alarmOnOff = True Then
        If Time >= alarmtime And Not AlarmSounded Then
            
            MMControl1.Command = "play"
            
                        
            AlarmSounded = True
        ElseIf Time < alarmtime Then
            AlarmSounded = False
        End If
        End If
        If WindowState = conMinimized Then
            ' If minimized, then update the form's Caption every minute.
            If Minute(CDate(Caption)) <> Minute(Time) Then SetCaptionTime
        Else
            ' Otherwise, update the label Caption in the form every second.
            lblTime.Caption = Time
        End If
    
    End If
End Sub

Private Sub mnuFileExit_Click()
'Ensuring The Contol is closed,
'if not the media would continue to
'play after the app is closed
    MMControl1.Command = "Close"
    Unload Me
End Sub


Private Sub Timer2_Timer()
If lblFileName.Left < Picture1.Width - Picture1.Width - lblFileName.Width Then
    lblFileName.Left = Picture1.Width - 1
    
    lblFileName.Left = lblFileName.Left - 5
    
Else
    lblFileName.Left = lblFileName.Left - 5
    
End If

End Sub




