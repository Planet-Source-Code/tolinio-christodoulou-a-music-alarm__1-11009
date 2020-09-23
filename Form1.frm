VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   4860
   ClientTop       =   4410
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4695
   Begin VB.PictureBox Picture2 
      BackColor       =   &H8000000A&
      Height          =   1575
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   840
      Width           =   4455
      Begin MSComctlLib.Slider Slider1 
         Height          =   1575
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   2778
         _Version        =   393216
         MousePointer    =   99
         MouseIcon       =   "Form1.frx":1C36E
         Orientation     =   1
         LargeChange     =   2
         SelectRange     =   -1  'True
         SelLength       =   10
         TickStyle       =   3
         TextPosition    =   1
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4080
         MouseIcon       =   "Form1.frx":1C688
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   615
      Left            =   120
      Picture         =   "Form1.frx":1C992
      ScaleHeight     =   555
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.Timer Timer3 
         Interval        =   1
         Left            =   3840
         Top             =   600
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   735
         Left            =   3960
         TabIndex        =   1
         Top             =   0
         Width           =   4455
      End
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   4680
      X2              =   4680
      Y1              =   2760
      Y2              =   840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   0
      X2              =   4680
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   840
      Y2              =   2760
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
ThreeDForm Form1
Left = frmMM.Left

Top = frmMM.Top + frmMM.Height

Label1.Caption = frmMM.lblFileName.Caption
If frmMM.mnuEq.Checked = False Then
frmMM.mnuEq.Checked = True

End If
End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.SetFocus
End Sub

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









Private Sub Form_Unload(Cancel As Integer)
If frmMM.mnuEq.Checked = True Then
frmMM.mnuEq.Checked = False

End If
End Sub

Private Sub Label1_Click()
form1Show = form1Show + 1
Unload Me
End Sub


Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
form1Show = form1Show + 1
Unload Me
End Sub

Private Sub Picture1_Click()
form1Show = form1Show + 1
Unload Me
End Sub

Private Sub Timer3_Timer()
If Label1.Left < Picture1.Width - Picture1.Width - Label1.Width Then
    Label1.Left = Picture1.Width - 1
    
    Label1.Left = Label1.Left - 5
    
Else
    Label1.Left = Label1.Left - 10
    
End If
End Sub
