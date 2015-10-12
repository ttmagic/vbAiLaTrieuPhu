VERSION 5.00
Begin VB.Form frmThangDiem 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmThangDiem.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerAn 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4320
      Top             =   1440
   End
   Begin VB.Timer timerHien 
      Interval        =   1
      Left            =   4320
      Top             =   840
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "1       -     $200"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   15
      Top             =   9960
      Width           =   2655
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "2       -     $400"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   14
      Top             =   9360
      Width           =   2535
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "3       -     $600"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   13
      Top             =   8760
      Width           =   5175
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "4       -     $1.000"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   12
      Top             =   8160
      Width           =   2775
   End
   Begin VB.Label lbl5 
      BackStyle       =   0  'Transparent
      Caption         =   "5       -     $2.000"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C1FF&
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   7560
      Width           =   2775
   End
   Begin VB.Label lbl6 
      BackStyle       =   0  'Transparent
      Caption         =   "6       -     $3.000"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label lbl7 
      BackStyle       =   0  'Transparent
      Caption         =   "7       -     $6.000"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Label lbl8 
      BackStyle       =   0  'Transparent
      Caption         =   "8       -     $10.000"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label lbl9 
      BackStyle       =   0  'Transparent
      Caption         =   "9       -     $14.000"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label lbl10 
      BackStyle       =   0  'Transparent
      Caption         =   "10     -     $22.000"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C1FF&
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Label lbl11 
      BackStyle       =   0  'Transparent
      Caption         =   "11     -     $30.000"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label lbl12 
      BackStyle       =   0  'Transparent
      Caption         =   "12     -     $40.000"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label lbl13 
      BackStyle       =   0  'Transparent
      Caption         =   "13     -     $60.000"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label lbl14 
      BackStyle       =   0  'Transparent
      Caption         =   "14     -     $85.000"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label lbl15 
      BackStyle       =   0  'Transparent
      Caption         =   "15     -     $150.000"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C1FF&
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Image thang 
      Height          =   645
      Left            =   240
      Picture         =   "frmThangDiem.frx":11128
      Stretch         =   -1  'True
      Top             =   10440
      Width           =   4860
   End
   Begin VB.Label title 
      BackStyle       =   0  'Transparent
      Caption         =   "ai lµ triÖu phó"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C1FF&
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.Image btnThuGon 
      Height          =   750
      Left            =   240
      Picture         =   "frmThangDiem.frx":1165E
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "frmThangDiem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cauhoihientai
Private Sub btnThuGon_Click()
timerAn.Enabled = True
End Sub
'1 nac => Top tru di 600

Private Sub Form_Click()
timerAn.Enabled = True
End Sub

Private Sub Form_Load()
frmGame.isShow = True
cauhoihientai = frmGame.cauhoihientai
thang.Top = thang.Top - (cauhoihientai * 600)
Me.Top = 0
Me.Left = Screen.Width 'bat dau tu ben phai
End Sub

Private Sub timerAn_Timer()
If Me.Left < Screen.Width Then
Me.Left = Me.Left + 120
Else
timerAn.Enabled = False
frmGame.Show
frmGame.isShow = False
Unload Me
End If
End Sub

Private Sub timerHien_Timer() 'xuat hien form
If Me.Left > Screen.Width - Me.Width + 120 Then
Me.Left = Me.Left - 120
Else
timerHien.Enabled = False
End If
End Sub

