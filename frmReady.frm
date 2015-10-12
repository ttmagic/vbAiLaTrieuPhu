VERSION 5.00
Begin VB.Form frmReady 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmReady.frx":0000
   ScaleHeight     =   4050
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timer 
      Interval        =   1000
      Left            =   8880
      Top             =   960
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C1FF&
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.Label btnCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ch­a s½n sµng"
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
      Height          =   735
      Left            =   1320
      TabIndex        =   2
      Top             =   2760
      Width           =   6495
   End
   Begin VB.Label btnOK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S½n sµng"
      Enabled         =   0   'False
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
      Height          =   735
      Left            =   1320
      TabIndex        =   1
      Top             =   1560
      Width           =   6495
   End
   Begin VB.Image frontB 
      Height          =   945
      Left            =   960
      Picture         =   "frmReady.frx":16225
      Top             =   2640
      Width           =   7260
   End
   Begin VB.Image frontA 
      Height          =   945
      Left            =   960
      Picture         =   "frmReady.frx":167FA
      Top             =   1440
      Width           =   7260
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   960
      Picture         =   "frmReady.frx":16DCF
      Top             =   2640
      Width           =   7260
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   960
      Picture         =   "frmReady.frx":17305
      Top             =   1440
      Width           =   7260
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B¹n ®· s½n sµng ch¬i víi chóng t«i ch­a?"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9015
   End
End
Attribute VB_Name = "frmReady"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias _
      "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As _
      Long) As Long
   Const SND_SYNC = &H0
   Const SND_ASYNC = &H1
   Const SND_NODEFAULT = &H2
   Const SND_LOOP = &H8
   Const SND_NOSTOP = &H10

      Dim demnguoc As Integer
   Private Sub btnCancel_Click()
   frmMenu.wmp.Controls.play
Call frmGame.disconnect
Unload frmGame
Unload frmThangDiem
Unload Me
End Sub
Private Sub btnCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frontA.Visible = True
frontB.Visible = False
End Sub

Private Sub btnOK_Click()
frmGame.wmp.Controls.pause
sndPlaySound App.Path + "\sounds\comeBack.wav", SND_SYNC
frmGame.wmp.settings.playCount = 100
frmGame.wmp.URL = App.Path + "\sounds\bg1-5.mp3"
frmGame.wmp.settings.playCount = 100
frmGame.Show
frmGame.Enabled = True
frmGame.BatDauCauHoi (frmGame.cauhoihientai)
Unload Me
End Sub

Private Sub btnOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frontA.Visible = False
frontB.Visible = True
End Sub

Private Sub Form_Load()
demnguoc = 5
frmGame.Enabled = False
frontA.Visible = True
frontB.Visible = True
End Sub

Private Sub timer_Timer()
demnguoc = demnguoc - 1
lblCount.Caption = Str(demnguoc)
If demnguoc = 0 Then
btnOK.Enabled = True
btnCancel.Enabled = True
timer.Enabled = False
End If
End Sub
