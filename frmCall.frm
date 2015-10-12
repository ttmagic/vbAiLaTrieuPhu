VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmCall 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "- Nãi to lªn, kh«ng nghe râ :v"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmCall.frx":0000
   ScaleHeight     =   5415
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timer 
      Interval        =   1000
      Left            =   4320
      Top             =   4080
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
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
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nhê ng­êi th©n trî gióp"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C1FF&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "- §¸p ¸n lµ XX nhÐ. Kh«ng ch¾c l¾m ®©u."
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "- Tõ tõ ®Ó tra google ®·. Mµ c©u hái lµ g× ý nhÓ :v"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "- Nãi to lªn, kh«ng nghe râ :v"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "- §äc c©u hái ®i xem nµo!"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5295
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   873
   End
End
Attribute VB_Name = "frmCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dem As Integer
Dim dapan As String
Dim cauhoi As Integer
Dim dapanfake As String

Sub taoDapAnFake()
Dim a As Integer
Randomize
a = Int(Rnd * 9) + 1   'a =1~~10
If a < 7 Then           '70% tra loi dung
dapanfake = dapan
ElseIf a = 7 Then   '30% tra loi sai
dapanfake = "A"
ElseIf a = 8 Then
dapanfake = "B"
ElseIf a = 9 Then
dapanfake = "C"
ElseIf a = 10 Then
dapanfake = "D"
End If
End Sub
Private Sub Form_Load()
'Set position
Me.Top = 15 * 70
Me.Left = 15 * 800
wmp.settings.volume = 100
wmp.URL = App.Path + "\sounds\call.wav"
dapan = UCase(frmGame.dapan)
Call taoDapAnFake
cauhoi = frmGame.cauhoihientai
If cauhoi < 8 Then
lbl4.Caption = "- §¸p ¸n lµ " + dapan + " nhÐ, chuÈn kh«ng ph¶i chØnh ®©u."
Else
lbl4.Caption = "- §¸p ¸n lµ " + dapanfake + " nhÐ. Kh«ng ch¾c l¾m ®©u :v"
End If
End Sub

Private Sub btnOK_Click()
frmGame.Show
frmGame.Enabled = True
Unload Me
End Sub

Private Sub timer_Timer()
dem = dem + 1
If dem = 4 Then
lbl2.Visible = True
ElseIf dem = 6 Then
lbl3.Visible = True
ElseIf dem = 9 Then
lbl4.Visible = True
ElseIf dem > 10 Then
Me.Show
timer.Enabled = False
btnOK.Enabled = True
frmGame.wmp.Controls.play
End If
End Sub
