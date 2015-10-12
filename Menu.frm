VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmMenu 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Menu.frx":08CA
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   4695
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
      _cx             =   8281
      _cy             =   1296
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "§å ¸n VB - NguyÔn Thanh Tïng - TH18.17 HUBT - at.6445022@gmail.com"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   11280
      Width           =   20295
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "tho¸t"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6600
      TabIndex        =   3
      Top             =   9360
      Width           =   7215
   End
   Begin VB.Image btnExit 
      Height          =   945
      Left            =   6600
      Picture         =   "Menu.frx":A713A
      Top             =   9240
      Width           =   7260
   End
   Begin VB.Label lblAdd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "thªm c©u hái"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6600
      TabIndex        =   2
      Top             =   8040
      Width           =   7215
   End
   Begin VB.Image btnAdd 
      Height          =   945
      Left            =   6600
      Picture         =   "Menu.frx":A770F
      Top             =   7920
      Width           =   7260
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "b¾t ®Çu"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6600
      TabIndex        =   1
      Top             =   5400
      Width           =   7215
   End
   Begin VB.Label lblHowto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "h­íng dÉn"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6600
      TabIndex        =   0
      Top             =   6720
      Width           =   7215
   End
   Begin VB.Image btnHowto 
      Height          =   945
      Left            =   6600
      Picture         =   "Menu.frx":A7CE4
      Top             =   6600
      Width           =   7260
   End
   Begin VB.Image btnStart 
      Height          =   945
      Left            =   6600
      Picture         =   "Menu.frx":A82B9
      Top             =   5280
      Width           =   7260
   End
   Begin VB.Image Image4 
      Height          =   945
      Left            =   6600
      Picture         =   "Menu.frx":A888E
      Top             =   7920
      Width           =   7260
   End
   Begin VB.Image Image3 
      Height          =   945
      Left            =   6600
      Picture         =   "Menu.frx":A8DC4
      Top             =   9240
      Width           =   7260
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6600
      Picture         =   "Menu.frx":A92FA
      Top             =   6600
      Width           =   7260
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   6600
      Picture         =   "Menu.frx":A9830
      Top             =   5280
      Width           =   7260
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'set bg music
wmp.URL = App.Path + "\sounds\introMenu.mp3"
wmp.settings.volume = 100
wmp.settings.playCount = 100
wmp.Visible = False
End Sub

Private Sub lblStart_Click()
frmGame.Show
frmReady.Show
frmHowto.Show
Me.wmp.Controls.pause
End Sub

Private Sub lblHowto_Click()
frmHowto.Show
frmMenu.Enabled = False
End Sub

Private Sub lblAdd_Click()
frmLogin.Show
frmMenu.Enabled = False
End Sub

Private Sub lblExit_Click()
Unload frmGame
Unload frmHowto
Unload frmAdd
Unload frmThangDiem
Unload Me
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnStart.Visible = True
btnHowto.Visible = True
btnAdd.Visible = True
btnExit.Visible = True
End Sub

Private Sub lblStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnStart.Visible = False
btnHowto.Visible = True
btnAdd.Visible = True
btnExit.Visible = True
End Sub

Private Sub lblHowto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnStart.Visible = True
btnHowto.Visible = False
btnAdd.Visible = True
btnExit.Visible = True
End Sub

Private Sub lblAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnStart.Visible = True
btnHowto.Visible = True
btnAdd.Visible = False
btnExit.Visible = True
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnStart.Visible = True
btnHowto.Visible = True
btnAdd.Visible = True
btnExit.Visible = False
End Sub
