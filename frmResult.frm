VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmResult 
   BackColor       =   &H00400040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   Picture         =   "frmResult.frx":0000
   ScaleHeight     =   5160
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timer 
      Interval        =   1000
      Left            =   360
      Top             =   3360
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   3255
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
      _cx             =   5741
      _cy             =   1296
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
      Left            =   7080
      TabIndex        =   2
      Top             =   4080
      Width           =   6255
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6600
      Picture         =   "frmResult.frx":16225
      Top             =   3960
      Width           =   7260
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$ 69000"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C1FF&
      Height          =   1455
      Left            =   2880
      TabIndex        =   1
      Top             =   1800
      Width           =   14535
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "th¾ng cuéc"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   20535
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   2880
      Picture         =   "frmResult.frx":1675B
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   14640
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public title As String
Public score As Integer
Public isMillionaire As Boolean
Public dotre As Integer, dem As Integer

Private Sub Form_Load()
dem = 0
lbl.Caption = title
lblScore.Caption = "$" + Str(score) + "00"
wmp.settings.volume = 100
If dotre = 0 Then
timer.Enabled = 0
Call checkTrieuPhu
End If
End Sub

Private Sub lblExit_Click()
frmMenu.Show
frmMenu.wmp.Controls.play
Call frmGame.disconnect
Unload Me
Unload frmGame
Unload frmThangDiem
End Sub

Private Sub timer_Timer()
dem = dem + 1
If dem = dotre Then
   Call checkTrieuPhu
End If
End Sub

Sub checkTrieuPhu()
 If isMillionaire = True Then
    wmp.URL = App.Path + "\sounds\TRIEUPHU.mp3"
    Else
    wmp.URL = App.Path + "\sounds\DUNGCHOI.wav"
    End If
End Sub
