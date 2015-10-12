VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmGuest 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
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
   Picture         =   "frmGuest.frx":0000
   ScaleHeight     =   5565
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timer 
      Interval        =   1000
      Left            =   3840
      Top             =   3360
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
      TabIndex        =   9
      Top             =   4800
      Width           =   4095
   End
   Begin VB.PictureBox CotD 
      Appearance      =   0  'Flat
      BackColor       =   &H00025DFF&
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   3240
      Picture         =   "frmGuest.frx":E33E
      ScaleHeight     =   3570
      ScaleWidth      =   585
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox CotC 
      Appearance      =   0  'Flat
      BackColor       =   &H00025DFF&
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   2280
      Picture         =   "frmGuest.frx":119A8
      ScaleHeight     =   3570
      ScaleWidth      =   585
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox cotB 
      Appearance      =   0  'Flat
      BackColor       =   &H00025DFF&
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   1320
      Picture         =   "frmGuest.frx":15012
      ScaleHeight     =   3570
      ScaleWidth      =   585
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox cotA 
      Appearance      =   0  'Flat
      BackColor       =   &H00025DFF&
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   360
      Picture         =   "frmGuest.frx":1867C
      ScaleHeight     =   3570
      ScaleWidth      =   585
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
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
      _cx             =   873
      _cy             =   873
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "A      B      c      d"
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
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label lblD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C1FF&
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C1FF&
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C1FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C1FF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmGuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dem As Integer
Dim dapan As String
Dim cauhoi As Integer
Dim cotdung As Integer      'nhieu% nhat
Dim cotsai1 As Integer      '3 cai o duoi
Dim cotsai2 As Integer      'chia se % cho nhau
Dim cotsai3 As Integer
Dim conlai As Integer, conlai2 As Integer

Private Sub btnOK_Click()
frmGame.Show
frmGame.Enabled = True
Unload Me
End Sub

'height 3600 = 100%
'top = 720+ phan chieu cao con lai
Private Sub Form_Load()
'Set position
Me.Top = 15 * 70
Me.Left = 15 * 800
dem = 0
cauhoi = frmGame.cauhoihientai
dapan = frmGame.dapanhientai
wmp.settings.volume = 100
wmp.URL = App.Path + "\sounds\guest.wav"
End Sub

Private Sub timer_Timer()  'random va hien thi luon
dem = dem + 1
If dem = 1 Then
    'Tinh phan tram cot dung. Cauhoi <6 - 90%, <10 - 50%
    If cauhoi < 6 Then
    Randomize
    cotdung = Int(Rnd * 20) + 75  'cot dung max la 95, minla 75
    ElseIf cauhoi < 10 Then
    Randomize
    cotdung = Int(Rnd * 35) + 50  'cot dung max la 85, minla 50
    Else
    Randomize
    cotdung = Int(Rnd * 35) + 40  'cot dung max la 75, minla 40
    End If
ElseIf dem = 3 Then
    conlai = 100 - cotdung      'con lai co the tu 10--60%
    Randomize
    cotsai1 = Int(Rnd * (conlai - 1)) + 1 'tu 1--conlai-1
ElseIf dem = 4 Then
    conlai2 = conlai - cotsai1  'phan con lai thu 2
    Randomize
    cotsai2 = Int(Rnd * (conlai2 - 1)) + 1
    cotsai3 = 100 - cotdung - cotsai1 - cotsai2
ElseIf dem = 7 Then     'bat dau hien thi day.
Call HienThiBieuDo

cotA.Visible = True
cotB.Visible = True
CotC.Visible = True
CotD.Visible = True
ElseIf dem > 7 Then
timer.Enabled = False
btnOK.Enabled = True
frmGame.wmp.Controls.play
End If
End Sub


'height 3600 = 100%
'top = 720+ phan chieu cao con lai
Sub HienThiBieuDo()
Me.Show
If dapan = "a" Then
    lblA.Caption = Str(cotdung) + "%"
    lblB.Caption = Str(cotsai1) + "%"
    lblC.Caption = Str(cotsai2) + "%"
    lblD.Caption = Str(cotsai3) + "%"
cotA.Height = cotdung * 36
cotB.Height = cotsai1 * 36
CotC.Height = cotsai2 * 36
CotD.Height = cotsai3 * 36
ElseIf dapan = "b" Then
    lblB.Caption = Str(cotdung) + "%"
    lblA.Caption = Str(cotsai1) + "%"
    lblC.Caption = Str(cotsai2) + "%"
    lblD.Caption = Str(cotsai3) + "%"
cotB.Height = cotdung * 36
cotA.Height = cotsai1 * 36
CotC.Height = cotsai2 * 36
CotD.Height = cotsai3 * 36
ElseIf dapan = "c" Then
    lblC.Caption = Str(cotdung) + "%"
    lblB.Caption = Str(cotsai1) + "%"
    lblA.Caption = Str(cotsai2) + "%"
    lblD.Caption = Str(cotsai3) + "%"
CotC.Height = cotdung * 36
cotB.Height = cotsai1 * 36
cotA.Height = cotsai2 * 36
CotD.Height = cotsai3 * 36
Else
    lblD.Caption = Str(cotdung) + "%"
    lblB.Caption = Str(cotsai1) + "%"
    lblC.Caption = Str(cotsai2) + "%"
    lblA.Caption = Str(cotsai3) + "%"
CotD.Height = cotdung * 36
cotB.Height = cotsai1 * 36
CotC.Height = cotsai2 * 36
cotA.Height = cotsai3 * 36
End If
cotA.Top = 600 + (3600 - cotA.Height)
cotB.Top = 600 + (3600 - cotB.Height)
CotC.Top = 600 + (3600 - CotC.Height)
CotD.Top = 600 + (3600 - CotD.Height)
End Sub
