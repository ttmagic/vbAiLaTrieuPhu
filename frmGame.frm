VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmGame 
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
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGame.frx":08CA
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer timerNhapNhay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   4560
   End
   Begin VB.Timer timerCauHoi 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   3960
   End
   Begin VB.Image btnTvtc 
      Height          =   1350
      Left            =   6360
      Picture         =   "frmGame.frx":4D320
      Top             =   120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Image Image6 
      Height          =   1350
      Left            =   6360
      Picture         =   "frmGame.frx":4DE45
      Top             =   120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Image btnMoRong 
      Height          =   750
      Left            =   15720
      Picture         =   "frmGame.frx":4EC14
      Top             =   120
      Width           =   750
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   855
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   5175
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
      _cx             =   9128
      _cy             =   1508
   End
   Begin VB.Label dapan 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   7920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label markCau 
      BackStyle       =   0  'Transparent
      Caption         =   "C©u 1:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cau 15 - $150000"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   16440
      TabIndex        =   9
      Top             =   240
      Width           =   3855
   End
   Begin VB.Image btnCall 
      Height          =   1350
      Left            =   2160
      Picture         =   "frmGame.frx":4EF07
      Top             =   120
      Width           =   1875
   End
   Begin VB.Image btnGuest 
      Height          =   1350
      Left            =   4200
      Picture         =   "frmGame.frx":4FA80
      Top             =   120
      Width           =   1875
   End
   Begin VB.Image btn5050 
      Height          =   1350
      Left            =   120
      Picture         =   "frmGame.frx":50656
      Top             =   120
      Width           =   1875
   End
   Begin VB.Image Image5 
      Height          =   1350
      Left            =   4200
      Picture         =   "frmGame.frx":50DDE
      Top             =   120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Image Image4 
      Height          =   1350
      Left            =   2160
      Picture         =   "frmGame.frx":51CC0
      Top             =   120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   1350
      Left            =   120
      Picture         =   "frmGame.frx":52B47
      Top             =   120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Image btnStop 
      Height          =   870
      Left            =   240
      Picture         =   "frmGame.frx":5382C
      Top             =   6120
      Width           =   2850
   End
   Begin VB.Image btnSure 
      Height          =   870
      Left            =   17400
      Picture         =   "frmGame.frx":54068
      Top             =   6120
      Width           =   2850
   End
   Begin VB.Label lblD 
      BackStyle       =   0  'Transparent
      Caption         =   "ChÝn"
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
      Left            =   10920
      TabIndex        =   8
      Top             =   10080
      Width           =   6135
   End
   Begin VB.Label lblC 
      BackStyle       =   0  'Transparent
      Caption         =   "S¸u"
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
      Left            =   3720
      TabIndex        =   7
      Top             =   10080
      Width           =   6135
   End
   Begin VB.Label lblB 
      BackStyle       =   0  'Transparent
      Caption         =   "Hai"
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
      Left            =   10920
      TabIndex        =   6
      Top             =   8880
      Width           =   6135
   End
   Begin VB.Label lblA 
      BackStyle       =   0  'Transparent
      Caption         =   "Mét"
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
      Left            =   3720
      TabIndex        =   5
      Top             =   8880
      Width           =   6135
   End
   Begin VB.Label markD 
      BackStyle       =   0  'Transparent
      Caption         =   "D:"
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   10440
      TabIndex        =   4
      Top             =   10080
      Width           =   375
   End
   Begin VB.Label markC 
      BackStyle       =   0  'Transparent
      Caption         =   "C:"
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   10080
      Width           =   375
   End
   Begin VB.Label markB 
      BackStyle       =   0  'Transparent
      Caption         =   "B:"
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   10440
      TabIndex        =   2
      Top             =   8880
      Width           =   375
   End
   Begin VB.Label markA 
      BackStyle       =   0  'Transparent
      Caption         =   "A:"
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   8880
      Width           =   375
   End
   Begin VB.Label lblCauhoi 
      BackStyle       =   0  'Transparent
      Caption         =   "Tõ nµo sau ®©y xuÊt hiÖn ®Çu tiªn trong tõ ®iÓn TiÕng ViÖt?"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   3720
      TabIndex        =   0
      Top             =   6960
      Width           =   12975
   End
   Begin VB.Image frontD 
      Height          =   945
      Left            =   10200
      Picture         =   "frmGame.frx":54794
      Top             =   9840
      Width           =   7260
   End
   Begin VB.Image frontC 
      Height          =   945
      Left            =   3000
      Picture         =   "frmGame.frx":54D69
      Top             =   9840
      Width           =   7260
   End
   Begin VB.Image frontB 
      Height          =   945
      Left            =   10200
      Picture         =   "frmGame.frx":5533E
      Top             =   8640
      Width           =   7260
   End
   Begin VB.Image frontA 
      Height          =   945
      Left            =   3000
      Picture         =   "frmGame.frx":55913
      Top             =   8640
      Width           =   7260
   End
   Begin VB.Image backD 
      Height          =   945
      Left            =   10200
      Picture         =   "frmGame.frx":55EE8
      Top             =   9840
      Width           =   7260
   End
   Begin VB.Image backC 
      Height          =   945
      Left            =   3000
      Picture         =   "frmGame.frx":5641E
      Top             =   9840
      Width           =   7260
   End
   Begin VB.Image backB 
      Height          =   945
      Left            =   10200
      Picture         =   "frmGame.frx":56954
      Top             =   8640
      Width           =   7260
   End
   Begin VB.Image backA 
      Height          =   945
      Left            =   3000
      Picture         =   "frmGame.frx":56E8A
      Top             =   8640
      Width           =   7260
   End
   Begin VB.Image backCH 
      Height          =   1470
      Left            =   2880
      Picture         =   "frmGame.frx":573C0
      Top             =   6840
      Width           =   14640
   End
End
Attribute VB_Name = "frmGame"
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
'-----------------PLAY SOUND----------------'
Dim con As New ADODB.Connection
Dim recDe As New ADODB.Recordset
Dim recThuong As New ADODB.Recordset
Dim recKho As New ADODB.Recordset

Dim maxDe As Integer, maxThuong As Integer, maxKho As Integer 'so luong cau hoi o 3 bang trong database
Dim cauhoi(1 To 15) As Integer '15 cau hoi, index tu 0 den 14
Dim muctien(1 To 15) As Integer '15 muc tien ung voi 15 cau hoi
Public money As Integer  'so tien dat duoc sau khi xong
Public cauhoihientai As Integer
Dim phuonganhientai As String
Public dapanhientai As String
Dim dem As Integer 'dem thoi gian de hien cau hoi
Dim isPaused As Boolean 'tam dung tro choi luc phat nhac
Public isShow As Boolean   'kiem tra co dang hien thi thang diem khong

Private Sub btnTvtc_Click()
btnTvtc.Visible = False
Image6.Visible = True
frmGame.Enabled = False
frmTVTC.Show
End Sub


Private Sub Form_Click()
If isShow = True Then
frmThangDiem.Show
frmThangDiem.timerAn.Enabled = True
End If
End Sub

Private Sub Form_Load()
'Connect
con.CursorLocation = adUseClient
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\ailatrieuphu.mdb;Persist Security Info=False;Jet OLEDB:Database Password=01653330406"
recDe.Open "de", con, adOpenDynamic, adLockOptimistic
recThuong.Open "thuong", con, adOpenDynamic, adLockOptimistic
recKho.Open "kho", con, adOpenDynamic, adLockOptimistic
Call setDataFields

dem = 0
cauhoihientai = 1
phuonganhientai = ""
isPaused = False
dapanhientai = dapan.Caption
Call AnCauHoi
Call Tao15CauHoi
Call SetMucTienThuong
setTextScore (cauhoihientai)
'----Music-----
wmp.URL = App.Path + "\sounds\welcome.wav"
wmp.settings.volume = 100
End Sub

Private Sub btnStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnStop.Top = btnStop.Top + 15
End Sub
'----button Dung cuoc choi----'
Private Sub btnStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnStop.Top = 6120
If cauhoihientai > 1 Then
frmResult.title = "th¾ng cuéc"
frmResult.score = muctien(cauhoihientai - 1)
frmResult.dotre = 0 'do tre phat nhac
Else
frmResult.title = "thua cuéc"
frmResult.score = 0
frmResult.dotre = 0
End If
frmResult.Show
frmGame.Enabled = False
wmp.Controls.pause
End Sub

Private Sub btnSure_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnSure.Top = btnSure.Top + 15
If phuonganhientai <> "" Then
btnSure.Enabled = False
isPaused = True
End If
End Sub
'-------button chac chan----------kiem tra cau tra loi dung hay sai, xu ly' khi thua-----------
Private Sub btnSure_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnSure.Top = 6120
If phuonganhientai = dapanhientai Then   'CHON DUNG----------------------------------
If cauhoihientai > 6 Then sndPlaySound App.Path + "\sounds\hoihop.wav", SND_SYNC 'hoi hop: 6s
Call NhapNhayDapAn
Call PlaySoundChonDung
    If cauhoihientai < 15 Then
    cauhoihientai = cauhoihientai + 1
        Else 'vuot qua 15 cau
    'YOU ARE MILLIONAIRE
    frmResult.title = "th¾ng cuéc"
    frmResult.score = muctien(cauhoihientai)
    frmResult.dotre = 3
    frmResult.isMillionaire = True  '>>> tao la` trieu phu
    frmGame.Enabled = False
    wmp.Controls.pause
    frmResult.Show
    End If
ElseIf phuonganhientai <> "" And phuonganhientai <> dapanhientai Then  'CHON SAI-------
wmp.Controls.pause
If cauhoihientai > 6 Then sndPlaySound App.Path + "\sounds\hoihop.wav", SND_SYNC 'hoi hop: 6s
Call NhapNhayDapAn
sndPlaySound App.Path + "\sounds\wrong.wav", SND_ASYNC
If cauhoihientai < 6 Then
frmResult.title = "thua cuéc"
frmResult.score = 0
frmResult.dotre = 3
ElseIf cauhoihientai < 11 Then
frmResult.title = "th¾ng cuéc"
frmResult.score = muctien(5)
frmResult.dotre = 3
Else
frmResult.title = "th¾ng cuéc"
frmResult.score = muctien(10)
frmResult.dotre = 3
End If
frmResult.Show
End If
End Sub

Sub Tao15CauHoi()       'tao 15 cau hoi ngau nhien, random trong 3 table
'Lay so luong cau hoi trong moi table
maxDe = recDe.RecordCount
maxThuong = recThuong.RecordCount
maxKho = recKho.RecordCount
'random khong + 1 nen khi move, khong can -1
Dim a As Integer, i As Integer
i = 1
Do While i < 6
Randomize
a = Int(Rnd * maxDe)
If isAdded(a, 1, 5) = False Then
cauhoi(i) = a
i = i + 1
End If
Loop
Do While i < 11
Randomize
a = Int(Rnd * maxThuong)
If isAdded(a, 6, 10) = False Then
cauhoi(i) = a
i = i + 1
End If
Loop
Do While i < 16
Randomize
a = Int(Rnd * maxKho)
If isAdded(a, 11, 15) = False Then
cauhoi(i) = a
i = i + 1
End If
Loop
End Sub

Function BatDauCauHoi(cau As Integer) '1--15
If cau <= 5 Then             'thay doi recordSource theo muc' cau hoi
recDe.MoveFirst
recDe.Move (cauhoi(cau))
Set lblCauhoi.DataSource = recDe
Set lblA.DataSource = recDe
Set lblB.DataSource = recDe
Set lblC.DataSource = recDe
Set lblD.DataSource = recDe
Set dapan.DataSource = recDe
ElseIf cau <= 10 Then
recThuong.MoveFirst
recThuong.Move (cauhoi(cau))
Set lblCauhoi.DataSource = recThuong
Set lblA.DataSource = recThuong
Set lblB.DataSource = recThuong
Set lblC.DataSource = recThuong
Set lblD.DataSource = recThuong
Set dapan.DataSource = recThuong
Else
recKho.MoveFirst
recKho.Move (cauhoi(cau))
Set lblCauhoi.DataSource = recKho
Set lblA.DataSource = recKho
Set lblB.DataSource = recKho
Set lblC.DataSource = recKho
Set lblD.DataSource = recKho
Set dapan.DataSource = recKho
End If
dapanhientai = dapan.Caption
Call HienCauHoi
End Function


Private Sub btn5050_Click() 'tro giup 5050  (dai` von` :v )
btn5050.Visible = False
Image3.Visible = True
sndPlaySound App.Path + "\sounds\5050.wav", SND_ASYNC
Dim giulai As Integer
Randomize
giulai = Int(Rnd * 2) + 1
If dapanhientai = "a" Then
    If giulai = 1 Then
    markC.Visible = False
    lblC.Visible = False
    markD.Visible = False
    lblD.Visible = False
    ElseIf giulai = 2 Then
    markB.Visible = False
    lblB.Visible = False
    markD.Visible = False
    lblD.Visible = False
    Else
    markC.Visible = False
    lblC.Visible = False
    markB.Visible = False
    lblB.Visible = False
    End If
ElseIf dapanhientai = "b" Then
    If giulai = 1 Then
    markC.Visible = False
    lblC.Visible = False
    markD.Visible = False
    lblD.Visible = False
    ElseIf giulai = 2 Then
    markA.Visible = False
    lblA.Visible = False
    markD.Visible = False
    lblD.Visible = False
    Else
    markC.Visible = False
    lblC.Visible = False
    markA.Visible = False
    lblA.Visible = False
    End If
ElseIf dapanhientai = "c" Then
    If giulai = 1 Then
    markB.Visible = False
    lblB.Visible = False
    markD.Visible = False
    lblD.Visible = False
    ElseIf giulai = 2 Then
    markA.Visible = False
    lblA.Visible = False
    markD.Visible = False
    lblD.Visible = False
    Else
    markA.Visible = False
    lblA.Visible = False
    markB.Visible = False
    lblB.Visible = False
    End If
ElseIf dapanhientai = "d" Then
    If giulai = 1 Then
    markC.Visible = False
    lblC.Visible = False
    markB.Visible = False
    lblB.Visible = False
    ElseIf giulai = 2 Then
    markC.Visible = False
    lblC.Visible = False
    markA.Visible = False
    lblA.Visible = False
    Else
    markA.Visible = False
    lblA.Visible = False
    markB.Visible = False
    lblB.Visible = False
    End If
End If
End Sub

Private Sub btnCall_Click() 'Goi dt cho nguoi than
btnCall.Visible = False
Image4.Visible = True
wmp.Controls.pause
Me.Enabled = False
frmCall.Show
End Sub

Private Sub btnGuest_Click() 'Hoi y kien khan gia
btnGuest.Visible = False
Image5.Visible = True
wmp.Controls.pause
Me.Enabled = False
frmGuest.Show
End Sub

Private Sub btnMoRong_Click() ' Xem thang diem
frmThangDiem.Show
End Sub


Sub KiemTra510()                'kiem tra co phai cau 5 hay 10 ko de play sound dac biet va doi nhac nen
If cauhoihientai = 6 Then
btnTvtc.Visible = True      'hien thi goi y moi'
wmp.Controls.pause
Call AnCauHoi
sndPlaySound App.Path + "\sounds\welcomeBack.wav", SND_SYNC     '7s
wmp.URL = App.Path + "\sounds\bg6-10.mp3"
End If
If cauhoihientai = 11 Then
wmp.Controls.pause
Call AnCauHoi
sndPlaySound App.Path + "\sounds\welcomeBack.wav", SND_SYNC
wmp.URL = App.Path + "\sounds\bg11-15.mp3"
End If
End Sub

Sub PlaySoundChonDung()     'Play sound khi xong 1 cau hoi
If cauhoihientai <= 5 Then
sndPlaySound App.Path + "\sounds\right1.wav", SND_ASYNC
ElseIf cauhoihientai <= 10 Then
sndPlaySound App.Path + "\sounds\right2.wav", SND_ASYNC
Else
sndPlaySound App.Path + "\sounds\right3.wav", SND_ASYNC
End If
End Sub

Sub NhapNhayDapAn()
If dapan.Caption = "a" Then
backA.Picture = LoadPicture(App.Path + "\pic\dapanTrue.gif")
ElseIf dapan.Caption = "b" Then
backB.Picture = LoadPicture(App.Path + "\pic\dapanTrue.gif")
ElseIf dapan.Caption = "c" Then
backC.Picture = LoadPicture(App.Path + "\pic\dapanTrue.gif")
Else
backD.Picture = LoadPicture(App.Path + "\pic\dapanTrue.gif")
End If
timerNhapNhay.Enabled = True
End Sub
Private Sub timerNhapNhay_Timer()   'Timer nhap nhay cau tra loi dung. Dong thoi kiem tra khi co dap an dung se ra cau hoi moi
dem = dem + 1
If dem Mod 2 = 1 And dem < 7 Then
    If dapanhientai = "a" Then frontA.Visible = True
    If dapanhientai = "b" Then frontB.Visible = True
    If dapanhientai = "c" Then frontC.Visible = True
    If dapanhientai = "d" Then frontD.Visible = True
ElseIf dem Mod 2 = 0 And dem < 7 Then
    If dapanhientai = "a" Then frontA.Visible = False
    If dapanhientai = "b" Then frontB.Visible = False
    If dapanhientai = "c" Then frontC.Visible = False
    If dapanhientai = "d" Then frontD.Visible = False
End If
If dem > 15 Then             'sau khi nhap nhay xong 1 luc
timerNhapNhay.Enabled = False
dem = 0
If phuonganhientai = dapanhientai Then  'neu nguoi choi chon dung thi an cau hoi va tiep tuc
backA.Picture = LoadPicture(App.Path + "\pic\dapanSelected.gif")
backB.Picture = LoadPicture(App.Path + "\pic\dapanSelected.gif")
backC.Picture = LoadPicture(App.Path + "\pic\dapanSelected.gif")
backD.Picture = LoadPicture(App.Path + "\pic\dapanSelected.gif")
Call AnCauHoi
Call KiemTra510
setTextScore (cauhoihientai)
Call BatDauCauHoi(cauhoihientai)
End If
End If
End Sub

Sub HienCauHoi()            'Hien cau hoi
If cauhoihientai <= 15 Then timerCauHoi.Enabled = True
isPaused = False
'show thang diem
If cauhoihientai = 1 Or cauhoihientai = 6 Or cauhoihientai = 11 Then
Unload frmThangDiem
frmThangDiem.Show
frmThangDiem.Left = Screen.Width
End If
End Sub
Private Sub timerCauHoi_Timer()         'hien thi dan dan cau hoi
dem = dem + 1
If dem < 2 Then
markCau.Visible = True
btnSure.Enabled = True
ElseIf dem < 3 Then
lblCauhoi.Visible = True
ElseIf dem < 4 Then
lblA.Visible = True
markA.Visible = True
ElseIf dem < 5 Then
lblB.Visible = True
markB.Visible = True
ElseIf dem < 6 Then
lblC.Visible = True
markC.Visible = True
ElseIf dem < 7 Then
lblD.Visible = True
markD.Visible = True
Else
dem = 0
timerCauHoi.Enabled = False
End If
End Sub

Sub AnCauHoi()              'An cau hoi va 4 dap an
lblCauhoi.Visible = False
lblA.Visible = False
lblB.Visible = False
lblC.Visible = False
lblD.Visible = False
markCau.Visible = False
markA.Visible = False
markB.Visible = False
markC.Visible = False
markD.Visible = False
                        'Lam luon viec. Reset phuong an da chon tu cau truoc
frontA.Visible = True
frontB.Visible = True
frontC.Visible = True
frontD.Visible = True
phuonganhientai = ""
End Sub

Function isAdded(a As Integer, dau As Integer, cuoi As Integer) As Boolean
isAdded = False
For i = dau To cuoi
If a = cauhoi(i) Then isAdded = True
Next
End Function

Function setTextScore(cau As Integer)   'set text so diem hien tai
markCau.Caption = "C©u" + Str(cau) + ":"
lblScore.Caption = "C©u" + Str(cau) + "  -  $" + Str(muctien(cau)) + "00"
End Function

Sub SetMucTienThuong()
muctien(1) = 2  '(1) - cau1
muctien(2) = 4  'don vi: tram ngan
muctien(3) = 6
muctien(4) = 10
muctien(5) = 20
muctien(6) = 30
muctien(7) = 60
muctien(8) = 100
muctien(9) = 140
muctien(10) = 220
muctien(11) = 300
muctien(12) = 400
muctien(13) = 600
muctien(14) = 850
muctien(15) = 1500
End Sub

Private Sub lblA_Click()
If isPaused = False Then 'tranh' click lung tung khi tam. dung` game de? phat' nhac.
frontA.Visible = False
frontB.Visible = True
frontC.Visible = True
frontD.Visible = True
phuonganhientai = "a"
End If
End Sub
Private Sub lblB_Click()
If isPaused = False Then
frontA.Visible = True
frontB.Visible = False
frontC.Visible = True
frontD.Visible = True
phuonganhientai = "b"
End If
End Sub
Private Sub lblC_Click()
If isPaused = False Then
frontA.Visible = True
frontB.Visible = True
frontC.Visible = False
frontD.Visible = True
phuonganhientai = "c"
End If
End Sub
Private Sub lblD_Click()
If isPaused = False Then
frontA.Visible = True
frontB.Visible = True
frontC.Visible = True
frontD.Visible = False
 phuonganhientai = "d"
 End If
End Sub

Public Sub disconnect() 'close connection database
recDe.Close
recThuong.Close
recKho.Close
con.Close
End Sub
Sub setDataFields() 'set datafields
lblCauhoi.DataField = "cauhoi"
lblA.DataField = "a"
lblB.DataField = "b"
lblC.DataField = "c"
lblD.DataField = "d"
dapan.DataField = "dapan"
End Sub
