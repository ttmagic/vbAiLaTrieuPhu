VERSION 5.00
Begin VB.Form frmAdd 
   BackColor       =   &H80000016&
   Caption         =   "Add questions"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnGo 
      Caption         =   "Go"
      Height          =   615
      Left            =   2880
      TabIndex        =   27
      Top             =   6480
      Width           =   615
   End
   Begin VB.TextBox txtGoto 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   26
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton btnLast 
      Caption         =   ">l"
      Height          =   615
      Left            =   1560
      TabIndex        =   25
      Top             =   6480
      Width           =   495
   End
   Begin VB.CommandButton btnNext 
      Caption         =   ">"
      Height          =   615
      Left            =   1080
      TabIndex        =   24
      Top             =   6480
      Width           =   495
   End
   Begin VB.CommandButton btnPrev 
      Caption         =   "<"
      Height          =   615
      Left            =   600
      TabIndex        =   23
      Top             =   6480
      Width           =   495
   End
   Begin VB.CommandButton btnFirst 
      Caption         =   "l<"
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   6480
      Width           =   495
   End
   Begin VB.CommandButton btnMinMaxMenu 
      Caption         =   "Min/Max MainMenu"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   7200
      Width           =   3375
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton btnKho 
      Caption         =   "Lo¹i khã"
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton btnThuong 
      Caption         =   "Lo¹i th­êng"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton btnDe 
      BackColor       =   &H0000C1FF&
      Caption         =   "Lo¹i dÔ"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txtDapan 
      Height          =   615
      Left            =   1440
      TabIndex        =   5
      Top             =   5640
      Width           =   7095
   End
   Begin VB.CommandButton btnXoa 
      Caption         =   "Xãa"
      Height          =   615
      Left            =   6240
      TabIndex        =   8
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton btnSua 
      Caption         =   "Söa"
      Height          =   615
      Left            =   5040
      TabIndex        =   7
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton btnThoat 
      Caption         =   "Tho¸t"
      Height          =   615
      Left            =   7440
      TabIndex        =   9
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton btnThem 
      Caption         =   "Thªm"
      Height          =   615
      Left            =   3840
      TabIndex        =   6
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox txtD 
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   4920
      Width           =   7815
   End
   Begin VB.TextBox txtC 
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   4200
      Width           =   7815
   End
   Begin VB.TextBox txtB 
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   3480
      Width           =   7815
   End
   Begin VB.TextBox txtA 
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   2760
      Width           =   7815
   End
   Begin VB.TextBox txtCauhoi 
      Height          =   1455
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   8295
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Chän b¶ng m· TCVN3 (ABC) ®Ó gâ TiÕng ViÖt"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   20
      Top             =   7320
      Width           =   5055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "§¸p ¸n:"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "D:"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "C:"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "B:"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "A:"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "C©u hái:"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rec As New ADODB.Recordset

Private Sub btnGo_Click()
On Error Resume Next
rec.MoveFirst
rec.Move (Val(txtGoto.Text) - 1)
End Sub

Private Sub Form_Load()
'connect db
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\ailatrieuphu.mdb;Persist Security Info=False;Jet OLEDB:Database Password=01653330406"
rec.CursorLocation = adUseClient
rec.Open "de", con, adOpenDynamic, adLockOptimistic
Call setResourceSource
'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VISUAL BASIC\PROJECT\AiLaTrieuPhu\ailatrieuphu.mdb;Persist Security Info=False;Jet OLEDB:Database Password=admin69
End Sub

Private Sub btnDe_Click()
btnDe.BackColor = &HC1FF&           'vang
btnThuong.BackColor = &H8000000F    'trang
btnKho.BackColor = &H8000000F       'trang
con.Close
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\ailatrieuphu.mdb;Persist Security Info=False;Jet OLEDB:Database Password=admin69"
rec.CursorLocation = adUseClient
rec.Open "de", con, adOpenDynamic, adLockOptimistic
Call setResourceSource
End Sub

Private Sub btnThuong_Click()
btnThuong.BackColor = &HC1FF&           'vang
btnDe.BackColor = &H8000000F            'trang
btnKho.BackColor = &H8000000F           'trang
con.Close
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\ailatrieuphu.mdb;Persist Security Info=False;Jet OLEDB:Database Password=admin69"
rec.CursorLocation = adUseClient
rec.Open "thuong", con, adOpenDynamic, adLockOptimistic
Call setResourceSource
End Sub

Private Sub btnKho_Click()
btnKho.BackColor = &HC1FF&           'vang
btnThuong.BackColor = &H8000000F     'trang
btnDe.BackColor = &H8000000F         'trang
con.Close
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\ailatrieuphu.mdb;Persist Security Info=False;Jet OLEDB:Database Password=admin69"
rec.CursorLocation = adUseClient
rec.Open "kho", con, adOpenDynamic, adLockOptimistic
Call setResourceSource
End Sub


Private Sub btnFirst_Click()
rec.MoveFirst
End Sub
Private Sub btnPrev_Click()
If rec.BOF = False Then rec.MovePrevious
End Sub
Private Sub btnNext_Click()
If rec.EOF = False Then rec.MoveNext
End Sub
Private Sub btnLast_Click()
rec.MoveLast
End Sub


Private Sub btnThem_Click()
rec.AddNew
txtCauhoi.SetFocus
End Sub
Private Sub btnSua_Click()
rec.Update
End Sub
Private Sub btnXoa_Click()
Dim warning As Integer
warning = MsgBox("Ban co muon xoa cau hoi nay khong?", vbCritical + vbYesNo, "Xac nhan")
If warning = vbYes Then
rec.Delete
rec.MovePrevious
End If
End Sub
Private Sub btnThoat_Click()
Unload Me
frmMenu.Show
frmMenu.Enabled = True
End Sub

'Su kien nhan nut X dong form
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMenu.WindowState = vbMaximized
frmMenu.Show
frmMenu.Enabled = True
End Sub

Private Sub btnMinMaxMenu_Click()
If frmMenu.WindowState = vbMaximized Then
frmMenu.WindowState = vbMinimized
Else
frmMenu.WindowState = vbMaximized
Me.Show
End If
End Sub

'Set Datasource va Datafields cho cac textbox
Sub setResourceSource()
Set txtID.DataSource = rec
txtID.DataField = "id"
Set txtCauhoi.DataSource = rec
txtCauhoi.DataField = "cauhoi"
Set txtA.DataSource = rec
txtA.DataField = "a"
Set txtB.DataSource = rec
txtB.DataField = "b"
Set txtC.DataSource = rec
txtC.DataField = "c"
Set txtD.DataSource = rec
txtD.DataField = "d"
Set txtDapan.DataSource = rec
txtDapan.DataField = "dapan"
End Sub


Private Sub txtGoto_Change()
btnGo.SetFocus
End Sub
