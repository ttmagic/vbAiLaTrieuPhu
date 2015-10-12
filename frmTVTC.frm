VERSION 5.00
Begin VB.Form frmTVTC 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmTVTC.frx":0000
   ScaleHeight     =   5415
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   240
      Top             =   1680
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   4935
   End
   Begin VB.Label pa3 
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      ForeColor       =   &H0000C1FF&
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label pa2 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      ForeColor       =   &H0000C1FF&
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label pa1 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H0000C1FF&
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.Label kg3 
      BackStyle       =   0  'Transparent
      Caption         =   "Kh¸n gi¶ 3:"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label kg2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kh¸n gi¶ 2:"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label kg1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kh¸n gi¶ 1:"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.Image img3 
      Height          =   945
      Left            =   960
      Picture         =   "frmTVTC.frx":11128
      Top             =   3360
      Width           =   7260
   End
   Begin VB.Image img2 
      Height          =   945
      Left            =   840
      Picture         =   "frmTVTC.frx":116FD
      Top             =   2040
      Width           =   7260
   End
   Begin VB.Image img1 
      Height          =   945
      Left            =   840
      Picture         =   "frmTVTC.frx":11CD2
      Top             =   720
      Width           =   7260
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "hái tæ t­ vÊn t¹i chç"
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
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmTVTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dem As Integer, cout As Integer
Dim ran1, ran2, ran3 As String
Dim dapan As String
Private Sub btnOK_Click()
frmGame.Show
frmGame.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
'Set position
Me.Top = 15 * 70
Me.Left = 15 * 800
dem = 0
cout = 0
dapan = frmGame.dapanhientai
Call AnCacPhuongAn
Call RandomCauTraLoi
Call setCauTraLoi
timer.Enabled = True
End Sub

Sub setCauTraLoi()
pa1.Caption = UCase(ran1)
pa2.Caption = UCase(ran2)
pa3.Caption = UCase(ran3)
End Sub

Sub AnCacPhuongAn() ' lam cho cac dap an' ra khoi? man` hinh`
img1.Left = img1.Left + 5000
img2.Left = img2.Left + 5000
img3.Left = img3.Left + 5000
kg1.Left = kg1.Left + 5000
kg2.Left = kg2.Left + 5000
kg3.Left = kg3.Left + 5000
pa1.Left = pa1.Left + 5000
pa2.Left = pa2.Left + 5000
pa3.Left = pa3.Left + 5000

End Sub

Private Sub timer_Timer()
dem = dem + 1
If dem < 20 Then
img1.Left = img1.Left - 250
kg1.Left = kg1.Left - 250
pa1.Left = pa1.Left - 250
ElseIf dem < 39 Then
img2.Left = img2.Left - 250
kg2.Left = kg2.Left - 250
pa2.Left = pa2.Left - 250
ElseIf dem < 58 Then
img3.Left = img3.Left - 250
kg3.Left = kg3.Left - 250
pa3.Left = pa3.Left - 250
Else
timer.Enabled = False
btnOK.Enabled = True
End If
End Sub

Sub RandomCauTraLoi()
Dim a, b, c As Byte         'a,b,c : 1...11
a = Int((Rnd * 10) + 1)     ' 1--8: dung, 9-11:sai
b = Int((Rnd * 10) + 1)
c = Int((Rnd * 10) + 1)

If dapan = "a" Then
    If a <= 8 Then
    ran1 = "a"
    ElseIf a <= 9 Then
    ran1 = "b"
    ElseIf a <= 8 Then
    ran1 = "c"
    ElseIf a <= 11 Then
    ran1 = "d"
    End If

    If b <= 8 Then
    ran2 = "a"
    ElseIf b <= 9 Then
    ran2 = "b"
    ElseIf b <= 10 Then
    ran2 = "c"
    ElseIf b <= 11 Then
    ran2 = "d"
    End If

    If c <= 8 Then
    ran3 = "a"
    ElseIf c <= 9 Then
    ran3 = "b"
    ElseIf c <= 10 Then
    ran3 = "c"
    ElseIf c <= 11 Then
    ran3 = "d"
    End If
ElseIf dapan = "b" Then
    If a <= 8 Then
    ran1 = "b"
    ElseIf a <= 9 Then
    ran1 = "a"
    ElseIf a <= 10 Then
    ran1 = "c"
    ElseIf a <= 11 Then
    ran1 = "d"
    End If

    If b <= 8 Then
    ran2 = "b"
    ElseIf b <= 9 Then
    ran2 = "a"
    ElseIf b <= 10 Then
    ran2 = "c"
    ElseIf b <= 11 Then
    ran2 = "d"
    End If

    If c <= 8 Then
    ran3 = "b"
    ElseIf c <= 9 Then
    ran3 = "a"
    ElseIf c <= 10 Then
    ran3 = "c"
    ElseIf c <= 11 Then
    ran3 = "d"
    End If
ElseIf dapan = "c" Then
    If a <= 8 Then
    ran1 = "c"
    ElseIf a <= 9 Then
    ran1 = "b"
    ElseIf a <= 10 Then
    ran1 = "a"
    ElseIf a <= 11 Then
    ran1 = "d"
    End If

    If b <= 8 Then
    ran2 = "c"
    ElseIf b <= 9 Then
    ran2 = "b"
    ElseIf b <= 10 Then
    ran2 = "a"
    ElseIf b <= 11 Then
    ran2 = "d"
    End If

    If c <= 8 Then
    ran3 = "c"
    ElseIf c <= 9 Then
    ran3 = "b"
    ElseIf c <= 10 Then
    ran3 = "a"
    ElseIf c <= 11 Then
    ran3 = "d"
    End If
ElseIf dapan = "d" Then
    If a <= 8 Then
    ran1 = "d"
    ElseIf a <= 9 Then
    ran1 = "b"
    ElseIf a <= 10 Then
    ran1 = "c"
    ElseIf a <= 11 Then
    ran1 = "a"
    End If

    If b <= 8 Then
    ran2 = "d"
    ElseIf b <= 9 Then
    ran2 = "b"
    ElseIf b <= 10 Then
    ran2 = "c"
    ElseIf b <= 11 Then
    ran2 = "a"
    End If

    If c <= 8 Then
    ran3 = "d"
    ElseIf c <= 9 Then
    ran3 = "b"
    ElseIf c <= 10 Then
    ran3 = "c"
    ElseIf c <= 11 Then
    ran3 = "a"
    End If
End If
End Sub

