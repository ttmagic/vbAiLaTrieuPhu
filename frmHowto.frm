VERSION 5.00
Begin VB.Form frmHowto 
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   Caption         =   "How to play"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9285
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   7440
      Width           =   9015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   6615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   9015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lu�t ch�i"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00025DFF&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmHowto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmMenu.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = "- Ng��i ch�i ph�i tr� l�i 15 c�u h�i tr�c nghi�m t� d� ��n kh�." + vbCrLf + "- C� 3 m�c c�u h�i quan tr�ng: 5 - 10 - 15. Khi ng��i ch�i tr� l�i sai th� s� ti�n nh�n ���c s� b� xu�ng m�c g�n nh�t. Tr��ng h�p ng��i ch�i D�ng cu�c ch�i th� s� ti�n s� ���c b�o to�n." + vbCrLf + "- Ng��i ch�i c� 3 s� tr� gi�p:" + vbCrLf + "   + 50:50: M�y t�nh s� lo�i �i 2 ph��ng �n sai." + vbCrLf + "   + G�i �i�n tho�i cho ng��i th�n: Nh� ng��i th�n tr� gi�p (T� l� tr� l�i ��ng ph� thu�c v�o �� kh� c�a c�u h�i)." + vbCrLf + "   + H�i � ki�n kh�n gi�: Nh� kh�n gi� trong tr��ng quay tr� gi�p." + vbCrLf + "B�t ��u t� c�u s� 6, b�n c� th�m s� tr� gi�p H�i t� t� v�n t�i ch�."
End Sub

