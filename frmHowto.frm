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
      Caption         =   "LuËt ch¬i"
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
Text1.Text = "- Ng­êi ch¬i ph¶i tr¶ lêi 15 c©u hái tr¾c nghiÖm tõ dÔ ®Õn khã." + vbCrLf + "- Cã 3 mèc c©u hái quan träng: 5 - 10 - 15. Khi ng­êi ch¬i tr¶ lêi sai th× sè tiÒn nhËn ®­îc sÏ bÞ xuèng mèc gÇn nhÊt. Tr­êng hîp ng­êi ch¬i Dõng cuéc ch¬i th× sè tiÒn sÏ ®­îc b¶o toµn." + vbCrLf + "- Ng­êi ch¬i cã 3 sù trî gióp:" + vbCrLf + "   + 50:50: M¸y tÝnh sÏ lo¹i ®i 2 ph­¬ng ¸n sai." + vbCrLf + "   + Gäi ®iÖn tho¹i cho ng­êi th©n: Nhê ng­êi th©n trî gióp (TØ lÖ tr¶ lêi ®óng phô thuéc vµo ®é khã cña c©u hái)." + vbCrLf + "   + Hái ý kiÕn kh¸n gi¶: Nhê kh¸n gi¶ trong tr­êng quay trî gióp." + vbCrLf + "B¾t ®Çu tõ c©u sè 6, b¹n cã thªm sù trî gióp Hái tæ t­ vÊn t¹i chç."
End Sub

