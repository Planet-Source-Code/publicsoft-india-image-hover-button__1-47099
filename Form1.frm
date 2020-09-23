VERSION 5.00
Object = "*\AimgButton.vbp"
Begin VB.Form frmImgContrl 
   Caption         =   "Image Control"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin imgTest.ImageButton ImageButton1 
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1508
      NormalPic       =   "Form1.frx":0000
      DownPic         =   "Form1.frx":03A2
      HoverPic        =   "Form1.frx":070F
      Stretch         =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "By Cyril M Gupta: cyril@cyrilgupta.com"
      Height          =   735
      Left            =   4920
      TabIndex        =   2
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "This is an Image hoverbutton."
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "frmImgContrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox imgButton1.CurPicture
End Sub

