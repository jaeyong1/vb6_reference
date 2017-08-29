VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4080
      Top             =   120
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3600
      Left            =   120
      ScaleHeight     =   3540
      ScaleWidth      =   4740
      TabIndex        =   0
      Top             =   600
      Width           =   4800
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Set Image1 = LoadPicture("c:\a.bmp")


End Sub

Private Sub Picture1_Click()
Set Picture1 = LoadPicture("c:\a.bmp")
Picture1.Refresh

Dim e
Dim r, g, b

e = Picture1.Point(100, 100)




End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo onerrors
Dim a As Long
Dim r, b, c As Long
Set Picture1 = LoadPicture("c:\a.bmp")
'Picture1.Refresh



'가져오기
a = Picture1.Point(X, Y)

'RGB숫자로 표시
   r = a Mod 256

    g = (a \ 256) Mod 256

    b = (a \ 65536) Mod 256

Text1 = r
Text2 = g
Text3 = b

'색깔로 표시
Label1.BackColor = a

onerrors:

End Sub

