VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "그리미"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   273
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   397
   StartUpPosition =   3  'Windows 기본값
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "그리미.frx":0000
      Left            =   2160
      List            =   "그리미.frx":0034
      TabIndex        =   4
      Text            =   "0"
      Top             =   3780
      Width           =   675
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "그리미.frx":006E
      Left            =   660
      List            =   "그리미.frx":0081
      TabIndex        =   1
      Text            =   "1"
      Top             =   3780
      Width           =   750
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3735
      Left            =   60
      ScaleHeight     =   245
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   385
      TabIndex        =   0
      Top             =   0
      Width           =   5835
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "색상"
      Height          =   180
      Left            =   1620
      TabIndex        =   3
      Top             =   3840
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "선굵기"
      Height          =   180
      Left            =   60
      TabIndex        =   2
      Top             =   3840
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Picture1.PSet (10, 10), QBColor(15)
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Picture1.Line (X, Y)-(X + Val(Combo1.Text) - 1, Y + Val(Combo1.Text) - 1), QBColor(Val(Combo2.Text)), BF
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Picture1.Line (X, Y)-(X + Val(Combo1.Text) - 1, Y + Val(Combo1.Text) - 1), QBColor(Val(Combo2.Text)), BF
    End If
End Sub
