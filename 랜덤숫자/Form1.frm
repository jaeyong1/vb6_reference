VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "랜덤으로 숫자 띄울까..."
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
Dim temp As String



For i = 1 To 10

     temp = Int(Rnd(1) * 10)

    '이렇게 하면 0 부터 10까지 뽑게 되죠.
    '원래 Rnd는 0부터 1까지의 중간 숫자를 구합니다.
    '그럼 10을 써놓으면 거기에 10을 곱한 효과가 나겠죠?
     'int문은 그렇게 하면 소수 점이 나오기 때문에 소수점을 없애기 위한 것이고요...

    Text1 = Text1 & " " & temp
Next i

End Sub
