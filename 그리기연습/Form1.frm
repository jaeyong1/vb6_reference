VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command4 
      Caption         =   "돌기"
      Height          =   615
      Left            =   6480
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "지우기"
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "선"
      Height          =   615
      Left            =   6480
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   4815
      Left            =   120
      ScaleHeight     =   4755
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   240
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "그리기연습"
      Height          =   615
      Left            =   6480
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '단일 고정
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   6120
      X2              =   6120
      Y1              =   3600
      Y2              =   4560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'위에서부터 이해해 나갈것!! 밑으로 갈수록 어려워지게..

'파이값 정의
Const pi = 3.14159265358979

'시계바늘 연습할때 각도 기억,
'함수안에 만든 지역변수는 함수끝나면
'값도 사라지므로 폼이 있는동안 계속 기억될수 있게
'여기에서 선언
Dim NowAngle As Integer


Private Sub Command1_Click() '그리기연습 버튼 누르면

' 점그리기
Picture1.DrawWidth = 10 '점 굵기
Picture1.PSet (500, 300), RGB(80, 100, 255) '(x,y)위치에 RGB색깔로 찍음


'선 그리기
Picture1.DrawWidth = 2 '선 굵기변경 (위에서 점찍는다고 너무 크게 했음)
Picture1.Line (1000, 1000)-(1500, 1500), RGB(100, 200, 0)


'원그리기
Picture1.Circle (3000, 1200), 300, RGB(180, 120, 255)



End Sub

Private Sub Command2_Click() '선 버튼 누르면
'위치 찍어주기
'Line1.X1 = 6000
'Line1.Y1 = 6000
'Line1.X2 = 200
'Line1.Y2 = 5000

'위치를 누적시키면..
Line1.X1 = Line1.X1 + 100
Line1.X2 = Line1.X2 + 100

End Sub

Private Sub Command3_Click() '지우기 버튼

'전부 지우기
Picture1.Picture = Nothing

End Sub


Private Sub Command4_Click() '바늘돌리기

NowAngle = NowAngle + 10 '현재 가리키는값 360도 기준

Label1 = NowAngle '레이블에 각 표시

'x1,y1은 중심점
'x2,y2는 변화는 지점
Line1.X1 = 6800 '임의의 지점
Line1.Y1 = 4800 '임의 지점
Line1.X2 = 6800 + 1000 * Cos((90 - NowAngle) * (pi / 180))
Line1.Y2 = 4800 - 1000 * Cos(NowAngle * (pi / 180))

End Sub


