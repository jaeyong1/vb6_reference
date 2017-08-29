VERSION 5.00
Begin VB.Form 시작 
   Caption         =   "에어디자인"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form3"
   ScaleHeight     =   4590
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command6 
      Caption         =   "      에 어 포 일       둘 레 계 산"
      Height          =   735
      Left            =   2640
      TabIndex        =   7
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "공기밀도계산"
      Height          =   735
      Left            =   2640
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "     필요마력계산     (속도)"
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "종  료"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3720
      Width           =   5775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "양력계수 검산"
      Height          =   735
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "양력계산"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Made in Korea  박재용"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "<- 서로 왔다갔다 가능 ->"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "시작"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Form1.Show
End Sub

Private Sub Command2_Click()
Unload Me
거꾸로양력계산.Show
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command6_Click()
Unload Me
frmdul.Show
End Sub
