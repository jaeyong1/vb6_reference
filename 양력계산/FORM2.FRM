VERSION 5.00
Begin VB.Form 거꾸로양력계산 
   Caption         =   "양력계수 검산"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   4590
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command3 
      Caption         =   "검산식으로 바로가기"
      Height          =   375
      Left            =   1680
      TabIndex        =   19
      Top             =   4200
      Width           =   3135
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "계 산 시 작"
      Default         =   -1  'True
      Height          =   615
      Left            =   4440
      TabIndex        =   18
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtV 
      Height          =   270
      Left            =   1800
      TabIndex        =   17
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtair 
      Height          =   270
      Left            =   1800
      TabIndex        =   16
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtS 
      Height          =   270
      Left            =   1800
      TabIndex        =   15
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtCl 
      Height          =   270
      Left            =   1800
      TabIndex        =   14
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "종  료"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "깔끄미~"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label14 
      Caption         =   "양력 (L) = 무게(W) ="
      BeginProperty Font 
         Name            =   "궁서"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblL 
      BeginProperty Font 
         Name            =   "바탕"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2880
      TabIndex        =   12
      Top             =   3000
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "양력계수(Cl)"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "날개면적(S)"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "공기밀도(p)"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "속   도(V) "
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '투명
      Caption         =   "cm"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '투명
      Caption         =   "kg/m"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "m/s"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '투명
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '투명
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  '단일 고정
      Height          =   1935
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   5775
   End
End
Attribute VB_Name = "거꾸로양력계산"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRun_Click()   '계산시작버튼
Dim w As Double   '무게
Dim S As Double   '면적
Dim p As Double   '공기밀도
Dim V As Double   '속도
Dim wl As Double  '윙로딩
Dim q As Double   '동압

Dim ll As Double  'CL
Dim sl As Double  'Cl
Dim Cl As Double  '양력계수

'입력확인1  ->  빈칸확인
If txtCl = "" Then MsgBox "빈칸을 채워주세요!": txtCl.SetFocus: Exit Sub
If txtS = "" Then MsgBox "빈칸을 채워주세요!": txtS.SetFocus: Exit Sub
If txtair = "" Then MsgBox "빈칸을 채워주세요!":  txtair.SetFocus: Exit Sub
If txtV = "" Then MsgBox "빈칸을 채워주세요!": txtV.SetFocus: Exit Sub

w = 0: S = 0: p = 0: V = 0: wl = 0: q = 0: b = 0: c = 0: ll = 0: sl = 0: Cl = 0:  '계산값 초기화
lblDAP = ""  '재 연산시 저번출력값 삭제..
lblrecl = ""  '        "
Cl = Val(Trim(txtCl))   '변수에 대입
S = Val(Trim(txtS))
p = Val(Trim(txtair))
V = Val(Trim(txtV))

'입력확인2  ->  숫자가 아니거나 0이하의 값들 확인요청
If Not IsNumeric(txtCl) Then MsgBox "값을 확인해주세요": txtCl.SetFocus: Exit Sub
If Not IsNumeric(txtS) Or (S < 0) Then MsgBox "값을 확인해주세요": txtS.SetFocus: Exit Sub
If Not IsNumeric(txtair) Or (p < 0) Then MsgBox "값을 확인해주세요": txtair.SetFocus: Exit Sub
If Not IsNumeric(txtV) Or (V < 0) Then MsgBox "값을 확인해주세요": txtV.SetFocus: Exit Sub

'연산시작
q = 0.5 * p * V * V
L = Cl * q * S
lblL = Format(L / 10000, "0.########0")

'양력계수 프로그램으로 값되돌림
Form1.txtS.Text = S
Form1.txtair.Text = p
Form1.txtV.Text = V

End Sub

Private Sub Command1_Click()  '종료버튼
Unload Me
시작.Show
End Sub

Private Sub Command3_Click()  '계산식으로 바로가기
Unload Me
Form1.Show
End Sub

Private Sub Form_Activate()  '프로그램 활동시작
txtCl.SetFocus  '처음입력칸으로 커서보내기
End Sub

