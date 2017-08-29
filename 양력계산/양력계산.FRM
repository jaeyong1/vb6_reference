VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "양력계수계산"
   ClientHeight    =   4590
   ClientLeft      =   1935
   ClientTop       =   1695
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   6540
   Begin VB.CommandButton Command3 
      Caption         =   "검산식으로 바로가기"
      Height          =   375
      Left            =   1680
      TabIndex        =   20
      Top             =   4200
      Width           =   3135
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "계 산 시 작"
      Default         =   -1  'True
      Height          =   615
      Left            =   4680
      TabIndex        =   19
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtV 
      Height          =   270
      Left            =   1920
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtair 
      Height          =   270
      Left            =   1920
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtS 
      Height          =   270
      Left            =   1920
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtweight 
      Height          =   270
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "깔끄미~"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "종  료"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   4200
      Width           =   1095
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
      Left            =   3960
      TabIndex        =   18
      Top             =   960
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
      Left            =   3840
      TabIndex        =   17
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "m/s"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '투명
      Caption         =   "kg/m"
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '투명
      Caption         =   "cm"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "N"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "양력계수 (Cl) = "
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
      Left            =   840
      TabIndex        =   12
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblDAP 
      BeginProperty Font 
         Name            =   "바탕"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  '단일 고정
      Height          =   1935
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   5775
   End
   Begin VB.Label Label4 
      Caption         =   "속   도(V) "
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "공기밀도(p)"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "날개면적(S)"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "무   게(W)  "
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRun_Click()
Dim w As Double   '무게
Dim S As Double   '면적
Dim p As Double   '공기밀도
Dim V As Double   '속도
Dim wl As Double  '윙로딩
Dim q As Double   '동압

Dim b As Double
Dim c As Double
Dim ll As Double  'CL
Dim sl As Double  'Cl
Dim Cl As Double  '양력계수

'입력확인1
If txtweight = "" Then MsgBox "빈칸을 채워주세요!": txtweight.SetFocus: Exit Sub
If txtS = "" Then MsgBox "빈칸을 채워주세요!": txtS.SetFocus: Exit Sub
If txtair = "" Then MsgBox "빈칸을 채워주세요!":  txtair.SetFocus: Exit Sub
If txtV = "" Then MsgBox "빈칸을 채워주세요!": txtV.SetFocus: Exit Sub

w = 0: S = 0: p = 0: V = 0: wl = 0: q = 0: b = 0: c = 0: ll = 0: sl = 0: Cl = 0
lblDAP = ""
lblrecl = ""
w = Val(Trim(txtweight))
S = Val(Trim(txtS))
p = Val(Trim(txtair))
V = Val(Trim(txtV))
'be = Val(Trim(txtbe))

'입력확인2
If Not IsNumeric(txtweight) Or (w < 0) Then MsgBox "값을 확인해주세요": txtweight.SetFocus: Exit Sub
If Not IsNumeric(txtS) Or (S < 0) Then MsgBox "값을 확인해주세요": txtS.SetFocus: Exit Sub
If Not IsNumeric(txtair) Or (p < 0) Then MsgBox "값을 확인해주세요": txtair.SetFocus: Exit Sub
If Not IsNumeric(txtV) Or (V < 0) Then MsgBox "값을 확인해주세요": txtV.SetFocus: Exit Sub

'입력값을 검산프로그램으로 보내기
거꾸로양력계산.txtS.Text = S
거꾸로양력계산.txtair.Text = p
거꾸로양력계산.txtV.Text = V

'양력값 계산
q = 0.5 * p * V * V
wl = w / S
S = w / wl
bj = 2 * w
bm = p * V * V * S
Cl = Format((2 * w) / (p * V * V * S), "0.#########0") * 10000
lblDAP = Cl  'Cl 출력
End Sub

Private Sub Command1_Click() '종료번튼 작동
Unload Me
시작.Show
End Sub

Private Sub Command2_Click()  '깔끄미버튼
txtweight.Text = ""
txtS.Text = ""
txtV.Text = ""
txtair.Text = ""
'txtbe.Text = ""
lblDAP.Caption = ""

txtweight.SetFocus  '처음입력칸으로 커서이동

End Sub

Private Sub Command3_Click()  '검산으로 바로가기
거꾸로양력계산.Show
Unload Me
End Sub

Private Sub Form_Activate()   '프로그램 활동시작
txtweight.SetFocus
End Sub

