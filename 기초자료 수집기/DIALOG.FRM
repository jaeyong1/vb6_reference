VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   3645
   ClientLeft      =   1920
   ClientTop       =   1770
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "확인"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   4935
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Activate()
 Label1 = "안녕하세요.. 는 일종의 단순한 나의 분신(?)으로써 지금의 기본적인 시스템 설정을 조사해서 디스켓에 저장하는 역할을 하고요.. 나쁜짓을 하는 악랄한 프로그램이 아니여요~ 대부분의 작업을 MS-DOS창을 열어 사용하기때문에 한글MS-DOS가 안된다면 아쉽지만 작동이 안되네여.. 이 프로그램 종료후 남은 MS-DOS창은 X표시로 닫아주세요.. 디스켓의 쓰기방지 구멍을 막았는가 확인해 주세요.(막아주세요)"
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub
