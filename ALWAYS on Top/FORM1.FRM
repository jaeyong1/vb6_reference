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
   Begin VB.Label Label1 
      Caption         =   "Always on top | 항상위에"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Always on top 기능..
   Private Declare Function SetWindowPos Lib "user32" _
         ( _
          ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
          ByVal x As Long, ByVal y As Long, _
          ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long _
         ) As Long
   Private Const HWND_TOPMOST = -1
   Private Const HWND_NOTOPMOST = -2
   Private Const SWP_NOSIZE = &H1

   Private Sub Form_Load()
       Call SetWindowPos(Form1.hwnd, HWND_TOPMOST, 0, 0, 1, 1, SWP_NOSIZE)
   End Sub

' >>> 설명
'
'   첫번째 인수
'     Always On Top을 사용할 작업대상의 핸들입니다!
'     즉 Form1에 Always On Top기능을 사용한다는 뜻입니다!
'
'   두번째 인수
'     Always On Top을 사용할지의 여부입니다!
'     -1은 사용, -2는 사용불가입니다.''
'
'   세번째/네번째 인수
'     폼의 X좌표, Y좌표입니다!
'     0,0은 모니터의 좌표에서 0,0에 뜨게 됩니다!
'


