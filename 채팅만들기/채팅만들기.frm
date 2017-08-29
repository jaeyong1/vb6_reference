VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "채팅 서버"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox text1 
      Height          =   6975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   7
      Top             =   480
      Width           =   4935
   End
   Begin VB.TextBox txtinput 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   7800
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "send"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "접속해제"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "프로그램 종료"
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      Top             =   8400
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   5760
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "~"
      Height          =   135
      Left            =   6360
      TabIndex        =   9
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "대화표시창을 누르면 채팅창이 깨끗하게 됨..."
      Height          =   1095
      Left            =   5640
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblconnect 
      Caption         =   "클라이언트 접속됨.."
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   8520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label ipview 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "서버측.."
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() '프로그램 종료
End
End Sub

Private Sub Command2_Click()  '끊기
Command2.Enabled = False
Command3.Enabled = False
txtinput.Enabled = False
text1.Enabled = False



Winsock1(1).Close
End Sub

Private Sub Command3_Click()  '글입력후 엔터
Winsock1(1).SendData txtinput.Text
 text1.Text = text1.Text + "나> " & txtinput.Text + vbNewLine
 txtinput.SetFocus
 txtinput.Text = ""
End Sub


Private Sub Form_Load()       '시작

Load Winsock1(0)
Winsock1(0).Protocol = sckTCPProtocol
Winsock1(0).LocalPort = 2000
Winsock1(0).Listen                '클라이언트 접속대기상태

End Sub

Private Sub Label3_Click()
MsgBox "<<사악한 치트코드>>" & vbCrLf & "  재용 시스템종료" & vbCrLf & "  메세지박스"
End Sub

Private Sub text1_Click()     '표시창 깔끔~
text1 = ""
txtinput.SetFocus

End Sub


Private Sub Winsock1_Close(Index As Integer)  '끊기
Winsock1(1).Close

End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
                                            '클라이언트에서 접속요구가 들어올경우

  'Load Winsock1(0)
  Winsock1(1).Accept requestID  '접속허락
  
  
  Command3.Enabled = True 'send 버튼 사용가
  txtinput.Enabled = True '입력창 사용가
  txtinput.SetFocus   '입력창에 커서이동
  lblconnect.Enabled = True  ' 클, 접속중 표시
  
  Command2.Enabled = True  '접속해제 버튼사용가
  text1.Enabled = True  '채팅창 클릭으로 깨끗가
  
  
 
 'Winsock1(Index).SendData "!!!" + CStr(Index)
 
End Sub
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)   '글도착
    Dim Gstr As String
    Winsock1(1).GetData Gstr
    text1.Text = text1.Text + "너> " & Gstr + vbNewLine
    txtinput.SetFocus


End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "뭐꼬! 에러발생"
Command2_Click
End Sub
