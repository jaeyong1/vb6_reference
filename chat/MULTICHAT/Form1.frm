VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Chatting Craft-Client"
   ClientHeight    =   4776
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6552
   LinkTopic       =   "Form1"
   ScaleHeight     =   4776
   ScaleWidth      =   6552
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox nick_bar 
      Height          =   372
      Left            =   960
      TabIndex        =   11
      Top             =   4920
      Width           =   1812
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6120
      Top             =   4320
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.TextBox port_bar 
      Height          =   372
      Left            =   960
      TabIndex        =   8
      Text            =   "2000"
      Top             =   6120
      Width           =   1812
   End
   Begin VB.TextBox ip_bar 
      Height          =   372
      Left            =   960
      TabIndex        =   7
      Text            =   "127.0.0.1"
      Top             =   5520
      Width           =   1812
   End
   Begin VB.CommandButton exit_com 
      Caption         =   "프로그램종료"
      Height          =   492
      Left            =   2880
      TabIndex        =   6
      Top             =   6000
      Width           =   3612
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   132
   End
   Begin VB.CommandButton dis_com 
      Caption         =   "끊   기"
      Height          =   492
      Left            =   4680
      TabIndex        =   3
      Top             =   5520
      Width           =   1812
   End
   Begin VB.CommandButton connect_com 
      Caption         =   "연   결"
      Height          =   492
      Left            =   2880
      TabIndex        =   2
      Top             =   5520
      Width           =   1812
   End
   Begin VB.TextBox chat_bar 
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   6372
   End
   Begin VB.TextBox chat_win 
      Height          =   3492
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   0
      Top             =   120
      Width           =   6372
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      Caption         =   "대화명"
      Height          =   252
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   732
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   "포트설정"
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   732
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "IP입력창"
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   732
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6360
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "설정창열기"
      Height          =   372
      Left            =   360
      TabIndex        =   5
      Top             =   4440
      Width           =   1452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim check As Integer
Dim ret, gointo, gotoval, i, counter As Integer
Dim tempid As String
Dim indexid As String


Private Sub dis_com_Click() '끊기버튼을 누를경우

 
 For i = 1 To 2
   If i = 1 Then            '끊기데이터를 먼저 서버에게 보낸다

        Winsock1.SendData "*X*"
        DoEvents
        

   Else                     '클라이언트의 소켓을닫는다

       Winsock1.Close
       connect_com.Enabled = True
       dis_com.Enabled = False
       ret = MsgBox("연결이 끊어졌습니다", 64, "연결상태")
       addtext "#서버와의 연결이 끊어졌습니다#"

   End If
Next i
 



End Sub

Sub Startrek(frm As Form)  '프로그램닫기 애니메이션

gotoval = Form1.Height / 2

For gointo = 1 To gotoval  ' 폼의 수직길이를 줄인다

DoEvents
Form1.Height = Form1.Height - 100
Form1.Top = (Screen.Height - Form1.Height) \ 2
If Form1.Height <= 500 Then Exit For

Next gointo

horiz:
Form1.Height = 30
gotoval = Form1.Width / 2

For gointo = 1 To gotoval  '폼의 수평길이를 줄인다

DoEvents
Form1.Width = Form1.Width - 100
Form1.Left = (Screen.Width - Form1.Width) \ 2
If Form1.Width <= 2000 Then Exit For

Next gointo

End Sub


Private Sub exit_com_Click() '프로그램종료 버튼을 누를경우

Call Startrek(Me)
End

End Sub


Private Sub connect_com_Click() '연결버튼을 누를경우

If ip_bar.Text = "" Or nick_bar.Text = "" Then  '대화명과 ip주소가 입력되지 않았을경우
   
   ret = MsgBox("서버측IP주소와 대화명을 입력요망", 64, "연결상태")

Else                                             '정상접속시도

   If port_bar <> 2000 Then                      'port번호가 바뀔경우
        ret = MsgBox("Port번호가" & port_bar.Text & "로변경되었습니다", 64, "연결상태")
        Winsock1.RemotePort = port_bar.Text
   End If
   Winsock1.RemoteHost = ip_bar.Text             '서버에 접속을 시도한다
   Winsock1.Connect
   connect_com.Enabled = False
   dis_com.Enabled = True
   addtext "#서버와 정상적으로 접속했습니다#"
   
End If

End Sub

Private Sub Form_Load()         '폼로드시 접속포트번호 초기화

Winsock1.RemotePort = 2000

End Sub

Private Sub chat_bar_keyPress(keyascii As Integer)     '채팅메세지 보낼경우


If keyascii = 13 And dis_com.Enabled = True Then       '연결상태에서 엔터를 누를경우
                                                  
      If tempid <> "" And tempid <> nick_bar.Text Then '대화명이 변경될경우 변경메세지 출력
         
            Winsock1.SendData "##########" + tempid + "가 " + nick_bar.Text + "로 대화명이 변경되었습니다" + "#########" + vbNewLine
            addtext "##########" + tempid + "가 " + nick_bar.Text + "로 대화명이 변경되었습니다" + "#########"
  
      End If                                           '채팅메세지 전송과 창에 표시
        
            tempid = nick_bar.Text
            Winsock1.SendData nick_bar.Text + ">>" + chat_bar.Text
            addtext nick_bar.Text + ">>" + chat_bar.Text
            chat_bar.Text = ""
      
      
End If

End Sub


Private Sub addtext(addline As String) '화면에 출력시킨다

chat_win.Text = chat_win.Text + addline + vbNewLine

End Sub

Private Sub Winsock1_Close()           '소켓을 닫는다

dis_com_Click

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
                                       '서버에서온 데이터를 받는다
Dim getda As String
Winsock1.GetData getda

If Left(getda, 3) <> "!!!" Then


Select Case getda

Case "boot"  '원격종료 명령어가 들어올경우

  ret = MsgBox("서버에 의해 윈도우가 자동종료됩니다", 52, "연결상태")
  
  If ret = 6 Then '확인버튼
       Call ExitWindowsEx(EWX_SHUTDOWN, 0) '윈도우종료함수호출
  End If
   
Case "logoff" '원격로그오프 명령어가 들어올경우
  
  ret = MsgBox("서버에 의해 윈도우가 자동로그오프됩니다", 52, "연결상태")
  
  If ret = 6 Then '확인버튼
       Call ExitWindowsEx(EWX_LOGOFF, 0) '윈도우로그오프함수호출
  End If
   
Case "reboot" '원격재시작 명령어가 들어올경우
  
  ret = MsgBox("서버에 의해 윈도우가 자동재시작됩니다", 52, "연결상태")
  
  If ret = 6 Then '확인버튼
       Call ExitWindowsEx(EWX_REBOOT, 0) '윈도우재시작함수호출
  End If
   

Case "#!#!"  '강제연결해제 명령이 들어올경우

       Winsock1.Close
       connect_com.Enabled = True
       dis_com.Enabled = False
       ret = MsgBox("연결이 끊어졌습니다", 64, "연결상태")
       addtext "#서버와의 연결이 끊어졌습니다#"

Case Else '채팅메세지가 들어올경우

addtext getda

End Select

Else

indexid = CStr(Right(getda, 1))

End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
                                        '접속시 에러발생
ret = MsgBox("연결시 에러발생했습니다 다시 시도하세요", 64, "연결상태")
Winsock1.Close
connect_com.Enabled = True
dis_com.Enabled = False

End Sub

Private Sub Check1_Click()  '설정창을 열때

If check = 0 Then     '설정창을 열경우
      Form1.Height = 6972
      check = check + 1
ElseIf check = 1 Then '설정창을 닫을경우
      Form1.Height = 5148
      check = 0
End If
 
End Sub
