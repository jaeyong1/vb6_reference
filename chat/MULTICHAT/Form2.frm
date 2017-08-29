VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BackColor       =   &H0080FF80&
   Caption         =   "Chatting Craft-Server"
   ClientHeight    =   4788
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8268
   LinkTopic       =   "Form2"
   ScaleHeight     =   4788
   ScaleWidth      =   8268
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CheckBox sCheck2 
      Caption         =   "Check1"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   5760
      Width           =   132
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   10
      Left            =   3000
      Top             =   4320
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.CommandButton slogoff_com 
      Caption         =   "원격로그오프"
      Height          =   492
      Left            =   3120
      TabIndex        =   11
      Top             =   5520
      Width           =   1692
   End
   Begin VB.CommandButton sreboot_com 
      Caption         =   "원격재부팅"
      Height          =   492
      Left            =   3120
      TabIndex        =   10
      Top             =   5040
      Width           =   1692
   End
   Begin VB.TextBox snick_bar 
      Height          =   372
      Left            =   720
      TabIndex        =   8
      Top             =   5040
      Width           =   2172
   End
   Begin VB.CommandButton sboot_com 
      Caption         =   "원격윈도우종료"
      Height          =   492
      Left            =   4800
      TabIndex        =   7
      Top             =   5520
      Width           =   1692
   End
   Begin VB.CommandButton Sconnect_com 
      Caption         =   "대기상태"
      Height          =   492
      Left            =   4800
      TabIndex        =   6
      Top             =   5040
      Width           =   1692
   End
   Begin VB.CommandButton sexit_com 
      Caption         =   "프로그램종료"
      Height          =   492
      Left            =   6480
      TabIndex        =   5
      Top             =   5520
      Width           =   1692
   End
   Begin VB.CommandButton sdis_com 
      Caption         =   "끊   기"
      Height          =   492
      Left            =   6480
      TabIndex        =   4
      Top             =   5040
      Width           =   1692
   End
   Begin VB.CheckBox sCheck1 
      Caption         =   "Check1"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   132
   End
   Begin VB.TextBox schat_bar 
      Height          =   372
      Left            =   0
      TabIndex        =   1
      Top             =   3840
      Width           =   8172
   End
   Begin VB.TextBox schat_win 
      ForeColor       =   &H80000001&
      Height          =   3492
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   0
      Top             =   120
      Width           =   8172
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "원격제어기능"
      Height          =   252
      Left            =   480
      TabIndex        =   13
      Top             =   5760
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "대화명"
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      Width           =   852
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8040
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "설정창열기"
      Height          =   252
      Left            =   360
      TabIndex        =   3
      Top             =   4440
      Width           =   972
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private connectindex(10) As Integer
Dim sscheck, gointo, gotoval As Integer
Dim j, i, countp As Integer
Dim tempid, tempstring As String

Private Sub Form_Load() ' 폼이 로드될때 초기화 하는 부분 각버튼의 초기화

sdis_com.Enabled = False
Sconnect_com.Enabled = True
sreboot_com.Enabled = False
sboot_com.Enabled = False
slogoff_com.Enabled = False
sexit_com.Enabled = True

For i = 1 To 10         '멀티체팅을 위한 소켓의 사용여부확인 배열 초기화
    
    connectindex(i) = 0

Next i


End Sub

Private Sub sboot_com_Click() '원격셧다운버튼을 누를때

For i = 1 To 10               '현재 연결되어있는 소켓을 검사하여 셧다운메시지 전송
    
    If connectindex(i) = 1 Then
          
          Winsock1(i).SendData "boot"
          DoEvents
    
    End If

Next i

End Sub

Private Sub sCheck1_Click() '설정창을 열때

If sscheck = 0 Then         '첫번째 체크시(설정)
   
   Form2.Height = 6684
   sscheck = sscheck + 1
 
 ElseIf sscheck = 1 Then    '두번째 체크시(해제)
   
   Form2.Height = 5184
   sscheck = 0
 
 End If

End Sub

Private Sub sCheck2_Click() '원격제어버튼을 사용코져 할때 각버튼 설정

If sscheck = 0 Then         '첫번째 체크시(설정)
   
      sreboot_com.Enabled = False
      sboot_com.Enabled = False
      slogoff_com.Enabled = False
      sscheck = sscheck + 1
 
 ElseIf sscheck = 1 Then
      
      sreboot_com.Enabled = True
      sboot_com.Enabled = True
      slogoff_com.Enabled = True
      sscheck = 0
 
 End If

End Sub

Private Sub Sconnect_com_Click() '연결을 누를때

Load Winsock1(0)                  '서버용 소켓을 생성하고 포트는2000번으로 한다
Winsock1(0).Protocol = sckTCPProtocol
Winsock1(0).LocalPort = 2000
Winsock1(0).Listen                '클라이언트 접속대기상태
sdis_com.Enabled = True
Sconnect_com.Enabled = False
addtext "#클라이언트의 접속을 대기중입니다#"

End Sub

Sub Startrek(frm As Form)  '종료시 창 에니메이션실행

gotoval = Form2.Height / 2

For gointo = 1 To gotoval  '수직으로 창크기를 줄인다
   
   DoEvents
   Form2.Height = Form2.Height - 100
   Form2.Top = (Screen.Height - Form2.Height) \ 2
   If Form2.Height <= 500 Then Exit For

Next gointo

horiz:
Form2.Height = 30
gotoval = Form2.Width / 2

For gointo = 1 To gotoval  '수평으로 창크기를 줄인다
 
   DoEvents
   Form2.Width = Form2.Width - 100
   Form2.Left = (Screen.Width - Form2.Width) \ 2
   If Form2.Width <= 2000 Then Exit For

Next gointo

End Sub

Private Sub sdis_com_Click() '모든 클라이언트 접속을 해제시킬경우

For i = 1 To 10
  
  If connectindex(i) = 1 Then  '접속되어있는 클라이언트들에게만 강제접속신호를 보낸다
    
    Winsock1(i).SendData "#!#!"
    DoEvents
    connectindex(i) = 0                '소켓배열을 0으로 초기화시킨다
    Winsock1(i).Close
    Unload Winsock1(i)
    addtext "#### " + CStr(i) + "번째 클라이언트가 접속해제 되었습니다####"
  
  
  End If

Next i

End Sub

Private Sub sexit_com_Click() '프로그램 종료버튼을 누를경우

Call Startrek(Me)             '종료에니메이션 실행
End

End Sub


Private Sub slogoff_com_Click() '원격로그오프버튼을 누를때

For i = 1 To 10                 '현재 연결되어있는 소켓을 검사하여 로그오프메시지 전송
    
    addtext CStr(connectindex(i))
    
    If connectindex(i) <> 0 Then
          
          Winsock1(i).SendData "logoff"
          DoEvents
    
    End If

Next i

End Sub

Private Sub sreboot_com_Click() '원격로그오프버튼을 누를때

For i = 1 To 10                 '현재 연결되어있는 소켓을 검사하여 재부팅메시지 전송
    
    If connectindex(i) <> 0 Then
          
          Winsock1(i).SendData "reboot"
          DoEvents
    
    End If

Next i

End Sub

Private Sub Winsock1_Close(Index As Integer) '소켓을 닫는다

connectindex(Index) = 0
Winsock1(Index).Close

End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
                                            '클라이언트에서 접속요구가 들어올경우
If countp <> 10 Then '9개이상의 클라이언트는 받지 않는다
  j = Index
  countp = countp + 1

  For i = 1 To 10    '현재 비어있는 소켓을 찾는다
     If connectindex(i) = 0 Then
          Index = i
          Exit For
      End If
  Next i
                      '비어있는 소켓에 클라이언트를 연결시킨다
  connectindex(Index) = 1
  Load Winsock1(Index)
  Winsock1(Index).Accept requestID
  j = Index
  addtext "#" + CStr(Index) + "번째 클라이언트가 접속했습니다#"

End If
 
 Winsock1(Index).SendData "!!!" + CStr(Index)
 
End Sub


Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
                                          '클라이언트에서 데이터를 보내올경우
Dim p As String
Winsock1(Index).GetData p                 '클라이언트에서 보내오는 데이터를 받는다

Select Case CStr(Left(p, 3))              '들어온데이터중 앞3자리가 *x* 이면 종료메세지로 인식

Case "*X*"                                '소켓을 닫는다
   
   connectindex(Index) = 0                '소켓배열을 0으로 초기화시킨다
   Winsock1(Index).Close
   Unload Winsock1(Index)
   addtext "#### " + CStr(Index) + "번째 클라이언트가 접속해제 되었습니다####"

Case Else                                 '그이외의 데이터는 채팅데이터로 간주한다
   
   For i = 1 To 10
                                          '들어온 데이터를 다른 클라이언트에게 보내준다
     If connectindex(i) = 1 And i <> Index Then
      
        Winsock1(i).SendData p
        DoEvents
        
     End If
   
   Next i
   
   addtext p                               '서버의 채팅창에 뿌려준다

End Select

End Sub

Private Sub schat_bar_keyPress(keyascii As Integer)
                                                          '서버에서 채팅메세지를 보낼때

If keyascii = 13 And sdis_com.Enabled = True Then         '접속상태에서 엔터를 칠경우

     If tempid <> "" And tempid <> snick_bar.Text Then    '대화명이 변경되었는지 확인
         
         For i = 1 To 10                                  '접속되어있는 클라이언트에게 대화명 변경을 전달한다
            
            If connectindex(i) = 1 Then
                     
                     Winsock1(i).SendData "##########" + tempid + "가 " + snick_bar.Text + "로 대화명이 변경되었습니다" + "#########" + vbNewLine
                     DoEvents
                     
            End If
         
         Next i                                            '서버측에도 변경내용을 보여준다
                     addtext "##########" + tempid + "가 " + snick_bar.Text + "로 대화명이 변경되었습니다" + "#########"
      End If
        
        For i = 1 To 10                                    '접속되어있는 클라이언트에게 대화내용을 전달한다
        If connectindex(i) = 1 Then
                    tempid = snick_bar.Text
                    Winsock1(i).SendData snick_bar.Text + ">>" + schat_bar.Text
                    DoEvents
                    
           End If
        Next i
                                                           '서버측에도 대화내용을 보여준다
      addtext snick_bar.Text + ">>" + schat_bar.Text
      schat_bar.Text = ""
End If

End Sub

Private Sub addtext(addline As String)      '들어온 내용을 화면에 출력시키는 함수


schat_win.Text = schat_win.Text + addline + vbNewLine

End Sub



