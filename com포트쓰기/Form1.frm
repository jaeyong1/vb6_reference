VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "보내기"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1320
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '단일 고정
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'일단 com포트를 쓸라면 왼쪽에 노란전화기를 추가해야함..
'추가하는 방법
'위에 메뉴에서
'프로젝트 -> 구성요소에서
'Microsoft Comm Control 6.0을 체크한후 확인하면
'노란전화기 아이콘이 생긴다.
'생긴 전화기를 클릭, 폼위에서 마우스로 드래그-> 폼위에 전화기가 놓인다.


Private Sub Command1_Click()

    MSComm1.CommPort = 1            'COM1을 사용.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '입력버퍼 크기를 1Byte로 함.
    MSComm1.PortOpen = True         '통신포트 열기.
    MSComm1.Output = "보내고싶은문자열" '실제로 이게 com1로 전송된다.
    MSComm1.PortOpen = False        '통신포트닫기 (계속 쓸꺼라면 안닫아야겠지)
 


End Sub



'글자 받는방법.. 쉽게말하면 com1에서 뭔가일이 있으면 인터럽터가 걸리듯이
'이 이벤트가 발생하는것이다.
'이벤트 실행됐을때 그 내용을 select로 확인해서 수신이면 글자를 받는다.

Private Sub MSComm1_OnComm()

Dim rcvtem
  
Select Case MSComm1.CommEvent
      Case comEvReceive '<- 수신이벤트 일때.. 이거말고도 원래는 엄청많다. 다른이벤트때 받을려고 시도하는걸 막으려는 것.

        If MSComm1.InBufferCount Then ' 비교구문이 없다? 아니다. 버퍼가 0이면 false 0이아닌숫자면 true인것!, 어떻게보면 위에꺼랑 같은뜻..
           rcvtemp = MSComm1.Input      '버퍼에 있는거 받음
           Label1 = rcvtemp             '화면에 출력
        End If

End Select
End Sub


'테스트는 안해봤지만 작동하는 소스일듯.. 복사-붙여넣기 신공이라..
