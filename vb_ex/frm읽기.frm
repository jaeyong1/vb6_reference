VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FRM읽기 
   Caption         =   "변수 읽기"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm MSComm1 
      Left            =   840
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   4200
      TabIndex        =   3
      Top             =   15
      Width           =   3735
      Begin VB.CommandButton Command6 
         Caption         =   "액세스 UDINT형 Array 변수 읽기"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "DINT_ARRAY 5개 읽기"
         Top             =   3240
         Width           =   3495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "액세스 INT형 Array 변수 읽기"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "INT_ARRAY 10개 읽기"
         Top             =   2640
         Width           =   3495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "액세스 INT형 변수 읽기"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "INT_CV,INT_CV1 읽기"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "액세스 WORD형 변수 읽기"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "액세스 변수 OUT_1,2. MOTOR1,2 읽기"
         Top             =   1440
         Width           =   3495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "직접 변수 연속 읽기"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "%QW0.3.0에서 10WORD를 읽음"
         Top             =   840
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "직접 변수 개별 읽기"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "%QW0.3.0 1WORD를 읽음"
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.TextBox TextRcvData 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "종료"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   1
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton CmdScreenClear 
      Caption         =   "화면삭제"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   0
      ToolTipText     =   "우측창의 내용을 지움"
      Top             =   3960
      Width           =   1815
   End
End
Attribute VB_Name = "FRM읽기"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InString As String
Dim Q As String
Dim Rcv As String

Private Sub CmdQuit_Click()
    Unload Me                            '프로그램의 수행 종료.
End Sub

Private Sub CmdScreenClear_Click()
    TextRcvData = ""                'TextBox의 문자열 삭제.
End Sub

Private Sub Command1_Click()
'직접 변수의 메모리 러드레스를 지정하여 데이터를 읽는 경우(명령어 → RSS)
    '통신 프레임 만들기
    Q = Chr(5) & "00RSS" & "01" & "08%QW0.3.0" & Chr(4)
'ERROR 추적기 설치
On Error GoTo ErrMsg                '통신포트 에러의 경우 ErrMsg:로 점프
    MSComm1.CommPort = 1            'COM1을 사용.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '입력버퍼 크기를 1Byte로 함.
    MSComm1.PortOpen = True         '통신포트 열기.
    MSComm1.Output = Q              '프레임 전송.
    
    Do
        DoEvents                    'Loop수행중 외부Event 검출.
        InString = MSComm1.Input    '입력버퍼로 부터 한문자를 읽어냄.
        Rcv = Rcv & InString
        If InString = "" Then       '입력버퍼에 수신된 값을 확인하여 내용이 없으면
            TextRcvData = "PLC와 통신중 잠시 기다려주세요."
            Rcv_No = Rcv_No + 1     '수신되지 않는 횟수를 COUNT.
        Else
            TextRcvData = ""
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '연속적으로 DATA가 수신되지 않은 횟수가 1000보다 크면.
            TextRcvData = ""
            Dummy = MsgBox("Time Out Error", 0, "에러") 'Time Out Error 메세지 표시.
            MSComm1.PortOpen = False 'COM1 통신 PORT 닫기.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX가 수신될 때까지 Do ... Loop간을 반복.
    
    TextRcvData = TextRcvData & Rcv '수신된 문자열을 Text Box에 출력.
    Rcv = ""                        '문자열 변수 초기화.
    MSComm1.PortOpen = False        'COM1 통신 PORT 닫기.
ErrMsg:
    TextRcvData = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "에러")
End Sub

Private Sub Command2_Click()
'직접 변수의 메모리 러드레스를 지정하여 데이터를 연속으로 읽는 경우(명령어 → RSB)
    '통신 프레임 만들기
    Q = Chr(5) & "00RSB" & "08%QW0.3.0" & "0A" & Chr(4)
'ERROR 추적기 설치
On Error GoTo ErrMsg                '통신포트 에러의 경우 ErrMsg:로 점프
    MSComm1.CommPort = 1            'COM1을 사용.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '입력버퍼 크기를 1Byte로 함.
    MSComm1.PortOpen = True         '통신포트 열기.
    MSComm1.Output = Q              '프레임 전송.
    
    Do
        DoEvents                    'Loop수행중 외부Event 검출.
        InString = MSComm1.Input    '입력버퍼로 부터 한문자를 읽어냄.
        Rcv = Rcv & InString
        If InString = "" Then       '입력버퍼에 수신된 값을 확인하여 내용이 없으면
            TextRcvData = "PLC와 통신중 잠시 기다려주세요."
            Rcv_No = Rcv_No + 1     '수신되지 않는 횟수를 COUNT.
        Else
            TextRcvData = ""
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '연속적으로 DATA가 수신되지 않은 횟수가 1000보다 크면.
            TextRcvData = ""
            Dummy = MsgBox("Time Out Error", 0, "에러") 'Time Out Error 메세지 표시.
            MSComm1.PortOpen = False 'COM1 통신 PORT 닫기.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX가 수신될 때까지 Do ... Loop간을 반복.
    
    TextRcvData = TextRcvData & Rcv '수신된 문자열을 Text Box에 출력.
    Rcv = ""                        '문자열 변수 초기화.
    MSComm1.PortOpen = False        'COM1 통신 PORT 닫기.
ErrMsg:
    TextRcvData = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "에러")
End Sub

Private Sub Command3_Click()
'액세스 변수에 등록된 글로벌 변수가 WORD형인 경우(명령어 → R02)
'액세스 변수 이름 : OUT_1, OUT_2, MOTOR1, MOTOR2
    '통신 프레임 만들기
    Q = Chr(5) & "00R02" & "04" & "05OUT_1" & "05OUT_2" & "06MOTOR1" & "06MOTOR2" & Chr(4)
'ERROR 추적기 설치
On Error GoTo ErrMsg                '통신포트 에러의 경우 ErrMsg:로 점프
    MSComm1.CommPort = 1            'COM1을 사용.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '입력버퍼 크기를 1Byte로 함.
    MSComm1.PortOpen = True         '통신포트 열기.
    MSComm1.Output = Q              '프레임 전송.
    
    Do
        DoEvents                    'Loop수행중 외부Event 검출.
        InString = MSComm1.Input    '입력버퍼로 부터 한문자를 읽어냄.
        Rcv = Rcv & InString
        If InString = "" Then       '입력버퍼에 수신된 값을 확인하여 내용이 없으면
            TextRcvData = "PLC와 통신중 잠시 기다려주세요."
            Rcv_No = Rcv_No + 1     '수신되지 않는 횟수를 COUNT.
        Else
            TextRcvData = ""
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '연속적으로 DATA가 수신되지 않은 횟수가 1000보다 크면.
            TextRcvData = ""
            Dummy = MsgBox("Time Out Error", 0, "에러") 'Time Out Error 메세지 표시.
            MSComm1.PortOpen = False 'COM1 통신 PORT 닫기.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX가 수신될 때까지 Do ... Loop간을 반복.
    
    TextRcvData = TextRcvData & Rcv '수신된 문자열을 Text Box에 출력.
    Rcv = ""                        '문자열 변수 초기화.
    MSComm1.PortOpen = False        'COM1 통신 PORT 닫기.
ErrMsg:
    TextRcvData = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "에러")
End Sub

Private Sub Command4_Click()
'액세스 변수에 등록된 글로벌 변수가 INT형인 경우(명령어 → R06)
'액세스 변수 이름 : INT_CV,INT_CV1
    '통신 프레임 만들기
    Q = Chr(5) & "00R06" & "02" & "06INT_CV" & "07INT_CV1" & Chr(4)
'ERROR 추적기 설치
On Error GoTo ErrMsg                '통신포트 에러의 경우 ErrMsg:로 점프
    MSComm1.CommPort = 1            'COM1을 사용.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '입력버퍼 크기를 1Byte로 함.
    MSComm1.PortOpen = True         '통신포트 열기.
    MSComm1.Output = Q              '프레임 전송.
    
    Do
        DoEvents                    'Loop수행중 외부Event 검출.
        InString = MSComm1.Input    '입력버퍼로 부터 한문자를 읽어냄.
        Rcv = Rcv & InString
        If InString = "" Then       '입력버퍼에 수신된 값을 확인하여 내용이 없으면
            TextRcvData = "PLC와 통신중 잠시 기다려주세요."
            Rcv_No = Rcv_No + 1     '수신되지 않는 횟수를 COUNT.
        Else
            TextRcvData = ""
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '연속적으로 DATA가 수신되지 않은 횟수가 1000보다 크면.
            TextRcvData = ""
            Dummy = MsgBox("Time Out Error", 0, "에러") 'Time Out Error 메세지 표시.
            MSComm1.PortOpen = False 'COM1 통신 PORT 닫기.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX가 수신될 때까지 Do ... Loop간을 반복.
    
    TextRcvData = TextRcvData & Rcv '수신된 문자열을 Text Box에 출력.
    Rcv = ""                        '문자열 변수 초기화.
    MSComm1.PortOpen = False        'COM1 통신 PORT 닫기.
ErrMsg:
    TextRcvData = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "에러")
End Sub

Private Sub Command5_Click()
'액세스 변수에 등록된 글로벌 변수가 INT형 Array인 경우(명령어 → R1B)
'액세스 변수 이름 : INT_ARRAY, ARRAY 원소 갯수 10개.
    '통신 프레임 만들기
    Q = Chr(5) & "00R1B" & "09INT_ARRAY" & "0A" & Chr(4)
'ERROR 추적기 설치
On Error GoTo ErrMsg                '통신포트 에러의 경우 ErrMsg:로 점프
    MSComm1.CommPort = 1            'COM1을 사용.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '입력버퍼 크기를 1Byte로 함.
    MSComm1.PortOpen = True         '통신포트 열기.
    MSComm1.Output = Q              '프레임 전송.
    
    Do
        DoEvents                    'Loop수행중 외부Event 검출.
        InString = MSComm1.Input    '입력버퍼로 부터 한문자를 읽어냄.
        Rcv = Rcv & InString
        If InString = "" Then       '입력버퍼에 수신된 값을 확인하여 내용이 없으면
            TextRcvData = "PLC와 통신중 잠시 기다려주세요."
            Rcv_No = Rcv_No + 1     '수신되지 않는 횟수를 COUNT.
        Else
            TextRcvData = ""
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '연속적으로 DATA가 수신되지 않은 횟수가 1000보다 크면.
            TextRcvData = ""
            Dummy = MsgBox("Time Out Error", 0, "에러") 'Time Out Error 메세지 표시.
            MSComm1.PortOpen = False 'COM1 통신 PORT 닫기.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX가 수신될 때까지 Do ... Loop간을 반복.
    
    TextRcvData = TextRcvData & Rcv '수신된 문자열을 Text Box에 출력.
    Rcv = ""                        '문자열 변수 초기화.
    MSComm1.PortOpen = False        'COM1 통신 PORT 닫기.
ErrMsg:
    TextRcvData = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "에러")
End Sub

Private Sub Command6_Click()
'액세스 변수에 등록된 글로벌 변수가 UDINT형 Array인 경우(명령어 → R20)
'액세스 변수 이름 : DINT_ARRAY, ARRAY 원소 갯수 5개.
    '통신 프레임 만들기
    Q = Chr(5) & "00R20" & "0ADINT_ARRAY" & "05" & Chr(4)
'ERROR 추적기 설치
On Error GoTo ErrMsg                '통신포트 에러의 경우 ErrMsg:로 점프
    MSComm1.CommPort = 1            'COM1을 사용.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '입력버퍼 크기를 1Byte로 함.
    MSComm1.PortOpen = True         '통신포트 열기.
    MSComm1.Output = Q              '프레임 전송.
    
    Do
        DoEvents                    'Loop수행중 외부Event 검출.
        InString = MSComm1.Input    '입력버퍼로 부터 한문자를 읽어냄.
        Rcv = Rcv & InString
        If InString = "" Then       '입력버퍼에 수신된 값을 확인하여 내용이 없으면
            TextRcvData = "PLC와 통신중 잠시 기다려주세요."
            Rcv_No = Rcv_No + 1     '수신되지 않는 횟수를 COUNT.
        Else
            TextRcvData = ""
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '연속적으로 DATA가 수신되지 않은 횟수가 1000보다 크면.
            TextRcvData = ""
            Dummy = MsgBox("Time Out Error", 0, "에러") 'Time Out Error 메세지 표시.
            MSComm1.PortOpen = False 'COM1 통신 PORT 닫기.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX가 수신될 때까지 Do ... Loop간을 반복.
    
    TextRcvData = TextRcvData & Rcv '수신된 문자열을 Text Box에 출력.
    Rcv = ""                        '문자열 변수 초기화.
    MSComm1.PortOpen = False        'COM1 통신 PORT 닫기.
ErrMsg:
    TextRcvData = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "에러")
End Sub
