VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FRM쓰기 
   Caption         =   "변수 쓰기"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9795
   BeginProperty Font 
      Name            =   "돋움체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Write Data"
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   3945
      TabIndex        =   13
      Top             =   1515
      Width           =   5760
      Begin VB.TextBox Text 
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   9
         Left            =   5085
         TabIndex        =   23
         Text            =   "0000"
         Top             =   225
         Width           =   560
      End
      Begin VB.TextBox Text 
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   8
         Left            =   4530
         TabIndex        =   22
         Text            =   "9999"
         Top             =   225
         Width           =   560
      End
      Begin VB.TextBox Text 
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   7
         Left            =   3975
         TabIndex        =   21
         Text            =   "8888"
         Top             =   225
         Width           =   560
      End
      Begin VB.TextBox Text 
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   6
         Left            =   3420
         TabIndex        =   20
         Text            =   "7777"
         Top             =   225
         Width           =   560
      End
      Begin VB.TextBox Text 
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   5
         Left            =   2865
         TabIndex        =   19
         Text            =   "6666"
         Top             =   225
         Width           =   560
      End
      Begin VB.TextBox Text 
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   4
         Left            =   2310
         TabIndex        =   18
         Text            =   "5555"
         Top             =   225
         Width           =   560
      End
      Begin VB.TextBox Text 
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   3
         Left            =   1770
         TabIndex        =   17
         Text            =   "4444"
         Top             =   225
         Width           =   560
      End
      Begin VB.TextBox Text 
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   2
         Left            =   1215
         TabIndex        =   16
         Text            =   "3333"
         Top             =   225
         Width           =   560
      End
      Begin VB.TextBox Text 
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   660
         TabIndex        =   15
         Text            =   "2222"
         Top             =   225
         Width           =   560
      End
      Begin VB.TextBox Text 
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   105
         TabIndex        =   14
         Text            =   "1111"
         Top             =   225
         Width           =   560
      End
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
      Height          =   480
      Left            =   3960
      TabIndex        =   12
      ToolTipText     =   "상단창의 내용을 지움"
      Top             =   3390
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      Caption         =   "수신 Form"
      ForeColor       =   &H00FF0000&
      Height          =   1185
      Left            =   3960
      TabIndex        =   9
      Top             =   2190
      Width           =   5745
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FF00&
         Height          =   855
         Left            =   105
         TabIndex        =   11
         Top             =   225
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "송신 Form"
      ForeColor       =   &H000000FF&
      Height          =   1530
      Left            =   3945
      TabIndex        =   8
      Top             =   -15
      Width           =   5760
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FFFF&
         Height          =   1185
         Left            =   105
         TabIndex        =   10
         Top             =   210
         Width           =   5565
      End
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
      Height          =   480
      Left            =   6840
      TabIndex        =   7
      Top             =   3390
      Width           =   2865
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin MSCommLib.MSComm MSComm1 
         Left            =   30
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   327680
         DTREnable       =   -1  'True
      End
      Begin VB.CommandButton Command1 
         Caption         =   "직접 변수 개별 쓰기"
         Height          =   495
         Left            =   135
         TabIndex        =   6
         ToolTipText     =   "%QW0.3.0 1WORD를 쓰기"
         Top             =   240
         Width           =   3480
      End
      Begin VB.CommandButton Command2 
         Caption         =   "직접 변수 연속 쓰기"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "%QW0.3.0에서 10WORD를 쓰기"
         Top             =   840
         Width           =   3495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "액세스 WORD형 변수 쓰기"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "액세스 변수 OUT_1,2. MOTOR1,2 쓰기"
         Top             =   1440
         Width           =   3495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "액세스 INT형 변수 쓰기"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "INT_CV,INT_CV1 쓰기"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "액세스 INT형 Array 변수 쓰기"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "INT_ARRAY 10개 쓰기"
         Top             =   2640
         Width           =   3495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "액세스 UDINT형 Array 변수 쓰기"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "DINT_ARRAY 5개 쓰기"
         Top             =   3240
         Width           =   3495
      End
   End
End
Attribute VB_Name = "FRM쓰기"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdQuit_Click()
    Unload Me
End Sub

Private Sub CmdScreenClear_Click()
    Label1 = ""
    Label2 = ""
End Sub

Private Sub Command1_Click()
'직접 변수의 메모리 어드레스를 지정하여 데이터를 쓰는 경우(명령어 → WSS)
    '통신 프레임 만들기
    Data = Text(0)
    Q = Chr(5) & "00WSS" & "01" & "08%QW0.3.0" & Data & Chr(4)
'ERROR 추적기 설치
On Error GoTo ErrMsg                '통신포트 에러의 경우 ErrMsg:로 점프
    MSComm1.CommPort = 1            'COM1을 사용.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '입력버퍼 크기를 1Byte로 함.
    MSComm1.PortOpen = True         '통신포트 열기.
    MSComm1.Output = Q              '프레임 전송.
    Label1 = Q
    Do
        DoEvents                    'Loop수행중 외부Event 검출.
        InString = MSComm1.Input    '입력버퍼로 부터 한문자를 읽어냄.
        Rcv = Rcv & InString
        Label2 = Rcv
        If InString = "" Then       '입력버퍼에 수신된 값을 확인하여 내용이 없으면
            Label2 = "PLC와 통신 확인중 잠시 기다려주세요."
            Rcv_No = Rcv_No + 1     '수신되지 않는 횟수를 COUNT.
        Else
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '연속적으로 DATA가 수신되지 않은 횟수가 1000보다 크면.
            Label2 = Err.Number
            Dummy = MsgBox("Time Out Error", 0, "에러") 'Time Out Error 메세지 표시.
            MSComm1.PortOpen = False 'COM1 통신 PORT 닫기.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX가 수신될 때까지 Do ... Loop간을 반복.
    
'    TextRcvData = TextRcvData & Rcv '수신된 문자열을 Text Box에 출력.
    Rcv = ""                        '문자열 변수 초기화.
    MSComm1.PortOpen = False        'COM1 통신 PORT 닫기.
ErrMsg:
    Label2 = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "에러")
End Sub

Private Sub Command2_Click()
'직접 변수의 메모리 러드레스를 지정하여 데이터를 연속으로 읽는 경우(명령어 → RSB)
    '통신 프레임 만들기
    For i = 0 To 9
        Data = Data & Text(i)
    Next i
    Q = Chr(5) & "00WSB" & "08%QW0.3.0" & "0A" & Data & Chr(4)
'ERROR 추적기 설치
On Error GoTo ErrMsg                '통신포트 에러의 경우 ErrMsg:로 점프
    MSComm1.CommPort = 1            'COM1을 사용.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '입력버퍼 크기를 1Byte로 함.
    MSComm1.PortOpen = True         '통신포트 열기.
    MSComm1.Output = Q              '프레임 전송.
    Label1 = Q
    Do
        DoEvents                    'Loop수행중 외부Event 검출.
        InString = MSComm1.Input    '입력버퍼로 부터 한문자를 읽어냄.
        Rcv = Rcv & InString
        Label2 = Rcv
        If InString = "" Then       '입력버퍼에 수신된 값을 확인하여 내용이 없으면
            Label2 = "PLC와 통신 확인중 잠시 기다려주세요."
            Rcv_No = Rcv_No + 1     '수신되지 않는 횟수를 COUNT.
        Else
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '연속적으로 DATA가 수신되지 않은 횟수가 1000보다 크면.
            Label2 = Err.Number
            Dummy = MsgBox("Time Out Error", 0, "에러") 'Time Out Error 메세지 표시.
            MSComm1.PortOpen = False 'COM1 통신 PORT 닫기.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX가 수신될 때까지 Do ... Loop간을 반복.
    
'    TextRcvData = TextRcvData & Rcv '수신된 문자열을 Text Box에 출력.
    Rcv = ""                        '문자열 변수 초기화.
    MSComm1.PortOpen = False        'COM1 통신 PORT 닫기.
ErrMsg:
    Label2 = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "에러")

End Sub

Private Sub Command3_Click()
'액세스 변수에 등록된 글로벌 변수가 WORD형인 경우(명령어 → R02)
'액세스 변수 이름 : OUT_1, OUT_2, MOTOR1, MOTOR2
    '통신 프레임 만들기
    OUT_1_DATA = Text(0)
    OUT_2_DATA = Text(1)
    MOTOR1_DATA = Text(2)
    MOTOR2_DATA = Text(3)
    Q = Chr(5) & "00W02" & "04" & "05OUT_1" & OUT_1_DATA & "05OUT_2" & OUT_2_DATA & "06MOTOR1" & MOTOR1_DATA & "06MOTOR2" & MOTOR2_DATA & Chr(4)
'ERROR 추적기 설치
On Error GoTo ErrMsg                '통신포트 에러의 경우 ErrMsg:로 점프
    MSComm1.CommPort = 1            'COM1을 사용.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '입력버퍼 크기를 1Byte로 함.
    MSComm1.PortOpen = True         '통신포트 열기.
    MSComm1.Output = Q              '프레임 전송.
    Label1 = Q
    Do
        DoEvents                    'Loop수행중 외부Event 검출.
        InString = MSComm1.Input    '입력버퍼로 부터 한문자를 읽어냄.
        Rcv = Rcv & InString
        Label2 = Rcv
        If InString = "" Then       '입력버퍼에 수신된 값을 확인하여 내용이 없으면
            Label2 = "PLC와 통신 확인중 잠시 기다려주세요."
            Rcv_No = Rcv_No + 1     '수신되지 않는 횟수를 COUNT.
        Else
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '연속적으로 DATA가 수신되지 않은 횟수가 1000보다 크면.
            Label2 = Err.Number
            Dummy = MsgBox("Time Out Error", 0, "에러") 'Time Out Error 메세지 표시.
            MSComm1.PortOpen = False 'COM1 통신 PORT 닫기.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX가 수신될 때까지 Do ... Loop간을 반복.
    
'    TextRcvData = TextRcvData & Rcv '수신된 문자열을 Text Box에 출력.
    Rcv = ""                        '문자열 변수 초기화.
    MSComm1.PortOpen = False        'COM1 통신 PORT 닫기.
ErrMsg:
    Label2 = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "에러")
End Sub

Private Sub Command4_Click()
'액세스 변수에 등록된 글로벌 변수가 INT형인 경우(명령어 → W02)
'액세스 변수 이름 : INT_CV, INT_CV1
    '통신 프레임 만들기
    INT_CV_DATA = Text(0)
    INT_CV1_DATA = Text(1)
    Q = Chr(5) & "00W06" & "02" & "06INT_CV" & INT_CV_DATA & "07INT_CV1" & INT_CV1_DATA & Chr(4)
'ERROR 추적기 설치
On Error GoTo ErrMsg                '통신포트 에러의 경우 ErrMsg:로 점프
    MSComm1.CommPort = 1            'COM1을 사용.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '입력버퍼 크기를 1Byte로 함.
    MSComm1.PortOpen = True         '통신포트 열기.
    MSComm1.Output = Q              '프레임 전송.
    Label1 = Q
    Do
        DoEvents                    'Loop수행중 외부Event 검출.
        InString = MSComm1.Input    '입력버퍼로 부터 한문자를 읽어냄.
        Rcv = Rcv & InString
        Label2 = Rcv
        If InString = "" Then       '입력버퍼에 수신된 값을 확인하여 내용이 없으면
            Label2 = "PLC와 통신 확인중 잠시 기다려주세요."
            Rcv_No = Rcv_No + 1     '수신되지 않는 횟수를 COUNT.
        Else
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '연속적으로 DATA가 수신되지 않은 횟수가 1000보다 크면.
            Label2 = Err.Number
            Dummy = MsgBox("Time Out Error", 0, "에러") 'Time Out Error 메세지 표시.
            MSComm1.PortOpen = False 'COM1 통신 PORT 닫기.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX가 수신될 때까지 Do ... Loop간을 반복.
    
'    TextRcvData = TextRcvData & Rcv '수신된 문자열을 Text Box에 출력.
    Rcv = ""                        '문자열 변수 초기화.
    MSComm1.PortOpen = False        'COM1 통신 PORT 닫기.
ErrMsg:
    Label2 = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "에러")
End Sub

Private Sub Command5_Click()
'액세스 변수에 등록된 글로벌 변수가 INT형 Array인 경우(명령어 → R1B)
'액세스 변수 이름 : INT_ARRAY, ARRAY 원소 갯수 10개.
    '통신 프레임 만들기
    For i = 0 To 9
        Data = Data & Text(i)
    Next i
    Q = Chr(5) & "00W1B" & "09INT_ARRAY" & "0A" & Data & Chr(4)
'ERROR 추적기 설치
On Error GoTo ErrMsg                '통신포트 에러의 경우 ErrMsg:로 점프
    MSComm1.CommPort = 1            'COM1을 사용.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '입력버퍼 크기를 1Byte로 함.
    MSComm1.PortOpen = True         '통신포트 열기.
    MSComm1.Output = Q              '프레임 전송.
    Label1 = Q
    Do
        DoEvents                    'Loop수행중 외부Event 검출.
        InString = MSComm1.Input    '입력버퍼로 부터 한문자를 읽어냄.
        Rcv = Rcv & InString
        Label2 = Rcv
        If InString = "" Then       '입력버퍼에 수신된 값을 확인하여 내용이 없으면
            Label2 = "PLC와 통신 확인중 잠시 기다려주세요."
            Rcv_No = Rcv_No + 1     '수신되지 않는 횟수를 COUNT.
        Else
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '연속적으로 DATA가 수신되지 않은 횟수가 1000보다 크면.
            Label2 = Err.Number
            Dummy = MsgBox("Time Out Error", 0, "에러") 'Time Out Error 메세지 표시.
            MSComm1.PortOpen = False 'COM1 통신 PORT 닫기.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX가 수신될 때까지 Do ... Loop간을 반복.
    
'    TextRcvData = TextRcvData & Rcv '수신된 문자열을 Text Box에 출력.
    Rcv = ""                        '문자열 변수 초기화.
    MSComm1.PortOpen = False        'COM1 통신 PORT 닫기.
ErrMsg:
    Label2 = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "에러")
End Sub

Private Sub Command6_Click()
'액세스 변수에 등록된 글로벌 변수가 UDINT형 Array인 경우(명령어 → R20)
'액세스 변수 이름 : DINT_ARRAY, ARRAY 원소 갯수 5개.
    '통신 프레임 만들기
    For i = 0 To 9
        Data = Data & Text(i)
    Next i
    Q = Chr(5) & "00R20" & "0ADINT_ARRAY" & "05" & Data & Chr(4)
'ERROR 추적기 설치
On Error GoTo ErrMsg                '통신포트 에러의 경우 ErrMsg:로 점프
    MSComm1.CommPort = 1            'COM1을 사용.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '입력버퍼 크기를 1Byte로 함.
    MSComm1.PortOpen = True         '통신포트 열기.
    MSComm1.Output = Q              '프레임 전송.
    Label1 = Q
    Do
        DoEvents                    'Loop수행중 외부Event 검출.
        InString = MSComm1.Input    '입력버퍼로 부터 한문자를 읽어냄.
        Rcv = Rcv & InString
        Label2 = Rcv
        If InString = "" Then       '입력버퍼에 수신된 값을 확인하여 내용이 없으면
            Label2 = "PLC와 통신 확인중 잠시 기다려주세요."
            Rcv_No = Rcv_No + 1     '수신되지 않는 횟수를 COUNT.
        Else
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '연속적으로 DATA가 수신되지 않은 횟수가 1000보다 크면.
            Label2 = Err.Number
            Dummy = MsgBox("Time Out Error", 0, "에러") 'Time Out Error 메세지 표시.
            MSComm1.PortOpen = False 'COM1 통신 PORT 닫기.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX가 수신될 때까지 Do ... Loop간을 반복.
    
'    TextRcvData = TextRcvData & Rcv '수신된 문자열을 Text Box에 출력.
    Rcv = ""                        '문자열 변수 초기화.
    MSComm1.PortOpen = False        'COM1 통신 PORT 닫기.
ErrMsg:
    Label2 = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "에러")
End Sub
