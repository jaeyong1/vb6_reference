VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmDialer 
   Caption         =   "Dialer"
   ClientHeight    =   5445
   ClientLeft      =   2880
   ClientTop       =   3360
   ClientWidth     =   6450
   Icon            =   "frmDialer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6450
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1260
      Left            =   2640
      TabIndex        =   1
      Top             =   2880
      Width           =   1380
   End
   Begin VB.TextBox PhoneNum2 
      Height          =   270
      Left            =   1980
      TabIndex        =   0
      Text            =   "016-535-6090"
      Top             =   510
      Width           =   3465
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2025
      Top             =   1845
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmDialer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PhoneNum As String

Private Sub Ring()
        
    bQuit = False
    '전화번호 있는가 확인
    If PhoneNum = "" Then
        MsgBox "전화번호를 입력하세요."
        Exit Sub
    ElseIf Left(PhoneNum, 4) = "0412" Then
        PhoneNum = Mid(PhoneNum, 6)
    End If
    '전화 걸기
    Dial PhoneNum
End Sub

Private Sub Dial(Number$)
    

    Dim dialstring$, FromModem$, dummy

    ' AT는 Hayse 호환 ATTENTION 명령어로 모뎀에 명령을 보낼 때 필요합니다.
    ' DT는 "Dial Tone"입니다. Dial 명령은 펄스와는 반대로 접촉음을 사용합니다(DP = Dial Pulse).
    ' Numbers$ 는 전화를 걸고 있는 전화 번호입니다.
    ' 세미콜론은 전화를 건 후 모뎀이 명령 모드로 반환할 것을 알려줍니다(중요).
    ' 캐리지 리턴인 vbCr는 모뎀에 명령을 보낼 경우 필요합니다.
    dialstring$ = "ATDT" + Number$ + ";" + vbCr

    ' 통신 포트 설정.
    ' 마우스는 COM1에 부착되어 있고 CommPort는 3에 설정되어 있는 것으로 생각됩니다.
    MSComm1.CommPort = 2
    MSComm1.Settings = "38400,N,8,1"
    
    ' 통신 포트를 엽니다.
    On Error Resume Next
    
    MSComm1.PortOpen = True
    
    If Err Then
       MsgBox "COM2 Port : not available. Change the CommPort property to another port."
       Exit Sub
    End If
    
    ' 입력 버퍼의 내용을 지웁니다.
    MSComm1.InBufferCount = 0
    
    ' 전화를 거십시오.
    MSComm1.Output = dialstring$
    
    ' 모뎀에서 빠져 나오려면 "확인" 메시지를 기다립니다.
    Do
       dummy = DoEvents()
       ' 버퍼에 데이터가 있는 경우 데이터를 읽습니다.
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' "확인"을 검사합니다.
          If InStr(FromModem$, "OK") Then
             ' 사용자가 전화기를 들도록 알립니다.
             Screen.MousePointer = vbDefault
             Response = MsgBox("'" & PhoneNum & "' 으로 전화를 걸고 있습니다." & vbCrLf & vbCrLf & "통화를 하려면 수화기를 들고 확인을 누르세요." & vbCrLf & "연결을 끊으려면 취소를 누르세요.", vbOKCancel + vbExclamation, PhoneNum)
             If Response = vbOK Then
                    Exit Do
                ElseIf MSComm1.PortOpen = False Then
                    Exit Sub
                Else
                    bQuit = True
                    MSComm1.PortOpen = False
                End If
          End If
       End If
        
       ' 사용자가 취소를 선택하였습니까?
       If bQuit Then
          bQuit = False
          Exit Do
       End If
    Loop
    
    ' 모뎀 연결이 끊어집니다.
    MSComm1.Output = "ATH" + vbCr
    
    ' 포트를 닫습니다.
    MSComm1.PortOpen = False
    
    End
End Sub

Private Sub Command1_Click()
    Screen.MousePointer = vbHourglass

    PhoneNum = PhoneNum2
    If Len(PhoneNum) > 0 And Len(PhoneNum) < 20 And InStr(PhoneNum, "") > 0 Then
    Ring
    Else
        MsgBox "전화번호 형식이 아닙니다.", vbExclamation, "오류"
        End
   End If

End Sub

'
Private Sub Form_Load()
End Sub

