VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   Caption         =   "한림 용해로 온도컨트롤 경보시스템"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmd_sample 
      Caption         =   "test샘플링"
      Height          =   375
      Left            =   4920
      TabIndex        =   23
      Top             =   2520
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6720
      Top             =   2400
   End
   Begin VB.CommandButton cmndSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmndOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmndClose 
      Caption         =   "Close"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtRx 
      Height          =   270
      Left            =   2520
      TabIndex        =   13
      Text            =   "00RSS0302123402567802ABCD"
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox txtTx 
      Height          =   270
      Left            =   2520
      TabIndex        =   12
      Text            =   "00RSS0106%MW001"
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox txtTxHead 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1920
      TabIndex        =   11
      Text            =   "ENQ"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtTxTail 
      Enabled         =   0   'False
      Height          =   270
      Left            =   5880
      TabIndex        =   10
      Text            =   "EOT"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtTxBcc 
      Height          =   270
      Left            =   6480
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtRxHead 
      Height          =   270
      Left            =   1920
      TabIndex        =   8
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtRxTail 
      Height          =   270
      Left            =   5880
      TabIndex        =   7
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtRxBcc 
      Height          =   270
      Left            =   6480
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.CheckBox chkBcc 
      Caption         =   "BCC"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox cmbPort 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   3120
      List            =   "Form1.frx":0010
      Style           =   2  '드롭다운 목록
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7560
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   4095
      ExtentX         =   7223
      ExtentY         =   6165
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command1 
      Caption         =   "인터넷으로 전송"
      Height          =   855
      Left            =   4320
      TabIndex        =   0
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Com port 통신에러 : "
      Height          =   180
      Left            =   1920
      TabIndex        =   22
      Top             =   1920
      Width           =   1740
   End
   Begin VB.Label lbltimeout 
      BorderStyle     =   1  '단일 고정
      Height          =   375
      Left            =   3840
      TabIndex        =   21
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Rx"
      Height          =   180
      Left            =   1560
      TabIndex        =   20
      Top             =   1440
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tx"
      Height          =   180
      Left            =   1560
      TabIndex        =   19
      Top             =   1080
      Width           =   225
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Head"
      Height          =   180
      Left            =   1920
      TabIndex        =   18
      Top             =   840
      Width           =   435
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Tail"
      Height          =   180
      Left            =   5880
      TabIndex        =   17
      Top             =   840
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "COM Port"
      Height          =   180
      Left            =   2040
      TabIndex        =   4
      Top             =   480
      Width           =   825
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label1"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   7440
      Width           =   6495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ACK00RSS03020000020000020000etX
Private IE As InternetExplorer


Private Sub cmd_sample_Click()
    sampling_data
End Sub

Private Sub cmndClose_Click()
    MSComm1.PortOpen = False
    
    cmbPort.Enabled = True
    cmndOpen.Enabled = True
    cmndSend.Enabled = False
    cmndClose.Enabled = False
End Sub

Private Sub cmndOpen_Click()

    OpenCommPort
    
    cmbPort.Enabled = False
    cmndOpen.Enabled = False
    cmndSend.Enabled = True
    cmndClose.Enabled = True
End Sub


' Send 버튼 Click Message Handler
' Data를 Send 합니다.
Private Sub cmndSend_Click()
   Dim Buffer As String
    Dim Head As String
    Dim Length As Long
    Dim Tx As String
    Dim Rx As String
    Dim lRefTime As Long
    Dim lCurTime As Long
    Dim Bcc As String
    Dim isrcv As Integer '수신완료됐는지? ack check
    
    
    isrcv = 0
     lbltimeout.Caption = ""
       
    txtRxHead.Text = ""
    txtRx.Text = ""
    txtRxTail.Text = ""
    txtTxBcc.Text = ""
    txtRxBcc.Text = ""
      
    ' 데이터를 Send 합니다.
    SendData

    ' Time Out을 계산하기 위해 데이터를 Send한 시간을 기록합니다.
    lRefTime = GetTickCount()


    ' ETX가 수신되거나 Time Out이 발생할 때까지 Loop를 돕니다.
    Do
        DoEvents
        Buffer$ = Buffer$ & MSComm1.Input
        
        ' ETX가 수신되었는지를 Check 합니다.
        Length = InStr(Buffer$, chr$(3))
        
        ' Time Out을 Check 합니다. (여기에서는 1000 msec로 설정하였습니다.)
        If ((GetTickCount() - lRefTime) > 1000) Then
            'MsgBox "Time Out Error !!!", vbOKOnly, "Error"
            lbltimeout.Caption = "Time Out Error !!!"
            Exit Sub
        End If
    Loop Until (Length)
      
    ' BCC가 설정된 경우에는 BCC의 수신을 확실히 하기위해 한번더 Input을 수행합니다.
    If chkBcc.Value = 1 Then
        Buffer$ = Buffer$ & MSComm1.Input
    End If
    
    Head = Left(Buffer$, 1)
   
    ' ACK가 수신된 경우
    If (Head = chr$(6)) Then
        txtRxHead.Text = "ACK"
        isrcv = 1
    ' NAK가 수신된 경우
    ElseIf (Head = chr$(&H15)) Then
        txtRxHead.Text = "NAK"
    ' ACK나 NAK가 수신되지 않은 경우
    Else
        'MsgBox "Unknown", vbOKOnly, "Rx Message"
         lbltimeout.Caption = "Unknown"
        Exit Sub
    End If
    
    txtRxTail.Text = "ETX"
    Rx = Mid(Buffer$, 2, Length - 2)
    txtRx.Text = Rx
    
    ' BCC가 선택된 경우에는 수신된 BCC를 화면에 출력합니다.
    If chkBcc.Value = 1 Then
        Bcc = Mid(Buffer$, Length + 1, 2)
        txtRxBcc.Text = Bcc
    End If
    
    If isrcv = 1 Then
        sampling_data
    End If
    
    
End Sub
Public Sub sampling_data()
    Dim RXstr As String '수신한 전체 문자열
    Dim chr As String   '추출할 글자
    Dim ar As String    '나온 1010같은 문자
    
    
    
    ' Mid(RXstr, 3, 1)  = R
       
    RXstr = txtRx
    
    If PlcSendData_iter = 0 Then  '1, 2, 3라인 경보 읽기 (대문자)
        ar = hextoarray(Mid(RXstr, 7, 1))
        Print ar
             
        
        
        
        
        
        
        
        
        
        
        
        
        
    ElseIf PlcSendData_iter = 1 Then '1 라인 ABC온도읽기
    
    ElseIf PlcSendData_iter = 2 Then '2 라인 ABC온도읽기
    
    ElseIf PlcSendData_iter = 3 Then '3 라인 ABC온도읽기
        
    End If
    
    
    
    


End Sub


Private Sub Command1_Click()
'인터넷으로 전송

On Error Resume Next
WebBrowser1.navigate "http://jy01.maru.net/p1/indata.html"

Label1.Caption = WebBrowser1.LocationURL

Do While WebBrowser1.Busy
     DoEvents
Loop




End Sub


Private Sub Form_Load()

    ' Default 값을 설정합니다.
    frmMain.cmbPort.ListIndex = 0             ' COM Port : COM1
    PlcSendData(0) = "00RSS0306%MW01006%MW02006%MW030" '1, 2, 3라인 경보 읽기 (대문자)
    PlcSendData(1) = "00RSS0306%MW11106%MW11206%MW113" '1 라인 ABC온도읽기
    PlcSendData(2) = "00RSS0306%MW12106%MW12206%MW123" '2 라인 ABC온도읽기
    PlcSendData(3) = "00RSS0306%MW13106%MW13206%MW133" '3 라인 ABC온도읽기
    








End Sub


Private Sub Timer1_Timer()

If cmndSend.Enabled = True Then

    If PlcSendData_iter = 0 Then
        txtTx.Text = PlcSendData(0)
        cmndSend_Click
        
        PlcSendData_iter = PlcSendData_iter + 1
    ElseIf PlcSendData_iter = 1 Then
        txtTx.Text = PlcSendData(1)
        cmndSend_Click
        PlcSendData_iter = PlcSendData_iter + 1
    ElseIf PlcSendData_iter = 2 Then
        txtTx.Text = PlcSendData(2)
        cmndSend_Click
        PlcSendData_iter = PlcSendData_iter + 1
    ElseIf PlcSendData_iter = 3 Then
        txtTx.Text = PlcSendData(3)
        cmndSend_Click
        PlcSendData_iter = 0
    End If
        
        
End If 'cmndSend

End Sub

'인터넷 페이지가 다 불려지면 할일(텍스트박스에 데이터 저장 + 저장버튼클릭)
Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next

WebBrowser1.document.Forms(0).t01.Value = t(1)
WebBrowser1.document.Forms(0).t02.Value = t(2)
WebBrowser1.document.Forms(0).t03.Value = t(3)
WebBrowser1.document.Forms(0).t04.Value = t(4)
WebBrowser1.document.Forms(0).t05.Value = t(5)
WebBrowser1.document.Forms(0).t06.Value = t(6)
WebBrowser1.document.Forms(0).t07.Value = t(7)
WebBrowser1.document.Forms(0).t08.Value = t(8)
WebBrowser1.document.Forms(0).t09.Value = t(9)
WebBrowser1.document.Forms(0).t10.Value = t(10)
WebBrowser1.document.Forms(0).t11.Value = t(11)
WebBrowser1.document.Forms(0).t12.Value = t(12)
WebBrowser1.document.Forms(0).t13.Value = t(13)
WebBrowser1.document.Forms(0).t14.Value = t(14)
WebBrowser1.document.Forms(0).t15.Value = t(15)
WebBrowser1.document.Forms(0).t16.Value = t(16)
WebBrowser1.document.Forms(0).t17.Value = t(17)
WebBrowser1.document.Forms(0).t18.Value = t(18)
WebBrowser1.document.Forms(0).t19.Value = t(19)
WebBrowser1.document.Forms(0).t20.Value = t(20)
WebBrowser1.document.Forms(0).t21.Value = t(21)
WebBrowser1.document.Forms(0).t22.Value = t(22)
WebBrowser1.document.Forms(0).t23.Value = t(23)
WebBrowser1.document.Forms(0).t24.Value = t(24)
WebBrowser1.document.Forms(0).t25.Value = t(25)
WebBrowser1.document.Forms(0).t26.Value = t(26)
WebBrowser1.document.Forms(0).t27.Value = t(27)
WebBrowser1.document.Forms(0).t28.Value = t(28)
WebBrowser1.document.Forms(0).t29.Value = t(29)
WebBrowser1.document.Forms(0).t30.Value = t(30)
WebBrowser1.document.Forms(0).t31.Value = t(31)
WebBrowser1.document.Forms(0).t32.Value = t(32)
WebBrowser1.document.Forms(0).t33.Value = t(33)
WebBrowser1.document.Forms(0).t34.Value = t(34)
WebBrowser1.document.Forms(0).t35.Value = t(35)
WebBrowser1.document.Forms(0).t36.Value = t(36)
WebBrowser1.document.Forms(0).t37.Value = t(37)
WebBrowser1.document.Forms(0).t38.Value = t(38)
WebBrowser1.document.Forms(0).t39.Value = t(39)
WebBrowser1.document.Forms(0).t40.Value = t(40)
WebBrowser1.document.Forms(0).t41.Value = t(41)
WebBrowser1.document.Forms(0).t42.Value = t(42)




WebBrowser1.document.Forms(0).submit
End Sub

