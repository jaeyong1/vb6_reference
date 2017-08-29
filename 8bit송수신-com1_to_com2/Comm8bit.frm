VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Comm8bit 
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows 기본값
   Begin RichTextLib.RichTextBox text2 
      Height          =   5655
      Left            =   6720
      TabIndex        =   6
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   9975
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Comm8bit.frx":0000
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   4560
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   3720
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox text1 
      Height          =   3375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Comm8bit.frx":0290
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "종료"
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "보내기"
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   3600
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1080
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
      OutBufferSize   =   1024
      RThreshold      =   1
      RTSEnable       =   -1  'True
      BaudRate        =   921600
      SThreshold      =   2
      EOFEnable       =   -1  'True
      InputMode       =   1
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   1680
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
      OutBufferSize   =   1024
      RThreshold      =   1
      RTSEnable       =   -1  'True
      BaudRate        =   19200
      SThreshold      =   2
      EOFEnable       =   -1  'True
      InputMode       =   1
   End
   Begin RichTextLib.RichTextBox text1 
      Height          =   135
      Index           =   1
      Left            =   3600
      TabIndex        =   4
      Top             =   3360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   238
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Comm8bit.frx":0525
   End
End
Attribute VB_Name = "Comm8bit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private send As Byte
Private AA As Integer, BB As Long, CC As Integer, dd As Integer
Private Ee As Byte, Ff As Byte, Gg As Byte
Private mm As String, nn As String
Private Start As Byte
Private Rmode As Byte, Bmode As Byte, Read16 As Single, Read24 As Single

Dim ii As Integer, jj As Integer, Num As Byte, Value As Byte



Private Sub cmdSend_Click()
If cmdSend.Caption = "보내기" Then
    cmdSend.Caption = "정지"
    Timer1.Enabled = True
Else
    cmdSend.Caption = "보내기"
    Timer1.Enabled = False
End If


End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
ii = 1
    text1(0).Text = ""
    text1(0).SelText = "my name is kim kwang seop" & vbCrLf
    text1(0).SelText = "my birthday is 22,08,1974" & vbCrLf
    text1(0).SelText = "my E_mail address is " & vbCrLf
    text1(0).SelText = "kks3825@hitel.net" & vbCrLf
    
'   text1(1).Text = "": text1(0).Text = ""
   
  
   Timer1.Enabled = False
   
    '프로토콜
    MSComm1.CommPort = 1 '2
    MSComm1.Settings = "19200,N,8,1"
    MSComm1.PortOpen = True
    
    
'    MSComm2.CommPort = 2 '2
'    MSComm2.Settings = "19200,N,8,1"
'    MSComm2.PortOpen = False
End Sub


Private Sub MSComm1_OnComm()            '통신 모드 수신기 1
    Dim strTemp As String, strSts As String, tmpstr As String
    Dim CommData As String
    Dim InBufCnt As Integer, ii As Integer
    Dim InCell As Byte
    Dim TChar(512) As Byte
        
    Select Case MSComm1.CommEvent
   ' Handle each error by placing code below each case statement

        Case comEvReceive                       '이벤트 구문
            'strSts = " Received chars."
                While MSComm1.InBufferCount > 0
                    tmpstr = MSComm1.Input
                    InCell = AscB(tmpstr) 'And &HFF)
                    
                    text2.SelText = Hex(InCell) & "  "
                    
                    'Call RcvMsgPrs(InCell)
                Wend
   End Select
End Sub

Private Sub MSComm2_OnComm()            '통신 모드 수신기 1
    Dim strTemp As String, strSts As String, tmpstr As String
    Dim CommData As String
    Dim InBufCnt As Integer, ii As Integer
    Dim InCell As Byte
    Dim TChar(512) As Byte
        
    Select Case MSComm2.CommEvent
        Case comEvReceive                       '이벤트 구문
            'strSts = " Received chars."
                While MSComm2.InBufferCount > 0
                    tmpstr = MSComm2.Input
                    InCell = AscB(tmpstr) 'And &HFF)
                    
                    Call RcvMsgPrs(InCell)
                    
            Text4.SelText = InCell & " "
            text1.SelText = Hex(InCell) & " "

                Wend
   End Select
End Sub

Private Sub RcvMsgPrs(InCell As Byte)

text2.SelText = Chr(InCell) & "  "

End Sub


Private Sub Timer1_Timer()
    Dim AA As String, BB As Integer
    'For AA = 1 To Len(text1(0).Text)
    Timer1.Enabled = True
    
    AA = Mid(text1(0).Text, ii, 1)
    Text4.Text = ii
    Text3.Text = Len(text1(0).Text)
        Num = Asc(AA)
    ii = ii + 1
    Call Send0(Num)
        If ii > Len(text1(0).Text) Then
            ii = 1
            'Timer1.Enabled = False
      '     Text2.Text = CByte(Num) & "  " & CByte((Value))
        End If
    
End Sub


Function Send0(Num As Byte)
Dim TChar(1) As Byte
    
    TChar(1) = CByte(Num)
    MSComm1.Output = TChar
 '   text2.SelText = Chr(Num)
    
End Function

