VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "직접변수 개별읽기"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   14.25
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text3 
      Height          =   4695
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Form1.frx":0000
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   4695
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form1.frx":0006
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   4695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form1.frx":000C
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "문제발생 연습용"
      Height          =   975
      Left            =   7440
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "중지"
      Height          =   615
      Left            =   7560
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "작동"
      Height          =   615
      Left            =   7560
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   1080
      Top             =   4680
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   360
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   19200
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "종료"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      TabIndex        =   0
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "master-k200H"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InString As String
Dim jup As Integer
Dim Q As String
Dim rcv As String
Dim j As Integer '시간간격
Dim data As String
Dim p(30) As arrar
Dim p1(30) As arrar
Dim p2(30) As arrar
Dim R As String
Dim k As Integer
Private Sub Command1_Click()
Timer1.Enabled = True
End Sub
Private Sub Command2_Click()
Timer1.Enabled = False
End Sub
Private Sub Command3_Click()
End
End Sub

Private Sub Timer1_Timer()
jup = 0 '연결안됬을때 따운방지용..
Cls
data = ""
k = 0
Dim etx As String: etx = Chr$(3)
Dim eot As String: eot = Chr$(4)
Dim enq As String: enq = Chr$(5)
Dim ack As String: ack = Chr$(6)
Dim nak As String: nak = Chr$(21)
Dim stx As String: stx = Chr$(2)
 
Dim address As String
Dim number As String

address = "C080"
number = "06"
  Q = enq + stx + "G" + address + number + eot  '전송 프레임 만들기
  
  Dim W As String
  'W = enq + stx + "W" + "C000" + "01" + "1" + eot '전송 프레임 만들기
  
With MSComm1        '그냥 .찍으면 .앞에 mscomm1을 생략했다는 뜻 정의하기
 .CommPort = 1             'Com1 사용
 .Settings = "9600,N,8,1"  '통신 9600bps, 패리티 없음, 데이타8 스톱1비트
 .PortOpen = True: ' Print '포트염"
 '.Output = W
 .Output = Q               'Print "명령전송"
 .InputLen = 1             '1자씩만 받아라..
    
RCVLOOP:

jup = jup + 1
If jup = 15 Then
.PortOpen = False
 MsgBox "응답없다~"
 End
End If

  rcv = .Input          '받기(1개씩)
  data = data + rcv     '받은거 계속 누적시키기
     Select Case rcv   '  받음거 체크해서
     Case eot          '  EOT (End of Text)
     GoTo 다받음p       '  받으면 받기 반복 빠져나감
     Case Else         '  데이타 받으면
     GoTo RCVLOOP      '  계속 돌아랏!!
     End Select        '  우핫핫핫
다받음p:
 .PortOpen = False '포트닫음
End With
'Print data
data = Mid(data, 4, 12)
Print data

'자료처리
'p0~p3
R = Mid(data, 2, 1)
 Select Case R
 Case 0
 p(3).a = 0: p(2).a = 0: p(1).a = 0: p(0).a = 0
 Case 1
 p(3).a = 0: p(2).a = 0: p(1).a = 0: p(0).a = 1
 Case 2
 p(3).a = 0: p(2).a = 0: p(1).a = 1: p(0).a = 0
 Case 3
 p(3).a = 0: p(2).a = 0: p(1).a = 1: p(0).a = 1
 Case 4
 p(3).a = 0: p(2).a = 1: p(1).a = 0: p(0).a = 0
 Case 5
 p(3).a = 0: p(2).a = 1: p(1).a = 0: p(0).a = 1
 Case 6
 p(3).a = 0: p(2).a = 1: p(1).a = 1: p(0).a = 0
 Case 7
 p(3).a = 0: p(2).a = 1: p(1).a = 1: p(0).a = 1
 Case 8
 p(3).a = 1: p(2).a = 0: p(1).a = 0: p(0).a = 0
 Case 9
 p(3).a = 1: p(2).a = 0: p(1).a = 0: p(0).a = 1
 Case "A"
 p(3).a = 1: p(2).a = 0: p(1).a = 1: p(0).a = 0
 Case "B"
 p(3).a = 1: p(2).a = 0: p(1).a = 1: p(0).a = 1
 Case "C"
 p(3).a = 1: p(2).a = 1: p(1).a = 0: p(0).a = 0
 Case "D"
 p(3).a = 1: p(2).a = 1: p(1).a = 0: p(0).a = 1
 Case "E"
 p(3).a = 1: p(2).a = 1: p(1).a = 1: p(0).a = 0
 Case "F"
 p(3).a = 1: p(2).a = 1: p(1).a = 1: p(0).a = 1
 End Select



'p4~p7
R = Mid(data, 1, 1)
 Select Case R
 Case 0
 p(7).a = 0: p(6).a = 0: p(5).a = 0: p(4).a = 0
 Case 1
 p(7).a = 0: p(6).a = 0: p(5).a = 0: p(4).a = 1
 Case 2
 p(7).a = 0: p(6).a = 0: p(5).a = 1: p(4).a = 0
 Case 3
 p(7).a = 0: p(6).a = 0: p(5).a = 1: p(4).a = 1
 Case 4
 p(7).a = 0: p(6).a = 1: p(5).a = 0: p(4).a = 0
 Case 5
 p(7).a = 0: p(6).a = 1: p(5).a = 0: p(4).a = 1
 Case 6
 p(7).a = 0: p(6).a = 1: p(5).a = 1: p(4).a = 0
 Case 7
 p(7).a = 0: p(6).a = 1: p(5).a = 1: p(4).a = 1
 Case 8
 p(7).a = 1: p(6).a = 0: p(5).a = 0: p(4).a = 0
 Case 9
 p(7).a = 1: p(6).a = 0: p(5).a = 0: p(4).a = 1
 Case "A"
 p(7).a = 1: p(6).a = 0: p(5).a = 1: p(4).a = 0
 Case "B"
 p(7).a = 1: p(6).a = 0: p(5).a = 1: p(4).a = 1
 Case "C"
 p(7).a = 1: p(6).a = 1: p(5).a = 0: p(4).a = 0
 Case "D"
 p(7).a = 1: p(6).a = 1: p(5).a = 0: p(4).a = 1
 Case "E"
 p(7).a = 1: p(6).a = 1: p(5).a = 1: p(4).a = 0
 Case "F"
 p(7).a = 1: p(6).a = 1: p(5).a = 1: p(4).a = 1
 End Select

'p8~pB
R = Mid(data, 4, 1)
 Select Case R
 Case 0
 p(11).a = 0: p(10).a = 0: p(9).a = 0: p(8).a = 0
 Case 1
 p(11).a = 0: p(10).a = 0: p(9).a = 0: p(8).a = 1
 Case 2
 p(11).a = 0: p(10).a = 0: p(9).a = 1: p(8).a = 0
 Case 3
 p(11).a = 0: p(10).a = 0: p(9).a = 1: p(8).a = 1
 Case 4
 p(11).a = 0: p(10).a = 1: p(9).a = 0: p(8).a = 0
 Case 5
 p(11).a = 0: p(10).a = 1: p(9).a = 0: p(8).a = 1
 Case 6
 p(11).a = 0: p(10).a = 1: p(9).a = 1: p(8).a = 0
 Case 7
 p(11).a = 0: p(10).a = 1: p(9).a = 1: p(8).a = 1
 Case 8
 p(11).a = 1: p(10).a = 0: p(9).a = 0: p(8).a = 0
 Case 9
 p(11).a = 1: p(10).a = 0: p(9).a = 0: p(8).a = 1
 Case "A"
 p(11).a = 1: p(10).a = 0: p(9).a = 1: p(8).a = 0
 Case "B"
 p(11).a = 1: p(10).a = 0: p(9).a = 1: p(8).a = 1
 Case "C"
 p(11).a = 1: p(10).a = 1: p(9).a = 0: p(8).a = 0
 Case "D"
 p(11).a = 1: p(10).a = 1: p(9).a = 0: p(8).a = 1
 Case "E"
 p(11).a = 1: p(10).a = 1: p(9).a = 1: p(8).a = 0
 Case "F"
 p(11).a = 1: p(10).a = 1: p(9).a = 1: p(8).a = 1
 End Select



'pC~pF
R = Mid(data, 3, 1)
 Select Case R
 Case 0
 p(15).a = 0: p(14).a = 0: p(13).a = 0: p(12).a = 0
 Case 1
 p(15).a = 0: p(14).a = 0: p(13).a = 0: p(12).a = 1
 Case 2
 p(15).a = 0: p(14).a = 0: p(13).a = 1: p(12).a = 0
 Case 3
 p(15).a = 0: p(14).a = 0: p(13).a = 1: p(12).a = 1
 Case 4
 p(15).a = 0: p(14).a = 1: p(13).a = 0: p(12).a = 0
 Case 5
 p(15).a = 0: p(14).a = 1: p(13).a = 0: p(12).a = 1
 Case 6
 p(15).a = 0: p(14).a = 1: p(13).a = 1: p(12).a = 0
 Case 7
 p(15).a = 0: p(14).a = 1: p(13).a = 1: p(12).a = 1
 Case 8
 p(15).a = 1: p(14).a = 0: p(13).a = 0: p(12).a = 0
 Case 9
 p(15).a = 1: p(14).a = 0: p(13).a = 0: p(12).a = 1
 Case "A"
 p(15).a = 1: p(14).a = 0: p(13).a = 1: p(12).a = 0
 Case "B"
 p(15).a = 1: p(14).a = 0: p(13).a = 1: p(12).a = 1
 Case "C"
 p(15).a = 1: p(14).a = 1: p(13).a = 0: p(12).a = 0
 Case "D"
 p(15).a = 1: p(14).a = 1: p(13).a = 0: p(12).a = 1
 Case "E"
 p(15).a = 1: p(14).a = 1: p(13).a = 1: p(12).a = 0
 Case "F"
 p(15).a = 1: p(14).a = 1: p(13).a = 1: p(12).a = 1
 End Select

'Print "P0 = "; p(0).a
'Print "P1 = "; p(1).a
'Print "P2 = "; p(2).a
'Print "P3 = "; p(3).a
'Print "P4 = "; p(4).a
'Print "P5 = "; p(5).a
'Print "P6 = "; p(6).a
'Print "P7 = "; p(7).a
'Print "P8 = "; p(8).a
'Print "P9 = "; p(9).a
'Print "PA = "; p(10).a
'Print "PB = "; p(11).a
'Print "PC = "; p(12).a
'Print "PD = "; p(13).a
'Print "PE = "; p(14).a
'Print "PF = "; p(15).a

 Text1 = ""
 Dim kl As Integer
 For kl = 0 To 15
  Text1 = Text1 & "P" & kl & " = " & p(kl).a & vbNewLine
 Next kl

'자료처리
'p0~p3
R = Mid(data, 6, 1)
 Select Case R
 Case 0
 p1(3).a = 0: p1(2).a = 0: p1(1).a = 0: p1(0).a = 0
 Case 1
 p1(3).a = 0: p1(2).a = 0: p1(1).a = 0: p1(0).a = 1
 Case 2
 p1(3).a = 0: p1(2).a = 0: p1(1).a = 1: p1(0).a = 0
 Case 3
 p1(3).a = 0: p1(2).a = 0: p1(1).a = 1: p1(0).a = 1
 Case 4
 p1(3).a = 0: p1(2).a = 1: p1(1).a = 0: p1(0).a = 0
 Case 5
 p1(3).a = 0: p1(2).a = 1: p1(1).a = 0: p1(0).a = 1
 Case 6
 p1(3).a = 0: p1(2).a = 1: p1(1).a = 1: p1(0).a = 0
 Case 7
 p1(3).a = 0: p1(2).a = 1: p1(1).a = 1: p1(0).a = 1
 Case 8
 p1(3).a = 1: p1(2).a = 0: p1(1).a = 0: p1(0).a = 0
 Case 9
 p1(3).a = 1: p1(2).a = 0: p1(1).a = 0: p1(0).a = 1
 Case "A"
 p1(3).a = 1: p1(2).a = 0: p1(1).a = 1: p1(0).a = 0
 Case "B"
 p1(3).a = 1: p1(2).a = 0: p1(1).a = 1: p1(0).a = 1
 Case "C"
 p1(3).a = 1: p1(2).a = 1: p1(1).a = 0: p1(0).a = 0
 Case "D"
 p1(3).a = 1: p1(2).a = 1: p1(1).a = 0: p1(0).a = 1
 Case "E"
 p1(3).a = 1: p1(2).a = 1: p1(1).a = 1: p1(0).a = 0
 Case "F"
 p1(3).a = 1: p1(2).a = 1: p1(1).a = 1: p1(0).a = 1
 End Select



'p4~p7
R = Mid(data, 5, 1)
 Select Case R
 Case 0
 p1(7).a = 0: p1(6).a = 0: p1(5).a = 0: p1(4).a = 0
 Case 1
 p1(7).a = 0: p1(6).a = 0: p1(5).a = 0: p1(4).a = 1
 Case 2
 p1(7).a = 0: p1(6).a = 0: p1(5).a = 1: p1(4).a = 0
 Case 3
 p1(7).a = 0: p1(6).a = 0: p1(5).a = 1: p1(4).a = 1
 Case 4
 p1(7).a = 0: p1(6).a = 1: p1(5).a = 0: p1(4).a = 0
 Case 5
 p1(7).a = 0: p1(6).a = 1: p1(5).a = 0: p1(4).a = 1
 Case 6
 p1(7).a = 0: p1(6).a = 1: p1(5).a = 1: p1(4).a = 0
 Case 7
 p1(7).a = 0: p1(6).a = 1: p1(5).a = 1: p1(4).a = 1
 Case 8
 p1(7).a = 1: p1(6).a = 0: p1(5).a = 0: p1(4).a = 0
 Case 9
 p1(7).a = 1: p1(6).a = 0: p1(5).a = 0: p1(4).a = 1
 Case "A"
 p1(7).a = 1: p1(6).a = 0: p1(5).a = 1: p1(4).a = 0
 Case "B"
 p1(7).a = 1: p1(6).a = 0: p1(5).a = 1: p1(4).a = 1
 Case "C"
 p1(7).a = 1: p1(6).a = 1: p1(5).a = 0: p1(4).a = 0
 Case "D"
 p1(7).a = 1: p1(6).a = 1: p1(5).a = 0: p1(4).a = 1
 Case "E"
 p1(7).a = 1: p1(6).a = 1: p1(5).a = 1: p1(4).a = 0
 Case "F"
 p1(7).a = 1: p1(6).a = 1: p1(5).a = 1: p1(4).a = 1
 End Select

'p8~pB
R = Mid(data, 8, 1)
 Select Case R
 Case 0
 p1(11).a = 0: p1(10).a = 0: p1(9).a = 0: p1(8).a = 0
 Case 1
 p1(11).a = 0: p1(10).a = 0: p1(9).a = 0: p1(8).a = 1
 Case 2
 p1(11).a = 0: p1(10).a = 0: p1(9).a = 1: p1(8).a = 0
 Case 3
 p1(11).a = 0: p1(10).a = 0: p1(9).a = 1: p1(8).a = 1
 Case 4
 p1(11).a = 0: p1(10).a = 1: p1(9).a = 0: p1(8).a = 0
 Case 5
 p1(11).a = 0: p1(10).a = 1: p1(9).a = 0: p1(8).a = 1
 Case 6
 p1(11).a = 0: p1(10).a = 1: p1(9).a = 1: p1(8).a = 0
 Case 7
 p1(11).a = 0: p1(10).a = 1: p1(9).a = 1: p1(8).a = 1
 Case 8
 p1(11).a = 1: p1(10).a = 0: p1(9).a = 0: p1(8).a = 0
 Case 9
 p1(11).a = 1: p1(10).a = 0: p1(9).a = 0: p1(8).a = 1
 Case "A"
 p1(11).a = 1: p1(10).a = 0: p1(9).a = 1: p1(8).a = 0
 Case "B"
 p1(11).a = 1: p1(10).a = 0: p1(9).a = 1: p1(8).a = 1
 Case "C"
 p1(11).a = 1: p1(10).a = 1: p1(9).a = 0: p1(8).a = 0
 Case "D"
 p1(11).a = 1: p1(10).a = 1: p1(9).a = 0: p1(8).a = 1
 Case "E"
 p1(11).a = 1: p1(10).a = 1: p1(9).a = 1: p1(8).a = 0
 Case "F"
 p1(11).a = 1: p1(10).a = 1: p1(9).a = 1: p1(8).a = 1
 End Select



'pC~pF
R = Mid(data, 5, 1)
 Select Case R
 Case 0
 p1(15).a = 0: p1(14).a = 0: p1(13).a = 0: p1(12).a = 0
 Case 1
 p1(15).a = 0: p1(14).a = 0: p1(13).a = 0: p1(12).a = 1
 Case 2
 p1(15).a = 0: p1(14).a = 0: p1(13).a = 1: p1(12).a = 0
 Case 3
 p1(15).a = 0: p1(14).a = 0: p1(13).a = 1: p1(12).a = 1
 Case 4
 p1(15).a = 0: p1(14).a = 1: p1(13).a = 0: p1(12).a = 0
 Case 5
 p1(15).a = 0: p1(14).a = 1: p1(13).a = 0: p1(12).a = 1
 Case 6
 p1(15).a = 0: p1(14).a = 1: p1(13).a = 1: p1(12).a = 0
 Case 7
 p1(15).a = 0: p1(14).a = 1: p1(13).a = 1: p1(12).a = 1
 Case 8
 p1(15).a = 1: p1(14).a = 0: p1(13).a = 0: p1(12).a = 0
 Case 9
 p1(15).a = 1: p1(14).a = 0: p1(13).a = 0: p1(12).a = 1
 Case "A"
 p1(15).a = 1: p1(14).a = 0: p1(13).a = 1: p1(12).a = 0
 Case "B"
 p1(15).a = 1: p1(14).a = 0: p1(13).a = 1: p1(12).a = 1
 Case "C"
 p1(15).a = 1: p1(14).a = 1: p1(13).a = 0: p1(12).a = 0
 Case "D"
 p1(15).a = 1: p1(14).a = 1: p1(13).a = 0: p1(12).a = 1
 Case "E"
 p1(15).a = 1: p1(14).a = 1: p1(13).a = 1: p1(12).a = 0
 Case "F"
 p1(15).a = 1: p1(14).a = 1: p1(13).a = 1: p1(12).a = 1
 End Select

'Print "P0 = "; p1(0).a
'Print "P1 = "; p1(1).a
'Print "P2 = "; p1(2).a
'Print "P3 = "; p1(3).a
'Print "P4 = "; p1(4).a
'Print "P5 = "; p1(5).a
'Print "P6 = "; p1(6).a
'Print "P7 = "; p1(7).a
'Print "P8 = "; p1(8).a
'Print "P9 = "; p1(9).a
'Print "PA = "; p1(10).a
'Print "PB = "; p1(11).a
'Print "PC = "; p1(12).a
'Print "PD = "; p1(13).a
'Print "PE = "; p1(14).a
'Print "PF = "; p1(15).a

 Text2 = ""
 Dim kk As Integer
 For kk = 0 To 15
  Text2 = Text2 & "P" & kk & " = " & p1(kk).a & vbNewLine
 Next kk

'자료처리
'p0~p3
R = Mid(data, 10, 1)
 Select Case R
 Case 0
 p2(3).a = 0: p2(2).a = 0: p2(1).a = 0: p2(0).a = 0
 Case 1
 p2(3).a = 0: p2(2).a = 0: p2(1).a = 0: p2(0).a = 1
 Case 2
 p2(3).a = 0: p2(2).a = 0: p2(1).a = 1: p2(0).a = 0
 Case 3
 p2(3).a = 0: p2(2).a = 0: p2(1).a = 1: p2(0).a = 1
 Case 4
 p2(3).a = 0: p2(2).a = 1: p2(1).a = 0: p2(0).a = 0
 Case 5
 p2(3).a = 0: p2(2).a = 1: p2(1).a = 0: p2(0).a = 1
 Case 6
 p2(3).a = 0: p2(2).a = 1: p2(1).a = 1: p2(0).a = 0
 Case 7
 p2(3).a = 0: p2(2).a = 1: p2(1).a = 1: p2(0).a = 1
 Case 8
 p2(3).a = 1: p2(2).a = 0: p2(1).a = 0: p2(0).a = 0
 Case 9
 p2(3).a = 1: p2(2).a = 0: p2(1).a = 0: p2(0).a = 1
 Case "A"
 p2(3).a = 1: p2(2).a = 0: p2(1).a = 1: p2(0).a = 0
 Case "B"
 p2(3).a = 1: p2(2).a = 0: p2(1).a = 1: p2(0).a = 1
 Case "C"
 p2(3).a = 1: p2(2).a = 1: p2(1).a = 0: p2(0).a = 0
 Case "D"
 p2(3).a = 1: p2(2).a = 1: p2(1).a = 0: p2(0).a = 1
 Case "E"
 p2(3).a = 1: p2(2).a = 1: p2(1).a = 1: p2(0).a = 0
 Case "F"
 p2(3).a = 1: p2(2).a = 1: p2(1).a = 1: p2(0).a = 1
 End Select



'p4~p7
R = Mid(data, 9, 1)
 Select Case R
 Case 0
 p2(7).a = 0: p2(6).a = 0: p2(5).a = 0: p2(4).a = 0
 Case 1
 p2(7).a = 0: p2(6).a = 0: p2(5).a = 0: p2(4).a = 1
 Case 2
 p2(7).a = 0: p2(6).a = 0: p2(5).a = 1: p2(4).a = 0
 Case 3
 p2(7).a = 0: p2(6).a = 0: p2(5).a = 1: p2(4).a = 1
 Case 4
 p2(7).a = 0: p2(6).a = 1: p2(5).a = 0: p2(4).a = 0
 Case 5
 p2(7).a = 0: p2(6).a = 1: p2(5).a = 0: p2(4).a = 1
 Case 6
 p2(7).a = 0: p2(6).a = 1: p2(5).a = 1: p2(4).a = 0
 Case 7
 p2(7).a = 0: p2(6).a = 1: p2(5).a = 1: p2(4).a = 1
 Case 8
 p2(7).a = 1: p2(6).a = 0: p2(5).a = 0: p2(4).a = 0
 Case 9
 p2(7).a = 1: p2(6).a = 0: p2(5).a = 0: p2(4).a = 1
 Case "A"
 p2(7).a = 1: p2(6).a = 0: p2(5).a = 1: p2(4).a = 0
 Case "B"
 p2(7).a = 1: p2(6).a = 0: p2(5).a = 1: p2(4).a = 1
 Case "C"
 p2(7).a = 1: p2(6).a = 1: p2(5).a = 0: p2(4).a = 0
 Case "D"
 p2(7).a = 1: p2(6).a = 1: p2(5).a = 0: p2(4).a = 1
 Case "E"
 p2(7).a = 1: p2(6).a = 1: p2(5).a = 1: p2(4).a = 0
 Case "F"
 p2(7).a = 1: p2(6).a = 1: p2(5).a = 1: p2(4).a = 1
 End Select

'p8~pB
R = Mid(data, 12, 1)
 Select Case R
 Case 0
 p2(11).a = 0: p2(10).a = 0: p2(9).a = 0: p2(8).a = 0
 Case 1
 p2(11).a = 0: p2(10).a = 0: p2(9).a = 0: p2(8).a = 1
 Case 2
 p2(11).a = 0: p2(10).a = 0: p2(9).a = 1: p2(8).a = 0
 Case 3
 p2(11).a = 0: p2(10).a = 0: p2(9).a = 1: p2(8).a = 1
 Case 4
 p2(11).a = 0: p2(10).a = 1: p2(9).a = 0: p2(8).a = 0
 Case 5
 p2(11).a = 0: p2(10).a = 1: p2(9).a = 0: p2(8).a = 1
 Case 6
 p2(11).a = 0: p2(10).a = 1: p2(9).a = 1: p2(8).a = 0
 Case 7
 p2(11).a = 0: p2(10).a = 1: p2(9).a = 1: p2(8).a = 1
 Case 8
 p2(11).a = 1: p2(10).a = 0: p2(9).a = 0: p2(8).a = 0
 Case 9
 p2(11).a = 1: p2(10).a = 0: p2(9).a = 0: p2(8).a = 1
 Case "A"
 p2(11).a = 1: p2(10).a = 0: p2(9).a = 1: p2(8).a = 0
 Case "B"
 p2(11).a = 1: p2(10).a = 0: p2(9).a = 1: p2(8).a = 1
 Case "C"
 p2(11).a = 1: p2(10).a = 1: p2(9).a = 0: p2(8).a = 0
 Case "D"
 p2(11).a = 1: p2(10).a = 1: p2(9).a = 0: p2(8).a = 1
 Case "E"
 p2(11).a = 1: p2(10).a = 1: p2(9).a = 1: p2(8).a = 0
 Case "F"
 p2(11).a = 1: p2(10).a = 1: p2(9).a = 1: p2(8).a = 1
 End Select



'pC~pF
R = Mid(data, 9, 1)
 Select Case R
 Case 0
 p2(15).a = 0: p2(14).a = 0: p2(13).a = 0: p2(12).a = 0
 Case 1
 p2(15).a = 0: p2(14).a = 0: p2(13).a = 0: p2(12).a = 1
 Case 2
 p2(15).a = 0: p2(14).a = 0: p2(13).a = 1: p2(12).a = 0
 Case 3
 p2(15).a = 0: p2(14).a = 0: p2(13).a = 1: p2(12).a = 1
 Case 4
 p2(15).a = 0: p2(14).a = 1: p2(13).a = 0: p2(12).a = 0
 Case 5
 p2(15).a = 0: p2(14).a = 1: p2(13).a = 0: p2(12).a = 1
 Case 6
 p2(15).a = 0: p2(14).a = 1: p2(13).a = 1: p2(12).a = 0
 Case 7
 p2(15).a = 0: p2(14).a = 1: p2(13).a = 1: p2(12).a = 1
 Case 8
 p2(15).a = 1: p2(14).a = 0: p2(13).a = 0: p2(12).a = 0
 Case 9
 p2(15).a = 1: p2(14).a = 0: p2(13).a = 0: p2(12).a = 1
 Case "A"
 p2(15).a = 1: p2(14).a = 0: p2(13).a = 1: p2(12).a = 0
 Case "B"
 p2(15).a = 1: p2(14).a = 0: p2(13).a = 1: p2(12).a = 1
 Case "C"
 p2(15).a = 1: p2(14).a = 1: p2(13).a = 0: p2(12).a = 0
 Case "D"
 p2(15).a = 1: p2(14).a = 1: p2(13).a = 0: p2(12).a = 1
 Case "E"
 p2(15).a = 1: p2(14).a = 1: p2(13).a = 1: p2(12).a = 0
 Case "F"
 p2(15).a = 1: p2(14).a = 1: p2(13).a = 1: p2(12).a = 1
 End Select

'Print "P0 = "; p2(0).a
'Print "P1 = "; p2(1).a
'Print "P2 = "; p2(2).a
'Print "P3 = "; p2(3).a
'Print "P4 = "; p2(4).a
'Print "P5 = "; p2(5).a
'Print "P6 = "; p2(6).a
'Print "P7 = "; p2(7).a
'Print "P8 = "; p2(8).a
'Print "P9 = "; p2(9).a
'Print "PA = "; p2(10).a
'Print "PB = "; p2(11).a
'Print "PC = "; p2(12).a
'Print "PD = "; p2(13).a
'Print "PE = "; p2(14).a
'Print "PF = "; p2(15).a

 Text3 = ""
 Dim kj As Integer
 For kj = 0 To 15
  Text3 = Text3 & "P" & kj & " = " & p2(kj).a & vbNewLine
 Next kj

End Sub
