VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows 기본값
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2400
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Cls

Dim etx As String: etx = Chr$(3)
Dim eot As String: eot = Chr$(4)
Dim enq As String: enq = Chr$(5)
Dim ack As String: ack = Chr$(6)
Dim nak As String: nak = Chr$(21)
Dim stx As String: stx = Chr$(2)
Dim Q As String

  Q = enq + "00RSS" + "01" + "06%MW100" + eot
   Q = enq + "00RSS0106%IW000" + eot
  
With MSComm1        '그냥 .찍으면 .앞에 mscomm1을 생략했다는 뜻 정의하기
 .CommPort = 1             'Com1 사용
 .Settings = "9600,N,8,1"  '통신 9600bps, 패리티 없음, 데이타8 스톱1비트
 .PortOpen = True: ' Print '포트염"
 Print ">OPEN"
 .Output = Q               'Print "명령전송"
 Print ">SND"
 .InputLen = 1             '1자씩만 받아라..
    
    Dim i As Integer
    For i = 1 To 20
Print i
  rcv = .Input          '받기(1개씩)
Print rcv

Next i


.PortOpen = False
Print ">CLOSE"
End With
Print ">END"

End Sub

