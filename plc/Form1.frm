VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "�������� �����б�"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "����"
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
   StartUpPosition =   3  'Windows �⺻��
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
      Caption         =   "�����߻� ������"
      Height          =   975
      Left            =   7440
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   615
      Left            =   7560
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�۵�"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
Dim j As Integer '�ð�����
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
jup = 0 '����ȉ����� ���������..
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
  Q = enq + stx + "G" + address + number + eot  '���� ������ �����
  
  Dim W As String
  'W = enq + stx + "W" + "C000" + "01" + "1" + eot '���� ������ �����
  
With MSComm1        '�׳� .������ .�տ� mscomm1�� �����ߴٴ� �� �����ϱ�
 .CommPort = 1             'Com1 ���
 .Settings = "9600,N,8,1"  '��� 9600bps, �и�Ƽ ����, ����Ÿ8 ����1��Ʈ
 .PortOpen = True: ' Print '��Ʈ��"
 '.Output = W
 .Output = Q               'Print "�������"
 .InputLen = 1             '1�ھ��� �޾ƶ�..
    
RCVLOOP:

jup = jup + 1
If jup = 15 Then
.PortOpen = False
 MsgBox "�������~"
 End
End If

  rcv = .Input          '�ޱ�(1����)
  data = data + rcv     '������ ��� ������Ű��
     Select Case rcv   '  ������ üũ�ؼ�
     Case eot          '  EOT (End of Text)
     GoTo �ٹ���p       '  ������ �ޱ� �ݺ� ��������
     Case Else         '  ����Ÿ ������
     GoTo RCVLOOP      '  ��� ���ƶ�!!
     End Select        '  ��������
�ٹ���p:
 .PortOpen = False '��Ʈ����
End With
'Print data
data = Mid(data, 4, 12)
Print data

'�ڷ�ó��
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

'�ڷ�ó��
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

'�ڷ�ó��
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
