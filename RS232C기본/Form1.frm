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
   StartUpPosition =   3  'Windows �⺻��
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
  
With MSComm1        '�׳� .������ .�տ� mscomm1�� �����ߴٴ� �� �����ϱ�
 .CommPort = 1             'Com1 ���
 .Settings = "9600,N,8,1"  '��� 9600bps, �и�Ƽ ����, ����Ÿ8 ����1��Ʈ
 .PortOpen = True: ' Print '��Ʈ��"
 Print ">OPEN"
 .Output = Q               'Print "�������"
 Print ">SND"
 .InputLen = 1             '1�ھ��� �޾ƶ�..
    
    Dim i As Integer
    For i = 1 To 20
Print i
  rcv = .Input          '�ޱ�(1����)
Print rcv

Next i


.PortOpen = False
Print ">CLOSE"
End With
Print ">END"

End Sub

