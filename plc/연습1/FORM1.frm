VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3360
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
Dim etx As String: etx = Chr$(3)
Dim eot As String: eot = Chr$(4)
Dim enq As String: enq = Chr$(5)
Dim ack As String: ack = Chr$(6)
Dim nak As String: nak = Chr$(21)
Dim stx As String: stx = Chr$(2)
Dim q As String
Dim I As Integer
Dim RCV As String
'q = enq + "RSS0106%MW100" + etx
q = Chr(5) & "00RSS" & "01" & "08%QW0.3.0" & Chr(4)
With MSComm1        '�׳� .������ .�տ� mscomm1�� �����ߴٴ� �� �����ϱ�
 .CommPort = 1             'Com1 ���
 .Settings = "9600,N,8,1"  '��� 9600bps, �и�Ƽ ����, ����Ÿ8 ����1��Ʈ
 .PortOpen = True: ' Print '��Ʈ��"
 '.Output = W
 .Output = q               'Print "�������"
 .InputLen = 1             '1�ھ��� �޾ƶ�..

RCVLO:

If RCV = "" Then
  I = I + 1
   If I = 10000 Then
   GoTo EXITS
   End If
End If
'-----
  RCV = .Input          '�ޱ�(1����)

Print RCV
'-----
GoTo RCVLO


EXITS:
Print "EE"

End With
End
End Sub

