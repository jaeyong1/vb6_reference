VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command1 
      Caption         =   "������"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1320
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '���� ����
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�ϴ� com��Ʈ�� ����� ���ʿ� �����ȭ�⸦ �߰��ؾ���..
'�߰��ϴ� ���
'���� �޴�����
'������Ʈ -> ������ҿ���
'Microsoft Comm Control 6.0�� üũ���� Ȯ���ϸ�
'�����ȭ�� �������� �����.
'���� ��ȭ�⸦ Ŭ��, �������� ���콺�� �巡��-> ������ ��ȭ�Ⱑ ���δ�.


Private Sub Command1_Click()

    MSComm1.CommPort = 1            'COM1�� ���.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '�Է¹��� ũ�⸦ 1Byte�� ��.
    MSComm1.PortOpen = True         '�����Ʈ ����.
    MSComm1.Output = "������������ڿ�" '������ �̰� com1�� ���۵ȴ�.
    MSComm1.PortOpen = False        '�����Ʈ�ݱ� (��� ������� �ȴݾƾ߰���)
 


End Sub



'���� �޴¹��.. ���Ը��ϸ� com1���� �������� ������ ���ͷ��Ͱ� �ɸ�����
'�� �̺�Ʈ�� �߻��ϴ°��̴�.
'�̺�Ʈ ��������� �� ������ select�� Ȯ���ؼ� �����̸� ���ڸ� �޴´�.

Private Sub MSComm1_OnComm()

Dim rcvtem
  
Select Case MSComm1.CommEvent
      Case comEvReceive '<- �����̺�Ʈ �϶�.. �̰Ÿ��� ������ ��û����. �ٸ��̺�Ʈ�� �������� �õ��ϴ°� �������� ��.

        If MSComm1.InBufferCount Then ' �񱳱����� ����? �ƴϴ�. ���۰� 0�̸� false 0�̾ƴѼ��ڸ� true�ΰ�!, ��Ժ��� �������� ������..
           rcvtemp = MSComm1.Input      '���ۿ� �ִ°� ����
           Label1 = rcvtemp             'ȭ�鿡 ���
        End If

End Select
End Sub


'�׽�Ʈ�� ���غ����� �۵��ϴ� �ҽ��ϵ�.. ����-�ٿ��ֱ� �Ű��̶�..
