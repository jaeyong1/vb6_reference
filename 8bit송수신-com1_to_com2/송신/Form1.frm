VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4080
      Top             =   120
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "�۽� "
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()

Dim q As String
End Sub

Private Sub Form_Click()
End
End Sub

Private Sub Timer1_Timer()

q = " ~" & i
i = i + 1
    MSComm1.CommPort = 1            'COM1�� ���.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '�Է¹��� ũ�⸦ 1Byte�� ��.
    MSComm1.PortOpen = True         '�����Ʈ ����.
    MSComm1.Output = q
    MSComm1.PortOpen = False

Label1 = Label1 & q
End Sub
