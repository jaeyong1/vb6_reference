VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  '���� ����
   Caption         =   "ä�� ����"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.TextBox text1 
      Height          =   6975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   7
      Top             =   480
      Width           =   4935
   End
   Begin VB.TextBox txtinput 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   7800
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "send"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��������"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���α׷� ����"
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      Top             =   8400
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   5760
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "~"
      Height          =   135
      Left            =   6360
      TabIndex        =   9
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "��ȭǥ��â�� ������ ä��â�� �����ϰ� ��..."
      Height          =   1095
      Left            =   5640
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblconnect 
      Caption         =   "Ŭ���̾�Ʈ ���ӵ�.."
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   8520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label ipview 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "������.."
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() '���α׷� ����
End
End Sub

Private Sub Command2_Click()  '����
Command2.Enabled = False
Command3.Enabled = False
txtinput.Enabled = False
text1.Enabled = False



Winsock1(1).Close
End Sub

Private Sub Command3_Click()  '���Է��� ����
Winsock1(1).SendData txtinput.Text
 text1.Text = text1.Text + "��> " & txtinput.Text + vbNewLine
 txtinput.SetFocus
 txtinput.Text = ""
End Sub


Private Sub Form_Load()       '����

Load Winsock1(0)
Winsock1(0).Protocol = sckTCPProtocol
Winsock1(0).LocalPort = 2000
Winsock1(0).Listen                'Ŭ���̾�Ʈ ���Ӵ�����

End Sub

Private Sub Label3_Click()
MsgBox "<<����� ġƮ�ڵ�>>" & vbCrLf & "  ��� �ý�������" & vbCrLf & "  �޼����ڽ�"
End Sub

Private Sub text1_Click()     'ǥ��â ���~
text1 = ""
txtinput.SetFocus

End Sub


Private Sub Winsock1_Close(Index As Integer)  '����
Winsock1(1).Close

End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
                                            'Ŭ���̾�Ʈ���� ���ӿ䱸�� ���ð��

  'Load Winsock1(0)
  Winsock1(1).Accept requestID  '�������
  
  
  Command3.Enabled = True 'send ��ư ��밡
  txtinput.Enabled = True '�Է�â ��밡
  txtinput.SetFocus   '�Է�â�� Ŀ���̵�
  lblconnect.Enabled = True  ' Ŭ, ������ ǥ��
  
  Command2.Enabled = True  '�������� ��ư��밡
  text1.Enabled = True  'ä��â Ŭ������ ������
  
  
 
 'Winsock1(Index).SendData "!!!" + CStr(Index)
 
End Sub
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)   '�۵���
    Dim Gstr As String
    Winsock1(1).GetData Gstr
    text1.Text = text1.Text + "��> " & Gstr + vbNewLine
    txtinput.SetFocus


End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "����! �����߻�"
Command2_Click
End Sub
