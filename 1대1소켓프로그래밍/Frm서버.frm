VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Frm���� 
   Caption         =   "����"
   ClientHeight    =   3390
   ClientLeft      =   1215
   ClientTop       =   1830
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   5925
   Begin VB.TextBox Txt���� 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox TxtPort 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Text            =   "2000"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox TxtIP 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Text            =   "61.84.8.254"
      Top             =   480
      Width           =   1335
   End
   Begin VB.ListBox Lst���� 
      Height          =   2040
      Left            =   2760
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "�����۵�"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "PORT"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "IP"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "Frm����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdStart_Click()

TxtIP = Winsock1.LocalIP
Winsock1.LocalPort = TxtPort
Winsock1.Listen
���ϻ��¾˸� 0
End Sub

Private Sub Form_Load()
FrmŬ���̾�Ʈ.Show

���ϻ��� = Array("����", "����", "���Ӵ��", "��������", "ȣ��Ʈ�˻���", _
    "ȣ��Ʈã��", "������", "�����", "�ݰ��ִ���", "����")
End Sub

Private Sub Txt����_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Winsock1.SendData Txt����.Text
    Lst����.AddItem Txt����
    Txt����.Text = ""
End If
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Lst����.AddItem "�����û"
Winsock1.Close
Winsock1.Accept requestID
'Lst����.AddItem requestID
���ϻ��¾˸� 0


'�ӽ÷� �߰�------
Winsock1.SendData "����"
'������� �߰��Ѱ� �����ϱ�..------
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData sData, vbString, bytesTotal
Lst����.AddItem sData
End Sub
