VERSION 5.00
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
   Begin VB.CommandButton Command1 
      Caption         =   "�������� ���� ����..."
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
Dim temp As String



For i = 1 To 10

     temp = Int(Rnd(1) * 10)

    '�̷��� �ϸ� 0 ���� 10���� �̰� ����.
    '���� Rnd�� 0���� 1������ �߰� ���ڸ� ���մϴ�.
    '�׷� 10�� ������� �ű⿡ 10�� ���� ȿ���� ������?
     'int���� �׷��� �ϸ� �Ҽ� ���� ������ ������ �Ҽ����� ���ֱ� ���� ���̰��...

    Text1 = Text1 & " " & temp
Next i

End Sub
