VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      Height          =   615
      Left            =   6480
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�����"
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��"
      Height          =   615
      Left            =   6480
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   4815
      Left            =   120
      ScaleHeight     =   4755
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   240
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�׸��⿬��"
      Height          =   615
      Left            =   6480
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '���� ����
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   6120
      X2              =   6120
      Y1              =   3600
      Y2              =   4560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���������� ������ ������!! ������ ������ ���������..

'���̰� ����
Const pi = 3.14159265358979

'�ð�ٴ� �����Ҷ� ���� ���,
'�Լ��ȿ� ���� ���������� �Լ�������
'���� ������Ƿ� ���� �ִµ��� ��� ���ɼ� �ְ�
'���⿡�� ����
Dim NowAngle As Integer


Private Sub Command1_Click() '�׸��⿬�� ��ư ������

' ���׸���
Picture1.DrawWidth = 10 '�� ����
Picture1.PSet (500, 300), RGB(80, 100, 255) '(x,y)��ġ�� RGB����� ����


'�� �׸���
Picture1.DrawWidth = 2 '�� ���⺯�� (������ ����´ٰ� �ʹ� ũ�� ����)
Picture1.Line (1000, 1000)-(1500, 1500), RGB(100, 200, 0)


'���׸���
Picture1.Circle (3000, 1200), 300, RGB(180, 120, 255)



End Sub

Private Sub Command2_Click() '�� ��ư ������
'��ġ ����ֱ�
'Line1.X1 = 6000
'Line1.Y1 = 6000
'Line1.X2 = 200
'Line1.Y2 = 5000

'��ġ�� ������Ű��..
Line1.X1 = Line1.X1 + 100
Line1.X2 = Line1.X2 + 100

End Sub

Private Sub Command3_Click() '����� ��ư

'���� �����
Picture1.Picture = Nothing

End Sub


Private Sub Command4_Click() '�ٴõ�����

NowAngle = NowAngle + 10 '���� ����Ű�°� 360�� ����

Label1 = NowAngle '���̺� �� ǥ��

'x1,y1�� �߽���
'x2,y2�� ��ȭ�� ����
Line1.X1 = 6800 '������ ����
Line1.Y1 = 4800 '���� ����
Line1.X2 = 6800 + 1000 * Cos((90 - NowAngle) * (pi / 180))
Line1.Y2 = 4800 - 1000 * Cos(NowAngle * (pi / 180))

End Sub


