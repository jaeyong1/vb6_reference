VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "c:\out.pla"
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Text            =   "c:\ex1.pla"
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim index() As String   'ũ����� �迭����
Dim sizeOfIndex As Integer



Private Sub Command1_Click()

Dim file_a_line As String '���� ���� �д³�
Dim filename1 As String '�����д³� ���ϸ�
Dim filename2 As String '���Ͼ��³� ���ϸ�

Dim filenum1 As Integer '�д³� �̸� (�������̸��̶�� ���� ��)
Dim filenum2 As Integer '���³� �̸�


filename1 = Text1.Text 'ȭ�� �ؽ�Ʈ�ڽ�1�� �ִ� ������ ���ϸ����� ����"
filename2 = Text2.Text 'ȭ�� �ؽ�Ʈ�ڽ�2�� �ִ� ������ ���ϸ����� ����"
filenum1 = FreeFile
filenum2 = FreeFile

Dim �ٹٲް�ģ��Ʈ�� As String


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���� �д°�
Open filename1 For Input As filenum1    '�б�� ����
Do Until EOF(filenum1)                   '����о�
   Line Input #filenum1, file_1_line
    �ٹٲް�ģ��Ʈ�� = Replace(file_1_line, vbLf, vbCrLf)
Loop '������� Do����

Close filenum1   '���ϴݱ�
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���� �д°� ��

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�ٹٲ� �ذ��� ���� ����
Dim dmn As Integer '�ӽ�����(�ٹٲ޵�)
dmn = FreeFile

Open "c:\dummy.dat" For Output As dmn
    Print #dmn, �ٹٲް�ģ��Ʈ��;
Close dmn
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�ٹٲ� �ذ��� ���� ���ⳡ

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�ٹٲ� �ذ��� ���� �б�
Dim isfinish As Boolean
Open "c:\dummy.dat" For Input As filenum1    '�б�� ����
Do Until EOF(filenum1)                 '����о�
   Line Input #filenum1, file_1_line
    func1 (file_1_line)
    'If isfinish = False Then Exit Do
    
Loop '������� Do����

Close filenum1   '���ϴݱ�
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�ٹٲ� �ذ��� ���� �бⳡ

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���� ���°�

Open filename2 For Output As FreeFile 'append: �ڿ��᳻������.
Print #filenum2, '�տ��� �ٹٲ޾��ؼ� ���ٳ���, �ƹ����� �Ⱦ��ϱ� �ٹٲ�..
Print #filenum2, �ٹٲ��ֱ�����Ʈ��; '��¥�ð����     ;�ϸ� �ٹٲ� ����.
Close #filenum2

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���� �д°� ��


End Sub

Function func1(strln As String) As Boolean



If Mid(strln, 1, 2) = ".i" Then '.i�� ��Ÿ����
    Debug.Print ".i�߰�"
    MsgBox "���� : " + Trim(Mid(strln, 3, 10)) 'trim:�¿���鹮������ , val:����->���ڷ� �ν�(��� ����)
    sizeOfIndex = Trim(Mid(strln, 3, 10))
    ReDim index(sizeOfIndex, 10000)   '2�����迭 �˳��� ����


ElseIf Mid(strln, 1, 2) = ".e" Then '.o�� ��Ÿ����
    Debug.Print ".e�߰�"
    
ElseIf Mid(strln, 1, 2) = ".o" Then '.o�� ��Ÿ����
    Debug.Print ".o�߰�"
ElseIf Mid(strln, 1, 1) = "#" Then
    Debug.Print "# �ּ�"
    

End If

func1 = True

End Function

