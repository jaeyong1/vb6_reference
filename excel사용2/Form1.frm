VERSION 5.00
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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���� ���ε� ���(������ �����ؼ� �����Ű�� ���)
'���� VB���� ������Ʈ �޴�-������  - Microsoft Excel 11.0 Object Library ��
'üũ�ϰ� Ȯ���� ������ ���� ��ü�� �����մϴ�.
'
'

Option Explicit
 
Dim xlApp As New Excel.Application
 
Sub Command1_Click()
MsgBox App.Path

    Const XL_NOTRUNNING As Long = 429 '������ ������ ����ǰ� ���� ������ 429 ������ �߻�
 
    On Error GoTo ShowName_Err '������ �߻��ϸ�(������ ������ ����ǰ� ���� �ʴٸ�) ShowName_Err ������ �̵�
    Set xlApp = GetObject(, "Excel.Application") '������ ����ǰ� �ֳ� üũ
    xlApp.Visible = True '���� ǥ��
    
    xlApp.DisplayAlerts = False
    
    
    xlApp.Workbooks.Open "C:\test.xls" '���� ����, �ݱ�, ����
    
    
ShowName_End:
    Exit Sub
ShowName_Err:
    If Err = XL_NOTRUNNING Then '������ ���������� ���� ���
        Set xlApp = New Excel.Application '���� ����
        xlApp.Workbooks.Add '��ũ�� �߰�
        Resume Next '���� ���� �߻� ��ġ(GetObject �� ��)�� ����
    Else
        MsgBox Err.Number & " - " & Err.Description '�׷��� ���� ������ �߻��ϸ� ���� ��ȣ �� ���� ���� ǥ��
    End If
    Resume ShowName_End '���ν����� ����
    
End Sub

Private Sub Command2_Click()
xlApp.Quit '���� ���α׷� ����
Set xlApp = Nothing '���� ���ø����̼� ��ü �޸𸮿��� ����
End Sub
