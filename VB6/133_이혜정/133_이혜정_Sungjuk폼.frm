VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "133_������_SungJuk ��"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form3"
   ScaleHeight     =   5490
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command6 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   16
      Top             =   3690
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   15
      Top             =   2805
      Width           =   975
   End
   Begin VB.TextBox Text4 
      DataField       =   "���"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   2865
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      DataField       =   "����"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   3750
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\VB6\133_������\133_������_�л�����.mdb"
      DefaultCursorType=   0  '�⺻ Ŀ��
      DefaultType     =   2  'ODBC���
      Exclusive       =   0   'False
      Height          =   405
      Left            =   390
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '���̳ʼ�
      RecordSource    =   "Student"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Է�"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˻�"
      Height          =   495
      Left            =   2360
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�л�����"
      Height          =   495
      Left            =   4360
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�ݱ�"
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "�й�"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1110
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      DataField       =   "�̸�"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1995
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      DataField       =   "�߰����"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2865
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      DataField       =   "�⸻���"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   3750
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�� �� �� ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2925
      TabIndex        =   12
      Top             =   240
      Width           =   2085
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��  �� :"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   390
      TabIndex        =   11
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "��  �� :"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   390
      TabIndex        =   10
      Top             =   2085
      Width           =   840
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "�߰���� :"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   390
      TabIndex        =   9
      Top             =   2955
      Width           =   1050
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "�⸻��� :"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   390
      TabIndex        =   8
      Top             =   3840
      Width           =   1050
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Data1.Recordset.AddNew
    Text1.SetFocus
End Sub

Private Sub Command2_Click()
    Dim result
    result = InputBox("�̸��� �Է��ϼ���", "ã��")
    If result = "" Then
        MsgBox "�̸��� �Էµ��� �ʾҽ��ϴ�.", vbOKOnly + vbCritical, "�����޽���"
        Exit Sub
    End If
    
    Data1.Recordset.FindNext "�̸�='" & result & "'"
    If Data1.Recordset.NoMatch Then
        Data1.Recordset.FindFirst "�̸�='" & result & "'"
    End If
    If Data1.Recordset.NoMatch Then
        MsgBox "ã�� �ڷᰡ �����ϴ�.", vbOKOnly + vbExclamation, "�޽���"
    End If
End Sub

Private Sub Command3_Click()
    Dim num1, num2
    num1 = Val(Text5.Text)
    num2 = Val(Text6.Text)
    Text4.Text = Str((num1 + num2) / 2)
End Sub

Private Sub Command4_Click()
    Form3.Hide
    Form2.Show
End Sub

Private Sub Command5_Click()
    Form3.Hide
    Form1.Show
End Sub

Private Sub Command6_Click()
    If Val(Text4.Text) >= 90 Then
        Text2.Text = "A"
    ElseIf Val(Text4.Text) >= 80 Then
        Text2.Text = "B"
    ElseIf Val(Text4.Text) >= 70 Then
        Text2.Text = "C"
    ElseIf Val(Text4.Text) >= 60 Then
        Text2.Text = "D"
    Else
        Text2.Text = "F"
    End If
End Sub
