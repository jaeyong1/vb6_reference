VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "133_������_�л����� ��"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form2"
   ScaleHeight     =   5490
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.TextBox Text6 
      DataField       =   "�ڰ���"
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
      Left            =   1320
      TabIndex        =   19
      Top             =   3750
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      DataField       =   "�ּ�"
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
      Left            =   1320
      TabIndex        =   18
      Top             =   2865
      Width           =   6135
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
      Left            =   1320
      TabIndex        =   17
      Top             =   1995
      Width           =   1575
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
      Left            =   1320
      TabIndex        =   16
      Top             =   1110
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      DataField       =   "�ڵ���"
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
      Left            =   5280
      TabIndex        =   15
      Top             =   3750
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      DataField       =   "��ȭ��ȣ"
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
      Left            =   5280
      TabIndex        =   14
      Top             =   1995
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      DataField       =   "�ֹε�Ϲ�ȣ"
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
      Left            =   5280
      TabIndex        =   13
      Top             =   1110
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�ݱ�"
      Height          =   495
      Left            =   6330
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��������"
      Height          =   495
      Left            =   4830
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Left            =   3330
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˻�"
      Height          =   495
      Left            =   1830
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Է�"
      Height          =   495
      Left            =   330
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\VB6\133_������\133_������_�л�����.mdb"
      DefaultCursorType=   0  '�⺻ Ŀ��
      DefaultType     =   2  'ODBC���
      Exclusive       =   0   'False
      Height          =   405
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '���̳ʼ�
      RecordSource    =   "Student"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "�ڵ��� :"
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
      Left            =   4320
      TabIndex        =   12
      Top             =   3840
      Width           =   840
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "�ڰ��� :"
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
      Left            =   360
      TabIndex        =   11
      Top             =   3840
      Width           =   840
   End
   Begin VB.Label Label6 
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
      Left            =   360
      TabIndex        =   10
      Top             =   2955
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "�� ȭ �� ȣ :"
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
      Left            =   3720
      TabIndex        =   9
      Top             =   2085
      Width           =   1365
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
      Left            =   360
      TabIndex        =   8
      Top             =   2085
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "�ֹε�Ϲ�ȣ :"
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
      Left            =   3720
      TabIndex        =   7
      Top             =   1200
      Width           =   1470
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
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   840
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
      Left            =   2895
      TabIndex        =   0
      Top             =   240
      Width           =   2085
   End
End
Attribute VB_Name = "Form2"
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
    Data1.Recordset.Delete
    Data1.Recordset.MoveNext
    
    If Data1.Recordset.EOF = True Then
        If Data1.Recordset.RecordCount = 0 Then
            Data1.Recordset.AddNew
            Text1.SetFocus
        Else
            Data1.Recordset.MoveLast
        End If
    End If
End Sub

Private Sub Command4_Click()
    Form2.Hide
    Form3.Show
End Sub

Private Sub Command5_Click()
    Form2.Hide
    Form1.Show
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text2_LostFocus()

End Sub
