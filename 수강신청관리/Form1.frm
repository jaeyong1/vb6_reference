VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '���� ����
   Caption         =   "������û ���α׷�"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9810
   ForeColor       =   &H00FFFFC0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command5 
      Caption         =   "DataBase"
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "�ð�ǥ����"
      Height          =   615
      Left            =   2400
      TabIndex        =   10
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���û���"
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�߰�"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�˻�"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   4020
      Left            =   6840
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '����
      Caption         =   "������ : 2005. 12. 10  ..."
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   17
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��� ����
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  '���� ����
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�����ڵ�"
      Height          =   180
      Left            =   960
      TabIndex        =   3
      Top             =   2880
      Width           =   720
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Height          =   2055
      Left            =   360
      TabIndex        =   16
      Top             =   2520
      Width           =   5535
   End
   Begin VB.Label Label7 
      Alignment       =   2  '��� ����
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  '���� ����
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "������û �������� (�̽�û/������)"
      Height          =   180
      Index           =   0
      Left            =   1200
      TabIndex        =   9
      Top             =   1320
      Width           =   2880
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "������û�� :"
      Height          =   180
      Left            =   720
      TabIndex        =   8
      Top             =   720
      Width           =   1020
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Height          =   1815
      Left            =   360
      TabIndex        =   15
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "<������û����>"
      Height          =   255
      Left            =   7440
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Height          =   6375
      Left            =   -120
      TabIndex        =   14
      Top             =   -120
      Width           =   9975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
getinfo (Text1)

If info.���� Then
    Label3 = info.�����
Else
    Label3 = "�ش������ �����ϴ�."
End If


End Sub

Private Sub Command2_Click()

If (info.����) Then
    
    getinfo (Text1) '������ ����
        
    If (Int(info.��û�ο�) >= Int(info.�ִ��ο�)) Then
        MsgBox "��û�ο��� �ʰ��Ǿ����ϴ�.", , "����"
        Exit Sub
    End If
    
    List1.AddItem (Label3) '����Ʈ�� �׸��߰�
    
    Form3.gridData.TextMatrix(Form3.gridData.Rows - 1, 0) = txtName     '�̸� �߰�
    Form3.gridData.TextMatrix(Form3.gridData.Rows - 1, 1) = Text1       '�����ڵ� �߰�
    Form3.gridData.Rows = Form3.gridData.Rows + 1
    Form2.gridSugang.TextMatrix(info.row, 4) = Form2.gridSugang.TextMatrix(info.row, 4) + 1
    
       
Else
    MsgBox "�˻������ �ùٸ��� �ʽ��ϴ�. �ٽ� �˻����ּ���.", , "�˸�"
End If

End Sub

Private Sub Command3_Click()
' MsgBox List1.ListCount ����

'MsgBox List1.ListIndex '������ġ



Dim tmpstr, code As String
tmpstr = List1.Text

'---------------- �����ڵ� ã�� -------------
Dim i, tmpcnt As Integer

For i = 1 To Form2.gridSugang.Rows - 1
    If Form2.gridSugang.TextMatrix(i, 2) = tmpstr Then
        info.�����ڵ� = Form2.gridSugang.TextMatrix(i, 1)
        info.���� = False
    End If
Next

code = info.�����ڵ�


'---------------- �������� �����Ϳ��� ���� -------------

Dim datacount As Integer
With Form3.gridData

datacount = .Rows - 2

For i = 1 To datacount
    If (.TextMatrix(i, 0) = txtName) And (.TextMatrix(i, 1) = code) Then
        .TextMatrix(i, 0) = ""
        .TextMatrix(i, 1) = ""
    
        getinfo (Text1)         '��������ã��
        Rhakjum = Rhakjum + info.�������   '�������� �������� ���
        '.Sort = 5
        .Rows = .Rows - 1
        
        
    
    End If

Next i

'ǥ��
Label7 = Format(Rhakjum, " #00 ") & "/" & Format(hakjum, " #00")
List1.RemoveItem (List1.ListIndex) '�����׸� ����

End With




' ��û�������� �߰�

End Sub

Private Sub Command4_Click()

Form2.Show

End Sub

Private Sub Command5_Click()
Form3.Show

End Sub

Private Sub Form_Activate()

If loading = 1 Then
    Exit Sub
End If
'------------------------------------------------------
loading = 1

MsgBox "c:\sugang.txt�� c:\sugangdata.txt������ �о���Դϴ�", , "�˷��帳�ϴ�."
hakjum = 30 '�⺻������ 30���� ������û ����



Dim Filename, Nextline, tmpst As String

Dim Filenum As Integer
Filenum = FreeFile
Filename = "c:\sugang.txt"
Open Filename For Input As Filenum

Dim i, preloc, tmprow, tmpcol As Integer
preloc = 1
tmpcol = 0
tmprow = 1  'col0, 1�����κ��� ����


Do Until EOF(Filenum)

  Line Input #Filenum, Nextline
  Form2.gridSugang.Rows = Form2.gridSugang.Rows + 1
 
For i = 1 To Len(Nextline)

    If Mid(Nextline, i, 1) = " " Then
    
        tmpst = Mid(Nextline, preloc, i - preloc)
        'MsgBox tmpst
        preloc = i + 1
        
        
        Form2.gridSugang.TextMatrix(tmprow, tmpcol) = tmpst
        
        tmpcol = tmpcol + 1
        'MsgBox tmpcol
    End If
Next i
        
tmpst = Mid(Nextline, preloc, i - preloc)
        Form2.gridSugang.TextMatrix(tmprow, tmpcol) = tmpst



tmpcol = 0
tmprow = tmprow + 1
preloc = 1
 
 
'  List1.AddItem (Nextline)
 
Loop
Close Filenum
  
Form2.gridSugang.ColWidth(2) = 3000






Filenum = FreeFile
Filename = "c:\sugangdata.txt"
Open Filename For Input As Filenum


preloc = 1
tmpcol = 0
tmprow = 1  'col0, 1�����κ��� ����


'���Ϸκ��� �о����
Do Until EOF(Filenum)

  Line Input #Filenum, Nextline
  
  For i = 1 To Len(Nextline) - 1
  If Mid(Nextline, i, 1) = " " Then
    tmpst = Trim(Mid(Nextline, 1, i))
'    MsgBox tmpst
    preloc = i + 1
   End If
   Next i
 
    
 
  Form3.gridData.TextMatrix(tmprow, 0) = tmpst
  Form3.gridData.TextMatrix(tmprow, 1) = Mid(Nextline, preloc, i - preloc + 2)

  Form3.gridData.Rows = Form3.gridData.Rows + 1
  tmprow = tmprow + 1
 
Loop
Close Filenum
   
  
Form3.gridData.ColWidth(0) = 3000


End Sub


Private Sub txtName_Change()

'����Ʈ ����
While (List1.ListCount <> 0)
  List1.RemoveItem (0)
Wend



' �ƹ��͵� ������������ ������ϰ� ��������
If txtName = "" Then
    Rhakjum = hakjum
    Label7 = Format(Rhakjum, " #00 ") & "/" & Format(hakjum, " #00")
Exit Sub
End If

'���� �������� ����
Rhakjum = hakjum


Dim i, datacount As Integer
With Form3.gridData

datacount = .Rows - 1

For i = 1 To datacount
    If .TextMatrix(i, 0) = txtName Then
        getinfo (.TextMatrix(i, 1))         '��������ã��
        
        Rhakjum = Rhakjum - info.�������   '�������� �������� ���
        
        List1.AddItem (info.�����)         '����Ʈ�� ǥ��
    End If

Next i

'ǥ��
Label7 = Format(Rhakjum, " #00 ") & "/" & Format(hakjum, " #00")


End With

End Sub
Private Sub getinfo(request) '�����ڵ�� ã��
If request = "" Then
    Exit Sub
End If

Dim i, tmpcnt As Integer
info.���� = False

For i = 1 To Form2.gridSugang.Rows - 1
    If Form2.gridSugang.TextMatrix(i, 1) = request Then
    info.����� = Form2.gridSugang.TextMatrix(i, 2)
    info.�ִ��ο� = Form2.gridSugang.TextMatrix(i, 3)
    info.��û�ο� = Form2.gridSugang.TextMatrix(i, 4)
    info.������� = Form2.gridSugang.TextMatrix(i, 5)
    info.row = i
    
    info.���� = True
    End If
Next

 
End Sub
