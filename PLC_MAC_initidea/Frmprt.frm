VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frmprt 
   Caption         =   "�μ� ������"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8955
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   518
   ScaleMode       =   3  '�ȼ�
   ScaleWidth      =   597
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CheckBox chkPnt4 
      Caption         =   "��������"
      Height          =   255
      Left            =   2400
      TabIndex        =   28
      Top             =   7320
      Value           =   1  'Ȯ��
      Width           =   1695
   End
   Begin VB.CheckBox chkPnt3 
      Caption         =   "�򰡰��"
      Height          =   255
      Left            =   2400
      TabIndex        =   27
      Top             =   6960
      Value           =   1  'Ȯ��
      Width           =   1455
   End
   Begin VB.CheckBox chkPnt2 
      Caption         =   "������"
      Height          =   255
      Left            =   600
      TabIndex        =   26
      Top             =   7320
      Value           =   1  'Ȯ��
      Width           =   1215
   End
   Begin VB.CheckBox chkPnt1 
      Caption         =   "ǥ��"
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   6960
      UseMaskColor    =   -1  'True
      Value           =   1  'Ȯ��
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   5775
      Left            =   4800
      TabIndex        =   23
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   22
      Top             =   6960
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker DTReqDay 
      Height          =   375
      Left            =   1680
      TabIndex        =   19
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      Format          =   25493505
      CurrentDate     =   40016
   End
   Begin VB.TextBox Txtprog 
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   18
      Top             =   6240
      Width           =   2775
   End
   Begin VB.TextBox TxtModelnum 
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Text            =   "PLC-AAA-BB001"
      Top             =   5640
      Width           =   2775
   End
   Begin VB.TextBox TxtMaker 
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Text            =   "ȫ�浿"
      Top             =   5040
      Width           =   2775
   End
   Begin VB.TextBox TxtReqA 
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Text            =   "�ѱ��������(��)"
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox TxtClk 
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Text            =   "������"
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox TxtUDay 
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Text            =   "�����Ϸκ��� 6����"
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox TxtAsso 
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Text            =   "�ѱ����⿬����"
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox TxtSpecType 
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Text            =   "KS X 4600-1 (Class-B)"
      Top             =   240
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTFinDay 
      Height          =   375
      Left            =   1680
      TabIndex        =   20
      Top             =   1440
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      Format          =   25493505
      CurrentDate     =   40016
   End
   Begin MSComCtl2.DTPicker DTPrtDay 
      Height          =   375
      Left            =   1680
      TabIndex        =   21
      Top             =   2040
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      Format          =   25493505
      CurrentDate     =   40016
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6360
      TabIndex        =   24
      Top             =   240
      Width           =   525
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   10
      Top             =   6240
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "DUT �𵨸�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   9
      Top             =   5640
      Width           =   1170
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "DUT ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   1170
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "�����û���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   1350
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "��ȿ�Ⱓ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "�� �� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "����Ϸ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�� û ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���� �԰�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmprt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlApp As New Excel.Application '������Ʈ�޴�-���� - Microsoft Excel 12.0 Object Library  (2007 ����, ���Ϲ����� 11 10..�� ����)

Private Sub CheckBox1_Click()

End Sub

Private Sub CheckBox2_Click()

End Sub

Private Sub Command1_Click()
  Const XL_NOTRUNNING As Long = 429 '������ ������ ����ǰ� ���� ������ 429 ������ �߻�
  On Error GoTo ShowName_Err '������ �߻��ϸ�(������ ������ ����ǰ� ���� �ʴٸ�) ShowName_Err ������ �̵�
  Set xlApp = GetObject(, "Excel.Application") '������ ����ǰ� �ֳ� üũ
    
    
        
  With xlApp  '���� �� �۾�
        .Visible = True '���� ǥ��
        '.Visible = False '���� ǥ�� ����
        
        .DisplayAlerts = False '��� ����
        .Workbooks.Open App.Path & "\cert.xlsx"  '������� ���� ����
    
    If chkPnt1.Value = 1 Then
        '#### ǥ�� �μ� ####
        .Sheets("ǥ��").Select
        .Range("e7").Select: .ActiveCell.FormulaR1C1 = "0000-0000-0000" ' ������ȣ
        .Range("e9").Select: .ActiveCell.FormulaR1C1 = TxtSpecType.Text  ' ����԰�
        
        .Range("e14").Select: .ActiveCell.FormulaR1C1 = TxtReqA.Text  ' �����û���
        .Range("e15").Select: .ActiveCell.FormulaR1C1 = TxtMaker.Text  ' DUT������
        .Range("e16").Select: .ActiveCell.FormulaR1C1 = TxtModelnum.Text  ' DUT�𵨸�
        
        .Range("d19").Select: .ActiveCell.FormulaR1C1 = DTReqDay.Value  ' ��û��
        .Range("d20").Select: .ActiveCell.FormulaR1C1 = DTFinDay.Value  ' ����Ϸ���
        .Range("d21").Select: .ActiveCell.FormulaR1C1 = DTPrtDay.Value  ' ������
        
        .Range("f19").Select: .ActiveCell.FormulaR1C1 = TxtAsso.Text  ' ������
        .Range("f20").Select: .ActiveCell.FormulaR1C1 = TxtClk.Text  ' ������
        .Range("f21").Select: .ActiveCell.FormulaR1C1 = TxtUDay.Text  ' ��ȿ�Ⱓ
        
        .ActiveWindow.SelectedSheets.PrintOut Copies:=1 '�μ�
    End If
    
    If chkPnt2.Value = 1 Then
        '#### ������ �μ� ####
         .Sheets("������").Select
    
    
    
    End If
    
    
    If chkPnt3.Value = 1 Then
        '#### �򰡰�� �μ� ####
         .Sheets("�򰡰��").Select
    
    
    
    End If
    
    
    If chkPnt4.Value = 1 Then
        '#### �������� �μ� ####
         .Sheets("��������").Select
    
    
    
    End If
        
  End With
    
    
'
'    '�÷��� �׸����� �ڷḦ ������ �����Ѵ�.
'    For iRow = 0 To VSFlexGrid1.Rows - 1
'        For iCol = 0 To VSFlexGrid1.Cols - 1
'            oExcel.Worksheets(1).Cells(iRow, iCol).Value = MSFlexGrid1.TextMatrix(iRow, iCol)
'        Next
'    Next
'
'    '���� ���Ϸ� �����Ѵ�.
'    oExcel.Worksheets(1).SaveAs "C:\test.xls"
'    'sPath = "http://" & window.location.host & "\eMES\reports\prodt\prd600p.xls"
'    'oExcel.Worksheets(1).SaveAs "F:\test.xls"
'    '��ȭ�� ���� ��ȯ�մϴ�.
  '  oExcel.Interactive = True
'
'
'
'    With xlApp
'
'
'    .Range("C3").Select :     .ActiveCell.FormulaR1C1 = "1"
'    .Range("C4").Select
'    .ActiveCell.FormulaR1C1 = "2"
'    .Range("C5").Select
'    .ActiveCell.FormulaR1C1 = "3"
'    .Range("C6").Select
'    .ActiveCell.FormulaR1C1 = "4"
'    .Range("C7").Select
'    .ActiveWindow.SelectedSheets.PrintOut Copies:=1
'
'
'    '   ������ �ٷ� ���
'End With
'    '���� ��ü�� �ݽ��ϴ�.
'    If Not (oExcel Is Nothing) Then
'        Set oExcel = Nothing
'    End If
''
'      '�÷��� �׸����� �ڷḦ ������ �����Ѵ�.
'    For iRow = 0 To VSFlexGrid1.Rows - 1
'        For iCol = 0 To VSFlexGrid1.Cols - 1
'            oExcel.Worksheets(1).Cells(iRow + 1, iCol + 1).Value = VSFlexGrid1.TextMatrix(iRow, iCol)
'        Next
'    Next
    
    
    
    
          
    


'xlApp.Quit '���� ���α׷� ����
'Set xlApp = Nothing '���� ���ø����̼� ��ü �޸𸮿��� ����
Exit Sub

''''''''''' ����ó��
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

Private Sub Form_Unload(Cancel As Integer)
 '
End Sub
