VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12720
   ClientLeft      =   240
   ClientTop       =   795
   ClientWidth     =   15975
   LinkTopic       =   "Form1"
   ScaleHeight     =   848
   ScaleMode       =   3  '�ȼ�
   ScaleWidth      =   1065
   Begin VB.CommandButton cmdprt 
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
      Height          =   615
      Left            =   120
      TabIndex        =   41
      Top             =   11040
      Width           =   3015
   End
   Begin VB.CommandButton cmdtest 
      Caption         =   "Test"
      Height          =   495
      Left            =   5640
      TabIndex        =   40
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox TxtRunResult 
      Height          =   11295
      Left            =   9960
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   36
      Top             =   600
      Width           =   5895
   End
   Begin VB.TextBox TxtRunState 
      Height          =   5655
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   35
      Top             =   6240
      Width           =   6255
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4935
      Left            =   3480
      TabIndex        =   34
      Top             =   600
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8705
      _Version        =   393217
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.CommandButton Command6 
      Caption         =   "����"
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
      Left            =   1680
      TabIndex        =   33
      Top             =   10320
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "����"
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
      Left            =   120
      TabIndex        =   32
      Top             =   10320
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
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
      Height          =   615
      Left            =   120
      TabIndex        =   31
      Top             =   11760
      Width           =   3015
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  '������ ����
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
      Left            =   120
      TabIndex        =   29
      Text            =   "60"
      Top             =   9720
      Width           =   3015
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  '������ ����
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
      Left            =   120
      TabIndex        =   27
      Text            =   "BB00000000000000"
      Top             =   9000
      Width           =   3015
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  '������ ����
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
      Left            =   120
      TabIndex        =   25
      Text            =   "AA00000000000000"
      Top             =   8280
      Width           =   3015
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  '������ ����
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
      Left            =   120
      TabIndex        =   23
      Text            =   "220000000000"
      Top             =   7560
      Width           =   3015
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  '������ ����
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
      Left            =   120
      TabIndex        =   21
      Text            =   "110000000000"
      Top             =   6840
      Width           =   3015
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  '������ ����
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
      Left            =   120
      TabIndex        =   19
      Text            =   "20"
      Top             =   6120
      Width           =   3015
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  '������ ����
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
      Left            =   120
      TabIndex        =   17
      Text            =   "5"
      Top             =   5400
      Width           =   3015
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  '������ ����
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
      Left            =   120
      TabIndex        =   15
      Text            =   "6000"
      Top             =   4680
      Width           =   3015
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  '������ ����
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
      Left            =   120
      TabIndex        =   14
      Text            =   "6000"
      Top             =   3960
      Width           =   3015
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  '��� ����
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
      Left            =   2520
      TabIndex        =   12
      Text            =   "2"
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  '��� ����
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
      Text            =   "0"
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  '��� ����
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
      Left            =   840
      TabIndex        =   10
      Text            =   "0"
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  '��� ����
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
      Left            =   120
      TabIndex        =   9
      Text            =   "10"
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '��� ����
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
      Left            =   2520
      TabIndex        =   7
      Text            =   "4"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '��� ����
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
      TabIndex        =   6
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '��� ����
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
      Left            =   840
      TabIndex        =   5
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '��� ����
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
      Left            =   120
      TabIndex        =   4
      Text            =   "10"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "������ �׽�Ʈ"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RTS/CTS �׽�Ʈ"
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
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Ϲ� �׽�Ʈ"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�׽�Ʈ ���� ���"
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
      Left            =   9960
      TabIndex        =   39
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�׽�Ʈ ���� ����"
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
      Left            =   3480
      TabIndex        =   38
      Top             =   5880
      Width           =   1725
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�׽�Ʈ ���̽� ����"
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
      Left            =   3480
      TabIndex        =   37
      Top             =   240
      Width           =   1950
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "TIME OUT"
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
      Left            =   120
      TabIndex        =   30
      Top             =   9480
      Width           =   1005
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "2nd ENCRYPTION KEY"
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
      Left            =   120
      TabIndex        =   28
      Top             =   8760
      Width           =   2205
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1st ENCRYPTION KEY"
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
      Left            =   120
      TabIndex        =   26
      Top             =   8040
      Width           =   2145
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "2nd Group ID"
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
      Left            =   120
      TabIndex        =   24
      Top             =   7320
      Width           =   1305
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1st Group ID"
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
      Left            =   120
      TabIndex        =   22
      Top             =   6600
      Width           =   1245
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "EXTENDED TEST TIME"
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
      Left            =   120
      TabIndex        =   20
      Top             =   5880
      Width           =   2250
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "TEST TIME"
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
      Left            =   120
      TabIndex        =   18
      Top             =   5160
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "LOCAL PORT NUMBER"
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
      Left            =   120
      TabIndex        =   16
      Top             =   4440
      Width           =   2355
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "REMOTE IP ADDRESS"
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
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   2265
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "TEST IP ADDRESS"
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
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "REMOTE IP ADDRESS"
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
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2265
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprt_Click()
'####### ��� ��ư #######
frmprt.Show 1

End Sub

Private Sub cmdtest_Click()
'test button
Add_Result ("aaaa")

End Sub

Public Sub Add_Result(str As String)
'###### �׽�Ʈ ��� �ؽ�Ʈ �߰� ######
    TxtRunResult.Text = TxtRunResult.Text + str
    TxtRunResult.SelStart = Len(TxtRunResult.Text)
End Sub

Public Sub Add_State(str As String)
'###### �׽�Ʈ ���� ���� �ؽ�Ʈ �߰� ######
    TxtRunState.Text = TxtRunState.Text + vbCrLf + str
    TxtRunState.SelStart = Len(TxtRunState.Text)
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command5_Click()
' ##### ���۹�ư ######
Dim nTest As Integer
nTest = TreeView1.Nodes.Count
Debug.Print "test start / num of test : " & nTest
 
Erase TestNode  '�����ڷ� ����
ReDim TestNode(nTest) As PLCTestNode '�׸� ������ŭ ����
 
Dim s
For NowTreeIndex = 1 To nTest
    If TreeView1.Nodes(NowTreeIndex).Checked = True Then
    '## �� �׸񺰷� üũ Ȯ���� ���� ##
        Add_State ("<" & TreeView1.Nodes(NowTreeIndex).Text & ">")    'ȭ�����
        Add_Result ("* " & TreeView1.Nodes(NowTreeIndex).Text & " : ")     'ȭ�����
        s = TreeView1.Nodes(NowTreeIndex).Key
        TestSpec (s)    '�׽�Ʈ ��ü ȣ��
        
        Add_State (vbCrLf + "-----------------" + vbCrLf + vbCrLf)
        Add_Result (vbCrLf)
        
        
     '+ vbCrLf + "-----------------" + vbCrLf + vbCrLf) 'ȭ�����

    End If
Next NowTreeIndex


 
End Sub

Private Sub Form_Load()
'�� �ҷ�����

'####### �׽�Ʈ ���̽� Ʈ�� ��� �Է� ########
Dim nod_x As Node
Set nod_x = TreeView1.Nodes.Add(, , "GTC", "General Test Cases")   '�ε������� 1
    Set nod_x = TreeView1.Nodes.Add("GTC", tvwChild, "CFF", "Conrol Frame ")
        Set nod_x = TreeView1.Nodes.Add("CFF", tvwChild, "1_1", "1.1 DT field of Control Frame")
        Set nod_x = TreeView1.Nodes.Add("CFF", tvwChild, "1_2", "1.2 VF field of Unicast Data Frame")
        Set nod_x = TreeView1.Nodes.Add("CFF", tvwChild, "1_3", "1.3 DT field of Management Frame")
                Set nod_x = TreeView1.Nodes.Add("CFF", tvwChild, "1_4", "1.4 DT field of Broadcast Data Frame")
    Set nod_x = TreeView1.Nodes.Add("GTC", tvwChild, "DFF", "Data Frame Format")
        Set nod_x = TreeView1.Nodes.Add("DFF", tvwChild, "2_1", "2.1 AAAAA")
        Set nod_x = TreeView1.Nodes.Add("DFF", tvwChild, "2_2", "2.2 BBBBB")
    Set nod_x = TreeView1.Nodes.Add("GTC", tvwChild, "IFS", "IFS(Inter-Frame Space")
    Set nod_x = TreeView1.Nodes.Add("GTC", tvwChild, "CE", "CE(Channel Estimation")


TreeView1.Nodes.Item(1).Expanded = True 'Root���� Ȯ��
'Debug.Print TreeView1.Nodes.Count '��� ����
Debug.Print vbCrLf & vbCrLf & vbCrLf

End Sub
 

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
'Ʈ�� üũ

'Debug.Print "me:" & Node.Index

'########  Rootüũ�� ����/���� �� ���  ########
If Node.Index = 1 Then
    For q = 1 To TreeView1.Nodes.Count
         TreeView1.Nodes(q).Checked = Node.Checked
    Next q
    Exit Sub
End If


'########  üũǥ�� ���� �� ���  ########
If (Node.Checked = True) And (Node.Index <> 1) Then
    Debug.Print "pa:" & Node.Parent.Index
    
    If Node.Parent.Checked = False Then    '�ڽ� üũ�ϸ� �θ� üũ�ǰ�..
        Node.Parent.Checked = True
    End If
    

    Debug.Print "node.Children" & Node.Children
    For q = Node.Index To (Node.Index + Node.Children)  '�θ� üũ�ϸ� �ڽĵ� üũ�ǰ�.
        TreeView1.Nodes(q).Checked = True
    Next q
    Exit Sub
End If


'########  üũǥ�� ���� �� ���  ########
If (Node.Checked = False) And (Node.Index <> 1) Then
    Debug.Print "pa:" & Node.Parent.Index

    Debug.Print "node.Children" & Node.Children
    For q = Node.Index To (Node.Index + Node.Children)  '�θ� üũ�ϸ� �ڽĵ� �����ǰ�.
        TreeView1.Nodes(q).Checked = False
    Next q
    Exit Sub
End If


End Sub

