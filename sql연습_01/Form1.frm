VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command3 
      Caption         =   "+"
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "->"
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   4080
      TabIndex        =   0
      Top             =   4080
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'프로젝트->참조->Microsoft ActiveX Data Object 2.5 Library
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Form_Load()
cn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=재용;Initial Catalog=pubs;Data Source=222.121.65.9"
cn.CursorLocation = adUseClient
cn.CommandTimeout = 30
cn.Open

End Sub

Private Sub Command1_Click()

'cn.ConnectionString = "Provider=SQLOLEDB;" & _
'"Data Source=top" & _
'"Initial Catalog=erp;" & _
'"uid=erp;" & _
'"Password=erp;" & _
'"Network Address= 192.168.0.2,1433;" & _
'"Trusted_Connection=yes;" & _
'"Network Library=dbmssocn"

'"Network Address=192.168.0.2,1433;" & _ <- 요넘은 서버 아이피 입니다. 서버 아이피를<주석>
'넣으시면 됩니다. <주석>

'Set rs = cn.Execute("SELECT * FROM pubs.dbo.table1") 'erp. < = 디비명, table1 <- 테이블명<주석>
Set rs = cn.Execute("SELECT * FROM test1") 'test1<- 테이블명

rs.MoveFirst

While Not (rs.EOF)
    Print rs(0) & "   " & rs(1) & "   " & rs(2)
    rs.MoveNext
Wend



rs.MoveFirst
'cn.Close

End Sub


Private Sub Command2_Click()


If Not rs.EOF Then
    rs.MoveNext  '<-key point!!!!!!!!
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Text3.Text = rs(2)
'Print rs(0) & "   " & rs(1) & "   " & rs(2)
End If




End Sub

Private Sub Command3_Click()

Dim sql
sql = "INSERT INTO TEST1(text_data, int_data) VALUES('" & _
            Text2 & _
            "' , '" & _
            Text3 & _
            "') "

Print sql

On Error GoTo er

cn.Execute (sql)

Exit Sub

er:
MsgBox "에러발생으로 처리되지 못했음", , "처리에러"
            
End Sub


