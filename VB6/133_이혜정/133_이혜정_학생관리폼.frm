VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "133_이혜정_학생관리 폼"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form2"
   ScaleHeight     =   5490
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text6 
      DataField       =   "자격증"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "굴림체"
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
      DataField       =   "주소"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "굴림체"
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
      DataField       =   "이름"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "굴림체"
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
      DataField       =   "학번"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "굴림체"
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
      DataField       =   "핸드폰"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "굴림체"
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
      DataField       =   "전화번호"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "굴림체"
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
      DataField       =   "주민등록번호"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "굴림체"
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
      Caption         =   "닫기"
      Height          =   495
      Left            =   6330
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "성적관리"
      Height          =   495
      Left            =   4830
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "삭제"
      Height          =   495
      Left            =   3330
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "검색"
      Height          =   495
      Left            =   1830
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "입력"
      Height          =   495
      Left            =   330
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\VB6\133_이혜정\133_이혜정_학생관리.mdb"
      DefaultCursorType=   0  '기본 커서
      DefaultType     =   2  'ODBC사용
      Exclusive       =   0   'False
      Height          =   405
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '다이너셋
      RecordSource    =   "Student"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "핸드폰 :"
      BeginProperty Font 
         Name            =   "굴림체"
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
      Caption         =   "자격증 :"
      BeginProperty Font 
         Name            =   "굴림체"
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
      Caption         =   "주  소 :"
      BeginProperty Font 
         Name            =   "굴림체"
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
      Caption         =   "전 화 번 호 :"
      BeginProperty Font 
         Name            =   "굴림체"
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
      Caption         =   "이  름 :"
      BeginProperty Font 
         Name            =   "굴림체"
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
      Caption         =   "주민등록번호 :"
      BeginProperty Font 
         Name            =   "굴림체"
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
      Caption         =   "학  번 :"
      BeginProperty Font 
         Name            =   "굴림체"
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
      Caption         =   "학 생 관 리"
      BeginProperty Font 
         Name            =   "굴림체"
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
    result = InputBox("이름을 입력하세요", "찾기")
    If result = "" Then
        MsgBox "이름이 입력되지 않았습니다.", vbOKOnly + vbCritical, "오류메시지"
        Exit Sub
    End If
    
    Data1.Recordset.FindNext "이름='" & result & "'"
    If Data1.Recordset.NoMatch Then
        Data1.Recordset.FindFirst "이름='" & result & "'"
    End If
    
    If Data1.Recordset.NoMatch Then
        MsgBox "찾는 자료가 없습니다.", vbOKOnly + vbExclamation, "메시지"
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
