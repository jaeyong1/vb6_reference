VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "전체화면 폼"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form4"
   ScaleHeight     =   5490
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command7 
      Caption         =   "평점"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3570
      TabIndex        =   26
      Top             =   3765
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "평균"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3570
      TabIndex        =   25
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text11 
      DataField       =   "평균"
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
      Left            =   4770
      TabIndex        =   24
      Top             =   3180
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      DataField       =   "평점"
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
      Left            =   4770
      TabIndex        =   23
      Top             =   3825
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      DataField       =   "중간고사"
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
      Left            =   1530
      TabIndex        =   20
      Top             =   3180
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      DataField       =   "기말고사"
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
      Left            =   1530
      TabIndex        =   19
      Top             =   3825
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\VB6\133_이혜정\133_이혜정_학생관리.mdb"
      DefaultCursorType=   0  '기본 커서
      DefaultType     =   2  'ODBC사용
      Exclusive       =   0   'False
      Height          =   405
      Left            =   390
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '다이너셋
      RecordSource    =   "Student"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "입력"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "검색"
      Height          =   495
      Left            =   2360
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "삭제"
      Height          =   495
      Left            =   4360
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "닫기"
      Height          =   495
      Left            =   6360
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
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
      Left            =   5310
      TabIndex        =   6
      Top             =   990
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
      Left            =   5310
      TabIndex        =   5
      Top             =   1515
      Width           =   2175
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
      Left            =   5310
      TabIndex        =   4
      Top             =   2550
      Width           =   2175
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
      Left            =   1350
      TabIndex        =   3
      Top             =   990
      Width           =   1575
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
      Left            =   1350
      TabIndex        =   2
      Top             =   1515
      Width           =   1575
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
      Left            =   1350
      TabIndex        =   1
      Top             =   2025
      Width           =   6135
   End
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
      Left            =   1350
      TabIndex        =   0
      Top             =   2550
      Width           =   2895
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "중간고사 :"
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
      TabIndex        =   22
      Top             =   3270
      Width           =   1050
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "기말고사 :"
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
      TabIndex        =   21
      Top             =   3915
      Width           =   1050
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
      Left            =   2925
      TabIndex        =   18
      Top             =   120
      Width           =   2085
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
      Left            =   390
      TabIndex        =   17
      Top             =   1080
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
      Left            =   3750
      TabIndex        =   16
      Top             =   1080
      Width           =   1470
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
      Left            =   390
      TabIndex        =   15
      Top             =   1605
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
      Left            =   3750
      TabIndex        =   14
      Top             =   1605
      Width           =   1365
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
      Left            =   390
      TabIndex        =   13
      Top             =   2115
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
      Left            =   390
      TabIndex        =   12
      Top             =   2640
      Width           =   840
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
      Left            =   4350
      TabIndex        =   11
      Top             =   2640
      Width           =   840
   End
End
Attribute VB_Name = "Form4"
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

Private Sub Command5_Click()
    Form4.Hide
    Form1.Show
End Sub

Private Sub Command6_Click()
    Dim num1, num2
    num1 = Val(Text9.Text)
    num2 = Val(Text8.Text)
    Text11.Text = Str((num1 + num2) / 2)
End Sub

Private Sub Command7_Click()
    If Val(Text11.Text) >= 90 Then
        Text10.Text = "A"
    ElseIf Val(Text11.Text) >= 80 Then
        Text10.Text = "B"
    ElseIf Val(Text11.Text) >= 70 Then
        Text10.Text = "C"
    ElseIf Val(Text11.Text) >= 60 Then
        Text10.Text = "D"
    Else
        Text10.Text = "F"
    End If
End Sub
