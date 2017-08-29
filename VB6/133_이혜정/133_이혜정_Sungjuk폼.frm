VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "133_이혜정_SungJuk 폼"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form3"
   ScaleHeight     =   5490
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command6 
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
      Left            =   3600
      TabIndex        =   16
      Top             =   3690
      Width           =   975
   End
   Begin VB.CommandButton Command3 
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
      Left            =   3600
      TabIndex        =   15
      Top             =   2805
      Width           =   975
   End
   Begin VB.TextBox Text4 
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
      Left            =   4800
      TabIndex        =   14
      Top             =   2865
      Width           =   1575
   End
   Begin VB.TextBox Text2 
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
      Left            =   4800
      TabIndex        =   13
      Top             =   3750
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
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "입력"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "검색"
      Height          =   495
      Left            =   2360
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "학생관리"
      Height          =   495
      Left            =   4360
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "닫기"
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
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
      Left            =   1560
      TabIndex        =   3
      Top             =   1110
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
      Left            =   1560
      TabIndex        =   2
      Top             =   1995
      Width           =   1575
   End
   Begin VB.TextBox Text5 
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
      Left            =   1560
      TabIndex        =   1
      Top             =   2865
      Width           =   1575
   End
   Begin VB.TextBox Text6 
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
      Left            =   1560
      TabIndex        =   0
      Top             =   3750
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "성 적 관 리"
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
      TabIndex        =   12
      Top             =   240
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
      TabIndex        =   11
      Top             =   1200
      Width           =   840
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
      TabIndex        =   10
      Top             =   2085
      Width           =   840
   End
   Begin VB.Label Label6 
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
      Left            =   390
      TabIndex        =   9
      Top             =   2955
      Width           =   1050
   End
   Begin VB.Label Label7 
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
