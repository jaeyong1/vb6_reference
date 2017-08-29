VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "c:\out.pla"
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Text            =   "c:\ex1.pla"
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim index() As String   '크기없이 배열생성
Dim sizeOfIndex As Integer



Private Sub Command1_Click()

Dim file_a_line As String '파일 한줄 읽는놈
Dim filename1 As String '파일읽는놈 파일명
Dim filename2 As String '파일쓰는놈 파일명

Dim filenum1 As Integer '읽는놈 이름 (포인터이름이라고 봐도 됨)
Dim filenum2 As Integer '쓰는놈 이름


filename1 = Text1.Text '화면 텍스트박스1에 있는 내용을 파일명으로 받음"
filename2 = Text2.Text '화면 텍스트박스2에 있는 내용을 파일명으로 받음"
filenum1 = FreeFile
filenum2 = FreeFile

Dim 줄바꿈고친스트링 As String


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''파일 읽는거
Open filename1 For Input As filenum1    '읽기로 열기
Do Until EOF(filenum1)                   '계속읽어
   Line Input #filenum1, file_1_line
    줄바꿈고친스트링 = Replace(file_1_line, vbLf, vbCrLf)
Loop '여기까지 Do루프

Close filenum1   '파일닫기
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''파일 읽는거 끝

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''줄바꿈 해결한 파일 쓰기
Dim dmn As Integer '임시파일(줄바꿈된)
dmn = FreeFile

Open "c:\dummy.dat" For Output As dmn
    Print #dmn, 줄바꿈고친스트링;
Close dmn
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''줄바꿈 해결한 파일 쓰기끝

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''줄바꿈 해결한 파일 읽기
Dim isfinish As Boolean
Open "c:\dummy.dat" For Input As filenum1    '읽기로 열기
Do Until EOF(filenum1)                 '계속읽어
   Line Input #filenum1, file_1_line
    func1 (file_1_line)
    'If isfinish = False Then Exit Do
    
Loop '여기까지 Do루프

Close filenum1   '파일닫기
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''줄바꿈 해결한 파일 읽기끝

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''파일 쓰는거

Open filename2 For Output As FreeFile 'append: 뒤에써내려간다.
Print #filenum2, '앞에서 줄바꿈안해서 한줄내림, 아무꺼도 안쓰니깐 줄바꿈..
Print #filenum2, 줄바꿈있기전스트링; '날짜시간기록     ;하면 줄바꿈 안함.
Close #filenum2

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''파일 읽는거 끝


End Sub

Function func1(strln As String) As Boolean



If Mid(strln, 1, 2) = ".i" Then '.i가 나타나면
    Debug.Print ".i발견"
    MsgBox "생성 : " + Trim(Mid(strln, 3, 10)) 'trim:좌우공백문자지움 , val:문자->숫자로 인식(없어도 가능)
    sizeOfIndex = Trim(Mid(strln, 3, 10))
    ReDim index(sizeOfIndex, 10000)   '2차원배열 넉넉히 생성


ElseIf Mid(strln, 1, 2) = ".e" Then '.o가 나타나면
    Debug.Print ".e발견"
    
ElseIf Mid(strln, 1, 2) = ".o" Then '.o가 나타나면
    Debug.Print ".o발견"
ElseIf Mid(strln, 1, 1) = "#" Then
    Debug.Print "# 주석"
    

End If

func1 = True

End Function

