VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "c:\write.txt에 쓰기파일"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "기록"
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txt 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'p340 + p340
'3
Dim filename, linesfromfile, nextline As String
Dim filenum As Integer
filename = "c:\write.txt"
filenum = FreeFile
'4
Open "c:\write.txt" For Append As FreeFile 'append: 뒤에써내려간다.
Print #filenum, '앞에서 줄바꿈안해서 한줄내림
Print #filenum, Date; Time; '날짜시간기록
Print #filenum, txt; '텍스트박스안의내용기록후 줄바꿈안함 => ; 붙임
Close #filenum

txt = ""
End Sub
