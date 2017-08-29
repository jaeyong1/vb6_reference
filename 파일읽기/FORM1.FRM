VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "c:\vbtxt.txt라는 파일올림"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label1"
      Height          =   1935
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Filename, Linesfromfile, Nextline As String
Dim Filenum As Integer
Filename = "c:\vbtxt.txt"
Filenum = FreeFile

Open Filename For Input As Filenum
Do Until EOF(Filenum)

  Line Input #Filenum, Nextline
  Linesfromfile = Linesfromfile + Nextline + Chr(13) + Chr(10)
  Label1 = Linesfromfile
Loop
Close Filenum
Debug.Print Linesfromfile

End Sub
