VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2040
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click() 'load
'c:\list.txt를 불러옵니다.


Dim NewItem As String
Dim filenum As Integer

filenum = FreeFile
Open "c:\list.txt" For Input As filenum

Do Until EOF(filenum)
    Line Input #filenum, NewItem
    List1.AddItem (NewItem)
Loop

Close filenum

End Sub

Private Sub Command2_Click() 'save
'c:\list.txt로 저장합니다.

Dim inF As Integer
inF = FreeFile
Open "c:\list.txt" For Output As inF
  
For i = 0 To List1.ListCount - 1
    Print #inF, List1.List(i)
Next i

Close inF

End Sub

Private Sub Command3_Click() 'for test
Dim i
List1.AddItem ("a")
List1.AddItem ("b")
List1.AddItem ("c")

Print List1.ListCount '아이템 갯수
For i = 0 To List1.ListCount - 1
    MsgBox List1.List(i)
Next i



End Sub
