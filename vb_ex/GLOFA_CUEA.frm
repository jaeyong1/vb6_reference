VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "GLOFA PLC CUEA 통신 예제 PROGRAM"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows 기본값
   WindowState     =   2  '최대화
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Write 
         Caption         =   "쓰기"
      End
      Begin VB.Menu Read 
         Caption         =   "읽기"
      End
      Begin VB.Menu Exit 
         Caption         =   "종료"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Exit_Click()
    End
End Sub

Private Sub Read_Click()
    FRM읽기.Show
End Sub

Private Sub Write_Click()
    FRM쓰기.Show
End Sub
