VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "GLOFA PLC CUEA ��� ���� PROGRAM"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows �⺻��
   WindowState     =   2  '�ִ�ȭ
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Write 
         Caption         =   "����"
      End
      Begin VB.Menu Read 
         Caption         =   "�б�"
      End
      Begin VB.Menu Exit 
         Caption         =   "����"
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
    FRM�б�.Show
End Sub

Private Sub Write_Click()
    FRM����.Show
End Sub
