VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   1200
      Top             =   1080
      Width           =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Set Image1 = LoadPicture("c:\a.bmp")
MsgBox ("Ȯ�ΰ�� ����: " & Image1.Width & " ����: " & Image1.Height)

'Stretch : �׸�ũ�⸦ ���� ��Ʈ��ũ�⿡ ����
Image1.Stretch = True

End Sub
