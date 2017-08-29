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
   StartUpPosition =   3  'Windows 기본값
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
MsgBox ("확인결과 가로: " & Image1.Width & " 세로: " & Image1.Height)

'Stretch : 그림크기를 현재 컨트롤크기에 맞춤
Image1.Stretch = True

End Sub
