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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim 사람() As String  '모듈에서 public으로 해도 됨.


Private Sub Form_Activate()

ReDim 사람(10) As String '10개 할당

Dim i
For i = 1 To 10
  사람(i) = i
  Print 사람(i)
Next i

Erase 사람 '해제

End Sub
