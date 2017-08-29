VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'배열에 저장된 값을 파일에 가장 빠르게 저장하는 방법

'배열에 저장된 값을 파일에 가장 빠르게 저장하는 방법
'
'배열에 저장된 값을 파일에 저장하거나 읽을 경우
'
'배열의 값을 순차적으로 저장하게 되면 저장 속도가
'
'느리게 됩니다.
'
'이럴경우는 바이나리 파일을 이용 하여 한꺼번에 저장하거나
'
'읽으면 순식간에 처리가 됩니다.
'
'아래에 10000 건의 배열이 있다고 가정하고
'
'그것을 저장하는 예제가 있습니다.
Private Sub Form_Activate()
   Dim arr(1 To 100000) As Long
   Dim fnum As Integer

       fnum = FreeFile
       Open "C:\Temp\xxx.dat" For Binary As fnum
       Put #fnum, , arr
       Close fnum
End Sub
'끝.
 


