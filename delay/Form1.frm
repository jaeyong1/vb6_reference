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
   StartUpPosition =   3  'Windows ±‚∫ª∞™
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
'delay(1000)¿∫ 1√ ∞£ ∏ÿ√·¥Ÿ.
End Sub

 Public Sub Delay(Num As Single)

    Dim St As Single

    St = Timer

    Do

      DoEvents

    Loop While (Timer - St) < (Num / 1000)

End Sub
