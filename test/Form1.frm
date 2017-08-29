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
   StartUpPosition =   3  'Windows ±âº»°ª
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()


Dim inputdata As String
inputdata = "2|0|"


'Print changeitem(inputdata, 2, getnextitem(inputdata, 2) + 1)

Call IncreaseNumber(inputdata)
Call addItem(inputdata, "ew")


Print "r : " & inputdata


'Print getnextitem(inputdata, 1)
'changeitem inputdata, 3, "±è¸»¶Ë"


'If getnextitem(inputdata, 1) = "" Then MsgBox "¾øÀ½"


'IncreaseNumber (inputdata)





End Sub

