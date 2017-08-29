VERSION 5.00
Object = "{7AAC1AE8-58E9-11D3-95BE-EA7F6012FA72}#1.0#0"; "LED.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin LED.LCD LCD1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      Caption         =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    LCD1.Caption = "HANSUNGWOO"
    LCD1.Speed = 200
    LCD1.AutoScroll = True

End Sub
