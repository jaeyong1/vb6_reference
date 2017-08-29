VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   3045
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "칸속의 루트값은?"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "결과"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "입력"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '단일 고정
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, b As Double

a = Text1.Text
b = Sqr(a)
Label1.Caption = b

End Sub
