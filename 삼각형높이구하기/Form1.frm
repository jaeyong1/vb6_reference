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
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Text            =   "100"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Text            =   "45"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Text            =   "45"
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "끼인변길이"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "각2"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "각1"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim th1, th2, dis As Double
Dim th3%, a, b, c, h As Double
th1 = Text1
th2 = Text2
a = Text3
th3 = 180 - th1 - th2
b = a / Sin(th3 * 3.14159265358979 / 180) * Sin(th1 * 3.14159265358979 / 180)
h = b * Sin(th2 * 3.14159265358979 / 180)
Label1 = h




End Sub
