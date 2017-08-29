VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command2 
      Caption         =   "끝"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   4080
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   4200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "시작"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "올라가는 텍스트 예제 입니다"
      Height          =   180
      Left            =   1380
      TabIndex        =   2
      Top             =   3540
      Width           =   2340
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Timer1.Enabled = True
    Timer1.Interval = 1
End Sub

Private Sub Command2_Click()
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    Label1.Top = Label1.Top - 50
    If Label1.Top <= 0 Then
        Label1.Top = 3540
    End If
End Sub
