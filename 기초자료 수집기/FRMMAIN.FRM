VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "이쯤이야"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "이게뭐꼬? 예전처럼 해도!!!!"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   4680
      Width           =   5295
   End
   Begin VB.CommandButton end 
      Caption         =   "종    료"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   3
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton rcvr 
      Caption         =   "마 무 리"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   2
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   $"frmmain.frx":0000
      Height          =   1215
      Left            =   960
      TabIndex        =   5
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "*** 이쯤이야 ***"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   1200
      Width           =   2835
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Made in Korea 재용 엔터테인먼트"
      Height          =   180
      Left            =   3960
      TabIndex        =   1
      Top             =   0
      Width           =   2805
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '단일 고정
      Height          =   2535
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   5775
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub end_Click()
End
End Sub

Private Sub rcvr_Click()
retval = ("a:\comeback.bat")

End Sub
