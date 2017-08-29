VERSION 5.00
Begin VB.Form FRMMA 
   Caption         =   "필요마력계산"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form2"
   ScaleHeight     =   4590
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtweight 
      Height          =   270
      Left            =   1680
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtS 
      Height          =   270
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtair 
      Height          =   270
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtV 
      Height          =   270
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "계 산 시 작"
      Default         =   -1  'True
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "무   게(W)  "
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "날개면적(S)"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "공기밀도(p)"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "속   도(V) "
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "N"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "cm"
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "kg/m"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "m/s"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '투명
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '투명
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "FRMMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
