VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00404040&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "DECA"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "공지사항"
      ForeColor       =   &H00C0FFC0&
      Height          =   1335
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   4
         Text            =   "Form2.frx":0000
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "감사합니다."
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "감사합니다."
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Bulletin Board Sample"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   525
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Bulletin Board Sample"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   525
      Index           =   1
      Left            =   390
      TabIndex        =   1
      Top             =   390
      Width           =   3810
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim i As Integer
  Label1(0).Visible = False
  Label1(1).Visible = False
  Frame1.Visible = False
  Command1.Visible = False
  Label2(0).Visible = True
  Label2(1).Visible = True
  For i = 1 To 100
    Me.Width = Me.Width - 20
    Me.Height = Me.Height - 20
    Label2(0).Move Me.Width / 2 - 544, Me.Height / 2 - 127
    Label2(1).Move Me.Width / 2 - 547, Me.Height / 2 - 132
    Me.WindowState = 0
  Next i
  Form1.Show
  Unload Me
End Sub
