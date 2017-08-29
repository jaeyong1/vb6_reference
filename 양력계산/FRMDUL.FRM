VERSION 5.00
Begin VB.Form frmdul 
   Caption         =   "에어포일 둘레계산"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   LinkTopic       =   "Form2"
   ScaleHeight     =   4590
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command4 
      Caption         =   "파일 불러오기"
      Height          =   975
      Left            =   5760
      TabIndex        =   13
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   3960
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1560
      TabIndex        =   5
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      TabIndex        =   4
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "실  행"
      Enabled         =   0   'False
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   4080
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "읽어주세요"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "종  료"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '투명
      Caption         =   "-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '투명
      Caption         =   "총 줄수 :"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '투명
      Height          =   495
      Left            =   1080
      TabIndex        =   15
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblne 
      BackStyle       =   0  '투명
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1560
      TabIndex        =   14
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "두 수치사이 간격수"
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      ToolTipText     =   "누르면 설명이.."
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Caption         =   "x.xx..."
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "소수점 자리수"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      ToolTipText     =   "누르면 설명이.."
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "앞쪽 공백"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      ToolTipText     =   "누르면 설명이.."
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "데이터 시작 줄"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "누르면 설명이.."
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "에어포일 데이터를 가져와 둘레를 계산합니다."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frmdul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
시작.Show
End Sub

Private Sub Command2_Click()
dulhelp.Show 1
End Sub

Private Sub Command3_Click()
Dim Filename, Linesfromfile, Nextline As String
Dim Filenum As Integer
Filename = "c:\airfoil.dat"
Filenum = FreeFile

Open Filename For Input As Filenum
Do Until EOF(Filenum)

  Line Input #Filenum, Nextline
  Linesfromfile = Linesfromfile + Nextline + Chr(13) + Chr(10)
  lblne = Linesfromfile

Loop

Close Filenum




End Sub

Private Sub Command4_Click()
Dim co%
Dim Filename, Linesfromfile, Nextline As String
Dim Filenum As Integer
Filename = "c:\airfoil.dat"
Filenum = FreeFile

Open Filename For Input As Filenum

Do Until EOF(Filenum)
  Line Input #Filenum, Nextline
'  Linesfromfile = Nextline ' + Chr(13) + Chr(10)
  'lblne = Linesfromfile
  co = co + 1
Loop
Label7 = co


Filename = "c:\airfoil.dat"
Filenum = FreeFile

Open Filename For Input As Filenum
Do Until EOF(Filenum)

  Line Input #Filenum, Nextline
  Linesfromfile = Linesfromfile + Nextline + Chr(13) + Chr(10)
  lblne = Linesfromfile

Loop

Close Filenum






'Dim i%
'For i = 1 To co - 1
' 'Line Input #Filenum
' Linesfromfile = Nextline
' lblne = Nextline
'Next i
'lblne = linefromfile

'Close Filenum



Command3.Enabled = True

End Sub

Private Sub Form_Activate()
Dim asdf%
End Sub

Private Sub Label2_Click()
asdf = MsgBox("몇번째 행이 실제 수치데이터가 나오기 시작하는 지를 입력합니다. (위에는 빈칸이나 이름등이 존재하므로) ", , "도움")
End Sub

Private Sub Label3_Click()
asdf = MsgBox("실제 데이터가 나오기 전에 빈칸수를 입력합니다.", , "도움")
End Sub

Private Sub Label4_Click()
asdf = MsgBox("소수 몇째자리까지 내려가는지를 입력합니다.(예>0.1122는 소수4자리) ", , "도움")
End Sub

Private Sub Label6_Click()
asdf = MsgBox("X값과 Y값 사이에 존재하는 빈칸수를 입력합니다", , "도움")
End Sub
