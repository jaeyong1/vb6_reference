VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0000C000&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4995
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CheckBox Check3 
      BackColor       =   &H0000C000&
      Enabled         =   0   'False
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   1920
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H0000C000&
      Enabled         =   0   'False
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000C000&
      Enabled         =   0   'False
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "닫    기"
      Height          =   375
      Left            =   1680
      MaskColor       =   &H00FFC0FF&
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label txtdial3 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label txtdial2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label txtdial1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "전화번호2 :"
      Height          =   180
      Left            =   480
      TabIndex        =   7
      Top             =   1440
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "전화번호3 :"
      Height          =   180
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "전화번호1 :"
      Height          =   180
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   930
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "<< 설   정 >>"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "모뎀Port :"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "(COM포트 숫자만)"
      Height          =   180
      Left            =   1800
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'modemcom = Text5


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 파일읽기위한 변수
'Dim filename, nextline As String
'Dim filenum As Integer
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ C:\PORT.TXT 읽어서 기억
'filename = "c:\sdwater\PORT.txt"
'filenum = FreeFile
'Open filename For Output As filenum
'Print #filenum, modemcom
'Close #filenum


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
txtdial1 = dial1
txtdial2 = dial2
txtdial3 = dial3
'Text5 = modemcom
Check1.Value = dialcheck1
Check2.Value = dialcheck2
Check3.Value = dialcheck3
End Sub


