VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0000C000&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5100
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CheckBox Check3 
      BackColor       =   &H0000C000&
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H0000C000&
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000C000&
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1800
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   1680
      TabIndex        =   15
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1680
      TabIndex        =   13
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1680
      TabIndex        =   11
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1680
      TabIndex        =   9
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox txtdial3 
      Height          =   270
      Left            =   1800
      TabIndex        =   7
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtdial2 
      Height          =   270
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtdial1 
      Height          =   270
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "닫    기"
      Height          =   375
      Left            =   2880
      MaskColor       =   &H00FFC0FF&
      TabIndex        =   1
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "저    장"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      MaskColor       =   &H00FFC0FF&
      TabIndex        =   0
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "<< 남  김  말 >>"
      Height          =   180
      Left            =   0
      TabIndex        =   24
      Top             =   2760
      Width           =   4980
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '투명
      Caption         =   "SMS사용"
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "(COM포트 숫자만)"
      Height          =   180
      Left            =   2280
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "모뎀Port :"
      Height          =   180
      Left            =   840
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "<< 설   정 >>"
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "휴 대 폰 1 :"
      Height          =   180
      Left            =   600
      TabIndex        =   16
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label Label8 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "염색폐수수위 :"
      Height          =   180
      Left            =   360
      TabIndex        =   14
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Label Label7 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "집수탱크수위 :"
      Height          =   180
      Left            =   360
      TabIndex        =   12
      Top             =   4080
      Width           =   1200
   End
   Begin VB.Label Label6 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "저장탱크수위 :"
      Height          =   180
      Left            =   360
      TabIndex        =   10
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Label Label5 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "우 물 수 위 :"
      Height          =   180
      Left            =   480
      TabIndex        =   8
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Label Label3 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "휴 대 폰 3 :"
      Height          =   180
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Width           =   1050
   End
   Begin VB.Label Label2 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "휴 대 폰 2 :"
      Height          =   180
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   1050
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
dial1 = txtdial1
dial2 = txtdial2
dial3 = txtdial3
set1 = Text1
set2 = Text2
set3 = Text3
set4 = Text4
modemcom = Text5
dialcheck1 = Check1.Value
dialcheck2 = Check2.Value
dialcheck3 = Check3.Value
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
txtdial1 = dial1
txtdial2 = dial2
txtdial3 = dial3
Text1 = set1
Text2 = set2
Text3 = set3
Text4 = set4
Text5 = modemcom
Check1.Value = dialcheck1
Check2.Value = dialcheck2
Check3.Value = dialcheck3
End Sub

