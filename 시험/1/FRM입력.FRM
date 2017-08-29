VERSION 5.00
Begin VB.Form frm입력 
   Caption         =   "입력화면"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4590
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmd출력 
      Cancel          =   -1  'True
      Caption         =   "출      력"
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmd입력 
      Caption         =   "입      력"
      Default         =   -1  'True
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox txt포트 
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txt코볼 
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txt어셈 
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txt이름 
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txt번호 
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "포 트 란 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1440
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "코     볼 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1440
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "어셈블리 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "이     름 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1440
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "번     호 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frm입력"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim 번호%, 이름$, 어셈%, 코볼%, 포트%

Private Sub cmd입력_Click()

번호 = Val(Trim(txt번호))
이름 = Trim(txt이름)
어셈 = Val(Trim(txt어셈))
코볼 = Val(Trim(txt코볼))
포트 = Val(Trim(txt포트))

If Not (IsNumeric(txt번호)) Or (번호 > 9) Or (번호 < 0) Then
  MsgBox "번호입력오류"
  txt번호 = ""
  txt번호.SetFocus
  Exit Sub
End If

Select Case 어셈
Case 1 To 100
Case Else
  MsgBox "어셈블리 입력오류"
  txt어셈 = ""
  txt어셈.SetFocus
  Exit Sub
End Select

cnt = cnt + 1
sung(cnt, 1) = 번호
sung(cnt, 2) = 이름
sung(cnt, 3) = 어셈
sung(cnt, 4) = 코볼
sung(cnt, 5) = 포트

txt번호 = ""
txt이름 = ""
txt어셈 = ""
txt코볼 = ""
txt포트 = ""

txt번호.SetFocus

If cnt = 10 Then
  MsgBox "10건의 입력이 완료되었습니다."
  cmd입력.Enabled = False
  cmd출력.SetFocus
End If
End Sub

Private Sub cmd출력_Click()
If cnt = 0 Then
End
End If

Unload Me
frm출력.Show
End Sub
