VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "클라이언트측 프로그램"
   ClientHeight    =   4770
   ClientLeft      =   1125
   ClientTop       =   1425
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   6840
   Begin VB.TextBox InputLine 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   135
      TabIndex        =   6
      Top             =   4170
      Width           =   5460
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   5265
      TabIndex        =   1
      Top             =   15
      Width           =   1530
      Begin VB.CommandButton 끊기 
         Caption         =   "끊기"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   165
         TabIndex        =   4
         Top             =   1665
         Width           =   1230
      End
      Begin VB.CommandButton 연결 
         Caption         =   "연결"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   165
         TabIndex        =   3
         Top             =   1155
         Width           =   1230
      End
      Begin VB.TextBox AddrBox 
         Height          =   300
         Left            =   90
         TabIndex        =   2
         Top             =   660
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "IP ;"
         Height          =   225
         Left            =   225
         TabIndex        =   5
         Top             =   390
         Width           =   780
      End
   End
   Begin VB.TextBox Text1 
      Height          =   3225
      Left            =   270
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   195
      Width           =   4125
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6150
      Top             =   4140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub 끊기_Click()
    Winsock1.Close
    연결.Enabled = True
    끊기.Enabled = False
End Sub

Private Sub 연결_Click()
    Winsock1.RemoteHost = AddrBox.Text   '주소 지정
    Winsock1.Connect                     '연결 시도
    연결.Enabled = False
    끊기.Enabled = True
End Sub

Private Sub Form_Load()
    Text1.Left = 0
    Text1.Top = 0
    Frame1.Top = 0
    InputLine.Left = 0
    
    Winsock1.RemotePort = 2345  '임의로 정한 포트
End Sub

Private Sub Form_Resize()
    Text1.Width = Form1.ScaleWidth - Frame1.Width
    Text1.Height = Form1.ScaleHeight - InputLine.Height
    Frame1.Left = Text1.Width
    Frame1.Height = Text1.Height
    InputLine.Top = Text1.Height
    InputLine.Width = Form1.ScaleWidth
End Sub

Private Sub InputLine_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And 끊기.Enabled = True Then
        Winsock1.SendData InputLine.Text
        AddText InputLine.Text
        InputLine.Text = ""
    End If
End Sub

Private Sub AddText(str As String)
    Text1.Text = Text1.Text + str + vbNewLine
End Sub

Private Sub Winsock1_Close()
    끊기_Click
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim Gstr As String
    Winsock1.GetData Gstr
    AddText ">" & Gstr
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, _
    ByVal Scode As Long, ByVal Source As String, _
    ByVal HelpFile As String, ByVal HelpContext As Long, _
    CancelDisplay As Boolean)
    MsgBox "에러 발생 :" & Description, vbCritical + vbOKOnly, "에러 발생!"
    Winsock1.Close
    연결.Enabled = True
    끊기.Enabled = False
End Sub
