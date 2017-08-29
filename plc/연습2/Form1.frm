VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   2280
      Top             =   2160
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   8175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   735
      Left            =   3840
      TabIndex        =   1
      Top             =   1920
      Width           =   855
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1080
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '단일 고정
      Caption         =   "...->"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1 = ""
q = Chr(5) & "00RSS0106%PW000" & Chr(4)
On Error GoTo errmsg
  MSComm1.Output = q
  Timer1.Enabled = True
  
  
  
  'Do
    ' DoEvents
     instring = MSComm1.Input
    ' Rcv = Rcv & instring
  'Loop Until instring = Chr(3)
  
  
  'Text1 = Text1 & Rcv
  'Rcv = ""
 ' MSComm1.PortOpen = False
  Exit Sub
  
errmsg:
 dummy = MsgBox("포트못열어", 0, "에러~")

  
End Sub

Private Sub Form_Activate()
  MSComm1.CommPort = 2
  MSComm1.Settings = "19200,n,8,1"
  MSComm1.InputLen = 3
  MSComm1.PortOpen = True

End Sub

Private Sub Form_Click()
End
End Sub

Private Sub Timer1_Timer()


   instring = MSComm1.Input    '입력버퍼로 부터 한문자를 읽어냄.
   Text1 = Text1 & instring

If instring = "" Then
   Text1 = Text1 & Chr(9)
      End If
End Sub

