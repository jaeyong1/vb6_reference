VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "¼­¹öÃø ÇÁ·Î±×·¥"
   ClientHeight    =   5175
   ClientLeft      =   1125
   ClientTop       =   1440
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   6060
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   5265
      Top             =   2715
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   630
      Left            =   105
      TabIndex        =   1
      Top             =   4275
      Width           =   5790
      Begin VB.TextBox InputLine 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   255
         TabIndex        =   3
         Top             =   225
         Width           =   3585
      End
      Begin VB.CommandButton ²÷±â 
         Caption         =   "²÷±â"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4575
         TabIndex        =   2
         Top             =   210
         Width           =   1020
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  '¼öÁ÷
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   330
      Width           =   4110
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   1
      Left            =   5235
      Top             =   3465
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
Option Explicit

Private Sub ²÷±â_Click()
    ²÷±â.Enabled = False
    Winsock(1).Close
End Sub

Private Sub Form_Load()
    InputLine.Left = 100
    Text1.Left = 0
    Text1.Top = 0
    Frame1.Left = 0

    Winsock(0).Protocol = sckTCPProtocol
    Winsock(0).LocalPort = 2345
    Winsock(0).Listen
End Sub

Private Sub Form_Resize()
    Text1.Width = Form1.ScaleWidth
    Text1.Height = Form1.ScaleHeight - Frame1.Height
    Frame1.Top = Text1.Height
    Frame1.Width = Text1.Width
    InputLine.Width = Frame1.Width - ²÷±â.Width - 300
    ²÷±â.Left = InputLine.Width + 200
End Sub

Private Sub InputLine_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And ²÷±â.Enabled = True Then
        Winsock(1).SendData InputLine.Text
        AddText InputLine.Text
        InputLine.Text = ""
    End If
End Sub

Private Sub AddText(Gstr As String)
    Text1.Text = Text1.Text + Gstr + vbNewLine
End Sub

Private Sub Winsock_Close(Index As Integer)
    ²÷±â_Click
End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    AddText "¿¬°áµÇ¾ú½À´Ï´Ù."
    ²÷±â.Enabled = True

    Winsock(1).Accept requestID
End Sub

Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Gstr As String
    Winsock(1).GetData Gstr
    AddText ">" & Gstr
End Sub

Private Sub Winsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "¿¡·¯ ¹ß»ý :" + Description, vbCritical + vbOKOnly, "¿¡·¯"
    ²÷±â_Click
End Sub
