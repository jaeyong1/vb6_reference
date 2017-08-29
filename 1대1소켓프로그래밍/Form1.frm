VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Frm클라이언트 
   Caption         =   "클라이언트"
   ClientHeight    =   3360
   ClientLeft      =   8505
   ClientTop       =   1455
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   6000
   Begin VB.TextBox Txt전송 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   3255
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "클라이언트작동"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.ListBox Lst상태 
      Height          =   2040
      Left            =   2880
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox TxtPort 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "2000"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox TxtIP 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "IP"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "PORT"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "Frm클라이언트"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdStart_Click()
Winsock1.RemoteHost = TxtIP.Text
Winsock1.RemotePort = TxtPort.Text
Winsock1.Connect
소켓상태알림 1
End Sub

Private Sub Winsock1_Connect()
'Lst상태.AddItem "연결완료"
소켓상태알림 1
End Sub

Private Sub Txt전송_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Winsock1.SendData Txt전송.Text
    Lst상태.AddItem Txt전송
    Txt전송.Text = ""
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData sData, vbString, bytesTotal
Lst상태.AddItem sData
End Sub
