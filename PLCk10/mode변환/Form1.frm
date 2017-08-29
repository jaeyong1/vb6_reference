VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3720
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Master-K 10S Run or Stop"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   3240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim MODE As String
MODE = "01"

With MSComm1
 .CommPort = 2
 .Settings = "9600,N,8,1"
 .PortOpen = True
 .Output = Chr$(5)
 .Output = Chr$(2) + "M" + MODE + Chr(4)
 .PortOpen = False
End With
End Sub

Private Sub Command2_Click()
Dim MODE As String
MODE = "02"

With MSComm1
 .CommPort = 2
 .Settings = "9600,N,8,1"
 .PortOpen = True
 .Output = Chr$(5)
 .Output = Chr$(2) + "M" + MODE + Chr(4)
 .PortOpen = False
End With

End Sub

