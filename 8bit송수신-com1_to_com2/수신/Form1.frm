VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3480
      Top             =   3960
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1080
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "수신 "
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   8535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()

Dim instring As String

    MSComm1.CommPort = 2            'COM1을 사용.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 10            '입력버퍼 크기를 1Byte로 함.
    MSComm1.PortOpen = True
End Sub

Private Sub Timer1_Timer()
Print ".";    '통신포트 열기.

   instring = MSComm1.Input    '입력버퍼로 부터 한문자를 읽어냄.
   Label1 = Label1 & instring

If instring = "" Then
   Label1 = Label1 & "`"
End If

    
    End Sub
