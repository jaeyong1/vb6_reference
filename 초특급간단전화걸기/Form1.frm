VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1740
      Left            =   1080
      TabIndex        =   0
      Top             =   2400
      Width           =   4410
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   360
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MSComm1.CommPort = 7
    MSComm1.Settings = "38400,N,8,1"
    MSComm1.PortOpen = True
    MSComm1.Output = "ATDT0165356090" + ";" + vbCr
MsgBox "^^"
    MSComm1.Output = "ATH" + vbCr
    
    ' 포트를 닫습니다.
    MSComm1.PortOpen = False
    

End Sub

