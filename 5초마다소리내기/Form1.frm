VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   4020
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command2 
      Caption         =   "시간다시시작"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "소리한번내기"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3120
      Top             =   0
   End
   Begin MCI.MMControl MMControl1 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1296
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "c:\1.wav를 5초마다 재생 "
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MMControl1.FileName = "c:\1.wav"

MMControl1.Command = "Open" '//이것은 컨트롤을 사용가능하게 하는것..
MMControl1.Command = "prev"   '       // 재생을 할것인지 여부
MMControl1.Command = "Play"   '       // 이것은 음악을 연주하는것 임
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
MMControl1.FileName = "c:\1.wav"

MMControl1.Command = "Open" '//이것은 컨트롤을 사용가능하게 하는것..
MMControl1.Command = "prev"   '       // 재생을 할것인지 여부
MMControl1.Command = "Play"   '       // 이것은 음악을 연주하는것 임
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Call Command1_Click

End Sub
