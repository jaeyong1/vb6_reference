VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdend 
      Cancel          =   -1  'True
      Caption         =   "뒈지라요~"
      Height          =   975
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   2160
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdend_Click()
End
End Sub
'c:\bgm\cleo.wav 를 연주해요
Private Sub form_load()
MMControl1.FileName = "c:\bgm\cleo.wav"
MMControl1.Command = "Open" '//이것은 컨트롤을 사용가능하게 하는것..
MMControl1.Command = "prev"   '       // 재생을 할것인지 여부
MMControl1.Command = "Play"   '       // 이것은 음악을 연주하는것 임
'-----------------------------------------------------------------
End Sub

Function MyPath()
   MyPath = App.Path
   If Right(MyPath, 1) <> "\" Then
   MyPath = MyPath & "\"
    End If
End Function


