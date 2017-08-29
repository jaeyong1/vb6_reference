VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form 접속도우미 
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   120
   ClientLeft      =   -645
   ClientTop       =   -825
   ClientWidth     =   120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   120
   ScaleWidth      =   120
   ShowInTaskbar   =   0   'False
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   480
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   0
      Top             =   240
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "접속도우미"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer


Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
'Text1.Text = Inet1.OpenURL("http://163.180.114.88/jy/jy_ip_save.php")
End Sub

Private Sub Form_Load()
'창크기 아주 작음
Me.Height = 1
Me.Width = 1

End Sub

Private Sub Timer1_Timer()
cnt = cnt + 1
On Error GoTo er

' 프로그램 실행후 아래 분(minute)에 접속 시도..

If (cnt = 3 Or cnt = 5 Or cnt = 6) Then
    Text1.Text = Inet1.OpenURL("http://cvs.khu.ac.kr/~jaeyong1/ip/jy_ip_save.php")
End If

If cnt = 7 Then
    Text1.Text = Inet1.OpenURL("http://cvs.khu.ac.kr/~jaeyong1/ip/jy_ip_save.php")
    Timer1.Enabled = False
    Unload Me
End If


Exit Sub



er:
If cnt >= 8 Then
    Unload Me
End If

End Sub

