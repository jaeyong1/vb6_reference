VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form2 
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6255
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "닫 기"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   5760
      Width           =   5535
   End
   Begin MSFlexGridLib.MSFlexGrid gridSugang 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      SelectionMode   =   1
      FormatString    =   "^index |^과목코드 |^과목명 |^최대인원 |^신청인원 |^학점"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Height          =   7215
      Left            =   -120
      TabIndex        =   2
      Top             =   -240
      Width           =   9135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide

End Sub

Private Sub gridSugang_Click()
'수강인원 보기

Dim tmpmsg, newline As String
Dim a

newline = Chr(13) + Chr(10)
tmpmsg = tmpmsg + "                   " + newline
tmpmsg = tmpmsg + "  수강신청 인원   " + newline
tmpmsg = tmpmsg + "----------------- " + newline + newline

Dim datacount As Integer
With Form3.gridData

datacount = .Rows - 1

For i = 1 To datacount
    If .TextMatrix(i, 1) = gridSugang.TextMatrix(gridSugang.MouseRow, 1) Then
        tmpmsg = tmpmsg + "  " + .TextMatrix(i, 0) + newline
 
    End If


Next i

a = MsgBox(tmpmsg, vb_ok, " 수강신청")




End With


End Sub
