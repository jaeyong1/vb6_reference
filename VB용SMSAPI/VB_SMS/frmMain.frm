VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "SMSWorld"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtCNum 
      Enabled         =   0   'False
      Height          =   330
      Left            =   1170
      TabIndex        =   9
      Text            =   "0"
      Top             =   2655
      Width           =   465
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "보내기(&S)"
      Height          =   435
      Left            =   1785
      TabIndex        =   8
      Top             =   2640
      Width           =   1200
   End
   Begin VB.TextBox txtSPhone 
      Height          =   330
      Left            =   1170
      TabIndex        =   6
      Top             =   2175
      Width           =   3075
   End
   Begin VB.TextBox txtMessage 
      Height          =   915
      Left            =   1185
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   2
      Top             =   1140
      Width           =   3060
   End
   Begin VB.TextBox txtRPhone 
      Height          =   330
      Left            =   1185
      TabIndex        =   1
      Top             =   705
      Width           =   3075
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "끝내기(&X)"
      Height          =   435
      Left            =   3045
      TabIndex        =   0
      Top             =   2640
      Width           =   1200
   End
   Begin VB.Label lblCNum 
      Caption         =   "글자수 :"
      Height          =   240
      Left            =   405
      TabIndex        =   10
      Top             =   2700
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "회신번호 :"
      Height          =   240
      Left            =   225
      TabIndex        =   7
      Top             =   2250
      Width           =   945
   End
   Begin VB.Label lblMessage 
      Caption         =   "보낼메세지 :"
      Height          =   240
      Left            =   60
      TabIndex        =   5
      Top             =   1215
      Width           =   1125
   End
   Begin VB.Label lblRPhone 
      Caption         =   "수신휴대폰 :"
      Height          =   240
      Left            =   75
      TabIndex        =   4
      Top             =   780
      Width           =   1125
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H80000011&
      Caption         =   "SMS 전송 프로그램"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1035
      TabIndex        =   3
      Top             =   195
      Width           =   2460
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SMSObj As SMSCOMLib.SMSAPI


Private Sub cmdExit_Click()
  End
End Sub

Private Sub cmdSend_Click()
  If txtRPhone.Text = "" Then
    MsgBox "수신휴대폰 번호를 입력해 주세요."
    Exit Sub
  End If

  If txtMessage.Text = "" Then
    MsgBox "보낼 메세지를 입력해 주세요."
    Exit Sub
  End If

  If txtCNum.Text > 80 Then
    MsgBox "본문의 길이는 80Bytes 이하로 가능합니다."
    Exit Sub
  End If

  SMSObj.ReCallNum = txtSPhone.Text
  SMSObj.SendSMS txtRPhone.Text, txtMessage.Text

  If SMSObj.RetCode = "000" Then
    MsgBox "정상적으로 발송되었습니다."
    txtRPhone.Text = ""
    txtMessage.Text = ""
    txtSPhone.Text = ""
  Else
    MsgBox "ErrCode : " & SMSObj.RetCode & vbCrLf & "ErrMessage : " & SMSObj.RetMsg, vbCritical
  End If

End Sub

Private Sub Form_Load()
  Set SMSObj = New SMSCOMLib.SMSAPI
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set SMSObj = Nothing
End Sub

Private Sub txtMessage_Change()
  txtCNum.Text = LenB(txtMessage.Text)
End Sub
