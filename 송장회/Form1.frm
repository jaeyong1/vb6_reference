VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "송.장.회."
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame2 
      Caption         =   "~문자메세지~"
      Height          =   6375
      Left            =   2640
      TabIndex        =   14
      Top             =   1560
      Width           =   7695
      Begin VB.CheckBox Check1 
         Caption         =   " 1. 구상건(011-884-6831)"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 1. 구상건(011-884-6831)"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 1. 구상건(011-884-6831)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "보내기(&S)"
         Height          =   435
         Left            =   4920
         TabIndex        =   23
         Top             =   2280
         Width           =   1200
      End
      Begin VB.TextBox txtCNum 
         Enabled         =   0   'False
         Height          =   330
         Left            =   4440
         TabIndex        =   19
         Text            =   "0"
         Top             =   2280
         Width           =   465
      End
      Begin VB.TextBox txtSPhone 
         Height          =   330
         Left            =   4800
         TabIndex        =   18
         Text            =   "000-0000-0000"
         Top             =   1800
         Width           =   1395
      End
      Begin VB.TextBox txtMessage 
         Height          =   915
         Left            =   4200
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   17
         Top             =   720
         Width           =   1980
      End
      Begin VB.CheckBox Check1 
         Caption         =   " 1. 구상건(011-884-6831)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblMessage 
         Caption         =   "보낼메세지 :"
         Height          =   240
         Left            =   3120
         TabIndex        =   22
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "회신번호 :"
         Height          =   240
         Left            =   3720
         TabIndex        =   21
         Top             =   1875
         Width           =   945
      End
      Begin VB.Label lblCNum 
         Caption         =   "글자수 :"
         Height          =   240
         Left            =   3705
         TabIndex        =   20
         Top             =   2325
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(대충 프리첼순)"
         Height          =   180
         Left            =   6240
         TabIndex        =   16
         Top             =   240
         Width           =   1290
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "자동로긴"
      Height          =   1575
      Left            =   480
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CommandButton Command6 
         Caption         =   "저 장"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   840
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   840
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "비번 :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblid 
         Caption         =   "아뒤 :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   4320
      TabIndex        =   7
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command5 
      Caption         =   "이야기"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "쿵쿵따"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   3300
      Left            =   7080
      Top             =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "앨범"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "게시판"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6975
      ExtentX         =   12303
      ExtentY         =   9763
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Caption         =   "<auto Refresh>"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Type 전화번호
' 이름 As String
' 번호 As String
'End Type
'Dim tel(1 To 19)  As 전화번호
'tel(1).이름 = "구": tel(1).번호 = "011-884-6831"
'tel(2).이름 = "김성식": tel(2).번호 = ""
'tel(3).이름 = "김태균": tel(3).번호 = ""
'tel(4).이름 = "김현우": tel(4).번호 = ""
'tel(5).이름 = "박재용": tel(5).번호 = ""
'tel(6).이름 = "박진재": tel(6).번호 = ""
'tel(7).이름 = "배현근": tel(7).번호 = ""
'tel(8).이름 = "손병찬": tel(8).번호 = ""
'tel(9).이름 = "송용찬": tel(9).번호 = ""
'tel(10).이름 = "이욱조": tel(10).번호 = ""
'tel(11).이름 = "이향미": tel(11).번호 = ""
'tel(12).이름 = "이희범": tel(12).번호 = ""
'tel(13).이름 = "장진영": tel(13).번호 = ""
'tel(14).이름 = "정주호": tel(14).번호 = ""
'tel(15).이름 = "홍길동": tel(15).번호 = ""
'tel(16).이름 = "": tel(16).번호 = ""
'tel(17).이름 = "": tel(17).번호 = ""
'tel(18).이름 = "": tel(18).번호 = ""
'tel(19).이름 = "": tel(19).번호 = ""
'Dim i As Integer
'For i = 1 To 3
'Check1(i).Caption = tel(i).이름 & "(" & tel(i).번호 & ")"
'Next i
 
 
 





Private Sub cmdSend_Click()
'  If txtRPhone.Text = "" Then
'    MsgBox "수신휴대폰 번호를 입력해 주세요."
'    Exit Sub
'  End If

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


Private Sub Command1_Click()
Frame1.Visible = True
WebBrowser1.Navigate2 (App.Path & "\aaa.html")

End Sub

Private Sub Command2_Click()
Frame1.Visible = False
WebBrowser1.Navigate2 ("http://bbs.freechal.com/ComService/Activity/BBS/CsBBSList.asp?GrpId=2078062&ObjSeq=1")
End Sub

Private Sub Command3_Click()
Frame1.Visible = False
WebBrowser1.Navigate2 ("http://community.freechal.com/ComService/Activity/Album/CsPhotoList.asp?GrpId=2078062&ObjSeq=1&grpurl=songjangclub")

End Sub

Private Sub Command4_Click()
Frame1.Visible = False
WebBrowser1.Navigate2 ("http://bbs.freechal.com/ComService/Activity/BBS/CsBBSList.asp?GrpId=2078062&ObjSeq=2")
End Sub

Private Sub Command5_Click()
Frame1.Visible = False
WebBrowser1.Navigate2 ("http://bbs.freechal.com/ComService/Activity/BBS/CsBBSList.asp?GrpId=2078062&ObjSeq=3")
End Sub

Private Sub Command6_Click()
Dim filename, nextline As String
Dim filenum As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ C:\DIAL1.TXT 읽어서 기억
filename = App.Path & "\aaa.html"
filenum = FreeFile
Open filename For Output As filenum
a = "<html>"
                         Print #filenum, a
a = "<body onload=""login.submit()"">"

                         Print #filenum, a
a = "<form method=""post"" name=""login"" action=""http://login.freechal.com/FcNwVerify.asp"">"
                         Print #filenum, a
a = "  <input type=""text"" name=""UserID"" value=""" + Text1 + """ style=""visibility:hidden"">"
                         Print #filenum, a
a = "  <input type=""password""  name=""Password"" value=""" + Text2 + """ style=""visibility:hidden"">"
                         Print #filenum, a
a = "</form>"
                         Print #filenum, a
a = "</body>"
                         Print #filenum, a
a = "</html>"
                         Print #filenum, a


Close #filenum

End Sub

Private Sub Form_Activate()
Dim ht As String
With Me
.Left = (Screen.Width - Me.ScaleWidth) / 2
.Top = (Screen.Height - Me.ScaleHeight) / 2
End With
Label1.Left = Form1.Width - 1435
WebBrowser1.Height = Form1.Height - 990
WebBrowser1.Width = Form1.Width - 345
WebBrowser1.Left = 120
ProgressBar1.Width = Form1.Width - 4520
'WebBrowser1.Navigate2 ("http://my-cgi.dreamwiz.com/desirelove/")
WebBrowser1.Navigate2 (App.Path & "\aaa.html")
Frame1.Visible = False
End Sub

Private Sub Form_Click()
On Error GoTo er
WebBrowser1.GoBack
er:
End Sub

Private Sub Form_LostFocus()
Timer1.Enabled = True
Label1.Caption = "<auto Refresh>"

End Sub

Private Sub Form_Resize()
On Error GoTo er

Label1.Left = Form1.Width - 1435
WebBrowser1.Height = Form1.Height - 990
WebBrowser1.Width = Form1.Width - 345
WebBrowser1.Left = 120
ProgressBar1.Width = Form1.Width - 4520

er:

End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = 100 Then
WebBrowser1.Refresh
ProgressBar1.Value = 0
End If
End Sub

Private Sub txtMessage_Change()
  txtCNum.Text = LenB(txtMessage.Text)
End Sub

Private Sub WebBrowser1_GotFocus()
Timer1.Enabled = False
Label1.Caption = "<user ready>"
ProgressBar1.Value = 0
End Sub

Private Sub WebBrowser1_LostFocus()
Timer1.Enabled = True
Label1.Caption = "<auto Refresh>"
End Sub


