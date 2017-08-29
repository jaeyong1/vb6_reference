VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   6120
      TabIndex        =   6
      Text            =   "10"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "즉구누르기(&a)"
      Height          =   735
      Left            =   6000
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   120
   End
   Begin SHDocVwCtl.WebBrowser WB2 
      Height          =   8655
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   5655
      ExtentX         =   9975
      ExtentY         =   15266
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
      Location        =   ""
   End
   Begin VB.CheckBox Check1 
      Caption         =   "auto refresh"
      Height          =   255
      Left            =   6840
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1320
      TabIndex        =   2
      Text            =   "http://www.auction.co.kr/buy/bid_event.asp?ItemNo=A026818636&BidType=1&strOptYn=N&strIlLegal=N"
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO(&G)"
      Default         =   -1  'True
      Height          =   300
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   810
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   11055
      Left            =   5880
      TabIndex        =   0
      Top             =   1440
      Width           =   8655
      ExtentX         =   15266
      ExtentY         =   19500
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
      Location        =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "보안숫자"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "refresh speed"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  팝업창이 왼쪽에 뜸..
'
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private IE As InternetExplorer
Private IE2 As InternetExplorer
Private Pop As Object
Public start As Integer
Public cnt As Integer

Private Sub Form_Activate()
    Text1.SetFocus
    Call Command1_Click
End Sub

Private Sub Check1_Click()
    Timer1.Enabled = Not (Timer1.Enabled)
End Sub

Private Sub Command1_Click()
    WebBrowser1.Navigate2 (Text1)
End Sub


Private Sub Timer1_Timer() '주기적으로 확인..
If WebBrowser1.Busy = True Then
    DoEvents
Else
    cnt = cnt + 1
    If cnt = Text2 Then
        cnt = 0
        WebBrowser1.Refresh
        Print "refresh"
    End If
End If

Call Command3_Click '계속 즉구누르기..
End Sub

Private Sub Command3_Click() '즉구누르기
'즉구버튼 찾아서 눌르는거..
If start = 1 Then

    Dim IE_Frm As HTMLFormElement
    Set IE_Frm = WebBrowser1.Document.getElementsByName("form")(0)      ' html 에서 form의 이름을 얻어온다. 소스보기에서 method=post 글자를 찾으면 됨. <form~~ 으로 시작하고 name값 을 가져와야함.. -재용-

    On Error GoTo ee '에러시건너뜀
        IE_Frm.submit
ee:

End If
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
start = 1
'팝업창뜨면 여기가 실행됨..(게임시작)
Timer1.Enabled = False '자동고침타이머끄고
Text4.SetFocus
Form1.Cls: Print "popup"
   
   WB2.RegisterAsBrowser = True
   Set ppDisp = WB2.object
 Do While WB2.Busy
        DoEvents
    Loop


'Pop.Navigate2 "http://cvs2.khu.ac.kr/"
End Sub

Private Sub Text4_Change()
'보안문자5자리 적으면 자동실행


If Len(Text4) = 5 Then
Form1.Cls: Print "보안문자입력"
On Error GoTo e3
    'Dim IE_id As HTMLInputElement
    Dim IE_pwd As HTMLInputElement  'as 뒤에꺼 정하는 방법 <input ~~~ class=input> 일경우 HTMLInputElement
    Dim IE_Frm As HTMLFormElement
    
    'Set IE_id = WB1.Document.getElementsByName("id")(0)     ' html 에서 input의 id 관련 태그 얻어온다.   소스보기에서 아이디 위치에 있는 name 값
    Set IE_pwd = WB2.Document.getElementsByName("txtSecText")(0) '  <- 이거 찾아서 바꿔줘야함. html 에서 input의 pwd 관련 태그 얻어온다.   소스보기에서 패스워드 위치에 있는 name 값
    Set IE_Frm = WB2.Document.getElementsByName("frmSecText")(0)      ' html 에서 form의 이름을 얻어온다. 소스보기에서 method=post 글자를 찾으면 됨. <form~~ 으로 시작하고 name값 을 가져와야함.. -재용-


'     If TypeName(IE_id) <> "Nothing" And TypeName(IE_pwd) <> "Nothing" And TypeName(IE_Frm) <> "Nothing" Then
'        IE_id.setAttribute "value", "jaeyong1" 'Text1.Text
        IE_pwd.setAttribute "value", Text4.Text  '보안문자값
        IE_Frm.submit

    'End If
e3:
End If



End Sub

