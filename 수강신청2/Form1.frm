VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "Shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text2 
      Height          =   2895
      Left            =   12840
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   11
      Text            =   "Form1.frx":0000
      Top             =   2400
      Width           =   5415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "테스트"
      Height          =   495
      Left            =   13560
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   10200
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser3 
      Height          =   2175
      Left            =   7200
      TabIndex        =   6
      Top             =   7800
      Width           =   6495
      ExtentX         =   11456
      ExtentY         =   3836
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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Text            =   "http://sugang.khu.ac.kr"
      Top             =   240
      Width           =   3615
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   2055
      Left            =   360
      TabIndex        =   1
      Top             =   7920
      Width           =   6135
      ExtentX         =   10821
      ExtentY         =   3625
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
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   12255
      ExtentX         =   21616
      ExtentY         =   11456
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
   Begin VB.Label Label5 
      Height          =   375
      Left            =   7200
      TabIndex        =   13
      Top             =   7440
      Width           =   6375
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   7560
      Width           =   5895
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label3"
      Height          =   615
      Left            =   13320
      TabIndex        =   10
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '단일 고정
      Caption         =   "Label2"
      Height          =   1935
      Left            =   14640
      TabIndex        =   9
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '단일 고정
      Caption         =   "위치보정"
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long





Private Sub Command1_Click()
    WebBrowser1.Navigate2 (Text1)
End Sub

Private Sub Command2_Click()
WebBrowser2.Navigate "javascript:It('1');" 'javascript:It('4');
End Sub

Private Sub Command3_Click()
With WebBrowser1
        .Document.script.attribute.Value = "login"
        .Document.script.Action = "/sugang/haksugang/login"
        .Document.script.submit
        
        End With
        
        End Sub

'소스보기로 내용 분리해내기
Private Sub Command4_Click()

'소스보기
Label2 = WebBrowser1.Document.documentElement.Outerhtml
Text2 = WebBrowser1.Document.documentElement.Outerhtml
'중간에 찾기
'Label3 = Mid(Label2, 10, 10)
Label3 = InStr(Label2, "<DIV class=newst>")


'''문자열에서 a 들어간 위치 찾아내는거.. 앞에서부터 순서대로..
'Dim s
's = "asdfasdf"             '기준문자열
'Label1 = InStr(s, "a")     '몇번째에 a가 있다.
'MsgBox "a"                 '한타임쉬고
's = Mid(s, Label1 + 1, Len(s)) '앞의a위치까지 빼고
'Label1 = InStr(s, "a")     '다음 a위치


End Sub

Private Sub Form_Load()
   Call Command1_Click
End Sub

Private Sub Label1_Click()
WebBrowser2.Left = 360
WebBrowser2.Top = 7920
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
   
   WebBrowser2.RegisterAsBrowser = True
   Set ppDisp = WebBrowser2.object
 Do While WebBrowser2.Busy
        DoEvents
    Loop

End Sub

Private Sub WebBrowser1_TitleChange(ByVal Text As String)
Form1.Caption = WebBrowser1.LocationName


End Sub

Private Sub WebBrowser2_NewWindow2(ppDisp1 As Object, Cancel As Boolean)
WebBrowser3.RegisterAsBrowser = True

Set ppDisp1 = WebBrowser3.object
'Set ppDisp1 = WebBrowser2.object

End Sub

Private Sub WebBrowser2_TitleChange(ByVal Text As String)
'타이틀표시
Label4 = WebBrowser2.LocationName

End Sub

Private Sub WebBrowser3_StatusTextChange(ByVal Text As String)
Label5 = WebBrowser3.LocationName
End Sub
