VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   ScaleHeight     =   11655
   ScaleWidth      =   13080
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text2 
      Height          =   8535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   2880
      Width           =   12975
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7800
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "http://stock.naver.com/item/sise_day.nhn?code=011810&page=1"
      Top             =   120
      Width           =   6015
   End
   Begin SHDocVwCtl.WebBrowser wb1 
      Height          =   2775
      Left            =   9600
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      ExtentX         =   6165
      ExtentY         =   4895
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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
wb1.Navigate (Text1.Text)

Cls
Dim str As String
Dim strlen As Integer

str = Inet1.OpenURL(Text1.Text)
strlen = Len(str)
Label1 = strlen
If strlen < 8000 Then
    Call Command1_Click
Else
    Text2 = str
'''''''''''''''''''''''''''''''''
Print "3번째는..." + kukfind(3, str)


    
'''''''''''''''''''''''''''''''''
End If
    


End Sub


'몇번째 꺽쇠 내용 가져올까..
Public Function kukfind(n As Integer, str As String) As String

Dim i As Integer
Dim cl '이전 >
Dim re As String

cl = 1

For i = 1 To n
    cl = InStr(cl + 1, str, ">")
    Print Mid(str, cl, 10) '>부터 10글자
    re = Mid(str, cl, 10)
    
Next i


kukfind = re




End Function





Private Sub wb1_DownloadComplete()
'Print wb1.Document.documentElement.innerText


End Sub

