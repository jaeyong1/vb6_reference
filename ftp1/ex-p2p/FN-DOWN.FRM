VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form main 
   Caption         =   "파일나라 Downloader"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   Icon            =   "fn-down.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5265
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3360
      Top             =   4800
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   4800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "닫기"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   4800
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "다운로드할 파일 정보"
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5175
      Begin VB.Label filesizez 
         AutoSize        =   -1  'True
         Caption         =   "None"
         Height          =   180
         Left            =   1080
         TabIndex        =   9
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "파일크기:"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   780
      End
      Begin VB.Label filenamez 
         AutoSize        =   -1  'True
         Caption         =   "None"
         Height          =   180
         Left            =   1920
         TabIndex        =   7
         Top             =   480
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "다운로드할 파일명:"
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1560
      End
      Begin VB.Label ipz 
         AutoSize        =   -1  'True
         Caption         =   "None"
         Height          =   180
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "공유자 아이피:"
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "다운로드 속도 및 시간"
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   3360
      Width           =   5175
      Begin VB.Label tottime 
         AutoSize        =   -1  'True
         Caption         =   "None"
         Height          =   180
         Left            =   1080
         TabIndex        =   22
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "예상시간:"
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   780
      End
      Begin VB.Label needtime 
         AutoSize        =   -1  'True
         Caption         =   "None"
         Height          =   180
         Left            =   1080
         TabIndex        =   20
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "남은시간:"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   780
      End
      Begin VB.Label dotime 
         AutoSize        =   -1  'True
         Caption         =   "None"
         Height          =   180
         Left            =   960
         TabIndex        =   18
         Top             =   480
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "작업시간:"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   780
      End
      Begin VB.Label speed 
         AutoSize        =   -1  'True
         Caption         =   "None"
         Height          =   180
         Left            =   720
         TabIndex        =   16
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "속도:"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "다운로드 작업 정보"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   5175
      Begin MSComctlLib.ProgressBar per 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.ListBox worklist 
         Height          =   960
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label percent 
         AutoSize        =   -1  'True
         Caption         =   "0%"
         Height          =   180
         Left            =   4800
         TabIndex        =   14
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "다운로드:"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "작업내용:"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   780
      End
   End
   Begin MSWinsockLib.Winsock down 
      Left            =   2880
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filename As String, filesize, remotez As String, worknum, myip As String
Dim secz As Single
Dim intFileNum As Integer
Dim strByte() As Byte
Dim beforebyte As Single
Dim dirTmp As String
Dim fileByte_Tot As Long
Dim downpath As String
Private Sub Command1_Click()
    End
End Sub
Private Sub down_Connect()
    worknum = 0
    working "연결되었습니다."
    down.SendData "ip:" + myip
    working "아이피 주소를 전송합니다."
End Sub
Function working(body As String)
    worklist.AddItem body
    worklist.TopIndex = worklist.ListCount - 1
End Function
Private Sub down_DataArrival(ByVal bytesTotal As Long)
    Select Case worknum
        Case 0
            Dim revdata As String
            down.GetData revdata
            
            If revdata = "ok-your-ip" Then
                working "전송할 파일명을 전송합니다."
                down.SendData "filename:" + filename
                Exit Sub
            End If
            
            If revdata = "no-your-ip" Then
                working "사용자가 제한한 사용자가 아닙니다."
                working "다운로드를 종료합니다."
                down.Close
                Timer1.Enabled = False
                Exit Sub
            End If
            
            If revdata = "findfile" Then
                working "파일의 존재가 확인되었습니다."
                working "다운로드를 시작합니다."
                worknum = 1
                down.SendData "send file now"
                Exit Sub
            Else
                working "파일이 없습니다."
                working "작업을 종료합니다."
                down.Close
            End If
        Case 1
            ReDim strByte(bytesTotal)
            down.GetData strByte, vbArray + vbByte, bytesTotal
            
            Open downpath + "\" + filename For Binary Access Write As #3
                Put #3, fileByte_Tot + 1, strByte
                fileByte_Tot = fileByte_Tot + bytesTotal
                If Int((fileByte_Tot / filesize) * 100) <= 100 Then
                    percent = nospace(Str(Int((fileByte_Tot / filesize) * 100)) + "%")
                    per.Value = Val(percent)
                Else
                    percent = nospace(Str(100) + "%")
                    per.Value = Val(percent)
                End If
            Close #3
            
            If fileByte_Tot >= filesize Then
                working "파일 내려받기가 완료되었습니다."
                down.SendData "download-ok"
            End If
    End Select
End Sub
Function middata(midbody, num As Single)
    x = 0
    For i = 1 To Len(midbody)
        If Mid$(midbody, i, 1) = "/" Then x = x + 1
        If x = num Then
            For j = i - 1 To 0 Step -1
                If j > 0 Then
                    If Mid$(midbody, j, 1) = "/" Then
                        middata = midz(midbody, j + 1, i - 1)
                        Exit Function
                    End If
                End If
                
                If j = 0 Then
                    middata = midz(midbody, j + 1, i - 1)
                    Exit Function
                End If
            Next
        End If
    Next
End Function
Function midz(midbody, startnum As Single, endnum As Single)
    midz = Mid$(midbody, startnum, endnum - startnum + 1)
End Function
Private Sub Form_Load()
'    Command$ = "211.199.231.242/127.0.0.1/tk3.lcd/70566448/"
    If Command$ = "" Then MsgBox "소리나라에 의해서만 작동하는 프로그램입니다." + vbCrLf + "소리나라에 의한 실행이 아니므로 종료합니다.", , "오류": End
    
    downpath = middata(Command$, 1)
    remotez = middata(Command$, 2)
    ipz = remotez
    myip = middata(Command$, 3)
    filename = middata(Command$, 4)
    filenamez = filename
    filesize = Val(middata(Command$, 5))
    filesizez = Str(filesize)
    
    If Dir(downpath, vbDirectory) = "" Then
        MsgBox "설정한 다운로드 파일 저장 폴더를 찾을 수 없습니다." + vbCrLf + "재설정을 하세요.(지금은 바탕화면에 저장됩니다)", , "오류"
        downpath = "c:\windows\바탕 화면"
    End If
    
    'remotez = "127.0.0.1"
    'myip = "127.0.0.1"
    'filename = "tk3.lcd"
    'filesize = 70566448
    
    down.RemoteHost = remotez
    down.RemotePort = 30303
    down.Connect
    
    working "연결중입니다."
End Sub
Function SecToTime(sec As Single)
    SecToTime = Str(Int(sec / 3600)) + "시간" + Str(Int((sec Mod 3600) / 60)) + "분" + Str(Int(sec Mod 60)) + "초"
End Function
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
Function nospace(getspace As String)
    Dim i As Single
    Dim tot As String
    
    tot = ""
    
    For i = 1 To Len(getspace)
        If Mid$(getspace, i, 1) <> " " Then tot = tot + Mid$(getspace, i, 1)
    Next
    
    nospace = tot
End Function
Private Sub Timer1_Timer()
    If per.Value = 100 Then Timer1.Enabled = False: Exit Sub
    secz = secz + 1
    dotime = SecToTime(secz)
    If secz = 10 And worklist.ListCount = 1 Then
        MsgBox "연결할 수 없습니다.", , "오류"
        End
    End If
End Sub
Private Sub timer2_Timer()
    If per.Value = 100 Then Timer2.Enabled = False: Exit Sub
    speed = Str(Int((fileByte_Tot - beforebyte) / 1000)) + "k/s"
    If (fileByte_Tot - beforebyte) > 0 Then
        needtime = SecToTime(Int((filesize - fileByte_Tot) / (fileByte_Tot - beforebyte)))
        tottime = SecToTime(Int(filesize / (fileByte_Tot - beforebyte)))
    End If
    beforebyte = fileByte_Tot
End Sub
