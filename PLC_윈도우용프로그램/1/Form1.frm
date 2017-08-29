VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "PLC-1"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command3 
      Caption         =   "모니터링 사이트"
      Height          =   615
      Left            =   240
      TabIndex        =   33
      Top             =   5520
      Width           =   3135
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5640
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "알람기능 Reset"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   240
      TabIndex        =   29
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2760
      Top             =   3000
   End
   Begin VB.Frame Frame1 
      Caption         =   "* 알람기능 *"
      Height          =   1575
      Left            =   3240
      TabIndex        =   24
      Top             =   2040
      Width           =   2775
      Begin VB.CommandButton Command2 
         Caption         =   "?"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1800
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         Caption         =   "사용안함"
         Height          =   180
         Left            =   360
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "사용"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label12 
         Height          =   255
         Left            =   2400
         TabIndex        =   30
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lbltel 
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1080
         Width           =   2295
      End
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   3120
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "종  료"
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   5520
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Interval        =   11000
      Left            =   2760
      Top             =   2520
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1455
      Left            =   2520
      TabIndex        =   0
      Top             =   3840
      Width           =   4215
      ExtentX         =   7435
      ExtentY         =   2566
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
   Begin VB.Label Label13 
      Caption         =   "Made by P.J.Y  2002. 02."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   32
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  '단일 고정
      Caption         =   " "
      Height          =   255
      Left            =   8400
      TabIndex        =   31
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblP12 
      Height          =   255
      Left            =   5400
      TabIndex        =   23
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblP11 
      Height          =   255
      Left            =   5400
      TabIndex        =   22
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblP10 
      Height          =   255
      Left            =   5400
      TabIndex        =   21
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblP09 
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblP08 
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblP07 
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblP06 
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblP05 
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblP02 
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblP01 
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "수중펌프 경고 :"
      Height          =   180
      Left            =   3720
      TabIndex        =   13
      Top             =   1680
      Width           =   1260
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "양수장모터2 경고 :"
      Height          =   180
      Left            =   3720
      TabIndex        =   12
      Top             =   1320
      Width           =   1530
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "양수장모터1 경고 :"
      Height          =   180
      Left            =   3720
      TabIndex        =   11
      Top             =   960
      Width           =   1530
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "메인펌프압력 :"
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   1200
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "양수장모터2 :"
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   1110
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "양수장모터1 : "
      Height          =   180
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "수중펌프압력 : "
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "수중펌프 :"
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "저장탱크수위 : "
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1260
   End
   Begin VB.Label LL 
      AutoSize        =   -1  'True
      Caption         =   "우물수위 : "
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BorderStyle     =   1  '단일 고정
      Caption         =   "PLC 수신내용 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1485
   End
   Begin VB.Label lbldata 
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SMSObj As SMSCOMLib.SMSAPI
Private Sub Command1_Click()
End
End Sub



Private Sub Command2_Click()
Form2.Show 1
End Sub



Private Sub Command3_Click()
Shell ("C:\Program Files\Internet Explorer\IEXPLORE.EXE" & Space(1) & site + "monitor.php3")
End Sub

Private Sub Command4_Click()
Dim k As Integer
k = MsgBox("알람을 다시 준비시키겠습니까?" & vbCrLf & "문제가 해결되지 않은 상황에서 리셋을 시킬경우 핸드폰문자메세지가 다시 발송되게 됩니다. ", 32 + 4 + 256, "알람Reset")
If k = vbYes Then
    에러접수 = "OFF"
    t = 0
    Command4.Enabled = False
    Option1.Enabled = True
    Option2.Enabled = True
    Option1.SetFocus
    Label12 = ""
    lbltel = ""
End If
End Sub

Private Sub Form_Activate()


               site = "http://www.i-pws.com/sdwater/monitor/"
             ' 웹사이트 변경시 위의 내용을 바꿔주면 됨.
             ' http로 시작해서 /까지 포함하는 전체적인 주소로 표시
             ' 웹사이트용 프로그램은 PHP등 모든 파일의 퍼미션은 777로 하면 작동됨
             

w = 0
에러접수 = "OFF"

If 설정다운로드 = "했음" Then
  Else
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 사이트에서 전화번호 읽어서 기억
  lbltel = "Loading..0%"
  dial1 = Inet1.OpenURL(site + "dial1.jy")
  lbltel = "Loading..10%"
  dial2 = Inet1.OpenURL(site + "dial2.jy")
  lbltel = "Loading..20%"
  dial3 = Inet1.OpenURL(site + "dial3.jy")
  lbltel = "Loading..30%"
  dialcheck1 = Inet1.OpenURL(site + "dialcheck1.jy")
  lbltel = "Loading..40%"
  dialcheck2 = Inet1.OpenURL(site + "dialcheck2.jy")
  lbltel = "Loading..50%"
  dialcheck3 = Inet1.OpenURL(site + "dialcheck3.jy")
  lbltel = "Loading..60%"
  set1 = Inet1.OpenURL(site + "set1.jy")
  lbltel = "Loading..70%"
  set2 = Inet1.OpenURL(site + "set2.jy")
  lbltel = "Loading..80%"
  set3 = Inet1.OpenURL(site + "set3.jy")
  lbltel = "Loading..90%"
  set4 = Inet1.OpenURL(site + "set4.jy")
  lbltel = "Loading..100%"
  설정다운로드 = "했음"
  Command2.Enabled = True
  lbltel = ""
End If


End Sub


Private Sub Form_Load()  'sms준비

Set SMSObj = New SMSCOMLib.SMSAPI
Shell ("Regsvr32 c:\sdwater\SMSCOM.dll /s")


End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set SMSObj = Nothing
End Sub

Private Sub Option2_Click()
t = 0

End Sub

Private Sub Timer1_Timer()

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PLC와의 통신부분
lbldata = ""
On Error GoTo errmsg   '에러나면 errmsg로 이동후 대기...
  MSComm1.CommPort = 1
  MSComm1.Settings = "19200,n,8,1"
  MSComm1.InputLen = 1
  MSComm1.PortOpen = True
   q = Chr(5) & "00RSS0206%PW00106%PW000" & Chr(4) '06%PW001 06%PW000 두워드의 데이타 요청
   
   MSComm1.Output = q

Do
     instring = MSComm1.Input
     Rcv = Rcv & instring
     data = Rcv
Loop Until instring = Chr(3)
     Rcv = ""
     MSComm1.PortOpen = False
     
lbldata.Caption = data

     
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 수신 데이터 분석
Select Case Mid(data, 20, 1)  '01 우물수위 체크
Case "0" '저수위
lblP01 = "LOW"
p01 = "1"
Case "1" '중수위
lblP01 = "MIDDLE"
p01 = "2"
Case "3" '고수위
lblP01 = "HIGH"
p01 = "3"
Case "7" '초과
lblP01 = "OVER"
p01 = "4"
Case Else '센서이상
lblP01 = "ERROR"
p01 = "0"
End Select

Select Case Mid(data, 19, 1)  '02 저장탱크수위 체크
Case "0" '저수위
lblP02 = "LOW": p02 = "1"
Case "1" '중수위
lblP02 = "MIDDLE": p02 = "2"
Case "3" '고수위
lblP02 = "HIGH": p02 = "3"
Case "7" '초과
lblP02 = "OVER": p02 = "4"
Case Else '센서이상
lblP02 = "ERROR": p02 = "0"
End Select

Select Case Mid(data, 18, 1)  '05 06 07 08 체크
Case "0"
p08 = "1": p07 = "1": p06 = "1": p05 = "1"
lblP08 = "OFF": lblP07 = "OFF": lblP06 = "OFF": lblP05 = "OFF"
Case "1"
lblP08 = "OFF": lblP07 = "OFF": lblP06 = "OFF": lblP05 = "ON"
p08 = "1": p07 = "1": p06 = "1": p05 = "2"
Case "2"
p08 = "1": p07 = "1": p06 = "2": p05 = "1"
lblP08 = "OFF": lblP07 = "OFF": lblP06 = "ON": lblP05 = "OFF"
Case "3"
p08 = "1": p07 = "1": p06 = "2": p05 = "2"
lblP08 = "OFF": lblP07 = "OFF": lblP06 = "ON": lblP05 = "ON"
Case "4"
p08 = "1": p07 = "2": p06 = "1": p05 = "1"
lblP08 = "OFF": lblP07 = "ON": lblP06 = "OFF": lblP05 = "OFF"
Case "5"
p08 = "1": p07 = "2": p06 = "1": p05 = "2"
lblP08 = "OFF": lblP07 = "ON": lblP06 = "OFF": lblP05 = "ON"

Case "6"
p08 = "1": p07 = "2": p06 = "2": p05 = "1"
lblP08 = "OFF": lblP07 = "ON": lblP06 = "ON": lblP05 = "OFF"
Case "7"
p08 = "1": p07 = "2": p06 = "2": p05 = "2"
lblP08 = "OFF": lblP07 = "ON": lblP06 = "ON": lblP05 = "ON"
Case "8"
p08 = "2": p07 = "1": p06 = "1": p05 = "1"
lblP08 = "ON": lblP07 = "OFF": lblP06 = "OFF": lblP05 = "OFF"
Case "9"
p08 = "2": p07 = "1": p06 = "1": p05 = "2"
lblP08 = "ON": lblP07 = "OFF": lblP06 = "OFF": lblP05 = "ON"
Case "A"
p08 = "2": p07 = "1": p06 = "2": p05 = "1"
lblP08 = "ON": lblP07 = "OFF": lblP06 = "ON": lblP05 = "OFF"

Case "B"
p08 = "2": p07 = "1": p06 = "2": p05 = "2"
lblP08 = "ON": lblP07 = "OFF": lblP06 = "ON": lblP05 = "ON"
Case "C"
p08 = "2": p07 = "2": p06 = "1": p05 = "1"
lblP08 = "ON": lblP07 = "ON": lblP06 = "OFF": lblP05 = "OFF"
Case "D"
p08 = "2": p07 = "2": p06 = "1": p05 = "2"
lblP08 = "ON": lblP07 = "ON": lblP06 = "OFF": lblP05 = "ON"
Case "E"
p08 = "2": p07 = "2": p06 = "2": p05 = "1"
lblP08 = "ON": lblP07 = "ON": lblP06 = "ON": lblP05 = "OFF"
Case "F"
p08 = "2": p07 = "2": p06 = "2": p05 = "2"
lblP08 = "ON": lblP07 = "ON": lblP06 = "ON": lblP05 = "ON"

End Select

Select Case Mid(data, 17, 1)  'P09 10 11 12 체크
Case "0"
p12 = "0": p11 = "0": p10 = "0": p09 = "1"
lblP12 = "OFF": lblP11 = "OFF": lblP10 = "OFF": lblP09 = "OFF"
Case "1"
p12 = "0": p11 = "0": p10 = "0": p09 = "2"
lblP12 = "OFF": lblP11 = "OFF": lblP10 = "OFF": lblP09 = "ON"
Case "2"
p12 = "0": p11 = "0": p10 = "1": p09 = "1"
lblP12 = "OFF": lblP11 = "OFF": lblP10 = "ON": lblP09 = "OFF"
Case "3"
p12 = "0": p11 = "0": p10 = "1": p09 = "2"
lblP12 = "OFF": lblP11 = "OFF": lblP10 = "ON": lblP09 = "ON"
Case "4"
p12 = "0": p11 = "1": p10 = "0": p09 = "1"
lblP12 = "OFF": lblP11 = "ON": lblP10 = "OFF": lblP09 = "OFF"
Case "5"
p12 = "0": p11 = "1": p10 = "0": p09 = "2"
lblP12 = "OFF": lblP11 = "ON": lblP10 = "OFF": lblP09 = "ON"

Case "6"
p12 = "0": p11 = "1": p10 = "1": p09 = "1"
lblP12 = "OFF": lblP11 = "ON": lblP10 = "ON": lblP09 = "OFF"
Case "7"
p12 = "0": p11 = "1": p10 = "1": p09 = "2"
lblP12 = "OFF": lblP11 = "ON": lblP10 = "ON": lblP09 = "ON"
Case "8"
p12 = "1": p11 = "0": p10 = "0": p09 = "1"
lblP12 = "ON": lblP11 = "OFF": lblP10 = "OFF": lblP09 = "OFF"
Case "9"
p12 = "1": p11 = "0": p10 = "0": p09 = "2"
lblP12 = "ON": lblP11 = "OFF": lblP10 = "OFF": lblP09 = "ON"
Case "A"
p12 = "1": p11 = "0": p10 = "1": p09 = "1"
lblP12 = "ON": lblP11 = "OFF": lblP10 = "ON": lblP09 = "OFF"

Case "B"
p12 = "1": p11 = "0": p10 = "1": p09 = "2"
lblP12 = "ON": lblP11 = "OFF": lblP10 = "ON": lblP09 = "ON"
Case "C"
p12 = "1": p11 = "1": p10 = "0": p09 = "1"
lblP12 = "ON": lblP11 = "ON": lblP10 = "OFF": lblP09 = "OFF"
Case "D"
p12 = "1": p11 = "1": p10 = "0": p09 = "2"
lblP12 = "ON": lblP11 = "ON": lblP10 = "OFF": lblP09 = "ON"
Case "E"
p12 = "1": p11 = "1": p10 = "1": p09 = "1"
lblP12 = "ON": lblP11 = "ON": lblP10 = "ON": lblP09 = "OFF"
Case "F"
p12 = "1": p11 = "1": p10 = "1": p09 = "2"
lblP12 = "ON": lblP11 = "ON": lblP10 = "ON": lblP09 = "ON"

End Select


  '~~~~~~ 인터넷으로 전송하기 위해서 수집데이터를 한줄로 요약
  webQ = site + "plcwrite-1.php3?p01=" & p01 & "&p02=" & p02 & "&p05=" & p05 & "&p06=" & p06 & "&p07=" & p07 & "&p08=" & p08 & "&p09=" & p09 & "&p10=" & p10 & "&p11=" & p11 & "&p12=" & p12 & "&w=" & w & "&dial1=" & dial1 & "&dial2=" & dial2 & "&dial3=" & dial3 & "&dialcheck1=" & dialcheck1 & "&dialcheck2=" & dialcheck2 & "&dialcheck3=" & dialcheck3 & "&set1=" & set1 & "&set2=" & set2 & "&set3=" & set3 & "&dial3=" & dial3 & "&set4=" & set4


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 인터넷으로 전송
w = w + 1
If w = 10 Then
 w = 0: WebBrowser1.Navigate ("kr.yahoo.com")
Else
    If 설정다운로드 = "했음" Then '설정상태를 다 받아오기 전까지는 데이터를 전송하지 않음..
     WebBrowser1.Navigate (webQ)
    End If
End If
'~~~~~~~~~~~~~~~~~~-<<<<  알람기능 >>>>>
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 알람기능
If 에러접수 = "OFF" Then

 If (p01 = "1") Or (p02 = "1") Or (p10 = "1") Or (p11 = "1") Or (p12 = "1") Then
    If Option1.Value = True Then

    에러접수 = "ON"
    Command4.Enabled = True
    'Timer2.Enabled = True '<-  전화알람 작동
             '~~~~~~~ SMS메세지 판단 및 전송부분
             If p01 = "1" Then
                smsmsg = smsmsg & "우물저수위"
             End If
             
             If p02 = "1" Then
                         smsmsg = smsmsg & "저장탱크저수위"
             End If
             
             If p10 = "1" Then
                         smsmsg = smsmsg & "양수장모터1과부하"
             End If
             
             If p11 = "1" Then
                         smsmsg = smsmsg & "양수장모터2과부하"
            End If
             
             If p12 = "1" Then
                         smsmsg = smsmsg & "수중펌프과부하"
            End If
             smsmsg = "*[삼도양수관리]*알람내용-" + smsmsg
                        If LenB(smsmsg) > 78 Then '---- 문자전송내용이 80자가 초과되었을때..
                        smsmsg = "*[삼도양수관리]*알람내용-복합적경보발생확인필요"
                        End If
              
               If (Left(dial1, 2) = "01") And (dialcheck1 = "1") Then '01로 시작하는 연락처 and SMS체크여부
                  lbltel = " 연락처1 문자전송중"
                  SMSObj.ReCallNum = "9999" '보내는휴대폰
                  SMSObj.SendSMS dial1, smsmsg
                  
               End If
               
               If (Left(dial2, 2) = "01") And (dialcheck2 = "1") Then
                  lbltel = " 연락처2 문자전송중"
                  SMSObj.ReCallNum = "9999" '보내는휴대폰
                  SMSObj.SendSMS dial2, smsmsg
               End If
               If (Left(dial3, 2) = "01") And (dialcheck3 = "1") Then
                  lbltel = " 연락처3 문자전송중"
                  SMSObj.ReCallNum = "9999" '보내는휴대폰
                  SMSObj.SendSMS dial3, smsmsg
               End If
               lbltel = " ALERT"
             smsmsg = ""
             '~~~~~~~ SMS끝
          
    Option2.Enabled = False
    Option1.Enabled = False
    
    'Label12 = "T"
    End If
 End If

End If



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 프로그래밍 에러발생시 메세지창 작동
Exit Sub
errmsg:
  MsgBox "프로그램작동중에 에러가 발생했습니다.", 0, "에러~"

End Sub

