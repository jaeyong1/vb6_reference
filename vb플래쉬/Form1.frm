VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   ScaleHeight     =   763
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   946
   StartUpPosition =   3  'Windows 기본값
   Begin MSACAL.Calendar Calendar1 
      Height          =   2055
      Left            =   6600
      TabIndex        =   4
      Top             =   120
      Width           =   4095
      _Version        =   524288
      _ExtentX        =   7223
      _ExtentY        =   3625
      _StockProps     =   1
      BackColor       =   12648384
      Year            =   2006
      Month           =   12
      Day             =   14
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   16384
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   16744576
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   0   'False
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   10500
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   13500
      _cx             =   23812
      _cy             =   18521
      FlashVars       =   ""
      Movie           =   "c:\schedule\aaa.swf"
      Src             =   "c:\schedule\aaa.swf"
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "NoBorder"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'프로젝트->참조->Microsoft ActiveX Data Object 2.5 Library
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sch_j, sch_i, sch_data, DayStr, Firstday
Dim toDB날짜, toDB기록시간, toDB내용, toDB지속시간


Private Sub Form_Load()
Call DBopen
    With ShockwaveFlash1
       .Move 10
        .Movie = "c:\schedule\aaa.swf"
    End With
Firstday = Date
Call scheduleView
End Sub

Public Function DBopen()
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\schedule\스케쥴.mdb;Persist Security Info=False"
cn.CursorLocation = adUseClient
cn.CommandTimeout = 30
cn.Open

End Function

Public Function insertDB()  'flash->VB->access
Dim sql
sql = "INSERT INTO schedule(날짜, 시간대, 일정내용, 지속시간) VALUES('" & _
            toDB날짜 & _
            "' , '" & _
            toDB기록시간 & _
            "' , '" & _
            toDB내용 & _
            "' , '" & _
            toDB지속시간 & _
            "') "

Debug.Print sql
On Error GoTo er
cn.Execute (sql)
Exit Function

er:
MsgBox "에러발생으로 처리되지 못했음", , "처리에러"

End Function

Private Sub Calendar1_Click()
    With Calendar1
        Firstday = DateSerial(.Year, .Month, .Day)
    End With
    Call scheduleView
End Sub

Public Function scheduleView()

'초기화
sch_j = 0
ShockwaveFlash1.SetVariable "setaction.allclear", ""


For sch_j = 0 To 7  '일(Day)별로 로드

    getday = Month(Firstday + sch_j) & "/" & Day(Firstday + sch_j) & "/" & Year(Firstday + sch_j)
    Set rs = cn.Execute("TRANSFORM Last(schedule.일정내용) AS [일정내용의마지막 값] SELECT time.시간대 FROM schedule INNER JOIN [time] ON schedule.시간대=time.시간대 Where (((schedule.날짜) = #" & getday & "#)) GROUP BY time.시간대 PIVOT schedule.날짜;")
    ShockwaveFlash1.SetVariable "setaction.day" & sch_j, Day(Firstday + sch_j) & "일"

'On Error GoTo skips '스케쥴 없는날..
If Not (rs.EOF) Then
    rs.MoveFirst
End If
'skips: 'on error 처리


While Not (rs.EOF)

sch_i = rs(0)
sch_data = rs(1)
    'vb -> flash
    With ShockwaveFlash1
        .SetVariable "setaction.day" & sch_j, Day(Firstday + sch_j) & "일"  '상단 날짜(일)표시
        .SetVariable "setaction.data1", "schdata" & sch_j & "-" & sch_i     '매트릭스내 표시위치 지정
        .SetVariable "setaction.data2", sch_data                            '매트릭스내 데이터 지정
    End With
    
'Debug.Print "schdata" & sch_j & "-" & sch_i  '3-11"
'Debug.Print sch_data

rs.MoveNext
Wend


Next ' sch_j
End Function

'플래시로부터 데이터수신
Private Sub ShockwaveFlash1_FSCommand(ByVal command As String, ByVal args As String)
Label1 = command
Label2 = args

Select Case command
Case "날짜"
    toDB날짜 = args

Case "기록시간"
    toDB기록시간 = args
    
Case "내용"
    toDB내용 = args
    
Case "지속시간"
    toDB지속시간 = args
    Call timeconvert
    Call insertDB
    Call scheduleView

Case "입력종료"
    Call scheduleView

Case Else
    MsgBox "약속안된명령발생 : " & command & " " & args
End Select

End Sub

Public Function timeconvert()

If Not (IsNumeric(toDB기록시간)) Then
    MsgBox "시간은 24시기준으로 숫자만 넣어야 합니다.", "에러"
    Exit Function
End If

If (5 <= toDB기록시간) And (toDB기록시간 <= 24) Then
    toDB기록시간 = toDB기록시간 - 4
ElseIf (toDB기록시간 < 5) Then
    toDB기록시간 = toDB기록시간 + 20
Else
    MsgBox "시간대오류입니다. 24보다 큰수는 안되요.", "에러"
End If

End Function


Private Sub Command1_Click() 'for test
With ShockwaveFlash1
    .SetVariable "setaction.day0", "4일" 'today
    .SetVariable "setaction.day1", "5일"
    .SetVariable "setaction.day2", "6일"
    .SetVariable "setaction.day3", "7일"
    .SetVariable "setaction.day4", "8일"
    .SetVariable "setaction.day5", "9일"
    .SetVariable "setaction.day6", "10일"
    .SetVariable "setaction.day7", "11일"
    .SetVariable "setaction.day8", "12일"
    
    .SetVariable "setaction.test1", "논다"
    
    .SetVariable "setaction.data1", "schdata3-12"
    .SetVariable "setaction.data2", "에헤라디12312312야~"
    
    .SetVariable "setaction.data1", "schdata3-11"
    .SetVariable "setaction.data2", "ㅁㅁㅁㅁㅁㅁㅁㅁㅁㅁㅁㅁㅁㅁㅁ"
    
    
    
   ' "schdata"+j+"-"+i
End With


With ShockwaveFlash1
'ShockwaveFlash1.SetVariable "setaction.sDown", "우홋" '변수명, 데이터
'.SetVariable ("_root.txt1.text","76");

End With

'ShockwaveFlash1.SetVariable "setaction.day1", "21일" '변수명, 데이터

End Sub
