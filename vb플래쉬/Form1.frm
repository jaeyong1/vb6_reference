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
   ScaleMode       =   3  '波漆
   ScaleWidth      =   946
   StartUpPosition =   3  'Windows 奄沙葵
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
         Name            =   "閏顕"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "閏顕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "閏顕"
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
'覗稽詮闘->凧繕->Microsoft ActiveX Data Object 2.5 Library
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sch_j, sch_i, sch_data, DayStr, Firstday
Dim toDB劾促, toDB奄系獣娃, toDB鎧遂, toDB走紗獣娃


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
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\schedule\什追糟.mdb;Persist Security Info=False"
cn.CursorLocation = adUseClient
cn.CommandTimeout = 30
cn.Open

End Function

Public Function insertDB()  'flash->VB->access
Dim sql
sql = "INSERT INTO schedule(劾促, 獣娃企, 析舛鎧遂, 走紗獣娃) VALUES('" & _
            toDB劾促 & _
            "' , '" & _
            toDB奄系獣娃 & _
            "' , '" & _
            toDB鎧遂 & _
            "' , '" & _
            toDB走紗獣娃 & _
            "') "

Debug.Print sql
On Error GoTo er
cn.Execute (sql)
Exit Function

er:
MsgBox "拭君降持生稽 坦軒鞠走 公梅製", , "坦軒拭君"

End Function

Private Sub Calendar1_Click()
    With Calendar1
        Firstday = DateSerial(.Year, .Month, .Day)
    End With
    Call scheduleView
End Sub

Public Function scheduleView()

'段奄鉢
sch_j = 0
ShockwaveFlash1.SetVariable "setaction.allclear", ""


For sch_j = 0 To 7  '析(Day)紺稽 稽球

    getday = Month(Firstday + sch_j) & "/" & Day(Firstday + sch_j) & "/" & Year(Firstday + sch_j)
    Set rs = cn.Execute("TRANSFORM Last(schedule.析舛鎧遂) AS [析舛鎧遂税原走厳 葵] SELECT time.獣娃企 FROM schedule INNER JOIN [time] ON schedule.獣娃企=time.獣娃企 Where (((schedule.劾促) = #" & getday & "#)) GROUP BY time.獣娃企 PIVOT schedule.劾促;")
    ShockwaveFlash1.SetVariable "setaction.day" & sch_j, Day(Firstday + sch_j) & "析"

'On Error GoTo skips '什追糟 蒸澗劾..
If Not (rs.EOF) Then
    rs.MoveFirst
End If
'skips: 'on error 坦軒


While Not (rs.EOF)

sch_i = rs(0)
sch_data = rs(1)
    'vb -> flash
    With ShockwaveFlash1
        .SetVariable "setaction.day" & sch_j, Day(Firstday + sch_j) & "析"  '雌舘 劾促(析)妊獣
        .SetVariable "setaction.data1", "schdata" & sch_j & "-" & sch_i     '古闘遣什鎧 妊獣是帖 走舛
        .SetVariable "setaction.data2", sch_data                            '古闘遣什鎧 汽戚斗 走舛
    End With
    
'Debug.Print "schdata" & sch_j & "-" & sch_i  '3-11"
'Debug.Print sch_data

rs.MoveNext
Wend


Next ' sch_j
End Function

'巴掘獣稽採斗 汽戚斗呪重
Private Sub ShockwaveFlash1_FSCommand(ByVal command As String, ByVal args As String)
Label1 = command
Label2 = args

Select Case command
Case "劾促"
    toDB劾促 = args

Case "奄系獣娃"
    toDB奄系獣娃 = args
    
Case "鎧遂"
    toDB鎧遂 = args
    
Case "走紗獣娃"
    toDB走紗獣娃 = args
    Call timeconvert
    Call insertDB
    Call scheduleView

Case "脊径曽戟"
    Call scheduleView

Case Else
    MsgBox "鉦紗照吉誤敬降持 : " & command & " " & args
End Select

End Sub

Public Function timeconvert()

If Not (IsNumeric(toDB奄系獣娃)) Then
    MsgBox "獣娃精 24獣奄層生稽 収切幻 隔嬢醤 杯艦陥.", "拭君"
    Exit Function
End If

If (5 <= toDB奄系獣娃) And (toDB奄系獣娃 <= 24) Then
    toDB奄系獣娃 = toDB奄系獣娃 - 4
ElseIf (toDB奄系獣娃 < 5) Then
    toDB奄系獣娃 = toDB奄系獣娃 + 20
Else
    MsgBox "獣娃企神嫌脊艦陥. 24左陥 笛呪澗 照鞠推.", "拭君"
End If

End Function


Private Sub Command1_Click() 'for test
With ShockwaveFlash1
    .SetVariable "setaction.day0", "4析" 'today
    .SetVariable "setaction.day1", "5析"
    .SetVariable "setaction.day2", "6析"
    .SetVariable "setaction.day3", "7析"
    .SetVariable "setaction.day4", "8析"
    .SetVariable "setaction.day5", "9析"
    .SetVariable "setaction.day6", "10析"
    .SetVariable "setaction.day7", "11析"
    .SetVariable "setaction.day8", "12析"
    
    .SetVariable "setaction.test1", "轄陥"
    
    .SetVariable "setaction.data1", "schdata3-12"
    .SetVariable "setaction.data2", "拭伯虞巨12312312醤~"
    
    .SetVariable "setaction.data1", "schdata3-11"
    .SetVariable "setaction.data2", "けけけけけけけけけけけけけけけ"
    
    
    
   ' "schdata"+j+"-"+i
End With


With ShockwaveFlash1
'ShockwaveFlash1.SetVariable "setaction.sDown", "酔畑" '痕呪誤, 汽戚斗
'.SetVariable ("_root.txt1.text","76");

End With

'ShockwaveFlash1.SetVariable "setaction.day1", "21析" '痕呪誤, 汽戚斗

End Sub
