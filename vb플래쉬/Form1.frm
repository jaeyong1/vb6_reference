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
   ScaleMode       =   3  '�ȼ�
   ScaleWidth      =   946
   StartUpPosition =   3  'Windows �⺻��
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
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
'������Ʈ->����->Microsoft ActiveX Data Object 2.5 Library
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sch_j, sch_i, sch_data, DayStr, Firstday
Dim toDB��¥, toDB��Ͻð�, toDB����, toDB���ӽð�


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
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\schedule\������.mdb;Persist Security Info=False"
cn.CursorLocation = adUseClient
cn.CommandTimeout = 30
cn.Open

End Function

Public Function insertDB()  'flash->VB->access
Dim sql
sql = "INSERT INTO schedule(��¥, �ð���, ��������, ���ӽð�) VALUES('" & _
            toDB��¥ & _
            "' , '" & _
            toDB��Ͻð� & _
            "' , '" & _
            toDB���� & _
            "' , '" & _
            toDB���ӽð� & _
            "') "

Debug.Print sql
On Error GoTo er
cn.Execute (sql)
Exit Function

er:
MsgBox "�����߻����� ó������ ������", , "ó������"

End Function

Private Sub Calendar1_Click()
    With Calendar1
        Firstday = DateSerial(.Year, .Month, .Day)
    End With
    Call scheduleView
End Sub

Public Function scheduleView()

'�ʱ�ȭ
sch_j = 0
ShockwaveFlash1.SetVariable "setaction.allclear", ""


For sch_j = 0 To 7  '��(Day)���� �ε�

    getday = Month(Firstday + sch_j) & "/" & Day(Firstday + sch_j) & "/" & Year(Firstday + sch_j)
    Set rs = cn.Execute("TRANSFORM Last(schedule.��������) AS [���������Ǹ����� ��] SELECT time.�ð��� FROM schedule INNER JOIN [time] ON schedule.�ð���=time.�ð��� Where (((schedule.��¥) = #" & getday & "#)) GROUP BY time.�ð��� PIVOT schedule.��¥;")
    ShockwaveFlash1.SetVariable "setaction.day" & sch_j, Day(Firstday + sch_j) & "��"

'On Error GoTo skips '������ ���³�..
If Not (rs.EOF) Then
    rs.MoveFirst
End If
'skips: 'on error ó��


While Not (rs.EOF)

sch_i = rs(0)
sch_data = rs(1)
    'vb -> flash
    With ShockwaveFlash1
        .SetVariable "setaction.day" & sch_j, Day(Firstday + sch_j) & "��"  '��� ��¥(��)ǥ��
        .SetVariable "setaction.data1", "schdata" & sch_j & "-" & sch_i     '��Ʈ������ ǥ����ġ ����
        .SetVariable "setaction.data2", sch_data                            '��Ʈ������ ������ ����
    End With
    
'Debug.Print "schdata" & sch_j & "-" & sch_i  '3-11"
'Debug.Print sch_data

rs.MoveNext
Wend


Next ' sch_j
End Function

'�÷��÷κ��� �����ͼ���
Private Sub ShockwaveFlash1_FSCommand(ByVal command As String, ByVal args As String)
Label1 = command
Label2 = args

Select Case command
Case "��¥"
    toDB��¥ = args

Case "��Ͻð�"
    toDB��Ͻð� = args
    
Case "����"
    toDB���� = args
    
Case "���ӽð�"
    toDB���ӽð� = args
    Call timeconvert
    Call insertDB
    Call scheduleView

Case "�Է�����"
    Call scheduleView

Case Else
    MsgBox "��Ӿȵȸ�ɹ߻� : " & command & " " & args
End Select

End Sub

Public Function timeconvert()

If Not (IsNumeric(toDB��Ͻð�)) Then
    MsgBox "�ð��� 24�ñ������� ���ڸ� �־�� �մϴ�.", "����"
    Exit Function
End If

If (5 <= toDB��Ͻð�) And (toDB��Ͻð� <= 24) Then
    toDB��Ͻð� = toDB��Ͻð� - 4
ElseIf (toDB��Ͻð� < 5) Then
    toDB��Ͻð� = toDB��Ͻð� + 20
Else
    MsgBox "�ð�������Դϴ�. 24���� ū���� �ȵǿ�.", "����"
End If

End Function


Private Sub Command1_Click() 'for test
With ShockwaveFlash1
    .SetVariable "setaction.day0", "4��" 'today
    .SetVariable "setaction.day1", "5��"
    .SetVariable "setaction.day2", "6��"
    .SetVariable "setaction.day3", "7��"
    .SetVariable "setaction.day4", "8��"
    .SetVariable "setaction.day5", "9��"
    .SetVariable "setaction.day6", "10��"
    .SetVariable "setaction.day7", "11��"
    .SetVariable "setaction.day8", "12��"
    
    .SetVariable "setaction.test1", "���"
    
    .SetVariable "setaction.data1", "schdata3-12"
    .SetVariable "setaction.data2", "������12312312��~"
    
    .SetVariable "setaction.data1", "schdata3-11"
    .SetVariable "setaction.data2", "������������������������������"
    
    
    
   ' "schdata"+j+"-"+i
End With


With ShockwaveFlash1
'ShockwaveFlash1.SetVariable "setaction.sDown", "��Ȫ" '������, ������
'.SetVariable ("_root.txt1.text","76");

End With

'ShockwaveFlash1.SetVariable "setaction.day1", "21��" '������, ������

End Sub
