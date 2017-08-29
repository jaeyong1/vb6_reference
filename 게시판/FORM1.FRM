VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "Bulletin Board Sample - Suport by DECA98"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   0
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "List View"
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   17
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton Command3 
         Caption         =   "삭제"
         Height          =   375
         Left            =   4080
         TabIndex        =   21
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "취소"
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   20
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "저장"
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   19
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   270
         IMEMode         =   3  '사용 못함
         Index           =   4
         Left            =   3360
         Locked          =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   4
         Tag             =   "PASS_F"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   2175
         Index           =   3
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   6
         Tag             =   "COMMENT_F"
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   2
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   5
         Tag             =   "SUB_F"
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   1
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   3
         Tag             =   "ID_F"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   0
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "EMAIL_F"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         Caption         =   "패스워드"
         Height          =   180
         Index           =   4
         Left            =   2460
         TabIndex        =   16
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         Caption         =   "내용"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         Caption         =   "제목"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         Caption         =   "아이디"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         Caption         =   "메일"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton Command4 
         Caption         =   ">"
         Height          =   300
         Index           =   1
         Left            =   3840
         TabIndex        =   23
         Top             =   3840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<"
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   3840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "신규"
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   3840
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H8000000B&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   480
         ScaleHeight     =   435
         ScaleWidth      =   3255
         TabIndex        =   15
         Top             =   3740
         Width           =   3255
         Begin VB.PictureBox Picture2 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  '없음
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   24
            Top             =   120
            Width           =   255
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "1"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   25
               Top             =   60
               Visible         =   0   'False
               Width           =   90
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5741
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         FixedCols       =   0
         BackColor       =   -2147483626
         BackColorFixed  =   -2147483645
         ForeColorSel    =   -2147483635
         BackColorBkg    =   -2147483633
         GridColor       =   -2147483644
         GridColorFixed  =   -2147483644
         AllowBigSelection=   -1  'True
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         Appearance      =   0
         FormatString    =   "<Number |<ID          |<Subject                                          |"
      End
      Begin VB.Label Label2 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         Caption         =   "총 아이템수 : 0개"
         Height          =   180
         Index           =   1
         Left            =   3675
         TabIndex        =   14
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Page : 1/1"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   885
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i        As Integer
Dim P        As Integer
Dim j        As String
Dim sql_TMP  As String
Dim PAGE_CNT As Integer
Dim PAGE_POS As Integer

Private Sub Command1_Click(Index As Integer)
  Select Case Index
    Case 0
       Select Case Command1(0).Caption
         Case "수정"
           j = InputBox("패스워드를 입력입력하십시요", "Password")
           If j <> "Supervisor" And j <> Text1(4).Text Then
             MsgBox "수정 권한이 없습니다", vbOKOnly, "Fault"
           Else
             Call Protect_RTN(True)
             Command1(0).Caption = "수정저장"
           End If
           Exit Sub
         Case "신규저장"
           Call NEW_INPUT
           Rs.Requery
           Call RE_FRE
         Case "수정저장"
           Call MODIFY_INPUT
           Rs.Requery
           Call RE_FRE
       End Select
    Case 3
      Unload Me
  End Select
  Call View_RTN(False)
End Sub

Private Sub Command2_Click()
  Call View_RTN(True)
  Command3.Enabled = False
  Command1(0).Caption = "신규저장"
  Call Protect_RTN(False)
  For i = 0 To 4
    Text1(i).Text = ""
  Next i
End Sub

Private Sub Command3_Click()
  j = InputBox("패스워드를 입력입력하십시요", "Password")
  If j <> "Supervisor" And j <> Text1(4).Text Then
    MsgBox "삭제 권한이 없습니다", vbOKOnly, "Fault"
    Exit Sub
  End If
  Rs_MOD.Close
  sql_TMP = "Delete * From BULLETIN_T Where NUM_F='" & Frame1(0).Tag & "' "
  For i = 0 To 2
    sql_TMP = sql_TMP & "AND " & Text1(i).Tag & "='" & Text1(i).Text & "'"
  Next i
  Db.Execute sql_TMP
  Rs.Requery
  Call DB_BULLSET(0)
  Call View_RTN(False)
End Sub

Private Sub Command4_Click(Index As Integer)
  Select Case Index
    Case 0
      Picture2.Left = Picture2.Left + 255
    Case 1
      Picture2.Left = Picture2.Left - 255
  End Select
End Sub

Private Sub Form_Load()
  PAGE_POS = 1000
  If Dir(App.Path & "\Data.mdb") <> "" Then
    sql_TMP = "select * from BULLETIN_T ORDER BY NUM_F DESC"
    Set Db = OpenDatabase(App.Path & "\Data.mdb", False, False, gsConnect)
    Set Rs = Db.OpenRecordset(sql_TMP, dbOpenDynaset, dbSeeChanges, dbPessimistic)
    Call RE_FRE
    USE_FLAG = True
  Else
    USE_FLAG = False
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If USE_FLAG = True Then
    Rs.Close
    Db.Close
  End If
End Sub

Private Sub Label3_Click(Index As Integer)
  For i = 1 To PAGE_CNT
    Label3(i).ForeColor = &H0&
  Next i
  Label3(Index).ForeColor = &HC000&
  PAGE_POS = Index
  Call DB_BULLSET(Index)
End Sub

Private Sub MSFlexGrid1_dblClick()
  If MSFlexGrid1.Row = 0 Then Exit Sub
  If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) = "" And MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = "" And MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) = "" Then Exit Sub
  sql_TMP = "Select * from BULLETIN_T where NUM_F='" & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) & "' AND ID_F = '" & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) & "'"
  Call View_RTN(True)
  Call Protect_RTN(True)
  Command3.Enabled = True
  Command1(0).Caption = "수정"
  Set Rs_MOD = Db.OpenRecordset(sql_TMP, dbOpenDynaset, dbSeeChanges, dbPessimistic)
  If Rs_MOD.EOF And Rs_MOD.BOF Then Exit Sub
  For i = 0 To 4
    If VarType(Rs_MOD(Text1(i).Tag)) <> vbNull Then
      Text1(i).Text = Rs_MOD(Text1(i).Tag)
    Else
      Text1(i).Text = ""
    End If
  Next i
  If VarType(Rs_MOD!NUM_F) <> vbNull Then
      Frame1(0).Tag = Rs_MOD!NUM_F
  Else
      Frame1(0).Tag = ""
  End If
  
End Sub

'######### here down : Call SUB_lib rootin #######

Sub View_RTN(B As Boolean)
  Frame1(0).Visible = B
  Frame1(1).Visible = Not (B)
  If B = False Then Command1(0).Caption = ""
End Sub

Sub Protect_RTN(B As Boolean)
  For i = 0 To 4
    Text1(i).Locked = B
  Next i
End Sub

Sub DB_BULLSET(K As Integer)
  Dim OLD_PGCNT As Integer
  If Rs.EOF And Rs.BOF Then Exit Sub
  Rs.MoveLast
  Rs.MoveFirst
  Label2(1).Caption = "총 아이템수 : " & Trim(Str(Rs.RecordCount - 1))
  OLD_PGCNT = PAGE_CNT
  PAGE_CNT = Rs.RecordCount / 20
  If (Rs.RecordCount / 20) > PAGE_CNT Then PAGE_CNT = PAGE_CNT + 1
'  PAGE_CNT = 30
  Label2(0).Tag = Trim(Str(PAGE_CNT))
  Picture2.Width = 255 * PAGE_CNT
  If OLD_PGCNT < PAGE_CNT Then
    For i = OLD_PGCNT + 1 To PAGE_CNT
      Load Label3(i)
      Label3(i).Caption = Trim(Str(i))
      Label3(i).Visible = True
      Label3(i).Move (i - 1) * 250, 60
    Next i
  ElseIf OLD_PGCNT > PAGE_CNT Then
    For i = OLD_PGCNT To PAGE_CNT + 1 Step -1
      Unload Label3(i)
    Next i
  End If
  If PAGE_CNT > 10 Then
    Command4(0).Visible = True
    Command4(1).Visible = True
  Else
    Command4(0).Visible = False
    Command4(1).Visible = False
  End If
  Label2(0).Caption = "Page : " & Trim(Str(K)) & "/" & Label2(0).Tag
  K = (K - 1) * 20
  If K < 0 Then
     PAGE_POS = 0
  Else
     PAGE_POS = K
  End If
  Rs.AbsolutePosition = PAGE_POS
  MSFlexGrid1.Clear
  MSFlexGrid1.FormatString = "<Number |<ID          |<Subject                                          |"
  For i = 1 To 20
    If VarType(Rs(Text1(1).Tag)) <> vbNull Then
      MSFlexGrid1.TextMatrix(i, 0) = Rs!NUM_F
    End If
    For P = 1 To 2
      If VarType(Rs(Text1(1).Tag)) <> vbNull Then
         MSFlexGrid1.TextMatrix(i, P) = Rs(Text1(P).Tag)
      Else
         MSFlexGrid1.TextMatrix(i, P) = ""
      End If
    Next P
    Rs.MoveNext
    If Rs.EOF Then Exit Sub
  Next i
End Sub

Sub MODIFY_INPUT()
  Rs_MOD.Edit
  For i = 0 To 4
    Rs_MOD(Text1(i).Tag) = Text1(i).Text
  Next i
  Rs_MOD.Update
  Rs_MOD.Close
End Sub

Sub NEW_INPUT()
  sql_TMP = "Insert INTO BULLETIN_T (" & Text1(0).Tag
  For i = 1 To 4
    sql_TMP = sql_TMP & ", " & Text1(i).Tag
  Next i
  sql_TMP = sql_TMP & ", NUM_F) VALUES ('" & Text1(0).Text
  For i = 1 To 4
    sql_TMP = sql_TMP & "', '" & Text1(i).Text
  Next i
  sql_TMP = sql_TMP & "', '" & Format$(Now, "yyyyMMddhhmmss") & "')"
  Db.Execute sql_TMP
End Sub

Sub RE_FRE()
  Call DB_BULLSET(0)
  For i = 2 To PAGE_CNT
    Label3(i).ForeColor = &H0&
  Next i
  Label3(1).ForeColor = &HC000&
End Sub
