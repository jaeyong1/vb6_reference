VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frmprt 
   Caption         =   "인쇄 관리자"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8955
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   518
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   597
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CheckBox chkPnt4 
      Caption         =   "인증문서"
      Height          =   255
      Left            =   2400
      TabIndex        =   28
      Top             =   7320
      Value           =   1  '확인
      Width           =   1695
   End
   Begin VB.CheckBox chkPnt3 
      Caption         =   "평가결과"
      Height          =   255
      Left            =   2400
      TabIndex        =   27
      Top             =   6960
      Value           =   1  '확인
      Width           =   1455
   End
   Begin VB.CheckBox chkPnt2 
      Caption         =   "시험결과"
      Height          =   255
      Left            =   600
      TabIndex        =   26
      Top             =   7320
      Value           =   1  '확인
      Width           =   1215
   End
   Begin VB.CheckBox chkPnt1 
      Caption         =   "표지"
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   6960
      UseMaskColor    =   -1  'True
      Value           =   1  '확인
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   5775
      Left            =   4800
      TabIndex        =   23
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "출  력"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   22
      Top             =   6960
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker DTReqDay 
      Height          =   375
      Left            =   1680
      TabIndex        =   19
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      Format          =   25493505
      CurrentDate     =   40016
   End
   Begin VB.TextBox Txtprog 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   18
      Top             =   6240
      Width           =   2775
   End
   Begin VB.TextBox TxtModelnum 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Text            =   "PLC-AAA-BB001"
      Top             =   5640
      Width           =   2775
   End
   Begin VB.TextBox TxtMaker 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Text            =   "홍길동"
      Top             =   5040
      Width           =   2775
   End
   Begin VB.TextBox TxtReqA 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Text            =   "한국정보통신(주)"
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox TxtClk 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Text            =   "시험자"
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox TxtUDay 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Text            =   "발행일로부터 6개월"
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox TxtAsso 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Text            =   "한국전기연구원"
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox TxtSpecType 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Text            =   "KS X 4600-1 (Class-B)"
      Top             =   240
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTFinDay 
      Height          =   375
      Left            =   1680
      TabIndex        =   20
      Top             =   1440
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      Format          =   25493505
      CurrentDate     =   40016
   End
   Begin MSComCtl2.DTPicker DTPrtDay 
      Height          =   375
      Left            =   1680
      TabIndex        =   21
      Top             =   2040
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      Format          =   25493505
      CurrentDate     =   40016
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "결 과"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6360
      TabIndex        =   24
      Top             =   240
      Width           =   525
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "진행상태"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   10
      Top             =   6240
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "DUT 모델명"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   9
      Top             =   5640
      Width           =   1170
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "DUT 생산자"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   1170
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "시험신청기관"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   1350
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "유효기간"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "시험자"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "시험기관"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "발 행 일"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "시험완료일"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "신 청 일"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "적용 규격"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmprt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlApp As New Excel.Application '프로젝트메뉴-참조 - Microsoft Excel 12.0 Object Library  (2007 기준, 이하버젼은 11 10..등 참조)

Private Sub CheckBox1_Click()

End Sub

Private Sub CheckBox2_Click()

End Sub

Private Sub Command1_Click()
  Const XL_NOTRUNNING As Long = 429 '엑셀이 기존에 실행되고 있지 않으면 429 에러가 발생
  On Error GoTo ShowName_Err '에러가 발생하면(엑셀이 기존에 실행되고 있지 않다면) ShowName_Err 문으로 이동
  Set xlApp = GetObject(, "Excel.Application") '엑셀이 실행되고 있나 체크
    
    
        
  With xlApp  '엑셀 내 작업
        .Visible = True '엑셀 표시
        '.Visible = False '엑셀 표시 안함
        
        .DisplayAlerts = False '경고 무시
        .Workbooks.Open App.Path & "\cert.xlsx"  '인증양식 문서 열기
    
    If chkPnt1.Value = 1 Then
        '#### 표지 인쇄 ####
        .Sheets("표지").Select
        .Range("e7").Select: .ActiveCell.FormulaR1C1 = "0000-0000-0000" ' 문서번호
        .Range("e9").Select: .ActiveCell.FormulaR1C1 = TxtSpecType.Text  ' 적용규격
        
        .Range("e14").Select: .ActiveCell.FormulaR1C1 = TxtReqA.Text  ' 시험신청기관
        .Range("e15").Select: .ActiveCell.FormulaR1C1 = TxtMaker.Text  ' DUT생산자
        .Range("e16").Select: .ActiveCell.FormulaR1C1 = TxtModelnum.Text  ' DUT모델명
        
        .Range("d19").Select: .ActiveCell.FormulaR1C1 = DTReqDay.Value  ' 신청일
        .Range("d20").Select: .ActiveCell.FormulaR1C1 = DTFinDay.Value  ' 시험완료일
        .Range("d21").Select: .ActiveCell.FormulaR1C1 = DTPrtDay.Value  ' 발행일
        
        .Range("f19").Select: .ActiveCell.FormulaR1C1 = TxtAsso.Text  ' 시험기관
        .Range("f20").Select: .ActiveCell.FormulaR1C1 = TxtClk.Text  ' 시험자
        .Range("f21").Select: .ActiveCell.FormulaR1C1 = TxtUDay.Text  ' 유효기간
        
        .ActiveWindow.SelectedSheets.PrintOut Copies:=1 '인쇄
    End If
    
    If chkPnt2.Value = 1 Then
        '#### 시험결과 인쇄 ####
         .Sheets("시험결과").Select
    
    
    
    End If
    
    
    If chkPnt3.Value = 1 Then
        '#### 평가결과 인쇄 ####
         .Sheets("평가결과").Select
    
    
    
    End If
    
    
    If chkPnt4.Value = 1 Then
        '#### 인증문서 인쇄 ####
         .Sheets("인증문서").Select
    
    
    
    End If
        
  End With
    
    
'
'    '플랙스 그리드의 자료를 엑셀로 복사한다.
'    For iRow = 0 To VSFlexGrid1.Rows - 1
'        For iCol = 0 To VSFlexGrid1.Cols - 1
'            oExcel.Worksheets(1).Cells(iRow, iCol).Value = MSFlexGrid1.TextMatrix(iRow, iCol)
'        Next
'    Next
'
'    '엑셀 파일로 저장한다.
'    oExcel.Worksheets(1).SaveAs "C:\test.xls"
'    'sPath = "http://" & window.location.host & "\eMES\reports\prodt\prd600p.xls"
'    'oExcel.Worksheets(1).SaveAs "F:\test.xls"
'    '대화형 모드로 전환합니다.
  '  oExcel.Interactive = True
'
'
'
'    With xlApp
'
'
'    .Range("C3").Select :     .ActiveCell.FormulaR1C1 = "1"
'    .Range("C4").Select
'    .ActiveCell.FormulaR1C1 = "2"
'    .Range("C5").Select
'    .ActiveCell.FormulaR1C1 = "3"
'    .Range("C6").Select
'    .ActiveCell.FormulaR1C1 = "4"
'    .Range("C7").Select
'    .ActiveWindow.SelectedSheets.PrintOut Copies:=1
'
'
'    '   엑셀로 바로 출력
'End With
'    '엑셀 개체를 닫습니다.
'    If Not (oExcel Is Nothing) Then
'        Set oExcel = Nothing
'    End If
''
'      '플랙스 그리드의 자료를 엑셀로 복사한다.
'    For iRow = 0 To VSFlexGrid1.Rows - 1
'        For iCol = 0 To VSFlexGrid1.Cols - 1
'            oExcel.Worksheets(1).Cells(iRow + 1, iCol + 1).Value = VSFlexGrid1.TextMatrix(iRow, iCol)
'        Next
'    Next
    
    
    
    
          
    


'xlApp.Quit '엑셀 프로그램 종료
'Set xlApp = Nothing '엑셀 어플리케이션 개체 메모리에서 제거
Exit Sub

''''''''''' 에러처리
ShowName_End:
    Exit Sub
ShowName_Err:
    If Err = XL_NOTRUNNING Then '엑셀이 실행중이지 않은 경우
        Set xlApp = New Excel.Application '엑셀 실행
        xlApp.Workbooks.Add '워크북 추가
        Resume Next '에러 다음 발생 위치(GetObject 문 뒤)로 복귀
    Else
        MsgBox Err.Number & " - " & Err.Description '그렇지 않은 에러가 발생하면 에러 번호 및 에러 내용 표시
    End If
    Resume ShowName_End '프로시저를 끝냄
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
 '
End Sub
