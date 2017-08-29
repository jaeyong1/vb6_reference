VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim oExcel    As Object
    Dim sPath
      On Error Resume Next
        
    '현재 엑셀이 실행중인지를 확인한다.
     Set oExcel = GetObject(, "Excel.Application")
    If Err = 429 Then
        Err = 0
      '엑셀이 실행중이지 않다면 엑셀을 실행한다.
        Set oExcel = CreateObject("Excel.Application")
        If Err = 429 Then
            MsgBox Err & ": " & Error, vbExclamation + vbOKOnly
            Exit Sub
        End If
    End If
    
    
    
    '새로운 엑셀문서를 만든다.
     oExcel.Workbooks.Add
        
    '시트 하나만 남기고 모두 삭제한다.
     oExcel.DisplayAlerts = False
     
         Workbooks.Open FileName:= _
        "C:\Documents and Settings\jaeyong\My Documents\ELL.xlsx"
    
    iSheetCount = oExcel.Worksheets.Count
    For i = 2 To iSheetCount
        oExcel.Worksheets(1).Delete
    Next
    oExcel.DisplayAlerts = True
    
    oExcel.Worksheets(1).Name = Null
    oExcel.Visible = True
    
    '플랙스 그리드의 자료를 엑셀로 복사한다.
    For iRow = 0 To VSFlexGrid1.Rows - 1
        For iCol = 0 To VSFlexGrid1.Cols - 1
            oExcel.Worksheets(1).Cells(iRow, iCol).Value = MSFlexGrid1.TextMatrix(iRow, iCol)
        Next
    Next
    
    '엑셀 파일로 저장한다.
    oExcel.Worksheets(1).SaveAs "C:\test.xls"
    'sPath = "http://" & window.location.host & "\eMES\reports\prodt\prd600p.xls"
    'oExcel.Worksheets(1).SaveAs "F:\test.xls"
    '대화형 모드로 전환합니다.
    oExcel.Interactive = True
    


    With oExcel
            
            
    .Range("C3").Select
    .ActiveCell.FormulaR1C1 = "1"
    .Range("C4").Select
    .ActiveCell.FormulaR1C1 = "2"
    .Range("C5").Select
    .ActiveCell.FormulaR1C1 = "3"
    .Range("C6").Select
    .ActiveCell.FormulaR1C1 = "4"
    .Range("C7").Select
    .ActiveWindow.SelectedSheets.PrintOut Copies:=1
            
    
    '   엑셀로 바로 출력
End With
    '엑셀 개체를 닫습니다.
    If Not (oExcel Is Nothing) Then
        Set oExcel = Nothing
    End If
'
      '플랙스 그리드의 자료를 엑셀로 복사한다.
    For iRow = 0 To VSFlexGrid1.Rows - 1
        For iCol = 0 To VSFlexGrid1.Cols - 1
            oExcel.Worksheets(1).Cells(iRow + 1, iCol + 1).Value = VSFlexGrid1.TextMatrix(iRow, iCol)
        Next
    Next

  End Sub
