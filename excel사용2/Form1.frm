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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'사전 바인딩 방법(엑셀을 참조해서 실행시키는 방법)
'먼저 VB에서 프로젝트 메뉴-참조를  - Microsoft Excel 11.0 Object Library 를
'체크하고 확인을 눌러서 엑셀 개체를 참조합니다.
'
'

Option Explicit
 
Dim xlApp As New Excel.Application
 
Sub Command1_Click()
MsgBox App.Path

    Const XL_NOTRUNNING As Long = 429 '엑셀이 기존에 실행되고 있지 않으면 429 에러가 발생
 
    On Error GoTo ShowName_Err '에러가 발생하면(엑셀이 기존에 실행되고 있지 않다면) ShowName_Err 문으로 이동
    Set xlApp = GetObject(, "Excel.Application") '엑셀이 실행되고 있나 체크
    xlApp.Visible = True '엑셀 표시
    
    xlApp.DisplayAlerts = False
    
    
    xlApp.Workbooks.Open "C:\test.xls" '문서 열기, 닫기, 저장
    
    
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

Private Sub Command2_Click()
xlApp.Quit '엑셀 프로그램 종료
Set xlApp = Nothing '엑셀 어플리케이션 개체 메모리에서 제거
End Sub
