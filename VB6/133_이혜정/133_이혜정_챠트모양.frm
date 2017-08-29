VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form Form6 
   Caption         =   "챠트 모양"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form6"
   ScaleHeight     =   5490
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command5 
      Caption         =   "닫기"
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "2차원 파이"
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   3630
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "2차원 콤비"
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   2460
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "3차원 선형"
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   1290
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3차원 막대"
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Bindings        =   "133_이혜정_챠트모양.frx":0000
      Height          =   5175
      Left            =   240
      OleObjectBlob   =   "133_이혜정_챠트모양.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MSChart1.chartType = VtChChartType3dBar
End Sub

Private Sub Command2_Click()
    MSChart1.chartType = VtChChartType3dLine
End Sub

Private Sub Command3_Click()
    MSChart1.chartType = VtChChartType2dXY
End Sub

Private Sub Command4_Click()
    MSChart1.chartType = VtChChartType2dPie
End Sub

Private Sub Command5_Click()
    Form6.Hide
    Form1.Show
End Sub

Private Sub Form_Load()
    Dim i, j
    MSChart1.chartType = VtChChartType2dBar
    
    MSChart1.ColumnCount = 2
   
'    MSChart1.RowCount = 2
    MSChart1.RowCount = Form5.DBGrid1.ApproxCount
    
    Form5.DBGrid1.Col = 1
    For i = 0 To MSChart1.RowCount - 1
        Form5.DBGrid1.Row = i
        MSChart1.Row = i + 1
        MSChart1.RowLabel = Form5.DBGrid1.Text
    Next i
    
    For i = 1 To MSChart1.RowCount
        Form5.DBGrid1.Row = i - 1
         For j = 1 To 2
            MSChart1.Column = j
            MSChart1.Row = i
            Form5.DBGrid1.Col = j + 6
            MSChart1.Data = Form5.DBGrid1.Text
         Next j
      Next i
    
End Sub
