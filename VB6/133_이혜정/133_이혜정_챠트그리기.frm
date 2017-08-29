VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form5 
   Caption         =   "챠트 그리기"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form5"
   ScaleHeight     =   5490
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command2 
      Caption         =   "닫기"
      Height          =   495
      Left            =   4830
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "챠트"
      Height          =   495
      Left            =   1830
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "133_이혜정_챠트그리기.frx":0000
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "133_이혜정_챠트그리기.frx":0014
      TabIndex        =   0
      Top             =   600
      Width           =   7575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\VB6\133_이혜정\133_이혜정_학생관리.mdb"
      DefaultCursorType=   0  '기본 커서
      DefaultType     =   2  'ODBC사용
      Exclusive       =   0   'False
      Height          =   285
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '다이너셋
      RecordSource    =   "Student"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Form5.Hide
    Form6.Show
End Sub

Private Sub Command2_Click()
    Form5.Hide
    Form1.Show
End Sub

