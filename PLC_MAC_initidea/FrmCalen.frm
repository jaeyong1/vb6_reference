VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCalen 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   2685
   ClientLeft      =   2760
   ClientTop       =   3360
   ClientWidth     =   2325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "확인"
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
   Begin MSComCtl2.MonthView MonthView1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   3
      EndProperty
      Height          =   2220
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   76087297
      CurrentDate     =   40016
      MaxDate         =   401768
      MinDate         =   36161
   End
End
Attribute VB_Name = "FrmCalen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
 

Private Sub CancelButton_Click()
    'ESC키 누르면 그냥 종료
    Unload Me
End Sub

Private Sub MonthView1_DateDblClick(ByVal DateDblClicked As Date)
    '날짜 더블클릭 -> 날짜선택, 달력감춤
    OKButton_Click
End Sub

Private Sub OKButton_Click()
    Cal_YYYY = MonthView1.Year
    Cal_MM = MonthView1.Month
    Cal_DD = MonthView1.Day
    YYYYMMDD = MonthView1.Value
    Unload Me
End Sub
