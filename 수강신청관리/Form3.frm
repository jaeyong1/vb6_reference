VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7260
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   7140
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "숨기기"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   6600
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid gridData 
      Bindings        =   "Form3.frx":0000
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11245
      _Version        =   393216
      FixedCols       =   0
      FormatString    =   "이름|신청과목코드"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide


End Sub
