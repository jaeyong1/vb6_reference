VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl UserControl1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0FF&
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9075
   ScaleHeight     =   4425
   ScaleWidth      =   9075
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Frame1"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   5280
         TabIndex        =   4
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command1 
         Caption         =   "test"
         Height          =   615
         Left            =   5520
         TabIndex        =   3
         Top             =   2520
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   240
         Top             =   1920
      End
      Begin MSACAL.Calendar Calendar2 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   1935
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   2295
         _Version        =   524288
         _ExtentX        =   4048
         _ExtentY        =   3413
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2006
         Month           =   11
         Day             =   22
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   0   'False
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   1095
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   1215
         _Version        =   524288
         _ExtentX        =   2143
         _ExtentY        =   1931
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2006
         Month           =   11
         Day             =   22
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'ÇÁ·ÎÁ§Æ®->ÂüÁ¶->Microsoft ActiveX Data Object 2.5 Library
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Calendar2_Click()
MsgBox Calendar2.Value

End Sub

Private Sub Command1_Click()
UserControl.Height = 0
UserControl.Width = 0



End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = 99 Then
ProgressBar1.Value = 0
End If


End Sub

'html ÅÂ±×¿¡..
'<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0">

Private Sub UserControl_Initialize()
 
Frame1.Left = 0 'Screen.Width - Frame1.Width
Frame1.Top = 0

With UserControl

'.Height = Screen.Height
'.Width = Screen.Width
'.Picture = "c:\\1.bmp"

End With
'MsgBox App.Path


End Sub

