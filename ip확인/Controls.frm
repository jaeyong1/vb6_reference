VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Controls 
   BackColor       =   &H00404080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "Controls.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3720
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "made by R3|K0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   2400
      Width           =   1140
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current ip address       :"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1635
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current machine name :"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1680
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Some Winsock controls."
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   2  'Horizontal Line
      Height          =   255
      Left            =   -120
      Top             =   0
      Width           =   4815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "Controls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*                    Winsock Controls made by [-=R3|K0=-]                      *
'*    This is a simple program for begginers that want to know some of the      *
'* mswinsock controls. Very easy to use and customize. For any problems or      *
'* suggestions e-mail me at : alex_tz@email.com or ICQ me at : 13244452         *
'*                              Thanks a lot                                    *
'*                                -=R3|K0=-                                     *
'********************************************************************************

Private Sub Form_Load()
Text1.Text = Winsock1.LocalHostName
Text2.Text = Winsock1.LocalIP
End Sub

Private Sub Label2_Click()
Me.WindowState = 1
End Sub

Private Sub Label3_Click()
End
End Sub
