VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "133_������_Main ȭ��"
   ClientHeight    =   5490
   ClientLeft      =   1905
   ClientTop       =   1740
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7875
   Begin VB.CommandButton Command5 
      Caption         =   "íƮ�׸���"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   2880
      Style           =   1  '�׷���
      TabIndex        =   4
      Top             =   3390
      Width           =   2100
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   2887
      Style           =   1  '�׷���
      TabIndex        =   3
      Top             =   4440
      Width           =   2100
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ü����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   2887
      Style           =   1  '�׷���
      TabIndex        =   2
      Top             =   2340
      Width           =   2100
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   2887
      Style           =   1  '�׷���
      TabIndex        =   1
      Top             =   1290
      Width           =   2100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�л�����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   2887
      Style           =   1  '�׷���
      TabIndex        =   0
      Top             =   240
      Width           =   2100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub Command2_Click()
    Form3.Show
End Sub

Private Sub Command3_Click()
    Form4.Show
End Sub

Private Sub Command4_Click()
    Dim res
    
    res = MsgBox("���� �����Ͻðڽ��ϱ�?", vbOKCancel + _
        vbInformation, "�޽�������")
    If res = vbOK Then
        End
    End If
End Sub

Private Sub Command5_Click()
    Form5.DBGrid1.Refresh
    Form1.Hide
    Form5.Show
End Sub
