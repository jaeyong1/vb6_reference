VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
      Dim intFnum As Integer

      ' Open pjy.bat.
      intFnum = FreeFile
      Open "C:\pjy.bat" For Append As intFnum
    
      ' ���ϴ� ���ڸ� �����δ�
      Print #intFnum, Text1.Text
      
    
      ' Close Auotexec.bat
      Close intFnum
    
      MsgBox "���� !! Autoexec.bat ������ üũ�غ��ʽÿ�"
  End Sub

