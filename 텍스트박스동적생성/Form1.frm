VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.TextBox txtRe 
      BackColor       =   &H008080FF&
      Height          =   270
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 
 
Dim i, j, k As Integer

k = 1
For j = 1 To 20
For i = 1 To 5
    Load txtRe(k)

    txtRe(k).Visible = True
     txtRe(k).Left = (i * 600) + txtRe(0).Left
    txtRe(k).Top = j * 300 + txtRe(0).Top
    k = k + 1
     
Next
Next
     
End Sub

Private Sub txtRe_Click(Index As Integer)
If (txtRe(Index).BackColor = &H80FF80) Then
    txtRe(Index).BackColor = &H8080FF
    Else
    txtRe(Index).BackColor = &H80FF80
    End If
    
    



End Sub
