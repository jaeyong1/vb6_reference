VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form frm1 
   BackColor       =   &H00000000&
   ClientHeight    =   8760
   ClientLeft      =   1380
   ClientTop       =   1200
   ClientWidth     =   11895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   11895
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox backpic 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  '없음
      Height          =   8775
      Left            =   0
      ScaleHeight     =   8775
      ScaleWidth      =   12015
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ss 
         Height          =   6135
         Left            =   6960
         TabIndex        =   2
         Top             =   360
         Width           =   5415
         _cx             =   9551
         _cy             =   10821
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   0   'False
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   0   'False
         Base            =   ""
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "X"
         Height          =   375
         Left            =   11400
         TabIndex        =   1
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Image1 
         Appearance      =   0  '평면
         DragMode        =   1  '자동
         Height          =   2490
         Left            =   4560
         Top             =   3480
         Width           =   3240
      End
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Activate()
ss.Movie = "c:\flip.swf"
ss.Play
End Sub

Private Sub ss_OnReadyStateChange(newState As Long)
Image1.Picture = LoadPicture("c:\jaeyong.jpg")

End Sub
