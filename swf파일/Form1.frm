VERSION 5.00
Object = "{166B1BC7-3F9C-11CF-8075-444553540000}#1.0#0"; "SWDIR.DLL"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   5280
      Width           =   735
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   3495
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   6615
      _cx             =   11668
      _cy             =   6165
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      Stacking        =   "below"
   End
   Begin DIRECTORSHOCKWAVELibCtl.ShockwaveCtl ShockwaveCtl1 
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   2295
      _cx             =   4048
      _cy             =   2143
      BGCOLOR         =   ""
      swURL           =   ""
      swText          =   ""
      swForeColor     =   ""
      swBackColor     =   ""
      swFrame         =   ""
      swColor         =   ""
      swName          =   ""
      swPassword      =   ""
      swBanner        =   ""
      swSound         =   ""
      swVolume        =   ""
      swPreloadTime   =   ""
      swAudio         =   ""
      swList          =   ""
      sw1             =   ""
      sw2             =   ""
      sw3             =   ""
      sw4             =   ""
      sw5             =   ""
      sw6             =   ""
      sw7             =   ""
      sw8             =   ""
      sw9             =   ""
      SRC             =   ""
      AutoStart       =   "TRUE"
      Sound           =   "TRUE"
      swRemote        =   ""
      logo            =   "TRUE"
      progress        =   "TRUE"
      PowerMenuEnabled=   "TRUE"
      swModifyReport  =   "FALSE"
      swClickThroughUrl=   ""
      swStretchStyle  =   "stage"
      swStretchHAlign =   "center"
      swStretchVAlign =   "center"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Integer
Private Sub Command1_Click()
I = I + 1
ShockwaveFlash1.Movie = "C:\지하탱크.swf"
ShockwaveFlash1.Play

ShockwaveFlash1.GotoFrame (I)

If I = 4 Then
I = 0
End If
End Sub

Private Sub Form_Load()
ShockwaveFlash1.Movie = "C:\지하탱크.swf"
ShockwaveFlash1.Play
ShockwaveFlash1.GotoFrame (I)
End Sub
