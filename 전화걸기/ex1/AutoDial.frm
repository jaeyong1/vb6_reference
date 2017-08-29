VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auto-Dial"
   ClientHeight    =   1830
   ClientLeft      =   3765
   ClientTop       =   3720
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1830
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdDial 
      Caption         =   "&Dial..."
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtNumber 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "&Phone Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AutoDial - Telephone dialer demo program
'Copyright (c) 1996-97 SoftCircuits
'Redistributed by Permission.
'
'This Visual Basic 5.0 example program demonstrates how an application
'can dial a telephone number under Windows 95 using Assisted Telephony
'which is a subset of TAPI. This code is simple because it relies on a
'call manager applet to perform the actual dialing.
'
'This program may be distributed on the condition that it is
'distributed in full and unchanged, and that no fee is charged for
'such distribution with the exception of reasonable shipping and media
'charged. In addition, the code in this program may be incorporated
'into your own programs and the resulting programs may be distributed
'without payment of royalties.
'
'This example program was provided by:
' SoftCircuits Programming
' http://www.softcircuits.com
' P.O.Box 16262
' Irvine, CA 92623
Option Explicit

Private Declare Function tapiRequestMakeCall& Lib "TAPI32.DLL" (ByVal DestAddress$, ByVal AppName$, ByVal CalledParty$, ByVal Comment$)
Private Const TAPIERR_NOREQUESTRECIPIENT = -2&
Private Const TAPIERR_REQUESTQUEUEFULL = -3&
Private Const TAPIERR_INVALDESTADDRESS = -4&

Private Sub Form_Load()
    EnableDial
End Sub

Private Sub txtNumber_Change()
    EnableDial
End Sub

Private Sub cmdDial_Click()
    Dim buff As String
    Dim nResult As Long

    'Invoke tapiRequestMakeCall. If tapiRequestMakeCall returns 0, the
    'request has been accepted. It is up to the call manager application
    'to do any further work. The second-to-last argument should be
    'changed to be the name of the person you are dialing.
    nResult = tapiRequestMakeCall&(Trim$(txtNumber), CStr(Caption), "Test Dial", "")
    'Display message if error
    If nResult <> 0 Then
        buff = "Error dialing number : "
        Select Case nResult
            Case TAPIERR_NOREQUESTRECIPIENT
                buff = buff & "No Windows Telephony dialing application is running and none could be started."
            Case TAPIERR_REQUESTQUEUEFULL
                buff = buff & "The queue of pending Windows Telephony dialing requests is full."
            Case TAPIERR_INVALDESTADDRESS
                buff = buff & "The phone number is not valid."
            Case Else
                buff = buff & "Unknown error."
        End Select
        MsgBox buff
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub EnableDial()
    cmdDial.Enabled = Len(Trim$(txtNumber)) > 0
End Sub
