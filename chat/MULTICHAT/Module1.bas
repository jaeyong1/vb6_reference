Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function ExitWindowsEx Lib "user32" ( _
    ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1

Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" ( _
    ByVal hwnd As Long, ByVal szApp As String, _
    ByVal szOtherStuff As String, ByVal hIcon As Long _
    ) As Long


