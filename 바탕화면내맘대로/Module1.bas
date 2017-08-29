Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SystemParametersInfo Lib "user32" _
                 Alias "SystemParametersInfoA" _
                       (ByVal uAction As Long, _
                        ByVal uParam As Long, _
                        ByVal lpvParam As Any, _
                        ByVal fuWinIni As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetShortPathName Lib "kernel32" _
                Alias "GetShortPathNameA" _
                (ByVal lpszLongPath As String, _
                ByVal lpszShortPath As String, _
                ByVal cchBuffer As Long) As Long

Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_SENDWININICHANGE = &H2
Public Const SPIF_UPDATEINIFILE = &H1

Public g_File As String
Public convertFile As String
Public TitleStyle As String
Public PaperStyle As String
Public Const RegPath = "Control Panel\Desktop"

' 바탕화면 이미지를 바꾸는 함수를 만듬 ^^ (넘 간단)

Public Sub SetWallpaper(ByVal strFile As String)

    Dim x As Long

    x = SystemParametersInfo(SPI_SETDESKWALLPAPER, _
        0&, strFile, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE)

End Sub

'''HKEY_CURRENT_USER\Control Panel\Desktop\WallpaperStyle

'''HKEY_CURRENT_USER\Control Panel\Desktop\TileWallpaper

'''        TileWallpaper WallpaperStyle
'''가운데          0           0
'''바둑판식        1           0
'''늘이기          0           2

