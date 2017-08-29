VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   LinkTopic       =   "Form2"
   ScaleHeight     =   9450
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Image Image1 
      Height          =   9345
      Left            =   30
      Top             =   30
      Width           =   10935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strName As String

Private Sub Form_Activate()
    Dim i As Long
    
    If g_File = "" Then Exit Sub
    
'    i = InStrRev(g_File, ".")
'    strName = Left(g_File, i) & "bmp"   '저장할 파일명 생성
    
    Image1.Stretch = False  '이미지 크기에 맞게 Image1의 크기가 변하도록 설정
    Image1.Picture = LoadPicture(g_File)
    
    Me.Height = Image1.Height + 550     'Image1의 크기보다 폼의 크기가 항상 커야만 제대로 저장됨
    Me.Width = Image1.Width + 180
    
    strName = "C:\WINDOWS\Wallpaper1.bmp"   'WINDOWS폴더가 있다고 가정하에(없으면 에러남)
    
    '생성할 파일명으로 파일이 있으면 삭제함
    If Dir(strName) <> vbNullString Then Kill strName
    
    SavePicture Image1, strName     'savepicture로 저장하면 jpg나 gif는 모두 bmp로 저장됨
    
    i = 0   '변수 초기화
    
ReChk:
'############ 파일을 저장하면 하드디스크에 파일이 생성되기 이전에 아래루틴이 먼저 수행되면 에러발생함 ###########
'           에러를 방지하기 위해서 프로세스를 순간 멈추고 파일이 생성될때까지 무한루프...
    Sleep 100   '0.1초간 프로세스 정지
    
    If Dir(strName) = vbNullString Then     '새로 생성한 bmp파일이 생성되었는지를 검사함
        If i >= 300 Then    '일정시간동안 파일생성이 않될경우 무한루프돌지 않도록 빠져나감
            MsgBox "파일저장 실패"
        Else
            i = i + 1
            GoTo ReChk
        End If
    Else
    '### 아래 레지스트리값("OriginalWallpaper")과 "Wallpaper"레지스트리값의 경로가 동일하면
    '폼1에서 저장한 "ConvertedWallpaper"값의 경로가 디스플레이 등록정보에서 표시됨
    '다를경우는 "Wallpaper"값을 표시
        Call UpdateKey(HKEY_CURRENT_USER, RegPath, "OriginalWallpaper", strName)
        Call SetWallpaper(strName)    '파일이 생성됬으면 바로 바탕화면 변경
    End If
    
    Unload Me   '저장후 폼을 닫음
End Sub

Private Sub Form_Load()
'###### 폼2가 보이지 않도록 멀리~ 아주멀~리...날려버림...ㅋ########
    Me.Left = -99999
    Me.Top = -99999
End Sub
