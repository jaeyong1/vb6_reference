VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   BorderStyle     =   1  '단일 고정
   Caption         =   "바탕화면 바꾸기"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   6210
   StartUpPosition =   2  '화면 가운데
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   4530
      TabIndex        =   6
      Top             =   4050
      Width           =   1545
   End
   Begin VB.ListBox List2 
      Appearance      =   0  '평면
      Height          =   2190
      Left            =   90
      TabIndex        =   5
      Top             =   3630
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Caption         =   "바탕화면 바꾸기"
      Height          =   495
      Left            =   4500
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   5340
      Width           =   1605
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Caption         =   "그림파일 검색"
      Height          =   495
      Left            =   4500
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   4650
      Width           =   1605
   End
   Begin VB.ListBox List1 
      Appearance      =   0  '평면
      Height          =   2190
      Left            =   150
      TabIndex        =   0
      Top             =   6420
      Width           =   5925
   End
   Begin VB.Label Label3 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "위치 :"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4500
      TabIndex        =   7
      Top             =   3750
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   1950
      Top             =   660
      Width           =   2325
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "2222222222222222222222"
      Height          =   180
      Left            =   1470
      TabIndex        =   4
      Top             =   6240
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "111111111111111111111111"
      Height          =   180
      Left            =   90
      TabIndex        =   2
      Top             =   3330
      Width           =   2160
   End
   Begin VB.Image Image2 
      Height          =   2580
      Left            =   1710
      Picture         =   "Form1.frx":0000
      Top             =   330
      Width           =   2835
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const strFolder = "C:\WINDOWS"     '읽어들일 폴더경로
Const Check_Name = "JPG,JPGE,GIF,BMP"           '읽어들일 파일 확장자 목록 ","로 구분
Dim strFile As String
Dim Cnt As Long

Private Sub Command2_Click()
    Call Set_List
End Sub

'List Box에 파일명을 읽어들임
Private Sub Set_List()

    Cnt = 0
    
    Command1.Enabled = False
    Command2.Enabled = False
    
    Label2.Caption = "그림 파일을 검색중 입니다..."
    List1.Clear
    
    Call GetFolderList(strFolder)   '하위폴더까지 검색하기 위해 재귀함수 호출
    
    Label1.Caption = "총 " & Cnt & "개의 그림파일을 찾았습니다."
    Label2.Caption = ""
    Command1.Enabled = True
    Command2.Enabled = True
    Exit Sub

Err:
    MsgBox Err.Number & "; " & Err.Description
End Sub

Private Sub GetFolderList(folderspec)
    Dim temp As String * 555
    Dim strName As String
    Dim fs, f, f1, s, sf, fCnt, f_File
    Dim fPath As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set sf = f.SubFolders
    Set fCnt = f.Files
    
    fPath = folderspec
    
    '파일의 경로가 긴경우 짧은 경로명으로 변경해줌
    strName = GetShortPathName(fPath, temp, 254)
    strName = Left$(temp, strName)
    Label1.Caption = "진행폴더: " & strName
    
    For Each f_File In fCnt
        DoEvents
        If Check_FileName(f_File.Name) = False Then GoTo Skip_Next  '읽어들일수 있는 그림파일인지 확인하기 위해 함수호출
        
        List1.AddItem fPath & "\" & f_File.Name     'list1에는 파일의 전체경로를 저장하고
        List2.AddItem f_File.Name       'list2에는 파일명만 저장하도록
        Cnt = Cnt + 1
Skip_Next:
    Next
    
    For Each f1 In sf       '하위폴더가 있다면 재귀호출...
        Call GetFolderList(fPath & "\" & f1.Name)
    Next
    
End Sub


Private Sub Form_Load()
    Combo1.AddItem "가운데"
    Combo1.AddItem "바둑판식"
    Combo1.AddItem "늘이기"
    
    Label1.Caption = ""
    Label2.Caption = ""
    Command1.Enabled = False
    
    g_File = GetKeyValue(HKEY_CURRENT_USER, RegPath, "Wallpaper")   '현재 저장된 바탕화면 그림을 읽어들임
    
    Image1.Stretch = True   'Image1에 그림크기를 마춤
    Image1.Picture = LoadPicture(g_File)
    
    '### 위치(가운데,바둑판식,늘이기) 정보를 가지고 있는 레지스트리 값을 읽어서 표시함 ###
    TitleStyle = GetKeyValue(HKEY_CURRENT_USER, RegPath, "TileWallpaper")
    PaperStyle = GetKeyValue(HKEY_CURRENT_USER, RegPath, "WallpaperStyle")
    
'    MsgBox GetKeyValue(HKEY_CURRENT_USER, RegPath, "ConvertedWallpaper Last WriteTime")
    
    If TitleStyle = "0" And PaperStyle = "0" Then   '가운데
        Combo1.ListIndex = 0
    ElseIf TitleStyle = "1" And PaperStyle = "0" Then   '바둑판식
        Combo1.ListIndex = 1
    ElseIf TitleStyle = "0" And PaperStyle = "2" Then   '늘이기
        Combo1.ListIndex = 2
    Else    '에러??
        
    End If
End Sub


'해당 파일이 그림파일인지 확인하는 함수
Private Function Check_FileName(ByVal strName As String) As Boolean
    Dim i As Integer        '루프를 위한 변수선언
    Dim strTemp As String   '임시로 사용할 변수
    Dim ArrName() As String '정의된 확장자들을 담을 배열변수
    
    ArrName() = Split(Check_Name, ",")  '정의된 확장자들을 ","를 구분으로 배열에 저장
    
    i = InStrRev(strName, ".")      '파일명의 뒤에서부터 "."를 찾아음
    strTemp = Mid(strName, i + 1)   '확장자명을 임시변수에 저장
    
    For i = 0 To UBound(ArrName())  '정의된 확장자의 갯수만큼 루프
        If UCase(strTemp) = UCase(ArrName(i)) Then  '대문자로 변경하여 비교
            Check_FileName = True   '일치하는 확장자가 있으면 빠져나감
            Exit Function
        End If
    Next
    
End Function

Private Sub Command1_Click()
    Dim i As Integer

    If strFile = "" Then
        MsgBox "파일을 선택하셈         ", vbInformation, "에러"
        Exit Sub
    End If
    
    i = InStrRev(strFile, ".")
    
    '################ 위치(가운데,바둑판식,늘이기) 확인후 기록 ############
    If Combo1.ListIndex = 0 Then    '가운데
        TitleStyle = "0"
        PaperStyle = "0"
    ElseIf Combo1.ListIndex = 1 Then    '바둑판식
        TitleStyle = "1"
        PaperStyle = "0"
    ElseIf Combo1.ListIndex = 2 Then    '늘이기
        TitleStyle = "0"
        PaperStyle = "2"
    End If
    '레지스트리값을 미리 변경해야만 적용됨
    Call UpdateKey(HKEY_CURRENT_USER, RegPath, "TileWallpaper", TitleStyle)
    Call UpdateKey(HKEY_CURRENT_USER, RegPath, "WallpaperStyle", PaperStyle)
    
    If UCase(Mid(strFile, i + 1)) = "BMP" Then  'bmp파일인지 확인함
        Call SetWallpaper(strFile)  'BMP라면 바로 바탕화면 변경함
    Else
        g_File = strFile
    '변경전 파일명을 기록해야지만 디스플레이 등록정보에서 리스트박스에 해당파일이 보여짐
        Call UpdateKey(HKEY_CURRENT_USER, RegPath, "ConvertedWallpaper", g_File)    'bmp로 변경전의 파일명을 기록
        Form2.Show vbModal  'bmp로 저장하기 위해 폼2를 로드
    End If
    g_File = ""
End Sub

Private Sub List2_Click()
    List1.ListIndex = List2.ListIndex
    
    strFile = strFolder & "\" & List1.List(List1.ListIndex) '선택된 파일명으로 로드할 그림파일의 전체경로를 설정
    strFile = List1.List(List1.ListIndex)  '선택된 파일명으로 로드할 그림파일의 전체경로를 설정
    
    Image1.Picture = LoadPicture(strFile)
End Sub
