VERSION 5.00
Begin VB.Form RTGraph 
   BackColor       =   &H00FFFFFF&
   Caption         =   "실시간 그래프"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   259
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox Picture1 
      Height          =   2325
      Left            =   45
      ScaleHeight     =   150
      ScaleMode       =   0  '사용자
      ScaleWidth      =   250
      TabIndex        =   0
      Top             =   810
      Width           =   3825
   End
   Begin VB.Timer Timer1 
      Left            =   45
      Top             =   2790
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '단일 고정
      Caption         =   "실시간 그래프"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   540
      Left            =   420
      TabIndex        =   1
      Top             =   135
      Width           =   3105
   End
End
Attribute VB_Name = "RTGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Task: Create a dynamically scrolling graph in Visual Basic. Creates a graph as data is generated in
'real-time, such as in a monitoring program.

'//선언부

Private Const SRCCOPY = &HCC0020        '(DWORD) dest = source
Private Const PS_SOLID = 0

Private Declare Function CreateCompatibleDC Lib "gdi32" ( _
                ByVal hdc As Long) _
                As Long
'기존 DC와 호환성 있는 DC를 생성한다
'(Screen 전체를 대상으로 Device Context Handle이 필요한 경우 사용)
'hdc : 메모리 DC를 생성

Private Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
                ByVal hdc As Long, _
                ByVal nWidth As Long, _
                ByVal nHeight As Long) _
                As Long
'작성된 디바이스와 호환성이 있는 비트맵을 생성한다.
'hdc     : 비트맵을 적용하길 원하는 DC의 핸들
'nWidth  : 생성될 비트맵의 수평크기
'nHeight : 생성될 비트맵의 수직 크기

Private Declare Function SelectObject Lib "gdi32" ( _
                ByVal hdc As Long, _
                ByVal hObject As Long) _
                As Long
'객체를 Device Context로 선택한다.
'hdc     : DC(Device Context)를 생성
'hObject : 오브젝트의 핸들 값

Private Declare Function CreatePen Lib "gdi32" ( _
                ByVal nPenStyle As Long, _
                ByVal nWidth As Long, _
                ByVal crColor As Long) _
                As Long
'논리적인 펜을 생성한다.
'nPenStyle : 펜의 형태 (PS_SOLID : 실선 펜)
'nWidth    : 논리적 단위로 나타낸 펜의 굵기
'crColor    : 펜의 색깔(RGB 매크로 사용)

Private Declare Function LineTo Lib "gdi32" ( _
                ByVal hdc As Long, _
                ByVal x As Long, _
                ByVal y As Long) _
                As Long
'CP(Current Point : 현재의 위치)로부터 지정된 좌표까지 선을 긋는 함수
'hdc : DC(Device Context) 핸들
'x    : 선분 끝점의 논리적 X좌표
'y    : 선분 끝점의 논리적 Y좌표

Private Declare Function MoveToEx Lib "gdi32" ( _
                ByVal hdc As Long, _
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal lpPoint As Long) _
                As Long
'CP를 지정된 좌표로 이동시키는 함수
'hdc     : 원하는 DC(Device Context)를 지시한다.
'x       : 목적점의 논리적 X좌표
'y       : 목적점의 논리적 Y좌표
'lpPoint  : POINTAPI 구조체

Private Declare Function BitBlt Lib "gdi32" ( _
                ByVal hDestDC As Long, _
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal nWidth As Long, _
                ByVal nHeight As Long, _
                ByVal hSrcDC As Long, _
                ByVal xSrc As Long, _
                ByVal ySrc As Long, _
                ByVal dwRop As Long) _
                As Long
'A란 화상에서 B란 화상으로 지정한 사각형 범위의 이미지를 비트맵으로 복사한다.
'hDestDC : 대상 화상의 DC(Device Context)의 핸들
'x        : 대상 사각형의 좌측 상단 X좌표
'y        : 대상 사각형의 좌측 상단 Y좌표
'nWidth   : 원본과 대상 사각형의 폭
'nHeight  : 원본과 대상 사각형의 높이
'hSrcDC  : 원본 화상의 DC(Device Context)의 핸들
'xSrc     : 원본 사각형의 좌측상단 X좌표
'ySrc     : 원본 사각형의 좌측상단 Y좌표
'dwRop   : 래스터의 연산코드

'상수 정의
Private Const pWidth = 250      'picture box의 폭 수치를 상수로 정의
Private Const pHeight = 150     'picture box의 높이 수치를 상수로 정의
Private Const pGrid = 25        'grid 선 사이의 간격 수치를 상수로 정의
Private Const tInterval = 100   'Timer의 Interval을 100으로 정의(1/10초)
Private Const pHeightHalf = pHeight \ 2
Dim counter As Long             'Number of data points logged so far. Used to sync grid.
Dim oldY As Long                'Contains the previous y coordinate.
Dim hDCh As Long, hPenB As Long, hPenC As Long, hPenW As Long


'Code부

'1) Visual Basic을 시작하여 표준 EXE를 선택. Form1의 속성을 기본으로...
'2) Timer와 PictureBox를 Form1에 만든다.

Private Sub Form_Load()
    Dim hBmp As Long
    Dim i As Integer
    Me.Show
    Picture1.ScaleMode = 3   'picture box의 ScaleMode를 3-픽셀로 설정
    Picture1.Left = 3
    Picture1.Top = 54
    RTGraph.ScaleMode = 3      'Form1의 ScaleMode를 3-픽셀로 설정
    Picture1.Width = 255
    Picture1.Height = 155
    Picture1.BackColor = vbBlack
    counter = 0

    hDCh = CreateCompatibleDC(Picture1.hdc)
    hBmp = CreateCompatibleBitmap(Picture1.hdc, _
                                  pWidth, _
                                  pHeight)

    Call SelectObject(hDCh, hBmp)
    hPenB = CreatePen(PS_SOLID, 0, vbBlack)
    hPenC = CreatePen(PS_SOLID, 0, vbRed)
    hPenW = CreatePen(PS_SOLID, 0, vbWhite)
    Call SelectObject(hDCh, hPenB)
    
    

'grid 간격인 25마다 가로 선을 그린다.
'    For i = pGrid To pHeight - 1 Step pGrid
'        Picture1.Line (0, i)-(pWidth, i)
'    Next

'grid 간격인 25마다 세로 선을 그린다.
'    For i = pGrid - (counter Mod pGrid) To _
'                     pWidth - 1 Step pGrid
'        Picture1.Line (i, 0)-(i, pHeight)
'    Next

    Call BitBlt(hDCh, _
                0, _
                0, _
                pWidth, _
                pHeight, _
                Picture1.hdc, _
                0, _
                0, _
                SRCCOPY)
                
    Timer1.Interval = 100
    Timer1.Enabled = True
'    oldY = pHeightHalf
    oldY = 0
    
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    Call BitBlt(hDCh, _
                  0, _
                  0, _
                  pWidth - 1, _
                  pHeight, _
                  hDCh, _
                  1, _
                  0, _
                  SRCCOPY)
    
'    Call SelectObject(hDCh, hPenC)

'    If counter Mod pGrid = 0 Then
'        Call MoveToEx(hDCh, pWidth - 2, 0, 0)
'        Call LineTo(hDCh, pWidth - 2, pHeight)
'    End If

    i = Sin(0.1 * counter) * _
         (pHeightHalf - 1) + _
         pHeightHalf

    Call SelectObject(hDCh, hPenW)
    Call MoveToEx(hDCh, pWidth - 3, oldY, 0)
    Call LineTo(hDCh, pWidth - 2, i)
    Call SelectObject(hDCh, hPenB)
    Call BitBlt(Picture1.hdc, _
                  0, _
                  0, _
                  pWidth, _
                  pHeight, _
                  hDCh, _
                  0, _
                  0, _
                  SRCCOPY)
    counter = counter + 1
    oldY = i
End Sub
 



