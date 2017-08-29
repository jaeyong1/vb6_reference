VERSION 5.00
Begin VB.Form RTGraph 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�ǽð� �׷���"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  '�ȼ�
   ScaleWidth      =   259
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.PictureBox Picture1 
      Height          =   2325
      Left            =   45
      ScaleHeight     =   150
      ScaleMode       =   0  '�����
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
      BorderStyle     =   1  '���� ����
      Caption         =   "�ǽð� �׷���"
      BeginProperty Font 
         Name            =   "����"
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

'//�����

Private Const SRCCOPY = &HCC0020        '(DWORD) dest = source
Private Const PS_SOLID = 0

Private Declare Function CreateCompatibleDC Lib "gdi32" ( _
                ByVal hdc As Long) _
                As Long
'���� DC�� ȣȯ�� �ִ� DC�� �����Ѵ�
'(Screen ��ü�� ������� Device Context Handle�� �ʿ��� ��� ���)
'hdc : �޸� DC�� ����

Private Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
                ByVal hdc As Long, _
                ByVal nWidth As Long, _
                ByVal nHeight As Long) _
                As Long
'�ۼ��� ����̽��� ȣȯ���� �ִ� ��Ʈ���� �����Ѵ�.
'hdc     : ��Ʈ���� �����ϱ� ���ϴ� DC�� �ڵ�
'nWidth  : ������ ��Ʈ���� ����ũ��
'nHeight : ������ ��Ʈ���� ���� ũ��

Private Declare Function SelectObject Lib "gdi32" ( _
                ByVal hdc As Long, _
                ByVal hObject As Long) _
                As Long
'��ü�� Device Context�� �����Ѵ�.
'hdc     : DC(Device Context)�� ����
'hObject : ������Ʈ�� �ڵ� ��

Private Declare Function CreatePen Lib "gdi32" ( _
                ByVal nPenStyle As Long, _
                ByVal nWidth As Long, _
                ByVal crColor As Long) _
                As Long
'������ ���� �����Ѵ�.
'nPenStyle : ���� ���� (PS_SOLID : �Ǽ� ��)
'nWidth    : ���� ������ ��Ÿ�� ���� ����
'crColor    : ���� ����(RGB ��ũ�� ���)

Private Declare Function LineTo Lib "gdi32" ( _
                ByVal hdc As Long, _
                ByVal x As Long, _
                ByVal y As Long) _
                As Long
'CP(Current Point : ������ ��ġ)�κ��� ������ ��ǥ���� ���� �ߴ� �Լ�
'hdc : DC(Device Context) �ڵ�
'x    : ���� ������ ���� X��ǥ
'y    : ���� ������ ���� Y��ǥ

Private Declare Function MoveToEx Lib "gdi32" ( _
                ByVal hdc As Long, _
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal lpPoint As Long) _
                As Long
'CP�� ������ ��ǥ�� �̵���Ű�� �Լ�
'hdc     : ���ϴ� DC(Device Context)�� �����Ѵ�.
'x       : �������� ���� X��ǥ
'y       : �������� ���� Y��ǥ
'lpPoint  : POINTAPI ����ü

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
'A�� ȭ�󿡼� B�� ȭ������ ������ �簢�� ������ �̹����� ��Ʈ������ �����Ѵ�.
'hDestDC : ��� ȭ���� DC(Device Context)�� �ڵ�
'x        : ��� �簢���� ���� ��� X��ǥ
'y        : ��� �簢���� ���� ��� Y��ǥ
'nWidth   : ������ ��� �簢���� ��
'nHeight  : ������ ��� �簢���� ����
'hSrcDC  : ���� ȭ���� DC(Device Context)�� �ڵ�
'xSrc     : ���� �簢���� ������� X��ǥ
'ySrc     : ���� �簢���� ������� Y��ǥ
'dwRop   : �������� �����ڵ�

'��� ����
Private Const pWidth = 250      'picture box�� �� ��ġ�� ����� ����
Private Const pHeight = 150     'picture box�� ���� ��ġ�� ����� ����
Private Const pGrid = 25        'grid �� ������ ���� ��ġ�� ����� ����
Private Const tInterval = 100   'Timer�� Interval�� 100���� ����(1/10��)
Private Const pHeightHalf = pHeight \ 2
Dim counter As Long             'Number of data points logged so far. Used to sync grid.
Dim oldY As Long                'Contains the previous y coordinate.
Dim hDCh As Long, hPenB As Long, hPenC As Long, hPenW As Long


'Code��

'1) Visual Basic�� �����Ͽ� ǥ�� EXE�� ����. Form1�� �Ӽ��� �⺻����...
'2) Timer�� PictureBox�� Form1�� �����.

Private Sub Form_Load()
    Dim hBmp As Long
    Dim i As Integer
    Me.Show
    Picture1.ScaleMode = 3   'picture box�� ScaleMode�� 3-�ȼ��� ����
    Picture1.Left = 3
    Picture1.Top = 54
    RTGraph.ScaleMode = 3      'Form1�� ScaleMode�� 3-�ȼ��� ����
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
    
    

'grid ������ 25���� ���� ���� �׸���.
'    For i = pGrid To pHeight - 1 Step pGrid
'        Picture1.Line (0, i)-(pWidth, i)
'    Next

'grid ������ 25���� ���� ���� �׸���.
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
 



