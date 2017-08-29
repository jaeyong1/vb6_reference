VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  '���
   BackColor       =   &H80000005&
   BorderStyle     =   1  '���� ����
   Caption         =   "����ȭ�� �ٲٱ�"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   6210
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   4530
      TabIndex        =   6
      Top             =   4050
      Width           =   1545
   End
   Begin VB.ListBox List2 
      Appearance      =   0  '���
      Height          =   2190
      Left            =   90
      TabIndex        =   5
      Top             =   3630
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
      Caption         =   "����ȭ�� �ٲٱ�"
      Height          =   495
      Left            =   4500
      Style           =   1  '�׷���
      TabIndex        =   1
      Top             =   5340
      Width           =   1605
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
      Caption         =   "�׸����� �˻�"
      Height          =   495
      Left            =   4500
      Style           =   1  '�׷���
      TabIndex        =   3
      Top             =   4650
      Width           =   1605
   End
   Begin VB.ListBox List1 
      Appearance      =   0  '���
      Height          =   2190
      Left            =   150
      TabIndex        =   0
      Top             =   6420
      Width           =   5925
   End
   Begin VB.Label Label3 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "��ġ :"
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
      BackStyle       =   0  '����
      Caption         =   "2222222222222222222222"
      Height          =   180
      Left            =   1470
      TabIndex        =   4
      Top             =   6240
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
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

Const strFolder = "C:\WINDOWS"     '�о���� �������
Const Check_Name = "JPG,JPGE,GIF,BMP"           '�о���� ���� Ȯ���� ��� ","�� ����
Dim strFile As String
Dim Cnt As Long

Private Sub Command2_Click()
    Call Set_List
End Sub

'List Box�� ���ϸ��� �о����
Private Sub Set_List()

    Cnt = 0
    
    Command1.Enabled = False
    Command2.Enabled = False
    
    Label2.Caption = "�׸� ������ �˻��� �Դϴ�..."
    List1.Clear
    
    Call GetFolderList(strFolder)   '������������ �˻��ϱ� ���� ����Լ� ȣ��
    
    Label1.Caption = "�� " & Cnt & "���� �׸������� ã�ҽ��ϴ�."
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
    
    '������ ��ΰ� ���� ª�� ��θ����� ��������
    strName = GetShortPathName(fPath, temp, 254)
    strName = Left$(temp, strName)
    Label1.Caption = "��������: " & strName
    
    For Each f_File In fCnt
        DoEvents
        If Check_FileName(f_File.Name) = False Then GoTo Skip_Next  '�о���ϼ� �ִ� �׸��������� Ȯ���ϱ� ���� �Լ�ȣ��
        
        List1.AddItem fPath & "\" & f_File.Name     'list1���� ������ ��ü��θ� �����ϰ�
        List2.AddItem f_File.Name       'list2���� ���ϸ� �����ϵ���
        Cnt = Cnt + 1
Skip_Next:
    Next
    
    For Each f1 In sf       '���������� �ִٸ� ���ȣ��...
        Call GetFolderList(fPath & "\" & f1.Name)
    Next
    
End Sub


Private Sub Form_Load()
    Combo1.AddItem "���"
    Combo1.AddItem "�ٵ��ǽ�"
    Combo1.AddItem "���̱�"
    
    Label1.Caption = ""
    Label2.Caption = ""
    Command1.Enabled = False
    
    g_File = GetKeyValue(HKEY_CURRENT_USER, RegPath, "Wallpaper")   '���� ����� ����ȭ�� �׸��� �о����
    
    Image1.Stretch = True   'Image1�� �׸�ũ�⸦ ����
    Image1.Picture = LoadPicture(g_File)
    
    '### ��ġ(���,�ٵ��ǽ�,���̱�) ������ ������ �ִ� ������Ʈ�� ���� �о ǥ���� ###
    TitleStyle = GetKeyValue(HKEY_CURRENT_USER, RegPath, "TileWallpaper")
    PaperStyle = GetKeyValue(HKEY_CURRENT_USER, RegPath, "WallpaperStyle")
    
'    MsgBox GetKeyValue(HKEY_CURRENT_USER, RegPath, "ConvertedWallpaper Last WriteTime")
    
    If TitleStyle = "0" And PaperStyle = "0" Then   '���
        Combo1.ListIndex = 0
    ElseIf TitleStyle = "1" And PaperStyle = "0" Then   '�ٵ��ǽ�
        Combo1.ListIndex = 1
    ElseIf TitleStyle = "0" And PaperStyle = "2" Then   '���̱�
        Combo1.ListIndex = 2
    Else    '����??
        
    End If
End Sub


'�ش� ������ �׸��������� Ȯ���ϴ� �Լ�
Private Function Check_FileName(ByVal strName As String) As Boolean
    Dim i As Integer        '������ ���� ��������
    Dim strTemp As String   '�ӽ÷� ����� ����
    Dim ArrName() As String '���ǵ� Ȯ���ڵ��� ���� �迭����
    
    ArrName() = Split(Check_Name, ",")  '���ǵ� Ȯ���ڵ��� ","�� �������� �迭�� ����
    
    i = InStrRev(strName, ".")      '���ϸ��� �ڿ������� "."�� ã����
    strTemp = Mid(strName, i + 1)   'Ȯ���ڸ��� �ӽú����� ����
    
    For i = 0 To UBound(ArrName())  '���ǵ� Ȯ������ ������ŭ ����
        If UCase(strTemp) = UCase(ArrName(i)) Then  '�빮�ڷ� �����Ͽ� ��
            Check_FileName = True   '��ġ�ϴ� Ȯ���ڰ� ������ ��������
            Exit Function
        End If
    Next
    
End Function

Private Sub Command1_Click()
    Dim i As Integer

    If strFile = "" Then
        MsgBox "������ �����ϼ�         ", vbInformation, "����"
        Exit Sub
    End If
    
    i = InStrRev(strFile, ".")
    
    '################ ��ġ(���,�ٵ��ǽ�,���̱�) Ȯ���� ��� ############
    If Combo1.ListIndex = 0 Then    '���
        TitleStyle = "0"
        PaperStyle = "0"
    ElseIf Combo1.ListIndex = 1 Then    '�ٵ��ǽ�
        TitleStyle = "1"
        PaperStyle = "0"
    ElseIf Combo1.ListIndex = 2 Then    '���̱�
        TitleStyle = "0"
        PaperStyle = "2"
    End If
    '������Ʈ������ �̸� �����ؾ߸� �����
    Call UpdateKey(HKEY_CURRENT_USER, RegPath, "TileWallpaper", TitleStyle)
    Call UpdateKey(HKEY_CURRENT_USER, RegPath, "WallpaperStyle", PaperStyle)
    
    If UCase(Mid(strFile, i + 1)) = "BMP" Then  'bmp�������� Ȯ����
        Call SetWallpaper(strFile)  'BMP��� �ٷ� ����ȭ�� ������
    Else
        g_File = strFile
    '������ ���ϸ��� ����ؾ����� ���÷��� ����������� ����Ʈ�ڽ��� �ش������� ������
        Call UpdateKey(HKEY_CURRENT_USER, RegPath, "ConvertedWallpaper", g_File)    'bmp�� �������� ���ϸ��� ���
        Form2.Show vbModal  'bmp�� �����ϱ� ���� ��2�� �ε�
    End If
    g_File = ""
End Sub

Private Sub List2_Click()
    List1.ListIndex = List2.ListIndex
    
    strFile = strFolder & "\" & List1.List(List1.ListIndex) '���õ� ���ϸ����� �ε��� �׸������� ��ü��θ� ����
    strFile = List1.List(List1.ListIndex)  '���õ� ���ϸ����� �ε��� �׸������� ��ü��θ� ����
    
    Image1.Picture = LoadPicture(strFile)
End Sub
