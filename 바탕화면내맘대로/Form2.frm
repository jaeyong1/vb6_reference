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
   StartUpPosition =   3  'Windows �⺻��
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
'    strName = Left(g_File, i) & "bmp"   '������ ���ϸ� ����
    
    Image1.Stretch = False  '�̹��� ũ�⿡ �°� Image1�� ũ�Ⱑ ���ϵ��� ����
    Image1.Picture = LoadPicture(g_File)
    
    Me.Height = Image1.Height + 550     'Image1�� ũ�⺸�� ���� ũ�Ⱑ �׻� Ŀ�߸� ����� �����
    Me.Width = Image1.Width + 180
    
    strName = "C:\WINDOWS\Wallpaper1.bmp"   'WINDOWS������ �ִٰ� �����Ͽ�(������ ������)
    
    '������ ���ϸ����� ������ ������ ������
    If Dir(strName) <> vbNullString Then Kill strName
    
    SavePicture Image1, strName     'savepicture�� �����ϸ� jpg�� gif�� ��� bmp�� �����
    
    i = 0   '���� �ʱ�ȭ
    
ReChk:
'############ ������ �����ϸ� �ϵ��ũ�� ������ �����Ǳ� ������ �Ʒ���ƾ�� ���� ����Ǹ� �����߻��� ###########
'           ������ �����ϱ� ���ؼ� ���μ����� ���� ���߰� ������ �����ɶ����� ���ѷ���...
    Sleep 100   '0.1�ʰ� ���μ��� ����
    
    If Dir(strName) = vbNullString Then     '���� ������ bmp������ �����Ǿ������� �˻���
        If i >= 300 Then    '�����ð����� ���ϻ����� �ʵɰ�� ���ѷ������� �ʵ��� ��������
            MsgBox "�������� ����"
        Else
            i = i + 1
            GoTo ReChk
        End If
    Else
    '### �Ʒ� ������Ʈ����("OriginalWallpaper")�� "Wallpaper"������Ʈ������ ��ΰ� �����ϸ�
    '��1���� ������ "ConvertedWallpaper"���� ��ΰ� ���÷��� ����������� ǥ�õ�
    '�ٸ����� "Wallpaper"���� ǥ��
        Call UpdateKey(HKEY_CURRENT_USER, RegPath, "OriginalWallpaper", strName)
        Call SetWallpaper(strName)    '������ ���������� �ٷ� ����ȭ�� ����
    End If
    
    Unload Me   '������ ���� ����
End Sub

Private Sub Form_Load()
'###### ��2�� ������ �ʵ��� �ָ�~ ���ָ�~��...��������...��########
    Me.Left = -99999
    Me.Top = -99999
End Sub
