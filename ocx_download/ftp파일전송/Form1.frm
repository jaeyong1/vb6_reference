VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '���� ����
   Caption         =   "Form1"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3315
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdDownload 
      Caption         =   "�ٿ�ε�"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1380
      Width           =   2355
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "���ε�"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   2355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "FTP ���� ���ε�/�ٿ�ε� ����"
      Height          =   180
      Left            =   420
      TabIndex        =   2
      Top             =   300
      Width           =   2595
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    ' FTP ���� ���ε�/�ٿ�ε� ����
    ' �и��ڱ�ȣ :  # (���������� ���ε�/�ٿ�ε�� ������ ����)
    ' ��������   : C:\Test.rar#C:\Test.r00#C:\Test1.r01
    ' ��������   : Test.rar#Test.r00#Test.r01
    ' �������   : /TEST1#/TEST2#/TEST3/TEST33
    ' ���۸��   : �ƽ�Ű or ���̳ʸ�
    
Option Explicit
    Private cFTPUp     As KSYSFileLib.clsUploadSys
    Private cFTPDown   As KSYSFileLib.clsDownloadSys

Private Sub cmdDownload_Click()
    Dim sDownloadFiles             As String           '������     ��������(�������)
    Dim sDownloadRemoteFiles       As String           '�ٿ�ε��� �������� �̸�
    Dim sDownloadRemotePath        As String           '�ٿ�ε��� �������� ���
        
        '�ٿ�ε��� ������ ���������̸�
        sDownloadFiles = App.Path & "\downtest.rar" & "#" & _
                         App.Path & "\downtest.r00" & "#" & _
                         App.Path & "\downtest.r01"
                         
        '�ٿ�ε��� ������ �����̸�
        sDownloadRemoteFiles = "test.rar" & "#" & _
                               "test.r00" & "#" & _
                               "test.r01"
       
        '�ٿ�ε��� ������ ���
        sDownloadRemotePath = "/" & "#" & _
                              "/" & "#" & _
                              "/"
        
        
    Set cFTPDown = New KSYSFileLib.clsDownloadSys    ' FTP���� �ٿ�ε� Ŭ��������
        cFTPDown.FtpServer = "211.218.221.11"        ' FTP���� �̸�����
        cFTPDown.RemotePort = 21                     ' FTP������ ��Ʈ
        cFTPDown.UserName = "anonymous"              ' FTP������ �����̸�
        cFTPDown.Password = "anonymous@anony.net"    ' FTP������ �н�����
        cFTPDown.PassiveMode = True                  ' Passive �������(�𸣽ø� �׳� ������.. ^_^;)
        cFTPDown.Timeout = 30                        ' FTP���� Ÿ�Ӿƿ�(��)
        cFTPDown.TransferMode = KSYS_FTP_BINARY_MODE
        cFTPDown.BufferSize = 4096
    
        ' �Լ�ȣ��
        ' cFTPDown.StartDownload(��������, �ٿ�ε� ��������, �ٿ�ε� �������, ���۸��)
        Call cFTPDown.StartDownload(sDownloadFiles, sDownloadRemoteFiles, sDownloadRemotePath, KSYS_FTP_BINARY_MODE)
    Set cFTPDown = Nothing

End Sub

Private Sub cmdUpload_Click()
    
    Dim sUploadFiles             As String           '���ε��� ��������(�������)
    Dim sUploadRemoteFiles       As String           '������ ��������   �̸�
    Dim sUploadRemotePath        As String           '������ ���������� ���

        '���ε��� ���������̸�
        sUploadFiles = App.Path & "\test.rar" & "#" & _
                       App.Path & "\test.r00" & "#" & _
                       App.Path & "\test.r01"

        '���ε��� ������ �����̸�
        sUploadRemoteFiles = "test.rar" & "#" & _
                             "test.r00" & "#" & _
                             "test.r01"
       '���ε��� ������ ���
        sUploadRemotePath = "/" & "#" & _
                            "/" & "#" & _
                            "/"
        
        
    Set cFTPUp = New KSYSFileLib.clsUploadSys      ' FTP���� ���ε� Ŭ��������
        cFTPUp.FtpServer = "211.218.221.11"        ' FTP���� �̸�����
        cFTPUp.RemotePort = 21                     ' FTP������ ��Ʈ
        cFTPUp.UserName = "anonymous"              ' FTP������ �����̸�
        cFTPUp.Password = "anonymous@anony.net"    ' FTP������ �н�����
        cFTPUp.PassiveMode = True                  ' Passive �������(�𸣽ø� �׳� ������.. ^_^;)
        cFTPUp.Timeout = 30                        ' FTP���� Ÿ�Ӿƿ�(��)
        cFTPUp.TransferMode = KSYS_FTP_BINARY_MODE
        cFTPUp.BufferSize = 4096

    
        ' �Լ�ȣ��
        ' cFTPUp.StartUpload(��������, ���ε� ��������, ���ε� �������, ���۸��)
        Call cFTPUp.StartUpload(sUploadFiles, sUploadRemoteFiles, sUploadRemotePath, KSYS_FTP_BINARY_MODE)
    Set cFTPUp = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
End Sub
