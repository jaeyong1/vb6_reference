VERSION 5.00
Object = "{6580F760-7819-11CF-B86C-444553540000}#1.0#0"; "EZFTP.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin EZFTPLib.EZFTP EZFTP1 
      Left            =   840
      Top             =   1800
      _Version        =   65536
      _ExtentX        =   800
      _ExtentY        =   800
      _StockProps     =   0
      LocalFile       =   ""
      RemoteFile      =   ""
      RemoteAddres    =   ""
      UserName        =   ""
      Password        =   ""
      Binary          =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    EZFTP1.RemoteAddress = "ftp.hanmir.com" '��)ftp.hanmir.com
    EZFTP1.UserName = "pjaeyong1"           '��)merong
    EZFTP1.Password = "6090"             '��)babo
    
    EZFTP1.Connect  '����
    EZFTP1.RemoteFile = "version.txt"  '������ �ִ� �����̸�
    EZFTP1.LocalFile = "c:\1.txt"      '�Ŀ� ������ �����̸�
    EZFTP1.GetFile   '�ޱ�..
    
    
   'EZFTP1.DeleteFile ("2.txt")  '������ �ִ� ���� �����..
      'file ������ �����߻�...
      
      
    EZFTP1.LocalFile = "c:\2.txt"   '�ø����� ����
    EZFTP1.RemoteFile = "version.txt"     '������ ����� �̸�
    EZFTP1.PutFile                  '�ö󰡶�!!
    
    
    EZFTP1.Disconnect  '����� ������..
    
    
End Sub
