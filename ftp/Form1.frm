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
   StartUpPosition =   3  'Windows 기본값
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
    EZFTP1.RemoteAddress = "ftp.hanmir.com" '예)ftp.hanmir.com
    EZFTP1.UserName = "pjaeyong1"           '예)merong
    EZFTP1.Password = "6090"             '예)babo
    
    EZFTP1.Connect  '접속
    EZFTP1.RemoteFile = "version.txt"  '서버에 있는 파일이름
    EZFTP1.LocalFile = "c:\1.txt"      '컴에 저장할 파일이름
    EZFTP1.GetFile   '받기..
    
    
   'EZFTP1.DeleteFile ("2.txt")  '서버에 있는 파일 지우기..
      'file 없으면 에러발생...
      
      
    EZFTP1.LocalFile = "c:\2.txt"   '올릴파일 선택
    EZFTP1.RemoteFile = "version.txt"     '서버에 저장될 이름
    EZFTP1.PutFile                  '올라가라!!
    
    
    EZFTP1.Disconnect  '끊기로 추정됨..
    
    
End Sub
