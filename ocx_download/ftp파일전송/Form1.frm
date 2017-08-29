VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
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
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdDownload 
      Caption         =   "다운로드"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1380
      Width           =   2355
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "업로드"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   2355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "FTP 파일 업로드/다운로드 예제"
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
    ' FTP 파일 업로드/다운로드 예제
    ' 분리자기호 :  # (다중파일을 업로드/다운로드시 파일을 구분)
    ' 로컬파일   : C:\Test.rar#C:\Test.r00#C:\Test1.r01
    ' 서버파일   : Test.rar#Test.r00#Test.r01
    ' 서버경로   : /TEST1#/TEST2#/TEST3/TEST33
    ' 전송모드   : 아스키 or 바이너리
    
Option Explicit
    Private cFTPUp     As KSYSFileLib.clsUploadSys
    Private cFTPDown   As KSYSFileLib.clsDownloadSys

Private Sub cmdDownload_Click()
    Dim sDownloadFiles             As String           '저장할     로컬파일(경로포함)
    Dim sDownloadRemoteFiles       As String           '다운로드할 서버파일 이름
    Dim sDownloadRemotePath        As String           '다운로드할 서버파일 경로
        
        '다운로드후 저장할 로컬파일이름
        sDownloadFiles = App.Path & "\downtest.rar" & "#" & _
                         App.Path & "\downtest.r00" & "#" & _
                         App.Path & "\downtest.r01"
                         
        '다운로드할 서버의 파일이름
        sDownloadRemoteFiles = "test.rar" & "#" & _
                               "test.r00" & "#" & _
                               "test.r01"
       
        '다운로드할 서버의 경로
        sDownloadRemotePath = "/" & "#" & _
                              "/" & "#" & _
                              "/"
        
        
    Set cFTPDown = New KSYSFileLib.clsDownloadSys    ' FTP파일 다운로드 클래스설정
        cFTPDown.FtpServer = "211.218.221.11"        ' FTP서버 이름지정
        cFTPDown.RemotePort = 21                     ' FTP서버의 포트
        cFTPDown.UserName = "anonymous"              ' FTP서버의 유저이름
        cFTPDown.Password = "anonymous@anony.net"    ' FTP서버의 패스워드
        cFTPDown.PassiveMode = True                  ' Passive 모드지정(모르시면 그냥 쓰세요.. ^_^;)
        cFTPDown.Timeout = 30                        ' FTP접속 타임아웃(초)
        cFTPDown.TransferMode = KSYS_FTP_BINARY_MODE
        cFTPDown.BufferSize = 4096
    
        ' 함수호출
        ' cFTPDown.StartDownload(로컬파일, 다운로드 서버파일, 다운로드 서버경로, 전송모드)
        Call cFTPDown.StartDownload(sDownloadFiles, sDownloadRemoteFiles, sDownloadRemotePath, KSYS_FTP_BINARY_MODE)
    Set cFTPDown = Nothing

End Sub

Private Sub cmdUpload_Click()
    
    Dim sUploadFiles             As String           '업로드할 로컬파일(경로포함)
    Dim sUploadRemoteFiles       As String           '저장할 서버파일   이름
    Dim sUploadRemotePath        As String           '저장할 서버파일의 경로

        '업로드할 로컬파일이름
        sUploadFiles = App.Path & "\test.rar" & "#" & _
                       App.Path & "\test.r00" & "#" & _
                       App.Path & "\test.r01"

        '업로드할 서버의 파일이름
        sUploadRemoteFiles = "test.rar" & "#" & _
                             "test.r00" & "#" & _
                             "test.r01"
       '업로드할 서버의 경로
        sUploadRemotePath = "/" & "#" & _
                            "/" & "#" & _
                            "/"
        
        
    Set cFTPUp = New KSYSFileLib.clsUploadSys      ' FTP파일 업로드 클래스설정
        cFTPUp.FtpServer = "211.218.221.11"        ' FTP서버 이름지정
        cFTPUp.RemotePort = 21                     ' FTP서버의 포트
        cFTPUp.UserName = "anonymous"              ' FTP서버의 유저이름
        cFTPUp.Password = "anonymous@anony.net"    ' FTP서버의 패스워드
        cFTPUp.PassiveMode = True                  ' Passive 모드지정(모르시면 그냥 쓰세요.. ^_^;)
        cFTPUp.Timeout = 30                        ' FTP접속 타임아웃(초)
        cFTPUp.TransferMode = KSYS_FTP_BINARY_MODE
        cFTPUp.BufferSize = 4096

    
        ' 함수호출
        ' cFTPUp.StartUpload(로컬파일, 업로드 서버파일, 업로드 서버경로, 전송모드)
        Call cFTPUp.StartUpload(sUploadFiles, sUploadRemoteFiles, sUploadRemotePath, KSYS_FTP_BINARY_MODE)
    Set cFTPUp = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
End Sub
