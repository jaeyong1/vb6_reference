VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6580F760-7819-11CF-B86C-444553540000}#1.0#0"; "EZFTP.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form main 
   Caption         =   "파일나라 1.05 Beta #1"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10590
   Icon            =   "filenara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   10590
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "종료"
      Height          =   375
      Left            =   9840
      TabIndex        =   31
      Top             =   6960
      Width           =   735
   End
   Begin InetCtlsObjects.Inet inet 
      Left            =   7320
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer nodisconnect 
      Interval        =   30000
      Left            =   2520
      Top             =   7080
   End
   Begin VB.ListBox iplist 
      Height          =   240
      Left            =   480
      TabIndex        =   2
      Top             =   7080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.FileListBox filelist 
      Height          =   270
      Left            =   1560
      TabIndex        =   1
      Top             =   7080
      Visible         =   0   'False
      Width           =   120
   End
   Begin EZFTPLib.EZFTP ftp 
      Left            =   960
      Top             =   6960
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
   Begin TabDlg.SSTab tabz2 
      Height          =   6975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   12303
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "P2P 파일검색"
      TabPicture(0)   =   "filenara.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "work"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "per"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "모니터링&&설정"
      TabPicture(1)   =   "filenara.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "upload(0)"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "Frame6"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "파일나라에 관하여"
      TabPicture(2)   =   "filenara.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(2)=   "Label4"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame6 
         Caption         =   "설정"
         Height          =   3495
         Left            =   -74880
         TabIndex        =   23
         Top             =   3240
         Width           =   10335
         Begin VB.CommandButton Command2 
            Caption         =   "설정 변경 완료[적용]"
            Height          =   255
            Left            =   7200
            TabIndex        =   34
            Top             =   3120
            Width           =   3015
         End
         Begin VB.Frame Frame10 
            Caption         =   "다운로드 파일 저장 폴더"
            Height          =   3135
            Left            =   3480
            TabIndex        =   33
            Top             =   240
            Width           =   3135
            Begin VB.DirListBox Dir2 
               Height          =   2400
               Left            =   120
               TabIndex        =   38
               ToolTipText     =   "반드시 더블클릭하세요."
               Top             =   600
               Width           =   2895
            End
            Begin VB.DriveListBox Drive2 
               Height          =   300
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   2895
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "공유 폴더"
            Height          =   3135
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   3015
            Begin VB.DirListBox Dir1 
               Height          =   2400
               Left            =   120
               TabIndex        =   36
               ToolTipText     =   "반드시 더블클릭하세요."
               Top             =   600
               Width           =   2775
            End
            Begin VB.DriveListBox Drive1 
               Height          =   300
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   2775
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "접속제한"
            Height          =   1455
            Left            =   7200
            TabIndex        =   27
            Top             =   1560
            Width           =   3015
            Begin VB.TextBox onlyyou 
               Height          =   270
               Left            =   360
               TabIndex        =   30
               Top             =   960
               Width           =   2415
            End
            Begin VB.CheckBox onlyuser 
               Caption         =   "특정 아이피만 접속 가능"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   240
               Width           =   2535
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "접속 가능한 사용자 IP:"
               Height          =   180
               Left            =   240
               TabIndex        =   29
               Top             =   600
               Width           =   1845
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "접속 설정"
            Height          =   1215
            Left            =   7200
            TabIndex        =   24
            Top             =   240
            Width           =   3015
            Begin VB.TextBox maxcnt 
               Height          =   270
               Left            =   720
               MaxLength       =   3
               TabIndex        =   26
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "최대 접속자 수:"
               Height          =   180
               Left            =   240
               TabIndex        =   25
               Top             =   360
               Width           =   1260
            End
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "파일나라는...."
         Height          =   4815
         Left            =   -74760
         TabIndex        =   19
         Top             =   600
         Width           =   10095
         Begin VB.TextBox Text1 
            Appearance      =   0  '평면
            Enabled         =   0   'False
            Height          =   4215
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   20
            Text            =   "filenara.frx":0496
            Top             =   360
            Width           =   9615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "IP로 검색"
         Height          =   975
         Left            =   5400
         TabIndex        =   15
         Top             =   480
         Width           =   4935
         Begin VB.CommandButton findnow2 
            Caption         =   "검색"
            Height          =   375
            Left            =   3720
            TabIndex        =   18
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox rip 
            Height          =   270
            Left            =   480
            TabIndex        =   17
            Top             =   600
            Width           =   3135
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "상대측 IP:"
            Height          =   180
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.Frame frame1 
         Caption         =   "검색어로 검색"
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   5055
         Begin VB.TextBox findword 
            Height          =   270
            Left            =   360
            TabIndex        =   11
            Top             =   600
            Width           =   3255
         End
         Begin VB.CommandButton findnow 
            Caption         =   "검색"
            Height          =   375
            Left            =   3840
            TabIndex        =   10
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "검색어:"
            Height          =   180
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   600
         End
      End
      Begin VB.Frame frame2 
         Caption         =   "검색 결과"
         Height          =   5055
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   10095
         Begin MSComctlLib.ListView listz 
            Height          =   4695
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   8281
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "파일명"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "아이피"
               Object.Width           =   4057
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "파일 크기"
               Object.Width           =   4057
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "주고 있는 상황"
         Height          =   2535
         Left            =   -74880
         TabIndex        =   5
         Top             =   600
         Width           =   10335
         Begin MSComctlLib.ListView sendlist 
            Height          =   2175
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   3836
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "아이피"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "파일명"
               Object.Width           =   5644
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "크기"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "전송량"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "전송률"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "상태"
               Object.Width           =   3175
            EndProperty
         End
      End
      Begin MSComctlLib.ProgressBar per 
         Height          =   135
         Left            =   840
         TabIndex        =   4
         Top             =   6720
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSWinsockLib.Winsock upload 
         Index           =   0
         Left            =   -65160
         Top             =   3000
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "=4uzr="
         Height          =   180
         Left            =   -69840
         TabIndex        =   22
         Top             =   6480
         Width           =   540
      End
      Begin VB.Label Label4 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "FILENARA.COM.NE.KR"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   27.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   -72480
         MouseIcon       =   "filenara.frx":088E
         MousePointer    =   99  '사용자 정의
         TabIndex        =   21
         Top             =   5760
         Width           =   5880
      End
      Begin VB.Label work 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   6720
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "작업:"
         Height          =   180
         Left            =   360
         TabIndex        =   13
         Top             =   6720
         Width           =   420
      End
   End
   Begin VB.Label logon 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "=== Login to Server... Please wait... ==="
      Height          =   225
      Left            =   3840
      TabIndex        =   0
      Top             =   7080
      Width           =   3315
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'이 소스에 대한 저작권(?)은 김현철(earus@hanmail.net)에게 있으며
'이미 인터넷상에 실행파일로 공개된 프로그램이므로 이 프로그램을 수정하여 배포하는
'일은 허락하지 않음을 밝힙니다.
'오로지 공부를 목표로 보세요.
'(이 소스를 볼 모든 분들께.. 이게 소스인지 쓰레기인지 구분이 않가더라도 이해해주시길
' 부탁드립니다)
'Site ::: HTTP://FILENARA.COM.NE.KR/ or HTTP://any.to/4user
'E-Mail ::: earus@hanmail.net
이 부분을 꼭 읽으시라고 일부러 오류나도록 해놨습니다.^^

Dim stopnow As Single, i As Single, j As Single, k As Single, x As Single
Dim cntnum As Single
Dim myip As String
Dim upn As Single
Dim dirpath As String, downpath As String
Dim getbodycount As Single
Function nospace(getspace As String)
    Dim i As Single
    Dim tot As String
    
    tot = ""
    
    For i = 1 To Len(getspace)
        If Mid$(getspace, i, 1) <> " " Then tot = tot + Mid$(getspace, i, 1)
    Next
    
    nospace = tot
End Function
Function midz(midbody, startnum As Single, endnum As Single)
    midz = Mid$(midbody, startnum, endnum - startnum + 1)
End Function
Private Sub Check1_Click()
    If Check1.Value = 1 Then onlyyou.Enabled = True Else onlyyou.Enabled = False
End Sub
Private Sub Command1_Click()
    On Error Resume Next
    main.Hide
    If maxcnt = "" Then maxcnt = "10"
    Open "filenara.cfg" For Output As #1
        Print #1, Dir1.Path
        Print #1, Dir2.Path
        Print #1, maxcnt
    Close #1

    Dim filez
    Cancel = 1
    filez = "?????"
    Open "c:\filelist.tmp" For Output As #1
        Print #1, filez
    Close #1
    ftp.LocalFile = "C:\filelist.tmp"
    myip = inet.OpenURL("http://202.31.225.227:8080/webchat5/yourip.shtml")
    myip = nospace(midz(myip, InStr(1, myip, "<font color=blue size=4><b>") + 27, InStr(1, myip, "</b></font><br>") - 1))
    ftp.DeleteFile (myip)
    ftp.RemoteFile = myip
    'ftp.PutFile
    DoEvents
    End
End Sub
Private Sub Command2_Click()
    If maxcnt = "" Then maxcnt = "10"
    
    Open "filenara.cfg" For Output As #1
        Print #1, Dir1.Path
        Print #1, Dir2.Path
        Print #1, maxcnt
    Close #1
    
    dirpath = Dir1.Path
    filelist.Path = dirpath

    If filelist.ListCount <> 0 Then
        For i = 0 To filelist.ListCount - 1
            filez = filez + filelist.List(i) + "/" + Str(FileLen(filelist.Path + "\" + filelist.List(i))) + "?"
        Next
    Else
        filez = "?????"
    End If

    Open "c:\filelist.tmp" For Output As #1
        Print #1, filez
    Close #1
        
    DoEvents
   ftp.LocalFile = "C:\filelist.tmp"
    ftp.DeleteFile (myip)
   DoEvents
    ftp.RemoteFile = myip
   DoEvents
    ftp.PutFile
   DoEvents
    
    'If dirpath <> Dir1.Path Or downpath <> Dir2.Path Then
        MsgBox "설정내용을 적용하였습니다. 프로그램을 재시작합니다.", , "적용완료"
       'Shell App.Path + "\" + App.EXEName, vbNormal
'        End
    'Else
        MsgBox "설정내용 적용 완료.", , "적용완료"
    'End If
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub
Private Sub Drive2_Change()
    On Error Resume Next
    Dir2.Path = Drive2.Drive
End Sub
Private Sub findnow_Click()
    On Error GoTo err
    If findnow.Caption = "검색" Then
            
    If findword = "" Then MsgBox "검색어를 입력해주세요.", , "오류": Exit Sub
    
    Frame4.Enabled = False
    
    listz.ListItems.Clear
    iplist.Clear
    
    Dim filelistz As String, temp As String, count As Single
    
    findnow.Caption = "중지"
    
    ftp.GetDirectory ("*.*")
    DoEvents
    stopnow = 0
    
    If iplist.ListCount <> 0 Then
        For i = 0 To iplist.ListCount - 1
            On Error Resume Next
            per.Value = (i + 1) / iplist.ListCount * 100
            If stopnow = 1 Then Exit Sub
            filelistz = ""
            getbodycount = getbodycount + 1
            'Load getbody(getbodycount)
            'filelistz = getbody(getbodycount).OpenURL("http://home.hanmir.com/~earus/" + iplist.List(i))
            If Dir("c:\filenara.tmp", vbNormal) <> "" Then Kill "c:\filenara.tmp"
            ftp.RemoteFile = iplist.List(i)
            ftp.LocalFile = "c:\filenara.tmp"
            ftp.GetFile
            DoEvents
            
            Do Until Dir("c:\filenara.tmp", vbNormal) <> ""
                DoEvents
            Loop
            
            Open "c:\filenara.tmp" For Input As #2
                Line Input #2, filelistz
            Close #2
            
            DoEvents
            If Left$(filelistz, 5) <> "?????" Then
            
start:
            If Len(filelistz) >= 2 Then
            
            For j = 1 To Len(filelistz)
                temp = midz(filelistz, 1, InStr(1, filelistz, "/") - 1)
                If stopnow = 1 Then Exit Sub
                If InStr(1, UCase$(temp), UCase$(findword)) <> 0 Then
                    DoEvents
                    If stopnow = 1 Then Exit Sub
                    listz.ListItems.Add , , midz(filelistz, 1, InStr(1, filelistz, "/") - 1)
                    DoEvents
                    listz.ListItems.Item(listz.ListItems.count).ListSubItems.Add , , iplist.List(i)
                    DoEvents
                    
                    count = 0
                    For k = 1 To Len(filelistz)
                        If Mid$(filelistz, k, 1) = "?" Then count = count + 1
                    Next
                    
                    If count > 1 Then
                        listz.ListItems.Item(listz.ListItems.count).ListSubItems.Add , , midz(filelistz, InStr(1, filelistz, "/") + 1, InStr(1, filelistz, "?") - 1)
                        DoEvents
                    Else
                        listz.ListItems.Item(listz.ListItems.count).ListSubItems.Add , , midz(filelistz, InStr(1, filelistz, "/") + 1, Len(filelistz) - 2)
                        DoEvents
                    End If
                    
                    DoEvents
                    
                    filelistz = midz(filelistz, InStr(1, filelistz, "?") + 1, Len(filelistz))
                    j = 0
                    GoTo start
                Else
                    If stopnow = 1 Then Exit Sub
                    DoEvents
                    If InStr(1, filelistz, "?") = 0 Then Exit For
                    If InStr(1, UCase$(filelistz), UCase$(findword)) <> 0 Then
                        filelistz = midz(filelistz, InStr(1, filelistz, "?") + 1, Len(filelistz))
                        j = 0
                        GoTo start
                    Else
                        Exit For
                    End If
                End If
            Next
            
            End If
            End If
        Next
    End If
    
    Frame4.Enabled = True
    findnow.Caption = "검색"
    Exit Sub
    
    Else
        stopnow = 1
        Frame4.Enabled = True
        findnow.Caption = "검색"
    End If
    
    Exit Sub
    
err:
    MsgBox "서버에 연결되지 않았습니다." + vbCrLf + "아시죠? 실패작인거..ㅡㅡ;; 오류원인은 알고있으나 시간이 없는 관계로..ㅡㅡ;;" + vbCrLf + "15일부터 다시 새로 만들꺼니까 기다려주세용^^", , "오류"
    End
End Sub

Private Sub findnow2_Click()
    If findnow2.Caption = "검색" Then
            
    If rip = "" Then MsgBox "검색어를 입력해주세요.", , "오류": Exit Sub
    
    frame1.Enabled = False
    
    listz.ListItems.Clear
    iplist.Clear
    
    Dim filelistz As String, temp As String, count As Single
    
    findnow2.Caption = "중지"
    
    ftp.GetDirectory ("*.*")
    DoEvents
    stopnow = 0
    
    If iplist.ListCount <> 0 Then
        For i = 0 To iplist.ListCount - 1
            On Error Resume Next
            per.Value = (i + 1) / iplist.ListCount * 100
            If stopnow = 1 Then Exit Sub
'            MsgBox iplist.List(i) + "=>" + Str(InStr(1, iplist.List(i), rip))
            If InStr(1, iplist.List(i), rip) <> 0 Then
            'filelistz = inet.OpenURL("http://home.hanmir.com/~earus/" + iplist.List(i))
            
            If Dir("c:\filenara.tmp", vbNormal) <> "" Then Kill "c:\filenara.tmp"
            ftp.RemoteFile = iplist.List(i)
            ftp.LocalFile = "c:\filenara.tmp"
            ftp.GetFile
            DoEvents
            
            Do Until Dir("c:\filenara.tmp", vbNormal) <> ""
                DoEvents
            Loop
            
            Open "c:\filenara.tmp" For Input As #2
                Line Input #2, filelistz
            Close #2
            
            DoEvents
            If Left$(filelistz, 5) <> "?????" Then
            
start:
            If Len(filelistz) >= 2 Then
            For j = 1 To Len(filelistz)
                DoEvents
                If InStr(1, filelistz, "?") = 0 Then Exit For

                temp = midz(filelistz, 1, InStr(1, filelistz, "/") - 1)
                
                If stopnow = 1 Then Exit Sub
                
                listz.ListItems.Add , , midz(filelistz, 1, InStr(1, filelistz, "/") - 1)
                DoEvents
                listz.ListItems.Item(listz.ListItems.count).ListSubItems.Add , , iplist.List(i)
                DoEvents
                
                count = 0
                For k = 1 To Len(filelistz)
                    If Mid$(filelistz, k, 1) = "?" Then count = count + 1
                Next
                    
                If count > 1 Then
                    listz.ListItems.Item(listz.ListItems.count).ListSubItems.Add , , midz(filelistz, InStr(1, filelistz, "/") + 1, InStr(1, filelistz, "?") - 1)
                    DoEvents
                Else
                    listz.ListItems.Item(listz.ListItems.count).ListSubItems.Add , , midz(filelistz, InStr(1, filelistz, "/") + 1, Len(filelistz) - 2)
                    DoEvents
                End If
                    
                DoEvents
                    
                filelistz = midz(filelistz, InStr(1, filelistz, "?") + 1, Len(filelistz))
                j = 0
                GoTo start
            Next
            End If
            End If
            End If
        Next
    End If
    frame1.Enabled = True
    findnow2.Caption = "검색"
    Exit Sub
    
    Else
        stopnow = 1
        frame1.Enabled = True
        findnow2.Caption = "검색"
    End If

End Sub

Private Sub findword_GotFocus()
    findnow2.Default = False
    findnow.Default = True
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    If Dir("filenara.cfg", vbNormal) = "" Then
        Open "filenara.cfg" For Output As #1
            Print #1, "c:"
            Print #1, "c:\windows\바탕 화면"
            Print #1, 10
        Close #1
        MsgBox "공유할 폴더와 다운로드 받을 디렉토리를 지정해주시기 바랍니다.", , "파일나라 안내"
        tabz.Tab = 1
        DoEvents
    End If
    
    Dim gt As String
    Open "filenara.cfg" For Input As #1
        For i = 1 To 3
            Line Input #1, gt
            Select Case i
                Case 1
                    If Dir(gt, vbDirectory) <> "" Then
                        dirpath = gt
                    Else
                        MsgBox "공유폴더로 설정한 폴더가 없습니다." + vbCrLf + "재설정 바랍니다.(Default=바탕화면)", , "오류"
                        dirpath = "c:\windows\바탕 화면"
                    End If
                Case 2
                    If Dir(gt, vbDirectory) = "" Then
                        MsgBox "다운로드폴더로 설정한 폴더가 없습니다." + vbCrLf + "재설정 바랍니다.(Default=바탕화면)", , "오류"
                        Dir2.Path = "c:\windows\바탕 화면"
                        downpath = "c:\windows\바탕 화면"
                    Else
                        Dir2.Path = gt
                        downpath = gt
                    End If
                Case 3
                    maxcnt = Val(gt)
            End Select
        Next
    Close #1
    
    main.Show
    tabz.Enabled = False
    DoEvents
    
    Dim filez
    
    ftp.RemoteAddress = "ftp서버주소" '예)ftp.hanmir.com
    ftp.UserName = "아이디"           '예)merong
    ftp.Password = "암호"             '예)babo
    
    ftp.Connect
    DoEvents
    
    Dir1.Path = dirpath
    
        If Dir(dirpath, vbDirectory) <> "" Then
            filelist.Path = dirpath
            DoEvents
            If filelist.ListCount <> 0 Then
                For i = 0 To filelist.ListCount - 1
                    filez = filez + filelist.List(i) + "/" + Str(FileLen(filelist.Path + "\" + filelist.List(i))) + "?"
                Next
            Else
                filez = "?????"
            End If
        Else
            MsgBox "공유 디렉토리가 존재하지 않습니다.", , "오류"
        End If
        
        Open "c:\filelist.tmp" For Output As #1
            Print #1, filez
        Close #1
        
        DoEvents
        ftp.LocalFile = "C:\filelist.tmp"
        myip = inet.OpenURL("http://202.31.225.227:8080/webchat5/yourip.shtml")
        myip = nospace(midz(myip, InStr(1, myip, "<font color=blue size=4><b>") + 27, InStr(1, myip, "</b></font><br>") - 1))
        ftp.DeleteFile (myip)
        ftp.RemoteFile = myip
        ftp.PutFile
        DoEvents
        
        tabz.Enabled = True
        logon = "Connected to Server... Ready..."
        upload(0).LocalPort = 30303
        upload(0).Listen

        main.Refresh
        findword.SetFocus
    
    
    Dim msgtouser As String, version As String
    msgtouser = inet.OpenURL("http://my.dreamwiz.com/traderz/msgtouser.txt")
    version = inet.OpenURL("http://my.dreamwiz.com/traderz/version.txt")
    If Left$(version, 5) = "1.05b" Then
        MsgBox "현재 컴퓨터에 깔려있는 파일나라는 최신버젼이 아닙니다." + vbCrLf + "FILENARA.COM.NE.KR에 접속하셔서 " + version + "버젼을 다운받으세요.", , "버젼오류"
        Shell "start http://filenara.com.ne.kr/", vbHide
        End
    End If
    
    MsgBox msgtouser
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If maxcnt = "" Then maxcnt = "10"
    Open "filenara.cfg" For Output As #1
        Print #1, Dir1.Path
        Print #1, Dir2.Path
        Print #1, maxcnt
    Close #1
    
    Dim filez
    Cancel = 1
    filez = "?????"
    Open "c:\filelist.tmp" For Output As #1
        Print #1, filez
    Close #1
    ftp.LocalFile = "C:\filelist.tmp"
    myip = inet.OpenURL("http://202.31.225.227:8080/webchat5/yourip.shtml")
    myip = nospace(midz(myip, InStr(1, myip, "<font color=blue size=4><b>") + 27, InStr(1, myip, "</b></font><br>") - 1))
    ftp.DeleteFile (myip)
    ftp.RemoteFile = myip
'    ftp.PutFile
    DoEvents
    End
End Sub
Private Sub ftp_NextDirectoryEntry(ByVal FileName As String, ByVal Attributes As Long, ByVal Length As Double)
    If (Attributes And 16) <> 16 And Attributes <> 0 Then
        If findnow.Caption = "중지" Or findnow2.Caption = "중지" Then iplist.AddItem FileName
    End If
End Sub

Private Sub Label4_Click()
    Shell "start http://filenara.com.ne.kr", vbHide
End Sub

Private Sub listz_DblClick()
    Dim snum As Single
    If listz.ListItems.count <> 0 Then
        snum = listz.SelectedItem.Index
        If Dir("fn-down.exe", vbNormal) = "" Then MsgBox "다운로더 프로그램이 없습니다. 파일나라를 재설치하세요.", , "오류": Exit Sub
        Shell "fn-down " + downpath + "/" + listz.ListItems.Item(snum).ListSubItems.Item(1) + "/" + myip + "/" + listz.ListItems.Item(snum) + "/" + nospace(listz.ListItems.Item(snum).ListSubItems.Item(2)) + "/", vbNormalFocus
    End If
End Sub
Private Sub maxcnt_KeyPress(KeyAscii As Integer)
   If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
       KeyAscii = 0
   End If
End Sub
Private Sub nodisconnect_Timer()
    On Error Resume Next
    If Left$(logon, 1) = "=" Then
        MsgBox "서버에 연결할 수 없습니다. 잠시후 다시 시도해주세요.", , "로그온 오류"
        End
    End If
    
    ftp.GetDirectory ("*.*")
End Sub
Private Sub rip_GotFocus()
    findnow.Default = False
    findnow2.Default = True
End Sub

Private Sub upload_Close(Index As Integer)
    cntnum = cntnum - 1
End Sub
Private Sub upload_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If maxcnt = "" Then maxcnt = "10"
    If cntnum + 1 >= Val(maxcnt) Then Exit Sub
    upn = upn + 1
    cntnum = cntnum + 1
    Load upload(upn)
    upload(upn).Accept requestID
    'sendlist.ListItems.Add , , upload(upn).RemoteHost
End Sub
Private Sub upload_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim revdata As String
    
    upload(Index).GetData revdata
    
    If Left$(revdata, 11) = "download-ok" Then
        sendlist.ListItems.Item(Index).ListSubItems.Add 5, , "전송완료"
        upload(Index).Close
    End If
    
    If Left$(revdata, 3) = "ip:" Then
        sendlist.ListItems.Add , , midz(revdata, 4, Len(revdata))
        If onlyuser.Value = 1 Then
            If midz(revdata, 4, Len(revdata)) = onlyyou Then
                upload(Index).SendData "ok-your-ip"
            Else
                upload(Index).SendData "no-your-ip"
            End If
        Else
            upload(Index).SendData "ok-your-ip"
        End If
    End If
    
    If Left$(revdata, 9) = "filename:" Then
'        MsgBox dirpath + "\" + midz(revdata, 10, Len(revdata))
        sendlist.ListItems.Item(Index).ListSubItems.Add , , midz(revdata, 10, Len(revdata))
        If Dir(dirpath + "\" + midz(revdata, 10, Len(revdata)), vbNormal) <> "" Then
            sendlist.ListItems.Item(Index).ListSubItems.Add , , FileLen(dirpath + "\" + midz(revdata, 10, Len(revdata)))
            sendlist.ListItems.Item(Index).ListSubItems.Add , , "0"
            sendlist.ListItems.Item(Index).ListSubItems.Add , , "0%"
            sendlist.ListItems.Item(Index).ListSubItems.Add , , "전송중"
            upload(Index).SendData "findfile"
        Else
            sendlist.ListItems.Item(Index).ListSubItems.Add , , "파일없음"
            sendlist.ListItems.Item(Index).ListSubItems.Add , , "파일없음"
            sendlist.ListItems.Item(Index).ListSubItems.Add , , "0%"
            sendlist.ListItems.Item(Index).ListSubItems.Add , , "파일없음"
            upload(Index).SendData "nofile"
        End If
    End If
    
    '아이템-아이피
    '서브아이템1-파일명
    '서브아이템2-크기
    '서브아이템3-전송량
    '서브아이템4-전송률
    '서브아이템5-상태
    
    If revdata = "send file now" Then
'    MsgBox dirpath + "\" + sendlist.ListItems.Item(Index).ListSubItems.Item(1)
    Dim FreeFileNum As Single
    FreeFileNum = FreeFile
    Open dirpath + "\" + sendlist.ListItems.Item(Index).ListSubItems.Item(1) For Binary Access Read As #FreeFileNum
       
       For i = 1 To Int(FileLen(dirpath + "\" + sendlist.ListItems.Item(Index).ListSubItems.Item(1)) / 5000)
            On Error GoTo err
            If upload(Index).State = sckClosed Then
                Close #FreeFileNum
                sendlist.ListItems.Item(Index).ListSubItems.Add 5, , "연결끊김"
                Exit Sub
            End If
            ReDim bdata(1 To 5000) As Byte
            Get #FreeFileNum, , bdata
            DoEvents
            upload(Index).SendData bdata
            DoEvents
            sendlist.ListItems.Item(Index).ListSubItems.Item(3) = Str(Val(sendlist.ListItems.Item(Index).ListSubItems.Item(3)) + 5000)
            sendlist.ListItems.Item(Index).ListSubItems.Item(4) = nospace(Str(Int(Val(sendlist.ListItems.Item(Index).ListSubItems.Item(3)) / Val(sendlist.ListItems.Item(Index).ListSubItems.Item(2)) * 100)) + "%")
            DoEvents
       Next
       
       If FileLen(dirpath + "\" + sendlist.ListItems.Item(Index).ListSubItems.Item(1)) Mod 5000 <> 0 Then
            ReDim bdata(FileLen(dirpath + "\" + sendlist.ListItems.Item(Index).ListSubItems.Item(1)) Mod 5000)
            Get #FreeFileNum, , bdata
            upload(Index).SendData bdata
            DoEvents
            sendlist.ListItems.Item(Index).ListSubItems.Item(3) = Str(Val(sendlist.ListItems.Item(Index).ListSubItems.Item(3)) + Val(sendlist.ListItems.Item(Index).ListSubItems.Item(1)) Mod 5000)
            sendlist.ListItems.Item(Index).ListSubItems.Item(4) = "100%"
       End If
       
        
    Close #FreeFileNum
       
    End If
    
    Exit Sub
    
err:
    sendlist.ListItems.Item(Index).ListSubItems.Add 5, , "연결끊김"
    Exit Sub

End Sub

