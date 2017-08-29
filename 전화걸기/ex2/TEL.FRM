VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmTel 
   Caption         =   "짱이 만든 전화걸기"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows 기본값
   Begin MSCommLib.MSComm Comm1 
      Left            =   120
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
   End
   Begin VB.Frame fraOption 
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   7095
      Begin VB.CommandButton cmdEsc 
         Caption         =   "취  소"
         Height          =   375
         Left            =   5880
         TabIndex        =   26
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "확  인"
         Height          =   375
         Left            =   5880
         TabIndex        =   25
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1440
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         Caption         =   " 통신포트 "
         Height          =   1455
         Left            =   2880
         TabIndex        =   9
         Top             =   1800
         Width           =   2655
         Begin VB.OptionButton optCom 
            Caption         =   "Com4"
            Height          =   375
            Index           =   3
            Left            =   1560
            TabIndex        =   24
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton optCom 
            Caption         =   "Com3"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   23
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton optCom 
            Caption         =   "Com2"
            Height          =   375
            Index           =   1
            Left            =   1560
            TabIndex        =   22
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optCom 
            Caption         =   "Com1"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " 데이타 길이 "
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   2655
         Begin VB.OptionButton optLen 
            Caption         =   "8 비트"
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   17
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optLen 
            Caption         =   "7 비트"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 패리티 "
         Height          =   975
         Left            =   2880
         TabIndex        =   7
         Top             =   720
         Width           =   2655
         Begin VB.OptionButton optParity 
            Caption         =   "우수"
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   20
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optParity 
            Caption         =   "기수"
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   19
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optParity 
            Caption         =   "없음"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 통신 속도 "
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2655
         Begin VB.OptionButton optSpeed 
            Caption         =   "38,400"
            Height          =   375
            Index           =   3
            Left            =   1440
            TabIndex        =   15
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton optSpeed 
            Caption         =   "28,800"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   14
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton optSpeed 
            Caption         =   "14,400"
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optSpeed 
            Caption         =   "9,600"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00400000&
         X1              =   5760
         X2              =   5760
         Y1              =   240
         Y2              =   3360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "전화걸기"
         Height          =   180
         Left            =   480
         TabIndex        =   10
         Top             =   280
         Width           =   720
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00400000&
         X1              =   120
         X2              =   5640
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.CommandButton cmdCansel 
      Caption         =   "취      소"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdOption 
      Caption         =   "설      정"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdTel 
      Caption         =   "전화걸기"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtTel 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      X1              =   5880
      X2              =   5880
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      X1              =   120
      X2              =   5760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "로 다이알을 돌립니다"
      Height          =   180
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1740
   End
End
Attribute VB_Name = "frmTel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Integer
Dim DB As Database
Dim ReModem As Recordset
Dim Rs, Rp, Rl, Rc, r As Integer

Private Sub cmdCansel_Click()
Dim Msg As String
Msg = MsgBox("전화걸기를 취소 하시겠습니까?", vbYesNo, "전화걸기")
If Msg = 6 Then
  Unload Me
Else
  Exit Sub
End If
End Sub

Private Sub cmdEsc_Click()
'옵션의 상태를 원래대로 한다
fraOption.Visible = False
Me.Height = 1485

End Sub

Private Sub cmdOk_Click()
'옵션의 상태를 저장후에 종료
'If ReModem.BOF = True And ReModem.EOF = True Then '레코드는 한개면 된다
   'ReModem.AddNew                    '레코드가 없으면
'Else
   ReModem.Edit                      '레코드가 있으면
'End If

For Rs = 0 To 3
  If optSpeed(Rs).Value = True Then
    Exit For
  End If
Next Rs

For Rp = 0 To 2
  If optParity(Rp).Value = True Then
    Exit For
  End If
Next Rp

For Rl = 0 To 1
  If optLen(Rl).Value = True Then
    Exit For
  End If
Next Rl

For Rc = 0 To 3
  If optCom(Rc).Value = True Then
    Exit For
  End If
Next Rc

ReModem!s = Rs
ReModem!p = Rp
ReModem!l = Rl
ReModem!C = Rc
ReModem.Update
fraOption.Visible = False
Me.Height = 1485

End Sub

Private Sub cmdOption_Click()
'옵션창을 보인다
SetModem
fraOption.Visible = True
Me.Height = 5220


End Sub

Private Sub cmdTel_Click()
If txtTel = "" Then
 MsgBox "전화번호를 입력하시오", vbInformation, "전화걸기"
 txtTel.SetFocus
 Exit Sub
End If
Dial txtTel


End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

'옵션상태를 보이지 않는다
fraOption.Visible = False
Me.Height = 1485

For I = 0 To 3
  If (I + 1) = Comm1.CommPort Then
    optCom(I).Value = True
    Exit For
  End If
Next I

Comm1.InputLen = 0


Dim Path
Path = App.Path
If Right(Path, 1) <> "\" Then
  Path = Path + "\"
End If

'모뎀설정이 DB에 저장 되었기 때문에..
Set DB = Workspaces(0).OpenDatabase(Path + "Modem.mdb")
Set ReModem = DB.OpenRecordset("Modem", dbOpenDynaset)

SetModem

End Sub


Private Sub Dial(Number As String)
Dim Dialstring As String
Dim fromModem As String
Dim dummy
Dim ComSpeed As Long
Dim ComLen As Integer
Dim ComParity As String
Dim ComPort As Integer

Dialstring = "ATDT" + Number + ";" + Chr(13)

For I = 0 To 3
  If optSpeed(I).Value = True Then
    Select Case I
      Case 0
       ComSpeed = 9600
      Case 1
       ComSpeed = 14400
      Case 2
       ComSpeed = 28800
      Case 3
       ComSpeed = 38400
    End Select
  End If
Next I

For I = o To 2
  If optParity(I).Value = True Then
    Select Case I
      Case 0
       ComParity = "N"
      Case 1
       ComParity = "O"
      Case 2
       ComParity = "E"
    End Select
  End If
Next I

For I = 0 To 1
  If optLen(I).Value = True Then
    Select Case I
     Case 0
      ComLen = 7
     Case 1
      ComLen = 8
    End Select
  End If
Next I

For I = 0 To 3
  If optCom(I).Value = True Then
    Select Case I
      Case 0
       ComPort = 1
      Case 1
       ComPort = 2
      Case 2
       ComPort = 3
      Case 3
       ComPort = 4
    End Select
  End If
Next I
'On Error GoTo ModemError     '에러발생하면
  
'통신포트의 설정( 전송속도,패리티,데이타 비트,중단비트)
Comm1.Settings = ComSpeed & "," & ComParity & "," _
                 & ComLen & "," & 1

'ModemError:
 'If Err = 380 Then   '에러번호가 380이면
  ' MsgBox "모뎀 설정이 않되었습니다." & Chr(13) & _
   '"설정 버튼을 누르신후 모뎀을 설정하시길...", vbInformation, "모뎀에러"
  ' cmdOption.SetFocus
  ' Exit Sub
' End If
 
  
Comm1.CommPort = ComPort
'통신포트를 연다
On Error Resume Next
Comm1.PortOpen = True
If Err Then
  MsgBox "현재의 포트를 사용할 수가 없습니다" & Chr(13) & _
         "다른 통신 포트를 사용하시오", vbInformation, "전화걸기"
  Exit Sub
End If

'입력 버퍼를 지운다
Comm1.InBufferCount = 0

'전화번호를 건다
Comm1.Output = Dialstring

'모뎀에서 결과를 기다린다
Do
 dummy = DoEvents()
 '버퍼에 데이타가 있으면 읽는다
 If Comm1.InBufferCount Then
   fromModem = fromModem + Comm1.Input
   'OK를 검사한다
   If InStr(fromModem, "OK") Then
    '사용자에게 전화기를 들도록 알려준다
     Beep
     MsgBox "전화기를 든 다음에 확인 버튼을 눌러주세요", vbInformation, "전화걸기"
     Exit Do
   End If
 End If
Loop

'모뎀을 차단한다
Comm1.Output = "ATH" + Chr(13)

'통신 포트를 닫는다
Comm1.PortOpen = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
 ReModem.Close
 DB.Close
End Sub

Private Sub SetModem()

Rs = ReModem!s
Rp = ReModem!p
Rl = ReModem!l
Rc = ReModem!C

optSpeed(Rs).Value = True
optParity(Rp).Value = True
optLen(Rl).Value = True
optCom(Rc).Value = True
End Sub

