VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CheckBox p56 
      Caption         =   "56"
      Height          =   255
      Left            =   2880
      TabIndex        =   27
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CheckBox p57 
      Caption         =   "57"
      Height          =   255
      Left            =   2880
      TabIndex        =   26
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CheckBox p55 
      Caption         =   "55"
      Height          =   255
      Left            =   2880
      TabIndex        =   25
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox p54 
      Caption         =   "54"
      Height          =   255
      Left            =   2880
      TabIndex        =   24
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CheckBox p53 
      Caption         =   "53"
      Height          =   255
      Left            =   2880
      TabIndex        =   23
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox p52 
      Caption         =   "52"
      Height          =   255
      Left            =   2880
      TabIndex        =   22
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CheckBox p51 
      Caption         =   "51"
      Height          =   255
      Left            =   2880
      TabIndex        =   21
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox p4F 
      Caption         =   "4F"
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox p50 
      Caption         =   "50"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CheckBox p4E 
      Caption         =   "4E"
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox p4d 
      Caption         =   "4D"
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox p4c 
      Caption         =   "4C"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox p4b 
      Caption         =   "4B"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox p4a 
      Caption         =   "4A"
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "READ"
      Height          =   375
      Left            =   720
      TabIndex        =   13
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "WRITE"
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CheckBox p40 
      Caption         =   "40"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox p41 
      Caption         =   "41"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox p42 
      Caption         =   "42"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CheckBox p43 
      Caption         =   "43"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox p45 
      Caption         =   "45"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CheckBox p44 
      Caption         =   "44"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox p46 
      Caption         =   "46"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CheckBox p47 
      Caption         =   "47"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox p48 
      Caption         =   "48"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox p49 
      Caption         =   "49"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   2160
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "|여기까지만쓰기구현|"
      Height          =   255
      Left            =   600
      TabIndex        =   29
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "읽기는 출력접전 전체범위,  쓰기는 8개만 구현해놨음"
      Height          =   180
      Left            =   120
      TabIndex        =   28
      Top             =   600
      Width           =   4290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Master-K "
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '단일 고정
      Caption         =   "수신데이터"
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    --  PLC 프로그램 (KGL 니모닉)  --
'
'  LOAD M000   -->P40이 On이되면   (16진수)
'  OUT P40     -->P40(접점40)을 On 시킨다. (16진수) 밑으로는 반복
'  LOAD M001
'  OUT P41
'  LOAD M002
'  OUT P42
'  LOAD M003
'  OUT P43
'  LOAD M004
'  OUT P44
'  LOAD M005
'  OUT P45
'  LOAD M006
'  OUT P46
'  LOAD M007
'  OUT P47
'  LOAD M008
'  OUT P48
'  LOAD M009
'  OUT P49
'  LOAD M00A
'  OUT P4A
'  LOAD M00B
'  OUT P4B
'  LOAD M00C
'  OUT P4C
'  LOAD M00D
'  OUT P4D
'  LOAD M00E
'  OUT P4E
'  LOAD M00F
'  OUT P4F
'  LOAD M010
'  OUT P50
'  LOAD M011
'  OUT P51
'  LOAD M012
'  OUT P52
'  LOAD M013
'  OUT P53
'  LOAD M014
'  OUT P54
'  LOAD M015
'  OUT P55
'  LOAD M016
'  OUT P56
'  LOAD M017
'  OUT P57
'  END



Private Sub Command1_Click()
Dim n20 As String '보낼데이터 기억하는 변수
Dim n19 As String
'~~~~ 집계
If (p43 = "0") And (p42 = "0") And (p41 = "0") And (p40 = "0") Then
n20 = "0"
ElseIf (p43 = "0") And (p42 = "0") And (p41 = "0") And (p40 = "1") Then
n20 = "1"
ElseIf (p43 = "0") And (p42 = "0") And (p41 = "1") And (p40 = "0") Then
n20 = "2"
ElseIf (p43 = "0") And (p42 = "0") And (p41 = "1") And (p40 = "1") Then
n20 = "3"
ElseIf (p43 = "0") And (p42 = "1") And (p41 = "0") And (p40 = "0") Then
n20 = "4"
ElseIf (p43 = "0") And (p42 = "1") And (p41 = "0") And (p40 = "1") Then
n20 = "5"
ElseIf (p43 = "0") And (p42 = "1") And (p41 = "1") And (p40 = "0") Then
n20 = "6"
ElseIf (p43 = "0") And (p42 = "1") And (p41 = "1") And (p40 = "1") Then
n20 = "7"
ElseIf (p43 = "1") And (p42 = "0") And (p41 = "0") And (p40 = "0") Then
n20 = "8"
ElseIf (p43 = "1") And (p42 = "0") And (p41 = "0") And (p40 = "1") Then
n20 = "9"
ElseIf (p43 = "1") And (p42 = "0") And (p41 = "1") And (p40 = "0") Then
n20 = "A"
ElseIf (p43 = "1") And (p42 = "0") And (p41 = "1") And (p40 = "1") Then
n20 = "B"
ElseIf (p43 = "1") And (p42 = "1") And (p41 = "0") And (p40 = "0") Then
n20 = "C"
ElseIf (p43 = "1") And (p42 = "1") And (p41 = "0") And (p40 = "1") Then
n20 = "D"
ElseIf (p43 = "1") And (p42 = "1") And (p41 = "1") And (p40 = "0") Then
n20 = "E"
ElseIf (p43 = "1") And (p42 = "1") And (p41 = "1") And (p40 = "1") Then
n20 = "F"
End If

'~~~~ 집계
If (p47 = "0") And (p46 = "0") And (p45 = "0") And (p44 = "0") Then
n19 = "0"
ElseIf (p47 = "0") And (p46 = "0") And (p45 = "0") And (p44 = "1") Then
n19 = "1"
ElseIf (p47 = "0") And (p46 = "0") And (p45 = "1") And (p44 = "0") Then
n19 = "2"
ElseIf (p47 = "0") And (p46 = "0") And (p45 = "1") And (p44 = "1") Then
n19 = "3"
ElseIf (p47 = "0") And (p46 = "1") And (p45 = "0") And (p44 = "0") Then
n19 = "4"
ElseIf (p47 = "0") And (p46 = "1") And (p45 = "0") And (p44 = "1") Then
n19 = "5"
ElseIf (p47 = "0") And (p46 = "1") And (p45 = "1") And (p44 = "0") Then
n19 = "6"
ElseIf (p47 = "0") And (p46 = "1") And (p45 = "1") And (p44 = "1") Then
n19 = "7"
ElseIf (p47 = "1") And (p46 = "0") And (p45 = "0") And (p44 = "0") Then
n19 = "8"
ElseIf (p47 = "1") And (p46 = "0") And (p45 = "0") And (p44 = "1") Then
n19 = "9"
ElseIf (p47 = "1") And (p46 = "0") And (p45 = "1") And (p44 = "0") Then
n19 = "A"
ElseIf (p47 = "1") And (p46 = "0") And (p45 = "1") And (p44 = "1") Then
n19 = "B"
ElseIf (p47 = "1") And (p46 = "1") And (p45 = "0") And (p44 = "0") Then
n19 = "C"
ElseIf (p47 = "1") And (p46 = "1") And (p45 = "0") And (p44 = "1") Then
n19 = "D"
ElseIf (p47 = "1") And (p46 = "1") And (p45 = "1") And (p44 = "0") Then
n19 = "E"
ElseIf (p47 = "1") And (p46 = "1") And (p45 = "1") And (p44 = "1") Then
n19 = "F"
End If

'~~~~전송
 MSComm1.CommPort = 1
  MSComm1.Settings = "19200,n,8,1"
  MSComm1.InputLen = 1
  MSComm1.PortOpen = True
   q = Chr(5) & "00WSS" + "02" + "06%MW001" & "FFFF" & "06%MW000" & "FF" & n19 & n20 & Chr(4)
'                                     임의데이터(FFFF FF)+n19+n20
 
   MSComm1.Output = q

Do
     instring = MSComm1.Input
     rcv = rcv & instring
     data = rcv
     
Loop Until instring = Chr(3)
     rcv = ""
     MSComm1.PortOpen = False

Label1 = data

End Sub

Private Sub Command2_Click()

Dim q As String
Dim rcv As String
Dim instring As String
Dim data  As String


p41 = "1"
On Error GoTo errmsg   '에러나면 errmsg로 이동후 대기...
  MSComm1.CommPort = 1
  MSComm1.Settings = "19200,n,8,1"
  MSComm1.InputLen = 1
  MSComm1.PortOpen = True
   q = Chr(5) & "00" + "RSS" + "02" + "06%MW001" + "06%MW000" & Chr(4) '06%PW001 06%PW000 두워드의 데이타 요청
'설명: 문의문자 국번 개별읽기 2개요청   06%MW001=> 뒤따라오는변수06자리(%포함)  M=M디바이스 W=워드 001번지...
'   (요청,명령의뜻)
   MSComm1.Output = q

Do
     instring = MSComm1.Input
     rcv = rcv & instring
     data = rcv
     
Loop Until instring = Chr(3)
     rcv = ""
     MSComm1.PortOpen = False

Label1 = data


'~~~~~~~~~~~~~~~판단
Select Case Mid(data, 20, 1)
Case "0"
p43 = "0": p42 = "0": p41 = "0": p40 = "0"
Case "1"
p43 = "0": p42 = "0": p41 = "0": p40 = "1"
Case "2"
p43 = "0": p42 = "0": p41 = "1": p40 = "0"
Case "3"
p43 = "0": p42 = "0": p41 = "1": p40 = "1"
Case "4"
p43 = "0": p42 = "1": p41 = "0": p40 = "0"
Case "5"
p43 = "0": p42 = "1": p41 = "0": p40 = "1"

Case "6"
p43 = "0": p42 = "1": p41 = "1": p40 = "0"
Case "7"
p43 = "0": p42 = "1": p41 = "1": p40 = "1"
Case "8"
p43 = "1": p42 = "0": p41 = "0": p40 = "0"
Case "9"
p43 = "1": p42 = "0": p41 = "0": p40 = "1"
Case "A"
p43 = "1": p42 = "0": p41 = "1": p40 = "0"

Case "B"
p43 = "1": p42 = "0": p41 = "1": p40 = "1"
Case "C"
p43 = "1": p42 = "1": p41 = "0": p40 = "0"
Case "D"
p43 = "1": p42 = "1": p41 = "0": p40 = "1"
Case "E"
p43 = "1": p42 = "1": p41 = "1": p40 = "0"
Case "F"
p43 = "1": p42 = "1": p41 = "1": p40 = "1"

End Select



'~~~~~~~~~~~~~~~판단
Select Case Mid(data, 19, 1)
Case "0"
p47 = "0": p46 = "0": p45 = "0": p44 = "0"
Case "1"
p47 = "0": p46 = "0": p45 = "0": p44 = "1"
Case "2"
p47 = "0": p46 = "0": p45 = "1": p44 = "0"
Case "3"
p47 = "0": p46 = "0": p45 = "1": p44 = "1"
Case "4"
p47 = "0": p46 = "1": p45 = "0": p44 = "0"
Case "5"
p47 = "0": p46 = "1": p45 = "0": p44 = "1"

Case "6"
p47 = "0": p46 = "1": p45 = "1": p44 = "0"
Case "7"
p47 = "0": p46 = "1": p45 = "1": p44 = "1"
Case "8"
p47 = "1": p46 = "0": p45 = "0": p44 = "0"
Case "9"
p47 = "1": p46 = "0": p45 = "0": p44 = "1"
Case "A"
p47 = "1": p46 = "0": p45 = "1": p44 = "0"

Case "B"
p47 = "1": p46 = "0": p45 = "1": p44 = "1"
Case "C"
p47 = "1": p46 = "1": p45 = "0": p44 = "0"
Case "D"
p47 = "1": p46 = "1": p45 = "0": p44 = "1"
Case "E"
p47 = "1": p46 = "1": p45 = "1": p44 = "0"
Case "F"
p47 = "1": p46 = "1": p45 = "1": p44 = "1"

End Select

'~~~~~~~~~~~~~~~판단
Select Case Mid(data, 18, 1)
Case "0"
p4b = "0": p4a = "0": p49 = "0": p48 = "0"
Case "1"
p4b = "0": p4a = "0": p49 = "0": p48 = "1"
Case "2"
p4b = "0": p4a = "0": p49 = "1": p48 = "0"
Case "3"
p4b = "0": p4a = "0": p49 = "1": p48 = "1"
Case "4"
p4b = "0": p4a = "1": p49 = "0": p48 = "0"
Case "5"
p4b = "0": p4a = "1": p49 = "0": p48 = "1"

Case "6"
p4b = "0": p4a = "1": p49 = "1": p48 = "0"
Case "7"
p4b = "0": p4a = "1": p49 = "1": p48 = "1"
Case "8"
p4b = "1": p4a = "0": p49 = "0": p48 = "0"
Case "9"
p4b = "1": p4a = "0": p49 = "0": p48 = "1"
Case "A"
p4b = "1": p4a = "0": p49 = "1": p48 = "0"

Case "B"
p4b = "1": p4a = "0": p49 = "1": p48 = "1"
Case "C"
p4b = "1": p4a = "1": p49 = "0": p48 = "0"
Case "D"
p4b = "1": p4a = "1": p49 = "0": p48 = "1"
Case "E"
p4b = "1": p4a = "1": p49 = "1": p48 = "0"
Case "F"
p4b = "1": p4a = "1": p49 = "1": p48 = "1"

End Select

'~~~~~~~~~~~~~~~판단
Select Case Mid(data, 17, 1)
Case "0"
p4F = "0": p4E = "0": p4d = "0": p4c = "0"
Case "1"
p4F = "0": p4E = "0": p4d = "0": p4c = "1"
Case "2"
p4F = "0": p4E = "0": p4d = "1": p4c = "0"
Case "3"
p4F = "0": p4E = "0": p4d = "1": p4c = "1"
Case "4"
p4F = "0": p4E = "1": p4d = "0": p4c = "0"
Case "5"
p4F = "0": p4E = "1": p4d = "0": p4c = "1"

Case "6"
p4F = "0": p4E = "1": p4d = "1": p4c = "0"
Case "7"
p4F = "0": p4E = "1": p4d = "1": p4c = "1"
Case "8"
p4F = "1": p4E = "0": p4d = "0": p4c = "0"
Case "9"
p4F = "1": p4E = "0": p4d = "0": p4c = "1"
Case "A"
p4F = "1": p4E = "0": p4d = "1": p4c = "0"

Case "B"
p4F = "1": p4E = "0": p4d = "1": p4c = "1"
Case "C"
p4F = "1": p4E = "1": p4d = "0": p4c = "0"
Case "D"
p4F = "1": p4E = "1": p4d = "0": p4c = "1"
Case "E"
p4F = "1": p4E = "1": p4d = "1": p4c = "0"
Case "F"
p4F = "1": p4E = "1": p4d = "1": p4c = "1"

End Select

Select Case Mid(data, 14, 1)
Case "0"
p53 = "0": p52 = "0": p51 = "0": p50 = "0"
Case "1"
p53 = "0": p52 = "0": p51 = "0": p50 = "1"
Case "2"
p53 = "0": p52 = "0": p51 = "1": p50 = "0"
Case "3"
p53 = "0": p52 = "0": p51 = "1": p50 = "1"
Case "4"
p53 = "0": p52 = "1": p51 = "0": p50 = "0"
Case "5"
p53 = "0": p52 = "1": p51 = "0": p50 = "1"

Case "6"
p53 = "0": p52 = "1": p51 = "1": p50 = "0"
Case "7"
p53 = "0": p52 = "1": p51 = "1": p50 = "1"
Case "8"
p53 = "1": p52 = "0": p51 = "0": p50 = "0"
Case "9"
p53 = "1": p52 = "0": p51 = "0": p50 = "1"
Case "A"
p53 = "1": p52 = "0": p51 = "1": p50 = "0"

Case "B"
p53 = "1": p52 = "0": p51 = "1": p50 = "1"
Case "C"
p53 = "1": p52 = "1": p51 = "0": p50 = "0"
Case "D"
p53 = "1": p52 = "1": p51 = "0": p50 = "1"
Case "E"
p53 = "1": p52 = "1": p51 = "1": p50 = "0"
Case "F"
p53 = "1": p52 = "1": p51 = "1": p50 = "1"

End Select


Select Case Mid(data, 13, 1)
Case "0"
p57 = "0": p56 = "0": p55 = "0": p54 = "0"
Case "1"
p57 = "0": p56 = "0": p55 = "0": p54 = "1"
Case "2"
p57 = "0": p56 = "0": p55 = "1": p54 = "0"
Case "3"
p57 = "0": p56 = "0": p55 = "1": p54 = "1"
Case "4"
p57 = "0": p56 = "1": p55 = "0": p54 = "0"
Case "5"
p57 = "0": p56 = "1": p55 = "0": p54 = "1"

Case "6"
p57 = "0": p56 = "1": p55 = "1": p54 = "0"
Case "7"
p57 = "0": p56 = "1": p55 = "1": p54 = "1"
Case "8"
p57 = "1": p56 = "0": p55 = "0": p54 = "0"
Case "9"
p57 = "1": p56 = "0": p55 = "0": p54 = "1"
Case "A"
p57 = "1": p56 = "0": p55 = "1": p54 = "0"

Case "B"
p57 = "1": p56 = "0": p55 = "1": p54 = "1"
Case "C"
p57 = "1": p56 = "1": p55 = "0": p54 = "0"
Case "D"
p57 = "1": p56 = "1": p55 = "0": p54 = "1"
Case "E"
p57 = "1": p56 = "1": p55 = "1": p54 = "0"
Case "F"
p57 = "1": p56 = "1": p55 = "1": p54 = "1"

End Select

errmsg:
Exit Sub
End Sub

