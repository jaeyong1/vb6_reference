VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "PLC-1"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton Command3 
      Caption         =   "����͸� ����Ʈ"
      Height          =   615
      Left            =   240
      TabIndex        =   33
      Top             =   5520
      Width           =   3135
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5640
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�˶���� Reset"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   240
      TabIndex        =   29
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2760
      Top             =   3000
   End
   Begin VB.Frame Frame1 
      Caption         =   "* �˶���� *"
      Height          =   1575
      Left            =   3240
      TabIndex        =   24
      Top             =   2040
      Width           =   2775
      Begin VB.CommandButton Command2 
         Caption         =   "?"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1800
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         Caption         =   "������"
         Height          =   180
         Left            =   360
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "���"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label12 
         Height          =   255
         Left            =   2400
         TabIndex        =   30
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lbltel 
         BorderStyle     =   1  '���� ����
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1080
         Width           =   2295
      End
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   3120
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "��  ��"
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   5520
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Interval        =   11000
      Left            =   2760
      Top             =   2520
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1455
      Left            =   2520
      TabIndex        =   0
      Top             =   3840
      Width           =   4215
      ExtentX         =   7435
      ExtentY         =   2566
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label13 
      Caption         =   "Made by P.J.Y  2002. 02."
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   32
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  '���� ����
      Caption         =   " "
      Height          =   255
      Left            =   8400
      TabIndex        =   31
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblP12 
      Height          =   255
      Left            =   5400
      TabIndex        =   23
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblP11 
      Height          =   255
      Left            =   5400
      TabIndex        =   22
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblP10 
      Height          =   255
      Left            =   5400
      TabIndex        =   21
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblP09 
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblP08 
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblP07 
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblP06 
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblP05 
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblP02 
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblP01 
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "�������� ��� :"
      Height          =   180
      Left            =   3720
      TabIndex        =   13
      Top             =   1680
      Width           =   1260
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "��������2 ��� :"
      Height          =   180
      Left            =   3720
      TabIndex        =   12
      Top             =   1320
      Width           =   1530
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "��������1 ��� :"
      Height          =   180
      Left            =   3720
      TabIndex        =   11
      Top             =   960
      Width           =   1530
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "���������з� :"
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   1200
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "��������2 :"
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   1110
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "��������1 : "
      Height          =   180
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "���������з� : "
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "�������� :"
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "������ũ���� : "
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1260
   End
   Begin VB.Label LL 
      AutoSize        =   -1  'True
      Caption         =   "�칰���� : "
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BorderStyle     =   1  '���� ����
      Caption         =   "PLC ���ų��� : "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1485
   End
   Begin VB.Label lbldata 
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SMSObj As SMSCOMLib.SMSAPI
Private Sub Command1_Click()
End
End Sub



Private Sub Command2_Click()
Form2.Show 1
End Sub



Private Sub Command3_Click()
Shell ("C:\Program Files\Internet Explorer\IEXPLORE.EXE" & Space(1) & site + "monitor.php3")
End Sub

Private Sub Command4_Click()
Dim k As Integer
k = MsgBox("�˶��� �ٽ� �غ��Ű�ڽ��ϱ�?" & vbCrLf & "������ �ذ���� ���� ��Ȳ���� ������ ��ų��� �ڵ������ڸ޼����� �ٽ� �߼۵ǰ� �˴ϴ�. ", 32 + 4 + 256, "�˶�Reset")
If k = vbYes Then
    �������� = "OFF"
    t = 0
    Command4.Enabled = False
    Option1.Enabled = True
    Option2.Enabled = True
    Option1.SetFocus
    Label12 = ""
    lbltel = ""
End If
End Sub

Private Sub Form_Activate()


               site = "http://www.i-pws.com/sdwater/monitor/"
             ' ������Ʈ ����� ���� ������ �ٲ��ָ� ��.
             ' http�� �����ؼ� /���� �����ϴ� ��ü���� �ּҷ� ǥ��
             ' ������Ʈ�� ���α׷��� PHP�� ��� ������ �۹̼��� 777�� �ϸ� �۵���
             

w = 0
�������� = "OFF"

If �����ٿ�ε� = "����" Then
  Else
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ����Ʈ���� ��ȭ��ȣ �о ���
  lbltel = "Loading..0%"
  dial1 = Inet1.OpenURL(site + "dial1.jy")
  lbltel = "Loading..10%"
  dial2 = Inet1.OpenURL(site + "dial2.jy")
  lbltel = "Loading..20%"
  dial3 = Inet1.OpenURL(site + "dial3.jy")
  lbltel = "Loading..30%"
  dialcheck1 = Inet1.OpenURL(site + "dialcheck1.jy")
  lbltel = "Loading..40%"
  dialcheck2 = Inet1.OpenURL(site + "dialcheck2.jy")
  lbltel = "Loading..50%"
  dialcheck3 = Inet1.OpenURL(site + "dialcheck3.jy")
  lbltel = "Loading..60%"
  set1 = Inet1.OpenURL(site + "set1.jy")
  lbltel = "Loading..70%"
  set2 = Inet1.OpenURL(site + "set2.jy")
  lbltel = "Loading..80%"
  set3 = Inet1.OpenURL(site + "set3.jy")
  lbltel = "Loading..90%"
  set4 = Inet1.OpenURL(site + "set4.jy")
  lbltel = "Loading..100%"
  �����ٿ�ε� = "����"
  Command2.Enabled = True
  lbltel = ""
End If


End Sub


Private Sub Form_Load()  'sms�غ�

Set SMSObj = New SMSCOMLib.SMSAPI
Shell ("Regsvr32 c:\sdwater\SMSCOM.dll /s")


End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set SMSObj = Nothing
End Sub

Private Sub Option2_Click()
t = 0

End Sub

Private Sub Timer1_Timer()

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PLC���� ��źκ�
lbldata = ""
On Error GoTo errmsg   '�������� errmsg�� �̵��� ���...
  MSComm1.CommPort = 1
  MSComm1.Settings = "19200,n,8,1"
  MSComm1.InputLen = 1
  MSComm1.PortOpen = True
   q = Chr(5) & "00RSS0206%PW00106%PW000" & Chr(4) '06%PW001 06%PW000 �ο����� ����Ÿ ��û
   
   MSComm1.Output = q

Do
     instring = MSComm1.Input
     Rcv = Rcv & instring
     data = Rcv
Loop Until instring = Chr(3)
     Rcv = ""
     MSComm1.PortOpen = False
     
lbldata.Caption = data

     
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ���� ������ �м�
Select Case Mid(data, 20, 1)  '01 �칰���� üũ
Case "0" '������
lblP01 = "LOW"
p01 = "1"
Case "1" '�߼���
lblP01 = "MIDDLE"
p01 = "2"
Case "3" '�����
lblP01 = "HIGH"
p01 = "3"
Case "7" '�ʰ�
lblP01 = "OVER"
p01 = "4"
Case Else '�����̻�
lblP01 = "ERROR"
p01 = "0"
End Select

Select Case Mid(data, 19, 1)  '02 ������ũ���� üũ
Case "0" '������
lblP02 = "LOW": p02 = "1"
Case "1" '�߼���
lblP02 = "MIDDLE": p02 = "2"
Case "3" '�����
lblP02 = "HIGH": p02 = "3"
Case "7" '�ʰ�
lblP02 = "OVER": p02 = "4"
Case Else '�����̻�
lblP02 = "ERROR": p02 = "0"
End Select

Select Case Mid(data, 18, 1)  '05 06 07 08 üũ
Case "0"
p08 = "1": p07 = "1": p06 = "1": p05 = "1"
lblP08 = "OFF": lblP07 = "OFF": lblP06 = "OFF": lblP05 = "OFF"
Case "1"
lblP08 = "OFF": lblP07 = "OFF": lblP06 = "OFF": lblP05 = "ON"
p08 = "1": p07 = "1": p06 = "1": p05 = "2"
Case "2"
p08 = "1": p07 = "1": p06 = "2": p05 = "1"
lblP08 = "OFF": lblP07 = "OFF": lblP06 = "ON": lblP05 = "OFF"
Case "3"
p08 = "1": p07 = "1": p06 = "2": p05 = "2"
lblP08 = "OFF": lblP07 = "OFF": lblP06 = "ON": lblP05 = "ON"
Case "4"
p08 = "1": p07 = "2": p06 = "1": p05 = "1"
lblP08 = "OFF": lblP07 = "ON": lblP06 = "OFF": lblP05 = "OFF"
Case "5"
p08 = "1": p07 = "2": p06 = "1": p05 = "2"
lblP08 = "OFF": lblP07 = "ON": lblP06 = "OFF": lblP05 = "ON"

Case "6"
p08 = "1": p07 = "2": p06 = "2": p05 = "1"
lblP08 = "OFF": lblP07 = "ON": lblP06 = "ON": lblP05 = "OFF"
Case "7"
p08 = "1": p07 = "2": p06 = "2": p05 = "2"
lblP08 = "OFF": lblP07 = "ON": lblP06 = "ON": lblP05 = "ON"
Case "8"
p08 = "2": p07 = "1": p06 = "1": p05 = "1"
lblP08 = "ON": lblP07 = "OFF": lblP06 = "OFF": lblP05 = "OFF"
Case "9"
p08 = "2": p07 = "1": p06 = "1": p05 = "2"
lblP08 = "ON": lblP07 = "OFF": lblP06 = "OFF": lblP05 = "ON"
Case "A"
p08 = "2": p07 = "1": p06 = "2": p05 = "1"
lblP08 = "ON": lblP07 = "OFF": lblP06 = "ON": lblP05 = "OFF"

Case "B"
p08 = "2": p07 = "1": p06 = "2": p05 = "2"
lblP08 = "ON": lblP07 = "OFF": lblP06 = "ON": lblP05 = "ON"
Case "C"
p08 = "2": p07 = "2": p06 = "1": p05 = "1"
lblP08 = "ON": lblP07 = "ON": lblP06 = "OFF": lblP05 = "OFF"
Case "D"
p08 = "2": p07 = "2": p06 = "1": p05 = "2"
lblP08 = "ON": lblP07 = "ON": lblP06 = "OFF": lblP05 = "ON"
Case "E"
p08 = "2": p07 = "2": p06 = "2": p05 = "1"
lblP08 = "ON": lblP07 = "ON": lblP06 = "ON": lblP05 = "OFF"
Case "F"
p08 = "2": p07 = "2": p06 = "2": p05 = "2"
lblP08 = "ON": lblP07 = "ON": lblP06 = "ON": lblP05 = "ON"

End Select

Select Case Mid(data, 17, 1)  'P09 10 11 12 üũ
Case "0"
p12 = "0": p11 = "0": p10 = "0": p09 = "1"
lblP12 = "OFF": lblP11 = "OFF": lblP10 = "OFF": lblP09 = "OFF"
Case "1"
p12 = "0": p11 = "0": p10 = "0": p09 = "2"
lblP12 = "OFF": lblP11 = "OFF": lblP10 = "OFF": lblP09 = "ON"
Case "2"
p12 = "0": p11 = "0": p10 = "1": p09 = "1"
lblP12 = "OFF": lblP11 = "OFF": lblP10 = "ON": lblP09 = "OFF"
Case "3"
p12 = "0": p11 = "0": p10 = "1": p09 = "2"
lblP12 = "OFF": lblP11 = "OFF": lblP10 = "ON": lblP09 = "ON"
Case "4"
p12 = "0": p11 = "1": p10 = "0": p09 = "1"
lblP12 = "OFF": lblP11 = "ON": lblP10 = "OFF": lblP09 = "OFF"
Case "5"
p12 = "0": p11 = "1": p10 = "0": p09 = "2"
lblP12 = "OFF": lblP11 = "ON": lblP10 = "OFF": lblP09 = "ON"

Case "6"
p12 = "0": p11 = "1": p10 = "1": p09 = "1"
lblP12 = "OFF": lblP11 = "ON": lblP10 = "ON": lblP09 = "OFF"
Case "7"
p12 = "0": p11 = "1": p10 = "1": p09 = "2"
lblP12 = "OFF": lblP11 = "ON": lblP10 = "ON": lblP09 = "ON"
Case "8"
p12 = "1": p11 = "0": p10 = "0": p09 = "1"
lblP12 = "ON": lblP11 = "OFF": lblP10 = "OFF": lblP09 = "OFF"
Case "9"
p12 = "1": p11 = "0": p10 = "0": p09 = "2"
lblP12 = "ON": lblP11 = "OFF": lblP10 = "OFF": lblP09 = "ON"
Case "A"
p12 = "1": p11 = "0": p10 = "1": p09 = "1"
lblP12 = "ON": lblP11 = "OFF": lblP10 = "ON": lblP09 = "OFF"

Case "B"
p12 = "1": p11 = "0": p10 = "1": p09 = "2"
lblP12 = "ON": lblP11 = "OFF": lblP10 = "ON": lblP09 = "ON"
Case "C"
p12 = "1": p11 = "1": p10 = "0": p09 = "1"
lblP12 = "ON": lblP11 = "ON": lblP10 = "OFF": lblP09 = "OFF"
Case "D"
p12 = "1": p11 = "1": p10 = "0": p09 = "2"
lblP12 = "ON": lblP11 = "ON": lblP10 = "OFF": lblP09 = "ON"
Case "E"
p12 = "1": p11 = "1": p10 = "1": p09 = "1"
lblP12 = "ON": lblP11 = "ON": lblP10 = "ON": lblP09 = "OFF"
Case "F"
p12 = "1": p11 = "1": p10 = "1": p09 = "2"
lblP12 = "ON": lblP11 = "ON": lblP10 = "ON": lblP09 = "ON"

End Select


  '~~~~~~ ���ͳ����� �����ϱ� ���ؼ� ���������͸� ���ٷ� ���
  webQ = site + "plcwrite-1.php3?p01=" & p01 & "&p02=" & p02 & "&p05=" & p05 & "&p06=" & p06 & "&p07=" & p07 & "&p08=" & p08 & "&p09=" & p09 & "&p10=" & p10 & "&p11=" & p11 & "&p12=" & p12 & "&w=" & w & "&dial1=" & dial1 & "&dial2=" & dial2 & "&dial3=" & dial3 & "&dialcheck1=" & dialcheck1 & "&dialcheck2=" & dialcheck2 & "&dialcheck3=" & dialcheck3 & "&set1=" & set1 & "&set2=" & set2 & "&set3=" & set3 & "&dial3=" & dial3 & "&set4=" & set4


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ���ͳ����� ����
w = w + 1
If w = 10 Then
 w = 0: WebBrowser1.Navigate ("kr.yahoo.com")
Else
    If �����ٿ�ε� = "����" Then '�������¸� �� �޾ƿ��� �������� �����͸� �������� ����..
     WebBrowser1.Navigate (webQ)
    End If
End If
'~~~~~~~~~~~~~~~~~~-<<<<  �˶���� >>>>>
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ �˶����
If �������� = "OFF" Then

 If (p01 = "1") Or (p02 = "1") Or (p10 = "1") Or (p11 = "1") Or (p12 = "1") Then
    If Option1.Value = True Then

    �������� = "ON"
    Command4.Enabled = True
    'Timer2.Enabled = True '<-  ��ȭ�˶� �۵�
             '~~~~~~~ SMS�޼��� �Ǵ� �� ���ۺκ�
             If p01 = "1" Then
                smsmsg = smsmsg & "�칰������"
             End If
             
             If p02 = "1" Then
                         smsmsg = smsmsg & "������ũ������"
             End If
             
             If p10 = "1" Then
                         smsmsg = smsmsg & "��������1������"
             End If
             
             If p11 = "1" Then
                         smsmsg = smsmsg & "��������2������"
            End If
             
             If p12 = "1" Then
                         smsmsg = smsmsg & "��������������"
            End If
             smsmsg = "*[�ﵵ�������]*�˶�����-" + smsmsg
                        If LenB(smsmsg) > 78 Then '---- �������۳����� 80�ڰ� �ʰ��Ǿ�����..
                        smsmsg = "*[�ﵵ�������]*�˶�����-�������溸�߻�Ȯ���ʿ�"
                        End If
              
               If (Left(dial1, 2) = "01") And (dialcheck1 = "1") Then '01�� �����ϴ� ����ó and SMSüũ����
                  lbltel = " ����ó1 ����������"
                  SMSObj.ReCallNum = "9999" '�������޴���
                  SMSObj.SendSMS dial1, smsmsg
                  
               End If
               
               If (Left(dial2, 2) = "01") And (dialcheck2 = "1") Then
                  lbltel = " ����ó2 ����������"
                  SMSObj.ReCallNum = "9999" '�������޴���
                  SMSObj.SendSMS dial2, smsmsg
               End If
               If (Left(dial3, 2) = "01") And (dialcheck3 = "1") Then
                  lbltel = " ����ó3 ����������"
                  SMSObj.ReCallNum = "9999" '�������޴���
                  SMSObj.SendSMS dial3, smsmsg
               End If
               lbltel = " ALERT"
             smsmsg = ""
             '~~~~~~~ SMS��
          
    Option2.Enabled = False
    Option1.Enabled = False
    
    'Label12 = "T"
    End If
 End If

End If



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ���α׷��� �����߻��� �޼���â �۵�
Exit Sub
errmsg:
  MsgBox "���α׷��۵��߿� ������ �߻��߽��ϴ�.", 0, "����~"

End Sub

