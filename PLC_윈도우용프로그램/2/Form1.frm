VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "PLC-2"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton Command3 
      Caption         =   "����͸� ����Ʈ"
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   5280
      Width           =   2295
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   1200
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�˶���� Reset"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   3120
      TabIndex        =   14
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "* �˶���� *"
      Height          =   1575
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   2655
      Begin VB.OptionButton Option1 
         Caption         =   "���"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "������"
         Height          =   180
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "?"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbltel 
         BorderStyle     =   1  '���� ����
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label12 
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   255
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "��  ��"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   12000
      Left            =   1920
      Top             =   5760
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   4935
      ExtentX         =   8705
      ExtentY         =   2778
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
   Begin MSCommLib.MSComm MSComm1 
      Left            =   600
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label lblp16 
      BorderStyle     =   1  '���� ����
      Caption         =   "visible �� ����� ����"
      Height          =   615
      Left            =   4200
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "����� P.H. :"
      Height          =   180
      Left            =   3120
      TabIndex        =   21
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "��ǰ����P.H. : "
      Height          =   180
      Left            =   3000
      TabIndex        =   20
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "��  Ÿ : "
      Height          =   180
      Left            =   3600
      TabIndex        =   19
      Top             =   1440
      Width           =   660
   End
   Begin VB.Label lblP13 
      Height          =   255
      Left            =   4440
      TabIndex        =   18
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblP14 
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblP15 
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   1440
      Width           =   975
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
      Left            =   3360
      TabIndex        =   15
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "�����ũ���� : "
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1260
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lbldata 
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   240
      Width           =   3135
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
      TabIndex        =   3
      Top             =   240
      Width           =   1485
   End
   Begin VB.Label LL 
      AutoSize        =   -1  'True
      Caption         =   "������ũ���� : "
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1260
   End
   Begin VB.Label lblP01 
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SMSObj As SMSCOMLib.SMSAPI

Private Sub Command3_Click()
Shell ("C:\Program Files\Internet Explorer\IEXPLORE.EXE" & Space(1) & site + "monitor.php3")
End Sub

Private Sub Form_Load()
Set SMSObj = New SMSCOMLib.SMSAPI
Shell ("Regsvr32 c:\sdwater\SMSCOM.dll /s")
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set SMSObj = Nothing
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Form2.Show 1
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
             '
             
w = 0
�������� = "OFF"
If �����ٿ�ε� = "����" Then

  Else
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ����Ʈ���� �������� �޾Ƽ� ���
  lbltel = "Loading..0%"
  dial1 = Inet1.OpenURL(site + "dial1.jy")
  lbltel = "Loading..20%"
  dial2 = Inet1.OpenURL(site + "dial2.jy")
  lbltel = "Loading..40%"
  dial3 = Inet1.OpenURL(site + "dial3.jy")
  lbltel = "Loading..60%"
  dialcheck1 = Inet1.OpenURL(site + "dialcheck1.jy")
  lbltel = "Loading..80%"
  dialcheck2 = Inet1.OpenURL(site + "dialcheck2.jy")
  lbltel = "Loading..90%"
  dialcheck3 = Inet1.OpenURL(site + "dialcheck3.jy")
  lbltel = "Loading..100%"
  �����ٿ�ε� = "����"
  Command2.Enabled = True
  lbltel = ""
End If

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
Select Case Mid(data, 20, 1)  '
Case "0" '������
lblP01 = "LOW"
p03 = "1"
Case "1" '�߼���
lblP01 = "MIDDLE"
p03 = "2"
Case "3" '�����
lblP01 = "HIGH"
p03 = "3"
Case "7" '�ʰ�
lblP01 = "OVER"
p03 = "4"
Case Else '�����̻�
lblP01 = "ERROR"
p03 = "0"
End Select

Select Case Mid(data, 19, 1)  '
Case "0" '������
Label1 = "LOW": p04 = "1"
Case "1" '�߼���
Label1 = "MIDDLE": p04 = "2"
Case "3" '�����
Label1 = "HIGH": p04 = "3"
Case "7" '�ʰ�
Label1 = "OVER": p04 = "4"
Case "4" '�� ���� ������ �۵��ص� over ������ ��û..
Label1 = "OVER": p04 = "4"
Case Else '�����̻�
Label1 = "ERROR": p04 = "0"
End Select


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ �߰��� �κ� on=1 off=0
Select Case Mid(data, 18, 1)  '13 14 15 üũ / 16�� �������� ��ġ
Case "0"
p16 = "0": p15 = "0": p14 = "0": p13 = "0"
lblp16 = "OFF": lblP15 = "OFF": lblP14 = "OFF": lblP13 = "OFF"
Case "1"
lblp16 = "OFF": lblP15 = "OFF": lblP14 = "OFF": lblP13 = "ON"
p16 = "0": p15 = "0": p14 = "0": p13 = "1"
Case "2"
p16 = "0": p15 = "0": p14 = "1": p13 = "0"
lblp16 = "OFF": lblP15 = "OFF": lblP14 = "ON": lblP13 = "OFF"
Case "3"
p16 = "0": p15 = "0": p14 = "1": p13 = "1"
lblp16 = "OFF": lblP15 = "OFF": lblP14 = "ON": lblP13 = "ON"
Case "4"
p16 = "0": p15 = "1": p14 = "0": p13 = "0"
lblp16 = "OFF": lblP15 = "ON": lblP14 = "OFF": lblP13 = "OFF"
Case "5"
p16 = "0": p15 = "1": p14 = "0": p13 = "1"
lblp16 = "OFF": lblP15 = "ON": lblP14 = "OFF": lblP13 = "ON"

Case "6"
p16 = "0": p15 = "1": p14 = "1": p13 = "0"
lblp16 = "OFF": lblP15 = "ON": lblP14 = "ON": lblP13 = "OFF"
Case "7"
p16 = "0": p15 = "1": p14 = "1": p13 = "1"
lblp16 = "OFF": lblP15 = "ON": lblP14 = "ON": lblP13 = "ON"
Case "8"
p16 = "1": p15 = "0": p14 = "0": p13 = "0"
lblp16 = "ON": lblP15 = "OFF": lblP14 = "OFF": lblP13 = "OFF"
Case "9"
p16 = "1": p15 = "0": p14 = "0": p13 = "1"
lblp16 = "ON": lblP15 = "OFF": lblP14 = "OFF": lblP13 = "ON"
Case "A"
p16 = "1": p15 = "0": p14 = "1": p13 = "0"
lblp16 = "ON": lblP15 = "OFF": lblP14 = "ON": lblP13 = "OFF"

Case "B"
p16 = "1": p15 = "0": p14 = "1": p13 = "1"
lblp16 = "ON": lblP15 = "OFF": lblP14 = "ON": lblP13 = "ON"
Case "C"
p16 = "1": p15 = "1": p14 = "0": p13 = "0"
lblp16 = "ON": lblP15 = "ON": lblP14 = "OFF": lblP13 = "OFF"
Case "D"
p16 = "1": p15 = "1": p14 = "0": p13 = "1"
lblp16 = "ON": lblP15 = "ON": lblP14 = "OFF": lblP13 = "ON"
Case "E"
p16 = "1": p15 = "1": p14 = "1": p13 = "0"
lblp16 = "ON": lblP15 = "ON": lblP14 = "ON": lblP13 = "OFF"
Case "F"
p16 = "1": p15 = "1": p14 = "1": p13 = "1"
lblp16 = "ON": lblP15 = "ON": lblP14 = "ON": lblP13 = "ON"

End Select

'~~~~ ���ͳ����� �����ϱ� ���ؼ� ���������͸� ���ٷ� ���
webQ = site + "plcwrite-2.php3?p03=" & p03 & "&p04=" & p04 & "&p13=" & p13 & "&p14=" & p14 & "&p15=" & p15



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ���ͳ����� ����

w = w + 1
If w = 10 Then
 w = 0: WebBrowser1.Navigate ("kr.yahoo.com")
Else
   If �����ٿ�ε� = "����" Then '�������¸� �� �޾ƿ��� �������� �����͸� �������� ����.
    WebBrowser1.Navigate (webQ)
   End If
End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ �˶����
If �������� = "OFF" Then
 If (p03 = "4") Or (p04 = "4") Or (p13 = "1") Then
    If Option1.Value = True Then
    
    �������� = "ON"
    Command4.Enabled = True
    'Timer2.Enabled = True '< -��ȭ�˶��۵�------------------------------------------������
             '~~~~~~~ SMS�޼��� �Ǵ� �� ���ۺκ�
             If p03 = "4" Then
                          smsmsg = smsmsg & "������ũ�����"
             End If
             
             If p04 = "4" Then
                         smsmsg = smsmsg & "���Ͽ�����������"
             End If
             
             If p13 = "1" Then
                         smsmsg = smsmsg & "�����PH���"
             End If
             
             smsmsg = "*[�ﵵ�������]*�˶�����-" + smsmsg
             
             If LenB(smsmsg) > 78 Then '---- �������۳����� 80�ڰ� �ʰ��Ǿ�����..
             smsmsg = "*[�ﵵ�������]*�˶�����-���������߻�Ȯ���ʿ�"
             End If
             
               If (Left(dial1, 2) = "01") And (dialcheck1 = "1") Then
                  lbltel = " ����ó1������ ����������"
                  SMSObj.ReCallNum = "9999" '�������޴���
                  SMSObj.SendSMS dial1, smsmsg
               End If
               
               If (Left(dial2, 2) = "01") And (dialcheck2 = "1") Then
                  lbltel = " ����ó2������ ����������"
                  SMSObj.ReCallNum = "9999" '�������޴���
                  SMSObj.SendSMS dial2, smsmsg
               End If
               
               If (Left(dial3, 2) = "01") And (dialcheck3 = "1") Then
                  lbltel = " ����ó3������ ����������"
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
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ �����߻��� �޼���â �۵�
Exit Sub
errmsg:
 MsgBox "���α׷��۵��߿� ������ �߻��߽��ϴ�.", 0, "����~"
 
 
 
End Sub

