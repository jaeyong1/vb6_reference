VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  '���� ����
   Caption         =   "Chatting Craft-Client"
   ClientHeight    =   4776
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6552
   LinkTopic       =   "Form1"
   ScaleHeight     =   4776
   ScaleWidth      =   6552
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.TextBox nick_bar 
      Height          =   372
      Left            =   960
      TabIndex        =   11
      Top             =   4920
      Width           =   1812
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6120
      Top             =   4320
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.TextBox port_bar 
      Height          =   372
      Left            =   960
      TabIndex        =   8
      Text            =   "2000"
      Top             =   6120
      Width           =   1812
   End
   Begin VB.TextBox ip_bar 
      Height          =   372
      Left            =   960
      TabIndex        =   7
      Text            =   "127.0.0.1"
      Top             =   5520
      Width           =   1812
   End
   Begin VB.CommandButton exit_com 
      Caption         =   "���α׷�����"
      Height          =   492
      Left            =   2880
      TabIndex        =   6
      Top             =   6000
      Width           =   3612
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   132
   End
   Begin VB.CommandButton dis_com 
      Caption         =   "��   ��"
      Height          =   492
      Left            =   4680
      TabIndex        =   3
      Top             =   5520
      Width           =   1812
   End
   Begin VB.CommandButton connect_com 
      Caption         =   "��   ��"
      Height          =   492
      Left            =   2880
      TabIndex        =   2
      Top             =   5520
      Width           =   1812
   End
   Begin VB.TextBox chat_bar 
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   6372
   End
   Begin VB.TextBox chat_win 
      Height          =   3492
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   0
      Top             =   120
      Width           =   6372
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      Caption         =   "��ȭ��"
      Height          =   252
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   732
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   "��Ʈ����"
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   732
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "IP�Է�â"
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   732
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6360
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "����â����"
      Height          =   372
      Left            =   360
      TabIndex        =   5
      Top             =   4440
      Width           =   1452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim check As Integer
Dim ret, gointo, gotoval, i, counter As Integer
Dim tempid As String
Dim indexid As String


Private Sub dis_com_Click() '�����ư�� �������

 
 For i = 1 To 2
   If i = 1 Then            '���ⵥ���͸� ���� �������� ������

        Winsock1.SendData "*X*"
        DoEvents
        

   Else                     'Ŭ���̾�Ʈ�� �������ݴ´�

       Winsock1.Close
       connect_com.Enabled = True
       dis_com.Enabled = False
       ret = MsgBox("������ ���������ϴ�", 64, "�������")
       addtext "#�������� ������ ���������ϴ�#"

   End If
Next i
 



End Sub

Sub Startrek(frm As Form)  '���α׷��ݱ� �ִϸ��̼�

gotoval = Form1.Height / 2

For gointo = 1 To gotoval  ' ���� �������̸� ���δ�

DoEvents
Form1.Height = Form1.Height - 100
Form1.Top = (Screen.Height - Form1.Height) \ 2
If Form1.Height <= 500 Then Exit For

Next gointo

horiz:
Form1.Height = 30
gotoval = Form1.Width / 2

For gointo = 1 To gotoval  '���� ������̸� ���δ�

DoEvents
Form1.Width = Form1.Width - 100
Form1.Left = (Screen.Width - Form1.Width) \ 2
If Form1.Width <= 2000 Then Exit For

Next gointo

End Sub


Private Sub exit_com_Click() '���α׷����� ��ư�� �������

Call Startrek(Me)
End

End Sub


Private Sub connect_com_Click() '�����ư�� �������

If ip_bar.Text = "" Or nick_bar.Text = "" Then  '��ȭ��� ip�ּҰ� �Էµ��� �ʾ������
   
   ret = MsgBox("������IP�ּҿ� ��ȭ���� �Է¿��", 64, "�������")

Else                                             '�������ӽõ�

   If port_bar <> 2000 Then                      'port��ȣ�� �ٲ���
        ret = MsgBox("Port��ȣ��" & port_bar.Text & "�κ���Ǿ����ϴ�", 64, "�������")
        Winsock1.RemotePort = port_bar.Text
   End If
   Winsock1.RemoteHost = ip_bar.Text             '������ ������ �õ��Ѵ�
   Winsock1.Connect
   connect_com.Enabled = False
   dis_com.Enabled = True
   addtext "#������ ���������� �����߽��ϴ�#"
   
End If

End Sub

Private Sub Form_Load()         '���ε�� ������Ʈ��ȣ �ʱ�ȭ

Winsock1.RemotePort = 2000

End Sub

Private Sub chat_bar_keyPress(keyascii As Integer)     'ä�ø޼��� �������


If keyascii = 13 And dis_com.Enabled = True Then       '������¿��� ���͸� �������
                                                  
      If tempid <> "" And tempid <> nick_bar.Text Then '��ȭ���� ����ɰ�� ����޼��� ���
         
            Winsock1.SendData "##########" + tempid + "�� " + nick_bar.Text + "�� ��ȭ���� ����Ǿ����ϴ�" + "#########" + vbNewLine
            addtext "##########" + tempid + "�� " + nick_bar.Text + "�� ��ȭ���� ����Ǿ����ϴ�" + "#########"
  
      End If                                           'ä�ø޼��� ���۰� â�� ǥ��
        
            tempid = nick_bar.Text
            Winsock1.SendData nick_bar.Text + ">>" + chat_bar.Text
            addtext nick_bar.Text + ">>" + chat_bar.Text
            chat_bar.Text = ""
      
      
End If

End Sub


Private Sub addtext(addline As String) 'ȭ�鿡 ��½�Ų��

chat_win.Text = chat_win.Text + addline + vbNewLine

End Sub

Private Sub Winsock1_Close()           '������ �ݴ´�

dis_com_Click

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
                                       '���������� �����͸� �޴´�
Dim getda As String
Winsock1.GetData getda

If Left(getda, 3) <> "!!!" Then


Select Case getda

Case "boot"  '�������� ��ɾ ���ð��

  ret = MsgBox("������ ���� �����찡 �ڵ�����˴ϴ�", 52, "�������")
  
  If ret = 6 Then 'Ȯ�ι�ư
       Call ExitWindowsEx(EWX_SHUTDOWN, 0) '�����������Լ�ȣ��
  End If
   
Case "logoff" '���ݷα׿��� ��ɾ ���ð��
  
  ret = MsgBox("������ ���� �����찡 �ڵ��α׿����˴ϴ�", 52, "�������")
  
  If ret = 6 Then 'Ȯ�ι�ư
       Call ExitWindowsEx(EWX_LOGOFF, 0) '������α׿����Լ�ȣ��
  End If
   
Case "reboot" '��������� ��ɾ ���ð��
  
  ret = MsgBox("������ ���� �����찡 �ڵ�����۵˴ϴ�", 52, "�������")
  
  If ret = 6 Then 'Ȯ�ι�ư
       Call ExitWindowsEx(EWX_REBOOT, 0) '������������Լ�ȣ��
  End If
   

Case "#!#!"  '������������ ����� ���ð��

       Winsock1.Close
       connect_com.Enabled = True
       dis_com.Enabled = False
       ret = MsgBox("������ ���������ϴ�", 64, "�������")
       addtext "#�������� ������ ���������ϴ�#"

Case Else 'ä�ø޼����� ���ð��

addtext getda

End Select

Else

indexid = CStr(Right(getda, 1))

End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
                                        '���ӽ� �����߻�
ret = MsgBox("����� �����߻��߽��ϴ� �ٽ� �õ��ϼ���", 64, "�������")
Winsock1.Close
connect_com.Enabled = True
dis_com.Enabled = False

End Sub

Private Sub Check1_Click()  '����â�� ����

If check = 0 Then     '����â�� �����
      Form1.Height = 6972
      check = check + 1
ElseIf check = 1 Then '����â�� �������
      Form1.Height = 5148
      check = 0
End If
 
End Sub
