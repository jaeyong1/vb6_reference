VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BackColor       =   &H0080FF80&
   Caption         =   "Chatting Craft-Server"
   ClientHeight    =   4788
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8268
   LinkTopic       =   "Form2"
   ScaleHeight     =   4788
   ScaleWidth      =   8268
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CheckBox sCheck2 
      Caption         =   "Check1"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   5760
      Width           =   132
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   10
      Left            =   3000
      Top             =   4320
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.CommandButton slogoff_com 
      Caption         =   "���ݷα׿���"
      Height          =   492
      Left            =   3120
      TabIndex        =   11
      Top             =   5520
      Width           =   1692
   End
   Begin VB.CommandButton sreboot_com 
      Caption         =   "���������"
      Height          =   492
      Left            =   3120
      TabIndex        =   10
      Top             =   5040
      Width           =   1692
   End
   Begin VB.TextBox snick_bar 
      Height          =   372
      Left            =   720
      TabIndex        =   8
      Top             =   5040
      Width           =   2172
   End
   Begin VB.CommandButton sboot_com 
      Caption         =   "��������������"
      Height          =   492
      Left            =   4800
      TabIndex        =   7
      Top             =   5520
      Width           =   1692
   End
   Begin VB.CommandButton Sconnect_com 
      Caption         =   "������"
      Height          =   492
      Left            =   4800
      TabIndex        =   6
      Top             =   5040
      Width           =   1692
   End
   Begin VB.CommandButton sexit_com 
      Caption         =   "���α׷�����"
      Height          =   492
      Left            =   6480
      TabIndex        =   5
      Top             =   5520
      Width           =   1692
   End
   Begin VB.CommandButton sdis_com 
      Caption         =   "��   ��"
      Height          =   492
      Left            =   6480
      TabIndex        =   4
      Top             =   5040
      Width           =   1692
   End
   Begin VB.CheckBox sCheck1 
      Caption         =   "Check1"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   132
   End
   Begin VB.TextBox schat_bar 
      Height          =   372
      Left            =   0
      TabIndex        =   1
      Top             =   3840
      Width           =   8172
   End
   Begin VB.TextBox schat_win 
      ForeColor       =   &H80000001&
      Height          =   3492
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   0
      Top             =   120
      Width           =   8172
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "����������"
      Height          =   252
      Left            =   480
      TabIndex        =   13
      Top             =   5760
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "��ȭ��"
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      Width           =   852
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8040
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "����â����"
      Height          =   252
      Left            =   360
      TabIndex        =   3
      Top             =   4440
      Width           =   972
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private connectindex(10) As Integer
Dim sscheck, gointo, gotoval As Integer
Dim j, i, countp As Integer
Dim tempid, tempstring As String

Private Sub Form_Load() ' ���� �ε�ɶ� �ʱ�ȭ �ϴ� �κ� ����ư�� �ʱ�ȭ

sdis_com.Enabled = False
Sconnect_com.Enabled = True
sreboot_com.Enabled = False
sboot_com.Enabled = False
slogoff_com.Enabled = False
sexit_com.Enabled = True

For i = 1 To 10         '��Ƽü���� ���� ������ ��뿩��Ȯ�� �迭 �ʱ�ȭ
    
    connectindex(i) = 0

Next i


End Sub

Private Sub sboot_com_Click() '���ݼ˴ٿ��ư�� ������

For i = 1 To 10               '���� ����Ǿ��ִ� ������ �˻��Ͽ� �˴ٿ�޽��� ����
    
    If connectindex(i) = 1 Then
          
          Winsock1(i).SendData "boot"
          DoEvents
    
    End If

Next i

End Sub

Private Sub sCheck1_Click() '����â�� ����

If sscheck = 0 Then         'ù��° üũ��(����)
   
   Form2.Height = 6684
   sscheck = sscheck + 1
 
 ElseIf sscheck = 1 Then    '�ι�° üũ��(����)
   
   Form2.Height = 5184
   sscheck = 0
 
 End If

End Sub

Private Sub sCheck2_Click() '���������ư�� ������� �Ҷ� ����ư ����

If sscheck = 0 Then         'ù��° üũ��(����)
   
      sreboot_com.Enabled = False
      sboot_com.Enabled = False
      slogoff_com.Enabled = False
      sscheck = sscheck + 1
 
 ElseIf sscheck = 1 Then
      
      sreboot_com.Enabled = True
      sboot_com.Enabled = True
      slogoff_com.Enabled = True
      sscheck = 0
 
 End If

End Sub

Private Sub Sconnect_com_Click() '������ ������

Load Winsock1(0)                  '������ ������ �����ϰ� ��Ʈ��2000������ �Ѵ�
Winsock1(0).Protocol = sckTCPProtocol
Winsock1(0).LocalPort = 2000
Winsock1(0).Listen                'Ŭ���̾�Ʈ ���Ӵ�����
sdis_com.Enabled = True
Sconnect_com.Enabled = False
addtext "#Ŭ���̾�Ʈ�� ������ ������Դϴ�#"

End Sub

Sub Startrek(frm As Form)  '����� â ���ϸ��̼ǽ���

gotoval = Form2.Height / 2

For gointo = 1 To gotoval  '�������� âũ�⸦ ���δ�
   
   DoEvents
   Form2.Height = Form2.Height - 100
   Form2.Top = (Screen.Height - Form2.Height) \ 2
   If Form2.Height <= 500 Then Exit For

Next gointo

horiz:
Form2.Height = 30
gotoval = Form2.Width / 2

For gointo = 1 To gotoval  '�������� âũ�⸦ ���δ�
 
   DoEvents
   Form2.Width = Form2.Width - 100
   Form2.Left = (Screen.Width - Form2.Width) \ 2
   If Form2.Width <= 2000 Then Exit For

Next gointo

End Sub

Private Sub sdis_com_Click() '��� Ŭ���̾�Ʈ ������ ������ų���

For i = 1 To 10
  
  If connectindex(i) = 1 Then  '���ӵǾ��ִ� Ŭ���̾�Ʈ�鿡�Ը� �������ӽ�ȣ�� ������
    
    Winsock1(i).SendData "#!#!"
    DoEvents
    connectindex(i) = 0                '���Ϲ迭�� 0���� �ʱ�ȭ��Ų��
    Winsock1(i).Close
    Unload Winsock1(i)
    addtext "#### " + CStr(i) + "��° Ŭ���̾�Ʈ�� �������� �Ǿ����ϴ�####"
  
  
  End If

Next i

End Sub

Private Sub sexit_com_Click() '���α׷� �����ư�� �������

Call Startrek(Me)             '���ῡ�ϸ��̼� ����
End

End Sub


Private Sub slogoff_com_Click() '���ݷα׿�����ư�� ������

For i = 1 To 10                 '���� ����Ǿ��ִ� ������ �˻��Ͽ� �α׿����޽��� ����
    
    addtext CStr(connectindex(i))
    
    If connectindex(i) <> 0 Then
          
          Winsock1(i).SendData "logoff"
          DoEvents
    
    End If

Next i

End Sub

Private Sub sreboot_com_Click() '���ݷα׿�����ư�� ������

For i = 1 To 10                 '���� ����Ǿ��ִ� ������ �˻��Ͽ� ����ø޽��� ����
    
    If connectindex(i) <> 0 Then
          
          Winsock1(i).SendData "reboot"
          DoEvents
    
    End If

Next i

End Sub

Private Sub Winsock1_Close(Index As Integer) '������ �ݴ´�

connectindex(Index) = 0
Winsock1(Index).Close

End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
                                            'Ŭ���̾�Ʈ���� ���ӿ䱸�� ���ð��
If countp <> 10 Then '9���̻��� Ŭ���̾�Ʈ�� ���� �ʴ´�
  j = Index
  countp = countp + 1

  For i = 1 To 10    '���� ����ִ� ������ ã�´�
     If connectindex(i) = 0 Then
          Index = i
          Exit For
      End If
  Next i
                      '����ִ� ���Ͽ� Ŭ���̾�Ʈ�� �����Ų��
  connectindex(Index) = 1
  Load Winsock1(Index)
  Winsock1(Index).Accept requestID
  j = Index
  addtext "#" + CStr(Index) + "��° Ŭ���̾�Ʈ�� �����߽��ϴ�#"

End If
 
 Winsock1(Index).SendData "!!!" + CStr(Index)
 
End Sub


Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
                                          'Ŭ���̾�Ʈ���� �����͸� �����ð��
Dim p As String
Winsock1(Index).GetData p                 'Ŭ���̾�Ʈ���� �������� �����͸� �޴´�

Select Case CStr(Left(p, 3))              '���µ������� ��3�ڸ��� *x* �̸� ����޼����� �ν�

Case "*X*"                                '������ �ݴ´�
   
   connectindex(Index) = 0                '���Ϲ迭�� 0���� �ʱ�ȭ��Ų��
   Winsock1(Index).Close
   Unload Winsock1(Index)
   addtext "#### " + CStr(Index) + "��° Ŭ���̾�Ʈ�� �������� �Ǿ����ϴ�####"

Case Else                                 '���̿��� �����ʹ� ä�õ����ͷ� �����Ѵ�
   
   For i = 1 To 10
                                          '���� �����͸� �ٸ� Ŭ���̾�Ʈ���� �����ش�
     If connectindex(i) = 1 And i <> Index Then
      
        Winsock1(i).SendData p
        DoEvents
        
     End If
   
   Next i
   
   addtext p                               '������ ä��â�� �ѷ��ش�

End Select

End Sub

Private Sub schat_bar_keyPress(keyascii As Integer)
                                                          '�������� ä�ø޼����� ������

If keyascii = 13 And sdis_com.Enabled = True Then         '���ӻ��¿��� ���͸� ĥ���

     If tempid <> "" And tempid <> snick_bar.Text Then    '��ȭ���� ����Ǿ����� Ȯ��
         
         For i = 1 To 10                                  '���ӵǾ��ִ� Ŭ���̾�Ʈ���� ��ȭ�� ������ �����Ѵ�
            
            If connectindex(i) = 1 Then
                     
                     Winsock1(i).SendData "##########" + tempid + "�� " + snick_bar.Text + "�� ��ȭ���� ����Ǿ����ϴ�" + "#########" + vbNewLine
                     DoEvents
                     
            End If
         
         Next i                                            '���������� ���泻���� �����ش�
                     addtext "##########" + tempid + "�� " + snick_bar.Text + "�� ��ȭ���� ����Ǿ����ϴ�" + "#########"
      End If
        
        For i = 1 To 10                                    '���ӵǾ��ִ� Ŭ���̾�Ʈ���� ��ȭ������ �����Ѵ�
        If connectindex(i) = 1 Then
                    tempid = snick_bar.Text
                    Winsock1(i).SendData snick_bar.Text + ">>" + schat_bar.Text
                    DoEvents
                    
           End If
        Next i
                                                           '���������� ��ȭ������ �����ش�
      addtext snick_bar.Text + ">>" + schat_bar.Text
      schat_bar.Text = ""
End If

End Sub

Private Sub addtext(addline As String)      '���� ������ ȭ�鿡 ��½�Ű�� �Լ�


schat_win.Text = schat_win.Text + addline + vbNewLine

End Sub



