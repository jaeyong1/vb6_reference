VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmDialer 
   Caption         =   "Dialer"
   ClientHeight    =   5445
   ClientLeft      =   2880
   ClientTop       =   3360
   ClientWidth     =   6450
   Icon            =   "frmDialer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6450
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1260
      Left            =   2640
      TabIndex        =   1
      Top             =   2880
      Width           =   1380
   End
   Begin VB.TextBox PhoneNum2 
      Height          =   270
      Left            =   1980
      TabIndex        =   0
      Text            =   "016-535-6090"
      Top             =   510
      Width           =   3465
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2025
      Top             =   1845
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmDialer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PhoneNum As String

Private Sub Ring()
        
    bQuit = False
    '��ȭ��ȣ �ִ°� Ȯ��
    If PhoneNum = "" Then
        MsgBox "��ȭ��ȣ�� �Է��ϼ���."
        Exit Sub
    ElseIf Left(PhoneNum, 4) = "0412" Then
        PhoneNum = Mid(PhoneNum, 6)
    End If
    '��ȭ �ɱ�
    Dial PhoneNum
End Sub

Private Sub Dial(Number$)
    

    Dim dialstring$, FromModem$, dummy

    ' AT�� Hayse ȣȯ ATTENTION ��ɾ�� �𵩿� ����� ���� �� �ʿ��մϴ�.
    ' DT�� "Dial Tone"�Դϴ�. Dial ����� �޽��ʹ� �ݴ�� �������� ����մϴ�(DP = Dial Pulse).
    ' Numbers$ �� ��ȭ�� �ɰ� �ִ� ��ȭ ��ȣ�Դϴ�.
    ' �����ݷ��� ��ȭ�� �� �� ���� ��� ���� ��ȯ�� ���� �˷��ݴϴ�(�߿�).
    ' ĳ���� ������ vbCr�� �𵩿� ����� ���� ��� �ʿ��մϴ�.
    dialstring$ = "ATDT" + Number$ + ";" + vbCr

    ' ��� ��Ʈ ����.
    ' ���콺�� COM1�� �����Ǿ� �ְ� CommPort�� 3�� �����Ǿ� �ִ� ������ �����˴ϴ�.
    MSComm1.CommPort = 2
    MSComm1.Settings = "38400,N,8,1"
    
    ' ��� ��Ʈ�� ���ϴ�.
    On Error Resume Next
    
    MSComm1.PortOpen = True
    
    If Err Then
       MsgBox "COM2 Port : not available. Change the CommPort property to another port."
       Exit Sub
    End If
    
    ' �Է� ������ ������ ����ϴ�.
    MSComm1.InBufferCount = 0
    
    ' ��ȭ�� �Žʽÿ�.
    MSComm1.Output = dialstring$
    
    ' �𵩿��� ���� �������� "Ȯ��" �޽����� ��ٸ��ϴ�.
    Do
       dummy = DoEvents()
       ' ���ۿ� �����Ͱ� �ִ� ��� �����͸� �н��ϴ�.
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' "Ȯ��"�� �˻��մϴ�.
          If InStr(FromModem$, "OK") Then
             ' ����ڰ� ��ȭ�⸦ �鵵�� �˸��ϴ�.
             Screen.MousePointer = vbDefault
             Response = MsgBox("'" & PhoneNum & "' ���� ��ȭ�� �ɰ� �ֽ��ϴ�." & vbCrLf & vbCrLf & "��ȭ�� �Ϸ��� ��ȭ�⸦ ��� Ȯ���� ��������." & vbCrLf & "������ �������� ��Ҹ� ��������.", vbOKCancel + vbExclamation, PhoneNum)
             If Response = vbOK Then
                    Exit Do
                ElseIf MSComm1.PortOpen = False Then
                    Exit Sub
                Else
                    bQuit = True
                    MSComm1.PortOpen = False
                End If
          End If
       End If
        
       ' ����ڰ� ��Ҹ� �����Ͽ����ϱ�?
       If bQuit Then
          bQuit = False
          Exit Do
       End If
    Loop
    
    ' �� ������ �������ϴ�.
    MSComm1.Output = "ATH" + vbCr
    
    ' ��Ʈ�� �ݽ��ϴ�.
    MSComm1.PortOpen = False
    
    End
End Sub

Private Sub Command1_Click()
    Screen.MousePointer = vbHourglass

    PhoneNum = PhoneNum2
    If Len(PhoneNum) > 0 And Len(PhoneNum) < 20 And InStr(PhoneNum, "") > 0 Then
    Ring
    Else
        MsgBox "��ȭ��ȣ ������ �ƴմϴ�.", vbExclamation, "����"
        End
   End If

End Sub

'
Private Sub Form_Load()
End Sub

