VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FRM�б� 
   Caption         =   "���� �б�"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm MSComm1 
      Left            =   840
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   4200
      TabIndex        =   3
      Top             =   15
      Width           =   3735
      Begin VB.CommandButton Command6 
         Caption         =   "�׼��� UDINT�� Array ���� �б�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "DINT_ARRAY 5�� �б�"
         Top             =   3240
         Width           =   3495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "�׼��� INT�� Array ���� �б�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "INT_ARRAY 10�� �б�"
         Top             =   2640
         Width           =   3495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "�׼��� INT�� ���� �б�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "INT_CV,INT_CV1 �б�"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "�׼��� WORD�� ���� �б�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "�׼��� ���� OUT_1,2. MOTOR1,2 �б�"
         Top             =   1440
         Width           =   3495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "���� ���� ���� �б�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "%QW0.3.0���� 10WORD�� ����"
         Top             =   840
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "���� ���� ���� �б�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "%QW0.3.0 1WORD�� ����"
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.TextBox TextRcvData 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   1
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton CmdScreenClear 
      Caption         =   "ȭ�����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   0
      ToolTipText     =   "����â�� ������ ����"
      Top             =   3960
      Width           =   1815
   End
End
Attribute VB_Name = "FRM�б�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InString As String
Dim Q As String
Dim Rcv As String

Private Sub CmdQuit_Click()
    Unload Me                            '���α׷��� ���� ����.
End Sub

Private Sub CmdScreenClear_Click()
    TextRcvData = ""                'TextBox�� ���ڿ� ����.
End Sub

Private Sub Command1_Click()
'���� ������ �޸� ���巹���� �����Ͽ� �����͸� �д� ���(��ɾ� �� RSS)
    '��� ������ �����
    Q = Chr(5) & "00RSS" & "01" & "08%QW0.3.0" & Chr(4)
'ERROR ������ ��ġ
On Error GoTo ErrMsg                '�����Ʈ ������ ��� ErrMsg:�� ����
    MSComm1.CommPort = 1            'COM1�� ���.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '�Է¹��� ũ�⸦ 1Byte�� ��.
    MSComm1.PortOpen = True         '�����Ʈ ����.
    MSComm1.Output = Q              '������ ����.
    
    Do
        DoEvents                    'Loop������ �ܺ�Event ����.
        InString = MSComm1.Input    '�Է¹��۷� ���� �ѹ��ڸ� �о.
        Rcv = Rcv & InString
        If InString = "" Then       '�Է¹��ۿ� ���ŵ� ���� Ȯ���Ͽ� ������ ������
            TextRcvData = "PLC�� ����� ��� ��ٷ��ּ���."
            Rcv_No = Rcv_No + 1     '���ŵ��� �ʴ� Ƚ���� COUNT.
        Else
            TextRcvData = ""
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '���������� DATA�� ���ŵ��� ���� Ƚ���� 1000���� ũ��.
            TextRcvData = ""
            Dummy = MsgBox("Time Out Error", 0, "����") 'Time Out Error �޼��� ǥ��.
            MSComm1.PortOpen = False 'COM1 ��� PORT �ݱ�.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX�� ���ŵ� ������ Do ... Loop���� �ݺ�.
    
    TextRcvData = TextRcvData & Rcv '���ŵ� ���ڿ��� Text Box�� ���.
    Rcv = ""                        '���ڿ� ���� �ʱ�ȭ.
    MSComm1.PortOpen = False        'COM1 ��� PORT �ݱ�.
ErrMsg:
    TextRcvData = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "����")
End Sub

Private Sub Command2_Click()
'���� ������ �޸� ���巹���� �����Ͽ� �����͸� �������� �д� ���(��ɾ� �� RSB)
    '��� ������ �����
    Q = Chr(5) & "00RSB" & "08%QW0.3.0" & "0A" & Chr(4)
'ERROR ������ ��ġ
On Error GoTo ErrMsg                '�����Ʈ ������ ��� ErrMsg:�� ����
    MSComm1.CommPort = 1            'COM1�� ���.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '�Է¹��� ũ�⸦ 1Byte�� ��.
    MSComm1.PortOpen = True         '�����Ʈ ����.
    MSComm1.Output = Q              '������ ����.
    
    Do
        DoEvents                    'Loop������ �ܺ�Event ����.
        InString = MSComm1.Input    '�Է¹��۷� ���� �ѹ��ڸ� �о.
        Rcv = Rcv & InString
        If InString = "" Then       '�Է¹��ۿ� ���ŵ� ���� Ȯ���Ͽ� ������ ������
            TextRcvData = "PLC�� ����� ��� ��ٷ��ּ���."
            Rcv_No = Rcv_No + 1     '���ŵ��� �ʴ� Ƚ���� COUNT.
        Else
            TextRcvData = ""
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '���������� DATA�� ���ŵ��� ���� Ƚ���� 1000���� ũ��.
            TextRcvData = ""
            Dummy = MsgBox("Time Out Error", 0, "����") 'Time Out Error �޼��� ǥ��.
            MSComm1.PortOpen = False 'COM1 ��� PORT �ݱ�.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX�� ���ŵ� ������ Do ... Loop���� �ݺ�.
    
    TextRcvData = TextRcvData & Rcv '���ŵ� ���ڿ��� Text Box�� ���.
    Rcv = ""                        '���ڿ� ���� �ʱ�ȭ.
    MSComm1.PortOpen = False        'COM1 ��� PORT �ݱ�.
ErrMsg:
    TextRcvData = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "����")
End Sub

Private Sub Command3_Click()
'�׼��� ������ ��ϵ� �۷ι� ������ WORD���� ���(��ɾ� �� R02)
'�׼��� ���� �̸� : OUT_1, OUT_2, MOTOR1, MOTOR2
    '��� ������ �����
    Q = Chr(5) & "00R02" & "04" & "05OUT_1" & "05OUT_2" & "06MOTOR1" & "06MOTOR2" & Chr(4)
'ERROR ������ ��ġ
On Error GoTo ErrMsg                '�����Ʈ ������ ��� ErrMsg:�� ����
    MSComm1.CommPort = 1            'COM1�� ���.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '�Է¹��� ũ�⸦ 1Byte�� ��.
    MSComm1.PortOpen = True         '�����Ʈ ����.
    MSComm1.Output = Q              '������ ����.
    
    Do
        DoEvents                    'Loop������ �ܺ�Event ����.
        InString = MSComm1.Input    '�Է¹��۷� ���� �ѹ��ڸ� �о.
        Rcv = Rcv & InString
        If InString = "" Then       '�Է¹��ۿ� ���ŵ� ���� Ȯ���Ͽ� ������ ������
            TextRcvData = "PLC�� ����� ��� ��ٷ��ּ���."
            Rcv_No = Rcv_No + 1     '���ŵ��� �ʴ� Ƚ���� COUNT.
        Else
            TextRcvData = ""
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '���������� DATA�� ���ŵ��� ���� Ƚ���� 1000���� ũ��.
            TextRcvData = ""
            Dummy = MsgBox("Time Out Error", 0, "����") 'Time Out Error �޼��� ǥ��.
            MSComm1.PortOpen = False 'COM1 ��� PORT �ݱ�.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX�� ���ŵ� ������ Do ... Loop���� �ݺ�.
    
    TextRcvData = TextRcvData & Rcv '���ŵ� ���ڿ��� Text Box�� ���.
    Rcv = ""                        '���ڿ� ���� �ʱ�ȭ.
    MSComm1.PortOpen = False        'COM1 ��� PORT �ݱ�.
ErrMsg:
    TextRcvData = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "����")
End Sub

Private Sub Command4_Click()
'�׼��� ������ ��ϵ� �۷ι� ������ INT���� ���(��ɾ� �� R06)
'�׼��� ���� �̸� : INT_CV,INT_CV1
    '��� ������ �����
    Q = Chr(5) & "00R06" & "02" & "06INT_CV" & "07INT_CV1" & Chr(4)
'ERROR ������ ��ġ
On Error GoTo ErrMsg                '�����Ʈ ������ ��� ErrMsg:�� ����
    MSComm1.CommPort = 1            'COM1�� ���.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '�Է¹��� ũ�⸦ 1Byte�� ��.
    MSComm1.PortOpen = True         '�����Ʈ ����.
    MSComm1.Output = Q              '������ ����.
    
    Do
        DoEvents                    'Loop������ �ܺ�Event ����.
        InString = MSComm1.Input    '�Է¹��۷� ���� �ѹ��ڸ� �о.
        Rcv = Rcv & InString
        If InString = "" Then       '�Է¹��ۿ� ���ŵ� ���� Ȯ���Ͽ� ������ ������
            TextRcvData = "PLC�� ����� ��� ��ٷ��ּ���."
            Rcv_No = Rcv_No + 1     '���ŵ��� �ʴ� Ƚ���� COUNT.
        Else
            TextRcvData = ""
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '���������� DATA�� ���ŵ��� ���� Ƚ���� 1000���� ũ��.
            TextRcvData = ""
            Dummy = MsgBox("Time Out Error", 0, "����") 'Time Out Error �޼��� ǥ��.
            MSComm1.PortOpen = False 'COM1 ��� PORT �ݱ�.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX�� ���ŵ� ������ Do ... Loop���� �ݺ�.
    
    TextRcvData = TextRcvData & Rcv '���ŵ� ���ڿ��� Text Box�� ���.
    Rcv = ""                        '���ڿ� ���� �ʱ�ȭ.
    MSComm1.PortOpen = False        'COM1 ��� PORT �ݱ�.
ErrMsg:
    TextRcvData = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "����")
End Sub

Private Sub Command5_Click()
'�׼��� ������ ��ϵ� �۷ι� ������ INT�� Array�� ���(��ɾ� �� R1B)
'�׼��� ���� �̸� : INT_ARRAY, ARRAY ���� ���� 10��.
    '��� ������ �����
    Q = Chr(5) & "00R1B" & "09INT_ARRAY" & "0A" & Chr(4)
'ERROR ������ ��ġ
On Error GoTo ErrMsg                '�����Ʈ ������ ��� ErrMsg:�� ����
    MSComm1.CommPort = 1            'COM1�� ���.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '�Է¹��� ũ�⸦ 1Byte�� ��.
    MSComm1.PortOpen = True         '�����Ʈ ����.
    MSComm1.Output = Q              '������ ����.
    
    Do
        DoEvents                    'Loop������ �ܺ�Event ����.
        InString = MSComm1.Input    '�Է¹��۷� ���� �ѹ��ڸ� �о.
        Rcv = Rcv & InString
        If InString = "" Then       '�Է¹��ۿ� ���ŵ� ���� Ȯ���Ͽ� ������ ������
            TextRcvData = "PLC�� ����� ��� ��ٷ��ּ���."
            Rcv_No = Rcv_No + 1     '���ŵ��� �ʴ� Ƚ���� COUNT.
        Else
            TextRcvData = ""
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '���������� DATA�� ���ŵ��� ���� Ƚ���� 1000���� ũ��.
            TextRcvData = ""
            Dummy = MsgBox("Time Out Error", 0, "����") 'Time Out Error �޼��� ǥ��.
            MSComm1.PortOpen = False 'COM1 ��� PORT �ݱ�.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX�� ���ŵ� ������ Do ... Loop���� �ݺ�.
    
    TextRcvData = TextRcvData & Rcv '���ŵ� ���ڿ��� Text Box�� ���.
    Rcv = ""                        '���ڿ� ���� �ʱ�ȭ.
    MSComm1.PortOpen = False        'COM1 ��� PORT �ݱ�.
ErrMsg:
    TextRcvData = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "����")
End Sub

Private Sub Command6_Click()
'�׼��� ������ ��ϵ� �۷ι� ������ UDINT�� Array�� ���(��ɾ� �� R20)
'�׼��� ���� �̸� : DINT_ARRAY, ARRAY ���� ���� 5��.
    '��� ������ �����
    Q = Chr(5) & "00R20" & "0ADINT_ARRAY" & "05" & Chr(4)
'ERROR ������ ��ġ
On Error GoTo ErrMsg                '�����Ʈ ������ ��� ErrMsg:�� ����
    MSComm1.CommPort = 1            'COM1�� ���.
    MSComm1.Settings = "9600,N,8,1" '9600bps,None Parity,8 Data Bit,1Stop Bit.
    MSComm1.InputLen = 1            '�Է¹��� ũ�⸦ 1Byte�� ��.
    MSComm1.PortOpen = True         '�����Ʈ ����.
    MSComm1.Output = Q              '������ ����.
    
    Do
        DoEvents                    'Loop������ �ܺ�Event ����.
        InString = MSComm1.Input    '�Է¹��۷� ���� �ѹ��ڸ� �о.
        Rcv = Rcv & InString
        If InString = "" Then       '�Է¹��ۿ� ���ŵ� ���� Ȯ���Ͽ� ������ ������
            TextRcvData = "PLC�� ����� ��� ��ٷ��ּ���."
            Rcv_No = Rcv_No + 1     '���ŵ��� �ʴ� Ƚ���� COUNT.
        Else
            TextRcvData = ""
            Rcv_No = 0
        End If
        If Rcv_No > 1000 Then       '���������� DATA�� ���ŵ��� ���� Ƚ���� 1000���� ũ��.
            TextRcvData = ""
            Dummy = MsgBox("Time Out Error", 0, "����") 'Time Out Error �޼��� ǥ��.
            MSComm1.PortOpen = False 'COM1 ��� PORT �ݱ�.
            Exit Sub
        End If
    Loop Until InString = Chr(3)    'ETX�� ���ŵ� ������ Do ... Loop���� �ݺ�.
    
    TextRcvData = TextRcvData & Rcv '���ŵ� ���ڿ��� Text Box�� ���.
    Rcv = ""                        '���ڿ� ���� �ʱ�ȭ.
    MSComm1.PortOpen = False        'COM1 ��� PORT �ݱ�.
ErrMsg:
    TextRcvData = Err.Number
    port_no = MSComm1.CommPort
    Dummy = MsgBox("COM" & port_no & " is " & Err.Description & ".", 0, "����")
End Sub
