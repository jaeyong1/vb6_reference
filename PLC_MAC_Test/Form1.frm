VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "Fm20.dll"
Begin VB.Form Form1 
   Caption         =   "������¼����(KS X 4600-1) Ŭ���� B ��ġ ����"
   ClientHeight    =   10500
   ClientLeft      =   240
   ClientTop       =   795
   ClientWidth     =   15780
   LinkTopic       =   "Form1"
   ScaleHeight     =   700
   ScaleMode       =   3  '�ȼ�
   ScaleWidth      =   1052
   Begin VB.TextBox Text21 
      Alignment       =   2  '��� ����
      BackColor       =   &H00D0D0D0&
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   12360
      TabIndex        =   60
      Text            =   "�����"
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  '��� ����
      BackColor       =   &H008080FF&
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   11280
      TabIndex        =   59
      Text            =   "���հ�"
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox tmpRE 
      Alignment       =   2  '��� ����
      BackColor       =   &H0080FF80&
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   10200
      TabIndex        =   58
      Text            =   "�հ�"
      Top             =   6600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "�׽�Ʈ ���� ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   3480
      TabIndex        =   39
      Top             =   7320
      Width           =   6015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   1560
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label19 
         Caption         =   "1.1.2.2  MAC Frame Boundary Offset (MFBO)"
         Height          =   255
         Left            =   1560
         TabIndex        =   47
         Top             =   2040
         Width           =   4215
      End
      Begin VB.Label Label17 
         Caption         =   "�׽�Ʈ �׸� : "
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "MPDU ������ [20%]"
         Height          =   255
         Left            =   2040
         TabIndex        =   45
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label Label18 
         Caption         =   "�׽�Ʈ ���� �ܰ� :"
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "46�� / 105�� [43%]"
         Height          =   255
         Left            =   2040
         TabIndex        =   43
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "�׽�Ʈ ������Ȳ  : "
         Height          =   180
         Left            =   360
         TabIndex        =   42
         Top             =   960
         Width           =   1560
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���� �׽�Ʈ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   38
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton cmdprt 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   37
      Top             =   9720
      Width           =   3015
   End
   Begin VB.CommandButton cmdtest 
      Caption         =   "Test"
      Height          =   495
      Left            =   3600
      TabIndex        =   36
      Top             =   11040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtRunState 
      Height          =   1215
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   34
      Top             =   11400
      Visible         =   0   'False
      Width           =   6255
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6135
      Left            =   3720
      TabIndex        =   33
      Top             =   600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   10821
      _Version        =   393217
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.CommandButton Command6 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   32
      Top             =   9000
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   31
      Top             =   9000
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   30
      Top             =   10440
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  '������ ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Text            =   "60"
      Top             =   8400
      Width           =   3015
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  '������ ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Text            =   "BB00000000000000"
      Top             =   7680
      Width           =   3015
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  '������ ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Text            =   "AA00000000000000"
      Top             =   6960
      Width           =   3015
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  '������ ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Text            =   "220000000000"
      Top             =   12120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  '������ ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Text            =   "110000000000"
      Top             =   11400
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  '������ ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Text            =   "20"
      Top             =   6120
      Width           =   3015
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  '������ ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Text            =   "5"
      Top             =   5400
      Width           =   3015
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  '������ ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Text            =   "6000"
      Top             =   4680
      Width           =   3015
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  '������ ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Text            =   "6000"
      Top             =   3240
      Width           =   3015
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  '��� ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Text            =   "2"
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  '��� ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Text            =   "0"
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  '��� ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Text            =   "0"
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  '��� ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Text            =   "10"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '��� ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Text            =   "4"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '��� ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '��� ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '��� ����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "10"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���� �׽�Ʈ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Ϲ� �׽�Ʈ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      Caption         =   "�׽�Ʈ ���� ���� ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   9840
      TabIndex        =   48
      Top             =   7320
      Width           =   5655
      Begin VB.TextBox TxtRunResult 
         Height          =   1215
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   51
         Top             =   1560
         Width           =   5175
      End
      Begin VB.TextBox Text18 
         Height          =   375
         Left            =   840
         TabIndex        =   50
         Text            =   "1.1.2.5	Oldest Pending Segment Flag (OPSF)"
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox Text19 
         Height          =   375
         Left            =   840
         TabIndex        =   49
         Text            =   "����"
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "�׸� : "
         Height          =   180
         Left            =   240
         TabIndex        =   54
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "���� :"
         Height          =   180
         Left            =   240
         TabIndex        =   53
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label22 
         Caption         =   "���� ��� :"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "�׽�Ʈ ���� ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   9840
      TabIndex        =   55
      Top             =   240
      Width           =   5655
      Begin MSForms.Frame Frame5 
         Height          =   615
         Left            =   240
         OleObjectBlob   =   "Form1.frx":0000
         TabIndex        =   61
         Top             =   6120
         Width           =   3375
      End
      Begin VB.TextBox txtRe 
         BackColor       =   &H0080FF80&
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Index           =   0
         Left            =   -600
         TabIndex        =   56
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "�׽�Ʈ ���̽� ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   3480
      TabIndex        =   57
      Top             =   240
      Width           =   6015
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9840
      TabIndex        =   35
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "TIME OUT"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   29
      Top             =   8160
      Width           =   1005
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "DUT Tester ENCRYPTION KEY"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   27
      Top             =   7440
      Width           =   2940
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Tester ENCRYPTION KEY"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   25
      Top             =   6600
      Width           =   2445
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "2nd Group ID"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   23
      Top             =   11880
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1st Group ID"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   21
      Top             =   11160
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "EXTENDED TEST TIME"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   19
      Top             =   5880
      Width           =   2250
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "TEST TIME"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   17
      Top             =   5160
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "DUT App. PORT NUMBER"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   2595
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "DUT IP ADDRESS"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   1785
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "DUT App. IP ADDRESS"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   2310
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "DUT IP ADDRESS"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1785
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprt_Click()
'####### ��� ��ư #######
frmprt.Show 1

End Sub

Private Sub cmdtest_Click()
'test button




Add_Result ("aaaa")


End Sub

Public Sub Add_Result(str As String)
'###### �׽�Ʈ ��� �ؽ�Ʈ �߰� ######
    TxtRunResult.Text = TxtRunResult.Text + str
    TxtRunResult.SelStart = Len(TxtRunResult.Text)
End Sub

Public Sub Add_State(str As String)
'###### �׽�Ʈ ���� ���� �ؽ�Ʈ �߰� ######
    TxtRunState.Text = TxtRunState.Text + vbCrLf + str
    TxtRunState.SelStart = Len(TxtRunState.Text)
End Sub

Private Sub Command5_Click()
' ##### ���۹�ư ######
Dim nTest As Integer
nTest = TreeView1.Nodes.Count
Debug.Print "test start / num of test : " & nTest
 
Erase TestNode  '�����ڷ� ����
ReDim TestNode(nTest) As PLCTestNode '�׸� ������ŭ ����
 
Dim s
For NowTreeIndex = 1 To nTest
    If TreeView1.Nodes(NowTreeIndex).Checked = True Then
    '## �� �׸񺰷� üũ Ȯ���� ���� ##
        Add_State ("<" & TreeView1.Nodes(NowTreeIndex).Text & ">")    'ȭ�����
        Add_Result ("* " & TreeView1.Nodes(NowTreeIndex).Text & " : ")     'ȭ�����
        s = TreeView1.Nodes(NowTreeIndex).Key
        TestSpec (s)    '�׽�Ʈ ��ü ȣ��
        
        Add_State (vbCrLf + "-----------------" + vbCrLf + vbCrLf)
        Add_Result (vbCrLf)
        
        
     '+ vbCrLf + "-----------------" + vbCrLf + vbCrLf) 'ȭ�����

    End If
Next NowTreeIndex


 
End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub Form_Load()
'�� �ҷ�����

'####### �׽�Ʈ ���̽� Ʈ�� ��� �Է� ########
Dim nod_x As Node
Set nod_x = TreeView1.Nodes.Add(, , "GTC", "General Test Cases")   '�ε������� 1
    Set nod_x = TreeView1.Nodes.Add("GTC", tvwChild, "CFF", "Conrol Frame ")
        Set nod_x = TreeView1.Nodes.Add("CFF", tvwChild, "1_1", "1.1 DT field of Control Frame")
        Set nod_x = TreeView1.Nodes.Add("CFF", tvwChild, "1_2", "1.2 VF field of Unicast Data Frame")
        Set nod_x = TreeView1.Nodes.Add("CFF", tvwChild, "1_3", "1.3 DT field of Management Frame")
                Set nod_x = TreeView1.Nodes.Add("CFF", tvwChild, "1_4", "1.4 DT field of Broadcast Data Frame")
    Set nod_x = TreeView1.Nodes.Add("GTC", tvwChild, "DFF", "Data Frame Format")
        Set nod_x = TreeView1.Nodes.Add("DFF", tvwChild, "2_1", "2.1 AAAAA")
        Set nod_x = TreeView1.Nodes.Add("DFF", tvwChild, "2_2", "2.2 BBBBB")
    Set nod_x = TreeView1.Nodes.Add("GTC", tvwChild, "IFS", "IFS(Inter-Frame Space")
    Set nod_x = TreeView1.Nodes.Add("GTC", tvwChild, "CE", "CE(Channel Estimation")


TreeView1.Nodes.Item(1).Expanded = True 'Root���� Ȯ��
'Debug.Print TreeView1.Nodes.Count '��� ����
Debug.Print vbCrLf & vbCrLf & vbCrLf

ProgressBar1.Value = 43
ProgressBar2.Value = 10

 
'########### �ڽ� ���� ���� ##########
    Dim i, j, k As Integer
    k = 1
    For j = 1 To 21
    For i = 1 To 5
        Load txtRe(k)
        txtRe(k).Visible = True
        txtRe(k).Left = (i * txtRe(0).Width) + txtRe(0).Left
        txtRe(k).Top = (j * txtRe(0).Height) + txtRe(0).Top
        
        If k > 47 Then
            txtRe(k).BackColor = &HD0D0D0
            End If
        
        k = k + 1
        
         
    Next
    Next
txtRe(4).BackColor = &H8080FF
txtRe(27).BackColor = &H8080FF
    
End Sub
 
 
'######### ��� Į�� �ڽ� ���� �Ҵ� #############
Private Sub txtRe_Click(Index As Integer)
If (txtRe(Index).BackColor = &H80FF80) Then
    txtRe(Index).BackColor = &H8080FF
    Else
    txtRe(Index).BackColor = &H80FF80
    End If
    
    



End Sub



Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
'Ʈ�� üũ

'Debug.Print "me:" & Node.Index

'########  Rootüũ�� ����/���� �� ���  ########
If Node.Index = 1 Then
    For q = 1 To TreeView1.Nodes.Count
         TreeView1.Nodes(q).Checked = Node.Checked
    Next q
    Exit Sub
End If


'########  üũǥ�� ���� �� ���  ########
If (Node.Checked = True) And (Node.Index <> 1) Then
    Debug.Print "pa:" & Node.Parent.Index
    
    If Node.Parent.Checked = False Then    '�ڽ� üũ�ϸ� �θ� üũ�ǰ�..
        Node.Parent.Checked = True
    End If
    

    Debug.Print "node.Children" & Node.Children
    For q = Node.Index To (Node.Index + Node.Children)  '�θ� üũ�ϸ� �ڽĵ� üũ�ǰ�.
        TreeView1.Nodes(q).Checked = True
    Next q
    Exit Sub
End If


'########  üũǥ�� ���� �� ���  ########
If (Node.Checked = False) And (Node.Index <> 1) Then
    Debug.Print "pa:" & Node.Parent.Index

    Debug.Print "node.Children" & Node.Children
    For q = Node.Index To (Node.Index + Node.Children)  '�θ� üũ�ϸ� �ڽĵ� �����ǰ�.
        TreeView1.Nodes(q).Checked = False
    Next q
    Exit Sub
End If


End Sub

