VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4320
      Top             =   240
   End
   Begin VB.Label Label9 
      Alignment       =   2  '��� ����
      BorderStyle     =   1  '���� ����
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "�̸�"
      Height          =   180
      Left            =   3000
      TabIndex        =   8
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "�й�"
      Height          =   180
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label Label6 
      Alignment       =   2  '��� ����
      BorderStyle     =   1  '���� ����
      Caption         =   "���� 00�� 00��"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "�ð�"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  '��� ����
      BorderStyle     =   1  '���� ����
      Caption         =   "���� 00�� 00��"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��� ����
      BorderStyle     =   1  '���� ����
      Caption         =   "0000�� 00�� 00��"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "�ð�"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "��¥"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Menu admin 
      Caption         =   "����"
      Index           =   0
      Begin VB.Menu member_edit 
         Caption         =   "�ٷ��л�����"
      End
      Begin VB.Menu worktime_edit 
         Caption         =   "�����ð�����"
      End
   End
   Begin VB.Menu dbprint 
      Caption         =   "���"
      Index           =   1
      Begin VB.Menu personal_work 
         Caption         =   "���κ� �ٹ�����"
      End
      Begin VB.Menu daily_work 
         Caption         =   "��¥�� �ٹ�����"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
