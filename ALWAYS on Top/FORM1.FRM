VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Label Label1 
      Caption         =   "Always on top | �׻�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Always on top ���..
   Private Declare Function SetWindowPos Lib "user32" _
         ( _
          ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
          ByVal x As Long, ByVal y As Long, _
          ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long _
         ) As Long
   Private Const HWND_TOPMOST = -1
   Private Const HWND_NOTOPMOST = -2
   Private Const SWP_NOSIZE = &H1

   Private Sub Form_Load()
       Call SetWindowPos(Form1.hwnd, HWND_TOPMOST, 0, 0, 1, 1, SWP_NOSIZE)
   End Sub

' >>> ����
'
'   ù��° �μ�
'     Always On Top�� ����� �۾������ �ڵ��Դϴ�!
'     �� Form1�� Always On Top����� ����Ѵٴ� ���Դϴ�!
'
'   �ι�° �μ�
'     Always On Top�� ��������� �����Դϴ�!
'     -1�� ���, -2�� ���Ұ��Դϴ�.''
'
'   ����°/�׹�° �μ�
'     ���� X��ǥ, Y��ǥ�Դϴ�!
'     0,0�� ������� ��ǥ���� 0,0�� �߰� �˴ϴ�!
'

