VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1215
      Left            =   3000
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Microsoft HTML Object Library,
'Microsoft Internet Controls �̷��� 2�� ���� �߰�

Private IE As InternetExplorer


'�ͽ��÷η� ����� ������ �ε�
Private Sub Command1_Click()
Set IE = New InternetExplorer
IE.Visible = True


'IE.navigate "http://www.naver.com"
IE.navigate "http://www.kbs.co.kr/asx/login/SSOLogon.php?from_url=http://www.kbs.co.kr/"


Do While IE.Busy
     DoEvents
Loop


'<input type="text"
'name�׸��� �ִ´�.
'<form



End Sub

' �α׿� ����
Private Sub Command2_Click()
Dim IE_id As HTMLInputElement
Dim IE_pwd As HTMLInputElement
Dim IE_Frm As HTMLFormElement

    

'site: naver
'Set IE_id = IE.document.getElementsByName("id")(0)     ' html ���� input�� id ���� �±� ���´�.
'Set IE_pwd = IE.document.getElementsByName("PASSWORD")(0)   ' html ���� input�� pwd ���� �±� ���´�.
'Set IE_Frm = IE.document.getElementsByName("NidLogin")(0)      ' html ���� form�� �̸��� ���´�.


Set IE_id = IE.document.getElementsByName("ID")(0)
' html ���� input�� id ���� �±� ���´�. type="text"  ã��..
Set IE_pwd = IE.document.getElementsByName("PASSWORD")(0)
' html ���� input�� pwd ���� �±� ���´�. type="password" ã��..
Set IE_Frm = IE.document.getElementsByName("login")(0)      ' html ���� form�� �̸��� ���´�.<form �� name�׸�
 



If TypeName(IE_id) <> "Nothing" And TypeName(IE_pwd) <> "Nothing" And TypeName(IE_Frm) <> "Nothing" Then

        IE_id.setAttribute "value", "jaeyong1"
        IE_pwd.setAttribute "value", "6090"
        IE_Frm.submit

End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    IE.Quit
End Sub
