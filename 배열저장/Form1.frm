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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�迭�� ����� ���� ���Ͽ� ���� ������ �����ϴ� ���

'�迭�� ����� ���� ���Ͽ� ���� ������ �����ϴ� ���
'
'�迭�� ����� ���� ���Ͽ� �����ϰų� ���� ���
'
'�迭�� ���� ���������� �����ϰ� �Ǹ� ���� �ӵ���
'
'������ �˴ϴ�.
'
'�̷����� ���̳��� ������ �̿� �Ͽ� �Ѳ����� �����ϰų�
'
'������ ���İ��� ó���� �˴ϴ�.
'
'�Ʒ��� 10000 ���� �迭�� �ִٰ� �����ϰ�
'
'�װ��� �����ϴ� ������ �ֽ��ϴ�.
Private Sub Form_Activate()
   Dim arr(1 To 100000) As Long
   Dim fnum As Integer

       fnum = FreeFile
       Open "C:\Temp\xxx.dat" For Binary As fnum
       Put #fnum, , arr
       Close fnum
End Sub
'��.
 


