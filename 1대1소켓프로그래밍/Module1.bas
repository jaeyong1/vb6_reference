Attribute VB_Name = "Module1"
Option Explicit
Public ���ϻ��� As Variant
Public I As Integer
Public sData As Variant

Public Sub ���ϻ��¾˸�(Index As Integer)
If Index = 0 Then   '0�϶� ����
Frm����.Lst����.AddItem (���ϻ���(Frm����.Winsock1.State))
Frm����.Lst����.ListIndex = Frm����.Lst����.ListCount - 1
End If

If Index = 1 Then   '1�϶� Ŭ���̾�Ʈ
FrmŬ���̾�Ʈ.Lst����.AddItem (���ϻ���(FrmŬ���̾�Ʈ.Winsock1.State))
FrmŬ���̾�Ʈ.Lst����.ListIndex = FrmŬ���̾�Ʈ.Lst����.ListCount - 1
End If
End Sub




