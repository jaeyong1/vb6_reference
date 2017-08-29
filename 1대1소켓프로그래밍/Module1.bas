Attribute VB_Name = "Module1"
Option Explicit
Public 소켓상태 As Variant
Public I As Integer
Public sData As Variant

Public Sub 소켓상태알림(Index As Integer)
If Index = 0 Then   '0일때 서버
Frm서버.Lst상태.AddItem (소켓상태(Frm서버.Winsock1.State))
Frm서버.Lst상태.ListIndex = Frm서버.Lst상태.ListCount - 1
End If

If Index = 1 Then   '1일때 클라이언트
Frm클라이언트.Lst상태.AddItem (소켓상태(Frm클라이언트.Winsock1.State))
Frm클라이언트.Lst상태.ListIndex = Frm클라이언트.Lst상태.ListCount - 1
End If
End Sub




