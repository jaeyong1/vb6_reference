Attribute VB_Name = "Module1"
'### �޷� ���� ###
Public Cal_YYYY As String
Public Cal_MM As String
Public Cal_DD As String
Public YYYYMMDD As String

Public Type PLCTestNode
    'memoery dynamic allo.. ( item : id, result(P/F), comement, errorcode )
    ID As String
    Result As String
    Comment As String
    ErrorCode As String
End Type

Public TestNode() As PLCTestNode

Public NowTreeIndex As Integer '���� �������� �׽�Ʈ�� Ʈ���ε���
 
