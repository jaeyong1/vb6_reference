VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private colData As Collection

Private Sub Class_Initialize()
    Set colData = New Collection
End Sub

Public Sub Clear()
    Set colData = New Collection
End Sub

Public Function Count() As Long
    Count = colData.Count
End Function

Public Sub Push(value As Variant)
    colData.Add value
End Sub

Public Function Pop() As Variant
    Pop = colData.Item(colData.Count)
    colData.Remove colData.Count
End Function

Public Function Peek() As Variant
    Peek = colData.Item(colData.Count)
End Function

Private Sub Class_Terminate()
    Set colData = Nothing
End Sub

