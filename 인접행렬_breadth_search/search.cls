VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "adjacencymatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'배열기반 인접행렬 구현

Const NULL_Edge = 0

Private numVertices As Integer
Private maxVertices As Integer
Private Vertices() As String 'vertex 정보저장(이름)
Private marks() As Boolean
Private edge() As Integer

'클래스 초기화
Private Sub Class_Initialize()
  Dim i
  numVertices = 0
  maxvertice = 50 '총 vertex수  (범위 0 ~ max-1 , max는 포함안됨)
  ReDim marks(maxvertice) As Boolean
    For i = 0 To maxvertice - 1: marks(i) = False: Next i '초기화
  ReDim edge(maxvertice, maxvertice) As Integer
  ReDim Vertices(maxvertice) As String
End Sub

Private Sub Class_Terminate()
 Debug.Print "graph end"
 
End Sub

Public Function addvertex(vertex As String)
'정점을 새로추가
  Dim index
    
  Vertices(numVertices) = vertex
  
  For index = 0 To numVertices - 1
    edge(numVertices, index) = NULL_Edge
    edge(index, numVertices) = NULL_Edge
  Next index
    
  numVertices = numVertices + 1
End Function


'정점의 이름을 받아서 그것의 위치를 반환
Public Function IndexIs(vertex As String)
  Dim index, i

  index = 0

  For i = 0 To numVertices - 1
    If Vertices(i) = vertex Then:   index = i:
  Next i

  IndexIs = index 'return value

End Function

'인접한 관계를 설정
Public Function addEdge(fromVertex As String, toVertex As String)
  Dim row, col
  
  row = IndexIs(fromVertex)
  col = IndexIs(toVertex)
  
  edge(row, col) = 1 '관계가 있는것으로 설정
  
  
  edge(col, row) = 1 '상호관계가 있는것으로 설정
  
  
  
End Function

'인접관계확인
Public Function WeightIs(fromVertex As String, toVertex As String)
  Dim row, col
  
  row = IndexIs(fromVertex)
  col = IndexIs(toVertex)
  
  WeightIs = edge(row, col)
End Function


Public Function GetToVertice(vertex As String, adjVertices As Queue)
'입력받은 정점의 인접한 정점들을 큐로 만들어 반환
  Dim fromIndex As Integer
  Dim toIndex As Integer
  
  fromIndex = IndexIs(vertex)
  
  For toIndex = 0 To numVertices - 1
    If edge(fromIndex, toIndex) <> NULL_Edge Then
       adjVertices.Enqueue (Vertices(toIndex))        '큐에 넣어서 리턴
     End If
  Next toIndex
End Function

'마크 전체 초기화
Public Function ClearMarks()
  Dim i
  For i = 0 To numVertices - 1
   marks(i) = False
  Next i
 
End Function

'정점에 마크함
Public Function MarkVertex(vertex As String)
  Dim i
  i = IndexIs(vertex)
  marks(i) = True
End Function

'정점에 마크 확인
Public Function IsMarked(vertex As String)
  Dim i
  i = IndexIs(vertex)
  IsMarked = marks(i)
End Function

