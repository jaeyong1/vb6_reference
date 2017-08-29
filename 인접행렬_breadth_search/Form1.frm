VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Print App.Path: Print

'Dim o As New Queue
'Dim p As New Class1
'Dim l As New Class1
'
'p.a = 1
'p.b = "aa"
'
'o.Enqueue p
'
'o.Dequeue


Dim w As New Queue
Dim q As New Queue
'
q.Enqueue "test1"
q.Enqueue "test2"
q.Enqueue "test3"
q.Enqueue "test4"


'Do While q.Count > 0
'  Print q.Dequeue

'Loop
'


Dim s As New adjacencymatrix
s.addvertex "a"
s.addvertex "b"
s.addvertex "c"
s.addvertex "d"
s.addvertex "e"
s.addvertex "f"
s.addvertex "g"
s.addvertex "h"
s.addvertex "z"
s.addvertex "i"
's.MarkVertex ("c")
's.ClearMarks
'Print IIf(s.IsMarked("c"), "t", "f")

s.addEdge "a", "b"
s.addEdge "b", "c"
s.addEdge "b", "d"
s.addEdge "d", "e"
s.addEdge "d", "f"
s.addEdge "d", "g"
s.addEdge "b", "z"
s.addEdge "z", "h"
s.addEdge "h", "i"


'
's.GetToVertice "b", w
'Do While w.Count > 0
'  Print w.Dequeue
'Loop

'Print s.WeightIs("b", "a")

BreadthFirstSearch s, "a", "f" ''''''''''''''''"

'shortestpath s, "d"


End Sub


Public Function IncreaseDistance(DisData As String)
 Dim depth As Integer
 depth = Left(DisData, 1)
 depth = depth + 1

End Function

Public Function getNext(instr As String)
'

End Function

'너비우선(Breadth-first) 탐색 방법
Public Function BreadthFirstSearch(graph As adjacencymatrix, startVertex As String, endVertex As String)

Dim q As New Queue 'queue
Dim vertexQ As New Queue
Dim distance As New Queue 'start - end 출발점에서의 거리 , 이동해온경로를 저장..
Dim Nowdistance  As String  ' ' As Integer
'Nowdistance = 0
Nowdistance = "2|0|"

Dim found As Boolean: found = False
Dim vertex As String
Dim item As String

graph.ClearMarks
q.Enqueue startVertex
distance.Enqueue Nowdistance

Do
  vertex = q.Dequeue
  
  'Print "distance (" & distance.Peek & ") ";
  Print "distance (" & getnextitem(distance.Peek, 2) & ") ";
  'Print "(" & distance.Peek & ")";
  Nowdistance = distance.Peek
  distance.Dequeue '''''

  If vertex = endVertex Then
    Print "Find " & vertex & " distance : "; Nowdistance '& getnextitem(Nowdistance, 2)  '& Nowdistance
    found = True
  Else
   
   
    If Not (graph.IsMarked(vertex)) Then
       graph.MarkVertex (vertex)
       Print "search : " & vertex
       
       graph.GetToVertice vertex, vertexQ
       
       Do While (Not (vertexQ.count = 0))
           item = vertexQ.Dequeue
           If (Not (graph.IsMarked(item))) Then
             q.Enqueue (item)
             
             'distance.Enqueue Nowdistance + 1
             
             Call IncreaseNumber(Nowdistance)
             Call addItem(Nowdistance, vertex)
             distance.Enqueue Nowdistance
             Call decreaseNumber(Nowdistance)
             
             
           End If
        
       Loop
     End If
  End If
Loop While (Not (q.count = 0) And (found = False))

If found = False Then
  Print "Path not found."
End If


End Function

Public Function qview(q As Queue)
'큐 내용을 출력함 (for test...)
Dim w As New Queue
Dim datas As String


Do While q.count > 0
  w.Enqueue q.Dequeue
Loop

Do While w.count > 0
  datas = w.Dequeue
  Print datas & " ";
  q.Enqueue datas
Loop
Print

End Function

Private Sub Command1_Click()
Unload Me
End Sub


'Public Function shortestpath(graph As adjacencymatrix, fromV As String, toV As String)
Public Function shortestpath(graph As adjacencymatrix, startVertex As String)
Print "---- shortest path ---------"

'Dim startVertex As ItemType
'startVertex.fromVertex = fromV
'startVertex.toVertex = toV

Dim item As ItemType
Dim mindistance As Integer
Dim pq As New Queue 'queue
Dim vertexQ As New Queue
Dim vertex As String

graph.ClearMarks
item.fromVertex = startVertex
item.toVertex = startVertex
item.distance = 0
'pq.Enqueue item



End Function

