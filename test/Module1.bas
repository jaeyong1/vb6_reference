Attribute VB_Name = "Module1"

'''''''''''''''''''''''
'
'  갯수 | 데이터1 | 데이터2 | 데이터3.... |
'  과 같은 형태의 자료구조를 처리하기 위한 함수
'
'''''''''''''''''''''''''


Public Function getnextitem(inputdata As String, count As Integer) 'count = 갯수부분포함, 데이터 몇번째

Dim outputdata, i, j, length
i = 1

'갯수를 읽어옴
While (Mid(inputdata, i, 1) <> "|")
     length = length & Mid(inputdata, i, 1)
     i = i + 1
Wend
Debug.Print "length : " & length

 
'예외.오류처리
If length < count Then
   MsgBox "요청길이오류"
   Exit Function
End If

If count = 1 Then
  getnextitem = length
  Exit Function
End If


'읽음
i = i + 1
For j = 2 To count
  If i > Len(inputdata) Then Exit Function
  
  While (Mid(inputdata, i, 1) <> "|")
     outputdata = outputdata & Mid(inputdata, i, 1)
     i = i + 1
  Wend
 
  If j <> count Then outputdata = ""
  i = i + 1
  
Next j

'반환
 getnextitem = outputdata

End Function
'데이터 일부를 교환
Public Function changeitem(indata As String, count As Integer, newdata As String)
Dim i As Integer

Dim newoutput
Dim length
length = getnextitem(indata, 1)

For i = 1 To length
   If count = i Then
    newoutput = newoutput & newdata & "|"
   Else
    newoutput = newoutput & getnextitem(indata, i) & "|"
   End If
Next i

'Print newoutput
changeitem = newoutput
indata = newoutput
End Function


'두번째데이터를 1증가 시킴
Public Function IncreaseNumber(indata As String)

Dim number
Dim outp

number = getnextitem(indata, 2)
number = number + 1
outp = changeitem(indata, 2, number & "")

IncreaseNumber = outp
indata = outp
End Function

'데이터 추가
Public Function addItem(indata As String, newdata As String) As String
Dim number
Dim outp

outp = changeitem(indata, 1, getnextitem(indata, 1) + 1)
indata = indata & newdata & "|"
End Function


