Attribute VB_Name = "Module1"

'''''''''''''''''''''''
'
'  ���� | ������1 | ������2 | ������3.... |
'  �� ���� ������ �ڷᱸ���� ó���ϱ� ���� �Լ�
'
'''''''''''''''''''''''''


Public Function getnextitem(inputdata As String, count As Integer) 'count = �����κ�����, ������ ���°

Dim outputdata, i, j, length
i = 1

'������ �о��
While (Mid(inputdata, i, 1) <> "|")
     length = length & Mid(inputdata, i, 1)
     i = i + 1
Wend
Debug.Print "length : " & length

 
'����.����ó��
If length < count Then
   MsgBox "��û���̿���"
   Exit Function
End If

If count = 1 Then
  getnextitem = length
  Exit Function
End If


'����
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

'��ȯ
 getnextitem = outputdata

End Function
'������ �Ϻθ� ��ȯ
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


'�ι�°�����͸� 1���� ��Ŵ
Public Function IncreaseNumber(indata As String)

Dim number
Dim outp

number = getnextitem(indata, 2)
number = number + 1
outp = changeitem(indata, 2, number & "")

IncreaseNumber = outp
indata = outp
End Function

'������ �߰�
Public Function addItem(indata As String, newdata As String) As String
Dim number
Dim outp

outp = changeitem(indata, 1, getnextitem(indata, 1) + 1)
indata = indata & newdata & "|"
End Function


