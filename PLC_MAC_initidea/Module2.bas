Attribute VB_Name = "Module2"
'######### �׸� �׽�Ʈ ��� ###########

Public Sub TestSpec(TC As String) ' input : Test Case
TestNode(NowTreeIndex).ID = TC

Select Case TC
    Case "GTC":
    
    
     TestNode(NowTreeIndex).Result = "PASS" '�޸𸮱��
     Form1.Add_State ("PASS")
     Form1.Add_Result ("PASS")
     '+ vbCrLf + "-----------------" + vbCrLf + vbCrLf) 'ȭ�����


    Case Else
     Debug.Print "�׸��� ���ǵ��� �ʾҽ��ϴ�."
     Form1.Add_State (TC & " �׽�Ʈ ������ ���ǵ��� �ʾҽ��ϴ�.")
     Form1.Add_Result ("No Test")
     


End Select
End Sub



