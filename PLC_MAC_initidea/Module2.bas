Attribute VB_Name = "Module2"
'######### 항목별 테스트 방법 ###########

Public Sub TestSpec(TC As String) ' input : Test Case
TestNode(NowTreeIndex).ID = TC

Select Case TC
    Case "GTC":
    
    
     TestNode(NowTreeIndex).Result = "PASS" '메모리기억
     Form1.Add_State ("PASS")
     Form1.Add_Result ("PASS")
     '+ vbCrLf + "-----------------" + vbCrLf + vbCrLf) '화면출력


    Case Else
     Debug.Print "항목이 정의되지 않았습니다."
     Form1.Add_State (TC & " 테스트 내용이 정의되지 않았습니다.")
     Form1.Add_Result ("No Test")
     


End Select
End Sub



