Attribute VB_Name = "Module1"
Option Explicit

Declare Function GetTickCount Lib "kernel32" () As Long


Public t(50) As String
' == 배열 변수별 내용 == (hwp 참고)
'no  경보내용  어드레스
'01  A 고온 경보  MX0100
'02  A 저온 경보  MX0101
'03  B 고온 경보  MX0102
'04  B 저온 경보  MX0103
'05  C 고온 경보  MX0104
'06  C 저온 경보  MX0105
'07  A 보온로 ON  MX0106
'08  B 보온로 ON  MX0107
'09  C 보온로 ON  MX0108
'10  좌 위치이탈  MX0109
'11  우 위치이탈  MX010A
'12  A 온도 (워드)  MW111
'13  B 온도 (워드)  MW112
'14  C 온도 (워드)  MW113
'15  A 고온 경보  MX0200
'16  A 저온 경보  MX0201
'17  B 고온 경보  MX0202
'18  B 저온 경보  MX0203
'19  C 고온 경보  MX0204
'20  C 저온 경보  MX0205
'21  A 보온로 ON  MX0206
'22  B 보온로 ON  MX0207
'23  C 보온로 ON  MX0208
'24  좌 위치이탈  MX0209
'25  우 위치이탈  MX020A
'26  A 온도 (워드)  MW121
'27  B 온도 (워드)  MW122
'28  C 온도 (워드)  MW123
'29  A 고온 경보  MX0300
'30  A 저온 경보  MX0301
'31  B 고온 경보  MX0302
'32  B 저온 경보  MX0303
'33  C 고온 경보  MX0304
'34  C 저온 경보  MX0305
'35  A 보온로 ON  MX0306
'36  B 보온로 ON  MX0307
'37  C 보온로 ON  MX0308
'38  좌 위치이탈  MX0309
'39  우 위치이탈  MX030A
'40  A 온도 (워드)  MW131
'41  B 온도 (워드)  MW132
'42  C 온도 (워드)  MW133
'
'

Public PlcSendData(5) As String     '문자열 저장
Public PlcSendData_iter As Integer  '전송한 문자열배열 기억


' "A" 를 8 4 2 1 로 쪼갠후 "1010" 로 리턴
Public Function hextoarray(inchar As String) As String

    Dim re As String

    If inchar = "0" Then
        re = "0000"
    ElseIf inchar = "1" Then
        re = "0001"
    ElseIf inchar = "2" Then
        re = "0010"
    ElseIf inchar = "3" Then
    re = "0011"
    ElseIf inchar = "4" Then
    re = "0100"
    ElseIf inchar = "5" Then
    re = "0101"
    ElseIf inchar = "6" Then
    re = "0110"
    ElseIf inchar = "7" Then
    re = "0111"
    ElseIf inchar = "8" Then
    re = "1000"
    ElseIf inchar = "9" Then
    re = "1001"
    ElseIf inchar = "A" Then
    re = "1010"
    ElseIf inchar = "B" Then
    re = "1011"
    ElseIf inchar = "C" Then
    re = "1100"
    ElseIf inchar = "D" Then
    re = "1101"
    ElseIf inchar = "E" Then
    re = "1110"
    Else
    re = "1111"
    
    End If

    hextoarray = re

End Function

' COM Port를 설정한 값으로 Open 합니다.
Public Sub OpenCommPort()

    Dim strBps(7) As String
    Dim strParity(2) As String
    Dim strDataBit(1) As String
    Dim strStopBit(1) As String

    Dim strCom As String
    
    ' COM Port를 선택합니다.
    frmMain.MSComm1.CommPort = frmMain.cmbPort.ListIndex + 1
   
    ' COM Port의 기본 설정을 합니다.
    frmMain.MSComm1.Settings = "38400,N,8,1"
    
    ' COM Port를 Open 합니다.
    frmMain.MSComm1.PortOpen = True
    
End Sub

' 설정한 데이터를 Send 합니다.
Public Sub SendData()

    Dim Bcc As Byte
    Dim iBcc As Integer
    Dim sBcc As String
    
    ' BCC를 선택한 경우
    If frmMain.chkBcc.Value = 1 Then
    
        ' BCC를 계산합니다.
        iBcc = 5 + ByteCheckSum(frmMain.txtTx.Text) + 4
        If iBcc > 255 Then iBcc = iBcc - 256
        Bcc = CByte(iBcc)
        sBcc = ByteToHexStr(Bcc)
        
        ' 헤더, 테일과 BCC를 포함한 프레임을 Send 합니다.
        frmMain.MSComm1.Output = chr$(5) + frmMain.txtTx.Text + chr$(4) + sBcc
        
        ' 계산된 BCC값을 화면에 출력합니다.
        frmMain.txtTxBcc.Text = sBcc
        
    Else
        ' 헤더, 테일을 포함한 프레임을 Send 합니다.
        frmMain.MSComm1.Output = chr$(5) + frmMain.txtTx.Text + chr$(4)
    End If

End Sub

' 지정한 스트링의 BCC를 계산합니다.
Public Function ByteCheckSum(strData As String) As Byte
    
    Dim i As Long
    Dim CheckSum As Integer
    Dim Length As Integer
    
    Length = Len(strData)
    
    CheckSum = 0
    For i = 1 To Length
        CheckSum = CheckSum + Asc(Mid(strData, i, 1))
        If CheckSum > 255 Then CheckSum = CheckSum - 256
    Next
    
    ByteCheckSum = CByte(CheckSum)
    
End Function

' BYTE DATA를 16진수 표현의 스트링으로 변환합니다.
Public Function ByteToHexStr(byData As Byte) As String
    Dim strHex As String
    
    strHex = Hex(byData)
    
    If Len(strHex) < 2 Then _
        strHex = "0" + strHex
    
    ByteToHexStr = strHex

End Function

