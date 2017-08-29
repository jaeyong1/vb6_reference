Attribute VB_Name = "Module1"
Option Explicit

Declare Function GetTickCount Lib "kernel32" () As Long


Public t(50) As String
' == �迭 ������ ���� == (hwp ����)
'no  �溸����  ��巹��
'01  A ��� �溸  MX0100
'02  A ���� �溸  MX0101
'03  B ��� �溸  MX0102
'04  B ���� �溸  MX0103
'05  C ��� �溸  MX0104
'06  C ���� �溸  MX0105
'07  A ���·� ON  MX0106
'08  B ���·� ON  MX0107
'09  C ���·� ON  MX0108
'10  �� ��ġ��Ż  MX0109
'11  �� ��ġ��Ż  MX010A
'12  A �µ� (����)  MW111
'13  B �µ� (����)  MW112
'14  C �µ� (����)  MW113
'15  A ��� �溸  MX0200
'16  A ���� �溸  MX0201
'17  B ��� �溸  MX0202
'18  B ���� �溸  MX0203
'19  C ��� �溸  MX0204
'20  C ���� �溸  MX0205
'21  A ���·� ON  MX0206
'22  B ���·� ON  MX0207
'23  C ���·� ON  MX0208
'24  �� ��ġ��Ż  MX0209
'25  �� ��ġ��Ż  MX020A
'26  A �µ� (����)  MW121
'27  B �µ� (����)  MW122
'28  C �µ� (����)  MW123
'29  A ��� �溸  MX0300
'30  A ���� �溸  MX0301
'31  B ��� �溸  MX0302
'32  B ���� �溸  MX0303
'33  C ��� �溸  MX0304
'34  C ���� �溸  MX0305
'35  A ���·� ON  MX0306
'36  B ���·� ON  MX0307
'37  C ���·� ON  MX0308
'38  �� ��ġ��Ż  MX0309
'39  �� ��ġ��Ż  MX030A
'40  A �µ� (����)  MW131
'41  B �µ� (����)  MW132
'42  C �µ� (����)  MW133
'
'

Public PlcSendData(5) As String     '���ڿ� ����
Public PlcSendData_iter As Integer  '������ ���ڿ��迭 ���


' "A" �� 8 4 2 1 �� �ɰ��� "1010" �� ����
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

' COM Port�� ������ ������ Open �մϴ�.
Public Sub OpenCommPort()

    Dim strBps(7) As String
    Dim strParity(2) As String
    Dim strDataBit(1) As String
    Dim strStopBit(1) As String

    Dim strCom As String
    
    ' COM Port�� �����մϴ�.
    frmMain.MSComm1.CommPort = frmMain.cmbPort.ListIndex + 1
   
    ' COM Port�� �⺻ ������ �մϴ�.
    frmMain.MSComm1.Settings = "38400,N,8,1"
    
    ' COM Port�� Open �մϴ�.
    frmMain.MSComm1.PortOpen = True
    
End Sub

' ������ �����͸� Send �մϴ�.
Public Sub SendData()

    Dim Bcc As Byte
    Dim iBcc As Integer
    Dim sBcc As String
    
    ' BCC�� ������ ���
    If frmMain.chkBcc.Value = 1 Then
    
        ' BCC�� ����մϴ�.
        iBcc = 5 + ByteCheckSum(frmMain.txtTx.Text) + 4
        If iBcc > 255 Then iBcc = iBcc - 256
        Bcc = CByte(iBcc)
        sBcc = ByteToHexStr(Bcc)
        
        ' ���, ���ϰ� BCC�� ������ �������� Send �մϴ�.
        frmMain.MSComm1.Output = chr$(5) + frmMain.txtTx.Text + chr$(4) + sBcc
        
        ' ���� BCC���� ȭ�鿡 ����մϴ�.
        frmMain.txtTxBcc.Text = sBcc
        
    Else
        ' ���, ������ ������ �������� Send �մϴ�.
        frmMain.MSComm1.Output = chr$(5) + frmMain.txtTx.Text + chr$(4)
    End If

End Sub

' ������ ��Ʈ���� BCC�� ����մϴ�.
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

' BYTE DATA�� 16���� ǥ���� ��Ʈ������ ��ȯ�մϴ�.
Public Function ByteToHexStr(byData As Byte) As String
    Dim strHex As String
    
    strHex = Hex(byData)
    
    If Len(strHex) < 2 Then _
        strHex = "0" + strHex
    
    ByteToHexStr = strHex

End Function

