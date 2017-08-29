Attribute VB_Name = "RegKeys"

'############ �Ʒ� ������Ʈ�� �׼��� ����� VB�� ��ġ�Ͻø�...
'C:\Program Files\Microsoft Visual Studio\VB98\Template\Code ������ �⺻������ ����ִ� ����Դϴ�...
'###################################################################################################


' �� ����� ������Ʈ�� Ű�� �а� ���ϴ�. VB�� ���� ������Ʈ��
' �׼��� ����� �޸� ���ڿ� ������ ������Ʈ�� Ű��
' �а� �� �� �ֽ��ϴ�.

Option Explicit
'---------------------------------------------------------------
'- ������Ʈ�� API ����...
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

'---------------------------------------------------------------
'- ������Ʈ�� API ���...
'---------------------------------------------------------------
' ������Ʈ�� ������ ����...
Const REG_SZ = 1                         ' Unicode null ���� ���ڿ�
Const REG_EXPAND_SZ = 2                  ' Unicode null ���� ���ڿ�
Const REG_BINARY = 3                     ' BINARY
Const REG_DWORD = 4                      ' 32��Ʈ ����

' ������Ʈ���� ���� ���� �ۼ��մϴ�...
Const REG_OPTION_NON_VOLATILE = 0       ' �ý����� ����õǾ Ű�� �����˴ϴ�.

' ������Ʈ�� Ű ���� �ɼ�...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ������Ʈ�� Ű ROOT ����...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

' ��ȯ��...
Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0

'---------------------------------------------------------------
'- ������Ʈ�� ���� Ư�� ����...
'---------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

' ���ҽ� ���ڿ��� ������ ���� ��Ʈ���� �Ӽ��� �ε�˴ϴ�.
' Object      Property
' Form        Caption
' Menu        Caption
' TabStrip    Caption, ToolTipText
' Toolbar     ToolTipText
' ListView    ColumnHeader.Text

'-------------------------------------------------------------------------------------------------
'���� ��� - Debug.Print UpodateKey(HKEY_CLASSES_ROOT, "keyname", "newvalue")
'-------------------------------------------------------------------------------------------------
Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
    Dim rc As Long                                      ' �ڵ� ��ȯ
    Dim hKey As Long                                    ' ������Ʈ�� Ű ó��
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' ������Ʈ�� ���� ����
    
    lpAttr.nLength = 50                                 ' ���� Ư���� �⺻���� ����...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- ������Ʈ�� Ű �����/����...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' �����/���� //KeyRoot//KeyName
    
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' ���� ó��...
    
    '------------------------------------------------------------
    '- Ű �� �����/����...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' RegSetValueEx()�� ����ϱ� ���� �� ĭ�� �ʿ��մϴ�...
    
    ' Create/Modify Key Value
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))
                       
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' ���� ó��
    '------------------------------------------------------------
    '- ������Ʈ�� Ű �ݱ�...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' Ű�� ����
    
    UpdateKey = True                                    ' ������ ��ȯ
    Exit Function                                       ' ����
CreateKeyError:
    UpdateKey = False                                   ' ���� ��ȯ �ڵ带 ����
    rc = RegCloseKey(hKey)                              ' Ű �ݱ⸦ �õ�
End Function

'-------------------------------------------------------------------------------------------------
'���� ���� - Debug.Print GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
'-------------------------------------------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
    Dim i As Long                                           ' ���� ī����
    Dim rc As Long                                          ' �ڵ� ��ȯ
    Dim hKey As Long                                        ' ���� ������Ʈ�� Ű�� �ڵ�
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' ������Ʈ�� Ű�� ������ ����
    Dim tmpVal As String                                    ' ������Ʈ�� Ű ���� �ӽ� ����
    Dim KeyValSize As Long                                  ' ������Ʈ�� Ű ������ ũ��
    
    ' KeyRoot {HKEY_LOCAL_MACHINE...} �Ʒ��� RegKey ����
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ������Ʈ�� Ű ����
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ���� ó��...
    
    tmpVal = String$(1024, 0)                             ' ���� ���� �Ҵ�
    KeyValSize = 1024                                       ' ���� ũ�� ǥ��
    
    '------------------------------------------------------------
    ' ������Ʈ�� Ű �� �˻�...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' Ű �� �˾Ƴ���/�����
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ���� ó��
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' ��ȯ�� ���� Ű �� ���� ����...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' ������ ���� �˻�...
    Case REG_SZ, REG_EXPAND_SZ                              ' ���ڿ� ������Ʈ�� Ű ������ ����
        sKeyVal = tmpVal                                     ' ���ڿ� �� ����
    Case REG_DWORD                                          ' Double Word ������Ʈ�� Ű ������ ����
        For i = Len(tmpVal) To 1 Step -1                    ' ��Ʈ�� ��ȯ
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Char ������ �� Char�� ����
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' Double Word�� String�� ��ȯ
    Case REG_BINARY
        
    End Select
    
    GetKeyValue = sKeyVal                                   ' �� ��ȯ
    rc = RegCloseKey(hKey)                                  ' ������Ʈ�� Ű �ݱ�
    Exit Function                                           ' ����
    
GetKeyError:    ' Cleanup After An Error Has Occured...
    GetKeyValue = vbNullString                              ' ����ִ� ���ڿ��� ��ȯ ���� ����
    rc = RegCloseKey(hKey)                                  ' ������Ʈ�� Ű�� ����
End Function
