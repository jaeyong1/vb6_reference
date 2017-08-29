Attribute VB_Name = "RegKeys"

'############ 아래 레지스트리 액세스 모듈은 VB를 설치하시면...
'C:\Program Files\Microsoft Visual Studio\VB98\Template\Code 폴더에 기본적으로 들어있는 모듈입니다...
'###################################################################################################


' 이 모듈은 레지스트리 키를 읽고 씁니다. VB의 내부 레지스트리
' 액세스 방법과 달리 문자열 값으로 레지스트리 키를
' 읽고 쓸 수 있습니다.

Option Explicit
'---------------------------------------------------------------
'- 레지스트리 API 선언...
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

'---------------------------------------------------------------
'- 레지스트리 API 상수...
'---------------------------------------------------------------
' 레지스트리 데이터 형식...
Const REG_SZ = 1                         ' Unicode null 종료 문자열
Const REG_EXPAND_SZ = 2                  ' Unicode null 종료 문자열
Const REG_BINARY = 3                     ' BINARY
Const REG_DWORD = 4                      ' 32비트 숫자

' 레지스트리는 형식 값을 작성합니다...
Const REG_OPTION_NON_VOLATILE = 0       ' 시스템이 재부팅되어도 키는 보존됩니다.

' 레지스트리 키 보안 옵션...
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
                     
' 레지스트리 키 ROOT 형식...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

' 반환값...
Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0

'---------------------------------------------------------------
'- 레지스트리 보안 특성 형식...
'---------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

' 리소스 문자열은 다음과 같이 컨트롤의 속성에 로드됩니다.
' Object      Property
' Form        Caption
' Menu        Caption
' TabStrip    Caption, ToolTipText
' Toolbar     ToolTipText
' ListView    ColumnHeader.Text

'-------------------------------------------------------------------------------------------------
'예제 사용 - Debug.Print UpodateKey(HKEY_CLASSES_ROOT, "keyname", "newvalue")
'-------------------------------------------------------------------------------------------------
Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
    Dim rc As Long                                      ' 코드 반환
    Dim hKey As Long                                    ' 레지스트리 키 처리
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' 레지스트리 보안 형식
    
    lpAttr.nLength = 50                                 ' 보안 특성을 기본으로 설정...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- 레지스트리 키 만들기/열기...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' 만들기/열기 //KeyRoot//KeyName
    
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' 오류 처리...
    
    '------------------------------------------------------------
    '- 키 값 만들기/열기...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' RegSetValueEx()를 사용하기 위해 빈 칸이 필요합니다...
    
    ' Create/Modify Key Value
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))
                       
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' 오류 처리
    '------------------------------------------------------------
    '- 레지스트리 키 닫기...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' 키를 닫음
    
    UpdateKey = True                                    ' 성공을 반환
    Exit Function                                       ' 끝냄
CreateKeyError:
    UpdateKey = False                                   ' 오류 반환 코드를 설정
    rc = RegCloseKey(hKey)                              ' 키 닫기를 시도
End Function

'-------------------------------------------------------------------------------------------------
'샘플 예제 - Debug.Print GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
'-------------------------------------------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
    Dim i As Long                                           ' 루프 카운터
    Dim rc As Long                                          ' 코드 반환
    Dim hKey As Long                                        ' 열린 레지스트리 키의 핸들
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' 레지스트리 키의 데이터 형식
    Dim tmpVal As String                                    ' 레지스트리 키 값의 임시 저장
    Dim KeyValSize As Long                                  ' 레지스트리 키 변수의 크기
    
    ' KeyRoot {HKEY_LOCAL_MACHINE...} 아래의 RegKey 열기
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 레지스트리 키 열기
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 오류 처리...
    
    tmpVal = String$(1024, 0)                             ' 변수 공간 할당
    KeyValSize = 1024                                       ' 변수 크기 표시
    
    '------------------------------------------------------------
    ' 레지스트리 키 값 검색...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' 키 값 알아내기/만들기
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 오류 처리
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' 변환을 위한 키 값 형식 결정...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' 데이터 형식 검색...
    Case REG_SZ, REG_EXPAND_SZ                              ' 문자열 레지스트리 키 데이터 형식
        sKeyVal = tmpVal                                     ' 문자열 값 복사
    Case REG_DWORD                                          ' Double Word 레지스트리 키 데이터 형식
        For i = Len(tmpVal) To 1 Step -1                    ' 비트를 변환
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Char 단위로 값 Char을 만듦
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' Double Word를 String로 변환
    Case REG_BINARY
        
    End Select
    
    GetKeyValue = sKeyVal                                   ' 값 반환
    rc = RegCloseKey(hKey)                                  ' 레지스트리 키 닫기
    Exit Function                                           ' 끝냄
    
GetKeyError:    ' Cleanup After An Error Has Occured...
    GetKeyValue = vbNullString                              ' 비어있는 문자열로 반환 값을 설정
    rc = RegCloseKey(hKey)                                  ' 레지스트리 키를 닫음
End Function
