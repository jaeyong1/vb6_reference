* SMS API 설치하기 - 요약본

  -zip 파일을 다운받아서 임시 디렉토리에 압축을 풀어 줍니다.

  - SMSCOM.dll과 SMSWorld.reg 파일을 보관하실 위치에 이동시킵니다. 
    (C:\WinNT\System32 로 이동 권장)

  - SMSCOM.dll을 경로명과 함께 regsvr32 명령을 이용하여 로드합니다.
    ex) C:\Program Files\SMSAPI\ 라는 경로에 파일을 옮겨놓으셨으면,
        [시작]메뉴의 [실행]을 눌러 다음의 명령어를 실행합니다.

        regsvr32 "C:\Program Files\SMSAPI\SMSCOM.dll"

  - SMSWorld.reg 파일을 메모장으로 열어서 고객의 SMS월드가입시 사용한
    ID와 PW로 내용을 수정해 주십시오.

  - SMSWorld.reg 파일을 더블클릭 하시면 새로운 registry 키를 추가할지 여부를 묻습니다.
    이때, "예"를 눌러주세요.
   - 추후 ID나 비밀번호가 변경될 경우도 똑같이 SMSWorld.reg 파일을 수정해서
     더블클릭하면 수정된 내용이 적용되게 됩니다.

  - 그럼, SMSAPI의 설치는 완료되었습니다.

  - 압축파일에 포함되어 있는 VB예제 파일을 참고하여 사용하시면 되겠습니다.

-------------------------
Shell("Regsvr32 /s "c:\winnt\aaa.dll")

vb내에서 Shell 함수 호출하셔서 등록하면 되고요. (s :silent messageBox)
