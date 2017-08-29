
cls
@echo off
@echo :--------------------------------------------------:
@echo :                                                  :
@echo :          ** 안녕하세요.. 작동중 **               :
@echo :                                                  :
@echo :--------------------------------------------------:
@echo :                                                  :
@echo :   준비중..                                       :
a:
cd\
md data
del v3*.*
cls
@echo off
@echo :--------------------------------------------------:
@echo :                                                  :
@echo :          ** 안녕하세요.. 작동중 **               :
@echo :                                                  :
@echo :--------------------------------------------------:
@echo :                                                  :
@echo :   준비중..                                       :
@echo :                                                  :
@echo :   디스크 정보저장                                :
ver >a:\data\ver.txt
chkdsk c: >a:\data\chkdsk.txt
@echo :                                                  :
@echo :   메모리 정보저장                                :
mem/c >a:\data\mem.txt
@echo :                                                  :
@echo :  전체 TREE구조 정보저장(Mdir)                    :
c:
cd\
del mdirtree.dir
a:
m /cp c:sssss
c:
cd\
copy mdirtree.dir a:\data
copy mdirtree.dir a:
cls
@echo off
@echo :--------------------------------------------------:
@echo :                                                  :
@echo :          ** 안녕하세요.. 작동중 **               :
@echo :                                                  :
@echo :--------------------------------------------------:
@echo :                                                  :
@echo :   준비중..                                       :
@echo :                                                  :
@echo :   디스크 정보저장                                :
@echo :                                                  :
@echo :   메모리 정보저장                                :
@echo :                                                  :
@echo :  전체 TREE구조 정보저장(Mdir)                    :
@echo :                                                  :
@echo :  시작-프로그램 방안의 목록 압축중..              :

arj a -r c:\start c:\windows\시작메~1\프로그램\
c:
cd\
cls
@echo off
@echo :--------------------------------------------------:
@echo :                                                  :
@echo :          ** 안녕하세요.. 작동중 **               :
@echo :                                                  :
@echo :--------------------------------------------------:
@echo :                                                  :
@echo :   준비중..                                       :
@echo :                                                  :
@echo :   디스크 정보저장                                :
@echo :                                                  :
@echo :   메모리 정보저장                                :
@echo :                                                  :
@echo :  전체 TREE구조 정보저장(Mdir)                    :
@echo :                                                  :
@echo :  시작-프로그램 방안의 목록 압축중..              :
@echo :                                                  :
@echo :  기타 데이터 수집중..                            :
@echo :                                                  :
copy start.arj a:\data
del start.arj
copy autoexec.bat a:\data
copy config.sys a:\data
copy c:\windows\dosstart.bat a:\data\
@echo :                                                  :
@echo :==================================================:
@echo :                                                  :
@echo :             프로그램 실행 종료                   :
@echo :                                                  :
@echo :--------------------------------------------------:


exit


