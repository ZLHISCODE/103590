@echo off
echo ���ڰ�װ TinyPDF...
copy /y /v tinypdf.chm %windir%\system32\spool\drivers\w32x86\
copy /y /v tinypdf.dll %windir%\system32\spool\drivers\w32x86\
copy /y /v tinypdf1.dll %windir%\system32\spool\drivers\w32x86\
copy /y /v tinypdf2.dll %windir%\system32\spool\drivers\w32x86\
copy /y /v tinypdf3.dll %windir%\system32\spool\drivers\w32x86\
copy /y /v tinypdf3.dll %windir%\system32\spool\drivers\w32x86\3\
InstallPrinter.exe
echo ...
echo TinyPDF ��װ�ɹ���  :)
@pause>nul