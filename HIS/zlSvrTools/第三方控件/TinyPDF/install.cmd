@echo off
echo 正在安装 TinyPDF...
copy /y /v tinypdf.chm %windir%\system32\spool\drivers\w32x86\
copy /y /v tinypdf.dll %windir%\system32\spool\drivers\w32x86\
copy /y /v tinypdf1.dll %windir%\system32\spool\drivers\w32x86\
copy /y /v tinypdf2.dll %windir%\system32\spool\drivers\w32x86\
copy /y /v tinypdf3.dll %windir%\system32\spool\drivers\w32x86\
copy /y /v tinypdf3.dll %windir%\system32\spool\drivers\w32x86\3\
InstallPrinter.exe
echo ...
echo TinyPDF 安装成功！  :)
@pause>nul