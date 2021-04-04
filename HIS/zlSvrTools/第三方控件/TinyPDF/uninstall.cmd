@echo off
echo 正在卸载 TinyPDF...
rundll32.exe printui.dll, PrintUIEntry /dl /n""TinyPDF""
rundll32.exe printui.dll, PrintUIEntry /dd /K /m ""TinyPDF"" /h ""Windows NT x86"" /v 3
If EXIST %windir%\system32\spool\drivers\w32x86\tinypdf.chm  Del /S /Q %windir%\system32\spool\drivers\w32x86\tinypdf.chm
If EXIST %windir%\system32\spool\drivers\w32x86\tinypdf.dll  Del /S /Q %windir%\system32\spool\drivers\w32x86\tinypdf.dll
If EXIST %windir%\system32\spool\drivers\w32x86\tinypdf1.dll Del /S /Q %windir%\system32\spool\drivers\w32x86\tinypdf1.dll
If EXIST %windir%\system32\spool\drivers\w32x86\tinypdf2.dll Del /S /Q %windir%\system32\spool\drivers\w32x86\tinypdf2.dll
If EXIST %windir%\system32\spool\drivers\w32x86\tinypdf3.dll Del /S /Q %windir%\system32\spool\drivers\w32x86\tinypdf3.dll
echo ...
echo TinyPDF 卸载成功！  :)
@pause>nul
