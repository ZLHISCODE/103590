copy .\HIS\�������ؼ�\OLEGUIDS.TLB c:\Windows\System32 /Y
copy .\HIS\�������ؼ�\olelib.tlb c:\Windows\System32 /Y
copy .\HIS\�������ؼ�\ISHF_Ex.tlb c:\Windows\System32 /Y
copy .\HIS\�������ؼ�\SHLEXT.tlb c:\Windows\System32 /Y
for %%c in (.\HIS\�������ؼ�\*.ocx) do regsvr32.exe /s %%c 
.\HIS\�������ؼ�\c1regsvr.exe .\HIS\�������ؼ�\olch2x8.ocx -s
%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm .\HIS\�������ؼ�\ZLSoft.BusinessHome.ClientControl.TimeLineBase.dll /tlb /codebase