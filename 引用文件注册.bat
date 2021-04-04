copy .\HIS\第三方控件\OLEGUIDS.TLB c:\Windows\System32 /Y
copy .\HIS\第三方控件\olelib.tlb c:\Windows\System32 /Y
copy .\HIS\第三方控件\ISHF_Ex.tlb c:\Windows\System32 /Y
copy .\HIS\第三方控件\SHLEXT.tlb c:\Windows\System32 /Y
for %%c in (.\HIS\第三方控件\*.ocx) do regsvr32.exe /s %%c 
.\HIS\第三方控件\c1regsvr.exe .\HIS\第三方控件\olch2x8.ocx -s
%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm .\HIS\第三方控件\ZLSoft.BusinessHome.ClientControl.TimeLineBase.dll /tlb /codebase