--本脚本支持从ZLTOOLS v10.35.90 升级到 v10.35.90
--请以管理工具所有者登录PLSQL并执行下列脚本
-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------
--127267:余智勇,2018-06-15
Update Zltools.zlRPTSubs A
Set a.功能 = 
    (Select 名称 
     From Zltools.zlReports B, Zltools.zlRPTSubs C 
     Where b.Id = c.报表id And b.Id = a.报表id And c.功能 Is Null)
Where a.功能 Is Null;


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------



------------------------------------------------------------------------------------
Commit;