
--本脚本支持从ZLHIS+ v10.35.80 升级到 v10.35.90
--请以系统所有者登录PLSQL并执行下列脚本
--脚本执行后，请手工升级导出报表
Define n_System=100;
-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------
Create Or Replace Procedure Zl_Optional_病历标记图_删除 Is
  n_Count Number;
Begin
  Select Count(1) Into n_count From all_tables Where table_name = '病历标记图';
  If n_count > 0 Then
    Execute Immediate 'DROP TABLE 病历标记图';
    Execute Immediate 'DROP PUBLIC SYNONYM 病历标记图';
  End If;
  Execute Immediate 'DELETE FROM ZLPROGPRIVS WHERE 对象=' || Chr(39) || '病历标记图' || Chr(39);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Optional_病历标记图_删除;
/



-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------





-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------






-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------




---------------------------------------------------------------------------------------------------
--更改系统及部件的版本号
-------------------------------------------------------------------------------------------------------
--系统版本号
--部件版本号
Commit;