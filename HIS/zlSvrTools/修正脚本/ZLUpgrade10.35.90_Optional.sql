----10.35.80---》10.35.90
--000000:蒋敏,2017-12-07,删除期间表ZLPERIODS
Create Or Replace Procedure Zltools.Zl_Optional_期间表_删除 Is
  n_Count Number;
Begin
  Select Count(1) Into n_count From all_tables Where owner = 'ZLTOOLS' And table_name = 'ZLPERIODS';
  If n_count > 0 Then
    Execute Immediate 'DROP TABLE ZLTOOLS.ZLPERIODS';
    Execute Immediate 'DROP PUBLIC SYNONYM ZLPERIODS';
  End If;
  Execute Immediate 'DELETE FROM ZLTABLES WHERE 表名=' || Chr(39) || 'ZLPERIODS' || Chr(39);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Optional_期间表_删除;
/

--117980:高腾,2017-12-26,删除操作模块编号_bak字段
Create Or Replace Procedure Zltools.Zl_Optional_日志表_删除字段 Is
Begin
  If ZL_CheckObject(2,'操作模块编号_bak','Zlauditlog') = 1 then
    Execute Immediate 'ALTER TABLE Zlauditlog DROP COLUMN 操作模块编号_bak';
  End if;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Optional_日志表_删除字段;
/