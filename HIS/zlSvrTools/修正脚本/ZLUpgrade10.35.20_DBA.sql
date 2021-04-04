----10.35.10---》10.35.20
--98732:刘硕,2016-08-13,特殊SP版本支持
declare
begin
  --修正所有的ZLBakINfo表。
  for rs in (select owner from all_tables b where b.TABLE_NAME = 'ZLBAKINFO') loop
      execute immediate 'alter table '||rs.owner||'.ZLBAKINFO modify 版本号 varchar2(20)';
  end loop;
end;
/


