----10.35.10---��10.35.20
--98732:��˶,2016-08-13,����SP�汾֧��
declare
begin
  --�������е�ZLBakINfo��
  for rs in (select owner from all_tables b where b.TABLE_NAME = 'ZLBAKINFO') loop
      execute immediate 'alter table '||rs.owner||'.ZLBAKINFO modify �汾�� varchar2(20)';
  end loop;
end;
/


