--lshuo注释
--1.当存在小于3.0版本的PLJSON时，执行该卸载会导致低于3.0版本的无法使用
--2、卸载低版本的PLJSON时，不影响该版本的使用。
declare 
  n_have1Version number(3);
begin
  --lshuo注释
  --先判断是否存在小于3版本PLJSON,通过JSON_LIST Type,JSON与JSON_VALUE这两个名称在12C存在，JSON_LIST更好
  Select Count(1) Into n_have1Version
  From All_Objects A
  Where a.Object_Name Like 'JSON_LIST' And a.Object_Type = 'TYPE'  And a.Owner = User;
  /* drop old install */
  --lshuo测试
  --不删除3.0以前的版本
  --begin execute immediate 'drop package json_parser'; exception when others then null; end;
  --begin execute immediate 'drop package json_printer'; exception when others then null; end;
  --begin execute immediate 'drop package json_ext'; exception when others then null; end;
  --begin execute immediate 'drop package json_dyn'; exception when others then null; end;
  --begin execute immediate 'drop package json_ml'; exception when others then null; end;
  --begin execute immediate 'drop package json_xml'; exception when others then null; end;
  --begin execute immediate 'drop package json_util_pkg'; exception when others then null; end;
  --begin execute immediate 'drop package json_helper'; exception when others then null; end;
  --begin execute immediate 'drop package json_ac'; exception when others then null; end;
  --begin execute immediate 'drop type json force'; exception when others then null; end;
  --begin execute immediate 'drop type json_list force'; exception when others then null; end;
  --begin execute immediate 'drop type json_value_array force'; exception when others then null; end;
  --begin execute immediate 'drop type json_value force'; exception when others then null; end;
  --begin execute immediate 'drop type pljson_narray force'; exception when others then null; end;
  --begin execute immediate 'drop type pljson_vtab force'; exception when others then null; end;
  --begin execute immediate 'drop type pljson_varray force'; exception when others then null; end;
  --begin execute immediate 'drop type pljson_table_impl force'; exception when others then null; end;
  --begin execute immediate 'drop synonym pljson_table'; exception when others then null; end;
  --begin execute immediate 'drop synonym json_table'; exception when others then null; end;

  /* drop new install */
  begin execute immediate 'drop package pljson_parser'; exception when others then null; end;
  begin execute immediate 'drop package pljson_printer'; exception when others then null; end;
  begin execute immediate 'drop package pljson_ext'; exception when others then null; end;
  begin execute immediate 'drop package pljson_dyn'; exception when others then null; end;
  begin execute immediate 'drop package pljson_ml'; exception when others then null; end;
  begin execute immediate 'drop package pljson_xml'; exception when others then null; end;
  begin execute immediate 'drop package pljson_util_pkg'; exception when others then null; end;
  begin execute immediate 'drop package pljson_helper'; exception when others then null; end;
  begin execute immediate 'drop package pljson_ac'; exception when others then null; end;
  begin execute immediate 'drop type pljson force'; exception when others then null; end;
  begin execute immediate 'drop type pljson_list force'; exception when others then null; end;
  begin execute immediate 'drop type pljson_value_array force'; exception when others then null; end;
  begin execute immediate 'drop type pljson_value force'; exception when others then null; end;
  begin execute immediate 'drop type pljson_element force'; exception when others then null; end;
  if n_have1Version=0 Then 
    begin execute immediate 'drop type pljson_narray force'; exception when others then null; end;
    begin execute immediate 'drop type pljson_vtab force'; exception when others then null; end;
    begin execute immediate 'drop type pljson_varray force'; exception when others then null; end;
    begin execute immediate 'drop type pljson_table_impl force'; exception when others then null; end;
  End if;
  --lshuo测试
  --没有单元测试
  --begin execute immediate 'drop package pljson_ut'; exception when others then null; end;
 -- begin execute immediate 'drop table pljson_testsuite'; exception when others then null; end;
  
  if n_have1Version=0 Then begin execute immediate 'drop synonym pljson_table'; exception when others then null; end; end if;
  
  begin execute immediate 'drop public synonym pljson_parser'; exception when others then null; end;
  begin execute immediate 'drop public synonym pljson_printer'; exception when others then null; end;
  begin execute immediate 'drop public synonym pljson_ext'; exception when others then null; end;
  begin execute immediate 'drop public synonym pljson_dyn'; exception when others then null; end;
  begin execute immediate 'drop public synonym pljson_ml'; exception when others then null; end;
  begin execute immediate 'drop public synonym pljson_xml'; exception when others then null; end;
  begin execute immediate 'drop public synonym pljson_util_pkg'; exception when others then null; end;
  begin execute immediate 'drop public synonym pljson_helper'; exception when others then null; end;
  begin execute immediate 'drop public synonym pljson_ac'; exception when others then null; end;
  begin execute immediate 'drop public synonym pljson'; exception when others then null; end;
  begin execute immediate 'drop public synonym pljson_list'; exception when others then null; end;
  begin execute immediate 'drop public synonym pljson_value_array'; exception when others then null; end;
  begin execute immediate 'drop public synonym pljson_value'; exception when others then null; end;
  begin execute immediate 'drop public synonym pljson_element'; exception when others then null; end;
  If n_have1Version=0 Then 
    begin execute immediate 'drop public synonym pljson_narray'; exception when others then null; end;
    begin execute immediate 'drop public synonym pljson_vtab'; exception when others then null; end;
    begin execute immediate 'drop public synonym pljson_varray'; exception when others then null; end;
    begin execute immediate 'drop public synonym pljson_table_impl'; exception when others then null; end;
    begin execute immediate 'drop public synonym pljson_table'; exception when others then null; end;
  End if;
  --lshuo注释
  --不需要单元测试
  --begin execute immediate 'drop public synonym pljson_ut'; exception when others then null; end;
  --begin execute immediate 'drop public synonym pljson_testsuite'; exception when others then null; end;

  --lshuo注释
  --不要卸载3.0版本以前对象
  --begin execute immediate 'drop public synonym json_parser'; exception when others then null; end;
  --begin execute immediate 'drop public synonym json_printer'; exception when others then null; end;
  --begin execute immediate 'drop public synonym json_ext'; exception when others then null; end;
  --begin execute immediate 'drop public synonym json_dyn'; exception when others then null; end;
  --begin execute immediate 'drop public synonym json_ml'; exception when others then null; end;
  --begin execute immediate 'drop public synonym json_xml'; exception when others then null; end;
  --begin execute immediate 'drop public synonym json_util_pkg'; exception when others then null; end;
  --begin execute immediate 'drop public synonym json_helper'; exception when others then null; end;
  --begin execute immediate 'drop public synonym json_ac'; exception when others then null; end;
  --begin execute immediate 'drop public synonym json'; exception when others then null; end;
  --begin execute immediate 'drop public synonym json_list'; exception when others then null; end;
  --begin execute immediate 'drop public synonym json_value_array'; exception when others then null; end;
  --begin execute immediate 'drop public synonym json_value'; exception when others then null; end;
  --begin execute immediate 'drop public synonym json_table'; exception when others then null; end;
  
  --begin execute immediate 'drop synonym json_parser'; exception when others then null; end;
  --begin execute immediate 'drop synonym json_printer'; exception when others then null; end;
  --begin execute immediate 'drop synonym json_ext'; exception when others then null; end;
  --begin execute immediate 'drop synonym json_dyn'; exception when others then null; end;
  --begin execute immediate 'drop synonym json_ml'; exception when others then null; end;
  --begin execute immediate 'drop synonym json_xml'; exception when others then null; end;
  --begin execute immediate 'drop synonym json_util_pkg'; exception when others then null; end;
  --begin execute immediate 'drop synonym json_helper'; exception when others then null; end;
  --begin execute immediate 'drop synonym json_ac'; exception when others then null; end;
  --begin execute immediate 'drop synonym json'; exception when others then null; end;
  --begin execute immediate 'drop synonym json_list'; exception when others then null; end;
  --begin execute immediate 'drop synonym json_value_array'; exception when others then null; end;
  --begin execute immediate 'drop synonym json_value'; exception when others then null; end;
  --begin execute immediate 'drop synonym json_table'; exception when others then null; end;
end;
/
