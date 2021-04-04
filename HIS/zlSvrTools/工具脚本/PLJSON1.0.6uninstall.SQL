--PLJSON1.0.6反安装处理
Declare
  Procedure Drop_Object(Name_In Varchar2) As
  Begin
    For Rsobject In (Select a.Owner, a.Object_Type, a.Object_Name
                     From All_Objects a
                     Where a.Object_Name = Upper(Name_In)) Loop
      If Rsobject.Object_Type = 'SYNONYM' Then
        If Rsobject.Owner = 'PUBLIC' Then
          Begin Execute Immediate 'Drop Public Synonym ' || Rsobject.Object_Name; Exception When Others Then Null;End;
        Else
          Begin Execute Immediate 'Drop Synonym ' || Rsobject.Owner || '.' || Rsobject.Object_Name; Exception When Others Then Null; End;
        End If;
      Elsif Rsobject.Object_Type = 'TYPE' Or Rsobject.Object_Type = 'TYPE BODY' Then
        Begin Execute Immediate 'Drop ' || Rsobject.Object_Type || ' ' || Rsobject.Owner || '.' || Rsobject.Object_Name || '  Force';Exception When Others Then Null; End;
      Else
        Begin Execute Immediate 'Drop ' || Rsobject.Object_Type || ' ' || Rsobject.Owner || '.' || Rsobject.Object_Name;Exception When Others Then Null;End;
      End If;
    End Loop;
  End;
Begin
  Drop_Object('Json_parser');
  Drop_Object('Json_printer');
  Drop_Object('Json_ext');
  Drop_Object('Json_dyn');
  Drop_Object('Json_ml');
  Drop_Object('Json_xml');
  Drop_Object('Json_util_pkg');
  Drop_Object('Json_helper');
  Drop_Object('Json_ac');
  Drop_Object('Json');
  Drop_Object('Json_member_array');
  Drop_Object('Json_member');
  Drop_Object('Json_list');
  Drop_Object('Json_element_array');
  Drop_Object('Json_element');
  Drop_Object('Json_bool');
  Drop_Object('Json_null');
  Drop_Object('Json_value_array');
  Drop_Object('Json_value');
  Drop_Object('Pljson_narray');
  Drop_Object('Pljson_vtab');
  Drop_Object('Pljson_varray');
  Drop_Object('Pljson_table_impl');
  Drop_Object('Pljson_table');
  Drop_Object('Json_table');
End;
/
