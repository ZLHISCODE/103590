----10.35.70---》10.35.80
--105014:刘硕,2017-10-19,用户创建的缺省表空间与缺省临时表空间修正
Declare
  Arr_Owner t_Strlist;
  Arr_User  t_Strlist;
  v_Sql     Varchar2(2000);
Begin
  --先获取所有的上机人员表所有者
  Select a.Owner Bulk Collect
  Into Arr_Owner
  From All_Objects a
  Where a.Owner In (Select Distinct b.所有者 From Zltools.Zlsystems b)
  And a.Object_Name = '上机人员表'
  And a.Object_Type = 'TABLE';
  For i In 1 .. Arr_Owner.Count Loop
    --1.先处理缺省表空间错误
    v_Sql := 'Select Username From Dba_Users Where Default_Tablespace Not In (' || Chr(39) || 'USERS' || Chr(39) || ', ' ||
             Chr(39) || 'ZLTOOLSTBS' || Chr(39) || ') And Username In (Select 用户名 From ' || Arr_Owner(i) || '.上机人员表)';
    Execute Immediate v_Sql Bulk Collect
      Into Arr_User;
    For j In 1 .. Arr_User.Count Loop
      Begin
        --优先设置为Users表空间
        Execute Immediate 'Alter User ' || Arr_User(j) || ' Default Tablespace USERS';
      Exception
        When Others Then
          Begin
            --设置错误设置到ZLTOOLSTBS
            Execute Immediate 'Alter User ' || Arr_User(j) || ' Default Tablespace ZLTOOLSTBS';
            Null;
          Exception
            When Others Then
              Null;
          End;
      End;
    End Loop;
    --2.再处理缺省临时表空间错误
    v_Sql := 'Select Username From Dba_Users Where Temporary_Tablespace Not In (' || Chr(39) || 'TEMP' || Chr(39) || ', ' ||
             Chr(39) || 'ZLTOOLSTMP' || Chr(39) || ') And Username In (Select 用户名 From ' || Arr_Owner(i) || '.上机人员表)';
    Execute Immediate v_Sql Bulk Collect
      Into Arr_User;
    For j In 1 .. Arr_User.Count Loop
      Begin
        --优先设置为TEMP表空间
        Execute Immediate 'Alter User ' || Arr_User(j) || ' Temporary Tablespace TEMP';
      Exception
        When Others Then
          Begin
            --设置错误设置到ZLTOOLSTMP
            Execute Immediate 'Alter User ' || Arr_User(j) || ' Temporary Tablespace ZLTOOLSTMP';
            Null;
          Exception
            When Others Then
              Null;
          End;
      End;
    End Loop;
  End Loop;
End;
/

--115010:高腾,2017-11-7,为zltools及各个系统所有者添加Dba_Roles的读权限
Declare
  V_Sql Varchar2(1000);
Begin
  For R In (Select Distinct 所有者 From Zlsystems Union All Select 'ZLTOOLS' From Dual) Loop
    Begin
      --对zltools及系统所有者授权
      V_Sql := 'Grant Select On Sys.Dba_Roles To ' || R.所有者;
      Execute Immediate V_Sql;
    Exception
      When Others Then
        Null;
        --所有者可能不存在
    End;
  End Loop;
End;
/

--116691:杨周一,2017-11-9 ,为zltools用户添加 ADMINISTER DATABASE Trigger权限
Grant ADMINISTER DATABASE Trigger To zltools;

