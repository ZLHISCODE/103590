----10.35.80---》10.35.90


--117279:刘硕,2017-11-30,给予系统所有者ALter User权限
Declare
  v_Sql Varchar2(2000);
Begin
  For Rsowner In (Select Distinct b.所有者 From Zltools.Zlsystems B) Loop
    Begin
      v_Sql := 'Grant Alter User To ' || Rsowner.所有者;
      Execute Immediate v_Sql;
    Exception
      When Others Then
        Null;
        --所有者可能不存在
    End;
  End Loop;
End;
/

