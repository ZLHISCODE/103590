----10.35.80---��10.35.90


--117279:��˶,2017-11-30,����ϵͳ������ALter UserȨ��
Declare
  v_Sql Varchar2(2000);
Begin
  For Rsowner In (Select Distinct b.������ From Zltools.Zlsystems B) Loop
    Begin
      v_Sql := 'Grant Alter User To ' || Rsowner.������;
      Execute Immediate v_Sql;
    Exception
      When Others Then
        Null;
        --�����߿��ܲ�����
    End;
  End Loop;
End;
/

