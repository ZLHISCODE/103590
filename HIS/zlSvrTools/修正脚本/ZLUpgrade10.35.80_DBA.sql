----10.35.70---��10.35.80
--105014:��˶,2017-10-19,�û�������ȱʡ��ռ���ȱʡ��ʱ��ռ�����
Declare
  Arr_Owner t_Strlist;
  Arr_User  t_Strlist;
  v_Sql     Varchar2(2000);
Begin
  --�Ȼ�ȡ���е��ϻ���Ա��������
  Select a.Owner Bulk Collect
  Into Arr_Owner
  From All_Objects a
  Where a.Owner In (Select Distinct b.������ From Zltools.Zlsystems b)
  And a.Object_Name = '�ϻ���Ա��'
  And a.Object_Type = 'TABLE';
  For i In 1 .. Arr_Owner.Count Loop
    --1.�ȴ���ȱʡ��ռ����
    v_Sql := 'Select Username From Dba_Users Where Default_Tablespace Not In (' || Chr(39) || 'USERS' || Chr(39) || ', ' ||
             Chr(39) || 'ZLTOOLSTBS' || Chr(39) || ') And Username In (Select �û��� From ' || Arr_Owner(i) || '.�ϻ���Ա��)';
    Execute Immediate v_Sql Bulk Collect
      Into Arr_User;
    For j In 1 .. Arr_User.Count Loop
      Begin
        --��������ΪUsers��ռ�
        Execute Immediate 'Alter User ' || Arr_User(j) || ' Default Tablespace USERS';
      Exception
        When Others Then
          Begin
            --���ô������õ�ZLTOOLSTBS
            Execute Immediate 'Alter User ' || Arr_User(j) || ' Default Tablespace ZLTOOLSTBS';
            Null;
          Exception
            When Others Then
              Null;
          End;
      End;
    End Loop;
    --2.�ٴ���ȱʡ��ʱ��ռ����
    v_Sql := 'Select Username From Dba_Users Where Temporary_Tablespace Not In (' || Chr(39) || 'TEMP' || Chr(39) || ', ' ||
             Chr(39) || 'ZLTOOLSTMP' || Chr(39) || ') And Username In (Select �û��� From ' || Arr_Owner(i) || '.�ϻ���Ա��)';
    Execute Immediate v_Sql Bulk Collect
      Into Arr_User;
    For j In 1 .. Arr_User.Count Loop
      Begin
        --��������ΪTEMP��ռ�
        Execute Immediate 'Alter User ' || Arr_User(j) || ' Temporary Tablespace TEMP';
      Exception
        When Others Then
          Begin
            --���ô������õ�ZLTOOLSTMP
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

--115010:����,2017-11-7,Ϊzltools������ϵͳ���������Dba_Roles�Ķ�Ȩ��
Declare
  V_Sql Varchar2(1000);
Begin
  For R In (Select Distinct ������ From Zlsystems Union All Select 'ZLTOOLS' From Dual) Loop
    Begin
      --��zltools��ϵͳ��������Ȩ
      V_Sql := 'Grant Select On Sys.Dba_Roles To ' || R.������;
      Execute Immediate V_Sql;
    Exception
      When Others Then
        Null;
        --�����߿��ܲ�����
    End;
  End Loop;
End;
/

--116691:����һ,2017-11-9 ,Ϊzltools�û���� ADMINISTER DATABASE TriggerȨ��
Grant ADMINISTER DATABASE Trigger To zltools;

