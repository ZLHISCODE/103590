--���ű�֧�ִ�ZLTOOLS v10.35.90 ������ v10.35.90
--���Թ����������ߵ�¼PLSQL��ִ�����нű�
-------------------------------------------------------------------------------
--�ṹ��������
-------------------------------------------------------------------------------

-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--140678:��˶,2019-05-07,�ϻ���Ա�䶯��Ϣ֪ͨ
Create Or Replace Procedure Zltools.Zlmanuser_Edit
(
  Type_In   Number,
  Owner_In  Varchar2,
  �û���_In Varchar2,
  ��Աid_In Number := Null
  --Type_IN:0-ɾ���û���Ӧ���ϻ���Ա��1-�����û���Ӧ���ϻ���Ա
  --Owner_IN:�ϻ���Ա���������
  --�û���_IN:�ϻ���Ա��.�û���
  --��Աid:�ϻ���Ա��.��Աid
) Authid Current_User As
  v_Pkgname_Lis Varchar2(100);
  v_Pkgname_His Varchar2(100);
  n_Type        Number(1) := 0;
  n_Return      Number(1) := 0;
  n_Count       Number(1) := 0;
  Arr_Manid     t_Numlist;
  Procedure Executesql(Sql_In Varchar2) Is
  Begin
    Execute Immediate Sql_In;
    n_Return := 1;
  Exception
    When Others Then
      n_Return := 0;
  End;
  Procedure Sendmessage As
  Begin
    Select Count(1) Into n_Count From zlSystems A Where a.��� = 100 And a.������ = Upper(Owner_In);
    If n_Count = 1 Then
      If Type_In = 0 Then
        v_Pkgname_His := Owner_In || '.' || 'b_Message.ZLTOOLS_USERS_001';
      Else
        v_Pkgname_His := Owner_In || '.' || 'b_Message.ZLTOOLS_USERS_002';
      End If;
    End If;
    Select Count(1) Into n_Count From zlSystems A Where a.��� = 2500 And a.������ = Upper(Owner_In);
    If n_Count = 1 Then
      If Type_In = 0 Then
        v_Pkgname_Lis := Owner_In || '.' || 'b_Message_LIS.ZLTOOLS_USERS_001';
      Else
        v_Pkgname_Lis := Owner_In || '.' || 'b_Message_LIS.ZLTOOLS_USERS_002';
      End If;
    End If;
  
    If Not v_Pkgname_Lis Is Null Or Not v_Pkgname_His Is Null Then
      If Type_In = 0 Then
        Execute Immediate 'Select ��ԱID From ' || Owner_In || '.�ϻ���Ա�� Where �û���=' || Chr(39) || �û���_In || Chr(39) Bulk
                          Collect
          Into Arr_Manid;
      Else
        Execute Immediate 'Select ' || ��Աid_In || ' From Dual' Bulk Collect
          Into Arr_Manid;
      End If;
      For I In 1 .. Arr_Manid.Count Loop
        If n_Type = 0 Then
          If Not v_Pkgname_His Is Null Then
            Executesql('Begin ' || v_Pkgname_His || '(' || Chr(39) || �û���_In || Chr(39) || ',' || Arr_Manid(I) ||
                       '); End;');
          
            If n_Return = 1 Then
              n_Type := 1;
            End If;
            --����HISû�ж�Ӧ�ӿ�
            If n_Type = 0 Then
              Executesql('Begin ' || v_Pkgname_Lis || '(' || Chr(39) || �û���_In || Chr(39) || ',' || Arr_Manid(I) ||
                         '); End;');
            
              If n_Return = 1 Then
                n_Type := 2;
              End If;
            End If;
          Else
            Executesql('Begin ' || v_Pkgname_Lis || '(' || Chr(39) || �û���_In || Chr(39) || ',' || Arr_Manid(I) ||
                       '); End;');
          
            If n_Return = 1 Then
              n_Type := 2;
            End If;
          End If;
          If n_Type = 0 Then
            --��������Ϣ,û�ж�Ӧ�ӿ�
            n_Type := -1;
          End If;
        Elsif n_Type = 1 Then
          Executesql('Begin ' || v_Pkgname_His || '(' || Chr(39) || �û���_In || Chr(39) || ',' || Arr_Manid(I) ||
                     '); End;');
        Elsif n_Type = 2 Then
          Executesql('Begin ' || v_Pkgname_Lis || '(' || Chr(39) || �û���_In || Chr(39) || ',' || Arr_Manid(I) ||
                     '); End;');
        End If;
      End Loop;
    End If;
  Exception
    --���������߻���û��Ȩ�޵��û���ѯ����
    When Others Then
      Null;
  End;
Begin

  If Type_In = 0 Then
    Sendmessage;
    Execute Immediate 'Delete From ' || Owner_In || '.�ϻ���Ա�� Where �û���=' || Chr(39) || �û���_In || Chr(39);
  Elsif Type_In = 1 Then
    Execute Immediate 'Insert Into ' || Owner_In || '.�ϻ���Ա��(�û���,��Աid) Values (' || Chr(39) || �û���_In || Chr(39) || ',' ||
                      ��Աid_In || ')';
    Sendmessage;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/


------------------------------------------------------------------------------------
Commit;