--���ű�֧�ִ�ZLTOOLS v10.35.90 ������ v10.35.90
--���Թ����������ߵ�¼PLSQL��ִ�����нű�
-------------------------------------------------------------------------------
--�ṹ��������
-------------------------------------------------------------------------------

-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--122680:��˶,2018-03-12,ͣ��ɾ����Ա���Ӧ�˻�����
--����˵����
--1���ýű�����������Ա�Ѿ��������Ƕ�Ӧ���ݿ��˻�û���������û��������û�����������¼����̨���������������������ʻ���
--2��������Ա��ɾ�������û�û��ͣ���˻�������Ҫ������������ȡ��ע���е������˻���䡣����ͨ��ע�ӵ�����ѯ�����˻���
--3���ýű�������SYS��¼ִ�С�
Declare
  Arr_Owner t_Strlist;
  Arr_User  t_Strlist;
  v_Sql     Varchar2(2000);
  v_Zlusers Varchar2(2000);
Begin
  --0.�Ȼ�ȡ���е��ϻ���Ա��������
  Select a.Owner Bulk Collect
  Into Arr_Owner
  From All_Objects A
  Where a.Owner In (Select Distinct b.������ From Zltools.Zlsystems B) And a.Object_Name = '�ϻ���Ա��' And
        a.Object_Type = 'TABLE';
  --0.1��ȡ���е���Ա
  For I In 1 .. Arr_Owner.Count Loop
    v_Zlusers := v_Zlusers || ' Union All ' || 'Select Upper(a.�û���) Username, Nvl(b.����ʱ��, Sysdate + 1) ����ʱ�� ' ||
                 ' From ' || Arr_Owner(I) || '.�ϻ���Ա�� A, ' || Arr_Owner(I) || '.��Ա�� B ' || ' Where a.��Աid = b.Id ';
  End Loop;
  --1.��ȡ����ͣ�õ������ݿ�δͣ�õ���Ա
  v_Zlusers := Substr(v_Zlusers, Length(' Union All ') + 1);
  v_Sql     := 'Select C.Username From (Select Username,Max(����ʱ��) ����ʱ�� From (' || v_Zlusers ||
               ') Group by Username ) C,SYS.Dba_Users D Where C.����ʱ��<Sysdate And C.Username=D.Username And D.Account_Status = ''OPEN''';
  --1.1ͣ���˻�
  Execute Immediate v_Sql Bulk Collect
    Into Arr_User;
  For J In 1 .. Arr_User.Count Loop
    Begin
      Execute Immediate 'Alter User ' || Arr_User(J) || '  Account Lock';
    Exception
      When Others Then
        Null;
        --û��Ȩ�ޣ���ǰϵͳ������û��ALter UserȨ��
      --��˲�ȡ��������
    End;
  End Loop;

  --2.��ȡ����ɾ�������ݿ�δͣ�õ���Ա
  --Select c.Username
  --From (Select Distinct a.Username, a.Account_Status
  --       From Sys.Dba_Users A, Sys.Dba_Role_Privs B
  --       Where a.Username Not In ('SYS', 'SYSTEM', 'SCOTT', 'OUTLN', 'DBSNMP', 'MTSSYS', 'MDSYS', 'ORDSYS', 'ORDPLUGINS',
  --                                'CTXSYS', 'ZLTOOLS', 'XDB', 'WMSYS', 'TSMSYS', 'SYSMAN', 'SI_INFORMTN_SCHEMA', 'OLAPSYS',
  --                                'MGMT_VIEW', 'MDDATA', 'EXFSYS', 'DMSYS', 'DIP', 'ANONYMOUS') And
  --             Not a.Default_Tablespace In ('SYSTEM', 'DRSYS') And a.Username Not Like 'ZLBAK%' And
  --             a.Username Not Like 'ZLHD%' And
  --             a.Username Not In (Select Upper(������)
  --                                From Zltools.Zlsystems
  --                                Union All
  --                                Select Upper(������) From Zltools.Zlbakspaces) And a.Account_Status = 'OPEN' And
  --             b.Granted_Role Like 'ZL_%' And a.Username = b.Grantee) C
  --Where c.Username Not in (Select Username
  --     From (Select Upper(a.�û���) Username, Nvl(b.����ʱ��, Sysdate + 1) ����ʱ��
  --              From Zlhis.�ϻ���Ա�� A, Zlhis.��Ա�� B
  --              Where a.��Աid = b.Id)
  --       Group By Username)
  v_Sql := 'Select C.Username From ( Select a.Username From SYS.Dba_Users A, SYS.Dba_Role_Privs B Where a.Username Not In (''SYS'', ''SYSTEM'', ''SCOTT'', ''OUTLN'', ''DBSNMP'', ''MTSSYS'', ''MDSYS'', ''ORDSYS'', ''ORDPLUGINS'',  ''CTXSYS'', ''ZLTOOLS'', ''XDB'', ''WMSYS'', ''TSMSYS'', ''SYSMAN'', ''SI_INFORMTN_SCHEMA'', ''OLAPSYS'',  ''MGMT_VIEW'', ''MDDATA'', ''EXFSYS'', ''DMSYS'', ''DIP'', ''ANONYMOUS'') And  Not a.Default_Tablespace In (''SYSTEM'', ''DRSYS'') And a.Username Not Like ''ZLBAK%'' And a.Username Not Like ''ZLHD%'' And  a.Username Not In (Select Upper(������)  From Zltools.Zlsystems  Union All Select Upper(������) From Zltools.Zlbakspaces) And a.Account_Status = ''OPEN'' And  b.Granted_Role Like ''ZL_%'' And a.Username = b.Grantee ) C Where C.Username Not In(Select Username From (' ||
           v_Zlusers || ') Group by Username)';
  --2.1ͣ���˻�
  Execute Immediate v_Sql Bulk Collect
    Into Arr_User;
  For J In 1 .. Arr_User.Count Loop
    Begin
      --��Ҫ����ɾ����Ա�����û�δͣ���˻�����ȡ�������ע�͡�
      --Execute Immediate 'Alter User ' || Arr_User(J) || '  Account Lock';
      NULL;
    Exception
      When Others Then
        Null;
        --û��Ȩ�ޣ���ǰϵͳ������û��ALter UserȨ��
      --��˲�ȡ��������
    End;
  End Loop;
End;
/

-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



------------------------------------------------------------------------------------
Commit;