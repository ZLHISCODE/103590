--本脚本支持从ZLTOOLS v10.35.90 升级到 v10.35.90
--请以管理工具所有者登录PLSQL并执行下列脚本
-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------

-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------
--122680:刘硕,2018-03-12,停用删除人员表对应账户锁定
--修正说明：
--1、该脚本用来处理人员已经撤档但是对应数据库账户没有锁定的用户，这类用户可以正常登录导航台。现在修正将锁定该类帐户。
--2、对于人员表删除但是用户没有停用账户，若需要修正，请自行取消注释中的锁定账户语句。可以通过注视的语句查询该类账户。
--3、该脚本建议以SYS登录执行。
Declare
  Arr_Owner t_Strlist;
  Arr_User  t_Strlist;
  v_Sql     Varchar2(2000);
  v_Zlusers Varchar2(2000);
Begin
  --0.先获取所有的上机人员表所有者
  Select a.Owner Bulk Collect
  Into Arr_Owner
  From All_Objects A
  Where a.Owner In (Select Distinct b.所有者 From Zltools.Zlsystems B) And a.Object_Name = '上机人员表' And
        a.Object_Type = 'TABLE';
  --0.1获取所有的人员
  For I In 1 .. Arr_Owner.Count Loop
    v_Zlusers := v_Zlusers || ' Union All ' || 'Select Upper(a.用户名) Username, Nvl(b.撤档时间, Sysdate + 1) 撤档时间 ' ||
                 ' From ' || Arr_Owner(I) || '.上机人员表 A, ' || Arr_Owner(I) || '.人员表 B ' || ' Where a.人员id = b.Id ';
  End Loop;
  --1.获取所有停用但是数据库未停用的人员
  v_Zlusers := Substr(v_Zlusers, Length(' Union All ') + 1);
  v_Sql     := 'Select C.Username From (Select Username,Max(撤档时间) 撤档时间 From (' || v_Zlusers ||
               ') Group by Username ) C,SYS.Dba_Users D Where C.撤档时间<Sysdate And C.Username=D.Username And D.Account_Status = ''OPEN''';
  --1.1停用账户
  Execute Immediate v_Sql Bulk Collect
    Into Arr_User;
  For J In 1 .. Arr_User.Count Loop
    Begin
      Execute Immediate 'Alter User ' || Arr_User(J) || '  Account Lock';
    Exception
      When Others Then
        Null;
        --没有权限，以前系统所有者没有ALter User权限
      --因此采取错误屏蔽
    End;
  End Loop;

  --2.获取所有删除但数据库未停用的人员
  --Select c.Username
  --From (Select Distinct a.Username, a.Account_Status
  --       From Sys.Dba_Users A, Sys.Dba_Role_Privs B
  --       Where a.Username Not In ('SYS', 'SYSTEM', 'SCOTT', 'OUTLN', 'DBSNMP', 'MTSSYS', 'MDSYS', 'ORDSYS', 'ORDPLUGINS',
  --                                'CTXSYS', 'ZLTOOLS', 'XDB', 'WMSYS', 'TSMSYS', 'SYSMAN', 'SI_INFORMTN_SCHEMA', 'OLAPSYS',
  --                                'MGMT_VIEW', 'MDDATA', 'EXFSYS', 'DMSYS', 'DIP', 'ANONYMOUS') And
  --             Not a.Default_Tablespace In ('SYSTEM', 'DRSYS') And a.Username Not Like 'ZLBAK%' And
  --             a.Username Not Like 'ZLHD%' And
  --             a.Username Not In (Select Upper(所有者)
  --                                From Zltools.Zlsystems
  --                                Union All
  --                                Select Upper(所有者) From Zltools.Zlbakspaces) And a.Account_Status = 'OPEN' And
  --             b.Granted_Role Like 'ZL_%' And a.Username = b.Grantee) C
  --Where c.Username Not in (Select Username
  --     From (Select Upper(a.用户名) Username, Nvl(b.撤档时间, Sysdate + 1) 撤档时间
  --              From Zlhis.上机人员表 A, Zlhis.人员表 B
  --              Where a.人员id = b.Id)
  --       Group By Username)
  v_Sql := 'Select C.Username From ( Select a.Username From SYS.Dba_Users A, SYS.Dba_Role_Privs B Where a.Username Not In (''SYS'', ''SYSTEM'', ''SCOTT'', ''OUTLN'', ''DBSNMP'', ''MTSSYS'', ''MDSYS'', ''ORDSYS'', ''ORDPLUGINS'',  ''CTXSYS'', ''ZLTOOLS'', ''XDB'', ''WMSYS'', ''TSMSYS'', ''SYSMAN'', ''SI_INFORMTN_SCHEMA'', ''OLAPSYS'',  ''MGMT_VIEW'', ''MDDATA'', ''EXFSYS'', ''DMSYS'', ''DIP'', ''ANONYMOUS'') And  Not a.Default_Tablespace In (''SYSTEM'', ''DRSYS'') And a.Username Not Like ''ZLBAK%'' And a.Username Not Like ''ZLHD%'' And  a.Username Not In (Select Upper(所有者)  From Zltools.Zlsystems  Union All Select Upper(所有者) From Zltools.Zlbakspaces) And a.Account_Status = ''OPEN'' And  b.Granted_Role Like ''ZL_%'' And a.Username = b.Grantee ) C Where C.Username Not In(Select Username From (' ||
           v_Zlusers || ') Group by Username)';
  --2.1停用账户
  Execute Immediate v_Sql Bulk Collect
    Into Arr_User;
  For J In 1 .. Arr_User.Count Loop
    Begin
      --若要处理删除人员表但是用户未停用账户，请取消下面的注释。
      --Execute Immediate 'Alter User ' || Arr_User(J) || '  Account Lock';
      NULL;
    Exception
      When Others Then
        Null;
        --没有权限，以前系统所有者没有ALter User权限
      --因此采取错误屏蔽
    End;
  End Loop;
End;
/

-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------



------------------------------------------------------------------------------------
Commit;