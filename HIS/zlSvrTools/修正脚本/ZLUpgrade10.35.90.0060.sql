--本脚本支持从ZLTOOLS v10.35.90 升级到 v10.35.90
--请以管理工具所有者登录PLSQL并执行下列脚本
-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------

-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--140678:刘硕,2019-05-07,上机人员变动消息通知
Create Or Replace Procedure Zltools.Zlmanuser_Edit
(
  Type_In   Number,
  Owner_In  Varchar2,
  用户名_In Varchar2,
  人员id_In Number := Null
  --Type_IN:0-删除用户对应的上机人员，1-新增用户对应的上机人员
  --Owner_IN:上机人员表的所有者
  --用户名_IN:上机人员表.用户名
  --人员id:上机人员表.人员id
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
    Select Count(1) Into n_Count From zlSystems A Where a.编号 = 100 And a.所有者 = Upper(Owner_In);
    If n_Count = 1 Then
      If Type_In = 0 Then
        v_Pkgname_His := Owner_In || '.' || 'b_Message.ZLTOOLS_USERS_001';
      Else
        v_Pkgname_His := Owner_In || '.' || 'b_Message.ZLTOOLS_USERS_002';
      End If;
    End If;
    Select Count(1) Into n_Count From zlSystems A Where a.编号 = 2500 And a.所有者 = Upper(Owner_In);
    If n_Count = 1 Then
      If Type_In = 0 Then
        v_Pkgname_Lis := Owner_In || '.' || 'b_Message_LIS.ZLTOOLS_USERS_001';
      Else
        v_Pkgname_Lis := Owner_In || '.' || 'b_Message_LIS.ZLTOOLS_USERS_002';
      End If;
    End If;
  
    If Not v_Pkgname_Lis Is Null Or Not v_Pkgname_His Is Null Then
      If Type_In = 0 Then
        Execute Immediate 'Select 人员ID From ' || Owner_In || '.上机人员表 Where 用户名=' || Chr(39) || 用户名_In || Chr(39) Bulk
                          Collect
          Into Arr_Manid;
      Else
        Execute Immediate 'Select ' || 人员id_In || ' From Dual' Bulk Collect
          Into Arr_Manid;
      End If;
      For I In 1 .. Arr_Manid.Count Loop
        If n_Type = 0 Then
          If Not v_Pkgname_His Is Null Then
            Executesql('Begin ' || v_Pkgname_His || '(' || Chr(39) || 用户名_In || Chr(39) || ',' || Arr_Manid(I) ||
                       '); End;');
          
            If n_Return = 1 Then
              n_Type := 1;
            End If;
            --可能HIS没有对应接口
            If n_Type = 0 Then
              Executesql('Begin ' || v_Pkgname_Lis || '(' || Chr(39) || 用户名_In || Chr(39) || ',' || Arr_Manid(I) ||
                         '); End;');
            
              If n_Return = 1 Then
                n_Type := 2;
              End If;
            End If;
          Else
            Executesql('Begin ' || v_Pkgname_Lis || '(' || Chr(39) || 用户名_In || Chr(39) || ',' || Arr_Manid(I) ||
                       '); End;');
          
            If n_Return = 1 Then
              n_Type := 2;
            End If;
          End If;
          If n_Type = 0 Then
            --不发送消息,没有对应接口
            n_Type := -1;
          End If;
        Elsif n_Type = 1 Then
          Executesql('Begin ' || v_Pkgname_His || '(' || Chr(39) || 用户名_In || Chr(39) || ',' || Arr_Manid(I) ||
                     '); End;');
        Elsif n_Type = 2 Then
          Executesql('Begin ' || v_Pkgname_Lis || '(' || Chr(39) || 用户名_In || Chr(39) || ',' || Arr_Manid(I) ||
                     '); End;');
        End If;
      End Loop;
    End If;
  Exception
    --可能所有者或者没有权限的用户查询出错
    When Others Then
      Null;
  End;
Begin

  If Type_In = 0 Then
    Sendmessage;
    Execute Immediate 'Delete From ' || Owner_In || '.上机人员表 Where 用户名=' || Chr(39) || 用户名_In || Chr(39);
  Elsif Type_In = 1 Then
    Execute Immediate 'Insert Into ' || Owner_In || '.上机人员表(用户名,人员id) Values (' || Chr(39) || 用户名_In || Chr(39) || ',' ||
                      人员id_In || ')';
    Sendmessage;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/


------------------------------------------------------------------------------------
Commit;