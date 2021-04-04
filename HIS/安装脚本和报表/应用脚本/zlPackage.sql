--分类目录
--1.公共基础,2.医保基础,3.病人病案基础,4.费用基础,5.药品卫材基础
--6.临床基础,7.临床路径基础,8.病历基础,9.护理基础,10.检验基础
--11.检查基础,12.医保业务,13.病人病案业务,14.费用业务,15.药品卫材业务
--16.临床医嘱,17.临床路径,18.病历业务,19.护理业务,20.检验业务,21.检查业务
----------------------------------------------------------------------------
--[[1.公共基础]]
----------------------------------------------------------------------------
Create Or Replace Package b_Einvoice_Request Is
  ------------------------------------------------------------------
  --电子票据业务处理
  --  1.Einvoice_Start-判断电子票据是否启用(返回:1-启用;0-未启用)
  --  2.EInvoice_Create-电子票据开具(返回1-成功;0-失败)
  --  3.Einvoice_Cancel_Check-电子票据作废前检查(返回:1-合法;0-不合法)
  --  4.Einvoice_Cancel-电子票据作废(返回1-成功;0-失败)
  ------------------------------------------------------------------
  --1.判断电子票据是否启用
  Function Einvoice_Start
  (
    业务场景_In Integer,
    险类_In     保险结算记录.险类%Type,
    类型_In     Integer := Null
  ) Return Number;

  --2.电子票据开具
  Function Einvoice_Create
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    冲销id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number;

  --3.电子票据作废检查
  Function Einvoice_Cancel_Check
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number;

  --4.电子票据作废
  Function Einvoice_Cancel
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number;
End b_Einvoice_Request;
/
Create Or Replace Package b_Common_Context As
  --速度：包缓存>全局上下文（3-6倍）>表查询（3-6倍）
  --用于设置全局上下文。
  Procedure Set_Context
  (
    Namespace_In In Varchar2,
    Name_In      In Varchar2,
    Value_In     In Varchar2
  );
  --清理上下文。
  Procedure Clear_Context
  (
    Namespace_In In Varchar2,
    Name_In      In Varchar2 := Null
  );
End b_Common_Context;
/
--144329:刘硕,2019-09-04,全局上下文代替缓存
CREATE OR REPLACE Package Body b_Einvoice_Request Is
  ------------------------------------------------------------------
  --电子票据业务处理
  --  1.Einvoice_Start-判断电子票据是否启用(返回:1-启用;0-未启用)
  --  2.EInvoice_Create-电子票据开具(返回1-成功;0-失败)
  --  3.Einvoice_Cancel_Check-电子票据作废前检查(返回:1-合法;0-不合法)
  --  4.Einvoice_Cancel-电子票据作废(返回1-成功;0-失败)
  ------------------------------------------------------------------

  Function Einvoice_Start
  (
    业务场景_In Integer,
    险类_In     保险结算记录.险类%Type,
    类型_In     Integer := Null
  ) Return Number Is
    ------------------------------------------------------------------
    --功能:判断电子票据是否启用
    --入参:业务场景_In-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    --     类型_in:NULL-不区分类型;针对场合为结账及预交:1-门诊;2-住院;
    --出参:错误信息_Out-返回的错误信息
    --返回:1-启用;0-未启用
    -------------------------------------------------------------------
    v_包名称   电子票据类别.包名称%Type;
    v_Sql      Varchar2(1000);
    n_Return   Number(2);
    n_启用     Number(2);
    n_Err_Code Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin

    If Nvl(业务场景_In, 0) = 2 And Nvl(类型_In, 0) = 1 Then
      --门诊预交，暂不支持
      Return 0;
    End If;

    Begin
      n_启用 := 1;
      Select 包名称 Into v_包名称 From 电子票据类别 Where 是否启用 = 1;
    Exception
      When Others Then
        n_启用 := 0;
    End;
    If n_启用 = 0 Or v_包名称 Is Null Then
      --未启用或无包名称，直接返回0，表示成功;
      Return 0;
    End If;

    n_Err_Code := Null;
    Begin
      v_Sql := 'begin :n_return:=' || v_包名称 || '.Einvoice_Start(:1,:2,:3); end;';
      Execute Immediate v_Sql
        Using Out n_Return, In 业务场景_In, 险类_In, 类型_In;
      Return n_Return;
    Exception
      When Others Then
        n_Err_Code := SQLCode;
    End;
    If n_Err_Code = -6550 Or n_Err_Code Is Null Then
      Return 0;
    End If;
    Return 0;

  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Start;

  Function Einvoice_Create
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    冲销id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number Is
    ------------------------------------------------------------------
    --功能:电子票据开具
    --入参:业务场景_In-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    --     结算ID_In-业务场景_In=2(预交)时：原预交ID,业务场景_In<>2(预交)时：原结帐ID
    --     冲销ID_In-业务场景_In=2(预交)时：退款的预交ID,业务场景_In<>2(预交)时：当前退费的结帐ID,部分退费时有效;
    --出参:错误信息_Out-返回的错误信息
    --返回:1-成功;0-失败
    -------------------------------------------------------------------
    v_包名称      电子票据类别.包名称%Type;
    v_Sql         Varchar2(1000);
    n_Return      Number(2);
    v_Err_Msg_Out Varchar2(4000);
    n_启用        Number(2);
    n_Err_Code    Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin

    Begin
      n_启用 := 1;
      Select 包名称 Into v_包名称 From 电子票据类别 Where 是否启用 = 1;
    Exception
      When Others Then
        n_启用 := 0;
    End;
    If n_启用 = 0 Or v_包名称 Is Null Then
      --未启用或无包名称，直接返回1，表示成功;
      Return 1;
    End If;

    n_Err_Code := Null;
    Begin
      v_Sql := 'begin :n_return:=' || v_包名称 || '.EInvoice_Create(:1,:2,:3,:v_Err_Msg_out); end;';
      Execute Immediate v_Sql
        Using Out n_Return, In 业务场景_In, 结算id_In, 冲销id_In, Out v_Err_Msg_Out;
      错误信息_Out := v_Err_Msg_Out;
      Return n_Return;
    Exception
      When Others Then
        n_Err_Code    := SQLCode;
        v_Err_Msg_Out := SQLErrM;
    End;
    If n_Err_Code = -6550 Or n_Err_Code Is Null Then
      --没有此过程，返回true
      Return 1;
    End If;
    Raise Err_Item;

  Exception
    When Err_Item Then
      zl_ErrorCenter(n_Err_Code, v_Err_Msg_Out);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Create;

  Function Einvoice_Cancel_Check
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number Is
    ------------------------------------------------------------------
    --功能:电子票据作废
    --入参:业务场景_In-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    --     结算ID_In-业务场景_In=2(预交)时：原预交ID,业务场景_In<>2(预交)时：原结帐ID
    --出参:错误信息_Out-返回的错误信息
    --返回:1-成功;0-失败
    -------------------------------------------------------------------
    v_包名称      电子票据类别.包名称%Type;
    v_Sql         Varchar2(1000);
    n_Return      Number(2);
    v_Err_Msg_Out Varchar2(4000);
    n_启用        Number(2);
    n_Err_Code    Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin

    If 业务场景_In = 2 Then
      --预交款
      Select Max(Nvl(预交电子票据, 0)) Into n_Return From 病人预交记录 Where 结帐id = 结算id_In;
    Else
      --非预交：收费、结帐、挂号及就诊卡
      Select Max(Nvl(是否电子票据, 0)) Into n_Return From 病人预交记录 Where 结帐id = 结算id_In;
    End If;
    If Nvl(n_Return, 0) = 0 Then
      --该记录是未启用电子票据的，直接返回1;
      Return 1;
    End If;

    Begin
      n_启用 := 1;
      Select 包名称 Into v_包名称 From 电子票据类别 Where 是否启用 = 1;
    Exception
      When Others Then
        n_启用 := 0;
    End;

    If n_启用 = 0 Or v_包名称 Is Null Then
      错误信息_Out := '电子票据未启用，请在窗口中进行退费或退款，在此处禁止发起电子票据冲红。';
      Return 0;
    End If;

    --检查是否换开过电子票据
    For c_电子票据 In (Select ID, 是否换开, 纸质发票号
                   From 电子票据使用记录
                   Where 票种 = 业务场景_In And 记录状态 = 1 And 结算id = 结算id_In) Loop
      --针对电子票据进行处理
      If Nvl(c_电子票据.是否换开, 0) = 1 Then
        --换开纸质发票号，禁止作废操作
        错误信息_Out := '已经换开纸质发票(' || c_电子票据.纸质发票号 || ')，在此处禁止发起电子票据冲红。';
        Return 0;
      End If;
      n_Err_Code := Null;
      Begin
        v_Sql := 'begin :n_return:=' || v_包名称 || '.Einvoice_Cancel_Check(:1,:2,:v_Err_Msg_out); end;';
        Execute Immediate v_Sql
          Using Out n_Return, In 业务场景_In, c_电子票据.Id, Out v_Err_Msg_Out;
        错误信息_Out := v_Err_Msg_Out;
        Return n_Return;
      Exception
        When Others Then
          n_Err_Code    := SQLCode;
          v_Err_Msg_Out := SQLErrM;
      End;

      If n_Err_Code = -6550 Then
        错误信息_Out := '电子票据未启用，请在窗口中进行退费或退款，在此处禁止发起电子票据冲红。';
        Return 0;
      End If;
      If Not n_Err_Code Is Null Then
        Raise Err_Item;
      End If;

    End Loop;
    Return 1;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg_Out || '[ZLSOFT]');
    When Err_Item Then
      zl_ErrorCenter(n_Err_Code, v_Err_Msg_Out);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Cancel_Check;

  Function Einvoice_Cancel
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number Is
    ------------------------------------------------------------------
    --功能:电子票据作废
    --入参:业务场景_In-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    --     结算ID_In-业务场景_In=2(预交)时：原预交ID,业务场景_In<>2(预交)时：原结帐ID
    --出参:错误信息_Out-返回的错误信息
    --返回:1-成功;0-失败
    -------------------------------------------------------------------
    v_包名称      电子票据类别.包名称%Type;
    v_Sql         Varchar2(1000);
    n_Return      Number(2);
    v_Err_Msg_Out Varchar2(4000);
    n_启用        Number(2);
    n_Err_Code    Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin

    If 业务场景_In = 2 Then
      --预交款
      Select Max(Nvl(预交电子票据, 0)) Into n_Return From 病人预交记录 Where 结帐id = 结算id_In;
    Else
      --非预交：收费、结帐、挂号及就诊卡
      Select Max(Nvl(是否电子票据, 0)) Into n_Return From 病人预交记录 Where 结帐id = 结算id_In;
    End If;
    If Nvl(n_Return, 0) = 0 Then
      --该记录是未启用电子票据的，直接返回1;
      Return 1;
    End If;

    Begin
      n_启用 := 1;
      Select 包名称 Into v_包名称 From 电子票据类别 Where 是否启用 = 1;
    Exception
      When Others Then
        n_启用 := 0;
    End;

    If n_启用 = 0 Or v_包名称 Is Null Then
      错误信息_Out := '电子票据未启用，请在窗口中进行退费或退款，在此处禁止发起电子票据冲红。';
      Return 0;
    End If;

    --检查是否换开过电子票据
    For c_电子票据 In (Select ID, 是否换开, 纸质发票号
                   From 电子票据使用记录
                   Where 票种 = 业务场景_In And 记录状态 = 1 And 结算id = 结算id_In) Loop
      --针对电子票据进行处理
      If Nvl(c_电子票据.是否换开, 0) = 1 Then
        --换开纸质发票号，禁止作废操作
        错误信息_Out := '已经换开纸质发票(' || c_电子票据.纸质发票号 || ')，在此处禁止发起电子票据冲红。';
        Return 0;
      End If;
      n_Err_Code := Null;

      --避免并发原因，还是需要先进行检查电子票据是否允许冲红。
      Begin
        v_Sql := 'begin :n_return:=' || v_包名称 || '.Einvoice_Cancel_Check(:1,:2,:v_Err_Msg_out); end;';
        Execute Immediate v_Sql
          Using Out n_Return, In 业务场景_In, 结算id_In, Out v_Err_Msg_Out;
        错误信息_Out := v_Err_Msg_Out;
      Exception
        When Others Then
          n_Err_Code    := SQLCode;
          v_Err_Msg_Out := SQLErrM;
      End;

      If n_Err_Code = -6550 Then
        错误信息_Out := '电子票据未启用，请在窗口中进行退费或退款，在此处禁止发起电子票据冲红。';
        Return 0;
      End If;
      If Not n_Err_Code Is Null Or n_Return = 0 Then
        Raise Err_Item;
      End If;

      --进行电子票据冲红处理
      n_Return := 0;
      Begin
        v_Sql := 'begin :n_return:=' || v_包名称 || '.Einvoice_Cancel(:1,:2,:v_Err_Msg_out); end;';
        Execute Immediate v_Sql
          Using Out n_Return, In 业务场景_In, 结算id_In, Out v_Err_Msg_Out;
        错误信息_Out := v_Err_Msg_Out;
        Return n_Return;
      Exception
        When Others Then
          n_Err_Code    := SQLCode;
          v_Err_Msg_Out := SQLErrM;
      End;

      If n_Err_Code = -6550 Then
        错误信息_Out := '电子票据未启用，请在窗口中进行退费或退款，在此处禁止发起电子票据冲红。';
        Return 0;
      End If;
      If Not n_Err_Code Is Null Then
        Raise Err_Item;
      End If;
    End Loop;
    Return 1;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg_Out || '[ZLSOFT]');
    When Err_Item Then
      zl_ErrorCenter(n_Err_Code, v_Err_Msg_Out);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Cancel;
End b_Einvoice_Request;
/
Create Or Replace Package Body b_Common_Context As
  Procedure Set_Context
  (
    Namespace_In In Varchar2,
    Name_In      In Varchar2,
    Value_In     In Varchar2
  ) As
  Begin
    Dbms_Session.Set_Context(Namespace_In, Name_In, Value_In, Null, Null);
  End Set_Context;
  Procedure Clear_Context
  (
    Namespace_In In Varchar2,
    Name_In      In Varchar2 := Null
  ) As
  Begin
    Dbms_Session.Clear_Context(Namespace_In, Null, Name_In);
  End Clear_Context;
End b_Common_Context;
/
create or replace context zlMessageCtx using b_Common_Context ACCESSED GLOBALLY;
Create Or Replace Package b_Zlmsg_Cache Is
  --1、名称和b_Message_Cache以及b_Message区别，否则可能会导致启停时会调用到
  --2、该包替代b_Message_Cache(10.35.130)，以及替代b_Message(<10.35.130)的缓存部分

  --独立消息包缓存，防止手工编译导致问题 
  --1、该包的修改以及编译请尽量放在业务低谷时候处理 
  --2、请勿在该包中增加全局变量，如： 
  --Message_Creator Zlmsg_Todo.Creator%Type := Null; 
  --如确实需要增加全局变量（或已经存在全局变量），请在PLSQL中编译后执行如下语句： 
  --ALTER PACKAGE b_zlMessage_Cache COMPILE SPECIFICATION 

  --判断消息是否启用 
  Function Is_Message_Using(Msg_Code_In Zlmsg_Lists.Code%Type) Return Number;
  --设置当前会话是平台调用 
  Procedure Set_Platform_Call(Platform_Call Number);
  --获取消息创建者 
  Function Get_Message_Creator Return Varchar2;
End b_Zlmsg_Cache;
/
Create Or Replace Package Body b_Zlmsg_Cache Is
  --1、名称和b_Message_Cache以及b_Message区别，否则可能会导致启停时会调用到
  --2、该包替代b_Message_Cache(10.35.130)，以及替代b_Message(<10.35.130)的缓存部分

  --独立消息包缓存，防止手工编译导致问题 
  --1、该包的修改以及编译请尽量放在业务低谷时候处理 
  --2、请勿在该包中增加全局变量，如： 
  --Message_Creator Zlmsg_Todo.Creator%Type := Null; 
  --如确实需要增加全局变量（或已经存在全局变量），请在PLSQL中编译后执行如下语句： 
  --ALTER PACKAGE b_zlMessage_Cache COMPILE SPECIFICATION 

  --是否是平台调用 
  Is_Platform_Call Number(1) := 0;
  --消息公共方法 
  Message_Creator Zlmsg_Todo.Creator%Type := Null;
  --是否启用消息
  Function Is_Message_Using(Msg_Code_In Zlmsg_Lists.Code%Type) Return Number As
    n_Using Zlmsg_Lists.Using%Type;
    v_Code  Zlmsg_Lists.Code%Type;
  Begin
    If Is_Platform_Call = 1 Then
      Return 0;
    End If;
    v_Code := Upper(Msg_Code_In);
    --类型转换，可以转换
    n_Using := Sys_Context('zlMessageCtx', v_Code);
    If n_Using Is Null Then
      --不采取Max容错处理，错误相当于外键,用户可能没有采取同步修改或自己增加了消息类型但是未注册到Zlmsg_Lists，这两种情况会出现错误。  
      Select Nvl(Using, 0) Into n_Using From Zlmsg_Lists A Where Code = v_Code;
      --暂时不考虑同一个实例下存在相同的Zlmsg_Lists表问题，认为整个实例只有一个Zlmsg_Lists
      b_Common_Context.Set_Context('zlMessageCtx', v_Code, n_Using);
    End If;
    Return n_Using;
  Exception
    When Others Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || '未在Zlmsg_Lists中找到消息"' || v_Code || '"！请联系管理员进行处理。' || '[ZLSOFT]');
      Return 0;
  End;
  --设置当前会话为平台调用 
  Procedure Set_Platform_Call(Platform_Call Number) Is
  Begin
    Is_Platform_Call := Platform_Call;
  End Set_Platform_Call;
  --查询生成消息的人员
  Function Get_Message_Creator Return Varchar2 As
  Begin
    Return Message_Creator;
  End;
Begin
  --包初次实例化执行一次
  Message_Creator := zl_UserName;
End b_Zlmsg_Cache;
/

Create Or Replace Package b_Message Is
  --独立消息包缓存，防止手工编译导致问题
  --1、该包的修改以及编译请尽量放在业务低谷时候处理
  --2、请勿在该包中增加全局变量，如：
  --Message_Creator Zlmsg_Todo.Creator%Type := Null;
  --如确实需要增加全局变量（或已经存在全局变量），请在PLSQL中编译后执行如下语句：
  --ALTER PACKAGE b_Message COMPILE SPECIFICATION

  Type c_Dynamic Is Ref Cursor;

  --设置平台调用类型
  Procedure Set_Platform_Call(Platform_Call Number);
  --新增部门
  Procedure Zlhis_Dict_001(Id_In Number);
  --修改部门
  Procedure Zlhis_Dict_002(部门id_In Number);
  --停用部门
  Procedure Zlhis_Dict_003(部门id_In Number);
  --启用部门
  Procedure Zlhis_Dict_004(部门id_In Number);
  --新增人员
  Procedure Zlhis_Dict_005(人员id_In Number);
  --修改人员
  Procedure Zlhis_Dict_006(人员id_In Number);
  --停用人员
  Procedure Zlhis_Dict_007(人员id_In Number);
  --启用人员
  Procedure Zlhis_Dict_008(人员id_In Number);
  --新增收费项目
  Procedure Zlhis_Dict_009(细目id_In Number);
  --修改收费项目
  Procedure Zlhis_Dict_010(细目id_In Number);
  --停用收费项目
  Procedure Zlhis_Dict_011(细目id_In Number);
  --启用收费项目
  Procedure Zlhis_Dict_012(细目id_In Number);
  --新增诊疗项目
  Procedure Zlhis_Dict_013(诊疗id_In Number);
  --修改诊疗项目
  Procedure Zlhis_Dict_014(诊疗id_In Number);
  --停用诊疗项目
  Procedure Zlhis_Dict_015(诊疗id_In Number);
  --启用诊疗项目
  Procedure Zlhis_Dict_016(诊疗id_In Number);
  --新增检验项目
  Procedure Zlhis_Dict_017(诊疗id_In Number);
  --修改检验项目
  Procedure Zlhis_Dict_018(诊疗id_In Number);
  --删除检验项目
  Procedure Zlhis_Dict_019
  (
    诊疗id_In Number,
    编码_In   Varchar2,
    中文名_In Varchar2,
    英文名_In Varchar2
  );

  --新增疾病编码目录
  Procedure Zlhis_Dict_021(疾病id_In Number);
  --修改疾病编码目录
  Procedure Zlhis_Dict_022(疾病id_In Number);
  --停用疾病编码目录
  Procedure Zlhis_Dict_023(疾病id_In Number);
  --启用疾病编码目录
  Procedure Zlhis_Dict_024(疾病id_In Number);
  --新增药品分类
  Procedure Zlhis_Dict_025
  (
    类型_In Number,
    Id_In   Number
  );
  --修改药品分类
  Procedure Zlhis_Dict_026
  (
    类型_In Number,
    Id_In   Number
  );
  --删除药品分类
  Procedure Zlhis_Dict_027
  (
    类型_In Number,
    Id_In   Number,
    编码_In Varchar2,
    名称_In Varchar2
  );
  --停用药品分类
  Procedure Zlhis_Dict_028
  (
    类型_In Number,
    Id_In   Number
  );
  --启用药品分类
  Procedure Zlhis_Dict_029
  (
    类型_In Number,
    Id_In   Number
  );
  --新增药品品种
  Procedure Zlhis_Dict_030
  (
    类别_In Varchar2,
    Id_In   Number
  );
  --修改药品品种
  Procedure Zlhis_Dict_031
  (
    类别_In Varchar2,
    Id_In   Number
  );
  --删除药品品种
  Procedure Zlhis_Dict_032
  (
    类别_In Varchar2,
    Id_In   Number,
    编码_In Varchar2,
    名称_In Varchar2
  );
  --停用药品品种
  Procedure Zlhis_Dict_033
  (
    类别_In Varchar2,
    Id_In   Number
  );
  --启用药品品种
  Procedure Zlhis_Dict_034
  (
    类别_In Varchar2,
    Id_In   Number
  );
  --新增药品规格
  Procedure Zlhis_Dict_035
  (
    类别_In   Varchar2,
    药品id_In Number
  );
  --修改药品规格
  Procedure Zlhis_Dict_036
  (
    类别_In   Varchar2,
    药品id_In Number
  );
  --删除药品规格
  Procedure Zlhis_Dict_037
  (
    类别_In   Varchar2,
    药品id_In Number,
    编码_In   Varchar2,
    名称_In   Varchar2,
    规格_In   Varchar2,
    产地_In   Varchar2
  );
  --停用药品规格
  Procedure Zlhis_Dict_038
  (
    类别_In   Varchar2,
    药品id_In Number
  );
  --启用药品规格
  Procedure Zlhis_Dict_039
  (
    类别_In   Varchar2,
    药品id_In Number
  );
  --设置药品存储库房
  Procedure Zlhis_Dict_040
  (
    类别_In   Varchar2,
    药品id_In Number
  );
  --设置药品储备限额
  Procedure Zlhis_Dict_041
  (
    类别_In   Varchar2,
    药品id_In Number
  );
  --新增卫材品种
  Procedure Zlhis_Dict_042(Id_In Number);
  --新增卫材规格
  Procedure Zlhis_Dict_043(Id_In Number);
  --修改卫材规格
  Procedure Zlhis_Dict_044(Id_In Number);
  --删除卫材规格
  Procedure Zlhis_Dict_045
  (
    Id_In   Number,
    编码_In Varchar2,
    名称_In Varchar2,
    规格_In Varchar2,
    产地_In Varchar2
  );
  --停用卫材规格
  Procedure Zlhis_Dict_046(Id_In Number);
  --启用卫材规格
  Procedure Zlhis_Dict_047(Id_In Number);
  --医保对码
  Procedure Zlhis_Dict_048
  (
    险类_In       In Number,
    收费细目id_In In Number
  );
  --删除医保对码
  Procedure Zlhis_Dict_049
  (
    险类_In       In Number,
    收费细目id_In In Number,
    项目编码_In   In Varchar2,
    项目名称_In   In Varchar2,
    医保编码_In   In Varchar2,
    医保名称_In   In Varchar2
  );
  --新增卫材分类
  Procedure Zlhis_Dict_050
  (
    类型_In Number,
    Id_In   Number
  );
  --修改卫材分类
  Procedure Zlhis_Dict_051
  (
    类型_In Number,
    Id_In   Number
  );
  --删除卫材分类
  Procedure Zlhis_Dict_052
  (
    类型_In Number,
    Id_In   Number,
    编码_In Varchar2,
    名称_In Varchar2
  );
  --收费价目变动
  Procedure Zlhis_Dict_053(收费项目id_In Number);
  --诊疗收费对照变动
  Procedure Zlhis_Dict_054(诊疗项目id_In Number);
  --卫材存储库房变动
  Procedure Zlhis_Dict_055(细目id_In Varchar2);
  --新增诊疗检查类型
  Procedure Zlhis_Dictpacs_001
  (
    编码_In   Varchar2,
    名称_In   Varchar2,
    简码_In   Varchar2,
    建病案_In Number
  );
  --修改诊疗检查类型
  Procedure Zlhis_Dictpacs_002
  (
    编码_In   Varchar2,
    名称_In   Varchar2,
    简码_In   Varchar2,
    建病案_In Number
  );
  --删除诊疗检查类型
  Procedure Zlhis_Dictpacs_003
  (
    编码_In   Varchar2,
    名称_In   Varchar2,
    简码_In   Varchar2,
    建病案_In Number
  );
  --新增诊疗检查部位
  Procedure Zlhis_Dictpacs_004
  (
    类型_In     Varchar2,
    编码_In     Varchar2,
    名称_In     Varchar2,
    分组_In     Varchar2,
    备注_In     Varchar2,
    方法_In     Varchar2,
    适用性别_In Number
  );
  --修改诊疗检查部位
  Procedure Zlhis_Dictpacs_005
  (
    类型_In     Varchar2,
    编码_In     Varchar2,
    名称_In     Varchar2,
    分组_In     Varchar2,
    备注_In     Varchar2,
    方法_In     Varchar2,
    适用性别_In Number
  );
  --删除诊疗检查部位
  Procedure Zlhis_Dictpacs_006
  (
    类型_In     Varchar2,
    编码_In     Varchar2,
    名称_In     Varchar2,
    分组_In     Varchar2,
    备注_In     Varchar2,
    方法_In     Varchar2,
    适用性别_In Number
  );
  --新增诊疗项目部位
  Procedure Zlhis_Dictpacs_007
  (
    Id_In       Number,
    项目id_In   Number,
    类型_In     Varchar2,
    部位_In     Varchar2,
    方法_In     Varchar2,
    默认_In     Number,
    上级方法_In Varchar2
  );
  --修改诊疗项目部位
  Procedure Zlhis_Dictpacs_008
  (
    Id_In       Number,
    项目id_In   Number,
    类型_In     Varchar2,
    部位_In     Varchar2,
    方法_In     Varchar2,
    默认_In     Number,
    上级方法_In Varchar2
  );
  --删除诊疗项目部位
  Procedure Zlhis_Dictpacs_009
  (
    Id_In       Number,
    项目id_In   Number,
    类型_In     Varchar2,
    部位_In     Varchar2,
    方法_In     Varchar2,
    默认_In     Number,
    上级方法_In Varchar2
  );
  --新增诊疗检验标本
  Procedure Zlhis_Dictlis_004
  (
    编码_In     Varchar2,
    名称_In     Varchar2,
    简码_In     Varchar2,
    适用性别_In Varchar2
  );
  --修改诊疗检验标本
  Procedure Zlhis_Dictlis_005
  (
    编码_In     Varchar2,
    名称_In     Varchar2,
    简码_In     Varchar2,
    适用性别_In Varchar2
  );
  --删除诊疗项目部位
  Procedure Zlhis_Dictlis_006
  (
    编码_In     Varchar2,
    名称_In     Varchar2,
    简码_In     Varchar2,
    适用性别_In Varchar2
  );
  --新增采血管类型
  Procedure Zlhis_Dictlis_007
  (
    编码_In   Varchar2,
    名称_In   Varchar2,
    简码_In   Varchar2,
    添加剂_In Varchar2,
    采血量_In Varchar2,
    规格_In   Varchar2,
    颜色_In   Number,
    材料id_In Number
  );
  --修改采血管类型
  Procedure Zlhis_Dictlis_008
  (
    编码_In   Varchar2,
    名称_In   Varchar2,
    简码_In   Varchar2,
    添加剂_In Varchar2,
    采血量_In Varchar2,
    规格_In   Varchar2,
    颜色_In   Number,
    材料id_In Number
  );
  --删除采血管类型
  Procedure Zlhis_Dictlis_009
  (
    编码_In   Varchar2,
    名称_In   Varchar2,
    简码_In   Varchar2,
    添加剂_In Varchar2,
    采血量_In Varchar2,
    规格_In   Varchar2,
    颜色_In   Number,
    材料id_In Number
  );

  --药品备药发送
  Procedure Zlhis_Drug_001(No_In Varchar2);
  --取消药品备药发送
  Procedure Zlhis_Drug_002(No_In Varchar2);
  --药品移库单接收
  Procedure Zlhis_Drug_003(No_In Varchar2);
  --药品移库单冲销
  Procedure Zlhis_Drug_004
  (
    No_In       Varchar2,
    序号_In     Number,
    记录状态_In Number
  );
  --部门发药
  Procedure Zlhis_Drug_005
  (
    库房id_In Number,
    收发id_In Number
  );
  --部门退药
  Procedure Zlhis_Drug_006
  (
    冲销收发id_In Number,
    待发收发id_In Number,
    数量_In       Number,
    费用id_In     Number
  );
  --药品调价
  Procedure Zlhis_Drug_007(价格id_In Number);
  --静配发送
  Procedure Zlhis_Drug_008(记录ids_In Varchar2);
  --药品调售价
  Procedure Zlhis_Drug_009
  (
    价格id_In Number,
    时价_In   Number
  );
  --卫材调成本价
  Procedure Zlhis_Drug_010(价格id_In Number);
  --卫材调售价
  Procedure Zlhis_Drug_011
  (
    价格id_In Number,
    时价_In   Number
  );
  --卫材发料
  Procedure Zlhis_Drug_012
  (
    库房id_In Number,
    收发id_In Number
  );
  --卫材退料
  Procedure Zlhis_Drug_013
  (
    冲销收发id_In Number,
    待发收发id_In Number,
    数量_In       Number,
    费用id_In     Number
  );
  --2.停止患者医嘱，住院
  Procedure Zlhis_Cis_002
  (
    病人id_In  In Number,
    主页id_In  In Number,
    医嘱id_In  In Number,
    医嘱ids_In In Varchar2
  );
  --3.作废患者医嘱，门诊/住院
  Procedure Zlhis_Cis_003
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    医嘱id_In In Number
  );

  --4.患者术后医嘱，住院
  Procedure Zlhis_Cis_004
  (
    病人id_In In Number,
    主页id_In In Number,
    医嘱id_In In Number
  );

  --5.撤消患者术后医嘱，住院
  Procedure Zlhis_Cis_005
  (
    病人id_In In Number,
    主页id_In In Number,
    医嘱id_In In Number
  );

  --6.患者护理常规医嘱，住院
  Procedure Zlhis_Cis_006
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number
  );

  --7.撤消患者护理常规医嘱，住院
  Procedure Zlhis_Cis_007
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number,
    No_In     In Varchar2
  );

  --门诊患者接诊
  Procedure Zlhis_Cis_008
  (
    病人id_In In Number,
    挂号单_In In Varchar2
  );

  --门诊患者取消接诊
  Procedure Zlhis_Cis_009
  (
    病人id_In In Number,
    挂号单_In In Varchar2
  );

  --10.下达患者诊断，门诊/住院
  Procedure Zlhis_Cis_010
  (
    病人id_In In Number,
    就诊id_In In Number, --门诊病人 挂号ID，住院病人 主页ID
    诊断id_In In Number
  );
  --11.撤消患者诊断
  Procedure Zlhis_Cis_011
  (
    病人id_In   In Number,
    就诊id_In   In Number, --门诊病人 挂号ID，住院病人 主页ID
    Id_In       In Number,
    疾病id_In   In Number,
    诊断id_In   In Number,
    诊断描述_In In Varchar2
  );

  --病区执行医嘱校对
  Procedure Zlhis_Cis_012
  (
    病人id_In In Number,
    主页id_In In Number,
    医嘱id_In In Number
  );

  --13.检验危急值阅读通知
  Procedure Zlhis_Cis_014
  (
    病人id_In In Number,
    就诊id_In In Number, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In Number,
    消息id_In In Number
  );

  --15.患者检验申请，门诊/住院
  Procedure Zlhis_Cis_016
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    发送号_In   In Number,
    医嘱id_In   In Number,
    病人来源_In In Number
  );
  --16.患者检查申请，门诊/住院
  Procedure Zlhis_Cis_017
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    发送号_In   In Number,
    医嘱id_In   In Number,
    病人来源_In In Number
  );
  --17.患者手术申请，门诊/住院
  Procedure Zlhis_Cis_018
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    发送号_In In Number,
    医嘱id_In In Number
  );
  --18.患者输血申请，住院
  Procedure Zlhis_Cis_019
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    发送号_In In Number,
    医嘱id_In In Number
  );
  --19.患者会诊申请，住院
  Procedure Zlhis_Cis_020
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number
  );
  --20.患者抢救医嘱，住院
  Procedure Zlhis_Cis_021
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number
  );
  --21.患者死亡医嘱，住院
  Procedure Zlhis_Cis_022
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number
  );
  --22.患者特殊治疗医嘱，住院
  Procedure Zlhis_Cis_023
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number
  );

  --24.检查危急值阅读通知
  Procedure Zlhis_Cis_025
  (
    病人id_In In Number,
    就诊id_In In Number, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In Number,
    消息id_In In Number
  );

  --病区执行医嘱发送
  Procedure Zlhis_Cis_026
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number
  );

  --撤消患者检验申请
  Procedure Zlhis_Cis_036
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    发送号_In   In Number,
    医嘱id_In   In Number,
    No_In       In Varchar2,
    病人来源_In In Number
  );

  --撤消患者检查申请
  Procedure Zlhis_Cis_037
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    发送号_In   In Number,
    医嘱id_In   In Number,
    No_In       In Varchar2,
    病人来源_In In Number
  );

  --撤消患者手术申请
  Procedure Zlhis_Cis_038
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    发送号_In In Number,
    医嘱id_In In Number,
    No_In     In Varchar2
  );

  --撤消患者输血申请
  Procedure Zlhis_Cis_039
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    发送号_In In Number,
    医嘱id_In In Number,
    No_In     In Varchar2
  );

  --撤消患者会诊申请
  Procedure Zlhis_Cis_040
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number,
    No_In     In Varchar2
  );

  --撤消患者抢救医嘱
  Procedure Zlhis_Cis_041
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number,
    No_In     In Varchar2
  );

  --撤消患者死亡医嘱
  Procedure Zlhis_Cis_042
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number,
    No_In     In Varchar2
  );

  --撤消特殊治疗医嘱
  Procedure Zlhis_Cis_043
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number,
    No_In     In Varchar2
  );

  --撤消病区执行医嘱
  Procedure Zlhis_Cis_044
  (
    病人id_In   In Number,
    主页id_In   In Number,
    发送号_In   In Number,
    医嘱id_In   In Number,
    No_In       In Varchar2,
    发送数次_In In Number,
    首次时间_In In Date,
    末次时间_In In Date,
    样本条码_In In Varchar2
  );
  --患者医嘱执行登记
  Procedure Zlhis_Cis_050
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    发送号_In   In Number,
    医嘱id_In   In Number,
    要求时间_In In Date,
    执行时间_In In Date
  );

  --患者医嘱取消执行登记
  Procedure Zlhis_Cis_051
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    发送号_In   In Number,
    医嘱id_In   In Number,
    要求时间_In In Date,
    执行时间_In In Date,
    本次数次_In In Number,
    执行结果_In In Number,
    执行摘要_In In Varchar2,
    执行科室_In In Number,
    执行人_In   In Varchar2,
    核对人_In   In Varchar2,
    记录来源_In In Number
  );
  --患者医嘱执行完成
  Procedure Zlhis_Cis_052
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    发送号_In In Number,
    医嘱id_In In Number
  );
  --患者医嘱撤消执行完成
  Procedure Zlhis_Cis_053
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    发送号_In In Number,
    医嘱id_In In Number
  );
  --病理申请发送后修改
  Procedure Zlhis_Cis_056
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    发送号_In   In Number,
    医嘱id_In   In Number,
    病人来源_In In Number
  );

  --门诊患者完成就诊
  Procedure Zlhis_Cis_057
  (
    病人id_In In Number,
    挂号单_In In Varchar2
  );

  --门诊患者取消完成就诊
  Procedure Zlhis_Cis_058
  (
    病人id_In In Number,
    挂号单_In In Varchar2
  );

  --确认停止患者医嘱
  Procedure Zlhis_Cis_059
  (
    病人id_In In Number,
    主页id_In In Number,
    医嘱id_In In Number
  );

  --病人危急值处理
  Procedure Zlhis_Cis_060
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    危急值id_In In Number,
    医嘱id_In   In Number,
    病人来源_In In Number
  );

  --病人皮试结果填写
  Procedure Zlhis_Cis_061
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    医嘱id_In In Number
  );

  --26.检查报告完成，检查完成时
  Procedure Zlhis_Pacs_001
  (
    医嘱id_In   In Number,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  );
  --27.检查状态同步，检查状态改变后
  Procedure Zlhis_Pacs_002
  (
    医嘱id_In In Number,
    原状态_In In Number,
    新状态_In In Number
  );
  --28.检查状态回退，检查状态回退后
  Procedure Zlhis_Pacs_003
  (
    医嘱id_In In Number,
    原状态_In In Number,
    新状态_In In Number
  );
  --29.检查报告撤销，撤销检查完成时
  Procedure Zlhis_Pacs_004
  (
    医嘱id_In   In Number,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  );
  --30.检查危急值通知，检查发生危急值时
  Procedure Zlhis_Pacs_005(医嘱id_In In Number);
  -- 检查预约通知，检查预约时
  Procedure Zlhis_Pacs_006
  (
    医嘱id_In In Number,
    预约id_In In Number
  );
  -- 取消检查预约，取消预约时
  Procedure Zlhis_Pacs_007
  (
    医嘱id_In       In Number,
    预约id_In       In Number,
    预约日期_In     In Date,
    预约序号_In     In Number,
    检查设备名称_In In Varchar2
  );

  --36.患者发卡
  Procedure Zlhis_Patient_018
  (
    变动id_In   In Number,
    病人id_In   In Number,
    卡类别id_In In Number,
    卡号_In     In Varchar2
  );

  --37.患者退卡
  Procedure Zlhis_Patient_019
  (
    变动id_In   In Number,
    病人id_In   In Number,
    卡类别id_In In Number,
    卡号_In     In Varchar2
  );

  --38.患者退卡
  Procedure Zlhis_Patient_020
  (
    变动id_In   In Number,
    病人id_In   In Number,
    卡类别id_In In Number,
    原卡号_In   In Varchar2,
    新卡号_In   In Varchar2
  );

  --39.病人挂号登记（包含预约登记)
  Procedure Zlhis_Regist_001
  (
    挂号id_In In Number,
    No_In     In Varchar2
  );

  --40.病人分诊
  Procedure Zlhis_Regist_002
  (
    挂号id_In In Number,
    No_In     In Varchar2,
    诊室_In   In Varchar2
  );

  --41.病人退号
  Procedure Zlhis_Regist_003
  (
    挂号id_In In Number,
    No_In     In Varchar2
  );

  --42.临床出诊安排调整
  Procedure Zlhis_Regist_004
  (
    变动原因_In In Integer, --1-停诊;2-替诊;3-诊室变动
    记录id_In   In Number,
    变动id_In   In Number
  );

  --43.门诊患者挂号换号操作
  Procedure Zlhis_Regist_005
  (
    No_In         In Varchar2,
    变动原因_In   In Integer, --1-替诊;2-换号;3-预约日期变动,
    就诊变动id_In In Number
  );

  --费用门诊收费及补充结算
  --结算类型_In:1-收费结算，2-补充结算
  Procedure Zlhis_Charge_002
  (
    结算类型_In In Number,
    结帐id_In   In Number
  );

  --46.门诊退费单据
  Procedure Zlhis_Charge_004
  (
    退费类型_In In Number,
    结帐id_In   In Number
  );

  --47.收预交款
  Procedure Zlhis_Charge_005
  (
    预交id_In In Number,
    单据号_In In Varchar2
  );

  --48.退预交款(包含负数退预交款部分)
  Procedure Zlhis_Charge_006
  (
    退预交id_In In Number,
    单据号_In   In Varchar2
  );

  --住院记帐单据
  Procedure Zlhis_Charge_007
  (
    收费类别_In In Varchar2,
    费用id_In   In Number
  );

  --住院记帐单据销账
  Procedure Zlhis_Charge_008
  (
    收费类别_In In Varchar2,
    费用id_In   In Number,
    收发ids_In  In Varchar2 := Null --可能费用ID对应多个收发id，对应格式：收发id,数量|收发id,数量；非药品不传
  );

  --费用销账申请，
  Procedure Zlhis_Charge_009
  (
    费用id_In   Number,
    申请类别_In Number,
    申请时间_In Date
  );

  --取消销账申请
  Procedure Zlhis_Charge_010
  (
    费用id_In     Number,
    申请类别_In   Number,
    申请时间_In   Date,
    数量_In       Number,
    申请部门id_In Number,
    申请人_In     Varchar2
  );

  --53.住院患者入院登记
  Procedure Zlhis_Patient_001
  (
    病人id_In In Number,
    主页id_In In Number
  );
  --54.住院患者入院入科
  Procedure Zlhis_Patient_002
  (
    病人id_In In Number,
    主页id_In In Number
  );
  --56.住院患者床位变更
  Procedure Zlhis_Patient_004
  (
    病人id_In In Number,
    主页id_In In Number
  );
  --57.住院患者病情变更
  Procedure Zlhis_Patient_005
  (
    病人id_In In Number,
    主页id_In In Number
  );
  --58.住院患者变更撤消
  Procedure Zlhis_Patient_006
  (
    病人id_In   In Number,
    主页id_In   In Number,
    撤销方式_In In Varchar2
  );
  --59.住院患者医护变更
  Procedure Zlhis_Patient_007
  (
    病人id_In In Number,
    主页id_In In Number
  );
  --住院患者护理等级变更
  Procedure Zlhis_Patient_008
  (
    病人id_In In Number,
    主页id_In In Number
  );
  --60.住院患者预出院
  Procedure Zlhis_Patient_009
  (
    病人id_In In Number,
    主页id_In In Number
  );
  --61.住院患者出院
  Procedure Zlhis_Patient_010
  (
    病人id_In In Number,
    主页id_In In Number
  );
  --62.住院患者新生儿登记
  Procedure Zlhis_Patient_011
  (
    病人id_In   In Number,
    主页id_In   In Number,
    婴儿序号_In Number
  );
  --63.住院患者转入科室
  Procedure Zlhis_Patient_012
  (
    病人id_In In Number,
    主页id_In In Number
  );
  --64.新生儿登记作废
  Procedure Zlhis_Patient_013
  (
    病人id_In   In Number,
    主页id_In   In Number,
    婴儿序号_In Number
  );
  --65.门诊患者登记
  Procedure Zlhis_Patient_015(病人id_In In Number);
  --66.患者信息修改
  Procedure Zlhis_Patient_016(病人id_In In Number);

  --67.患者合并
  Procedure Zlhis_Patient_017
  (
    病人id_In   In Number,
    原病人id_In In Number,
    变化ids_In  In Varchar2
  );

  --69.患者转病区转入
  Procedure Zlhis_Patient_026
  (
    病人id_In In Number,
    主页id_In In Number
  );

  Procedure Zlhis_Patient_028(病人id_In In Number);

  --79.留观病人转住院病人
  Procedure Zlhis_Patient_029
  (
    病人id_In In Number,
    主页id_In In Number
  );

  --血库:科室配血完成
  Procedure Zlhis_Blood_001(医嘱id_In In Number);
  --血库:科室配血拒绝
  Procedure Zlhis_Blood_002(医嘱id_In In Number);

  --70.检验标本审核
  Procedure Zlhis_Lis_001(标本id_In In Number);
  --71.检验标本审核撤消
  Procedure Zlhis_Lis_002(标本id_In In Number);
  --73.检验标本条码打印
  Procedure Zlhis_Lis_004
  (
    样本条码_In In Varchar2,
    医嘱id_In   In Number,
    医嘱ids_In  In Varchar2
  );
  --74.检验标本条码打印撤销
  Procedure Zlhis_Lis_005
  (
    样本条码_In In Varchar2,
    医嘱id_In   In Number,
    医嘱ids_In  In Varchar2
  );
  --75.检验标本核收
  Procedure Zlhis_Lis_006(标本id_In In Number);
  --76.检验标本核收撤销
  Procedure Zlhis_Lis_007(标本id_In In Number);
  --77.检验标本拒收
  Procedure Zlhis_Lis_008(标本id_In In Number);
  --病历保存
  Procedure Zlhis_Emr_018
  (
    病人id_In In Number,
    主页id_In In Number,
    文件id_In In Number
  );
  --管理工具上机人员变动消息
  Procedure Zltools_Users_001
  (
    用户名_In In Varchar2,
    人员id_In In Number
  );
  Procedure Zltools_Users_002
  (
    用户名_In In Varchar2,
    人员id_In In Number
  );
End b_Message;
/

Create Or Replace Package Body b_Message Is
  --独立消息包缓存，防止手工编译导致问题
  --1、该包的修改以及编译请尽量放在业务低谷时候处理
  --2、请勿在该包中增加全局变量，如：
  --Message_Creator Zlmsg_Todo.Creator%Type := Null;
  --如确实需要增加全局变量（或已经存在全局变量），请在PLSQL中编译后执行如下语句：
  --ALTER PACKAGE b_Message COMPILE SPECIFICATION

  Procedure p_Msg_Todo_Insert
  (
    Msg_Code_In  Zlmsg_Todo.Msg_Code%Type,
    Key_Value_In Zlmsg_Todo.Key_Value%Type
  ) Is
  Begin
    If b_Zlmsg_Cache.Is_Message_Using(Msg_Code_In) = 0 Then
      Return;
    End If;
    Insert Into Zlmsg_Todo
      (ID, Msg_Code, Key_Value, State, Create_Time, Creator)
    Values
      (Zlmsg_Todo_Id.Nextval, Upper(Msg_Code_In), Key_Value_In, 0, Sysdate, b_Zlmsg_Cache.Get_Message_Creator);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Msg_Todo_Insert;
  --设置当前会话为平台调用
  Procedure Set_Platform_Call(Platform_Call Number) Is
  Begin
    b_Zlmsg_Cache.Set_Platform_Call(Platform_Call);
  End Set_Platform_Call;
  --消息Zlhis_Dict_001
  Procedure Zlhis_Dict_001(Id_In Number) Is
    v_Define Xmltype;
    v_Value  Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_001', v_Value);
  End Zlhis_Dict_001;
  --修改部门
  Procedure Zlhis_Dict_002(部门id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><部门ID>' || 部门id_In || '</部门ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_002', v_Value);
  End Zlhis_Dict_002;
  --停用部门
  Procedure Zlhis_Dict_003(部门id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><部门ID>' || 部门id_In || '</部门ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_003', v_Value);
  End Zlhis_Dict_003;
  --启用部门
  Procedure Zlhis_Dict_004(部门id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><部门ID>' || 部门id_In || '</部门ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_004', v_Value);
  End Zlhis_Dict_004;
  --新增人员
  Procedure Zlhis_Dict_005(人员id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_005', v_Value);
  End Zlhis_Dict_005;
  --修改人员
  Procedure Zlhis_Dict_006(人员id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_006', v_Value);
  End Zlhis_Dict_006;
  --停用人员
  Procedure Zlhis_Dict_007(人员id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_007', v_Value);
  End Zlhis_Dict_007;
  --启用人员
  Procedure Zlhis_Dict_008(人员id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_008', v_Value);
  End Zlhis_Dict_008;
  --新增收费项目
  Procedure Zlhis_Dict_009(细目id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_009', v_Value);
  End Zlhis_Dict_009;
  --修改收费项目
  Procedure Zlhis_Dict_010(细目id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_010', v_Value);
  End Zlhis_Dict_010;
  --停用收费项目
  Procedure Zlhis_Dict_011(细目id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_011', v_Value);
  End Zlhis_Dict_011;
  --启用收费项目
  Procedure Zlhis_Dict_012(细目id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_012', v_Value);
  End Zlhis_Dict_012;
  --新增诊疗项目
  Procedure Zlhis_Dict_013(诊疗id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_013', v_Value);
  End Zlhis_Dict_013;
  --修改诊疗项目
  Procedure Zlhis_Dict_014(诊疗id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_014', v_Value);
  End Zlhis_Dict_014;
  --停用诊疗项目
  Procedure Zlhis_Dict_015(诊疗id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_015', v_Value);
  End Zlhis_Dict_015;
  --启用诊疗项目
  Procedure Zlhis_Dict_016(诊疗id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_016', v_Value);
  End Zlhis_Dict_016;
  --新增检验项目
  Procedure Zlhis_Dict_017(诊疗id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID><系统>1</系统></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_017', v_Value);
  End Zlhis_Dict_017;
  --修改检验项目
  Procedure Zlhis_Dict_018(诊疗id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID><系统>1</系统></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_018', v_Value);
  End Zlhis_Dict_018;
  --删除检验项目
  Procedure Zlhis_Dict_019
  (
    诊疗id_In Number,
    编码_In   Varchar2,
    中文名_In Varchar2,
    英文名_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID>' || '<编码>' || 编码_In || '</编码>' || '<中文名>' || 中文名_In || '</中文名>' ||
               '<英文名>' || 英文名_In || '</英文名>' || '<系统>1</系统></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_019', v_Value);
  End Zlhis_Dict_019;
  --新增疾病编码目录
  Procedure Zlhis_Dict_021(疾病id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_021', v_Value);
  End Zlhis_Dict_021;
  --修改疾病编码目录
  Procedure Zlhis_Dict_022(疾病id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_022', v_Value);
  End Zlhis_Dict_022;
  --停用疾病编码目录
  Procedure Zlhis_Dict_023(疾病id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_023', v_Value);
  End Zlhis_Dict_023;
  --启用疾病编码目录
  Procedure Zlhis_Dict_024(疾病id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_024', v_Value);
  End Zlhis_Dict_024;
  --新增药品分类
  Procedure Zlhis_Dict_025
  (
    类型_In Number,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_025', v_Value);
  End Zlhis_Dict_025;
  --修改药品分类
  Procedure Zlhis_Dict_026
  (
    类型_In Number,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_026', v_Value);
  End Zlhis_Dict_026;
  --删除药品分类
  Procedure Zlhis_Dict_027
  (
    类型_In Number,
    Id_In   Number,
    编码_In Varchar2,
    名称_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID><编码>' || 编码_In || '</编码><名称>' || 名称_In ||
               '</名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_027', v_Value);
  End Zlhis_Dict_027;
  --停用药品分类
  Procedure Zlhis_Dict_028
  (
    类型_In Number,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_028', v_Value);
  End Zlhis_Dict_028;
  --启用药品分类
  Procedure Zlhis_Dict_029
  (
    类型_In Number,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_029', v_Value);
  End Zlhis_Dict_029;
  --新增药品品种
  Procedure Zlhis_Dict_030
  (
    类别_In Varchar2,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_030', v_Value);
  End Zlhis_Dict_030;
  --修改药品品种
  Procedure Zlhis_Dict_031
  (
    类别_In Varchar2,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_031', v_Value);
  End Zlhis_Dict_031;
  --删除药品品种
  Procedure Zlhis_Dict_032
  (
    类别_In Varchar2,
    Id_In   Number,
    编码_In Varchar2,
    名称_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID><编码>' || 编码_In || '</编码><名称>' || 名称_In ||
               '</名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_032', v_Value);
  End Zlhis_Dict_032;
  --停用药品品种
  Procedure Zlhis_Dict_033
  (
    类别_In Varchar2,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_033', v_Value);
  End Zlhis_Dict_033;
  --启用药品品种
  Procedure Zlhis_Dict_034
  (
    类别_In Varchar2,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_034', v_Value);
  End Zlhis_Dict_034;
  --新增药品规格
  Procedure Zlhis_Dict_035
  (
    类别_In   Varchar2,
    药品id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_035', v_Value);
  End Zlhis_Dict_035;
  --修改药品规格
  Procedure Zlhis_Dict_036
  (
    类别_In   Varchar2,
    药品id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_036', v_Value);
  End Zlhis_Dict_036;
  --删除药品规格
  Procedure Zlhis_Dict_037
  (
    类别_In   Varchar2,
    药品id_In Number,
    编码_In   Varchar2,
    名称_In   Varchar2,
    规格_In   Varchar2,
    产地_In   Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID><编码>' || 编码_In || '</编码><名称>' || 名称_In ||
               '</名称><规格>' || 规格_In || '</规格><产地>' || 产地_In || '</产地></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_037', v_Value);
  End Zlhis_Dict_037;
  --停用药品规格
  Procedure Zlhis_Dict_038
  (
    类别_In   Varchar2,
    药品id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_038', v_Value);
  End Zlhis_Dict_038;
  --启用药品规格
  Procedure Zlhis_Dict_039
  (
    类别_In   Varchar2,
    药品id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_039', v_Value);
  End Zlhis_Dict_039;
  --设置药品存储库房
  Procedure Zlhis_Dict_040
  (
    类别_In   Varchar2,
    药品id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_040', v_Value);
  End Zlhis_Dict_040;
  --设置药品储备限额
  Procedure Zlhis_Dict_041
  (
    类别_In   Varchar2,
    药品id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_041', v_Value);
  End Zlhis_Dict_041;
  --新增卫材品种
  Procedure Zlhis_Dict_042(Id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_042', v_Value);
  End Zlhis_Dict_042;
  --新增卫材规格
  Procedure Zlhis_Dict_043(Id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_043', v_Value);
  End Zlhis_Dict_043;
  --修改卫材规格
  Procedure Zlhis_Dict_044(Id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_044', v_Value);
  End Zlhis_Dict_044;
  --删除卫材规格
  Procedure Zlhis_Dict_045
  (
    Id_In   Number,
    编码_In Varchar2,
    名称_In Varchar2,
    规格_In Varchar2,
    产地_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><规格>' || 规格_In ||
               '</规格><产地>' || 产地_In || '</产地></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_045', v_Value);
  End Zlhis_Dict_045;
  --停用卫材规格
  Procedure Zlhis_Dict_046(Id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_046', v_Value);
  End Zlhis_Dict_046;
  --启用卫材规格
  Procedure Zlhis_Dict_047(Id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_047', v_Value);
  End Zlhis_Dict_047;
  --医保对码
  Procedure Zlhis_Dict_048
  (
    险类_In       In Number,
    收费细目id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><险类>' || 险类_In || '</险类><收费细目ID>' || 收费细目id_In || '</收费细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_048', v_Value);
  End Zlhis_Dict_048;
  --删除医保对码
  Procedure Zlhis_Dict_049
  (
    险类_In       In Number,
    收费细目id_In In Number,
    项目编码_In   In Varchar2,
    项目名称_In   In Varchar2,
    医保编码_In   In Varchar2,
    医保名称_In   In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><险类>' || 险类_In || '</险类><收费细目ID>' || 收费细目id_In || '</收费细目ID><项目编码>' || 项目编码_In || '</项目编码><项目名称>' ||
               项目名称_In || '</项目名称><医保编码>' || 医保编码_In || '</医保编码><医保名称>' || 医保名称_In || '</医保名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_049', v_Value);
  End Zlhis_Dict_049;
  --新增卫材分类
  Procedure Zlhis_Dict_050
  (
    类型_In Number,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_050', v_Value);
  End Zlhis_Dict_050;
  --修改卫材分类
  Procedure Zlhis_Dict_051
  (
    类型_In Number,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_051', v_Value);
  End Zlhis_Dict_051;
  --删除卫材分类
  Procedure Zlhis_Dict_052
  (
    类型_In Number,
    Id_In   Number,
    编码_In Varchar2,
    名称_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID><编码>' || 编码_In || '</编码><名称>' || 名称_In ||
               '</名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_052', v_Value);
  End Zlhis_Dict_052;
  --收费价目变动
  Procedure Zlhis_Dict_053(收费项目id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><收费项目ID>' || 收费项目id_In || '</收费项目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_053', v_Value);
  End Zlhis_Dict_053;

  --诊疗收费对照变动
  Procedure Zlhis_Dict_054(诊疗项目id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗项目ID>' || 诊疗项目id_In || '</诊疗项目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_054', v_Value);
  End Zlhis_Dict_054;

  --卫材存储库房变动
  Procedure Zlhis_Dict_055(细目id_In Varchar2) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><收费细目ID>' || 细目id_In || '</收费细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_055', v_Value);
  End Zlhis_Dict_055;

  --新增诊疗检查类型
  Procedure Zlhis_Dictpacs_001
  (
    编码_In   Varchar2,
    名称_In   Varchar2,
    简码_In   Varchar2,
    建病案_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><建病案>' || 建病案_In ||
               '</建病案></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_001', v_Value);
  End Zlhis_Dictpacs_001;

  --修改诊疗检查类型
  Procedure Zlhis_Dictpacs_002
  (
    编码_In   Varchar2,
    名称_In   Varchar2,
    简码_In   Varchar2,
    建病案_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><建病案>' || 建病案_In ||
               '</建病案></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_002', v_Value);
  End Zlhis_Dictpacs_002;
  --删除诊疗检查类型
  Procedure Zlhis_Dictpacs_003
  (
    编码_In   Varchar2,
    名称_In   Varchar2,
    简码_In   Varchar2,
    建病案_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><建病案>' || 建病案_In ||
               '</建病案></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_003', v_Value);
  End Zlhis_Dictpacs_003;
  --新增诊疗检查部位
  Procedure Zlhis_Dictpacs_004
  (
    类型_In     Varchar2,
    编码_In     Varchar2,
    名称_In     Varchar2,
    分组_In     Varchar2,
    备注_In     Varchar2,
    方法_In     Varchar2,
    适用性别_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><分组>' || 分组_In ||
               '</分组><备注>' || 备注_In || '</备注><方法>' || 方法_In || '</方法><适用性别>' || 适用性别_In || '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_004', v_Value);
  End Zlhis_Dictpacs_004;
  --修改诊疗检查部位
  Procedure Zlhis_Dictpacs_005
  (
    类型_In     Varchar2,
    编码_In     Varchar2,
    名称_In     Varchar2,
    分组_In     Varchar2,
    备注_In     Varchar2,
    方法_In     Varchar2,
    适用性别_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><分组>' || 分组_In ||
               '</分组><备注>' || 备注_In || '</备注><方法>' || 方法_In || '</方法><适用性别>' || 适用性别_In || '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_005', v_Value);
  End Zlhis_Dictpacs_005;
  --删除诊疗检查部位
  Procedure Zlhis_Dictpacs_006
  (
    类型_In     Varchar2,
    编码_In     Varchar2,
    名称_In     Varchar2,
    分组_In     Varchar2,
    备注_In     Varchar2,
    方法_In     Varchar2,
    适用性别_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><分组>' || 分组_In ||
               '</分组><备注>' || 备注_In || '</备注><方法>' || 方法_In || '</方法><适用性别>' || 适用性别_In || '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_006', v_Value);
  End Zlhis_Dictpacs_006;
  --新增诊疗项目部位
  Procedure Zlhis_Dictpacs_007
  (
    Id_In       Number,
    项目id_In   Number,
    类型_In     Varchar2,
    部位_In     Varchar2,
    方法_In     Varchar2,
    默认_In     Number,
    上级方法_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><项目ID>' || 项目id_In || '</项目ID><类型>' || 类型_In || '</类型><部位>' || 部位_In ||
               '</部位><方法>' || 方法_In || '</方法><默认>' || 默认_In || '</默认><上级方法>' || 上级方法_In || '</上级方法></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_007', v_Value);
  End Zlhis_Dictpacs_007;
  --修改诊疗项目部位
  Procedure Zlhis_Dictpacs_008
  (
    Id_In       Number,
    项目id_In   Number,
    类型_In     Varchar2,
    部位_In     Varchar2,
    方法_In     Varchar2,
    默认_In     Number,
    上级方法_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><项目ID>' || 项目id_In || '</项目ID><类型>' || 类型_In || '</类型><部位>' || 部位_In ||
               '</部位><方法>' || 方法_In || '</方法><默认>' || 默认_In || '</默认><上级方法>' || 上级方法_In || '</上级方法></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_008', v_Value);
  End Zlhis_Dictpacs_008;
  --删除诊疗项目部位
  Procedure Zlhis_Dictpacs_009
  (
    Id_In       Number,
    项目id_In   Number,
    类型_In     Varchar2,
    部位_In     Varchar2,
    方法_In     Varchar2,
    默认_In     Number,
    上级方法_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><项目ID>' || 项目id_In || '</项目ID><类型>' || 类型_In || '</类型><部位>' || 部位_In ||
               '</部位><方法>' || 方法_In || '</方法><默认>' || 默认_In || '</默认><上级方法>' || 上级方法_In || '</上级方法></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_009', v_Value);
  End Zlhis_Dictpacs_009;
  --新增诊疗项目部位
  Procedure Zlhis_Dictlis_004
  (
    编码_In     Varchar2,
    名称_In     Varchar2,
    简码_In     Varchar2,
    适用性别_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><适用性别>' || 适用性别_In ||
               '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_004', v_Value);
  End Zlhis_Dictlis_004;
  --修改诊疗项目部位
  Procedure Zlhis_Dictlis_005
  (
    编码_In     Varchar2,
    名称_In     Varchar2,
    简码_In     Varchar2,
    适用性别_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><适用性别>' || 适用性别_In ||
               '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_005', v_Value);
  End Zlhis_Dictlis_005;
  --删除诊疗项目部位
  Procedure Zlhis_Dictlis_006
  (
    编码_In     Varchar2,
    名称_In     Varchar2,
    简码_In     Varchar2,
    适用性别_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><适用性别>' || 适用性别_In ||
               '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_006', v_Value);
  End Zlhis_Dictlis_006;
  --新增采血管类型
  Procedure Zlhis_Dictlis_007
  (
    编码_In   Varchar2,
    名称_In   Varchar2,
    简码_In   Varchar2,
    添加剂_In Varchar2,
    采血量_In Varchar2,
    规格_In   Varchar2,
    颜色_In   Number,
    材料id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><添加剂>' || 添加剂_In ||
               '</添加剂><采血量>' || 采血量_In || '</采血量><颜色>' || 颜色_In || '</颜色><规格>' || 规格_In || '</规格><材料ID_In>' || 材料id_In ||
               '</材料ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_007', v_Value);
  End Zlhis_Dictlis_007;
  --新增采血管类型
  Procedure Zlhis_Dictlis_008
  (
    编码_In   Varchar2,
    名称_In   Varchar2,
    简码_In   Varchar2,
    添加剂_In Varchar2,
    采血量_In Varchar2,
    规格_In   Varchar2,
    颜色_In   Number,
    材料id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><添加剂>' || 添加剂_In ||
               '</添加剂><采血量>' || 采血量_In || '</采血量><颜色>' || 颜色_In || '</颜色><规格>' || 规格_In || '</规格><材料ID_In>' || 材料id_In ||
               '</材料ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_008', v_Value);
  End Zlhis_Dictlis_008;
  --新增采血管类型
  Procedure Zlhis_Dictlis_009
  (
    编码_In   Varchar2,
    名称_In   Varchar2,
    简码_In   Varchar2,
    添加剂_In Varchar2,
    采血量_In Varchar2,
    规格_In   Varchar2,
    颜色_In   Number,
    材料id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><添加剂>' || 添加剂_In ||
               '</添加剂><采血量>' || 采血量_In || '</采血量><颜色>' || 颜色_In || '</颜色><规格>' || 规格_In || '</规格><材料ID_In>' || 材料id_In ||
               '</材料ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_009', v_Value);
  End Zlhis_Dictlis_009;
  --药品备药发送
  Procedure Zlhis_Drug_001(No_In Varchar2) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_001', v_Value);
  End Zlhis_Drug_001;
  --取消药品备药发送
  Procedure Zlhis_Drug_002(No_In Varchar2) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_002', v_Value);
  End Zlhis_Drug_002;
  --药品移库单接收
  Procedure Zlhis_Drug_003(No_In Varchar2) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_003', v_Value);
  End Zlhis_Drug_003;
  --药品移库单冲销
  Procedure Zlhis_Drug_004
  (
    No_In       Varchar2,
    序号_In     Number,
    记录状态_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号><序号>' || 序号_In || '</序号><记录状态>' || 记录状态_In || '</记录状态></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_004', v_Value);
  End Zlhis_Drug_004;
  --部门发药
  Procedure Zlhis_Drug_005
  (
    库房id_In Number,
    收发id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><库房ID>' || 库房id_In || '</库房ID><收发ID>' || 收发id_In || '</收发ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_005', v_Value);
  End Zlhis_Drug_005;
  --部门退药
  Procedure Zlhis_Drug_006
  (
    冲销收发id_In Number,
    待发收发id_In Number,
    数量_In       Number,
    费用id_In     Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><冲销记录ID>' || 冲销收发id_In || '</冲销记录ID><待发记录ID>' || 待发收发id_In || '</待发记录ID><数量>' || 数量_In ||
               '</数量><费用ID>' || 费用id_In || '</费用ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_006', v_Value);
  End Zlhis_Drug_006;
  --药品调价
  Procedure Zlhis_Drug_007(价格id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格id_In || '</价格ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_007', v_Value);
  End Zlhis_Drug_007;
  --静配发送
  Procedure Zlhis_Drug_008(记录ids_In Varchar2) Is
    v_Value  Zlmsg_Todo.Key_Value%Type;
    n_记录id Number(18);
    v_Tmp    Varchar2(4000);
    n_Length Number(18);
  Begin
    If 记录ids_In Is Null Then
      v_Tmp := Null;
    Else
      v_Tmp := 记录ids_In || ',';
    End If;
  
    v_Value := '<root><记录IDS>';
  
    While v_Tmp Is Not Null Loop
      --分解单据ID串
      n_记录id := To_Number(Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1));
      v_Tmp    := Replace(',' || v_Tmp, ',' || n_记录id || ',');
    
      --判断当前长度是否即将超过缓存
      Select Lengthb(v_Value || '<记录ID>' || n_记录id || '</记录ID>') Into n_Length From Dual;
      If n_Length > 950 Then
        v_Value := v_Value || '</记录IDs></root>';
        b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);
        v_Value := '<root><记录IDs>';
      End If;
    
      v_Value := v_Value || '<记录ID>' || n_记录id || '</记录ID>';
    End Loop;
  
    v_Value := v_Value || '</记录IDS></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);
  End Zlhis_Drug_008;
  --药品调售价
  Procedure Zlhis_Drug_009
  (
    价格id_In Number,
    时价_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格id_In || '</价格ID><时价>' || 时价_In || '</时价></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_009', v_Value);
  End Zlhis_Drug_009;
  --卫材调成本价
  Procedure Zlhis_Drug_010(价格id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格id_In || '</价格ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_010', v_Value);
  End Zlhis_Drug_010;
  --卫材调售价
  Procedure Zlhis_Drug_011
  (
    价格id_In Number,
    时价_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格id_In || '</价格ID><时价>' || 时价_In || '</时价></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_011', v_Value);
  End Zlhis_Drug_011;
  --卫材发料
  Procedure Zlhis_Drug_012
  (
    库房id_In Number,
    收发id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><库房ID>' || 库房id_In || '</库房ID><收发ID>' || 收发id_In || '</收发ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_012', v_Value);
  End Zlhis_Drug_012;
  --卫材退料
  Procedure Zlhis_Drug_013
  (
    冲销收发id_In Number,
    待发收发id_In Number,
    数量_In       Number,
    费用id_In     Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><冲销记录ID>' || 冲销收发id_In || '</冲销记录ID><待发记录ID>' || 待发收发id_In || '</待发记录ID><数量>' || 数量_In ||
               '</数量><费用ID>' || 费用id_In || '</费用ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_013', v_Value);
  End Zlhis_Drug_013;
  --2.停止患者医嘱，住院
  Procedure Zlhis_Cis_002
  (
    病人id_In  In Number,
    主页id_In  In Number,
    医嘱id_In  In Number,
    医嘱ids_In In Varchar2
  ) Is
    r_Data c_Dynamic;
    v_Id   Number(18);
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_CIS_002') = 0 Then
      Return;
    End If;
    If 医嘱id_In Is Not Null Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002',
                                  '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                   '</ID></root>');
    Else
      Open r_Data For 'Select ID From 病人医嘱记录 Where ID In (Select Column_Value From Table(f_Num2list(:1))) And 相关id Is Null'
        Using 医嘱ids_In;
      Loop
        Fetch r_Data
          Into v_Id;
        Exit When r_Data%NotFound;
      
        b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002',
                                    '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || v_Id ||
                                     '</ID></root>');
      End Loop;
    End If;
  End Zlhis_Cis_002;

  --3.作废患者医嘱，门诊/住院
  Procedure Zlhis_Cis_003
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    医嘱id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_003', v_Value);
  End Zlhis_Cis_003;

  --4.患者术后医嘱，住院
  Procedure Zlhis_Cis_004
  (
    病人id_In In Number,
    主页id_In In Number,
    医嘱id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_004',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_004;

  --5.撤消患者术后医嘱，住院
  Procedure Zlhis_Cis_005
  (
    病人id_In In Number,
    主页id_In In Number,
    医嘱id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_005',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_005;

  --6.患者护理常规医嘱，住院
  Procedure Zlhis_Cis_006
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_006',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In ||
                                 '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_006;

  --7.撤消患者护理常规医嘱，住院
  Procedure Zlhis_Cis_007
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_007', v_Value);
  End Zlhis_Cis_007;

  --门诊患者接诊
  Procedure Zlhis_Cis_008
  (
    病人id_In In Number,
    挂号单_In In Varchar2
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_008', '<root><病人ID>' || 病人id_In || '</病人ID><NO>' || 挂号单_In || '</NO></root>');
  End Zlhis_Cis_008;

  --门诊患者取消接诊
  Procedure Zlhis_Cis_009
  (
    病人id_In In Number,
    挂号单_In In Varchar2
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_009', '<root><病人ID>' || 病人id_In || '</病人ID><NO>' || 挂号单_In || '</NO></root>');
  End Zlhis_Cis_009;

  --10.下达患者诊断，门诊/住院
  Procedure Zlhis_Cis_010
  (
    病人id_In In Number,
    就诊id_In In Number, --门诊病人 挂号ID，住院病人 主页ID
    诊断id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_010',
                                '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><ID>' || 诊断id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_010;
  --11.撤消患者诊断
  Procedure Zlhis_Cis_011
  (
    病人id_In   In Number,
    就诊id_In   In Number, --门诊病人 挂号ID，住院病人 主页ID
    Id_In       In Number,
    疾病id_In   In Number,
    诊断id_In   In Number,
    诊断描述_In In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><ID>' || Id_In || '</ID><疾病ID>' ||
               疾病id_In || '</疾病ID><诊断ID>' || 诊断id_In || '</诊断ID><诊断描述>' || 诊断描述_In || '</诊断描述></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_011', v_Value);
  End Zlhis_Cis_011;

  --病区执行医嘱校对
  Procedure Zlhis_Cis_012
  (
    病人id_In In Number,
    主页id_In In Number,
    医嘱id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_012',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_012;

  --13.检验危急值阅读通知
  Procedure Zlhis_Cis_014
  (
    病人id_In In Number,
    就诊id_In In Number, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In Number,
    消息id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><医嘱ID>' || 医嘱id_In || '</医嘱ID><ID>' ||
               消息id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_014', v_Value);
  End Zlhis_Cis_014;
  --15.患者检验申请，门诊/住院
  Procedure Zlhis_Cis_016
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    发送号_In   In Number,
    医嘱id_In   In Number,
    病人来源_In In Number --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><病人来源>' || 病人来源_In || '</病人来源></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_016', v_Value);
  End Zlhis_Cis_016;
  --16.患者检查申请，门诊/住院
  Procedure Zlhis_Cis_017
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    发送号_In   In Number,
    医嘱id_In   In Number,
    病人来源_In In Number --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_操作类型 Varchar2(20);
  Begin
    Execute Immediate 'Select Max(a.操作类型) From 诊疗项目目录 A, 病人医嘱记录 B Where b.诊疗项目id = a.Id And b.Id = :1'
      Into v_操作类型
      Using 医嘱id_In;
  
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><病人来源>' || 病人来源_In || '</病人来源></root>';
    If v_操作类型 = '病理' Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_054', v_Value);
    Else
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_017', v_Value);
    End If;
  End Zlhis_Cis_017;
  --17.患者手术申请，门诊/住院
  Procedure Zlhis_Cis_018
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    发送号_In In Number,
    医嘱id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_018', v_Value);
  End Zlhis_Cis_018;
  --18.患者输血申请，住院
  Procedure Zlhis_Cis_019
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    发送号_In In Number,
    医嘱id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_019', v_Value);
  End Zlhis_Cis_019;
  --19.患者会诊申请，住院
  Procedure Zlhis_Cis_020
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_020', v_Value);
  End Zlhis_Cis_020;
  --20.患者抢救医嘱，住院
  Procedure Zlhis_Cis_021
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_021', v_Value);
  End Zlhis_Cis_021;
  --21.患者死亡医嘱，住院
  Procedure Zlhis_Cis_022
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_022', v_Value);
  End Zlhis_Cis_022;
  --22.患者特殊治疗医嘱，住院
  Procedure Zlhis_Cis_023
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_023', v_Value);
  End Zlhis_Cis_023;

  --24.检查危急值阅读通知
  Procedure Zlhis_Cis_025
  (
    病人id_In In Number,
    就诊id_In In Number, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In Number,
    消息id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><医嘱ID>' || 医嘱id_In || '</医嘱ID><ID>' ||
               消息id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_025', v_Value);
  End Zlhis_Cis_025;

  --病区执行医嘱发送
  Procedure Zlhis_Cis_026
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_026',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In ||
                                 '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_026;

  --撤消患者检验申请
  Procedure Zlhis_Cis_036
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    发送号_In   In Number,
    医嘱id_In   In Number,
    No_In       In Varchar2,
    病人来源_In In Number --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO><病人来源>' || 病人来源_In ||
               '</病人来源></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_036', v_Value);
  End Zlhis_Cis_036;

  --撤消患者检查申请
  Procedure Zlhis_Cis_037
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    发送号_In   In Number,
    医嘱id_In   In Number,
    No_In       In Varchar2,
    病人来源_In In Number --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_操作类型 Varchar2(20);
  Begin
    Execute Immediate 'Select Max(a.操作类型) From 诊疗项目目录 A, 病人医嘱记录 B Where b.诊疗项目id = a.Id And b.Id = :1'
      Into v_操作类型
      Using 医嘱id_In;
  
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO><病人来源>' || 病人来源_In ||
               '</病人来源></root>';
    If v_操作类型 = '病理' Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_055', v_Value);
    Else
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_037', v_Value);
    End If;
  End Zlhis_Cis_037;

  --撤消患者手术申请
  Procedure Zlhis_Cis_038
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    发送号_In In Number,
    医嘱id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_038', v_Value);
  End Zlhis_Cis_038;

  --撤消患者输血申请
  Procedure Zlhis_Cis_039
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    发送号_In In Number,
    医嘱id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_039', v_Value);
  End Zlhis_Cis_039;

  --撤消患者会诊申请
  Procedure Zlhis_Cis_040
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_040', v_Value);
  End Zlhis_Cis_040;

  --撤消患者抢救医嘱
  Procedure Zlhis_Cis_041
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_041', v_Value);
  End Zlhis_Cis_041;

  --撤消患者死亡医嘱
  Procedure Zlhis_Cis_042
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_042', v_Value);
  End Zlhis_Cis_042;

  --撤消特殊治疗医嘱
  Procedure Zlhis_Cis_043
  (
    病人id_In In Number,
    主页id_In In Number,
    发送号_In In Number,
    医嘱id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_043', v_Value);
  End Zlhis_Cis_043;

  --撤消病区执行医嘱
  Procedure Zlhis_Cis_044
  (
    病人id_In   In Number,
    主页id_In   In Number,
    发送号_In   In Number,
    医嘱id_In   In Number,
    No_In       In Varchar2,
    发送数次_In In Number,
    首次时间_In In Date,
    末次时间_In In Date,
    样本条码_In In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO><发送数次>' || 发送数次_In || '</发送数次><首次时间>' ||
               To_Char(首次时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</首次时间><末次时间>' ||
               To_Char(末次时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</末次时间><样本条码>' || 样本条码_In || '</样本条码></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_044', v_Value);
  End Zlhis_Cis_044;

  --患者医嘱执行登记
  Procedure Zlhis_Cis_050
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    发送号_In   In Number,
    医嘱id_In   In Number,
    要求时间_In In Date,
    执行时间_In In Date
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><要求时间>' || To_Char(要求时间_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</要求时间><执行时间>' || To_Char(执行时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</执行时间></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_050', v_Value);
  End Zlhis_Cis_050;

  --患者医嘱取消执行登记
  Procedure Zlhis_Cis_051
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    发送号_In   In Number,
    医嘱id_In   In Number,
    要求时间_In In Date,
    执行时间_In In Date,
    本次数次_In In Number,
    执行结果_In In Number,
    执行摘要_In In Varchar2,
    执行科室_In In Number,
    执行人_In   In Varchar2,
    核对人_In   In Varchar2,
    记录来源_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><要求时间>' || To_Char(要求时间_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</要求时间><执行时间>' || To_Char(执行时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</执行时间><本次数次>' || 本次数次_In ||
               '</本次数次><执行结果>' || 执行结果_In || '</执行结果><执行摘要>' || 执行摘要_In || '</执行摘要><执行科室ID>' || 执行科室_In ||
               '</执行科室ID><执行人>' || 执行人_In || '</执行人><核对人>' || 核对人_In || '</核对人><记录来源>' || 记录来源_In || '</记录来源></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_051', v_Value);
  End Zlhis_Cis_051;
  --患者医嘱执行完成
  Procedure Zlhis_Cis_052
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    发送号_In In Number,
    医嘱id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_052',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In ||
                                 '</挂号单><发送号>' || 发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_052;
  --患者医嘱撤消执行完成
  Procedure Zlhis_Cis_053
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    发送号_In In Number,
    医嘱id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_053',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In ||
                                 '</挂号单><发送号>' || 发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_053;

  --病理申请发送后修改
  Procedure Zlhis_Cis_056
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    发送号_In   In Number,
    医嘱id_In   In Number,
    病人来源_In In Number --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><病人来源>' || 病人来源_In || '</病人来源></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_056', v_Value);
  End Zlhis_Cis_056;

  --门诊患者完成就诊
  Procedure Zlhis_Cis_057
  (
    病人id_In In Number,
    挂号单_In In Varchar2
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_057', '<root><病人ID>' || 病人id_In || '</病人ID><NO>' || 挂号单_In || '</NO></root>');
  End Zlhis_Cis_057;

  --门诊患者取消完成就诊
  Procedure Zlhis_Cis_058
  (
    病人id_In In Number,
    挂号单_In In Varchar2
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_058', '<root><病人ID>' || 病人id_In || '</病人ID><NO>' || 挂号单_In || '</NO></root>');
  End Zlhis_Cis_058;

  --确认停止患者医嘱
  Procedure Zlhis_Cis_059
  (
    病人id_In In Number,
    主页id_In In Number,
    医嘱id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_059',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_059;

  --病人危急值处理
  Procedure Zlhis_Cis_060
  (
    病人id_In   In Number,
    主页id_In   In Number,
    挂号单_In   In Varchar2,
    危急值id_In In Number,
    医嘱id_In   In Number,
    病人来源_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_060',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In ||
                                 '</挂号单><ID>' || 危急值id_In || '</ID><医嘱ID>' || 医嘱id_In || '</医嘱ID><病人来源>' || 病人来源_In ||
                                 '</病人来源></root>');
  End Zlhis_Cis_060;

  --病人皮试结果填写
  Procedure Zlhis_Cis_061
  (
    病人id_In In Number,
    主页id_In In Number,
    挂号单_In In Varchar2,
    医嘱id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_061',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In ||
                                 '</挂号单><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_061;

  --26.检查报告完成，检查完成时
  Procedure Zlhis_Pacs_001
  (
    医嘱id_In   In Number,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  ) Is
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PACS_001') = 0 Then
      Return;
    End If;
    For R In (Select '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><报告ID>' || Column_Value || '</报告ID><报告类型>' || 报告类型_In ||
                      '</报告类型></root>' As Xml_Value
              From Table(f_Str2list(报告id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_001', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_001;
  --27.检查状态同步，检查状态改变后
  Procedure Zlhis_Pacs_002
  (
    医嘱id_In In Number,
    原状态_In In Number,
    新状态_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><原状态>' || 原状态_In || '</原状态><新状态>' || 新状态_In || '</新状态></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_002', v_Value);
  End Zlhis_Pacs_002;
  --28.检查状态回退，检查状态回退后
  Procedure Zlhis_Pacs_003
  (
    医嘱id_In In Number,
    原状态_In In Number,
    新状态_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><原状态>' || 原状态_In || '</原状态><新状态>' || 新状态_In || '</新状态></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_003', v_Value);
  End Zlhis_Pacs_003;
  --29.检查报告撤销，撤销检查完成时
  Procedure Zlhis_Pacs_004
  (
    医嘱id_In   In Number,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  ) Is
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PACS_004') = 0 Then
      Return;
    End If;
    For R In (Select '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><报告ID>' || Column_Value || '</报告ID><报告类型>' || 报告类型_In ||
                      '</报告类型></root>' As Xml_Value
              From Table(f_Str2list(报告id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_004', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_004;
  --30.检查危急值通知，检查发生危急值时
  Procedure Zlhis_Pacs_005(医嘱id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_005', v_Value);
  End Zlhis_Pacs_005;
  -- 检查预约通知，检查预约时
  Procedure Zlhis_Pacs_006
  (
    医嘱id_In In Number,
    预约id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><预约ID>' || 预约id_In || '</预约ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_006', v_Value);
  End Zlhis_Pacs_006;
  -- 取消检查预约，取消预约时
  Procedure Zlhis_Pacs_007
  (
    医嘱id_In       In Number,
    预约id_In       In Number,
    预约日期_In     In Date,
    预约序号_In     In Number,
    检查设备名称_In In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><预约ID>' || 预约id_In || '</预约ID><预约日期>' || 预约日期_In || '</预约日期><预约序号>' ||
               预约序号_In || '</预约序号><检查设备名称>' || 检查设备名称_In || '</检查设备名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_007', v_Value);
  End Zlhis_Pacs_007;

  --36.患者发卡或绑定卡
  Procedure Zlhis_Patient_018
  (
    变动id_In   In Number,
    病人id_In   In Number,
    卡类别id_In In Number,
    卡号_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动ID>' || 变动id_In || '</变动ID><病人ID>' || 病人id_In || '</病人ID><卡类别ID>' || 卡类别id_In ||
               '</卡类别ID><卡号>' || 卡号_In || '</卡号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_018', v_Value);
  End;

  --37.患者退卡
  Procedure Zlhis_Patient_019
  (
    变动id_In   In Number,
    病人id_In   In Number,
    卡类别id_In In Number,
    卡号_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动ID>' || 变动id_In || '</变动ID><病人ID>' || 病人id_In || '</病人ID><卡类别ID>' || 卡类别id_In ||
               '</卡类别ID><卡号>' || 卡号_In || '</卡号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_019', v_Value);
  End;

  --38.患者补卡/换卡
  Procedure Zlhis_Patient_020
  (
    变动id_In   In Number,
    病人id_In   In Number,
    卡类别id_In In Number,
    原卡号_In   In Varchar2,
    新卡号_In   In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动ID>' || 变动id_In || '</变动ID><病人ID>' || 病人id_In || '</病人ID><卡类别ID>' || 卡类别id_In ||
               '</卡类别ID><原卡号>' || 原卡号_In || '</原卡号><新卡号>' || 新卡号_In || '</新卡号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_020', v_Value);
  End;

  --39.病人挂号登记（包含预约登记)
  Procedure Zlhis_Regist_001
  (
    挂号id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><挂号ID>' || 挂号id_In || '</挂号ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_001', v_Value);
  End;

  --40.病人分诊
  Procedure Zlhis_Regist_002
  (
    挂号id_In In Number,
    No_In     In Varchar2,
    诊室_In   In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><挂号ID>' || 挂号id_In || '</挂号ID><NO>' || No_In || '</NO><诊室>' || Nvl(诊室_In, '') || '</诊室></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_002', v_Value);
  End;

  --41.病人退号（含取消预约)
  Procedure Zlhis_Regist_003
  (
    挂号id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><挂号ID>' || 挂号id_In || '</挂号ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_003', v_Value);
  End;

  --42.临床出诊安排调整
  Procedure Zlhis_Regist_004
  (
    变动原因_In In Integer, --1-停诊;2-替诊;3-诊室变动
    记录id_In   In Number,
    变动id_In   In Number
    
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动原因>' || 变动原因_In || '</变动原因><记录ID>' || 记录id_In || '</记录ID><变动ID>' || 变动id_In ||
               '</变动ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_004', v_Value);
  End;

  --43.门诊患者挂号换号操作
  Procedure Zlhis_Regist_005
  (
    No_In         In Varchar2,
    变动原因_In   Integer, --1-替诊;2-换号;3-预约日期变动,
    就诊变动id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><NO>' || No_In || '</NO><变动原因>' || 变动原因_In || '</变动原因><就诊变动ID>' || 就诊变动id_In ||
               '</就诊变动ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_005', v_Value);
  End;

  --费用门诊收费及补充结算
  Procedure Zlhis_Charge_002
  (
    结算类型_In In Number,
    结帐id_In   In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --结算类型_In:1-收费结算，2-补充结算
    v_Value := '<root><结算类型>' || 结算类型_In || '</结算类型><结帐ID>' || 结帐id_In || '</结帐ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_002', v_Value);
  End;

  --46.门诊退费单据
  Procedure Zlhis_Charge_004
  (
    退费类型_In In Number,
    结帐id_In   In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --退费类型_In:1-收费结算，2-补充结算
    v_Value := '<root><退费类型>' || 退费类型_In || '</退费类型><结帐ID>' || 结帐id_In || '</结帐ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_004', v_Value);
  End;

  --47.收预交款
  Procedure Zlhis_Charge_005
  (
    预交id_In In Number,
    单据号_In In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><预交ID>' || 预交id_In || '</预交ID><单据号>' || 单据号_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_005', v_Value);
  End;

  --48.退预交款(包含负数退预交款部分)
  Procedure Zlhis_Charge_006
  (
    退预交id_In In Number,
    单据号_In   In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><退预交ID>' || 退预交id_In || '</退预交ID><单据号>' || 单据号_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_006', v_Value);
  End;

  --住院记帐单据
  Procedure Zlhis_Charge_007
  (
    收费类别_In In Varchar2,
    费用id_In   In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><收费类别>' || 收费类别_In || '</收费类别><费用ID>' || 费用id_In || '</费用ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_007', v_Value);
  End;

  --住院记帐单据销账
  Procedure Zlhis_Charge_008
  (
    收费类别_In In Varchar2,
    费用id_In   In Number,
    收发ids_In  In Varchar2 := Null --可能费用ID对应多个收发id，对应格式：收发id,数量|收发id,数量；非药品不传
  ) Is
    v_Value   Zlmsg_Todo.Key_Value%Type;
    v_Tmp     Varchar2(4000);
    v_Infotmp Varchar2(4000);
    v_Fields  Varchar2(4000);
    v_收发id  Varchar2(50);
    v_数量    Varchar2(20);
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_CHARGE_008') = 0 Then
      Return;
    End If;
    v_Value := '<root><收费类别>' || 收费类别_In || '</收费类别><费用ID>' || 费用id_In || '</费用ID>';
  
    If 收发ids_In Is Null Then
      v_Infotmp := Null;
      v_Tmp     := '<收发IDS>' || '<收发ID>' || '</收发ID>' || '<数量>' || '</数量>' || '</收发IDS>';
    Else
      v_Infotmp := 收发ids_In || '|';
      While v_Infotmp Is Not Null Loop
        --分解收发ID串
        v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
        v_收发id  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
        v_数量    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
        v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
      
        v_Tmp := v_Tmp || '<收发IDS>' || '<收发ID>' || v_收发id || '</收发ID>' || '<数量>' || v_数量 || '</数量>' || '</收发IDS>';
      End Loop;
    End If;
  
    v_Value := v_Value || v_Tmp || '</root>';
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_008', v_Value);
  End;

  --费用销账申请
  Procedure Zlhis_Charge_009
  (
    费用id_In   Number,
    申请类别_In Number,
    申请时间_In Date
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_CHARGE_009') = 0 Then
      Return;
    End If;
    v_Value := '<root><申请类别>' || 申请类别_In || '</申请类别><费用ID>' || 费用id_In || '</费用ID><申请时间>' ||
               To_Char(申请时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</申请时间>' || '</root>';
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_009', v_Value);
  End;

  --取消销账申请
  Procedure Zlhis_Charge_010
  (
    费用id_In     Number,
    申请类别_In   Number,
    申请时间_In   Date,
    数量_In       Number,
    申请部门id_In Number,
    申请人_In     Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_CHARGE_010') = 0 Then
      Return;
    End If;
    v_Value := '<root><申请类别>' || 申请类别_In || '</申请类别><费用ID>' || 费用id_In || '</费用ID><数量>' || 数量_In || '</数量><申请部门ID>' ||
               申请部门id_In || '</申请部门ID><申请人>' || 申请人_In || '</申请人><申请时间>' || To_Char(申请时间_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</申请时间></root>';
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_010', v_Value);
  End;

  --53.住院患者入院登记
  Procedure Zlhis_Patient_001
  (
    病人id_In In Number,
    主页id_In In Number
  ) Is
    n_变动id Number;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_001') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select Max(ID) From 病人变动记录 Where 病人id = :1 And 主页id = :2 And 开始原因 = 1 And Nvl(附加床位, 0) = 0'
      Into n_变动id
      Using 病人id_In, 主页id_In;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_001',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_001;
  --54.住院患者入院入科
  Procedure Zlhis_Patient_002
  (
    病人id_In In Number,
    主页id_In In Number
  ) Is
    n_变动id Number;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_002') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select Max(ID) From 病人变动记录 Where 病人id = :1 And 主页id = :2 And 终止时间 Is Null And Nvl(附加床位, 0) = 0'
      Into n_变动id
      Using 病人id_In, 主页id_In;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_002',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_002;
  --56.住院患者床位变更
  Procedure Zlhis_Patient_004
  (
    病人id_In In Number,
    主页id_In In Number
  ) Is
    v_原床号   Varchar2(255);
    v_新床号   Varchar2(255);
    n_变动id   Number(18);
    n_开始原因 Number(3);
    d_开始时间 Date;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_004') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select ID, 床号, 开始时间, 开始原因 From 病人变动记录 Where 病人id = :1 And 主页id = :2 And 终止时间 Is Null And Nvl(附加床位, 0) = 0'
      Into n_变动id, v_新床号, d_开始时间, n_开始原因
      Using 病人id_In, 主页id_In;
  
    Execute Immediate 'Select Max(床号) From 病人变动记录 Where 病人id = :1 And 主页id = :2 And 终止时间 = :3 And 终止原因 = :4 And Nvl(附加床位, 0) = 0'
      Into v_原床号
      Using 病人id_In, 主页id_In, d_开始时间, n_开始原因;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_004',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<原床号>' ||
                                 v_原床号 || '</原床号>' || '<新床号>' || v_新床号 || '</新床号>' || '<变动ID>' || n_变动id || '</变动ID>' ||
                                 '</root>');
  End Zlhis_Patient_004;
  --57.住院患者病情变更
  Procedure Zlhis_Patient_005
  (
    病人id_In In Number,
    主页id_In In Number
  ) Is
    n_变动id Number;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_005') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select Max(ID) From 病人变动记录 Where 病人id = :1 And 主页id = :2 And 终止时间 Is Null And Nvl(附加床位, 0) = 0'
      Into n_变动id
      Using 病人id_In, 主页id_In;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_005',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_005;
  --58.住院患者变更撤消
  Procedure Zlhis_Patient_006
  (
    病人id_In   In Number,
    主页id_In   In Number,
    撤销方式_In In Varchar2
  ) Is
    n_Id         Number(18);
    n_科室id     Number(18);
    n_病区id     Number(18);
    n_护理等级id Number(18);
    n_医疗小组id Number(18);
    v_床号       Varchar2(20);
    v_责任护士   Varchar2(50);
    v_主任医师   Varchar2(50);
    v_主治医师   Varchar2(50);
    v_经治医师   Varchar2(50);
    v_病情       Varchar2(50);
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_006') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select  Max(id), Max(科室id), Max(病区id), Max(护理等级id), Max(医疗小组id), Max(床号), Max(责任护士), Max(主任医师), Max(主治医师), Max(经治医师), Max(病情) ' ||
                      'From 病人变动记录 Where 病人id = :1 And 主页id = :2 And (终止时间 Is Null Or 终止原因 = 1) And Nvl(附加床位, 0) = 0'
      Into n_Id, n_科室id, n_病区id, n_护理等级id, n_医疗小组id, v_床号, v_责任护士, v_主任医师, v_主治医师, v_经治医师, v_病情
      Using 病人id_In, 主页id_In;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_006',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><撤销方式>' || 撤销方式_In ||
                                 '</撤销方式><科室ID>' || n_科室id || '</科室ID>' || '<病区ID>' || n_病区id || '</病区ID>' || '<护理等级ID>' ||
                                 n_护理等级id || '</护理等级ID>' || '<医疗小组ID>' || n_医疗小组id || '</医疗小组ID>' || '<床号>' || v_床号 ||
                                 '</床号>' || '<责任护士>' || v_责任护士 || '</责任护士>' || '<主任医师>' || v_主任医师 || '</主任医师>' ||
                                 '<主治医师>' || v_主治医师 || '</主治医师>' || '<经治医师>' || v_经治医师 || '</经治医师>' || '<病情>' || v_病情 ||
                                 '</病情>' || '<ID>' || n_Id || '</ID>' || '</root>');
  End Zlhis_Patient_006;
  --59.住院患者医护变更
  Procedure Zlhis_Patient_007
  (
    病人id_In In Number,
    主页id_In In Number
  ) Is
    v_原住院医生 Varchar2(100);
    v_新住院医生 Varchar2(100);
    v_原主治医生 Varchar2(100);
    v_新主治医生 Varchar2(100);
    v_原主任医生 Varchar2(100);
    v_新主任医生 Varchar2(100);
    v_原责任护士 Varchar2(100);
    v_新责任护士 Varchar2(100);
    n_变动id     Number(18);
    n_开始原因   Number(3);
    d_开始时间   Date;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_007') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select ID, 经治医师, 主治医师, 主任医师, 责任护士, 开始时间, 开始原因 ' ||
                      'From 病人变动记录 Where 病人id = :1 And 主页id = :2 And 终止时间 Is Null And Nvl(附加床位, 0) = 0'
      Into n_变动id, v_新住院医生, v_新主治医生, v_新主任医生, v_新责任护士, d_开始时间, n_开始原因
      Using 病人id_In, 主页id_In;
  
    Execute Immediate 'Select Max(经治医师), Max(主治医师), Max(主任医师), Max(责任护士) ' ||
                      'From 病人变动记录 Where 病人id = :1 And 主页id = :2 And 终止时间 = :3 And 终止原因 = :4 And Nvl(附加床位, 0) = 0'
      Into v_原住院医生, v_原主治医生, v_原主任医生, v_原责任护士
      Using 病人id_In, 主页id_In, d_开始时间, n_开始原因;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_007',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<原住院医生>' ||
                                 v_原住院医生 || '</原住院医生>' || '<新住院医生>' || v_新住院医生 || '</新住院医生>' || '<原主治医生>' || v_原主治医生 ||
                                 '</原主治医生>' || '<新主治医生>' || v_新主治医生 || '</新主治医生>' || '<原主任医生>' || v_原主任医生 || '</原主任医生>' ||
                                 '<新主任医生>' || v_新主任医生 || '</新主任医生>' || '<原责任护士>' || v_原责任护士 || '</原责任护士>' || '<新责任护士>' ||
                                 v_新责任护士 || '</新责任护士>' || '<变动ID>' || n_变动id || '</变动ID>' || '</root>');
  End Zlhis_Patient_007;
  --住院患者护理等级变更
  Procedure Zlhis_Patient_008
  (
    病人id_In In Number,
    主页id_In In Number
  ) Is
    v_原护理等级id Number(18);
    v_新护理等级id Number(18);
    n_变动id       Number(18);
    n_开始原因     Number(3);
    d_开始时间     Date;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_008') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select ID, 护理等级id, 开始时间, 开始原因 From 病人变动记录 ' ||
                      'Where 病人id = :1 And 主页id = :2 And 终止时间 Is Null And Nvl(附加床位, 0) = 0'
      Into n_变动id, v_新护理等级id, d_开始时间, n_开始原因
      Using 病人id_In, 主页id_In;
  
    Execute Immediate 'Select Max(护理等级id) From 病人变动记录 Where 病人id = :1 And 主页id = :2 And 终止时间 = :3 And 终止原因 = :4 And Nvl(附加床位, 0) = 0'
      Into v_原护理等级id
      Using 病人id_In, 主页id_In, d_开始时间, n_开始原因;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_008',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<原护理等级ID>' ||
                                 v_原护理等级id || '</原护理等级ID>' || '<新护理等级ID>' || v_新护理等级id || '</新护理等级ID>' || '<变动ID>' ||
                                 n_变动id || '</变动ID>' || '</root>');
  End Zlhis_Patient_008;
  --60.住院患者预出院
  Procedure Zlhis_Patient_009
  (
    病人id_In In Number,
    主页id_In In Number
  ) Is
    n_变动id Number;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_009') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select Max(ID) From 病人变动记录 Where 病人id = :1 And 主页id = :2 And 终止时间 Is Null And Nvl(附加床位, 0) = 0'
      Into n_变动id
      Using 病人id_In, 主页id_In;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_009',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_009;
  --61.住院患者出院
  Procedure Zlhis_Patient_010
  (
    病人id_In In Number,
    主页id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_010',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID></root>');
  End Zlhis_Patient_010;
  --62.住院患者新生儿登记
  Procedure Zlhis_Patient_011
  (
    病人id_In   In Number,
    主页id_In   In Number,
    婴儿序号_In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_011',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><婴儿序号>' || 婴儿序号_In ||
                                 '</婴儿序号></root>');
  End Zlhis_Patient_011;
  --63.住院患者转入科室
  Procedure Zlhis_Patient_012
  (
    病人id_In In Number,
    主页id_In In Number
  ) Is
    v_转出科室id Number(18);
    v_转入科室id Number(18);
    n_变动id     Number(18);
    n_开始原因   Number(3);
    d_开始时间   Date;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_012') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select ID, 科室id, 开始时间, 开始原因 From 病人变动记录 Where 病人id = :1 And 主页id = :2 And 终止时间 Is Null And Nvl(附加床位, 0) = 0'
      Into n_变动id, v_转入科室id, d_开始时间, n_开始原因
      Using 病人id_In, 主页id_In;
  
    Execute Immediate 'Select Max(科室id) From 病人变动记录 Where 病人id = :1 And 主页id = :2 And 终止时间 = :3 And 终止原因 = :4 And Nvl(附加床位, 0) = 0'
      Into v_转出科室id
      Using 病人id_In, 主页id_In, d_开始时间, n_开始原因;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_012',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<转出科室ID>' ||
                                 v_转出科室id || '</转出科室ID>' || '<转入科室ID>' || v_转入科室id || '</转入科室ID>' || '<变动ID>' || n_变动id ||
                                 '</变动ID>' || '</root>');
  End Zlhis_Patient_012;
  --64.新生儿登记作废
  Procedure Zlhis_Patient_013
  (
    病人id_In   In Number,
    主页id_In   In Number,
    婴儿序号_In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_013',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><婴儿序号>' || 婴儿序号_In ||
                                 '</婴儿序号></root>');
  End Zlhis_Patient_013;
  --65.门诊患者登记
  Procedure Zlhis_Patient_015(病人id_In In Number) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_015', '<root><病人ID>' || 病人id_In || '</病人ID></root>');
  End Zlhis_Patient_015;
  --66.患者信息修改
  Procedure Zlhis_Patient_016(病人id_In In Number) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_016', '<root><病人ID>' || 病人id_In || '</病人ID></root>');
  End Zlhis_Patient_016;

  --67.患者合并
  Procedure Zlhis_Patient_017
  (
    病人id_In   In Number,
    原病人id_In In Number,
    变化ids_In  In Varchar2
  ) Is
    --参数： 1病人id,1主页id:1原病人id,1原主页id; 2病人id,2主页id:2原病人id,2原主页id;….
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_017',
                                '<root><病人ID>' || 病人id_In || '</病人ID><原病人ID>' || 原病人id_In || '</原病人ID><CINFO>' ||
                                 变化ids_In || '</CINFO></root>');
  End Zlhis_Patient_017;

  --69.住院患者转入病区
  Procedure Zlhis_Patient_026
  (
    病人id_In In Number,
    主页id_In In Number
  ) Is
    v_转出病区id Number(18);
    v_转入病区id Number(18);
    n_变动id     Number(18);
    n_开始原因   Number(3);
    d_开始时间   Date;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_026') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select ID, 病区id, 开始时间, 开始原因 From 病人变动记录 Where 病人id = :1 And 主页id = :2 And 终止时间 Is Null And Nvl(附加床位, 0) = 0'
      Into n_变动id, v_转入病区id, d_开始时间, n_开始原因
      Using 病人id_In, 主页id_In;
  
    Execute Immediate 'Select Max(病区id) From 病人变动记录 Where 病人id = :1 And 主页id = :2 And 终止时间 = :3 And 终止原因 = :4 And Nvl(附加床位, 0) = 0'
      Into v_转出病区id
      Using 病人id_In, 主页id_In, d_开始时间, n_开始原因;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_026',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<转出病区ID>' ||
                                 v_转出病区id || '</转出病区ID>' || '<转入病区ID>' || v_转入病区id || '</转入病区ID>' || '<变动ID>' || n_变动id ||
                                 '</变动ID>' || '</root>');
  End Zlhis_Patient_026;

  Procedure Zlhis_Patient_028(病人id_In In Number) Is
    v_姓名     Varchar2(100);
    v_性别     Varchar2(10);
    v_年龄     Varchar2(20);
    v_门诊号   Number(18);
    v_身份证号 Varchar2(20);
    v_出生日期 Varchar2(50);
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_028') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select 姓名, 性别, 年龄, To_Char(出生日期, ''yyyymmdd''), 门诊号, 身份证号 From 病人信息 Where 病人id = :1'
      Into v_姓名, v_性别, v_年龄, v_出生日期, v_门诊号, v_身份证号
      Using 病人id_In;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_028',
                                '<root><病人ID>' || 病人id_In || '</病人ID><姓名>' || v_姓名 || '</姓名>' || '<性别>' || v_性别 ||
                                 '</性别>' || '<年龄>' || v_年龄 || '</年龄>' || '<出生日期>' || v_出生日期 || '</出生日期>' || '<门诊号>' ||
                                 v_门诊号 || '</门诊号>' || '<身份证号>' || v_身份证号 || '</身份证号>' || '</root>');
  End Zlhis_Patient_028;

  --79.留观病人转住院病人
  Procedure Zlhis_Patient_029
  (
    病人id_In In Number,
    主页id_In In Number
  ) Is
    n_变动id Number(18);
  Begin
    Execute Immediate 'Select max(ID) From 病人变动记录 Where 病人id = :1 And 主页id = :2 And 终止时间 Is Null And Nvl(附加床位, 0) = 0 And 开始原因 = 9'
      Into n_变动id
      Using 病人id_In, 主页id_In;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_029',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_029;

  --血库:科室配血完成
  Procedure Zlhis_Blood_001(医嘱id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_001', v_Value);
    End If;
  End Zlhis_Blood_001;

  --血库:科室拒绝配血
  Procedure Zlhis_Blood_002(医嘱id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_002', v_Value);
    End If;
  End Zlhis_Blood_002;

  --70.检验报告审核
  Procedure Zlhis_Lis_001(标本id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_001', v_Value);
    End If;
  End Zlhis_Lis_001;
  --71.检验报告审核撤消
  Procedure Zlhis_Lis_002(标本id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_002', v_Value);
    End If;
  End Zlhis_Lis_002;
  --73.检验标本条码打印
  Procedure Zlhis_Lis_004
  (
    样本条码_In In Varchar2,
    医嘱id_In   In Number,
    医嘱ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
    r_Data  c_Dynamic;
    v_Id    Number(18);
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_LIS_004') = 0 Then
      Return;
    End If;
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || 医嘱id_In || '</医嘱ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004', v_Value);
    Else
      Open r_Data For 'Select 医嘱ID From 病人医嘱发送 Where 医嘱id In (Select Column_Value From Table(f_Num2list(:1)))'
        Using 医嘱ids_In;
      Loop
        Fetch r_Data
          Into v_Id;
        Exit When r_Data%NotFound;
      
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004',
                                    '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || v_Id || '</医嘱ID><系统>1</系统></root>');
      End Loop;
    End If;
  End Zlhis_Lis_004;
  --74.检验标本条码打印撤销
  Procedure Zlhis_Lis_005
  (
    样本条码_In In Varchar2,
    医嘱id_In   In Number,
    医嘱ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
    r_Data  c_Dynamic;
    v_Id    Number(18);
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_LIS_005') = 0 Then
      Return;
    End If;
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || 医嘱id_In || '</医嘱ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005', v_Value);
    Else
      Open r_Data For 'Select 医嘱id From 病人医嘱发送 Where 医嘱id In (Select Column_Value From Table(f_Num2list(:1)))'
        Using 医嘱ids_In;
      Loop
        Fetch r_Data
          Into v_Id;
        Exit When r_Data%NotFound;
      
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005',
                                    '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || v_Id || '</医嘱ID><系统>1</系统></root>');
      End Loop;
    End If;
  End Zlhis_Lis_005;
  --75.检验标本核收
  Procedure Zlhis_Lis_006(标本id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_006', v_Value);
    End If;
  End Zlhis_Lis_006;
  --76.检验标本核收撤销
  Procedure Zlhis_Lis_007(标本id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_007', v_Value);
    End If;
  End Zlhis_Lis_007;
  --77.检验标本拒收
  Procedure Zlhis_Lis_008(标本id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_008', v_Value);
    End If;
  End Zlhis_Lis_008;

  --病历保存
  Procedure Zlhis_Emr_018
  (
    病人id_In In Number,
    主页id_In In Number,
    文件id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><文件ID>' || 文件id_In ||
               '</文件ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_EMR_018', v_Value);
  End Zlhis_Emr_018;
  --管理工具上机人员变动消息
  Procedure Zltools_Users_001
  (
    用户名_In In Varchar2,
    人员id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><用户名>' || 用户名_In || '</用户名><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLTOOLS_USERS_001', v_Value);
  End Zltools_Users_001;
  Procedure Zltools_Users_002
  (
    用户名_In In Varchar2,
    人员id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><用户名>' || 用户名_In || '</用户名><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLTOOLS_USERS_002', v_Value);
  End Zltools_Users_002;
End b_Message;
/

----------------------------------------------------------------------------
--[[11.检查基础]]
----------------------------------------------------------------------------



----------------------------------------------------------------------------
--[[16.临床医嘱]]
----------------------------------------------------------------------------

Create Or Replace Package Pkg_Zyedit As
  -----------------------------------------------------
  --获取中药疾病
  -----------------------------------------------------
  Procedure Get_Distype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --获取中药证型
  -----------------------------------------------------
  Procedure Get_Zxtype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取中药方剂
  -----------------------------------------------------
  Procedure Get_Fjlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取诊断列表
  -----------------------------------------------------
  Procedure Get_Diaglist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --获取中药方剂组成
  -----------------------------------------------------
  Procedure Get_Fjitems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取临证加症
  -----------------------------------------------------
  Procedure Get_Adddis
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取加症治法
  -----------------------------------------------------
  Procedure Get_Addzf
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取加症治法对应草药
  -----------------------------------------------------
  Procedure Get_Additems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取中药脚注(直接取对接系统的脚注)
  -----------------------------------------------------
  Procedure Get_Jzitems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取中药煎法(直接取对接系统的煎法)
  -----------------------------------------------------
  Procedure Get_Jftype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取中药用法(直接取对接系统的用法)
  -----------------------------------------------------
  Procedure Get_Usetype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取中药频率(直接取对接系统的频率)
  -----------------------------------------------------
  Procedure Get_Usetime
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取可用的发药药房(直接取对接系统的发药药房)
  -----------------------------------------------------
  Procedure Get_Drugdept
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取药品库存(直接取对接系统的药品库存)
  -----------------------------------------------------
  Procedure Get_Drugstock
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --加载中医诊断信息
  -----------------------------------------------------
  Procedure Load_Zyedit
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --加载草药明细
  -----------------------------------------------------
  Procedure Load_Zyinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --保存中医诊断信息
  -----------------------------------------------------
  Procedure Save_Zyedit
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --删除中医诊断(单独删除)
  -----------------------------------------------------
  Procedure Del_Zyedit
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --通过HIS医嘱ID获取处方ID和诊断ID
  -----------------------------------------------------
  Procedure Get_Diagid
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取所有方剂
  -----------------------------------------------------
  Procedure Get_Fjall
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取草药目录
  -----------------------------------------------------
  Procedure Get_Drugitems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取HIS品种列表
  -----------------------------------------------------
  Procedure Get_Hisdrug
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --修改草药信息
  -----------------------------------------------------
  Procedure Save_Drugitem
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --HIS品种批量对码
  -----------------------------------------------------
  Procedure Set_Autodrug
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --修改方剂信息
  -----------------------------------------------------
  Procedure Save_Fjitem
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --修改中医疾病
  -----------------------------------------------------
  Procedure Set_Distype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --修改中医证型
  -----------------------------------------------------
  Procedure Set_Zxtype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --设置证型方剂对应
  -----------------------------------------------------
  Procedure Set_Zxtofj
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --修改临证加症
  -----------------------------------------------------
  Procedure Set_Adddis
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --修改加症治法
  -----------------------------------------------------
  Procedure Set_Addzf
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --设置治法草药对应
  -----------------------------------------------------
  Procedure Set_Zftozy
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --项目删除处理
  -----------------------------------------------------
  Procedure Del_Zydata
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --获取数据库系统时间
  -----------------------------------------------------
  Procedure Get_Now_Time
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --异常处理
  -----------------------------------------------------
  Procedure Errorcenter
  (
    Err_Num In Number,
    Err_Msg In Varchar2
  );

End Pkg_Zyedit;
/
Create Or Replace Package Body Pkg_Zyedit As
  -----------------------------------------------------
  --获取中药疾病
  -----------------------------------------------------
  Procedure Get_Distype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select a.疾病id As ID, a.科别, a.疾病名称, a.简码 From 中医疾病 A Order By a.科别, a.疾病名称;
  End Get_Distype;

  -----------------------------------------------------
  --获取中药证型
  -----------------------------------------------------
  Procedure Get_Zxtype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj  Pljson;
    n_疾病id 中医证型.疾病id%Type;
  Begin
    Jsonobj  := Pljson(Input_In);
    n_疾病id := To_Number(Jsonobj.Get_String('疾病ID'));
  
    Open Output_Out For
      Select a.证型id As ID, a.证型名称, a.简码, a.证型治法, a.证型描述, a.症状表现
      From 中医证型 A
      Where a.疾病id = n_疾病id
      Order By a.证型名称;
  End Get_Zxtype;

  -----------------------------------------------------
  --获取中药方剂
  -----------------------------------------------------
  Procedure Get_Fjlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj   Pljson;
    n_证型id  证型方剂对照.证型id%Type;
    v_匹配项  Varchar2(100);
    n_Usetype Number;
  Begin
    Jsonobj   := Pljson(Input_In);
    n_证型id  := To_Number(Nvl(Jsonobj.Get_String('证型ID'), 0));
    n_Usetype := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
  
    If n_证型id = 0 Then
      v_匹配项 := '%' || Upper(Jsonobj.Get_String('匹配项')) || '%';
    
      If n_Usetype = 0 Then
        Open Output_Out For
          Select b.方剂id As ID, b.方剂名称, b.简码 As 简码, b.来源, Decode(Nvl(b.是否保密, 0), 0, b.组成摘要, '保密方剂') As 组成摘要, b.作用描述,
                 b.适应证描述, b.别名, b.别名简码, b.是否保密
          From 治法方剂 B
          Where b.简码 Like v_匹配项 Or b.方剂名称 Like v_匹配项 Or b.别名 Like v_匹配项 Or b.别名简码 Like v_匹配项
          Order By b.方剂名称;
      Else
        Open Output_Out For
          Select b.方剂id As ID, b.方剂名称, b.简码 As 简码, b.来源, b. 组成摘要, b.作用描述, b.适应证描述, b.别名, b.别名简码
          From 治法方剂 B
          Where b.简码 Like v_匹配项 Or b.方剂名称 Like v_匹配项 Or b.别名 Like v_匹配项 Or b.别名简码 Like v_匹配项
          Order By b.方剂名称;
      End If;
    Else
      If n_Usetype = 0 Then
        Open Output_Out For
          Select a.方剂id As ID, b.方剂名称, b.简码 As 简码, b.别名, b.别名简码, b.来源, b.作用描述,
                 Decode(Nvl(b.是否保密, 0), 0, b.组成摘要, '保密方剂') As 组成摘要, b.适应证描述, b.是否保密
          From 证型方剂对照 A, 治法方剂 B
          Where a.方剂id = b.方剂id And a.证型id = n_证型id And a.状态 = 1
          Order By a.对照id, b.方剂名称;
      Else
        Open Output_Out For
          Select a.方剂id As ID, b.方剂名称, b.简码 As 简码, b.别名, b.别名简码, b.来源, b.作用描述, b.组成摘要, b.适应证描述, a.状态, a.对照id
          From 证型方剂对照 A, 治法方剂 B
          Where a.方剂id = b.方剂id And a.证型id = n_证型id
          Order By -a.状态, a.对照id;
      End If;
    End If;
  End Get_Fjlist;

  -----------------------------------------------------
  --获取诊断列表
  -----------------------------------------------------
  Procedure Get_Diaglist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj  Pljson;
    v_匹配项 Varchar2(100);
  Begin
    Jsonobj := Pljson(Input_In);
  
    v_匹配项 := '%' || Upper(Jsonobj.Get_String('匹配项')) || '%';
    Open Output_Out For
      Select a.证型id As ID, b.疾病名称 || '-' || a.证型名称 As 诊断名称, b.简码 || a.简码 As 简码, a.证型治法 As 治法, a.证型描述 As 描述, a.症状表现
      From 中医证型 A, 中医疾病 B
      Where a.疾病id = b.疾病id And (b.疾病名称 || '-' || a.证型名称 Like v_匹配项 Or b.简码 || a.简码 Like v_匹配项)
      Order By b.疾病id;
  End Get_Diaglist;

  -----------------------------------------------------
  --获取中药方剂组成
  -----------------------------------------------------
  Procedure Get_Fjitems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj  Pljson;
    n_方剂id 方剂构成.方剂id%Type;
  Begin
    Jsonobj  := Pljson(Input_In);
    n_方剂id := To_Number(Nvl(Jsonobj.Get_String('方剂ID'), 0));
  
    Open Output_Out For
      Select b.构成id, b.草药id, a.草药名称, b.用法备注, b.古法用量, b.用量, a.单位, Nvl(a.His品种id, 0) As His品种id
      From 方剂构成 B, 草药目录 A
      Where b.草药id = a.草药id And b.方剂id = n_方剂id
      Order By b.构成id;
  End Get_Fjitems;

  -----------------------------------------------------
  --获取临证加症
  -----------------------------------------------------
  Procedure Get_Adddis
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj   Pljson;
    n_Usetype Number; --0 有效数据/-1 全部数据
  Begin
    Jsonobj   := Pljson(Input_In);
    n_Usetype := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
  
    If n_Usetype = 0 Then
      Open Output_Out For
        Select b.加症id As ID, b.加症名称, b.简码 From 临证加症 B Where b.状态 = 1 Order By b.加症id;
    Else
      Open Output_Out For
        Select b.状态, b.加症id As ID, b.加症名称, b.简码 From 临证加症 B Order By -b.状态, b.加症id;
    End If;
  End Get_Adddis;

  -----------------------------------------------------
  --获取加症治法
  -----------------------------------------------------
  Procedure Get_Addzf
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj   Pljson;
    n_加症id  加症治法.加症id%Type;
    n_Usetype Number; --0 有效数据/-1 全部数据
  Begin
    Jsonobj   := Pljson(Input_In);
    n_加症id  := To_Number(Nvl(Jsonobj.Get_String('加症ID'), 0));
    n_Usetype := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
  
    If n_Usetype = 0 Then
      Open Output_Out For
        Select b.治法id As ID, b.治法名称, b.简码
        From 加症治法 B
        Where b.加症id = n_加症id And b.状态 = 1
        Order By b.治法id;
    Else
      Open Output_Out For
        Select b.治法id As ID, b.治法名称, b.简码, b.状态, b.加症id
        From 加症治法 B
        Where b.加症id = n_加症id
        Order By -b.状态, b.治法id;
    End If;
  End Get_Addzf;

  -----------------------------------------------------
  --获取加症治法对应草药
  -----------------------------------------------------
  Procedure Get_Additems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj   Pljson;
    n_治法id  加症用药.治法id%Type;
    n_Usetype Number; --0 有效数据/-1 全部数据
  Begin
    Jsonobj   := Pljson(Input_In);
    n_治法id  := To_Number(Nvl(Jsonobj.Get_String('治法ID'), 0));
    n_Usetype := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
  
    If n_Usetype = 0 Then
      Open Output_Out For
        Select b.草药id As ID, a.草药名称, a.简码, b.用量, a.单位, Nvl(a.His品种id, 0) As His品种id
        From 草药目录 A, 加症用药 B
        Where a.草药id = b.草药id And b.治法id = n_治法id And b.状态 = 1
        Order By b.用药id;
    Else
      Open Output_Out For
        Select b.草药id, a.草药名称, a.简码, b.用量, a.单位, b.状态, b.用药id As ID, b.治法id
        From 草药目录 A, 加症用药 B
        Where a.草药id = b.草药id And b.治法id = n_治法id
        Order By -b.状态, b.用药id;
    End If;
  
  End Get_Additems;

  -----------------------------------------------------
  --获取中药脚注(直接取对接系统的脚注)
  -----------------------------------------------------
  Procedure Get_Jzitems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_中药数据_Edit(2, Null, Output_Out);
  End Get_Jzitems;

  -----------------------------------------------------
  --获取中药煎法(直接取对接系统的煎法)
  -----------------------------------------------------
  Procedure Get_Jftype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_中药数据_Edit(1, Input_In, Output_Out);
  End Get_Jftype;

  -----------------------------------------------------
  --获取中药用法(直接取对接系统的用法)
  -----------------------------------------------------
  Procedure Get_Usetype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_中药数据_Edit(3, Input_In, Output_Out);
  End Get_Usetype;

  -----------------------------------------------------
  --获取中药频率(直接取对接系统的用法)
  -----------------------------------------------------
  Procedure Get_Usetime
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_中药数据_Edit(4, Input_In, Output_Out);
  End Get_Usetime;

  -----------------------------------------------------
  --获取可用的发药药房(直接取对接系统的发药药房)
  -----------------------------------------------------
  Procedure Get_Drugdept
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_中药数据_Edit(5, Input_In, Output_Out);
  End Get_Drugdept;

  -----------------------------------------------------
  --获取药品库存(直接取对接系统的药品库存)
  -----------------------------------------------------
  Procedure Get_Drugstock
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_中药数据_Edit(6, Input_In, Output_Out);
  End Get_Drugstock;

  -----------------------------------------------------
  --加载中医诊断信息
  -----------------------------------------------------
  Procedure Load_Zyedit
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj  Pljson;
    n_诊断id 病人中医诊断记录.诊断id%Type;
  Begin
    Jsonobj  := Pljson(Input_In);
    n_诊断id := To_Number(Jsonobj.Get_String('诊断ID'));
  
    Open Output_Out For
      Select a.处方id, a.方剂id, a.方剂名称, a.付数, a.中药用法, a.中药煎法, a.煎量, a.用药频率, a.频率次数, a.频率间隔, a.间隔单位, a.医生嘱托, a.His煎法id,
             a.His用法id, a.His药房id, b.诊断id, b.就诊方式, b.科别, b.疾病id, b.疾病名称, b.证型id, b.证型名称, b.中医诊断, b.中医治法, b.操作时间, b.操作人,
             b.His诊断id, b.His医嘱id, Nvl(c.是否保密, 0) As 是否保密
      From 病人中医处方记录 A, 病人中医诊断记录 B, 治法方剂 C
      Where b.处方id = a.处方id And a.方剂id = c.方剂id And b.诊断id = n_诊断id;
  End Load_Zyedit;

  -----------------------------------------------------
  --加载草药明细
  -----------------------------------------------------
  Procedure Load_Zyinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj  Pljson;
    n_处方id 病人中医诊断记录.处方id%Type;
  Begin
    Jsonobj  := Pljson(Input_In);
    n_处方id := To_Number(Jsonobj.Get_String('处方ID'));
  
    Open Output_Out For
      Select a.处方明细id, a.处方id, a.序号, a.草药id, a.是否加药, a.来源, a.草药名称, a.用量, a.单位, a.脚注, a.His品种id, a.His规格id
      From 病人中医处方明细 A
      Where a.处方id = n_处方id
      Order By a.序号;
  End Load_Zyinfo;

  -----------------------------------------------------
  --保存中医诊断信息
  -----------------------------------------------------
  Procedure Save_Zyedit
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-新增/1-修改
    n_病人id     病人中医诊断记录.病人id%Type;
    v_挂号单     病人中医诊断记录.挂号单%Type;
    n_门诊号     病人中医诊断记录.门诊号%Type;
    n_诊断id     病人中医诊断记录.诊断id%Type;
    n_科室id     病人中医处方记录.His药房id%Type;
    v_科室名称   Varchar2(100);
    n_操作员id   Number;
    v_操作员姓名 病人中医诊断记录.操作人%Type;
    v_姓名       病人中医诊断记录.姓名%Type;
    v_性别       病人中医诊断记录.性别%Type;
    v_年龄       病人中医诊断记录.年龄%Type;
    v_民族       病人中医诊断记录.民族%Type;
    v_出生日期   Varchar2(100);
    n_就诊方式   病人中医诊断记录.就诊方式%Type;
    v_科别       病人中医诊断记录.科别%Type;
    n_疾病id     病人中医诊断记录.疾病id%Type;
    v_疾病名称   病人中医诊断记录.疾病名称%Type;
    n_证型id     病人中医诊断记录.证型id%Type;
    v_证型名称   病人中医诊断记录.证型名称%Type;
    v_中医诊断   病人中医诊断记录.中医诊断%Type;
    v_中医治法   病人中医诊断记录.中医治法%Type;
    n_方剂id     病人中医处方记录.方剂id%Type;
    v_方剂名称   病人中医处方记录.方剂名称%Type;
    n_付数       病人中医处方记录.付数%Type;
    v_中药用法   病人中医处方记录.中药用法%Type;
    n_His用法id  病人中医处方记录.His用法id%Type;
    v_中药煎法   病人中医处方记录.中药煎法%Type;
    n_His煎法id  病人中医处方记录.His煎法id%Type;
    v_煎量       病人中医处方记录.煎量%Type;
    v_用药频率   病人中医处方记录.用药频率%Type;
    n_频率次数   病人中医处方记录.频率次数%Type;
    n_频率间隔   病人中医处方记录.频率间隔%Type;
    v_间隔单位   病人中医处方记录.间隔单位%Type;
    v_医生嘱托   病人中医处方记录.医生嘱托%Type;
    n_His药房id  病人中医处方记录.His药房id%Type;
  
    n_草药id    病人中医处方明细.草药id%Type;
    n_是否加药  病人中医处方明细.是否加药%Type;
    v_来源      病人中医处方明细.来源%Type;
    v_草药名称  病人中医处方明细.草药名称%Type;
    n_用量      病人中医处方明细.用量%Type;
    v_单位      病人中医处方明细.单位%Type;
    v_脚注      病人中医处方明细.脚注%Type;
    n_His品种id 病人中医处方明细.His品种id%Type;
    n_His规格id 病人中医处方明细.His规格id%Type;
  
    n_His医嘱id   Number;
    n_His诊断id   Number;
    n_是否保密    Number;
    v_中医诊断old 病人中医诊断记录.中医诊断%Type;
  
    n_处方id    病人中医处方记录.处方id%Type;
    d_Now       Date;
    Jsonlistobj Pljson_List;
  
    v_Out Clob;
  
    v_Err_Msg Varchar2(255);
    Err_Item Exception;
  Begin
    --解析入参
    Jsonobj      := Pljson(Input_In);
    n_Usetype    := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_病人id     := To_Number(Jsonobj.Get_String('病人ID'));
    v_挂号单     := Jsonobj.Get_String('挂号单');
    n_门诊号     := To_Number(Jsonobj.Get_String('门诊号'));
    n_诊断id     := To_Number(Jsonobj.Get_String('诊断ID'));
    n_科室id     := To_Number(Jsonobj.Get_String('科室ID'));
    v_科室名称   := Jsonobj.Get_String('科室名称');
    n_操作员id   := To_Number(Jsonobj.Get_String('操作员ID'));
    v_操作员姓名 := Jsonobj.Get_String('操作员姓名');
    v_姓名       := Jsonobj.Get_String('姓名');
    v_性别       := Jsonobj.Get_String('性别');
    v_年龄       := Jsonobj.Get_String('年龄');
    v_民族       := Jsonobj.Get_String('民族');
    v_出生日期   := Jsonobj.Get_String('出生日期');
    n_就诊方式   := To_Number(Jsonobj.Get_String('就诊方式'));
    v_科别       := Jsonobj.Get_String('科别');
    n_疾病id     := To_Number(Jsonobj.Get_String('疾病ID'));
    v_疾病名称   := Jsonobj.Get_String('疾病名称');
    n_证型id     := To_Number(Jsonobj.Get_String('证型ID'));
    v_证型名称   := Jsonobj.Get_String('证型名称');
    v_中医诊断   := Jsonobj.Get_String('中医诊断');
    v_中医治法   := Jsonobj.Get_String('中医治法');
    n_方剂id     := To_Number(Jsonobj.Get_String('方剂ID'));
    v_方剂名称   := Jsonobj.Get_String('方剂名称');
    n_付数       := To_Number(Jsonobj.Get_String('付数'));
    v_中药用法   := Jsonobj.Get_String('中药用法');
    n_His用法id  := To_Number(Jsonobj.Get_String('HIS用法ID'));
    v_中药煎法   := Jsonobj.Get_String('中药煎法');
    n_His煎法id  := To_Number(Jsonobj.Get_String('HIS煎法ID'));
    v_煎量       := Jsonobj.Get_String('煎量');
    v_用药频率   := Jsonobj.Get_String('用药频率');
    n_频率次数   := To_Number(Jsonobj.Get_String('频率次数'));
    n_频率间隔   := To_Number(Jsonobj.Get_String('频率间隔'));
    v_间隔单位   := Jsonobj.Get_String('间隔单位');
    v_医生嘱托   := Jsonobj.Get_String('医生嘱托');
    n_His药房id  := To_Number(Jsonobj.Get_String('HIS药房ID'));
    Jsonlistobj  := Jsonobj.Get_Pljson_List('处方明细');
  
    Select Sysdate Into d_Now From Dual;
  
    --新增
    If n_Usetype = 0 Then
      --编辑数据
      Select 病人中医诊断记录_诊断id.Nextval Into n_诊断id From Dual;
      Select 病人中医处方记录_处方id.Nextval Into n_处方id From Dual;
    
      Insert Into 病人中医处方记录
        (处方id, 方剂id, 方剂名称, 付数, 中药用法, 中药煎法, 煎量, 用药频率, 频率次数, 频率间隔, 间隔单位, 医生嘱托, His煎法id, His用法id, His药房id)
      Values
        (n_处方id, n_方剂id, v_方剂名称, n_付数, v_中药用法, v_中药煎法, v_煎量, v_用药频率, n_频率次数, n_频率间隔, v_间隔单位, v_医生嘱托, n_His煎法id,
         n_His用法id, n_His药房id);
    
      Insert Into 病人中医诊断记录
        (诊断id, 病人id, 挂号单, 姓名, 门诊号, 性别, 年龄, 民族, 出生日期, 就诊方式, 科别, 疾病id, 疾病名称, 证型id, 证型名称, 中医诊断, 中医治法, 处方id, 操作时间, 操作人)
      Values
        (n_诊断id, n_病人id, v_挂号单, v_姓名, n_门诊号, v_性别, v_年龄, v_民族, To_Date(v_出生日期, 'yyyy-mm-dd'), n_就诊方式, v_科别, n_疾病id,
         v_疾病名称, n_证型id, v_证型名称, v_中医诊断, v_中医治法, n_处方id, d_Now, v_操作员姓名);
    
      For I In 1 .. Jsonlistobj.Count Loop
        Jsonobj     := Pljson();
        Jsonobj     := Pljson(Jsonlistobj.Get(I));
        n_草药id    := To_Number(Jsonobj.Get_String('草药ID'));
        n_是否加药  := To_Number(Jsonobj.Get_String('是否加药'));
        v_来源      := Jsonobj.Get_String('来源');
        v_草药名称  := Jsonobj.Get_String('草药名称');
        n_用量      := To_Number(Jsonobj.Get_String('用量'));
        v_单位      := Jsonobj.Get_String('单位');
        v_脚注      := Jsonobj.Get_String('脚注');
        n_His品种id := To_Number(Jsonobj.Get_String('HIS品种ID'));
        n_His规格id := To_Number(Jsonobj.Get_String('HIS规格ID'));
      
        Insert Into 病人中医处方明细
          (处方明细id, 处方id, 序号, 草药id, 是否加药, 来源, 草药名称, 用量, 单位, 脚注, His品种id, His规格id)
        Values
          (病人中医处方明细_处方明细id.Nextval, n_处方id, I, n_草药id, n_是否加药, v_来源, v_草药名称, n_用量, v_单位, v_脚注, n_His品种id, n_His规格id);
      End Loop;
    Else
      Select Max(处方id), Max(His医嘱id), Max(His诊断id), Max(中医诊断)
      Into n_处方id, n_His医嘱id, n_His诊断id, v_中医诊断old
      From 病人中医诊断记录
      Where 诊断id = n_诊断id;
    
      If Nvl(n_诊断id, 0) = 0 Or Nvl(n_处方id, 0) = 0 Then
        v_Err_Msg := '未找到病人诊断对应的处方数据。';
        Raise Err_Item;
      End If;
    
      Update 病人中医处方记录
      Set 方剂id = n_方剂id, 方剂名称 = v_方剂名称, 付数 = n_付数, 中药用法 = v_中药用法, 中药煎法 = v_中药煎法, 煎量 = v_煎量, 用药频率 = v_用药频率, 频率次数 = n_频率次数,
          频率间隔 = n_频率间隔, 间隔单位 = v_间隔单位, 医生嘱托 = v_医生嘱托, His煎法id = n_His煎法id, His用法id = n_His用法id, His药房id = n_His药房id
      Where 处方id = n_处方id;
    
      Update 病人中医诊断记录
      Set 病人id = n_病人id, 挂号单 = v_挂号单, 姓名 = v_姓名, 门诊号 = n_门诊号, 性别 = v_性别, 年龄 = v_年龄, 民族 = v_民族,
          出生日期 = To_Date(v_出生日期, 'yyyy-mm-dd'), 就诊方式 = n_就诊方式, 科别 = v_科别, 疾病id = n_疾病id, 疾病名称 = v_疾病名称, 证型id = n_证型id,
          证型名称 = v_证型名称, 中医诊断 = v_中医诊断, 中医治法 = v_中医治法, 处方id = n_处方id, 操作时间 = d_Now, 操作人 = v_操作员姓名
      
      Where 诊断id = n_诊断id;
    
      Delete From 病人中医处方明细 Where 处方id = n_处方id;
    
      For I In 1 .. Jsonlistobj.Count Loop
        Jsonobj     := Pljson();
        Jsonobj     := Pljson(Jsonlistobj.Get(I));
        n_草药id    := To_Number(Jsonobj.Get_String('草药ID'));
        n_是否加药  := To_Number(Jsonobj.Get_String('是否加药'));
        v_来源      := Jsonobj.Get_String('来源');
        v_草药名称  := Jsonobj.Get_String('草药名称');
        n_用量      := To_Number(Jsonobj.Get_String('用量'));
        v_单位      := Jsonobj.Get_String('单位');
        v_脚注      := Jsonobj.Get_String('脚注');
        n_His品种id := To_Number(Jsonobj.Get_String('HIS品种ID'));
        n_His规格id := To_Number(Jsonobj.Get_String('HIS规格ID'));
      
        Insert Into 病人中医处方明细
          (处方明细id, 处方id, 序号, 草药id, 是否加药, 来源, 草药名称, 用量, 单位, 脚注, His品种id, His规格id)
        Values
          (病人中医处方明细_处方明细id.Nextval, n_处方id, I, n_草药id, n_是否加药, v_来源, v_草药名称, n_用量, v_单位, v_脚注, n_His品种id, n_His规格id);
      End Loop;
    End If;
  
    Select Nvl(Max(是否保密), 0) Into n_是否保密 From 治法方剂 Where 方剂id = n_方剂id;
  
    --同步HIS医嘱诊断
    Zl_中医诊断_Save(Input_In, n_His医嘱id, n_His诊断id, v_中医诊断old, n_是否保密, v_Out);
    If Nvl(v_Out, '空') != '空' Then
      Jsonobj := Pljson();
      Jsonobj := Pljson(v_Out);
      If To_Number(Jsonobj.Get_String('His诊断id')) != 0 And To_Number(Nvl(Jsonobj.Get_String('His诊断id'), 0)) != 0 Then
        Update 病人中医诊断记录
        Set His医嘱id = To_Number(Nvl(Jsonobj.Get_String('His医嘱id'), 0)),
            His诊断id = To_Number(Nvl(Jsonobj.Get_String('His诊断id'), 0))
        Where 诊断id = n_诊断id;
      Else
        v_Err_Msg := '中医诊断保存失败,请检查HIS同步接口。';
        Raise Err_Item;
      End If;
    Else
      v_Err_Msg := '中医诊断保存失败,请检查HIS同步接口。';
      Raise Err_Item;
    End If;
  
    Open Output_Out For
      Select To_Number(Nvl(Jsonobj.Get_String('His医嘱id'), 0)) As His医嘱id,
             To_Number(Nvl(Jsonobj.Get_String('His诊断id'), 0)) As His诊断id, n_诊断id As 诊断id, n_处方id As 处方id
      From Dual;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  End Save_Zyedit;

  -----------------------------------------------------
  --删除中医诊断(单独删除)
  -----------------------------------------------------
  Procedure Del_Zyedit
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj  Pljson;
    n_诊断id 病人中医诊断记录.诊断id%Type;
    n_处方id 病人中医诊断记录.处方id%Type;
  
    v_Err_Msg Varchar2(255);
    Err_Item Exception;
  Begin
    Jsonobj  := Pljson(Input_In);
    n_诊断id := To_Number(Jsonobj.Get_String('诊断ID'));
  
    Select Max(处方id) Into n_处方id From 病人中医诊断记录 Where 诊断id = n_诊断id;
  
    If Nvl(n_诊断id, 0) = 0 Or Nvl(n_处方id, 0) = 0 Then
      v_Err_Msg := '未找到病人诊断对应的处方数据。';
      Raise Err_Item;
    End If;
  
    Delete From 病人中医处方记录 Where 处方id = n_处方id;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  End Del_Zyedit;

  -----------------------------------------------------
  --通过HIS医嘱ID或者HIS诊断ID获取处方ID和诊断ID
  -----------------------------------------------------
  Procedure Get_Diagid
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj     Pljson;
    n_His医嘱id 病人中医诊断记录.His医嘱id%Type;
    n_His诊断id 病人中医诊断记录.His诊断id%Type;
  Begin
    Jsonobj     := Pljson(Input_In);
    n_His医嘱id := To_Number(Jsonobj.Get_String('HIS医嘱ID'));
    n_His诊断id := To_Number(Jsonobj.Get_String('HIS诊断ID'));
    Open Output_Out For
      Select a.处方id, a.诊断id From 病人中医诊断记录 A Where a.His医嘱id = n_His医嘱id Or a.His诊断id = n_His诊断id;
  End Get_Diagid;

  -----------------------------------------------------
  --获取草药目录
  -----------------------------------------------------
  Procedure Get_Drugitems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select a.草药id As ID, a.草药名称, a.简码, a.别名, a.别名简码, a.单位, Null As His对码, a.来源, a.His品种id, a.草药描述, a.性状, a.药性, a.适应证,
             a.用法, a.服法, a.禁忌, a.成分, a.药理作用, a.创建人, To_Char(a.创建时间, 'yyyy-MM-dd hh24:mi') As 创建时间, a.最后修改人,
             To_Char(a.最后修改时间, 'yyyy-MM-dd hh24:mi') As 最后修改时间
      From 草药目录 A
      Order By a.草药id, a.草药名称;
  End Get_Drugitems;

  -----------------------------------------------------
  --获取所有方剂
  -----------------------------------------------------
  Procedure Get_Fjall
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 方剂id As ID, 方剂名称, 简码, 别名, 别名简码, 来源, 组成摘要, 服法描述, 作用描述, 制法描述, 适应证描述, 方剂组成作用描述, 创建人,
             To_Char(创建时间, 'yyyy-MM-dd hh24:mi') As 创建时间, 最后修改人, To_Char(最后修改时间, 'yyyy-MM-dd hh24:mi') As 最后修改时间,
             Nvl(是否保密, 0) As 是否保密, Decode(Nvl(是否保密, 0), 1, '√', '') As 密
      From 治法方剂
      Order By 方剂id, 方剂名称;
  End Get_Fjall;

  -----------------------------------------------------
  --获取HIS品种列表
  -----------------------------------------------------
  Procedure Get_Hisdrug
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_中药数据_Edit(7, Input_In, Output_Out);
  End Get_Hisdrug;

  -----------------------------------------------------
  --修改草药信息
  -----------------------------------------------------
  Procedure Save_Drugitem
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-新增/1-修改/2-修改HIS对码
    n_草药id     草药目录.草药id%Type;
    v_草药名称   草药目录.草药名称%Type;
    v_简码       草药目录.简码%Type;
    v_别名       草药目录.别名%Type;
    v_别名简码   草药目录.别名简码%Type;
    v_来源       草药目录.来源%Type;
    v_单位       草药目录.单位%Type;
    v_草药描述   草药目录.草药描述%Type;
    v_性状       草药目录.性状%Type;
    v_药性       草药目录.药性%Type;
    v_适应证     草药目录.适应证%Type;
    v_用法       草药目录.用法%Type;
    v_服法       草药目录.服法%Type;
    v_禁忌       草药目录.禁忌%Type;
    v_成分       草药目录.成分%Type;
    v_药理作用   草药目录.药理作用%Type;
    n_His品种id  草药目录.His品种id%Type;
    v_操作员名称 草药目录.最后修改人%Type;
    n_操作员id   Number;
  Begin
  
    --解析入参
    Jsonobj      := Pljson(Input_In);
    n_Usetype    := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_草药id     := To_Number(Nvl(Jsonobj.Get_String('草药ID'), 0));
    v_草药名称   := Jsonobj.Get_String('草药名称');
    v_简码       := Jsonobj.Get_String('简码');
    v_别名       := Jsonobj.Get_String('别名');
    v_别名简码   := Jsonobj.Get_String('别名简码');
    v_单位       := Jsonobj.Get_String('单位');
    v_来源       := Jsonobj.Get_String('来源');
    v_草药描述   := Jsonobj.Get_String('草药描述');
    v_性状       := Jsonobj.Get_String('性状');
    v_药性       := Jsonobj.Get_String('药性');
    v_适应证     := Jsonobj.Get_String('适应证');
    v_用法       := Jsonobj.Get_String('用法');
    v_服法       := Jsonobj.Get_String('服法');
    v_禁忌       := Jsonobj.Get_String('禁忌');
    v_成分       := Jsonobj.Get_String('成分');
    v_药理作用   := Jsonobj.Get_String('药理作用');
    n_His品种id  := To_Number(Nvl(Jsonobj.Get_String('HIS品种ID'), 0));
    v_操作员名称 := Jsonobj.Get_String('操作员名称');
    n_操作员id   := To_Number(Nvl(Jsonobj.Get_String('操作员ID'), 0));
    If n_His品种id = 0 Then
      n_His品种id := Null;
    End If;
    If n_Usetype = 0 Then
      Select 草药目录_草药id.Nextval Into n_草药id From Dual;
      Insert Into 草药目录
        (草药id, 草药名称, 简码, 别名, 别名简码, 单位, 来源, 草药描述, 性状, 药性, 适应证, 用法, 服法, 禁忌, 成分, 药理作用, His品种id, 创建人, 创建时间, 最后修改人, 最后修改时间)
      Values
        (n_草药id, v_草药名称, v_简码, v_别名, v_别名简码, v_单位, v_来源, v_草药描述, v_性状, v_药性, v_适应证, v_用法, v_服法, v_禁忌, v_成分, v_药理作用,
         n_His品种id, v_操作员名称, Sysdate, v_操作员名称, Sysdate);
    Elsif n_Usetype = 1 Then
      Update 草药目录
      Set 草药名称 = v_草药名称, 简码 = v_简码, 别名 = v_别名, 别名简码 = v_别名简码, 单位 = v_单位, 来源 = v_来源, 草药描述 = v_草药描述, 性状 = v_性状, 药性 = v_药性,
          适应证 = v_适应证, 用法 = v_用法, 服法 = v_服法, 禁忌 = v_禁忌, 成分 = v_成分, 药理作用 = v_药理作用, His品种id = n_His品种id, 最后修改人 = v_操作员名称,
          最后修改时间 = Sysdate
      Where 草药id = n_草药id;
    Elsif n_Usetype = 2 Then
      Update 草药目录
      Set His品种id = n_His品种id, 最后修改人 = v_操作员名称, 最后修改时间 = Sysdate
      Where 草药id = n_草药id;
    End If;
    Open Output_Out For
      Select n_草药id As 草药id From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Save_Drugitem;

  -----------------------------------------------------
  --HIS品种批量对码
  -----------------------------------------------------
  Procedure Set_Autodrug
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_中药数据_Edit(8, Input_In, Output_Out);
  End Set_Autodrug;

  -----------------------------------------------------
  --修改方剂信息
  -----------------------------------------------------
  Procedure Save_Fjitem
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj            Pljson;
    n_Usetype          Number; --0-新增/1-修改
    n_方剂id           治法方剂.方剂id%Type;
    v_方剂名称         治法方剂.方剂名称%Type;
    v_简码             治法方剂.简码%Type;
    v_别名             治法方剂.别名%Type;
    v_别名简码         治法方剂.别名简码%Type;
    v_来源             治法方剂.来源%Type;
    v_组成摘要         治法方剂.组成摘要%Type;
    v_服法描述         治法方剂.服法描述%Type;
    v_作用描述         治法方剂.作用描述%Type;
    v_制法描述         治法方剂.制法描述%Type;
    v_适应证描述       治法方剂.适应证描述%Type;
    v_方剂组成作用描述 治法方剂.方剂组成作用描述%Type;
    n_是否保密         治法方剂.是否保密%Type;
    v_操作员名称       治法方剂.最后修改人%Type;
    n_操作员id         Number;
    Jsonlistobj        Pljson_List;
  
    n_草药id   方剂构成.草药id%Type;
    n_用量     方剂构成.用量%Type;
    v_用法备注 方剂构成.用法备注%Type;
    v_古法用量 方剂构成.古法用量%Type;
  Begin
  
    --解析入参
    Jsonobj            := Pljson(Input_In);
    n_Usetype          := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_方剂id           := To_Number(Nvl(Jsonobj.Get_String('方剂ID'), 0));
    v_方剂名称         := Jsonobj.Get_String('方剂名称');
    v_简码             := Jsonobj.Get_String('简码');
    v_别名             := Jsonobj.Get_String('别名');
    v_别名简码         := Jsonobj.Get_String('别名简码');
    v_来源             := Jsonobj.Get_String('来源');
    v_组成摘要         := Jsonobj.Get_String('组成摘要');
    v_服法描述         := Jsonobj.Get_String('服法描述');
    v_作用描述         := Jsonobj.Get_String('作用描述');
    v_制法描述         := Jsonobj.Get_String('制法描述');
    v_适应证描述       := Jsonobj.Get_String('适应证描述');
    v_方剂组成作用描述 := Jsonobj.Get_String('方剂组成作用描述');
    n_是否保密         := To_Number(Nvl(Jsonobj.Get_String('是否保密'), 0));
    v_操作员名称       := Jsonobj.Get_String('操作员名称');
    n_操作员id         := To_Number(Nvl(Jsonobj.Get_String('操作员ID'), 0));
    Jsonlistobj        := Jsonobj.Get_Pljson_List('方剂构成');
  
    If n_Usetype = 0 Then
      Select 治法方剂_方剂id.Nextval Into n_方剂id From Dual;
      Insert Into 治法方剂
        (方剂id, 方剂名称, 简码, 别名, 别名简码, 来源, 组成摘要, 服法描述, 作用描述, 制法描述, 适应证描述, 方剂组成作用描述, 是否保密, 创建人, 创建时间, 最后修改人, 最后修改时间)
      Values
        (n_方剂id, v_方剂名称, v_简码, v_别名, v_别名简码, v_来源, v_组成摘要, v_服法描述, v_作用描述, v_制法描述, v_适应证描述, v_方剂组成作用描述, n_是否保密, v_操作员名称,
         Sysdate, v_操作员名称, Sysdate);
    Else
      Update 治法方剂
      Set 方剂名称 = v_方剂名称, 简码 = v_简码, 别名 = v_别名, 别名简码 = v_别名简码, 来源 = v_来源, 组成摘要 = v_组成摘要, 服法描述 = v_服法描述, 作用描述 = v_作用描述,
          制法描述 = v_制法描述, 适应证描述 = v_适应证描述, 方剂组成作用描述 = v_方剂组成作用描述, 是否保密 = n_是否保密, 最后修改人 = v_操作员名称, 最后修改时间 = Sysdate
      Where 方剂id = n_方剂id;
    End If;
  
    --更新方剂构成
    If n_Usetype = 1 Then
      Delete From 方剂构成 Where 方剂id = n_方剂id;
    End If;
    For I In 1 .. Jsonlistobj.Count Loop
      Jsonobj    := Pljson(Jsonlistobj.Get(I));
      n_草药id   := To_Number(Jsonobj.Get_String('草药ID'));
      n_用量     := To_Number(Jsonobj.Get_String('用量'));
      v_用法备注 := Jsonobj.Get_String('用法备注');
      v_古法用量 := Jsonobj.Get_String('古法用量');
    
      Insert Into 方剂构成
        (构成id, 方剂id, 草药id, 用法备注, 古法用量, 用量, 创建人, 创建时间)
      Values
        (方剂构成_构成id.Nextval, n_方剂id, n_草药id, v_用法备注, v_古法用量, n_用量, v_操作员名称, Sysdate);
    End Loop;
  
    Open Output_Out For
      Select n_方剂id As 方剂id From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Save_Fjitem;

  -----------------------------------------------------
  --修改中医疾病
  -----------------------------------------------------
  Procedure Set_Distype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-新增/1-修改
    n_疾病id     中医疾病.疾病id%Type;
    v_疾病名称   中医疾病.疾病名称%Type;
    v_简码       中医疾病.简码%Type;
    v_科别       中医疾病.科别%Type;
    v_操作员名称 中医疾病.最后修改人%Type;
    n_操作员id   Number;
  Begin
  
    --解析入参
    Jsonobj      := Pljson(Input_In);
    n_Usetype    := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_疾病id     := To_Number(Nvl(Jsonobj.Get_String('疾病ID'), 0));
    v_疾病名称   := Jsonobj.Get_String('疾病名称');
    v_简码       := Jsonobj.Get_String('简码');
    v_科别       := Jsonobj.Get_String('科别');
    v_操作员名称 := Jsonobj.Get_String('操作员名称');
    n_操作员id   := To_Number(Nvl(Jsonobj.Get_String('操作员ID'), 0));
    If n_Usetype = 0 Then
      Select 中医疾病_疾病id.Nextval Into n_疾病id From Dual;
      Insert Into 中医疾病
        (疾病id, 疾病名称, 简码, 科别, 创建人, 创建时间, 最后修改人, 最后修改时间)
      Values
        (n_疾病id, v_疾病名称, v_简码, v_科别, v_操作员名称, Sysdate, v_操作员名称, Sysdate);
    Else
      Update 中医疾病
      Set 疾病名称 = v_疾病名称, 简码 = v_简码, 科别 = v_科别, 最后修改人 = v_操作员名称, 最后修改时间 = Sysdate
      Where 疾病id = n_疾病id;
    End If;
    Open Output_Out For
      Select n_疾病id As 疾病id From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Set_Distype;

  -----------------------------------------------------
  --修改中医证型
  -----------------------------------------------------
  Procedure Set_Zxtype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-新增/1-修改
    n_证型id     中医证型.证型id%Type;
    v_证型名称   中医证型.证型名称%Type;
    v_简码       中医证型.简码%Type;
    n_疾病id     中医证型.疾病id%Type;
    v_证型描述   中医证型.证型描述%Type;
    v_证型治法   中医证型.证型治法%Type;
    v_症状表现   中医证型.症状表现%Type;
    v_操作员名称 中医证型.最后修改人%Type;
    n_操作员id   Number;
  Begin
  
    --解析入参
    Jsonobj    := Pljson(Input_In);
    n_Usetype  := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_证型id   := To_Number(Nvl(Jsonobj.Get_String('证型ID'), 0));
    v_证型名称 := Jsonobj.Get_String('证型名称');
    v_简码     := Jsonobj.Get_String('简码');
    n_疾病id   := To_Number(Nvl(Jsonobj.Get_String('疾病ID'), 0));
    v_证型描述 := Jsonobj.Get_String('证型描述');
    v_证型治法 := Jsonobj.Get_String('证型治法');
    v_症状表现 := Jsonobj.Get_String('症状表现');
  
    v_操作员名称 := Jsonobj.Get_String('操作员名称');
    n_操作员id   := To_Number(Nvl(Jsonobj.Get_String('操作员ID'), 0));
    If n_Usetype = 0 Then
      Select 中医证型_证型id.Nextval Into n_证型id From Dual;
      Insert Into 中医证型
        (证型id, 证型名称, 简码, 疾病id, 证型描述, 证型治法, 症状表现, 创建人, 创建时间, 最后修改人, 最后修改时间)
      Values
        (n_证型id, v_证型名称, v_简码, n_疾病id, v_证型描述, v_证型治法, v_症状表现, v_操作员名称, Sysdate, v_操作员名称, Sysdate);
    Else
      Update 中医证型
      Set 证型名称 = v_证型名称, 简码 = v_简码, 证型描述 = v_证型描述, 证型治法 = v_证型治法, 症状表现 = v_症状表现, 最后修改人 = v_操作员名称, 最后修改时间 = Sysdate
      Where 证型id = n_证型id;
    End If;
    Open Output_Out For
      Select n_证型id As 证型id From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Set_Zxtype;

  -----------------------------------------------------
  --设置证型方剂对应
  -----------------------------------------------------
  Procedure Set_Zxtofj
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-新增/1-更改状态
    n_对照id     证型方剂对照.对照id%Type;
    n_证型id     证型方剂对照.证型id%Type;
    n_方剂id     证型方剂对照.方剂id%Type;
    n_状态       证型方剂对照.状态%Type;
    v_操作员名称 中医证型.最后修改人%Type;
    n_操作员id   Number;
  Begin
    --解析入参
    Jsonobj   := Pljson(Input_In);
    n_Usetype := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
  
    v_操作员名称 := Jsonobj.Get_String('操作员名称');
    n_操作员id   := To_Number(Nvl(Jsonobj.Get_String('操作员ID'), 0));
    If n_Usetype = 0 Then
      n_证型id := To_Number(Nvl(Jsonobj.Get_String('证型ID'), 0));
      n_方剂id := To_Number(Nvl(Jsonobj.Get_String('方剂ID'), 0));
      Select 证型方剂对照_对照id.Nextval Into n_对照id From Dual;
      Insert Into 证型方剂对照
        (对照id, 证型id, 方剂id, 状态, 创建人, 创建时间, 最后修改人, 最后修改时间)
      Values
        (n_对照id, n_证型id, n_方剂id, 1, v_操作员名称, Sysdate, v_操作员名称, Sysdate);
    Elsif n_Usetype = 1 Then
      n_状态   := To_Number(Nvl(Jsonobj.Get_String('状态'), 0));
      n_对照id := To_Number(Nvl(Jsonobj.Get_String('对照ID'), 0));
      Update 证型方剂对照 Set 状态 = n_状态, 最后修改人 = v_操作员名称, 最后修改时间 = Sysdate Where 对照id = n_对照id;
    End If;
    Open Output_Out For
      Select n_对照id As 对照id From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Set_Zxtofj;

  -----------------------------------------------------
  --修改临证加症
  -----------------------------------------------------
  Procedure Set_Adddis
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-新增/1-修改/2-改变状态
    n_加症id     临证加症.加症id%Type;
    v_加症名称   临证加症.加症名称%Type;
    v_简码       临证加症.简码%Type;
    n_状态       临证加症.状态%Type;
    v_操作员名称 临证加症.最后修改人%Type;
    n_操作员id   Number;
  Begin
  
    --解析入参
    Jsonobj      := Pljson(Input_In);
    n_Usetype    := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_加症id     := To_Number(Nvl(Jsonobj.Get_String('加症ID'), 0));
    v_加症名称   := Jsonobj.Get_String('加症名称');
    v_简码       := Jsonobj.Get_String('简码');
    v_操作员名称 := Jsonobj.Get_String('操作员名称');
    n_操作员id   := To_Number(Nvl(Jsonobj.Get_String('操作员ID'), 0));
    If n_Usetype = 0 Then
      Select 临证加症_加症id.Nextval Into n_加症id From Dual;
      Insert Into 临证加症
        (加症id, 加症名称, 简码, 状态, 创建人, 创建时间, 最后修改人, 最后修改时间)
      Values
        (n_加症id, v_加症名称, v_简码, 1, v_操作员名称, Sysdate, v_操作员名称, Sysdate);
    Elsif n_Usetype = 1 Then
      Update 临证加症
      Set 加症名称 = v_加症名称, 简码 = v_简码, 最后修改人 = v_操作员名称, 最后修改时间 = Sysdate
      Where 加症id = n_加症id;
    Elsif n_Usetype = 2 Then
      n_状态 := To_Number(Nvl(Jsonobj.Get_String('状态'), 0));
      Update 临证加症 Set 状态 = n_状态, 最后修改人 = v_操作员名称, 最后修改时间 = Sysdate Where 加症id = n_加症id;
    End If;
    Open Output_Out For
      Select n_加症id As 加症id From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Set_Adddis;

  -----------------------------------------------------
  --修改加症治法
  -----------------------------------------------------
  Procedure Set_Addzf
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-新增/1-修改/2-改变状态
    n_治法id     加症治法.治法id%Type;
    v_治法名称   加症治法.治法名称%Type;
    n_加症id     加症治法.加症id%Type;
    v_简码       加症治法.简码%Type;
    n_状态       加症治法.状态%Type;
    v_操作员名称 加症治法.最后修改人%Type;
    n_操作员id   Number;
  Begin
    --解析入参
    Jsonobj      := Pljson(Input_In);
    n_Usetype    := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_治法id     := To_Number(Nvl(Jsonobj.Get_String('治法ID'), 0));
    n_加症id     := To_Number(Nvl(Jsonobj.Get_String('加症ID'), 0));
    v_治法名称   := Jsonobj.Get_String('治法名称');
    v_简码       := Jsonobj.Get_String('简码');
    v_操作员名称 := Jsonobj.Get_String('操作员名称');
    n_操作员id   := To_Number(Nvl(Jsonobj.Get_String('操作员ID'), 0));
    If n_Usetype = 0 Then
      Select 加症治法_治法id.Nextval Into n_治法id From Dual;
      Insert Into 加症治法
        (治法id, 治法名称, 简码, 加症id, 状态, 创建人, 创建时间, 最后修改人, 最后修改时间)
      Values
        (n_治法id, v_治法名称, v_简码, n_加症id, 1, v_操作员名称, Sysdate, v_操作员名称, Sysdate);
    Elsif n_Usetype = 1 Then
      Update 加症治法
      Set 治法名称 = v_治法名称, 简码 = v_简码, 最后修改人 = v_操作员名称, 最后修改时间 = Sysdate
      Where 治法id = n_治法id;
    Elsif n_Usetype = 2 Then
      n_状态 := To_Number(Nvl(Jsonobj.Get_String('状态'), 0));
      Update 加症治法 Set 状态 = n_状态, 最后修改人 = v_操作员名称, 最后修改时间 = Sysdate Where 治法id = n_治法id;
    End If;
    Open Output_Out For
      Select n_治法id As 治法id From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Set_Addzf;

  -----------------------------------------------------
  --设置治法草药对应
  -----------------------------------------------------
  Procedure Set_Zftozy
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-新增/2-修改/1-更改状态
    n_用药id     加症用药.用药id%Type;
    n_治法id     加症用药.治法id%Type;
    n_草药id     加症用药.草药id%Type;
    n_用量       加症用药.用量%Type;
    n_状态       加症用药.状态%Type;
    v_操作员名称 加症用药.最后修改人%Type;
    n_操作员id   Number;
  Begin
    --解析入参
    Jsonobj   := Pljson(Input_In);
    n_Usetype := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
  
    v_操作员名称 := Jsonobj.Get_String('操作员名称');
    n_操作员id   := To_Number(Nvl(Jsonobj.Get_String('操作员ID'), 0));
    If n_Usetype = 0 Then
      n_治法id := To_Number(Nvl(Jsonobj.Get_String('治法ID'), 0));
      n_草药id := To_Number(Nvl(Jsonobj.Get_String('草药ID'), 0));
      n_用量   := To_Number(Nvl(Jsonobj.Get_String('用量'), 0));
      Select 加症用药_用药id.Nextval Into n_用药id From Dual;
    
      Insert Into 加症用药
        (用药id, 治法id, 草药id, 用量, 状态, 创建人, 创建时间, 最后修改人, 最后修改时间)
      Values
        (n_用药id, n_治法id, n_草药id, n_用量, 1, v_操作员名称, Sysdate, v_操作员名称, Sysdate);
    Elsif n_Usetype = 1 Then
      n_用药id := To_Number(Nvl(Jsonobj.Get_String('用药ID'), 0));
      n_治法id := To_Number(Nvl(Jsonobj.Get_String('治法ID'), 0));
      n_草药id := To_Number(Nvl(Jsonobj.Get_String('草药ID'), 0));
      n_用量   := To_Number(Nvl(Jsonobj.Get_String('用量'), 0));
      n_用药id := To_Number(Nvl(Jsonobj.Get_String('用药ID'), 0));
      Update 加症用药
      Set 治法id = n_治法id, 草药id = n_草药id, 用量 = n_用量, 最后修改人 = v_操作员名称, 最后修改时间 = Sysdate
      Where 用药id = n_用药id;
    Elsif n_Usetype = 2 Then
      n_状态   := To_Number(Nvl(Jsonobj.Get_String('状态'), 0));
      n_用药id := To_Number(Nvl(Jsonobj.Get_String('用药ID'), 0));
      Update 加症用药 Set 状态 = n_状态, 最后修改人 = v_操作员名称, 最后修改时间 = Sysdate Where 用药id = n_用药id;
    End If;
    Open Output_Out For
      Select n_用药id As 用药id From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Set_Zftozy;

  -----------------------------------------------------
  --项目删除处理
  -----------------------------------------------------
  Procedure Del_Zydata
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj   Pljson;
    n_Usetype Number;
    n_Id      Number;
    v_Name    Varchar2(50);
    v_Table   Varchar2(4000);
    v_Err_Msg Varchar2(255);
    Err_Item Exception;
  Begin
    --解析入参
    Jsonobj   := Pljson(Input_In);
    n_Usetype := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_Id      := To_Number(Nvl(Jsonobj.Get_String('ID'), 0));
  
    If n_Usetype = 0 Then
      --草药目录
      Select Nvl(Max(创建人), '空') Into v_Name From 草药目录 Where 草药id = n_Id;
      If v_Name = '系统创建' Then
        v_Err_Msg := '当前删除项目为系统创建项目,不能删除。';
        Raise Err_Item;
      Else
        Delete From 草药目录 Where 草药id = n_Id;
      End If;
    Elsif n_Usetype = 1 Then
      --中医疾病
      Select Nvl(Max(创建人), '空') Into v_Name From 中医疾病 Where 疾病id = n_Id;
      If v_Name = '系统创建' Then
        v_Err_Msg := '当前删除项目为系统创建项目,不能删除。';
        Raise Err_Item;
      Else
        Delete From 中医疾病 Where 疾病id = n_Id;
      End If;
    Elsif n_Usetype = 2 Then
      --中医证型
      Select Nvl(Max(创建人), '空') Into v_Name From 中医证型 Where 证型id = n_Id;
      If v_Name = '系统创建' Then
        v_Err_Msg := '当前删除项目为系统创建项目,不能删除。';
        Raise Err_Item;
      Else
        Delete From 中医证型 Where 证型id = n_Id;
      End If;
    Elsif n_Usetype = 3 Then
      --证型方剂对照
      Select Nvl(Max(创建人), '空') Into v_Name From 证型方剂对照 Where 对照id = n_Id;
      If v_Name = '系统创建' Then
        v_Err_Msg := '当前删除项目为系统创建项目,不能删除。';
        Raise Err_Item;
      Else
        Delete From 证型方剂对照 Where 对照id = n_Id;
      End If;
    Elsif n_Usetype = 4 Then
      --治法方剂
      Select Nvl(Max(创建人), '空') Into v_Name From 治法方剂 Where 方剂id = n_Id;
      If v_Name = '系统创建' Then
        v_Err_Msg := '当前删除项目为系统创建项目,不能删除。';
        Raise Err_Item;
      Else
        Delete From 治法方剂 Where 方剂id = n_Id;
      End If;
    Elsif n_Usetype = 5 Then
      --临证加症
      Select Nvl(Max(创建人), '空') Into v_Name From 临证加症 Where 加症id = n_Id;
      If v_Name = '系统创建' Then
        v_Err_Msg := '当前删除项目为系统创建项目,不能删除。';
        Raise Err_Item;
      Else
        Delete From 临证加症 Where 加症id = n_Id;
      End If;
    Elsif n_Usetype = 6 Then
      --加症治法
      Select Nvl(Max(创建人), '空') Into v_Name From 加症治法 Where 治法id = n_Id;
      If v_Name = '系统创建' Then
        v_Err_Msg := '当前删除项目为系统创建项目,不能删除。';
        Raise Err_Item;
      Else
        Delete From 加症治法 Where 治法id = n_Id;
      End If;
    Elsif n_Usetype = 7 Then
      --加症用药
      Select Nvl(Max(创建人), '空') Into v_Name From 加症用药 Where 用药id = n_Id;
      If v_Name = '系统创建' Then
        v_Err_Msg := '当前删除项目为系统创建项目,不能删除。';
        Raise Err_Item;
      Else
        Delete From 加症用药 Where 用药id = n_Id;
      End If;
    End If;
    Open Output_Out For
      Select '1' As 结果 From Dual;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
    When Others Then
      v_Table := SQLErrM;
      If SQLCode = -2292 Then
        Select Table_Name
        Into v_Table
        From All_Constraints
        Where Instr(v_Table, Owner || '.' || Constraint_Name) > 1 And Rownum < 2;
        v_Err_Msg := '[ZLSOFT]该记录在 ' || v_Table || ' 中已经使用,' || Chr(13) || '不能删除或修改[ZLSOFT]';
        Raise_Application_Error(-20005, v_Err_Msg);
      End If;
  End Del_Zydata;

  -----------------------------------------------------
  --获取数据库系统时间
  -----------------------------------------------------
  Procedure Get_Now_Time
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select Sysdate As 当前时间 From Dual;
  End Get_Now_Time;

  -----------------------------------------------------
  --异常处理
  -----------------------------------------------------
  Procedure Errorcenter
  (
    Err_Num In Number,
    Err_Msg In Varchar2
  ) Is
    v_Outnum Number := 0;
    v_Outmsg Varchar2(1000) := '';
    v_Count  Number;
    v_Temp   Varchar2(1000) := '';
  
    Cursor Cur_Ind_Cols Is
      Select Table_Name, Column_Name From All_Ind_Columns Where Instr(Err_Msg, Index_Owner || '.' || Index_Name) > 0;
  
    Cursor Cur_Con_Cols Is
      Select Table_Name, Column_Name
      From All_Cons_Columns
      Where (Owner, Constraint_Name) =
            (Select r_Owner, r_Constraint_Name
             From All_Constraints
             Where Instr(Err_Msg, Owner || '.' || Constraint_Name) > 0 And Rownum < 2);
  Begin
    If Err_Num = -1 Then
      For Row_Cols In Cur_Ind_Cols Loop
        v_Temp   := Row_Cols.Table_Name;
        v_Outmsg := v_Outmsg || '、' || Row_Cols.Column_Name;
      End Loop;
    
      v_Outmsg := '[ZLSOFT]' || v_Temp || '的(' || Substr(v_Outmsg, 2) || ')出现重复！[ZLSOFT]';
      v_Outnum := -20000;
    Elsif Err_Num = -1000 Then
      v_Outmsg := '[ZLSOFT]打开的数据表太多，必要时请系统管理员修改数据库的Open_Cursors配置。';
      v_Outnum := -20001;
    Elsif Err_Num = -1400 Or Err_Num = -1407 Then
      Select Table_Name, Column_Name
      Into v_Temp, v_Outmsg
      From All_Tab_Columns
      Where Instr(Err_Msg, '"' || Owner || '"."' || Table_Name || '"."' || Column_Name || '"') > 0 And Rownum < 2;
      v_Outmsg := '[ZLSOFT]' || v_Temp || '的(' || v_Outmsg || ')必须输入！[ZLSOFT]';
      v_Outnum := -20002;
    Elsif Err_Num = -1401 Then
      v_Outmsg := '[ZLSOFT]由于赋予的值超过了列宽限制，导致增加或更新失败。[ZLSOFT]';
      v_Outnum := -20003;
    Elsif Err_Num = -2290 Then
      Select Table_Name, Search_Condition
      Into v_Temp, v_Outmsg
      From All_Constraints
      Where Instr(Err_Msg, Owner || '.' || Constraint_Name) > 1 And Rownum < 2;
    
      If Instr(v_Outmsg, 'IS NOT NULL') > 0 Then
        v_Outmsg := '[ZLSOFT]' || v_Temp || ' 的 ' || Replace(v_Outmsg, 'IS NOT NULL', '必须输入！') || '[ZLSOFT]';
        v_Outnum := -20004;
      Else
        v_Outmsg := Err_Msg;
        v_Outnum := -20999;
      End If;
    Elsif Err_Num = -2292 Then
      Select Table_Name
      Into v_Temp
      From All_Constraints
      Where Instr(Err_Msg, Owner || '.' || Constraint_Name) > 1 And Rownum < 2;
    
      For Row_Cols In Cur_Con_Cols Loop
        v_Outmsg := v_Outmsg || '、' || Row_Cols.Column_Name;
      End Loop;
    
      v_Outmsg := '[ZLSOFT]该记录在 ' || v_Temp || ' 中已经使用,' || Chr(13) || '不能删除或修改(' || Substr(v_Outmsg, 2) || ')[ZLSOFT]';
      v_Outnum := -20005;
    Else
      v_Outmsg := Err_Msg;
      v_Outnum := -20999;
    End If;
  
    ------------------------
    --今后补充填写错误记录的代码
    ------------------------
    Raise_Application_Error(v_Outnum, Substr(v_Outmsg, 1, 100));
  End Errorcenter;

End Pkg_Zyedit;
/




----------------------------------------------------------------------------
--[[21.检查业务]]
----------------------------------------------------------------------------
--专业版RIS接口
CREATE OR REPLACE Package b_Zlxwinterface Is
  Type t_Refcur Is Ref Cursor;

  --1、接收RIS状态改变
  Procedure Receiverisstate
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    Risid_In    病人医嘱报告.Risid%Type,
    状态_In     Number,
    操作人员_In 病人医嘱发送.完成人%Type,
    执行时间_In 病人医嘱发送.完成时间%Type := Null,
    执行说明_In 病人医嘱发送.执行说明%Type := Null,
    单独执行_In Number := 0
  );

  --2、费用确认
  Procedure 影像费用执行
  (
    医嘱id_In     影像检查记录.医嘱id%Type,
    单独执行_In   Number := 0,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := Null
  );

  --3、取消费用确认
  Procedure 影像费用执行_Cancel
  (
    医嘱id_In     影像检查记录.医嘱id%Type,
    单独执行_In   Number := 0,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := Null
  );

  --4、接收RIS的报告
  Procedure Receivereport
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    Risid_In    病人医嘱报告.Risid%Type,
    报告所见_In 电子病历内容.内容文本%Type,
    报告意见_In 电子病历内容.内容文本%Type,
    报告建议_In 电子病历内容.内容文本%Type,
    报告医生_In 电子病历记录.创建人%Type
  );

  --5、修改申请单信息
  Procedure 影像病人信息_修改
  (
    医嘱id_In       病人医嘱记录.Id%Type,
    姓名_In         病人信息.姓名%Type,
    性别_In         病人信息.性别%Type,
    年龄_In         病人信息.年龄%Type,
    费别_In         病人信息.费别%Type,
    医疗付款方式_In 病人信息.医疗付款方式%Type,
    民族_In         病人信息.民族%Type,
    婚姻_In         病人信息.婚姻状况%Type,
    职业_In         病人信息.职业%Type,
    身份证号_In     病人信息.身份证号%Type,
    家庭地址_In     病人信息.家庭地址%Type,
    家庭电话_In     病人信息.家庭电话%Type,
    家庭地址邮编_In 病人信息.家庭地址邮编%Type,
    出生日期_In     病人信息.出生日期%Type := Null
  );

  --6、取消申请单信息
  Procedure 取消检查申请单
  (
    医嘱id_In     病人医嘱执行.医嘱id%Type,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := 0,
    拒绝原因_In   病人医嘱发送.执行说明%Type := Null
  );

  --7、插入医嘱操作失败记录
  Procedure Ris医嘱失败记录_Insert
  (
    医嘱ID_In   In Ris医嘱失败记录.医嘱id%Type,
    发送类型_In In Ris医嘱失败记录.发送类型%Type
  );

  --8、更新医嘱操作失败记录
  Procedure Ris医嘱失败记录_重发
  (
    Id_In       In Ris医嘱失败记录.Id%Type,
    操作类型_In In Number
  );

  --9、销账后新建住院记账单据
  Procedure 病人医嘱_重建单据
  (
    医嘱id_In In 病人医嘱发送.医嘱id%Type,
    No_In     In 病人医嘱发送.No%Type,
    Action_In In Number
  );

  --10、打印RIS检查预约通知单
  Procedure Ris检查预约_打印(医嘱id_In In Ris检查预约.医嘱id%Type);

  --11、更新RIS分科室启用参数
  Procedure Ris启用控制_Update
  (
    检查类型_In Ris启用控制.检查类型%Type,
    场合_In     Ris启用控制.场合%Type,
    部门ids_In  Varchar2,
    启用类型_In Number
  );

  --12、删除RIS分科室启用参数
  Procedure Ris启用控制_Delete;

  --13、根据元素名提取信息
  Function Ris_Replace_Element_Value
  (
    元素名_In   In 诊治所见项目.中文名%Type,
    病人id_In   In 电子病历记录.病人id%Type,
    就诊id_In   In 电子病历记录.主页id%Type,
    病人来源_In In 电子病历记录.病人来源%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type
  ) Return Varchar2;

  --14、删除RIS分院设置参数
  Procedure Ris分院设置_Delete;

  --15、更新RISRis分院设置参数
  Procedure Ris分院设置_Update
  (
    Id_In           Ris分院设置.Id%Type,
    医院名称_In     Ris分院设置.医院名称%Type,
    医院代码_In     Ris分院设置.医院代码%Type,
    用户名_In       Ris分院设置.用户名%Type,
    密码_In         Ris分院设置.密码%Type,
    数据库服务名_In Ris分院设置.数据库服务名%Type
  );

  --16、登记危急值
  Procedure 病人危急值记录_Insert
  (
    Id_In         In 病人危急值记录.Id%Type,
    数据来源_In   In 病人危急值记录.数据来源%Type,
    病人id_In     In 病人危急值记录.病人id%Type,
    主页id_In     In 病人危急值记录.主页id%Type,
    挂号单_In     In 病人危急值记录.挂号单%Type,
    婴儿_In       In 病人危急值记录.婴儿%Type,
    姓名_In       In 病人危急值记录.姓名%Type,
    性别_In       In 病人危急值记录.性别%Type,
    年龄_In       In 病人危急值记录.年龄%Type,
    医嘱id_In     In 病人危急值记录.医嘱id%Type,
    标本id_In     In 病人危急值记录.标本id%Type,
    危急值描述_In In 病人危急值记录.危急值描述%Type,
    报告时间_In   In 病人危急值记录.报告时间%Type,
    报告科室id_In In 病人危急值记录.报告科室id%Type,
    报告人_In     In 病人危急值记录.报告人%Type
  );

  --17、取消危急值
  Procedure 病人危急值记录_Delete(医嘱id_In In 病人危急值记录.医嘱id%Type);

  --18、发送临床医嘱
  Function 病人医嘱记录_Send(医嘱id_In In 病人医嘱发送.医嘱id%Type) Return Varchar2;

End b_Zlxwinterface;
/

CREATE OR REPLACE Package Body b_Zlxwinterface Is

  --1、接收RIS状态改变
  Procedure Receiverisstate
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    Risid_In    病人医嘱报告.Risid%Type,
    状态_In     Number,
    操作人员_In 病人医嘱发送.完成人%Type,
    执行时间_In 病人医嘱发送.完成时间%Type := Null,
    执行说明_In 病人医嘱发送.执行说明%Type := Null,
    单独执行_In Number := 0
  ) Is
  
    --参数：医嘱ID_IN - 单独执行的医嘱ID。
    --      状态_IN - -1-删除；0-预约；1-登记；3-检查完成；4-检查中止；9-初步报告；12-报告审核；15-发放
    --     单独执行_In -0-全部执行；1-单独执行；检查医嘱组合是否采用对每个项目分散单独执行的方式
  
    Cursor c_Adviceinfo Is
      Select a.Id, a.相关id, Nvl(a.相关id, a.Id) As 组id, a.诊疗类别, a.病人来源, a.执行科室id, b.执行过程
      From 病人医嘱记录 A, 病人医嘱发送 B
      Where a.Id = b.医嘱id And ID = 医嘱id_In;
    r_Adviceinfo c_Adviceinfo%RowType;
  
    v_执行状态 病人医嘱发送.执行状态%Type;
    v_执行过程 病人医嘱发送.执行过程%Type;
    n_执行     Number; --标记是否需要更新状态，1：需要更新，其他不需要更新
    v_Count    Number;
    v_完成人   病人医嘱发送.完成人%Type;
    v_完成时间 病人医嘱发送.完成时间%Type;
    v_Error    Varchar2(255);
    Err_Custom Exception;
  
  Begin
  
    v_执行状态 := 0;
    v_执行过程 := 0;
  
    --提取医嘱的主医嘱ID，及组ID
    Open c_Adviceinfo;
    Fetch c_Adviceinfo
      Into r_Adviceinfo;
    Close c_Adviceinfo;
  
    --根据状态_IN执行医嘱
    ---1-删除；0-预约(在RIS中实际上就是删除)；1-登记；3-检查完成；4-检查中止；9-初步报告；12-报告审核；13-取消审核；14-报告删除；15-发放
  
    If 状态_In = -1 Or 状态_In = 0 Then
      v_执行状态 := 0; --未执行
      v_执行过程 := 0;
    Elsif 状态_In = 1 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 2; --已报到
    Elsif 状态_In = 3 Or 状态_In = 14 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 3; --已检查
    Elsif 状态_In = 4 Then
      --不改变
      v_执行状态 := v_执行状态;
    Elsif 状态_In = 9 Or 状态_In = 13 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 4; --已报告
    Elsif 状态_In = 12 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 5; --已审核
    Elsif 状态_In = 15 Then
      v_执行状态 := 1; --完全执行
      v_执行过程 := 6; --已完成
      v_完成人   := 操作人员_In;
      v_完成时间 := 执行时间_In;
    End If;
  
    n_执行 := 1; --默认都要更新状态
  
    If 状态_In = 13 Or 状态_In = 14 Then
      --删除对应报告数据
      Delete From 电子病历记录
      Where ID = (Select 病历id From 病人医嘱报告 Where 医嘱id = 医嘱id_In And Risid = Risid_In);
      Delete From 病人医嘱报告 Where 医嘱id = 医嘱id_In And Risid = Risid_In;
    
      --删除后判断是否还存在报告，若存在则医嘱状态保持不变，若报告全部删除则更新医嘱状态
      Select Count(1) Into v_Count From 病人医嘱报告 Where 医嘱id = 医嘱id_In;
    
      If v_Count > 0 Then
        n_执行 := 0; --若存在则医嘱状态保持不变
      End If;
    End If;
  
    --如果是删除，则删除已有的预约信息
    If 状态_In = -1 Or 状态_In = 0 Then
      Zl_Ris检查预约_Delete(医嘱id_In);
    End If;
  
    --如果是登记，先判断此检查是否未执行
    If 状态_In = 1 Then
      If r_Adviceinfo.执行过程 >= 3 Then
        v_Error := '患者已经做过检查了，不能重复登记。';
        Raise Err_Custom;
      End If;
    End If;
  
    --开始执行医嘱
    If n_执行 = 1 Then
      If Nvl(单独执行_In, 0) = 1 Then
        -- 单个部位医嘱单独执行
        Update 病人医嘱发送
        Set 执行状态 = v_执行状态, 执行过程 = v_执行过程, 执行说明 = 执行说明_In, 完成人 = v_完成人, 完成时间 = v_完成时间
        Where 医嘱id = 医嘱id_In;
      Else
        Update 病人医嘱发送
        Set 执行状态 = v_执行状态, 执行过程 = v_执行过程, 执行说明 = 执行说明_In, 完成人 = v_完成人, 完成时间 = v_完成时间
        Where 医嘱id In (Select ID From 病人医嘱记录 Where (ID = r_Adviceinfo.组id Or 相关id = r_Adviceinfo.组id));
      End If;
    End If;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Receiverisstate;

  --2、费用确认
  Procedure 影像费用执行
  (
    医嘱id_In     影像检查记录.医嘱id%Type,
    单独执行_In   Number := 0,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := Null
  ) Is
    --参数：医嘱ID_IN=单独执行的医嘱ID。
    --      单独执行_In=检查医嘱组合是否采用对每个项目分散单独执行的方式,0-不单独执行
    Cursor c_Advice Is
      Select ID, 相关id, Nvl(相关id, ID) As 组id, 诊疗类别, 病人来源 From 病人医嘱记录 Where ID = 医嘱id_In;
    r_Advice c_Advice%RowType;
  
    v_Temp     Varchar2(255);
    v_人员编号 人员表.编号%Type;
    v_人员姓名 人员表.姓名%Type;
    v_部门id   部门表.Id%Type;
    v_费用性质 病人医嘱发送.记录性质%Type;
    v_发送号   病人医嘱发送.发送号%Type;
    v_执行过程 病人医嘱发送.执行过程%Type;
    v_Error    Varchar2(255);
    Err_Custom Exception;
  Begin
  
    --取主医嘱ID
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
    Close c_Advice;
  
    Select 发送号, 执行过程 Into v_发送号, v_执行过程 From 病人医嘱发送 Where 医嘱id = r_Advice.组id;
  
    --登记和完成才执行费用  2-登记，3-检查，4-报告，5-审核，6-完成
    If v_执行过程 >= 2 Or v_执行过程 <= 6 Then
      --取当前操作人员
      If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null And 执行部门id_In Is Not Null Then
        v_人员编号 := 操作员编号_In;
        v_人员姓名 := 操作员姓名_In;
        v_部门id   := 执行部门id_In;
      Else
        v_Temp     := Zl_Identity;
        v_部门id   := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      End If;
    
      If r_Advice.病人来源 = 2 Then
        Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
        Into v_费用性质
        From 病人医嘱发送
        Where 发送号 = v_发送号 And 医嘱id = 医嘱id_In;
      Else
        v_费用性质 := 1;
      End If;
    
      --执行费用和自动发料
      If v_费用性质 = 1 Then
        Zl_门诊医嘱执行_Finish(医嘱id_In, v_发送号, 单独执行_In, v_人员编号, v_人员姓名, r_Advice.组id, r_Advice.诊疗类别, v_部门id);
      Else
        Zl_住院医嘱执行_Finish(医嘱id_In, v_发送号, 单独执行_In, v_人员编号, v_人员姓名, r_Advice.组id, r_Advice.诊疗类别, v_部门id);
      End If;
    End If;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 影像费用执行;

  --3、取消费用确认
  Procedure 影像费用执行_Cancel
  (
    医嘱id_In     影像检查记录.医嘱id%Type,
    单独执行_In   Number := 0,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := Null
  ) Is
    --参数：
    --      医嘱ID_IN=单独执行的医嘱ID。
    --      单独执行_In=检查医嘱组合是否采用对每个项目分散单独执行的方式,0-不单独执行
  
    Cursor c_Advice Is
      Select ID, 相关id, Nvl(相关id, ID) As 组id From 病人医嘱记录 Where ID = 医嘱id_In;
    r_Advice c_Advice%RowType;
  
    v_发送号 病人医嘱发送.发送号%Type;
    v_Count  Number;
    v_Error  Varchar2(255);
    Err_Custom Exception;
  
  Begin
  
    --取主医嘱ID
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
    Close c_Advice;
  
    --先检查是否已经出院的住院病人，已经预出院或者出院的检查申请，不允执行费用
    Select Count(*)
    Into v_Count
    From 病人医嘱记录 A, 病案主页 B
    Where a.病人id = b.病人id And a.主页id = b.主页id And (b.出院日期 Is Not Null Or b.状态 = 3) And a.Id = r_Advice.组id;
  
    If v_Count > 0 Then
      v_Error := '住院病人已经出院或者预出院，不能取消费用。';
      Raise Err_Custom;
    End If;
  
    Select 发送号 Into v_发送号 From 病人医嘱发送 Where 医嘱id = r_Advice.组id;
  
    --调用统一的医嘱执行Cancel过程
    Zl_病人医嘱执行_Cancel(医嘱id_In, v_发送号, Null, 单独执行_In, 执行部门id_In, 操作员编号_In, 操作员姓名_In);
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 影像费用执行_Cancel;

  --4、接收RIS的报告
  Procedure Receivereport
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    Risid_In    病人医嘱报告.Risid%Type,
    报告所见_In 电子病历内容.内容文本%Type,
    报告意见_In 电子病历内容.内容文本%Type,
    报告建议_In 电子病历内容.内容文本%Type,
    报告医生_In 电子病历记录.创建人%Type
  ) Is
    --提取病人医嘱及报告的相关信息
    Cursor c_Advice
    (
      v_组id  Number,
      v_Risid Number
    ) Is
      Select e.Id, e.病人来源, e.病人id, e.主页id, e.婴儿, e.病人科室id, e.文件id, e.病历种类, e.病历名称, f.病历id, e.执行科室id
      From (Select c.Id, c.病人来源, c.病人id, c.主页id, c.婴儿, c.病人科室id, c.文件id, d.种类 病历种类, d.名称 病历名称, c.执行科室id
             From (Select a.Id, a.病人来源, a.病人id, a.主页id, a.婴儿, a.病人科室id, b.病历文件id 文件id, a.执行科室id
                    From 病人医嘱记录 A, 病历单据应用 B
                    Where a.Id = v_组id And a.诊疗项目id = b.诊疗项目id(+) And b.应用场合(+) = Decode(a.病人来源, 2, 2, 4, 4, 1)) C,
                  病历文件列表 D
             Where c.文件id = d.Id(+)) E, 病人医嘱报告 F
      Where e.Id = f.医嘱id(+) And f.Risid(+) = v_Risid;
  
    --查找文件的组成元素
    Cursor c_File(v_File Number) Is
      Select a.Id, a.文件id, a.父id, a.对象序号, a.对象类型, a.对象标记, a.保留对象, a.对象属性, a.内容行次, a.内容文本, a.是否换行, a.预制提纲id, a.复用提纲,
             a.使用时机, a.诊治要素id, a.替换域, a.要素名称, a.要素类型, a.要素长度, a.要素小数, a.要素单位, a.要素表示, a.输入形态, a.要素值域
      From 病历文件结构 A
      Where a.文件id = v_File
      Order By a.对象序号;
  
    Cursor c_Report(v_电子病历记录id Number) Is
      Select b.Id, a.内容文本
      From 电子病历内容 A, 电子病历内容 B
      Where a.对象类型 = 3 And a.Id = b.父id And b.对象类型 = 2 And b.终止版 = 0 And a.文件id = v_电子病历记录id;
  
    Cursor c_Content
    (
      v_文件id Number,
      v_表格id Number
    ) Is
      Select a.Id, a.文件id, a.父id, a.对象序号, a.对象类型, a.对象标记, a.保留对象, a.对象属性, a.内容行次, a.内容文本, a.是否换行, a.预制提纲id, a.复用提纲,
             a.使用时机, a.诊治要素id, a.替换域, a.要素名称, a.要素类型, a.要素长度, a.要素小数, a.要素单位, a.要素表示, a.输入形态, a.要素值域
      From 病历文件结构 A
      Where 文件id = v_文件id And 父id = v_表格id;
  
    r_Advice        c_Advice%RowType;
    v_病历id        电子病历内容.文件id%Type;
    v_病历内容id    电子病历内容.Id%Type;
    v_病历内容idnew 电子病历内容.Id%Type;
    v_对象序号      电子病历内容.对象序号%Type;
    v_父id          电子病历内容.父id%Type;
    v_内容文本      电子病历内容.内容文本%Type;
    v_定义提纲id    电子病历内容.定义提纲id%Type;
    --v_格式内容    电子病历格式.内容%Type;
    v_Error Varchar2(255);
    Err_Custom Exception;
    v_主医嘱id 病人医嘱发送.医嘱id%Type;
    v_表格     Varchar2(300);
    n_数量     Number;
    n_Rptcount Number;
    v_病历名称 电子病历记录.病历名称%Type;
    v_挂号单id 病人挂号记录.Id%Type;
  
    Function Getrptno
    (
      v_医嘱idin   病人医嘱发送.医嘱id%Type,
      v_病历名称in 电子病历记录.病历名称%Type
    ) Return Varchar As
      v_Return Number;
      v_No     Number;
      v_Count  Number;
    Begin
      Select Count(医嘱id) + 1 Into v_No From 病人医嘱报告 Where 医嘱id = v_医嘱idin;
      v_Count := 1;
      While v_Count = 1 Loop
        Select Count(ID)
        Into v_Count
        From 病人医嘱报告 A, 电子病历记录 B
        Where a.医嘱id = v_医嘱idin And a.病历id = b.Id And b.病历名称 = v_病历名称in || v_No;
        If v_Count = 1 Then
          v_No := v_No + 1;
        End If;
      End Loop;
      v_Return := v_No;
      Return v_Return;
    End Getrptno;
  
  Begin
  
    -- 提取主医嘱ID ，防止因为传入部位医嘱，导致报告保存出错
    Select Nvl(相关id, ID) As 组id Into v_主医嘱id From 病人医嘱记录 Where ID = 医嘱id_In;
  
    Open c_Advice(v_主医嘱id, Nvl(Risid_In, 0));
    Fetch c_Advice
      Into r_Advice;
  
    If Nvl(r_Advice.文件id, 0) = 0 Then
      v_Error := '本次检查项目没有对应相关的检查报告，请与管理员联系！';
      Raise Err_Custom;
    Else
      If Nvl(r_Advice.病历id, 0) > 0 Then
        ----产生过报告
        --找出检查已填写的报告提纲中含有"%所见%","%描述%","%建议%","%意见%",并用传入的参数更新
        For r_Report In c_Report(r_Advice.病历id) Loop
          If r_Report.内容文本 Like '%所见%' Then
            Update 电子病历内容 Set 内容文本 = 报告所见_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          Elsif r_Report.内容文本 Like '%意见%' Then
            Update 电子病历内容 Set 内容文本 = 报告意见_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          Elsif r_Report.内容文本 Like '%建议%' Then
            Update 电子病历内容 Set 内容文本 = 报告建议_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          End If;
        End Loop;
        --更新保存时间
        Update 电子病历记录
        Set 完成时间 = Sysdate, 保存人 = 报告医生_In, 保存时间 = Sysdate
        Where ID = r_Advice.病历id;
      Else
        --先判断单据中是否有对应的提纲和表格
        If Nvl(报告所见_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_数量
          From 病历文件结构 A, 病历文件结构 B
          Where a.父id = b.Id And a.对象类型 = 3 And b.对象类型 = 1 And a.内容文本 Like '%所见%' And a.文件id = r_Advice.文件id;
        
          If n_数量 <= 0 Then
            v_Error := '在诊疗单据中没有找到【所见】对应的提纲或表格，请联系管理员设置！';
            Raise Err_Custom;
          End If;
        End If;
        If Nvl(报告意见_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_数量
          From 病历文件结构 A, 病历文件结构 B
          Where a.父id = b.Id And a.对象类型 = 3 And b.对象类型 = 1 And a.内容文本 Like '%意见%' And a.文件id = r_Advice.文件id;
        
          If n_数量 <= 0 Then
            v_Error := '在诊疗单据中没有找到【意见】对应的提纲或表格，请联系管理员设置！';
            Raise Err_Custom;
          End If;
        End If;
        If Nvl(报告建议_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_数量
          From 病历文件结构 A, 病历文件结构 B
          Where a.父id = b.Id And a.对象类型 = 3 And b.对象类型 = 1 And a.内容文本 Like '%建议%' And a.文件id = r_Advice.文件id;
        
          If n_数量 <= 0 Then
            v_Error := '在诊疗单据中没有找到【建议】对应的提纲或表格，请联系管理员设置！';
            Raise Err_Custom;
          End If;
        End If;
      
        If r_Advice.病人来源 = 1 Then
          --门诊，提取挂号单ID
          Select nvl(Max(c.Id), 0)
          Into v_挂号单id
          From 病人医嘱记录 B, 病人挂号记录 C
          Where b.挂号单 = c.No(+) And c.记录状态 In (1, 3) And b.Id = v_主医嘱id;
        Else
          --体检或者外诊，无挂号单ID，直接设置为0
          v_挂号单id := 0;
        End If;
      
        --产生电子病历记录
        Select 电子病历记录_Id.Nextval Into v_病历id From Dual;
        n_Rptcount := Getrptno(医嘱id_In, r_Advice.病历名称);
        If n_Rptcount > 1 Then
          v_病历名称 := r_Advice.病历名称 || n_Rptcount;
        Else
          v_病历名称 := r_Advice.病历名称;
        End If;
        Insert Into 电子病历记录
          (ID, 病人来源, 病人id, 主页id, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 创建人, 创建时间, 完成时间, 保存人, 保存时间, 最后版本, 签名级别)
        Values
          (v_病历id, r_Advice.病人来源, r_Advice.病人id, Decode(r_Advice.病人来源, 2, r_Advice.主页id, v_挂号单id), r_Advice.婴儿,
           r_Advice.病人科室id, r_Advice.病历种类, r_Advice.文件id, v_病历名称, 报告医生_In, Sysdate, Sysdate, 报告医生_In, Sysdate, 1, 2);
      
        --产生医嘱报告记录
        Insert Into 病人医嘱报告 (医嘱id, 病历id, Risid) Values (v_主医嘱id, v_病历id, Risid_In);
      
        v_对象序号 := 0;
      
        --新产生报告内容
        For r_File In c_File(r_Advice.文件id) Loop
          Select 电子病历内容_Id.Nextval Into v_病历内容id From Dual;
          v_内容文本   := r_File.内容文本;
          v_定义提纲id := 0;
        
          If Nvl(r_File.对象类型, 0) = 1 And Nvl(r_File.父id, 0) = 0 Then
            --提纲
            v_定义提纲id := r_File.Id;
            v_父id       := v_病历内容id;
          End If;
        
          If Nvl(r_File.对象类型, 0) = 4 And r_File.要素名称 Is Not Null Then
            --元素
            v_内容文本 := Zl_Replace_Element_Value(r_File.要素名称, r_Advice.病人id, r_Advice.主页id, r_Advice.病人来源, r_Advice.Id);
          End If;
        
          If Nvl(r_File.父id, 0) <> 0 Then
            v_定义提纲id := 0;
          End If;
        
          v_对象序号 := v_对象序号 + 1;
        
          If Instr(v_表格, '|' || r_File.父id || '|') > 0 Then
            Null;
          Else
            Insert Into 电子病历内容
              (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机, 诊治要素id, 替换域,
               要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域, 定义提纲id)
            Values
              (v_病历内容id, v_病历id, 1, 0, Decode(v_定义提纲id, 0, v_父id, Null), v_对象序号, r_File.对象类型, r_File.对象标记, r_File.保留对象,
               r_File.对象属性, Null, v_内容文本, r_File.是否换行, r_File.预制提纲id, r_File.复用提纲, r_File.使用时机, r_File.诊治要素id,
               r_File.替换域, r_File.要素名称, r_File.要素类型, r_File.要素长度, r_File.要素小数, r_File.要素单位, r_File.要素表示, r_File.输入形态,
               r_File.要素值域, Decode(v_定义提纲id, 0, Null, v_定义提纲id));
          End If;
        
          --为表格时，插入文本内容
          If Nvl(r_File.对象类型, 0) = 3 And Nvl(r_File.父id, 0) <> 0 Then
            v_表格 := v_表格 || ',|' || r_File.Id || '|';
          
            If r_File.内容文本 Like '%所见%' Then
              v_内容文本 := 报告所见_In || Chr(13) || Chr(13);
            Elsif r_File.内容文本 Like '%意见%' Then
              v_内容文本 := 报告意见_In || Chr(13) || Chr(13);
            Else
              v_内容文本 := 报告建议_In || Chr(13) || Chr(13);
            End If;
          
            For r_Con In c_Content(r_Advice.文件id, r_File.Id) Loop
              Select 电子病历内容_Id.Nextval Into v_病历内容idnew From Dual;
              v_对象序号 := v_对象序号 + 1;
            
              Insert Into 电子病历内容
                (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机, 诊治要素id,
                 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域, 定义提纲id)
              Values
                (v_病历内容idnew, v_病历id, 1, 0, v_病历内容id, v_对象序号, 2, r_Con.对象标记, r_Con.保留对象, r_Con.对象属性, Null, v_内容文本,
                 r_Con.是否换行, r_Con.预制提纲id, r_Con.复用提纲, r_Con.使用时机, r_Con.诊治要素id, r_Con.替换域, r_Con.要素名称, r_Con.要素类型,
                 r_Con.要素长度, r_Con.要素小数, r_Con.要素单位, r_Con.要素表示, r_Con.输入形态, r_Con.要素值域,
                 Decode(v_定义提纲id, 0, Null, v_定义提纲id));
            End Loop;
          End If;
        End Loop;
      
        --因电子病历格式中含了内容文字格式，此种方法导入之后内容文字将不可见
        --Select 内容 Into v_格式内容 From 病历文件格式 Where 文件ID=r_Advice.文件ID;
        --Insert Into 电子病历格式 (文件ID,内容) Values (v_病历id,v_格式内容);
      
      End If;
    End If;
    Close c_Advice;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Receivereport;

  --5、修改申请单信息
  Procedure 影像病人信息_修改
  (
    医嘱id_In       病人医嘱记录.Id%Type,
    姓名_In         病人信息.姓名%Type,
    性别_In         病人信息.性别%Type,
    年龄_In         病人信息.年龄%Type,
    费别_In         病人信息.费别%Type,
    医疗付款方式_In 病人信息.医疗付款方式%Type,
    民族_In         病人信息.民族%Type,
    婚姻_In         病人信息.婚姻状况%Type,
    职业_In         病人信息.职业%Type,
    身份证号_In     病人信息.身份证号%Type,
    家庭地址_In     病人信息.家庭地址%Type,
    家庭电话_In     病人信息.家庭电话%Type,
    家庭地址邮编_In 病人信息.家庭地址邮编%Type,
    出生日期_In     病人信息.出生日期%Type := Null
  ) As
  
    v_年龄        Varchar2(20);
    v_年龄单位    Varchar2(20);
    v_出生日期    Date;
    v_病人来源    病人医嘱记录.病人来源%Type;
    v_病人id      病人医嘱记录.病人id%Type;
    v_Strtmpbefor Varchar2(4000);
    v_Msg         Varchar2(4000);
  Begin
    Begin
      Select 病人来源, 病人id Into v_病人来源, v_病人id From 病人医嘱记录 Where ID = 医嘱id_In;
    Exception
      When Others Then
        Return;
    End;
  
    If 出生日期_In Is Null And 年龄_In Is Not Null Then
      --根据年龄求出生日期
      v_年龄单位 := Substr(年龄_In, Length(年龄_In), 1);
      If Instr('岁,月,天', v_年龄单位) <= 0 Then
        v_年龄单位 := Null;
      Else
        v_年龄 := Replace(年龄_In, v_年龄单位, '');
      End If;
      Begin
        v_年龄 := To_Number(v_年龄);
      Exception
        When Others Then
          v_年龄 := Null;
      End;
      If v_年龄 Is Not Null And v_年龄单位 Is Not Null Then
        Select Decode(v_年龄单位, '岁', Add_Months(Sysdate, -12 * v_年龄), '月', Add_Months(Sysdate, -1 * v_年龄), '天',
                       Sysdate - v_年龄)
        Into v_出生日期
        From Dual;
      End If;
    Else
      v_出生日期 := 出生日期_In;
    End If;
    Select Zl_Fun_Checkidentify(0, v_病人id, v_Strtmpbefor) Into v_Strtmpbefor From Dual;
    If v_病人来源 = 3 Then
      Update 病人信息
      Set 姓名 = 姓名_In, 性别 = Nvl(性别_In, 性别), 年龄 = 年龄_In, 出生日期 = v_出生日期, 费别 = Nvl(费别_In, 费别),
          医疗付款方式 = Nvl(医疗付款方式_In, 医疗付款方式), 民族 = Nvl(民族_In, 民族), 婚姻状况 = Nvl(婚姻_In, 婚姻状况), 职业 = Nvl(职业_In, 职业),
          身份证号 = 身份证号_In, 家庭地址 = 家庭地址_In, 家庭电话 = 家庭电话_In, 家庭地址邮编 = 家庭地址邮编_In
      Where 病人id = v_病人id;
      --修改对应的医嘱记录
      Update 病人医嘱记录
      Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In
      Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
    Else
      Update 病人信息
      Set 民族 = Nvl(民族_In, 民族), 婚姻状况 = Nvl(婚姻_In, 婚姻状况), 职业 = Nvl(职业_In, 职业), 家庭地址 = 家庭地址_In, 家庭电话 = 家庭电话_In,
          家庭地址邮编 = 家庭地址邮编_In
      Where 病人id = v_病人id;
    End If;
    Select Zl_Fun_Checkidentify(1, v_病人id, v_Strtmpbefor) Into v_Msg From Dual;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 影像病人信息_修改;

  --6、取消申请单信息
  Procedure 取消检查申请单
  (
    医嘱id_In     病人医嘱执行.医嘱id%Type,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := 0,
    拒绝原因_In   病人医嘱发送.执行说明%Type := Null
  ) As
    --参数：医嘱ID_IN=单独执行的医嘱ID
  
    v_发送号 病人医嘱执行.发送号%Type;
  
  Begin
  
    Begin
      Select 发送号 Into v_发送号 From 病人医嘱发送 Where 医嘱id = 医嘱id_In;
    Exception
      When Others Then
        Return;
    End;
  
    Zl_病人医嘱执行_拒绝执行(医嘱id_In, v_发送号, 操作员编号_In, 操作员姓名_In, 执行部门id_In, 拒绝原因_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 取消检查申请单;

  --7、插入医嘱操作失败记录
  Procedure Ris医嘱失败记录_Insert
  (
    医嘱ID_In   In Ris医嘱失败记录.医嘱id%Type,
    发送类型_In In Ris医嘱失败记录.发送类型%Type
  ) Is
  Begin
    Insert Into Ris医嘱失败记录
      (ID, 医嘱ID, 发送类型, 发送时间, 重发次数)
    Values
      (Ris医嘱失败记录_Id.Nextval, 医嘱ID_In, 发送类型_In, Sysdate, 0);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris医嘱失败记录_Insert;

  --8、更新医嘱操作失败记录
  Procedure Ris医嘱失败记录_重发
  (
    Id_In       In Ris医嘱失败记录.Id%Type,
    操作类型_In In Number
  ) Is
    v_重发次数 Ris医嘱失败记录.重发次数%Type;
  Begin
    --操作类型_In -- 1 重发成功，删除记录；2--重发失败
  
    If 操作类型_In = 1 Then
      Delete From Ris医嘱失败记录 Where ID = Id_In;
    Else
      Select 重发次数 Into v_重发次数 From Ris医嘱失败记录 Where ID = Id_In;
      If v_重发次数 >= 99 Then
        v_重发次数 := 99;
      Else
        v_重发次数 := v_重发次数 + 1;
      End If;
      Update Ris医嘱失败记录 Set 发送时间 = Sysdate, 重发次数 = v_重发次数 Where ID = Id_In;
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris医嘱失败记录_重发;

  --9、销账后新建住院记账单据
  Procedure 病人医嘱_重建单据
  (
    医嘱id_In In 病人医嘱发送.医嘱id%Type,
    No_In     In 病人医嘱发送.No%Type,
    Action_In In Number
  ) Is
    -- Action_In: 1 重建单据；2 取消重建单据
    v_No 病人医嘱发送.No%Type;
  Begin
    If Action_In = 1 Then
      Select Nextno(14) Into v_No From Dual;
    
      Update 病人医嘱发送
      Set NO = v_No, 计费状态 = 0
      Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In);
      Update 住院费用记录 Set 医嘱序号 = Null Where NO = No_In;
    Elsif Action_In = 2 Then
      Update 住院费用记录 Set 医嘱序号 = 医嘱id_In Where NO = No_In;
      Update 病人医嘱发送
      Set NO = No_In, 计费状态 = 4
      Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In);
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 病人医嘱_重建单据;

  --10、打印RIS检查预约通知单
  Procedure Ris检查预约_打印(医嘱id_In In Ris检查预约.医嘱id%Type) Is
    v_Temp     Varchar2(255);
    v_人员姓名 人员表.姓名%Type;
  Begin
    --取当前操作人员
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  
    Update Ris检查预约 Set 是否打印 = 1, 打印人 = v_人员姓名, 打印时间 = Sysdate Where 医嘱id = 医嘱id_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris检查预约_打印;

  --11、更新RIS分科室启用参数
  Procedure Ris启用控制_Update
  (
    检查类型_In Ris启用控制.检查类型%Type,
    场合_In     Ris启用控制.场合%Type,
    部门ids_In  Varchar2,
    启用类型_In Number
  ) Is
  
    l_部门id   t_Numlist := t_Numlist();
    v_启用ris  Ris启用控制.是否启用ris%Type;
    v_启用预约 Ris启用控制.是否启用预约%Type;
  
    Cursor c_Dept(Dept_In Varchar2) Is
      Select Column_Value From Table(f_Num2list(Dept_In));
  Begin
  
    If 启用类型_In = 1 Then
      v_启用ris  := 1;
      v_启用预约 := Null;
      Delete From Ris启用控制 Where 检查类型 = 检查类型_In And 场合 = 场合_In And 是否启用ris = 1;
    Else
      v_启用ris  := Null;
      v_启用预约 := 1;
      Delete From Ris启用控制 Where 检查类型 = 检查类型_In And 场合 = 场合_In And 是否启用预约 = 1;
    End If;
  
    If 部门ids_In Is Null Then
      Insert Into Ris启用控制
        (ID, 检查类型, 场合, 部门id, 是否启用ris, 是否启用预约)
      Values
        (Ris启用控制_Id.Nextval, 检查类型_In, 场合_In, Null, v_启用ris, v_启用预约);
    Else
      Open c_Dept(部门ids_In);
      Fetch c_Dept Bulk Collect
        Into l_部门id;
      Close c_Dept;
    
      Forall I In 1 .. l_部门id.Count
        Insert Into Ris启用控制
          (ID, 检查类型, 场合, 部门id, 是否启用ris, 是否启用预约)
        Values
          (Ris启用控制_Id.Nextval, 检查类型_In, 场合_In, l_部门id(I), v_启用ris, v_启用预约);
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris启用控制_Update;

  --12、删除RIS分科室启用参数
  Procedure Ris启用控制_Delete Is
  
  Begin
    Delete From Ris启用控制;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris启用控制_Delete;

  --13、根据元素名提取信息
  Function Ris_Replace_Element_Value
  (
    元素名_In   In 诊治所见项目.中文名%Type,
    病人id_In   In 电子病历记录.病人id%Type,
    就诊id_In   In 电子病历记录.主页id%Type,
    病人来源_In In 电子病历记录.病人来源%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type
  ) Return Varchar2 Is
    v_Return Varchar2(4000) := Null;
    Cursor c_Patient Is
      Select 姓名, 性别, Decode(性别, '男', 'M', '女', 'F', 'O') As 性别编码, 出生日期, 病人id, 联系人地址, 家庭电话, 联系人电话, 婚姻状况, 身份证号, 当前科室id,
             当前病区id, 当前床号 As 床号, 就诊卡号, 入院时间, 出院时间
      From 病人信息
      Where 病人id = 病人id_In;
    r_Patient c_Patient%RowType;
  
    Cursor c_Order Is
      Select 主页id, 婴儿, Decode(病人来源, 1, 'OUTPAT', 2, 'INPAT', 'UNK') As 病人来源, 开嘱医生, 开嘱时间, 校对护士, 医嘱内容, 紧急标志, 执行科室id
      From 病人医嘱记录
      Where ID = 医嘱id_In;
    r_Order c_Order%RowType;
  
    Cursor c_Diagnose Is
      Select 诊断描述 || Decode(Nvl(是否疑诊, 0), 0, '', ' (？)') As 临床诊断
      From 病人诊断医嘱 A, 病人诊断记录 B
      Where a.医嘱id = 医嘱id_In And a.诊断id = b.Id;
    r_Diagnose c_Diagnose%RowType;
  
    --获取指定表的行类型
    Procedure p_Get_Rowtype(Table_In In Varchar2) Is
    Begin
      If Table_In = '病人信息' Then
        Open c_Patient;
        Fetch c_Patient
          Into r_Patient;
      Elsif Table_In = '病人医嘱记录' Then
        Open c_Order;
        Fetch c_Order
          Into r_Order;
      Elsif Table_In = '病人诊断记录' Then
        Open c_Diagnose;
        Fetch c_Diagnose
          Into r_Diagnose;
      End If;
    Exception
      When Others Then
        Null;
    End p_Get_Rowtype;
  
  Begin
    Case
    --直接返回的输入元素
      When 元素名_In = '医嘱ID' Then
        v_Return := 医嘱id_In;
      When 元素名_In = '病人ID' Then
        v_Return := 病人id_In;
      
    --姓名，性别单独处理，可能是婴儿
      When Instr(',姓名,性别,性别编码,出生日期,', ',' || 元素名_In || ',') > 0 Then
        p_Get_Rowtype('病人医嘱记录');
        p_Get_Rowtype('病人信息');
        If Nvl(r_Order.婴儿, 0) = 0 Then
          If 元素名_In = '姓名' Then
            v_Return := r_Patient.姓名;
          Elsif 元素名_In = '性别' Then
            v_Return := r_Patient.性别;
          Elsif 元素名_In = '性别编码' Then
            v_Return := r_Patient.性别编码;
          Elsif 元素名_In = '出生日期' Then
            v_Return := To_Char(r_Patient.出生日期, 'YYYYMMDDMISS');
          End If;
        Else
          If 元素名_In = '姓名' Then
            Select Decode(婴儿姓名, Null, r_Patient.姓名 || '之婴' || Trim(To_Char(序号, '9')), 婴儿姓名) As 婴儿姓名
            Into v_Return
            From 病人新生儿记录
            Where 病人id = 病人id_In And 主页id = r_Order.主页id And 序号 = Nvl(r_Order.婴儿, 0);
          Elsif Instr('性别', 元素名_In) > 0 Then
            Select 婴儿性别
            Into v_Return
            From 病人新生儿记录
            Where 病人id = 病人id_In And 主页id = r_Order.主页id And 序号 = Nvl(r_Order.婴儿, 0);
            If 元素名_In = '性别编码' Then
              Select Decode(v_Return, '男', 'M', '女', 'F', 'O') Into v_Return From Dual;
            End If;
          Elsif 元素名_In = '出生日期' Then
            Select 出生时间
            Into v_Return
            From 病人新生儿记录
            Where 病人id = 病人id_In And 主页id = r_Order.主页id And 序号 = Nvl(r_Order.婴儿, 0);
            v_Return := To_Char(v_Return, 'YYYYMMDDMISS');
          End If;
        End If;
      
    --查询病人信息表返回的元素
      When Instr(',联系人地址,家庭电话,联系人电话,婚姻状况,身份证号,床号,就诊卡号,入院时间,出院时间,', ',' || 元素名_In || ',') > 0 Then
        p_Get_Rowtype('病人信息');
        Case 元素名_In
          When '联系人地址' Then
            v_Return := r_Patient.联系人地址;
          When '家庭电话' Then
            v_Return := r_Patient.家庭电话;
          When '联系人电话' Then
            v_Return := r_Patient.联系人电话;
          When '婚姻状况' Then
            v_Return := r_Patient.婚姻状况;
          When '身份证号' Then
            v_Return := r_Patient.身份证号;
          When '床号' Then
            v_Return := r_Patient.床号;
          When '就诊卡号' Then
            v_Return := r_Patient.就诊卡号;
          When '入院时间' Then
            v_Return := To_Char(r_Patient.入院时间, 'YYYYMMDDMISS');
          When '出院时间' Then
            v_Return := To_Char(r_Patient.出院时间, 'YYYYMMDDMISS');
          Else
            v_Return := '';
        End Case;
        --查询医嘱表返回的元素
      When Instr(',病人来源,开嘱医生,开嘱时间,校对护士,医嘱内容,紧急标志,紧急标志对码,', ',' || 元素名_In || ',') > 0 Then
        p_Get_Rowtype('病人医嘱记录');
        Case 元素名_In
          When '病人来源' Then
            v_Return := r_Order.病人来源;
          When '开嘱医生' Then
            v_Return := r_Order.开嘱医生;
          When '开嘱时间' Then
            v_Return := To_Char(r_Order.开嘱时间, 'YYYYMMDDMISS');
          When '校对护士' Then
            v_Return := r_Order.校对护士;
          When '医嘱内容' Then
            v_Return := r_Order.医嘱内容;
          When '紧急标志' Then
            v_Return := r_Order.紧急标志;
        End Case;
        --查询诊断记录返回的元素
      When 元素名_In = '临床诊断' Then
        p_Get_Rowtype('病人诊断记录');
        v_Return := r_Diagnose.临床诊断;
      
      Else
        --自行查询SQL返回值的元素
        If 元素名_In = '执行站点' Then
          p_Get_Rowtype('病人医嘱记录');
          Select Decode(站点, 1, 'SITE0002', 2, 'SITE0001', 3, 'SITE0003', 'SITE0001')
          Into v_Return
          From 部门表
          Where ID = r_Order.执行科室id;
        End If;
        If 元素名_In = '当前科室名称' Then
          p_Get_Rowtype('病人信息');
          Select 名称 Into v_Return From 部门表 Where ID = r_Patient.当前科室id;
        End If;
        If 元素名_In = '病区名称' Then
          p_Get_Rowtype('病人信息');
          Select 名称 Into v_Return From 部门表 Where ID = r_Patient.当前病区id;
        End If;
        If 元素名_In = '标识号' Then
          Select Decode(a.病人来源, 1, c.门诊号, 2, Decode(c.住院号, Null, c.门诊号, c.住院号), 4, c.健康号, c.门诊号)
          Into v_Return
          From 病人医嘱记录 A, 病人信息 C
          Where a.病人id = c.病人id And a.Id = 医嘱id_In;
        End If;
    End Case;
  
    Return Trim(v_Return);
  Exception
    When Others Then
      Return Null;
  End Ris_Replace_Element_Value;

  --14、删除RIS分院设置参数
  Procedure Ris分院设置_Delete Is
  Begin
    Delete From Ris分院设置;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris分院设置_Delete;

  --15、更新RISRis分院设置参数
  Procedure Ris分院设置_Update
  (
    Id_In           Ris分院设置.Id%Type,
    医院名称_In     Ris分院设置.医院名称%Type,
    医院代码_In     Ris分院设置.医院代码%Type,
    用户名_In       Ris分院设置.用户名%Type,
    密码_In         Ris分院设置.密码%Type,
    数据库服务名_In Ris分院设置.数据库服务名%Type
  ) Is
  
  Begin
  
    Insert Into Ris分院设置
      (ID, 医院名称, 医院代码, 用户名, 密码, 数据库服务名)
    Values
      (Id_In, 医院名称_In, 医院代码_In, 用户名_In, 密码_In, 数据库服务名_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris分院设置_Update;

  --16、登记危急值
  Procedure 病人危急值记录_Insert
  (
    Id_In         In 病人危急值记录.Id%Type,
    数据来源_In   In 病人危急值记录.数据来源%Type,
    病人id_In     In 病人危急值记录.病人id%Type,
    主页id_In     In 病人危急值记录.主页id%Type,
    挂号单_In     In 病人危急值记录.挂号单%Type,
    婴儿_In       In 病人危急值记录.婴儿%Type,
    姓名_In       In 病人危急值记录.姓名%Type,
    性别_In       In 病人危急值记录.性别%Type,
    年龄_In       In 病人危急值记录.年龄%Type,
    医嘱id_In     In 病人危急值记录.医嘱id%Type,
    标本id_In     In 病人危急值记录.标本id%Type,
    危急值描述_In In 病人危急值记录.危急值描述%Type,
    报告时间_In   In 病人危急值记录.报告时间%Type,
    报告科室id_In In 病人危急值记录.报告科室id%Type,
    报告人_In     In 病人危急值记录.报告人%Type
  ) Is
  Begin
  
    Zl_病人危急值记录_Insert(Id_In, 数据来源_In, 病人id_In, 主页id_In, 挂号单_In, 婴儿_In, 姓名_In, 性别_In, 年龄_In, 医嘱id_In, 标本id_In, 危急值描述_In,
                      报告时间_In, 报告科室id_In, 报告人_In);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 病人危急值记录_Insert;

  --17、取消危急值
  Procedure 病人危急值记录_Delete(医嘱id_In In 病人危急值记录.医嘱id%Type) Is
    Cursor c_Critical Is
      Select a.id From 病人危急值记录 A Where a.医嘱id = 医嘱id_In;
  Begin
    For r_Critical In c_Critical Loop
      zl_病人危急值记录_delete(r_Critical.id);
    End Loop;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 病人危急值记录_Delete;

  --18、发送临床医嘱
  Function 病人医嘱记录_Send(医嘱id_In In 病人医嘱发送.医嘱id%Type) Return Varchar2 Is
    --返回值分几种情况：
    --1、查询不到医嘱，返回空；
    --2、查询到医嘱，组织医嘱信息；
    --3、查询到医嘱，组织医嘱信息失败，返回空；
    Cursor c_Order Is
      Select a.id, a.相关id, a.病人id, a.主页ID, a.病人来源, a.挂号单 As NO, a.姓名, a.性别, a.年龄, e.健康号, a.病人科室id,
             To_Char(e.出生日期, 'YYYY-MM-DD') As 出生日期, e.家庭地址 As 地址, e.家庭电话 As 联系电话, a.开嘱医生 As 申请医生, e.门诊号, e.住院号,
             To_Char(a.开嘱时间, 'YYYY-MM-DD HH24:MI:SS') As 申请日期, c.影像类别 As 检查类型, a.诊疗项目id As 项目编码, a.医嘱内容 As 项目名称, a.紧急标志,
             a.诊疗项目id || ';' || a.标本部位 || ';' || a.检查方法 As 检查项目名称, a.开嘱科室ID As 申请科室ID, a.执行科室ID, b.No As 单据号, e.费别,
             e.医疗付款方式, nvl(a.婴儿, 0) As 婴儿, e.就诊卡号, a.医生嘱托, e.民族, e.身份证号, b.发送号, Null As 检查目的, 0 As 状态
      From 病人医嘱记录 A, 病人医嘱发送 B, 影像检查项目 C, 病人信息 E
      Where a.id = b.医嘱id And a.诊疗项目id = c.诊疗项目id And a.病人id = e.病人id And (a.ID = 医嘱id_In Or a.相关ID = 医嘱id_In) And
            a.医嘱状态 = 8;
    r_Order c_Order%RowType;
  
    v_Return       Varchar2(4000) := Null;
    v_临床诊断     Varchar2(2000) := Null;
    v_医嘱附件     Varchar2(800) := Null;
    v_Val          Varchar2(4000) := Null;
    v_Col          Varchar2(1000) := Null;
    v_病人科室     部门表.名称%Type := Null;
    v_执行科室     部门表.名称%Type := Null;
    v_申请科室     部门表.名称%Type := Null;
    v_病区         部门表.名称%Type;
    v_病区ID       病案主页.当前病区id%Type;
    v_病人类型     病案主页.病人类型%Type;
    v_床号         病案主页.出院病床%Type;
    v_急诊         病人挂号记录.急诊%Type;
    v_挂号ID       病人挂号记录.id%Type;
    n_Baby         病人医嘱记录.婴儿%Type;
    v_婴儿姓名     病人信息.姓名%Type;
    v_婴儿性别     病人信息.性别%Type;
    v_婴儿年龄     病人信息.年龄%Type;
    v_婴儿出生日期 病人信息.出生日期%Type;
    v_是否急诊     病人挂号记录.急诊%Type;
    v_年龄         病人信息.年龄%Type;
    v_医嘱类型     Number; --1 主医嘱；2 部位医嘱；
  
    --提取临床诊断
    Function f_GetDiagnose
    (
      v_病人来源 In 病人医嘱记录.病人来源%Type,
      v_医嘱id   In 病人医嘱记录.id%Type
    ) Return Varchar2 Is
    
      --住院临床诊断，只提取主要诊断
      Cursor c_DiagnoseIn Is
        Select a.id, e.诊断描述
        From 病人医嘱记录 A, 病人诊断记录 E
        Where a.病人ID = e.病人id And a.主页id = e.主页id And e.记录来源 = 3 And e.诊断类型 In (2, 12) And e.诊断次序 = 1 And e.编码序号 = 1 And
              a.id = v_医嘱id;
      r_DiagnoseIn c_DiagnoseIn%RowType;
    
      --门诊和体检的临床诊断，提取医嘱对应的诊断
      Cursor c_DiagnoseOut Is
        Select a.id, e.诊断描述
        From 病人医嘱记录 A, 病人诊断医嘱 D, 病人诊断记录 E
        Where d.医嘱id = a.id And d.诊断id = e.id And a.id = v_医嘱id;
      r_DiagnoseOut c_DiagnoseOut%RowType;
    
      v_Return Varchar2(2000);
      iCount   Number;
    Begin
      iCount := 0;
      If v_病人来源 = 2 Then
        Open c_DiagnoseIn;
        Fetch c_DiagnoseIn
          Into r_DiagnoseIn;
        While c_DiagnoseIn%Found Loop
          iCount := iCount + 1;
          If iCount = 1 Then
            If lengthb(iCount || '、' || r_DiagnoseIn.诊断描述 || '。') < 2000 Then
              v_Return := iCount || '、' || r_DiagnoseIn.诊断描述 || '。';
            End If;
          Else
            If lengthb(v_Return || Chr(10) || iCount || '、' || r_DiagnoseIn.诊断描述 || '。') < 2000 Then
              v_Return := v_Return || Chr(10) || iCount || '、' || r_DiagnoseIn.诊断描述 || '。';
            End If;
          End If;
        
          Fetch c_DiagnoseIn
            Into r_DiagnoseIn;
        End Loop;
      
      Else
        Open c_DiagnoseOut;
        Fetch c_DiagnoseOut
          Into r_DiagnoseOut;
        While c_DiagnoseOut%Found Loop
          iCount := iCount + 1;
          If iCount = 1 Then
            If lengthb(iCount || '、' || r_DiagnoseOut.诊断描述 || '。') < 2000 Then
              v_Return := iCount || '、' || r_DiagnoseOut.诊断描述 || '。';
            End If;
          Else
            If lengthb(v_Return || Chr(10) || iCount || '、' || r_DiagnoseOut.诊断描述 || '。') < 2000 Then
              v_Return := v_Return || Chr(10) || iCount || '、' || r_DiagnoseOut.诊断描述 || '。';
            End If;
          End If;
        
          Fetch c_DiagnoseOut
            Into r_DiagnoseOut;
        End Loop;
      End If;
    
      If iCount = 1 Then
        v_Return := substr(v_return, 3);
      End If;
      Return v_Return;
    
    End f_GetDiagnose;
  
    --提取医嘱附件
    Function f_GetAttachment(v_医嘱id In 病人医嘱记录.id%Type) Return Varchar2 Is
      Cursor c_Attachment Is
        Select a.项目, a.内容 From 病人医嘱附件 A Where a.医嘱ID = v_医嘱id Order By 排列;
      r_Attachment c_Attachment%RowType;
    
      v_Return Varchar2(800);
    Begin
      Open c_Attachment;
      Fetch c_Attachment
        Into r_Attachment;
      While c_Attachment%Found Loop
        If r_Attachment.内容 Is Not Null Then
          If v_Return Is Null Then
            If lengthb('【' || nvl(r_Attachment.项目, '') || '】' || chr(10) || nvl(r_Attachment.内容, '')) < 800 Then
              v_Return := '【' || nvl(r_Attachment.项目, '') || '】' || chr(10) || nvl(r_Attachment.内容, '');
            End If;
          Else
            If lengthb(v_Return || Chr(10) || '【' || nvl(r_Attachment.项目, '') || '】' || chr(10) ||
                       nvl(r_Attachment.内容, '')) < 800 Then
              v_Return := v_Return || Chr(10) || '【' || nvl(r_Attachment.项目, '') || '】' || chr(10) ||
                          nvl(r_Attachment.内容, '');
            End If;
          End If;
        End If;
        Fetch c_Attachment
          Into r_Attachment;
      End Loop;
      Return v_Return;
    End f_GetAttachment;
  
    --提取科室名称
    Function f_GetDeptName(v_科室id In 部门表.id%Type) Return Varchar2 Is
      v_Return 部门表.名称%Type;
    Begin
      Select Max(名称) Into v_Return From 部门表 Where ID = v_科室id;
      Return v_Return;
    End f_GetDeptName;
  
  Begin
  
    Open c_Order;
    Fetch c_Order
      Into r_Order;
    While c_Order%Found Loop
      --根据病人来源，查询患者的 病案主页，病人挂号记录,临床诊断等信息，只查一次
      If v_病人科室 Is Null Then
        v_急诊     := 0;
        v_挂号ID   := '';
        v_病人类型 := '';
        v_病区ID   := '';
        v_床号     := '';
      
        --只有住院和门诊，才提取 病案主页，挂号记录，临床诊断
        If r_Order.病人来源 = 2 Then
          Select b.病人类型, b.当前病区id As 病区id, b.出院病床 As 床号
          Into v_病人类型, v_病区ID, v_床号
          From 病人医嘱记录 A, 病案主页 B
          Where a.病人id = b.病人id And a.主页id = b.主页id And a.id = 医嘱id_In;
        
          v_临床诊断 := f_GetDiagnose(r_Order.病人来源, 医嘱id_In);
        Elsif r_Order.病人来源 = 1 Then
          Select b.急诊, b.id As 挂号id
          Into v_急诊, v_挂号ID
          From 病人医嘱记录 A, 病人挂号记录 B
          Where a.挂号单 = b.no And a.id = 医嘱id_In;
        
          v_临床诊断 := f_GetDiagnose(r_Order.病人来源, 医嘱id_In);
        End If;
      
        Select decode(nvl(r_Order.紧急标志, 0), 1, 1, nvl(v_急诊, 0)) Into v_是否急诊 From dual;
        v_医嘱附件 := f_GetAttachment(医嘱id_In);
      
        v_病区     := f_GetDeptName(v_病区ID);
        v_执行科室 := f_GetDeptName(r_Order.执行科室id);
        v_申请科室 := f_GetDeptName(r_Order.申请科室ID);
        v_病人科室 := f_GetDeptName(r_Order.病人科室id);
      End If;
    
      --申请科室 ，修改成提取ID，然后通过f_GetDeptName获取名称
    
      --循环医嘱记录，逐个发送
      If r_Order.相关id Is Null Then
        --发送主医嘱，需要处理婴儿医嘱
        v_医嘱类型 := 1;
        If r_Order.婴儿 = 0 Then
          v_Val  := r_Order.ID || '[;]' || r_Order.病人ID || '[;]' || r_Order.病人来源 || '[;]' || r_Order.主页ID || '[;]' ||
                    r_Order.NO || '[;]' || r_Order.姓名 || '[;]' || r_Order.性别 || '[;]' || r_Order.出生日期 || '[;]' ||
                    r_Order.地址 || '[;]' || r_Order.联系电话 || '[;]' || v_申请科室 || '[;]' || r_Order.申请医生 || '[;]' ||
                    r_Order.门诊号 || '[;]' || r_Order.住院号 || '[;]' || v_病区 || '[;]' || v_床号;
          v_年龄 := r_Order.年龄;
        Else
          n_Baby := r_Order.婴儿;
          Select Decode(a.婴儿姓名, Null, b.姓名 || '之子' || Trim(To_Char(a.序号, '9')), a.婴儿姓名) As 婴儿姓名, 婴儿性别,
                 round(Sysdate - a.出生时间) || '天' As 婴儿年龄, To_Char(a.出生时间, 'YYYY-MM-DD') As 出生时间
          Into v_婴儿姓名, v_婴儿性别, v_婴儿年龄, v_婴儿出生日期
          From 病人新生儿记录 A, 病人信息 B
          Where a.病人id = r_Order.病人ID And a.主页id = r_Order.主页ID And a.病人id = b.病人id And a.序号 = n_Baby;
        
          v_年龄 := v_婴儿年龄;
          v_Val  := r_Order.ID || '[;]' || r_Order.病人ID || '[;]' || r_Order.病人来源 || '[;]' || r_Order.主页ID || '[;]' ||
                    r_Order.NO || '[;]' || v_婴儿姓名 || '[;]' || v_婴儿性别 || '[;]' || v_婴儿出生日期 || '[;]' || r_Order.地址 ||
                    '[;]' || r_Order.联系电话 || '[;]' || v_申请科室 || '[;]' || r_Order.申请医生 || '[;]' || r_Order.门诊号 || '[;]' ||
                    r_Order.住院号 || '[;]' || v_病区 || '[;]' || v_床号;
        End If;
        --v_科室ID 为空，是否会出错？                
        v_Val := v_Val || '[;]' || Trim(v_临床诊断) || '[;]' || r_Order.申请日期 || '[;]' || r_Order.检查目的 || '[;]' ||
                 r_Order.检查类型 || '[;]' || r_Order.项目编码 || '[;]' || r_Order.项目名称 || '[;]' || r_Order.状态 || '[;]' ||
                 v_是否急诊 || '[;]' || v_病人科室 || '[;]' || v_年龄 || '[;]' || r_Order.健康号 || '[;]' || r_Order.发送号 || '[;]' ||
                 nvl(r_Order.申请科室ID, '') || '[;]' || r_Order.单据号 || '[;]' || v_挂号ID || '[;]' || Trim(v_医嘱附件) || '[;]' ||
                 v_执行科室 || '[;]' || r_Order.费别 || '[;]' || r_Order.医疗付款方式 || '[;]' || r_Order.执行科室id || '[;]' ||
                 nvl(r_Order.就诊卡号, 0) || '[;]' || r_Order.医生嘱托 || '[;]' || r_Order.民族 || '[;]' || v_病人类型 || '[;]' ||
                 r_Order.身份证号;
        v_col := 'appno[;]patid[;]patsource[;]pageid[;]regno[;]name[;]sex[;]birthdate[;]address[;]phoneno[;]dept[;]doctor[;]outpatno[;]inpatno[;]ward[;]bedno[;]clinicdiag[;]appdate[;]clinicdesc[;]modality[;]patno[;]partname[;]status[;]emergency[;]patdept[;]age[;]physicalexamid[;]sendno[;]deptno[;]billno[;]regid[;]clinicdiagex[;]executdept[;]feekind[;]paykind[;]executdeptID[;]medicalCardID[;]DoctorEntrust[;]Nation[;]PatientType[;]IDCard';
      
      Else
        --发送部位医嘱
        v_医嘱类型 := 2;
        v_Val      := r_Order.相关id || '[;]' || r_Order.ID || '[;]' || r_Order.检查项目名称;
        v_col      := 'AppNO[;]AppPartNo[;]ExamPlace';
      End If;
    
      If v_Return Is Null Then
        v_Return := v_医嘱类型 || '[:]' || v_Col || '[:]' || v_Val;
      Else
        v_Return := v_Return || '{;}' || v_医嘱类型 || '[:]' || v_Col || '[:]' || v_Val;
      End If;
    
      Fetch c_Order
        Into r_Order;
    End Loop;
  
    If c_Order%RowCount = 0 Then
      v_Return := '';
    End If;
  
    Return Trim(v_Return);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
      Return Null;
  End 病人医嘱记录_Send;

End b_Zlxwinterface;
/

--Pacs文档编辑器

--影像报告原型管理(---定义部分---)***************************************************
CREATE OR REPLACE Package b_PACS_Common Is
  Type t_Refcur Is Ref Cursor;

--1 获取参数的缓冲数据
Procedure p_GetParInfBuf(
  Val           Out t_Refcur,
  系统_In In 影像参数说明.系统%Type,
  模块_In  In 影像参数说明.模块%Type
  );
  
--2 功能：将由逗号分隔的不带引号的字符序列转换为单列数据表
  Function f_Str2list
  (
    Str_In   In Varchar2,
    Split_In In Varchar2 := ','
  ) Return t_Strlist
    Pipelined;
  
--3  获取参数值的缓冲数据
--当前用户所在计算机的参数值
Procedure p_GetParValueBuf(
  Val           Out t_Refcur,
  系统_In In 影像参数说明.系统%Type,
  模块_In In 影像参数说明.模块%Type,
  科室ID_In In Varchar2,
  机器名_In In Varchar2,
  用户ID_In In Number);

--4  获取权限的缓冲数据
Procedure p_GetPopedomBuf(
  Val           Out t_Refcur,
  系统_In In 影像参数说明.系统%Type,
  模块_IN In 影像参数说明.模块%Type,
  用户名_In In Varchar2);

--5  设置参数值
Procedure p_SetParameterValue(
  参数ID_In    In 影像参数取值.参数ID%Type,
  参数标识_In In 影像参数取值.参数标识%Type,
  参数值_In    In 影像参数取值.参数值%Type);

--6  获取用户账号信息
Function f_Get_Personal_Info_By_Account(
	Account_In In Varchar2
) Return Xmltype;

end b_PACS_Common ;

/



--*************************************************************************************
--*                  影像报告原型管理(---实现部分---)                                                        *
--*************************************************************************************
CREATE OR REPLACE Package Body b_PACS_Common  Is

--1 获取参数的缓冲数据
Procedure p_GetParInfBuf(
  Val           Out t_Refcur,
  系统_In In 影像参数说明.系统%Type,
  模块_In  In 影像参数说明.模块%Type
) Is
Begin
Open  Val For
   Select RawToHex(ID) As ID,RawToHex(PID) As PID,系统,模块,参数名,默认值,参数级别,启用条件
   From 影像参数说明
   Where 系统=系统_In And 模块=模块_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End p_GetParInfBuf;

--2 功能：将由逗号分隔的不带引号的字符序列转换为单列数据表
Function f_Str2list(
    Str_In   In Varchar2,
    Split_In In Varchar2 := ','
  ) Return t_Strlist
    Pipelined As
    v_Str Long;
    P     Number;
    --功能：将由逗号分隔的不带引号的字符序列转换为单列数据表
    --参数：Str_In,如:甲抗,胃溃疡,胃出血...,Split_In,分隔符,缺省为,号
    --说明：
    --1．当SQL语句中涉及“IN(常量1, 常量2,…) ”子句时使用这种方式以便利用绑定变量。
    --2．使用这两个函数时，需要在SQL语句中加入“/*+ Rule*/”提示，因为Cbo下临时内存表没有统计数据,。
    --3．两种调用示例
    --Select /*+ Rule*/ * From Sample_List Where Title In (Select * From Table(f_Str2list('甲抗,胃溃疡,胃出血'));
    --Select /*+ Rule*/ A.* From Sample_List A, Table(f_Str2list('甲抗,胃溃疡,胃出血')) B Where A.Title = B.Column_Value;
  Begin
    If Str_In Is Null Then
      Return;
    End If;
    v_Str := Str_In || Split_In;
    Loop
      P := Instr(v_Str, Split_In);
      Exit When(Nvl(P, 0) = 0);
      Pipe Row(Trim(Substr(v_Str, 1, P - 1)));
      v_Str := Substr(v_Str, P + 1);
    End Loop;
    Return;
  End;

--3  获取参数值的缓冲数据
--当前用户所在计算机的参数值
Procedure p_GetParValueBuf(
  Val           Out t_Refcur,
  系统_In In 影像参数说明.系统%Type,
  模块_In In 影像参数说明.模块%Type,
  科室ID_In In Varchar2,
  机器名_In In Varchar2,
  用户ID_In In Number
) Is
Begin
	Open  Val For
	    Select RawToHex(b.ID) As ID, RawToHex(b.参数ID) As 参数ID,b.参数标识,b.参数值
		From 影像参数说明 a, 影像参数取值 b
		Where  a.id=b.参数id And a.系统=系统_In And a.模块=模块_In and (a.参数级别=0 or a.参数级别=1 or (a.参数级别=2 and b.参数标识=科室ID_In) or (a.参数级别=3 and b.参数标识=用户ID_In) or (a.参数级别=4 and b.参数标识=机器名_In));
Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
End 	p_GetParValueBuf;

--4  获取权限的缓冲数据
Procedure p_GetPopedomBuf(
  Val           Out t_Refcur,
  系统_In In 影像参数说明.系统%Type,
  模块_In In 影像参数说明.模块%Type,
  用户名_In In Varchar2
)Is
Begin
    --返回用户, 模块号,功能
	Open  Val For
	    Select a.用户,b.系统, b.序号 as 模块, b.功能
		From zluserroles a, zlrolegrant b
		Where a.角色=b.角色 And a.用户=用户名_In And b.系统=系统_In And b.序号=模块_In;
Exception
  When Others Then
  Zl_Errorcenter(Sqlcode, Sqlerrm);
End p_GetPopedomBuf;

--5  更新参数值
Procedure p_SetParameterValue(
  参数ID_In    In 影像参数取值.参数ID%Type,
  参数标识_In In 影像参数取值.参数标识%Type,
  参数值_In    In 影像参数取值.参数值%Type
)Is
Begin
	Update 影像参数取值 Set 参数值=参数值_In Where 参数ID=参数ID_In And 参数标识=参数标识_In;
	If Sql%RowCount = 0 Then
	  Insert Into 影像参数取值(ID, 参数ID,参数标识,参数值)
	  Values(sys_guid(), 参数ID_In,参数标识_In,参数值_In);
	End If;
Exception
  When Others Then
  Zl_Errorcenter(Sqlcode, Sqlerrm);
End p_SetParameterValue;

--6  获取用户账号信息
Function f_Get_Personal_Info_By_Account(
	Account_In In Varchar2
) Return Xmltype Is
  Docxml   Xmltype;
Begin 
  Select Xmltype('<root></root>') Into Docxml From Dual;  
  Select Appendchildxml(Docxml, '/root',
                         Xmlconcat(Xmlelement("code", a.Id), Xmlelement("full_name", a.姓名),
                                    Xmlelement("sex", Xmlattributes(a.性别 As "display"),
                                                Decode(a.性别, '男', '1', '女', '2', '未知', '0', '9')),
                                    Xmlelement("birthday", To_Char(a.出生日期, 'yyyy-mm-dd')),
                                    Xmlelement("idcard_num", a.身份证号)))
  Into Docxml
  From 人员表 A, 上机人员表 B
  Where b.用户名 = Account_In And b.人员id = a.Id And Nvl(a.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate;  

  Select Appendchildxml(Docxml, '/root',
                         Xmlelement("departments",
                                     Xmlagg(Xmlelement("department", Xmlattributes(c.名称 As "display", d.缺省 As "current"),
                                                        Xmlelement("dept_value", c.Id)))))
  Into Docxml
  From 上机人员表 B, 部门表 C, 部门人员 D
  Where b.用户名 = Account_In And b.人员id = d.人员id And d.部门id = c.Id;

  For r_Record In (Select d.部门id, Xmlelement("subjects", Xmlagg(Xmlelement("subject", c.名称))) As 部门学科
                   From 临床部门 A, 上机人员表 B, 部门人员 D, 临床性质 C
                   Where b.用户名 = Account_In And b.人员id = d.人员id And a.部门id = d.部门id And a.工作性质 = c.编码
                   Group By d.部门id
                   Order By d.部门id) Loop
    Select Appendchildxml(Docxml, '/root/departments/department[dept_value=' || r_Record.部门id || ']', r_Record.部门学科)
    Into Docxml
    From Dual;
  End Loop;

  Return Docxml;
Exception
  When Others Then
    Return Null;
End f_Get_Personal_Info_By_Account;

End b_PACS_Common;

/



--*************************************************************************************
--*								   (---声明部分---)                                    *
--*************************************************************************************
Create Or Replace Package b_PACS_Config Is
  Type t_Refcur Is Ref Cursor;

  -- 功    能：获取影像字典清单
  Procedure p_GetAllDictList(
	Val			Out t_Refcur
  );

  -- 功    能：获取影像字典内容
  Procedure p_GetAllDictItems(
    Val           Out t_Refcur,
	字典ID_In	In 影像字典内容.字典ID%Type
  );

  -- 功    能：新增或修改影像字典内容
  Procedure p_EditDictItem(
	字典ID_In		In 影像字典内容.字典ID%Type,
	旧编号_In		In 影像字典内容.编号%Type,
	编号_In			In 影像字典内容.编号%Type,
	名称_In			In 影像字典内容.名称%Type,
	简码_In			In 影像字典内容.简码%Type,
	说明_In			In 影像字典内容.说明%Type
  );

  -- 功    能：删除影像字典内容
  Procedure p_DelDictItem(
	字典ID_In		In 影像字典内容.字典ID%Type,
	编号_In			In 影像字典内容.编号%Type
  );
End b_PACS_Config;
/

--*************************************************************************************
--*								   (---实现部分---)                                    *
--*************************************************************************************
Create Or Replace Package Body b_PACS_Config  Is
  -- 功    能：获取影像字典清单
  Procedure p_GetAllDictList(
	Val			Out t_Refcur
  )
  Is
	strSql varchar2(100);
  Begin
	strSql := 'select Rawtohex(ID) ID,编号,名称,说明,是否系统,分组 From 影像字典清单';
	Open Val For strSql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetAllDictList;

  -- 功    能：获取影像字典内容
  Procedure p_GetAllDictItems(
    Val           Out t_Refcur,
	字典ID_In		In 影像字典内容.字典ID%Type
  )
  Is
	strSql varchar2(200);
  Begin
	strSql := 'Select Rawtohex(A.字典id) Rid, A.编号, A.名称, A.简码, A.说明 '||
			  'From 影像字典内容 A Where A.字典id = '''|| 字典ID_In ||'''';
	Open Val For strSql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetAllDictItems;

  -- 功    能：新增或修改影像字典内容
  Procedure p_EditDictItem(
	字典ID_In		In 影像字典内容.字典ID%Type,
	旧编号_In		In 影像字典内容.编号%Type,
	编号_In			In 影像字典内容.编号%Type,
	名称_In			In 影像字典内容.名称%Type,
	简码_In			In 影像字典内容.简码%Type,
	说明_In			In 影像字典内容.说明%Type
  )
  Is
	n_Num Number;
	v_Msg Varchar2(50);
	Err	  Exception;
  Begin
	If 旧编号_In<>'-1' Then
	  Select Count(字典ID) Into n_Num From 影像字典内容 Where 字典ID = 字典ID_In And 编号 = 编号_In And 编号<>旧编号_In;
	  If n_Num > 0 Then
		v_Msg:='所属字典ID和编号重复!';
		Raise Err;
	  End IF;

	  Update 影像字典内容 A
	  Set A.编号 = 编号_In,A.名称 = 名称_In,A.简码 = 简码_In,A.说明 = 说明_In
	  Where A.字典ID = 字典ID_In And A.编号 = 旧编号_In;
	Else
	  Select Count(字典ID) Into n_Num From 影像字典内容 Where 字典ID = 字典ID_In And 编号 = 编号_In;
	  If n_Num > 0 Then
		v_Msg:='所属字典ID和编号重复!';
		Raise Err;
	  End IF;
	
	  Insert Into 影像字典内容(字典ID,编号,名称,简码,说明)
	  Values(字典ID_In,编号_In,名称_In,简码_In,说明_In);
	End If;
  Exception
	When Err Then
	  Raise_Application_Error(-20101,v_Msg);
	When Others Then
	  Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditDictItem;

  -- 功    能：删除影像字典内容
  Procedure p_DelDictItem(
	字典ID_In		In 影像字典内容.字典ID%Type,
	编号_In			In 影像字典内容.编号%Type
  )
  Is
  Begin
	Delete From 影像字典内容 Where 字典ID = 字典ID_In And 编号 = 编号_In;
  Exception
	When Others Then
	  Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelDictItem;
End b_PACS_Config;
/





Create Or Replace Package b_PACS_RptPublic Is
  Type t_Strlist Is Table Of Varchar2(4000);
  --功能：将由逗号分隔的不带引号的字符序列转换为单列数据表
  Function f_Str2list(
    Str_In   In Varchar2,
    Split_In In Varchar2 := ','
    ) Return t_Strlist
    Pipelined;

  --功能:跟据传入的表名称提取表中最大的Code并递增
  Function f_Get_Nextcode(
    Tablename_In Varchar2,
    Len_In       Number := 0,
    Mount_In     Number := 0,
    Pre_In       Varchar2 := Null
    ) Return Varchar2;

  --功能：生成字符串拼音首码
  Function f_Spellcode(
    v_Instr  In Varchar2,
    v_Outnum In Integer := 10
    ) Return Varchar2;

  --错误处理中心
  Procedure zl_ErrorCenter(
    Err_Num In Number,
    Err_Msg In Varchar2
    );

 --从传入的XML中提取编辑记录
  Function f_Geteditlist(
    Content_In In Xmltype
	) Return t_Editlist;

  Function Xml2clob(
    Xml_In Xmltype
	) Return Clob;

  Function f_Getlastedit(
    Content_In In Xmltype
	) Return t_Editlist;
  
  ----拆分匿名数据
  --Function f_Disjoin_Anonym
  --(
    --Content_In    In Xmltype,
    --x_Anonym_Data Out Xmltype
  --) Return Xmltype;

  ----合并匿名数据
  --Function f_Incorporate_Anonym
  --(
    --Content_In    In Xmltype,
    --x_Anonym_Data In Xmltype
  --) Return Clob;

  --设置XML的节点值,当节点为<ele></ele>此类闭合节点时，Updatexml函数无效
  Procedure p_Set_Elementtext(
    Texture In Out Xmltype,
    Ename   In Varchar2,
    Eaname  In Varchar2,
    Eatext  In Varchar2,
    Etext   In Varchar2
    );
  --根据子文档ID提取子文档当前状态
 Function f_Get_Docstatus(
    Content_In In Xmltype
    ) Return Varchar2;

  Function f_If_Intersect(
    Str1 Varchar2,
    Str2 Varchar2
    ) Return Number;

End b_PACS_RptPublic;
/

Create Or Replace Package Body b_PACS_RptPublic Is

  Function Xml2clob(
    Xml_In Xmltype
	) Return Clob As
  Begin
    Return Xml_In.Getclobval();
  End Xml2clob;

  Function f_Str2list(
    Str_In   In Varchar2,
    Split_In In Varchar2 := ','
  ) Return t_Strlist
    Pipelined As
    v_Str Long;
    P     Number;
    --功能：将由逗号分隔的不带引号的字符序列转换为单列数据表
    --参数：Str_In,如:甲抗,胃溃疡,胃出血...,Split_In,分隔符,缺省为,号
    --说明：
    --1．当SQL语句中涉及“IN(常量1, 常量2,…) ”子句时使用这种方式以便利用绑定变量。
    --2．使用这两个函数时，需要在SQL语句中加入“/*+ Rule*/”提示，因为Cbo下临时内存表没有统计数据,。
    --3．两种调用示例
    --Select /*+ Rule*/ * From Sample_List Where Title In (Select * From Table(f_Str2list('甲抗,胃溃疡,胃出血'));
    --Select /*+ Rule*/ A.* From Sample_List A, Table(f_Str2list('甲抗,胃溃疡,胃出血')) B Where A.Title = B.Column_Value;
  Begin
    If Str_In Is Null Then
      Return;
    End If;
    v_Str := Str_In || Split_In;
    Loop
      P := Instr(v_Str, Split_In);
      Exit When(Nvl(P, 0) = 0);
      Pipe Row(Trim(Substr(v_Str, 1, P - 1)));
      v_Str := Substr(v_Str, P + 1);
    End Loop;
    Return;
  End;

  Function f_Get_Nextcode(
    Tablename_In Varchar2,
    Len_In       Number := 0,
    Mount_In     Number := 0,
    Pre_In       Varchar2 := Null
  ) Return Varchar2 Is
    --跟据传入的表名称提取表中最大的Code并递增
    --Len_In        当指定长度时，按指定长度最大Code递增，大于表字段长度或不传时为表字段长度
    --Mount_In      当指定长度时，最大Code每位不能再进位时是否扩容长度，=0不扩容，＝1扩容 比如当前最大Code为 Z99,如果扩容则返回1A00否则返回Z99
    --递增规则：数字0123456789 字母A...Z,递增到达9或Z时前一位增长并且当前位转为0或A，如果前一个为非数字或字母则向前一位递增
    --字母全部为大写，无小写,当无指定长度Code时返回指定长度01，比如指定长度为5，则返回 00001
    v_Sql      Varchar2(100);
    v_Maxcode  Varchar2(150);
    v_Origcode Varchar2(150);
    v_Old      Varchar2(6);
    v_New      Varchar2(6);
    v_Return   Varchar2(150);
    n_Collen   Number;
    n_Length   Number;
    Err_Custom Exception;
    v_Msg Varchar2(200);
    Function f_Char_Carry(Word_In Varchar2) Return Varchar2 As
      v_Temp Varchar2(6);
      n_Asc  Number;
    Begin
      Select Ascii(Upper(Word_In)) Into n_Asc From Dual;
      If n_Asc = 57 Then
        v_Temp := '0';
      Elsif n_Asc = 90 Then
        v_Temp := 'A';
      Elsif n_Asc >= 48 And n_Asc <= 56 Or n_Asc >= 65 And n_Asc <= 89 Then
        v_Temp := Chr(Ascii(Word_In) + 1);
      Else
        v_Temp := Word_In;
      End If;
      Return v_Temp;
    End;
  Begin
    Begin
      Select Data_Length
      Into n_Collen
      From User_Tab_Cols
      Where Table_Name = Upper(Tablename_In) And Upper(Column_Name) = '编码';
    Exception
      When Others Then
        Null;
        v_Msg := '没有当前要查找的表，或表中没有【编码】字段！';
        Raise Err_Custom;
    End;
  
    --当传入长度为0或大于字段长度时取当前最大长度，否则取传入长度相当的最大编码
    If Len_In = 0 Or Len_In > n_Collen Then
      v_Sql := 'Select Max(Length(编码)) From ' || Tablename_In;
      Execute Immediate v_Sql
        Into n_Length;
    Else
      n_Length := Len_In;
    End If;
  
    If Nvl(n_Length, 0) = 0 Then
      Return '1';
    End If;
    
    If (Pre_In Is Not Null) And Length(Pre_In) >= n_Length Then
      v_Msg := '指定编码的长度应该大于前缀长度';
      Raise Err_Custom;
    End If;
  
    --查找指定前缀编码的最大值
    If Pre_In Is Not Null Then
      v_Sql := 'Select Max(编码) From ' || Tablename_In || ' Where upper(substr(code,1,length(''' || Pre_In ||
               '''))) =' || '' || 'upper(''' || Pre_In || ''')';
      Execute Immediate v_Sql
        Into v_Maxcode;
    Else
      v_Sql := 'Select Max(编码) From ' || Tablename_In || ' Where Length(编码)=' || n_Length;
      Execute Immediate v_Sql
        Into v_Maxcode;
    End If;
  
    --如果最大的code为空，那么赋值为1,前面增加0的数量由最长长度决定
    If v_Maxcode Is Null Then
      If Pre_In Is Null Then
        Select LPad('1', n_Length, '0') Into v_Maxcode From Dual;
        Return v_Maxcode;
      Else
        Select Pre_In || LPad('1', n_Length - Length(Pre_In), '0') Into v_Maxcode From Dual;
        Return v_Maxcode;
      End If;
    Else
      If Pre_In Is Null Then
        v_Maxcode  := Upper(v_Maxcode);
        v_Origcode := v_Maxcode;
      Else
        --补充为指定长度
        v_Maxcode := Upper(Pre_In || LPad(Substr(v_Maxcode, Length(Pre_In) + 1), n_Length - Length(Pre_In), '0'));
        --指定前缀时，仅以前缀后的字符串作为计算的值
        v_Origcode := Substr(v_Maxcode, Length(Pre_In) + 1);
        v_Maxcode  := v_Origcode;
      End If;
      
      For I In 0 .. Length(v_Maxcode) Loop
        If I = Length(v_Maxcode) Then
          If Len_In <> 0 And Mount_In = 0 Then
            --指定长度并且不扩容长度，已在到最大时，不再进位
            Return v_Origcode;
          Else
            v_Old := 'Z';
            v_New := '1';
          End If;
        Else
          v_Old := Substr(v_Maxcode, Length(v_Maxcode) - I, 1);
          v_New := f_Char_Carry(v_Old);
        End If;
      
        --新旧值相等表明为非数字或字母,需要向前查找
        If v_Old != v_New Then
          v_Return := Substr(v_Maxcode, 0, Length(v_Maxcode) - I - 1) || v_New ||
                      Substr(v_Maxcode, Length(v_Maxcode) - I + 1);
          If v_New != '0' And v_New != 'A' Then
            If Pre_In Is Null Then
              Return v_Return;
            Else
              Return Pre_In || v_Return;
            End If;
          Else
            --等于‘0’表明当前进位,前一位递增
            v_Maxcode := v_Return;
          End If;
        End If;
      End Loop;
    End If;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, v_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

  --功能：生成字符串拼音首码
  --参数：v_Instr需要生成拼音的字符串；v_Outnum 生成首码长度，默认10，超过40个字符最大10
  Function f_Spellcode(
    v_Instr  In Varchar2,
    v_Outnum In Integer := 10
  ) Return Varchar2 Is
    v_Spell     Varchar2(40);
    v_Input     Varchar2(1000);
    v_Bitchar   Varchar2(100);
    r_Bitchar   Varchar2(100);
    v_Bitnum    Integer;
    v_Outmaxnum Integer;
    Function f_Nlssort(p_Word In Varchar2) Return Varchar2 As
    Begin
      Return Nlssort(p_Word, 'NLS_SORT=SCHINESE_PINYIN_M');
    End;
  Begin
    If v_Outnum < 1 Or v_Outnum > 40 Then
      v_Outmaxnum := 10;
    Else
      v_Outmaxnum := v_Outnum;
    End If;
  
    If v_Instr Is Null Or Length(LTrim(v_Instr)) = 0 Then
      v_Spell := '';
    Else
      v_Input := Upper(v_Instr);
      v_Spell := '';
      For v_Bitnum In 1 .. Length(v_Input) Loop
        v_Bitchar := Substr(v_Input, v_Bitnum, 1);
        r_Bitchar := f_Nlssort(v_Bitchar);
        If r_Bitchar >= f_Nlssort('吖') And r_Bitchar <= f_Nlssort('驁') Then
          v_Spell := v_Spell || 'A';
        Elsif r_Bitchar >= f_Nlssort('八') And r_Bitchar <= f_Nlssort('簿') Then
          v_Spell := v_Spell || 'B';
        Elsif r_Bitchar >= f_Nlssort('嚓') And r_Bitchar <= f_Nlssort('錯') Then
          v_Spell := v_Spell || 'C';
        Elsif r_Bitchar >= f_Nlssort('咑') And r_Bitchar <= f_Nlssort('鵽') Then
          v_Spell := v_Spell || 'D';
        Elsif r_Bitchar >= f_Nlssort('妸') And r_Bitchar <= f_Nlssort('樲') Then
          v_Spell := v_Spell || 'E';
        Elsif r_Bitchar >= f_Nlssort('发') And r_Bitchar <= f_Nlssort('猤') Then
          v_Spell := v_Spell || 'F';
        Elsif r_Bitchar >= f_Nlssort('旮') And r_Bitchar <= f_Nlssort('腂') Then
          v_Spell := v_Spell || 'G';
        Elsif r_Bitchar >= f_Nlssort('妎') And r_Bitchar <= f_Nlssort('夻') Then
          v_Spell := v_Spell || 'H';
        Elsif r_Bitchar >= f_Nlssort('丌') And r_Bitchar <= f_Nlssort('攈') Then
          v_Spell := v_Spell || 'J';
        Elsif r_Bitchar >= f_Nlssort('咔') And r_Bitchar <= f_Nlssort('穒') Then
          v_Spell := v_Spell || 'K';
        Elsif r_Bitchar >= f_Nlssort('垃') And r_Bitchar <= f_Nlssort('擽') Then
          v_Spell := v_Spell || 'L';
        Elsif r_Bitchar >= f_Nlssort('嘸') And r_Bitchar <= f_Nlssort('椧') Then
          v_Spell := v_Spell || 'M';
        Elsif r_Bitchar >= f_Nlssort('拏') And r_Bitchar <= f_Nlssort('瘧') Then
          v_Spell := v_Spell || 'N';
        Elsif r_Bitchar >= f_Nlssort('筽') And r_Bitchar <= f_Nlssort('漚') Then
          v_Spell := v_Spell || 'O';
        Elsif r_Bitchar >= f_Nlssort('妑') And r_Bitchar <= f_Nlssort('曝') Then
          v_Spell := v_Spell || 'P';
        Elsif r_Bitchar >= f_Nlssort('七') And r_Bitchar <= f_Nlssort('裠') Then
          v_Spell := v_Spell || 'Q';
        Elsif r_Bitchar >= f_Nlssort('亽') And r_Bitchar <= f_Nlssort('鶸') Then
          v_Spell := v_Spell || 'R';
        Elsif r_Bitchar >= f_Nlssort('仨') And r_Bitchar <= f_Nlssort('蜶') Then
          v_Spell := v_Spell || 'S';
        Elsif r_Bitchar >= f_Nlssort('侤') And r_Bitchar <= f_Nlssort('籜') Then
          v_Spell := v_Spell || 'T';
        Elsif r_Bitchar >= f_Nlssort('屲') And r_Bitchar <= f_Nlssort('鶩') Then
          v_Spell := v_Spell || 'W';
        Elsif r_Bitchar >= f_Nlssort('夕') And r_Bitchar <= f_Nlssort('鑂') Then
          v_Spell := v_Spell || 'X';
        Elsif r_Bitchar >= f_Nlssort('丫') And r_Bitchar <= f_Nlssort('韻') Then
          v_Spell := v_Spell || 'Y';
        Elsif r_Bitchar >= f_Nlssort('帀') And r_Bitchar <= f_Nlssort('咗') Then
          v_Spell := v_Spell || 'Z';
        Elsif Instr('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.+-*/', v_Bitchar) > 0 Then
          v_Spell := v_Spell || v_Bitchar;
        Elsif Instr('ⅠⅡⅢⅣⅤⅥⅦⅧⅨ', v_Bitchar) > 0 Then
          v_Spell := v_Spell || Chr(Ascii(v_Bitchar) - 41664);
        Elsif Instr('ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ', v_Bitchar) > 0 Then
          v_Spell := v_Spell || Chr(Ascii(v_Bitchar) - 41856);
        Elsif Instr('Αα', v_Bitchar) > 0 Then
          v_Spell := v_Spell || 'A';
        Elsif Instr('Ββ', v_Bitchar) > 0 Then
          v_Spell := v_Spell || 'B';
        Elsif Instr('Γγ', v_Bitchar) > 0 Then
          v_Spell := v_Spell || 'G';
        End If;
        Exit When Length(v_Spell) > Nvl(v_Outmaxnum, 40) - 1;
      End Loop;
    End If;
    Return(v_Spell);
  End;

  Procedure zl_ErrorCenter(
    Err_Num In Number,
    Err_Msg In Varchar2
  ) Is
    v_Outnum Number := 0;
    v_Outmsg Varchar2(1000) := '';
    v_Temp   Varchar2(1000) := '';
  
    Cursor Cur_Ind_Cols Is
      Select Table_Name, Column_Name From All_Ind_Columns Where Instr(Err_Msg, Index_Owner || '.' || Index_Name) > 0;
  
    Cursor Cur_Con_Cols Is
      Select Table_Name, Column_Name
      From All_Cons_Columns
      Where (Owner, Constraint_Name) =
            (Select r_Owner, r_Constraint_Name
             From All_Constraints
             Where Instr(Err_Msg, Owner || '.' || Constraint_Name) > 0 And Rownum < 2);
  Begin
    If Err_Num = -1 Then
      For Row_Cols In Cur_Ind_Cols Loop
        v_Temp   := Row_Cols.Table_Name;
        v_Outmsg := v_Outmsg || '、' || Row_Cols.Column_Name;
      End Loop;
    
      v_Outmsg := '[ZLSOFT]' || v_Temp || '的(' || Substr(v_Outmsg, 2) || ')出现重复！[ZLSOFT]';
      v_Outnum := -20000;
    Elsif Err_Num = -1000 Then
      v_Outmsg := '[ZLSOFT]打开的数据表太多，必要时请系统管理员修改数据库的Open_Cursors配置。';
      v_Outnum := -20001;
    Elsif Err_Num = -1400 Or Err_Num = -1407 Then
      Select Table_Name, Column_Name
      Into v_Temp, v_Outmsg
      From All_Tab_Columns
      Where Instr(Err_Msg, '"' || Owner || '"."' || Table_Name || '"."' || Column_Name || '"') > 0 And Rownum < 2;
      v_Outmsg := '[ZLSOFT]' || v_Temp || '的(' || v_Outmsg || ')必须输入！[ZLSOFT]';
      v_Outnum := -20002;
    Elsif Err_Num = -1401 Then
      v_Outmsg := '[ZLSOFT]由于赋予的值超过了列宽限制，导致增加或更新失败。[ZLSOFT]';
      v_Outnum := -20003;
    Elsif Err_Num = -2290 Then
      Select Table_Name, Search_Condition
      Into v_Temp, v_Outmsg
      From All_Constraints
      Where Instr(Err_Msg, Owner || '.' || Constraint_Name) > 1 And Rownum < 2;
    
      If Instr(v_Outmsg, 'IS NOT NULL') > 0 Then
        v_Outmsg := '[ZLSOFT]' || v_Temp || ' 的 ' || Replace(v_Outmsg, 'IS NOT NULL', '必须输入！') || '[ZLSOFT]';
        v_Outnum := -20004;
      Else
        v_Outmsg := Err_Msg;
        v_Outnum := -20999;
      End If;
    Elsif Err_Num = -2292 Then
      Select Table_Name
      Into v_Temp
      From All_Constraints
      Where Instr(Err_Msg, Owner || '.' || Constraint_Name) > 1 And Rownum < 2;
    
      For Row_Cols In Cur_Con_Cols Loop
        v_Outmsg := v_Outmsg || '、' || Row_Cols.Column_Name;
      End Loop;
    
      v_Outmsg := '[ZLSOFT]该记录在 ' || v_Temp || ' 中已经使用,' || Chr(13) || '不能删除或修改(' || Substr(v_Outmsg, 2) || ')[ZLSOFT]';
      v_Outnum := -20005;
    Else
      v_Outmsg := Err_Msg;
      v_Outnum := -20999;
    End If;
    Raise_Application_Error(v_Outnum, Substr(v_Outmsg, 1, 100));
  End zl_ErrorCenter;

  --从文档中提取的编辑记录
  Function f_Geteditlist(
    Content_In In Xmltype
	) Return t_Editlist As
    --提取文档编辑、签名及修订记录,返回格式可能如下（独立文档SUBIID为空）：
    --  Subiid  Subaid  编辑人  编辑时间             签名 审订签名
    --  AAAAAA  AID     Null    Null                 0    0     第一条用于表示创建记录
    --  AAAAAA  AID     张险华  2012-05-31 12:01:02  1    0
    --  AAAAAA  AID     韩洪    2012-05-31 12:05:02  0    0
    --  AAAAAA  AID     韩洪    2012-05-31 12:06:02  1    1
    --  AAAAAA  AID     韩洪    2012-05-31 12:07:02  0    0
    --  AAAAAA  AID     张险华  2012-05-31 12:08:02  0    0
    --  AAAAAA  AID     张险华  2012-05-31 12:09:02  1    1
    --  BBBBBB............
    Content_c Clob;
    Xcdoc     Xmldom.Domdocument;
  
    Targetdoc Dbms_Xmldom.Domdocument;
  
    Signlist     Xmldom.Domnodelist;
    l_s          Number;
    n_Isnull     Number;
    Signname     Varchar2(64);
    Signtime     Date;
    Isaduit      Number(1);
    Xxdoc        Xmltype;
    Xaudit       Xmltype;
    Xa_Text      Xmltype;
    Textlist     Xmldom.Domnodelist;
    l_t          Number;
    Starttime    Date;
    Aftertime    Date;
    Aduitname    Varchar2(64);
    Aduittime    Date;
    Revisiontime Varchar2(20);
    Ts_Editlist  t_Editlist;
    Ta_Editlist  t_Editlist;
    r_Editlist   t_Editlist := t_Editlist();
    
    Function Sortbytime(t_e t_Editlist) Return t_Editlist Is
      Tm_List t_Editlist := t_Editlist();
    Begin
      For Rs In (Select * From Table(Cast(t_e As t_Editlist)) A Order By a.编辑时间) Loop
        Tm_List.Extend;
        Tm_List(Tm_List.Count) := t_Edits( Rs.编辑人, Rs.编辑时间, Rs.签名, Rs.审订签名);
      End Loop;
      Return Tm_List;
    End Sortbytime;
    
  Begin
    If Content_In Is Null Then
      r_Editlist := t_Editlist();
      Return r_Editlist;
    End If;
    --图片可能会超过64K,单节点超过64K时Newdomdocument会宕机,Newdomdocument(clob)方式会偶尔导致"通信通道文件结束"
    Content_c := Xml2clob(Content_In);
    
    --独立文档,直接给文档赋值
    Xcdoc   := Xmldom.Newdomdocument(Content_c);
    
    Signlist    := Xmldom.Getelementsbytagname(Xcdoc, 'signature');
    l_s         := Xmldom.Getlength(Signlist);
    l_s         := Nvl(l_s, 0);
    Ts_Editlist := t_Editlist();
    
    --遍历所有签名记录
    For L In 0 .. l_s - 1 Loop
      --提取签名位；1-签名位；0-真实签名
      n_Isnull := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Signlist, L)), 'isnull');
      
      If Nvl(n_Isnull, 0) = 0 Then
        --如果不是签名位，那就是一个真实签名
        --displayinfo 签名显示信息
        Signname := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Signlist, L)), 'displayinfo');
        If Signname Is Not Null Then
          --“签名显示信息”不为空
          Select Substr(Signname, 1, Decode(Instr(Signname, ','), 0, Length(Signname) + 1, Instr(Signname, ',')) - 1)
          Into Signname
          From Dual;
          --signtime 签名时间
          Revisiontime := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Signlist, L)), 'signtime');
          Signtime     := To_Date(Revisiontime, 'yyyy-mm-dd hh24:mi:ss');
          --isaudit 审签标记
          Isaduit      := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Signlist, L)), 'isaudit');
          Isaduit      := Nvl(Isaduit, 0);
          Ts_Editlist.Extend;
          Ts_Editlist(Ts_Editlist.Count) := t_Edits( Signname, Signtime, 1, Isaduit);
        End If;
      End If;
    End Loop;
    
    Xxdoc := Xmldom.Getxmltype(Xcdoc);
    --遍历所有以新增或删除为标记的修订记录
    --ratag 修订新增标记,取值为系统登录账号名称;  rdtag 修订删除标记,取值为系统登录账号名称
    Xa_Text     := Xxdoc.Extract('//*[@ratag!=""]|//*[@rdtag!=""]');
    Xaudit      := Xmltype('<root></root>');
    Xaudit      := Xaudit.Appendchildxml('/root', Xa_Text);
    Textlist    := Xmldom.Getchildnodes(Xmldom.Getfirstchild(Xmldom.Makenode(Xmldom.Newdomdocument(Xaudit))));
    l_t         := Xmldom.Getlength(Textlist);
    Ta_Editlist := t_Editlist();
    For L In 0 .. l_t - 1 Loop
      --ratag 修订新增标记,取值为系统登录账号名称
      Aduitname := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Textlist, L)), 'ratag');
      If Nvl(Aduitname, 'a') = 'a' Then
        --从ratag取不到值，说明是删除记录
        Revisiontime := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Textlist, L)), 'rdtime');
        Aduittime    := To_Date(Revisiontime, 'yyyy-mm-dd hh24:mi:ss');
        Aduitname    := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Textlist, L)), 'rdtag');
      Else
        Revisiontime := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Textlist, L)), 'ratime');
        Aduittime    := To_Date(Revisiontime, 'yyyy-mm-dd hh24:mi:ss');
      End If;
      Ta_Editlist.Extend;
      Ta_Editlist(Ta_Editlist.Count) := t_Edits( Aduitname, Aduittime, 0, 0);
    End Loop;
    --先按时间排序
    Ts_Editlist := Sortbytime(Ts_Editlist);
    Ta_Editlist := Sortbytime(Ta_Editlist);
    --虚拟第一条用于表示创建记录
    r_Editlist.Extend;
    r_Editlist(r_Editlist.Count) := t_Edits( Null,
                                            To_Date('1945-08-06 09:16:02', 'yyyy-mm-dd hh24:mi:ss'), 0, 0);
    
    --当次签名与下次签名之间被认定为修订记录,重新生成编辑列表
    For L In 1 .. Ts_Editlist.Count Loop
      Starttime := Ts_Editlist(L).编辑时间;
      If L = Ts_Editlist.Count Then
        --只是作为审订时间在两次签名时间之间的判断参照，当循环到最后一次签名时，此参照失去意义,所以赋值可以大于当前系统时间
        Aftertime := Sysdate + 1;
      Else
        Aftertime := Ts_Editlist(L + 1).编辑时间;
      End If;
      r_Editlist.Extend;
      r_Editlist(r_Editlist.Count) := Ts_Editlist(L);
      
      Aduitname := 'A';
      For N In 1 .. Ta_Editlist.Count Loop
        Aduittime := Ta_Editlist(N).编辑时间;
        If Aduittime Between Starttime And Aftertime Then
          Starttime := Aduittime;
          If Aduitname <> Ta_Editlist(N).编辑人 Then
            --不同人的修订记录或第一条修订记录
            Aduitname := Ta_Editlist(N).编辑人;
            r_Editlist.Extend;
            r_Editlist(r_Editlist.Count) := t_Edits( Aduitname, Aduittime, 0, 0);
          Else
            --同一人不同时间多处修订,只取最后一次时间
            r_Editlist(r_Editlist.Count).编辑时间 := Aduittime;
          End If;
        Elsif Aduittime>Aftertime then
          --因为修订记录经过时间排序，如果不在签名记录之间则是下次签名后的修订
          Exit;
        End If;
      End Loop;
    End Loop;
  
    If Not Xmldom.Isnull(Xcdoc) Then
      Xmldom.Freedocument(Xcdoc);
    End If;
  
    If Not Xmldom.Isnull(Targetdoc) Then
      Xmldom.Freedocument(Targetdoc);
    End If;
  
    Return r_Editlist;
  End f_Geteditlist;

  Function f_Getlastedit(
    Content_In In Xmltype
	) Return t_Editlist As
    t_List   t_Editlist := t_Editlist();
    r_List   t_Editlist := t_Editlist();
    
    Function Lastlist(
      t_e       t_Editlist
    ) Return t_Editlist Is
      Tm_List t_Editlist := t_Editlist();
    Begin
      For Rs In (Select * From Table(Cast(t_e As t_Editlist)) A  Order By a.编辑时间 Desc) Loop
        Tm_List.Extend;
        Tm_List(Tm_List.Count) := t_Edits( Rs.编辑人, Rs.编辑时间, Rs.签名, Rs.审订签名);
        Return Tm_List;
      End Loop;
    End Lastlist;
    
  Begin
    Select f_Geteditlist(Content_In) Into t_List From Dual;
    r_List.Extend;
    r_List(r_List.Count) := Lastlist(t_List) (1);
    Return r_List;
  End f_Getlastedit;
  
  
  --设置XML的节点值,当节点为<ele></ele>此类闭合节点时，Updatexml函数无效
  Procedure p_Set_Elementtext(
    Texture In Out Xmltype,
    Ename   In Varchar2,
    Eaname  In Varchar2,
    Eatext  In Varchar2,
    Etext   In Varchar2
  ) Is
    --参数：     Texture 操作的XML
    --           Ename 需要设置的节点名称
    --           Eaname 需要设置的节点内属性名称，用于精确定位
    --           Eatext 需要设置的节点内属性值，用于精确定位
    --           Ttext 需要设置的元素值
    x_Dom   Xmldom.Domdocument;
    x_Nlist Xmldom.Domnodelist;
    x_Text  Xmldom.Domnode;
    x_Node  Xmldom.Domnode;
    n_Len   Number;
    v_Val   Varchar2(2000);
    Procedure Freeall Is
    Begin
      If Not Xmldom.Isnull(x_Text) Then
        Xmldom.Freenode(x_Text);
      End If;
    
      If Not Xmldom.Isnull(x_Node) Then
        Xmldom.Freenode(x_Node);
      End If;
    
      If Not Xmldom.Isnull(x_Dom) Then
        Xmldom.Freedocument(x_Dom);
      End If;
    End Freeall;
  Begin
    If Texture Is Null Then
      Return;
    End If;
  
    x_Dom   := Xmldom.Newdomdocument(Texture);
    x_Nlist := Xmldom.Getelementsbytagname(x_Dom, Ename);
    n_Len   := Xmldom.Getlength(x_Nlist);
    For I In 0 .. n_Len - 1 Loop
      If Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(x_Nlist, I)), Eaname) = Eatext Then
        For J In 0 .. Xmldom.Getlength(Xmldom.Getchildnodes(Xmldom.Item(x_Nlist, I))) - 1 Loop
          x_Node := Xmldom.Item(Xmldom.Getchildnodes(Xmldom.Item(x_Nlist, I)), J);
          If Xmldom.Getnodetype(x_Node) = Xmldom.Text_Node Then
            --找到文本节点
            v_Val := Xmldom.Getnodevalue(x_Node);
            Exit;
          Else
            v_Val := Null;
          End If;
        End Loop;
      
        If v_Val Is Null Then
          x_Text := Xmldom.Makenode(Xmldom.Createtextnode(x_Dom, Etext));
          x_Text := Xmldom.Importnode(x_Dom, x_Text, True);
          x_Node := Xmldom.Appendchild(Xmldom.Item(x_Nlist, I), x_Text);
        Else
          Xmldom.Setnodevalue(x_Node, Etext);
        End If;
        Texture := Xmldom.Getxmltype(x_Dom);
        Freeall;
        Return;
      End If;
    End Loop;
  
    Freeall;
  End p_Set_Elementtext;

  --根据子文档ID提取子文档当前状态
  Function f_Get_Docstatus(
    Content_In In Xmltype
  ) Return Varchar2 Is
    n_Sign   Number;
    n_Audit  Number;
    v_Editor Varchar2(200);
    v_n      Varchar2(20);
  Begin
    For Rs In (Select *
               From Table(Cast((Select b_PACS_RptPublic.f_Geteditlist(Content_In) From Dual) As t_Editlist))
               Order By 编辑时间 Asc) Loop
      n_Sign   := Rs.签名;
      n_Audit  := Rs.审订签名;
      v_Editor := Rs.编辑人;
    End Loop;
    If n_Sign = 0 And n_Audit = 0 And v_Editor Is Null Then
      v_n := '编辑中';
    Elsif n_Sign = 1 And n_Audit = 0 Then
      v_n := '已签名';
    Elsif n_Sign = 0 And n_Audit = 0 And v_Editor Is Not Null Then
      v_n := '审订中';
    Elsif n_Sign = 1 And n_Audit = 1 Then
      v_n := '已审签';
    End If;
    Return v_n;
  End f_Get_Docstatus;

  Function f_If_Intersect
  (
    Str1 Varchar2,
    Str2 Varchar2
  ) Return Number As
    n_Num Number;
  Begin
  
    Select Count(*)
    Into n_Num
    From (Select a.Column_Value Value
           From Table(b_PACS_RptPublic.f_Str2list(Str1, ',')) A
           Intersect
           Select b.Column_Value Value From Table(b_PACS_RptPublic.f_Str2list(Str2, ',')) B);
  
    Return n_Num;
  End;

Begin
  -- Initialization
  Null;

End b_PACS_RptPublic;
/


--影像报告原型管理(---定义部分---)***************************************************
CREATE OR REPLACE Package b_PACS_RptCommon Is
  Type t_Refcur Is Ref Cursor;

  --获取预备提纲>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Pre_Outline(
    Val Out t_Refcur
	);

  --元素分类>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Eleclass(
    Val Out t_Refcur
	);

  --原型片段>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetype_Fragments(
    Val           Out t_Refcur,
	Aid_In 影像报告原型片段.原型ID%Type
	);

  --原型列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetypelist_By_Id(
    Val           Out t_Refcur,
	Id_In 影像报告原型清单.Id%Type
	);

  --原型内容>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetypelist_Content(
    Val           Out t_Refcur,
	Id_In 影像报告原型清单.Id%Type
	);

  --范文清单>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Samplelist_By_Aid(
    Val           Out t_Refcur,
	Antetypelist_Id_In Varchar2,
	Condition_In       影像报告范文清单.名称%Type,
	Author_In          影像报告范文清单.作者%Type,
	Subjects_In        影像报告范文清单.学科%Type
	);

  --获取插件配置根据插件ID获取>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Plugin_ConfigById(
    Val           Out t_Refcur,
	Id_In 影像报告插件.ID%Type
	);

  --获取插件配置根据原型清单获取>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Plugin_ConfigByAId(
    Val           Out t_Refcur,
	Aid_In 影像报告原型清单.ID%Type
	);

  --获取元素>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Element(
    Val Out t_Refcur
	);

  --获取片段列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Fragment_List(
    Val Out t_Refcur
	);

  --获取值域列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Range_List(
    Val Out t_Refcur
	);

  --获取值域列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Combo_List(
    Val Out t_Refcur
	);

  --获取原型片段根据原型ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_FragmentDirectory_ByAid(
    Val           Out t_Refcur,
	Aid_In 影像报告原型片段.原型ID%Type
	);

  --根据原型id获取片段数据
  Procedure p_Get_FragmentData_ByAid(
    Val           Out t_Refcur,
	Aid_In 影像报告原型片段.原型ID%Type
	);

  --获取数据表的最后更新时间>>>>>>>>>>>>>>>>>>>>>>>>>>>
  procedure p_Get_Data_Last_Edit_Time(
    Val           Out t_Refcur,
	Table_Name_In Varchar2
	);

  --获取片段列表根据上级ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Fragment_List_Bypid(
    Val           Out t_Refcur,
	Pid_In 影像报告片段清单.上级ID%Type
	);

  --获取片段列表根据节点类型>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Fragment_List_Byleaf(
    Val           Out t_Refcur,
	Leaf_In 影像报告片段清单.节点类型%Type
	);

  --获取值域信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Range_List_Byid(
    Val           Out t_Refcur,
	Id_In 影像报告值域清单.Id%Type
	);

  --根据元素ID获取值域ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Getelementrid_By_Eid(
    Val           Out t_Refcur,
	Eid_In 影像报告元素清单.Id%Type
	);

  --获取计量单位列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_GetMasure_UnitList(
    Val Out t_Refcur
	);

  --获取文档种类信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Doc_Kind(
    Val Out t_Refcur
	);

  --功能：获取所有学科信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Subjects(
    Val Out t_Refcur
	);

  --查看是否存在相应的编码或者名称(用于导入导出)>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_If_Exits_Doc_Kinds(
    Val           Out t_Refcur,
	编码_In      Varchar2,
	名称_In      Varchar2,
	Tablename_In Varchar2
	);

  --是否存在相同的ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_If_Exist_Id(
    Val           Out t_Refcur,
	Id_In        Number,
	Tablename_In Varchar2
	);

  --通过名称获取ID信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Id_By_Title(
    Val           Out t_Refcur,
	名称_In      Varchar2,
	Tablename_In Varchar2,
	Type_In      Varchar2
	);

  --通过简称片段清单
  procedure p_Get_FragmentSampleName(
    Val           Out t_Refcur,
	简称_In Varchar2
	);

  --更新ID对应的片段内容
  procedure p_Update_PhraseContent(
    Id_In      影像报告片段清单.ID%type,
	Name_In		影像报告片段清单.名称%Type,
	Content_In Varchar2
	);
  --获取原型ID对应的第一层片段节点
  procedure p_Get_FragmentData_LevelOne(
    Val           Out t_Refcur,
	AId_In 影像报告原型清单.ID%type
	);

  -- 获取片段的下层节点
  procedure p_GetFragmentDataListByFID(
    Val           Out t_Refcur,
	FId_In 影像报告片段清单.ID%type
	);
end b_PACS_RptCommon;
/



--*************************************************************************************
--*                  影像报告原型管理(---实现部分---)                                                        *
--*************************************************************************************
CREATE OR REPLACE Package Body b_PACS_RptCommon Is
  -- 功    能：该方法只用于演示...

  --获取预备提纲>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Pre_Outline(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(ID) As ID, a.编码, a.名称, a.说明
        From 影像报告预备提纲 A
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --元素分类>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Eleclass(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(A.ID) As ID,
             A.编码,
             A.名称,
             A.说明,
             RawToHex(A.上级ID) 上级ID
        From 影像报告元素分类 A
       Order By 编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --原型片段>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetype_Fragments(
    Val           Out t_Refcur,
	Aid_In 影像报告原型片段.原型ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(A.片段ID) As 片段ID
        From 影像报告原型片段 A
       Where a.原型ID = Aid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --原型清单>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetypelist_By_Id(
    Val           Out t_Refcur,
	Id_In 影像报告原型清单.Id%Type
	) As
  Begin
    Open Val For
      Select /*+rule*/
       RawToHex(A.ID) As ID,
       a.种类,
       a.编码,
       a.名称,
       a.说明,
       a.可否重置页面 As 页面重置,
       a.可否重置格式 As 格式重置,
       Extractvalue(b.Column_Value, '/root/print_hf_mode') Printhfmode,
       Extractvalue(b.Column_Value, '/root/print_follow_pages') Printfollowpages,
       Nvl(Extractvalue(b.Column_Value, '/root/print_limit'), 0) Printlimit,
       (Nvl(a.控制选项, XmlType('<NULL/>'))).GetClobVal() As 控制选项,
       a.创建人,
       a.创建时间,
       a.修改人,
       a.修改时间,
       a.是否禁用,
       A.分组
        From 影像报告原型清单 A,
             Table(Xmlsequence(Extract(a.控制选项, '/root'))) B
       Where a.Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --原型内容>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetypelist_Content(
    Val           Out t_Refcur,
	Id_In 影像报告原型清单.Id%Type
	) As
  Begin
    Open Val For
      Select (Nvl(a.内容, XmlType('<ZLXML/>'))).GetClobVal() As 内容
        From 影像报告原型清单 A
       Where a.Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --通过原型ID获得相应的范文信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Samplelist_By_Aid(
    Val           Out t_Refcur,
	Antetypelist_Id_In Varchar2,
	Condition_In       影像报告范文清单.名称%Type,
	Author_In          影像报告范文清单.作者%Type,
	Subjects_In        影像报告范文清单.学科%Type
	) As
  Begin
    --直接获取该原型下的范文列表
    If Length(Antetypelist_Id_In) > 30 Then
      Open Val For
        Select /*+rule*/
         RawToHex(A.ID) as ID,
         a.名称,
         a.作者,
         a.说明,
         a.学科,
         a.编号,
         a.标签,
         a.是否私有
          From 影像报告范文清单 A
         Where a.原型ID = Hextoraw(Antetypelist_Id_In)
           And (a.作者 = Author_In Or (a.学科 Is Null And a.是否私有 = 0) Or
               Subjects_In Is Null Or
               (a.学科 Is Not Null And
               b_PACS_RptPublic.f_If_Intersect(a.学科, Subjects_In) > 0 And
               a.是否私有 = 0));
    Else
      --获得一个存在原型信息的范文树形结构
      Open Val For
        Select Distinct a.分组 As ID,
                        a.分组 As 名称,
                        '' as 说明,
                        '' As 原型ID,
                        'category' As 类型,
                        '' As 作者,
                        '' As 学科,
                        Null as 最后编辑时间,
                        '' As 标签,
                        0 As 是否私有,
                        0 As Imgindex
          From 影像报告原型清单 A
         Where a.种类 = Antetypelist_Id_In
           And Exists
         (Select ID From 影像报告范文清单 C Where c.原型ID = a.Id)
           And a.分组 Is Not Null
        Union
        Select m.*
          From (Select RawToHex(B.ID) As ID,
                       b.名称,
                       b.说明,
                       b.分组 As 原型ID,
                       'antetype' As 类型,
                       '' As 作者,
                       '' As 学科,
                       Null as 最后编辑时间,
                       '' As 标签,
                       0 As 是否私有,
                       0 As Imgindex
                  From 影像报告原型清单 B
                 Where b.种类 = Antetypelist_Id_In
                   And Exists (Select ID
                          From 影像报告范文清单 C
                         Where c.原型ID = b.Id)
                 Order By b.编码) M
        
        Union All
        Select n.*
          From (Select /*+rule*/
                 RawToHex(A.ID) As ID,
                 a.名称,
                 a.说明,
                 RawToHex(A.原型ID) As 原型ID,
                 'sample' As 类型,
                 a.作者,
                 a.学科,
                 a.最后编辑时间,
                 a.标签,
                 a.是否私有,
                 Decode(a.是否私有, 1, 2, 1) As Imgindex
                  From 影像报告范文清单 A, 影像报告原型清单 C
                 Where a.原型ID = c.Id
                   And c.种类 = Antetypelist_Id_In
                   And ((a.名称 Like '%' || Condition_In || '%' And
                       Condition_In Is Not Null) Or Condition_In Is Null)
                   And (a.作者 = Author_In Or (a.学科 Is Null And a.是否私有 = 0) Or
                       Subjects_In Is Null Or
                       (a.学科 Is Not Null And
                       b_PACS_RptPublic.f_If_Intersect(a.学科, Subjects_In) > 0 And
                       a.是否私有 = 0))
                 Order By a.编号, a.名称) N;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取插件配置根据插件ID获取>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Plugin_ConfigById(
    Val           Out t_Refcur,
	Id_In 影像报告插件.ID%Type
	) As
    v_Sql Varchar2(1000);
  Begin
    If (Id_In Is Not Null) Then
      Begin
        v_Sql := 'Select RawToHex(T.ID) as ID, t.编码, t.说明, t.名称, t.类名, t.库名, t.是否禁用,t.种类,T.显示样式  From 影像报告插件 T Where t.Id =:Id_In And Rownum = 1';
      
        Open Val For v_Sql
          Using Id_In;
      End;
    Else
      Begin
        v_Sql := 'Select RawToHex(T.ID) as ID, t.编码, t.说明, t.名称, t.类名, t.库名, t.是否禁用,t.种类,T.显示样式 From 影像报告插件 T where 是否禁用 = 0 order by t.编码';
      
        Open Val For v_Sql;
      End;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取插件配置根据原型清单获取>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Plugin_ConfigByAId(
    Val           Out t_Refcur,
	Aid_In 影像报告原型清单.ID%Type
	) As
    v_Sql Varchar2(1000);
  Begin
    If (Aid_In Is Not Null) Then
      Begin
        v_Sql := 'Select RawToHex(T.ID) as ID, t.编码, t.说明, t.名称, t.类名, t.库名, t.是否禁用,t.种类,T.显示样式 ' ||
                 ' From 影像报告插件 T ' || ' Where t.Id in ( ' ||
                 ' Select X.pluginid from 影像报告原型清单 K, ' ||
                 '  (XMLTable(''*//pluginid''  Passing K.专用插件 Columns pluginid varchar2(32) Path ''/pluginid''))  X ' ||
                 ' Where K.id=:Aid_In) And 是否禁用 = 0' || ' Union All ' ||
                 'Select RawToHex(T.ID) as ID, t.编码, t.说明, t.名称, t.类名, t.库名, t.是否禁用,t.种类,T.显示样式 ' ||
                 ' From 影像报告插件 T ' || ' Where 是否禁用 = 0 And t.种类=0 ';
      
        Open Val For v_Sql
          Using Aid_In;
      End;
    Else
      Begin
        v_Sql := 'Select RawToHex(T.ID) as ID, t.编码, t.说明, t.名称, t.类名, t.库名, t.是否禁用,t.种类,T.显示样式 From 影像报告插件 T where 是否禁用 = 0 order by t.编码';
      
        Open Val For v_Sql;
      End;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取所有元素>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Element(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) as ID,
             RawToHex(T.分类ID) as 分类ID,
             T.编码,
             T.名称,
             T.前缀,
             T.后缀,
             T.说明,
             T.数据类型,
             T.数值形态,
             T.最小长度,
             T.最大长度,
             T.最小小数位,
             T.最大小数位,
             T.计量单位,
             (Nvl(T.扩展描述, XmlType('<NULL/>'))).GetClobVal() As 扩展描述,
             T.值域ID,
             T.值域种类
        From 影像报告元素清单 T;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取片段列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Fragment_List(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.上级ID) As 上级ID,
             t.编码,
             t.名称,
             t.说明,
             t.节点类型,
             (Nvl(t.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
             t.学科,
             t.标签,
             t.是否私有,
             t.作者,
             t.最后编辑时间,
			 (Nvl(t.适应条件, XmlType('<NULL/>'))).GetClobVal() As 适应条件
        From 影像报告片段清单 T;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取值域列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Range_List(
    Val Out t_Refcur
	) as
  begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.分类ID) As 分类ID,
             T.编码,
             T.名称,
             T.说明,
             T.数据类型,
             T.值域种类,
             (Nvl(t.值域描述, XmlType('<NULL/>'))).GetClobVal() As 值域描述,
             t.最后编辑时间
        From 影像报告值域清单 T;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end;

  --获取组句列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Combo_List(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             T.编码,
             T.名称,
             T.说明,
             T.多组,
             (Nvl(t.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
             T.编辑人,
             T.最后编辑时间,
             T.分组
        From 影像报告组句清单 T;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取原型片段目录根据原型ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_FragmentDirectory_ByAid(
    Val           Out t_Refcur,
	Aid_In 影像报告原型片段.原型ID%Type
	) As
  Begin
    Open Val For
      Select 名称,
             RawToHex(ID) As ID,
             4 as ImageIndex,
             RawToHex(上级ID) As 上级ID,
             '<NULL/>' As 组成,
             编码,
             节点类型,
             是否私有,
             作者,
             标签,
			 说明,
             学科
        From 影像报告片段清单
       Where ID In (Select ID
                      From 影像报告片段清单
                     Start With ID In (Select 片段ID
                                         From 影像报告原型片段
                                        Where 原型ID = Aid_In)
                    Connect By Prior 上级ID = ID)
       order by 编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取原型片段数据根据原型ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_FragmentData_ByAid(
    Val           Out t_Refcur,
	Aid_In 影像报告原型片段.原型ID%Type
	) As
  Begin
    Open Val For
      With TabFragmentId As
       (Select 片段ID From 影像报告原型片段 Where 原型ID = Aid_In)
      Select RawToHex(ID) As ID,
             RawToHex(上级ID) As 上级ID,
             编码,
             名称,
             说明,
             节点类型,
             (Nvl(t.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
             学科,
             标签,
             是否私有,
             作者,
             最后编辑时间,
             EXTRACTValue( 适应条件, '/Root/Rad') 项目,
             EXTRACTValue( 适应条件, '/Root/Part') 部位, 
             EXTRACTValue( 适应条件, '/Root/Kind') 类别,
             EXTRACTValue( 适应条件, '/Root/Sex') 性别,
             0 as 提纲状态,
             0 as 适应状态
        From 影像报告片段清单 t
       Where Id Not In (Select 片段ID From TabFragmentId)
       Start With ID In (Select 片段ID From TabFragmentId)
      Connect By Prior 上级ID = ID
      Union All
      Select RawToHex(ID) As ID,
             RawToHex(上级ID) As 上级ID,
             编码,
             名称,
             说明,
             节点类型,
             (Nvl(t.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
             学科,
             标签,
             是否私有,
             作者,
             最后编辑时间,
             EXTRACTValue( 适应条件, '/Root/Rad') 项目,
             EXTRACTValue( 适应条件, '/Root/Part') 部位, 
             EXTRACTValue( 适应条件, '/Root/Kind') 类别,
             EXTRACTValue( 适应条件, '/Root/Sex') 性别,
             0 as 提纲状态,
             0 as 适应状态
        From 影像报告片段清单 t
       Start With ID In (Select 片段ID From TabFragmentId)
      Connect By Prior ID = 上级ID
       order by 编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取数据表的最后更新时间>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Data_Last_Edit_Time(
    Val           Out t_Refcur,
	Table_Name_In Varchar2
	) As
    v_sql Varchar2(4000);
  Begin
    v_sql := 'select max(最后编辑时间) maxvalue from ' || Table_Name_In;
    Open val For v_sql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取片段列表根据上级ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Fragment_List_Bypid(
    Val           Out t_Refcur,
	Pid_In 影像报告片段清单.上级ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.上级ID) As 上级ID,
             T.编码,
             T.名称,
             T.说明,
             T.节点类型,
             (Nvl(T.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
             T.学科,
             T.标签,
             T.是否私有,
             T.作者,
			 (Nvl(T.适应条件, XmlType('<NULL/>'))).GetClobVal() As 适应条件,
             T.最后编辑时间
        From 影像报告片段清单 T
       Where T.上级ID = Pid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --通过节点类型获取词句列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Fragment_List_Byleaf(
    Val           Out t_Refcur,
	Leaf_In 影像报告片段清单.节点类型%Type
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.上级ID) As 上级ID,
             T.编码,
             T.名称,
             T.说明,
             T.节点类型,
             (Nvl(T.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
             T.学科,
             T.标签,
             T.是否私有,
             T.作者,
             T.最后编辑时间
        From 影像报告片段清单 T
       Where t.节点类型 = Leaf_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取值域信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Range_List_Byid(
    Val           Out t_Refcur,
	Id_In 影像报告值域清单.Id%Type
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.分类ID) As 分类ID,
             T.编码,
             T.名称,
             T.说明,
             T.数据类型,
             T.值域种类,
             (Nvl(T.值域描述, XmlType('<NULL/>'))).GetClobVal() As 值域描述
        From 影像报告值域清单 T
       Where ID = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --根据元素ID获取值域ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Getelementrid_By_Eid(
    Val           Out t_Refcur,
	Eid_In 影像报告元素清单.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(A.值域ID) As 值域ID
        From 影像报告元素清单 A
       Where a.Id = Eid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --获取计量单位列表>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_GetMasure_UnitList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select 编码, 名称, 说明, 前缀 From 影像报告计量单位;
  End p_GetMasure_UnitList;

  --获取文档种类信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Doc_Kind(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select a.编码, a.名称, a.说明 From 影像报告种类 A Order By a.编码;
  End;

  --功能：获取所有学科信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Subjects(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select rawtohex(b.字典id) ID, b.编号 As 编码, b.名称, b.简码, b.说明
        From 影像字典清单 A, 影像字典内容 B
       Where a.名称 = '专业学科'
         And a.Id = b.字典id
       Order By 编码;
  End;

  --查看是否存在相应的编码或者名称(用于导入导出)>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_If_Exits_Doc_Kinds(
    Val           Out t_Refcur,
	编码_In      Varchar2,
	名称_In      Varchar2,
	Tablename_In Varchar2
	) As
    v_Type Varchar2(50);
    n_Num  Number;
    v_Sql  Varchar2(100);
  Begin
    v_Sql := 'SELECT COUNT(*) FROM ' || Tablename_In || ' WHERE 编码 =''' ||
             编码_In || ''' AND 名称 =''' || 名称_In || '''';
    Execute Immediate v_Sql
      Into n_Num;
    If n_Num > 0 Then
      v_Type := '1';
    End If;
    v_Sql := 'SELECT COUNT(*) FROM ' || Tablename_In || ' WHERE 编码 <>''' ||
             编码_In || ''' AND 名称 = ''' || 名称_In || '''';
    Execute Immediate v_Sql
      Into n_Num;
    If n_Num > 0 Then
      v_Type := '2';
    End If;
    v_Sql := 'SELECT COUNT(*) FROM ' || Tablename_In || ' WHERE 编码 =''' ||
             编码_In || ''' or 名称 =''' || 名称_In || '''';
    Execute Immediate v_Sql
      Into n_Num;
    If n_Num = 0 Then
      v_Type := '3';
    End If;
    v_Sql := 'SELECT COUNT(*) FROM ' || Tablename_In || ' WHERE 编码 =''' ||
             编码_In || ''' AND 名称 <>''' || 名称_In || '''';
    Execute Immediate v_Sql
      Into n_Num;
    If n_Num > 0 Then
      v_Type := '4';
    End If;
    Open Val For
      Select v_Type As Type From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --是否存在相同的ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_If_Exist_Id(
    Val           Out t_Refcur,
	Id_In        Number,
	Tablename_In Varchar2
	) As
    v_Sql Varchar2(100);
    n_Num Number;
  Begin
    v_Sql := 'select count(id) from ' || Tablename_In || ' where id=''' ||
             Id_In || '''';
    Execute Immediate v_Sql
      Into n_Num;
    Open Val For
      Select Decode(n_Num, 0, 0, 1) Num From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --通过名称获取ID信息>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Id_By_Title(
    Val           Out t_Refcur,
	名称_In      Varchar2,
	Tablename_In Varchar2,
	Type_In      Varchar2
	) As
    v_Id  Varchar2(50);
    v_Sql Varchar2(100);
  Begin
    If Type_In = '1' Then
      v_Sql := 'select id from ' || Tablename_In || ' where 名称=''' || 名称_In || '''';
    Else
      v_Sql := 'select 编码 from ' || Tablename_In || ' where 名称=''' || 名称_In || '''';
    End If;
    v_Id := '';
    Execute Immediate v_Sql
      Into v_Id;
    Open Val For
      Select v_Id ID From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --通过简称片段清单
  Procedure p_Get_FragmentSampleName(
	Val           Out t_Refcur,
	简称_In Varchar2
	) as
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.上级ID) As 上级ID,
             t.编码,
             t.名称,
             t.说明,
             t.节点类型,
             (Nvl(t.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
             t.学科,
             t.标签,
             t.是否私有,
             t.作者,
             t.最后编辑时间
        From 影像报告片段清单 T
      Where t.名称 LIKE '%' || 简称_In || '%';
      --Where  F_TRANS_PINYIN_CAPITAL(t.名称) LIKE '%' || 简称_In || '%';
  End p_Get_FragmentSampleName;

  --更新ID对应的片段内容
  Procedure p_Update_PhraseContent(
	Id_In      影像报告片段清单.ID%Type,
	Name_In		影像报告片段清单.名称%Type,
	Content_In Varchar2
	) as
  Begin
    Update 影像报告片段清单 t 
	Set t.组成 = Content_In, t.名称=Name_In 
	Where t.id = Id_In;
  End p_Update_PhraseContent;

  --获取原型ID对应的第一层片段节点
  Procedure p_Get_FragmentData_LevelOne(
	Val           Out t_Refcur,
	AId_In 影像报告原型清单.ID%Type
	) as
  Begin
    Open Val For
      With TabFragmentId As
       (Select 片段ID From 影像报告原型片段 Where 原型ID = Aid_In)
      Select RawToHex(ID) As ID,
             RawToHex(上级ID) As 上级ID,
             编码,
             名称,
             说明,
             节点类型,
             (Nvl(t.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
             学科,
             标签,
             是否私有,
             作者,
             最后编辑时间
        From 影像报告片段清单 t
       Where Id In (Select 片段ID From TabFragmentId);
  
  End p_Get_FragmentData_LevelOne;

  -- 获取片段的下层节点
  Procedure p_GetFragmentDataListByFID(
	Val           Out t_Refcur,
	FId_In 影像报告片段清单.ID%Type
	) as
  Begin
    Open Val For
      Select RawToHex(ID) As ID,
             RawToHex(上级ID) As 上级ID,
             编码,
             名称,
             说明,
             节点类型,
             (Nvl(t.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
             学科,
             标签,
             是否私有,
             作者,
             最后编辑时间
        From 影像报告片段清单 t
       Where 上级ID = FId_In;
  End p_GetFragmentDataListByFID;

End b_PACS_RptCommon;

/




   --影像报告参数---定义部分---)***************************************************
CREATE OR REPLACE Package b_PACS_RptParam Is
  Type t_Refcur Is Ref Cursor;

  --功能1：获得活动的参数列表
  Procedure p_GetPrograms(
    Val Out t_Refcur
	);
  --功能2：通过模块号获取影像参数信息
  Procedure p_GetParamByQum(
    Val           Out t_Refcur,
	模块_In 影像参数说明.模块%Type
	);
  --功能3：通过参数ID号获取影像参数取值信息
  Procedure p_GetParamValue(
    Val           Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type
	);
  --功能4：获得部门信息
  Procedure p_GetDepart(
    Val Out t_Refcur
	);
  --功能5：获得人员信息
  Procedure p_GetUsersInfo(
    Val Out t_Refcur
	);
  --功能6：获得机器名信息
  Procedure p_GetMachinesInfo(
    Val Out t_Refcur
	);
  --功能7：获取所有的影像参数信息
  Procedure p_GetAllParam(
    Val Out t_Refcur
	);
  --功能8：获得所有部门的所有的参数取值信息
  Procedure p_GetValueAllDepart(
    Val Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type
	);
  --功能9：得ID对应部门的所有的参数取值信息
  Procedure p_GetValueByDepart(
    Val Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type,
	部门ID_In 部门表.ID%Type
	);
  --功能9：获取所有的用户对应的参数值
  Procedure p_GetValueAllUser(
    Val Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type
	);
  --功能10：获取用户ID对应的参数信息
  Procedure p_GetValueByUser(
    Val Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type,
	用户ID_In 人员表.ID%Type
	);
  --功能11：获取所有的工作站对应的参数值
  Procedure p_GetValueAllMachine(
    Val Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type
	);
  --功能12：获取工作站名称对应的参数信息
  Procedure p_GetValueByMachine(
    Val Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type,
	机器名_In zlclients.工作站%Type
	);
  --功能13：添加参数信息
  Procedure p_AddParamValue(
    ID_In       影像参数取值.ID%Type,
    参数ID_In   影像参数取值.参数ID%Type,
    参数标识_In 影像参数取值.参数标识%Type,
    参数值_In   影像参数取值.参数值%Type
	);

  --功能14：修改参数信息
  Procedure p_EditParamValue(
    ID_In       影像参数取值.ID%Type,
    参数标识_In 影像参数取值.参数标识%Type,
    参数值_In   影像参数取值.参数值%Type
	);

  --功能15:通过ID获得参数信息
  Procedure p_GetParamByID(
    Val Out t_Refcur,
	ID_In 影像参数说明.ID%Type
	);
  --功能16：修改ID对应的参数级别
  Procedure p_ChangeAdjustByID(
    ID_In     影像参数说明.ID%Type,
	Adjust_In 影像参数说明.参数级别%Type
	);
  --功能17：获得对应参数标识的参数取值信息
  Procedure p_GetValueBySign(
    Val Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type
	);
  --功能18：修改ID对应的参数信息的默认值
  Procedure p_EditDaultValue(
    ID_In     影像参数说明.ID%Type,
	默认值_In 影像参数说明.默认值%Type);

  --功能19：通过部门获得人员信息
  Procedure p_GetUserByDID(
    Val Out t_Refcur,
	DID_In 部门人员.部门ID%Type
	);
  --功能21:通过ID获得参数取值
  Procedure p_GetParamValueByCID(
    Val Out t_Refcur,
	CID_In 影像参数取值.参数ID%Type
	);
  --功能22:通过ID获得模块号的参数取值
  Procedure p_GetValueLevel0(
    Val Out t_Refcur,
	参数ID_In   影像参数取值.参数ID%Type,
	参数标识_In 影像参数取值.参数标识%Type
	);
end b_PACS_RptParam;
/

--影像报告参数---实现部分---)***************************************************
CREATE OR REPLACE Package Body b_PACS_RptParam Is
  --create by hwei;

  --功能1：获得活动的参数列表
  Procedure p_GetPrograms(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select 序号, 标题, Decode(A.标题, '', 序号, 序号 || '-' || 标题) 名称
        From (Select Distinct (t.模块) 序号,
                              (Select y.标题
                                 From zlprograms y
                                Where to_char(y.序号) = t.模块) 标题
                From 影像参数说明 t) A;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetPrograms;
  --功能2：通过模块号获取影像参数信息
  Procedure p_GetParamByQum(
    Val Out t_Refcur,
	模块_In 影像参数说明.模块%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(PID) PID,
             t.分组,
             t.参数序号,
             t.参数名,
             t.默认值,
             t.参数级别,
             t.取值范围,
             t.启用条件,
             t.说明,
             '— —' 参数值
        From 影像参数说明 t
       Where t.模块 = 模块_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetParamByQum;

  --功能3：通过参数ID号获取影像参数取值信息
  Procedure p_GetParamValue(
    Val Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type
	) As
    paramPiont nvarchar2(50);
    paramLevel Number := -1;
    paramCount Number := -1;
  Begin
    Select a.参数级别
      Into paramLevel
      From 影像参数说明 a
     Where a.id = 参数ID_In
       And rownum <= 1;
    If paramLevel = 1 Then
      Select count(t.id)
        Into paramCount
        From 影像参数取值 t
       Where t.参数id = 参数ID_In;
      IF paramCount <> 0 THEN
        Select t.参数标识
          Into paramPiont
          From 影像参数取值 t
         Where t.参数id = 参数ID_In;
      End If;
      IF paramCount <> 0 And
         Replace(translate(paramPiont, '0123456789', '0'), '0', '') IS NULL THEN
        Open Val For
          Select RawToHex(ID) ID,
                 RawToHex(参数ID) 参数ID,
                 t.参数标识,
                 (Select a.标题
                    From zlprograms a
                   Where a.序号 = t.参数标识
                     And rownum <= 1) As 标识名称,
                 t.参数值
            From 影像参数取值 t
           Where t.参数id = 参数ID_In;
      Else
        Open Val For
          Select RawToHex(ID) ID,
                 RawToHex(参数ID) 参数ID,
                 t.参数标识,
                 t.参数标识 As 标识名称,
                 t.参数值
            From 影像参数取值 t
           Where t.参数id = 参数ID_In;
      End If;
    Elsif paramLevel = 2 Then
      Open Val For
        Select RawToHex(ID) ID,
               RawToHex(参数ID) 参数ID,
               t.参数标识,
               (Select a.名称 From 部门表 a Where a.id = t.参数标识) as 标识名称,
               t.参数值
          From 影像参数取值 t
         Where t.参数id = 参数ID_In;
    Elsif paramLevel = 3 Then
      Open Val For
        Select RawToHex(ID) ID,
               RawToHex(参数ID) 参数ID,
               t.参数标识,
               (Select a.姓名
                  From 人员表 a
                 Where a.id = t.参数标识
                   And rownum <= 1) As 标识名称,
               t.参数值
          From 影像参数取值 t
         Where t.参数id = 参数ID_In;
    Else
      Open Val For
        Select RawToHex(ID) ID,
               RawToHex(参数ID) 参数ID,
               t.参数标识,
               t.参数标识 As 标识名称,
               t.参数值
          From 影像参数取值 t
         Where t.参数id = 参数ID_In;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetParamValue;

  --功能4：获得部门信息
  Procedure p_GetDepart(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select ID, 上级ID, t.编码, t.名称 from 部门表 t Order by t.名称;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDepart;
  --功能5：获得人员信息
  Procedure p_GetUsersInfo(Val Out t_Refcur) As
  Begin
    Open Val For
      Select ID, t.编号, t.姓名, t.简码, t.身份证号
        From 人员表 t
       Order by t.姓名;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetUsersInfo;

  --功能6：获得机器名信息
  Procedure p_GetMachinesInfo(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select t.工作站, t.ip, t.部门 From zlclients t Order by t.工作站;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetMachinesInfo;

  --功能7：获取所有的影像参数信息
  Procedure p_GetAllParam(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(PID) PID,
             t.分组,
             t.参数序号,
             t.参数名,
             t.默认值,
             t.参数级别,
             t.取值范围,
             t.启用条件,
             t.说明
        From 影像参数说明 t;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetAllParam;
  --功能8：获得所有部门的所有的参数取值信息
  Procedure p_GetValueAllDepart(
    Val Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type
	) As
  Begin
    Open Val For
      Select t.id, t.参数id, t.参数标识, t.参数值
        From 影像参数取值 t
       Where t.参数id = 参数ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueAllDepart;

  --功能9：得ID对应部门的所有的参数取值信息
  Procedure p_GetValueByDepart(
    Val Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type,
	部门ID_In 部门表.ID%Type
	) As
  Begin
    Open Val For
      Select (Select RawToHex(ID) ID
                From 影像参数取值 t
               Where t.参数id = 参数ID_In
                 And t.参数标识 = s.id) As ID,
             RawToHex(参数ID_In) As 参数ID,
             s.id As 参数标识,
             (Select t.参数值
                From 影像参数取值 t
               Where t.参数id = 参数ID_In
                 And t.参数标识 = s.id) As 参数值
        From 部门表 s
       Where s.id = 部门ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueByDepart;

  --功能9：获取所有的用户对应的参数值
  Procedure p_GetValueAllUser(
    Val Out t_Refcur,
    参数ID_In 影像参数取值.参数ID%Type
	) As
  Begin
    Open Val For
      Select (Select RawToHex(ID) ID
                From 影像参数取值 t
               Where t.参数id = 参数ID_In
                 And t.参数标识 = s.id) As ID,
             RawToHex(参数ID_In) As 参数ID,
             s.id As 参数标识,
             (Select t.参数值
                From 影像参数取值 t
               Where t.参数id = 参数ID_In
                 And t.参数标识 = s.id) As 参数值
        From 人员表 s;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueAllUser;

  --功能10：获取用户ID对应的参数信息
  Procedure p_GetValueByUser(
    Val Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type,
	用户ID_In 人员表.ID%Type
	) As
  Begin
    Open Val For
      Select (Select RawToHex(ID) ID
                From 影像参数取值 t
               Where t.参数id = 参数ID_In
                 And t.参数标识 = s.id) As ID,
             RawToHex(参数ID_In) As 参数ID,
             s.id As 参数标识,
             (Select t.参数值
                From 影像参数取值 t
               Where t.参数id = 参数ID_In
                 And t.参数标识 = s.id) As 参数值
        From 人员表 s
       Where s.id = 用户ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueByUser;

  --功能11：获取所有的工作站对应的参数值
  Procedure p_GetValueAllMachine(
    Val Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type
	) As
  Begin
    Open Val For
      Select (Select RawToHex(ID) ID
                From 影像参数取值 t
               Where t.参数id = 参数ID_In
                 And t.参数标识 = s.工作站) As ID,
             RawToHex(参数ID_In) As 参数ID,
             s.工作站 As 参数标识,
             (Select t.参数值
                From 影像参数取值 t
               Where t.参数id = 参数ID_In
                 And t.参数标识 = s.工作站) As 参数值
        From zlclients s;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueAllMachine;

  --功能12：获取工作站名称对应的参数信息
  Procedure p_GetValueByMachine(
    Val Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type,
	机器名_In zlclients.工作站%Type
	) As
  Begin
    Open Val For
      Select (Select RawToHex(ID) ID
                From 影像参数取值 t
               Where t.参数id = 参数ID_In
                 And t.参数标识 = s.工作站) as ID,
             RawToHex(参数ID_In) as 参数ID,
             s.工作站 as 参数标识,
             (Select t.参数值
                From 影像参数取值 t
               Where t.参数id = 参数ID_In
                 And t.参数标识 = s.工作站) as 参数值
        From zlclients s
       Where s.工作站 = 机器名_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueByMachine;
  --功能13：添加参数信息
  Procedure p_AddParamValue(
    ID_In       影像参数取值.ID%Type,
    参数ID_In   影像参数取值.参数ID%Type,
    参数标识_In 影像参数取值.参数标识%Type,
    参数值_In   影像参数取值.参数值%Type
	) As
  Begin
    Insert Into 影像参数取值 t
      (ID, 参数ID, 参数标识, 参数值)
    ValueS
      (ID_In, 参数ID_In, 参数标识_In, 参数值_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddParamValue;

  --功能14：修改参数信息
  Procedure p_EditParamValue(
    ID_In       影像参数取值.ID%Type,
    参数标识_In 影像参数取值.参数标识%Type,
    参数值_In   影像参数取值.参数值%Type
	) As
  Begin
    Update 影像参数取值 t
       Set 参数标识 = 参数标识_In, 参数值 = 参数值_In
     Where t.id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditParamValue;
  --功能15:通过ID获得参数信息
  Procedure p_GetParamByID(
    Val Out t_Refcur,
	ID_In 影像参数说明.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(PID) PID,
             t.分组,
             t.参数序号,
             t.参数名,
             t.默认值,
             t.参数级别,
             t.取值范围,
             t.启用条件,
             t.说明
        From 影像参数说明 t
       Where t.id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetParamByID;
  --功能16：修改ID对应的参数级别
  Procedure p_ChangeAdjustByID(
    ID_In     影像参数说明.ID%Type,
	Adjust_In 影像参数说明.参数级别%Type) As
  Begin
    Delete From 影像参数取值 a Where a.参数id = ID_In;
    Update 影像参数说明 t Set t.参数级别 = Adjust_In Where t.id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_ChangeAdjustByID;

  --功能17：获得对应参数标识的参数取值信息
  Procedure p_GetValueBySign(
    Val Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type
	) As
  Begin
    Open Val For
      Select t.id, t.参数id, t.参数标识, t.参数值
        From 影像参数取值 t
       Where t.参数id = 参数ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueBySign;

  --功能18：修改ID对应的参数信息的默认值
  Procedure p_EditDaultValue(
    ID_In     影像参数说明.ID%Type,
	默认值_In 影像参数说明.默认值%Type) As
  Begin
    Update 影像参数说明 t Set t.默认值 = 默认值_In Where t.id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditDaultValue;

  --功能19：通过部门获得人员信息
  Procedure p_GetUserByDID(
    Val Out t_Refcur,
	DID_In 部门人员.部门ID%Type
	) As
  Begin
    Open Val For
      Select ID, t.编号, t.姓名, t.简码, t.身份证号
        From 人员表 t
       Where t.id In
             (Select a.人员id From 部门人员 a Where a.部门id = DID_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetUserByDID;
  --功能20：通过参数ID号获取影像参数取值信息
  Procedure p_GetParamValue1(
    Val Out t_Refcur,
	参数ID_In 影像参数取值.参数ID%Type
	) As
    paramLevel number := -1;
    paramPiont nvarchar2(50);
    paramCount number := -1;
  Begin
    Select a.参数级别
      Into paramLevel
      From 影像参数说明 a
     Where a.id = 参数ID_In
       And rownum <= 1;
    If paramLevel = 1 then
      Select Count(t.id)
        into paramCount
        From 影像参数取值 t
       Where t.参数id = 参数ID_In;
      Select t.参数标识
        into paramPiont
        From 影像参数取值 t
       Where t.参数id = 参数ID_In;
      IF paramCount <> 0 and
         Replace(translate(paramPiont, '0123456789', '0'), '0', '') IS NULL THEN
        Open Val For
          Select RawToHex(ID) ID,
                 RawToHex(参数ID) 参数ID,
                 (Select a.标题
                    From zlprograms a
                   Where a.序号 = t.参数标识
                     And rownum <= 1) As 参数标识,
                 t.参数值
            From 影像参数取值 t
           Where t.参数id = 参数ID_In;
      Else
        Open Val For
          Select RawToHex(ID) ID,
                 RawToHex(参数ID) 参数ID,
                 t.参数标识,
                 t.参数值
            From 影像参数取值 t
           Where t.参数id = 参数ID_In;
      End if;
    Elsif paramLevel = 2 Then
      Open Val For
        Select RawToHex(ID) ID,
               RawToHex(参数ID) 参数ID,
               (Select a.名称
                  From 部门表 a
                 Where a.id = t.参数标识
                   And rownum <= 1) As 参数标识,
               t.参数值
          From 影像参数取值 t
         Where t.参数id = 参数ID_In;
    Elsif paramLevel = 3 Then
      Open Val For
        Select RawToHex(ID) ID,
               RawToHex(参数ID) 参数ID,
               (Select a.姓名
                  From 人员表 a
                 Where a.id = t.参数标识
                   And rownum <= 1) As 参数标识,
               t.参数值
          From 影像参数取值 t
         Where t.参数id = 参数ID_In;
    Else
      Open Val For
        Select RawToHex(ID) ID,
               RawToHex(参数ID) 参数ID,
               t.参数标识,
               t.参数值
          From 影像参数取值 t
         Where t.参数id = 参数ID_In;
    End If;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetParamValue1;
  --功能21:通过ID获得参数取值
  Procedure p_GetParamValueByCID(
    Val Out t_Refcur,
	CID_In 影像参数取值.参数ID%Type
	) As
  Begin
    Open Val For
      Select t.参数标识, t.参数值
        From 影像参数取值 t
       Where t.参数id = CID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetParamValueByCID;

  --功能22:通过ID获得模块号的参数取值
  Procedure p_GetValueLevel0(
    Val Out t_Refcur,
	参数ID_In   影像参数取值.参数ID%Type,
	参数标识_In 影像参数取值.参数标识%Type
	) As
  Begin
    Open Val For
      Select t.参数标识, t.参数值
        From 影像参数取值 t
       Where t.参数id = 参数ID_In
         And t.参数标识 = 参数标识_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueLevel0;
End b_PACS_RptParam;
/


--影像报告元素值域(---定义部分---)***************************************************
CREATE OR REPLACE Package b_PACS_RptElement Is
  --Create By Hwei;
  --2014/11/25
  Type t_Refcur Is Ref Cursor;

  Procedure p_GetElementClassList(
    Val Out t_Refcur
	);

  --2.功  能：新增影像报告元素分类信息
  Procedure p_AddElementClass(
    ID_In   In 影像报告元素分类.ID%Type,
	编码_In In 影像报告元素分类.编码%Type,
	名称_In In 影像报告元素分类.名称%Type,
	说明_In In 影像报告元素分类.说明%Type,
	上级ID_In In 影像报告元素分类.上级ID%Type
	);

  --3.功  能：修改影像报告元素分类信息
  Procedure p_EditElementClass(
    ID_In   In 影像报告元素分类.ID%Type,
	编码_In In 影像报告元素分类.编码%Type,
	名称_In In 影像报告元素分类.名称%Type,
	说明_In In 影像报告元素分类.说明%Type,
	上级ID_In In 影像报告元素分类.上级ID%Type
	);

  --4.功  能：删除影像报告元素分类信息
  Procedure p_DelelEmentClass(
    ID_In In 影像报告元素分类.ID%Type
	);

  --5.功  能：获得分类对应的影像报告值域信息列表
  Procedure p_GetRangeByClass(
    Val           Out t_Refcur,
	分类ID_In In 影像报告值域清单.分类ID%Type
	);

  --6.功  能：获得ID对应的影像报告值域信息
  Procedure p_GetRangeByID(
    Val           Out t_Refcur,
	ID_In In 影像报告值域清单.ID%Type
	);

  --7.功  能：新增影像报告值域信息
  Procedure p_AddRange(
    ID_In          In 影像报告值域清单.ID%Type,
	分类ID_In    In 影像报告值域清单.分类ID%Type,
	编码_In       In 影像报告值域清单.编码%Type,
	名称_In       In 影像报告值域清单.名称%Type,
	说明_In       In 影像报告值域清单.说明%Type,
	数据类型_In In 影像报告值域清单.数据类型%Type,
	值域种类_In In 影像报告值域清单.值域种类%Type,
	值域描述_In In Varchar2);

  --8.功  能：修改影像报告值域信息
  Procedure p_EditRange(
    ID_In         In 影像报告值域清单.ID%Type,
	分类ID_In   In 影像报告值域清单.分类ID%Type,
	编码_In       In 影像报告值域清单.编码%Type,
	名称_In       In 影像报告值域清单.名称%Type,
	说明_In       In 影像报告值域清单.说明%Type,
	数据类型_In In 影像报告值域清单.数据类型%Type,
	值域种类_In In 影像报告值域清单.值域种类%Type,
	值域描述_In In Varchar2
	);

  --9.功  能：删除影像报告值域信息
  Procedure p_DelRange(
    ID_In In 影像报告值域清单.ID%Type
	);

  --10.功  能：获得分类对应的影像报告元素列表
  Procedure p_GetElementByClass(
    Val           Out t_Refcur,
    分类ID_In In 影像报告元素清单.分类ID%Type
	);

  --11.功  能：获得ID对应的影像报告元素信息
  Procedure p_GetElementByID(
    Val           Out t_Refcur,
	ID_In In 影像报告元素清单.ID%Type
	);

  --12.功 能：新增影像报告元素信息
  Procedure p_AddElement(
    ID_In           In 影像报告元素清单.ID%Type,
	分类ID_In     In 影像报告元素清单.分类ID%Type,
	编码_In         In 影像报告元素清单.编码%Type,
	名称_In         In 影像报告元素清单.名称%Type,
	说明_In         In 影像报告元素清单.说明%Type,
	前缀_In         In 影像报告元素清单.前缀%Type,
	后缀_In         In 影像报告元素清单.后缀%Type,
	数据类型_In   In 影像报告元素清单.数据类型%Type,
	数值形态_In   In 影像报告元素清单.数值形态%Type,
	最小长度_In   In 影像报告元素清单.最小长度%Type,
	最大长度_In   In 影像报告元素清单.最大长度%Type,
	最小小数位_In In 影像报告元素清单.最小小数位%Type,
	最大小数位_In In 影像报告元素清单.最大小数位%Type,
	计量单位_In   In 影像报告元素清单.计量单位%Type,
	扩展描述_In   In Varchar2,
	值域ID_In      In 影像报告元素清单.值域ID%Type,
	值域种类_In   In 影像报告元素清单.值域种类%Type
	);

  --13.功 能：修改影像报告元素信息
  Procedure p_EditElement(
    ID_In         In 影像报告元素清单.ID%Type,
	分类ID_In     In 影像报告元素清单.分类ID%Type,
	编码_In       In 影像报告元素清单.编码%Type,
	名称_In       In 影像报告元素清单.名称%Type,
	前缀_In       In 影像报告元素清单.前缀%Type,
    后缀_In       In 影像报告元素清单.后缀%Type,
    说明_In       In 影像报告元素清单.说明%Type,
    数据类型_In   In 影像报告元素清单.数据类型%Type,
    数值形态_In   In 影像报告元素清单.数值形态%Type,
    最小长度_In   In 影像报告元素清单.最小长度%Type,
    最大长度_In   In 影像报告元素清单.最大长度%Type,
    最小小数位_In In 影像报告元素清单.最小小数位%Type,
    最大小数位_In In 影像报告元素清单.最大小数位%Type,
    计量单位_In   In 影像报告元素清单.计量单位%Type,
    扩展描述_In   In Varchar2,
    值域ID_In     In 影像报告元素清单.值域ID%Type,
    值域种类_In   In 影像报告元素清单.值域种类%Type
	);

  --14.功 能：删除影像报告元素信息
  Procedure p_DelElement(
    ID_In 影像报告元素清单.ID%Type
	);

  --15.功 能：通过ID获取影像报告分类信息
  Procedure p_GetElementClassByID(
    Val           Out t_Refcur,
	ID_In In 影像报告元素分类.ID%Type
	);
  --16.功  能：获取元素的下一个编码
  Procedure p_Get_ElementNextCode(
    Val Out t_Refcur
	);
  --17.功  能：获取元素分类的下一个编码
  Procedure p_Get_ElementClassNextCode(
    Val Out t_Refcur
	);
  --18.功  能：获取对应的值域类型所在的元素类别
  Procedure p_Get_ElementClassByKind(
    Val           Out t_Refcur,
	值域种类_In In 影像报告值域清单.值域种类%Type
	);
  --19.功  能：获取值域类型对应的值域信息
  Procedure p_Get_RangeByKind(
    Val           Out t_Refcur,
	值域种类_In In 影像报告值域清单.值域种类%Type
	);
  --20.功  能：获取对应的值域类型和数据类型所在的元素类别
  Procedure p_Get_ElementClassByKindType(
    Val           Out t_Refcur,
	值域种类_In In 影像报告值域清单.值域种类%Type,
	数据类型_In In 影像报告值域清单.数据类型%Type
	);
  --21.功  能：获取值域类型和数据类型对应的值域信息
  Procedure p_Get_RangeByKindAndType(
    Val           Out t_Refcur,
	值域种类_In In 影像报告值域清单.值域种类%Type,
	数据类型_In In 影像报告值域清单.数据类型%Type
	);
  --22.功 能：获取最后修改影像报告元素分类信息
  Procedure p_GetElementClassLastID(
    Val Out t_Refcur
	);
  --23.功 能：获取编辑人对应的最后修改影像报告元素信息ID
  Procedure p_GetElementLastID(
    Val Out t_Refcur
	);
  --24.功 能：获取最后修改影像报告值域信息ID
  Procedure p_GetRangeLastID(
    Val Out t_Refcur
	);
  --25.功 能：添加计量单位信息
  Procedure p_AddMasure_Unit(
    编码_In 影像报告计量单位.编码%Type,
    名称_In 影像报告计量单位.名称%Type,
    说明_In 影像报告计量单位.说明%Type,
    前缀_In 影像报告计量单位.前缀%Type
	);
  --26.功  能：修改计量单位
  Procedure p_EditMasure_Unit(
    原编码_In 影像报告计量单位.编码%Type,
    编码_In   影像报告计量单位.编码%Type,
    名称_In   影像报告计量单位.名称%Type,
    说明_In   影像报告计量单位.说明%Type,
    前缀_In   影像报告计量单位.前缀%Type
	);
  --27.功  能：删除计量单位
  Procedure p_DelMasure_Unit(
    编码_In 影像报告计量单位.编码%Type
	);
  --28.功  能：判断计量单位的编码是否已存在
  Procedure p_If_Exist_Masure_Unit(
    Val           Out t_Refcur,
	编码_In 影像报告计量单位.编码%Type
	);
  --29.功  能： 判断元素编码是否已存在
  Procedure p_If_Exist_ElementCode(
    Val           Out t_Refcur,
	ID_In   影像报告元素清单.ID%Type,
	编码_In 影像报告元素清单.编码%Type
	);
  --30.功  能： 判断元素名称是否已存在
  Procedure p_If_Exist_ElementName(
    Val           Out t_Refcur,
	ID_In   影像报告元素清单.ID%Type,
	名称_In 影像报告元素清单.名称%Type
	);
  --31.功  能： 判断值域编码是否已存在
  Procedure p_If_Exist_RangeCode(
    Val           Out t_Refcur,
	ID_In   影像报告值域清单.ID%Type,
	编码_In 影像报告值域清单.编码%Type
	);
  --32.功  能： 判断值域标题是否已存在
  Procedure p_If_Exist_RangeName(
    Val           Out t_Refcur,
	ID_In   影像报告值域清单.ID%Type,
	名称_In 影像报告值域清单.名称%Type
	);
  --33.获得元素列表
  Procedure p_GetElementList(
    Val Out t_Refcur
	);
  --34.获得值域列表
  Procedure p_GetRangeList(
    Val Out t_Refcur
	);
  --35.判断元素分类的标题和编码是否存在
  Procedure p_If_Exist_ElementClass(
    Val           Out t_Refcur,
	ID_In   影像报告元素分类.ID%Type,
	编码_In 影像报告元素分类.编码%Type,
	名称_In 影像报告元素分类.名称%Type
	) ;
    --36.判断该元素分类下面是否有值域或者元素
  Procedure p_Is_CanDel_ElementClass(
    Val           Out t_Refcur,
	ID_In 影像报告元素分类.ID%Type
	);
End b_PACS_RptElement;
/

	--影像报告元素值域(---实现部分---)***************************************************
CREATE OR REPLACE Package Body b_PACS_RptElement Is

  --1.功   能：获取全部的影像报告元素分类
  Procedure p_GetElementClassList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(上级ID) 上级ID,
             编码,
             名称,
             '[' || 编码 || ']' || 名称 As 标题,
             说明,
             最后编辑时间
        From 影像报告元素分类;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetElementClassList;

  --2.功  能：新增影像报告元素分类信息
  Procedure p_AddElementClass(
    ID_In   In 影像报告元素分类.ID%Type,
	编码_In In 影像报告元素分类.编码%Type,
	名称_In In 影像报告元素分类.名称%Type,
	说明_In In 影像报告元素分类.说明%Type,
	上级ID_In In 影像报告元素分类.上级ID%Type
	) As
  Begin
    Insert Into 影像报告元素分类
      (ID, 编码, 名称, 说明, 最后编辑时间,上级ID)
    Values
      (ID_In, 编码_In, 名称_In, 说明_In, Sysdate,上级ID_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddElementClass;

  --3.功  能：修改影像报告元素分类信息
  Procedure p_EditElementClass(
    ID_In   In 影像报告元素分类.ID%Type,
	编码_In In 影像报告元素分类.编码%Type,
	名称_In In 影像报告元素分类.名称%Type,
	说明_In In 影像报告元素分类.说明%Type,
	上级ID_In In 影像报告元素分类.上级ID%Type
	) As
  Begin
    Update 影像报告元素分类
       Set 编码         = 编码_In,
           名称         = 名称_In,
           说明         = 说明_In,
           最后编辑时间 = Sysdate,
           上级ID=上级ID_In
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditElementClass;

  --4.功  能：删除影像报告元素分类信息
  Procedure p_DelelEmentClass(
    ID_In In 影像报告元素分类.ID%Type
	) As
  Begin
    Delete From 影像报告元素分类 Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelelEmentClass;

  --5.功  能：获得分类对应的影像报告值域信息列表
  Procedure p_GetRangeByClass(
    Val           Out t_Refcur,
	分类ID_In In 影像报告值域清单.分类ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(分类ID) 分类ID,
             编码,
             名称,
             说明,
             数据类型,
             值域种类,
             (Nvl(t.值域描述, XmlType('<NULL/>'))).GetClobVal() As 值域描述,
             最后编辑时间
        From 影像报告值域清单 t
       Where 分类ID = 分类ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetRangeByClass;

  --6.功  能：获得ID对应的影像报告值域信息
  Procedure p_GetRangeByID(
    Val           Out t_Refcur,
	ID_In In 影像报告值域清单.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(分类ID) 分类ID,
             编码,
             名称,
             '[' || 编码 || ']' || 名称 标题,
             说明,
             数据类型,
             值域种类,
             (Nvl(t.值域描述, XmlType('<NULL/>'))).GetClobVal() As 值域描述,
             最后编辑时间
        From 影像报告值域清单 t
       Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetRangeByID;

  --7.功  能：新增影像报告值域信息
  Procedure p_AddRange(
    ID_In       In 影像报告值域清单.ID%Type,
	分类ID_In   In 影像报告值域清单.分类ID%Type,
    编码_In     In 影像报告值域清单.编码%Type,
    名称_In     In 影像报告值域清单.名称%Type,
    说明_In     In 影像报告值域清单.说明%Type,
    数据类型_In In 影像报告值域清单.数据类型%Type,
    值域种类_In In 影像报告值域清单.值域种类%Type,
    值域描述_In In Varchar2
	) As
  Begin
    Insert Into 影像报告值域清单
      (ID, 分类ID, 编码, 名称, 说明, 数据类型, 值域种类, 值域描述, 最后编辑时间)
    Values
      (ID_In, 分类ID_In, 编码_In, 名称_In, 说明_In, 数据类型_In, 值域种类_In, 值域描述_In, Sysdate);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddRange;

  --8.功  能：修改影像报告值域信息
  Procedure p_EditRange(
    ID_In       In 影像报告值域清单.ID%Type,
    分类ID_In   In 影像报告值域清单.分类ID%Type,
    编码_In     In 影像报告值域清单.编码%Type,
    名称_In     In 影像报告值域清单.名称%Type,
    说明_In     In 影像报告值域清单.说明%Type,
    数据类型_In In 影像报告值域清单.数据类型%Type,
    值域种类_In In 影像报告值域清单.值域种类%Type,
    值域描述_In In Varchar2
	) As
  Begin
    Update 影像报告值域清单
       Set 分类ID       = 分类ID_In,
           编码         = 编码_In,
           名称         = 名称_In,
           说明         = 说明_In,
           数据类型     = 数据类型_In,
           值域种类     = 值域种类_In,
           值域描述     = 值域描述_In,
           最后编辑时间 = Sysdate
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditRange;

  --9.功  能：删除影像报告值域信息
  Procedure p_DelRange(
    ID_In In 影像报告值域清单.ID%Type
	) As
  Begin
    Delete From 影像报告值域清单 Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelRange;

  --10.功  能：获得分类对应的影像报告元素列表
  Procedure p_GetElementByClass(
    Val           Out t_Refcur,
	分类ID_In In 影像报告元素清单.分类ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(分类ID) 分类ID,
             编码,
             名称,
             前缀,
             后缀,
             说明,
             数据类型,
             数值形态,
             最小长度,
             最大长度,
             最小小数位,
             最大小数位,
             计量单位,
             (Nvl(t.扩展描述, XmlType('<NULL/>'))).GetClobVal() As 扩展描述,
             RawToHex(值域ID) 值域ID,
             值域种类,
             最后编辑时间
        From 影像报告元素清单 t
       Where 分类ID = 分类ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetElementByClass;

  --11.功  能：获得ID对应的影像报告元素信息
  Procedure p_GetElementByID(
    Val           Out t_Refcur,
	ID_In In 影像报告元素清单.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(分类ID) 分类ID,
             编码,
             名称,
             前缀,
             后缀,
             说明,
             数据类型,
             数值形态,
             最小长度,
             最大长度,
             最小小数位,
             最大小数位,
             计量单位,
             (Nvl(t.扩展描述, XmlType('<NULL/>'))).GetClobVal() As 扩展描述,
             RawToHex(值域ID) 值域ID,
             值域种类,
             最后编辑时间
        From 影像报告元素清单 t
       Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetElementByID;

  --12.功 能：新增影像报告元素信息
  Procedure p_AddElement(
    ID_In            In 影像报告元素清单.ID%Type,
    分类ID_In     In 影像报告元素清单.分类ID%Type,
    编码_In         In 影像报告元素清单.编码%Type,
    名称_In         In 影像报告元素清单.名称%Type,
    说明_In         In 影像报告元素清单.说明%Type,
    前缀_In         In 影像报告元素清单.前缀%Type,
    后缀_In         In 影像报告元素清单.后缀%Type,
    数据类型_In   In 影像报告元素清单.数据类型%Type,
    数值形态_In   In 影像报告元素清单.数值形态%Type,
    最小长度_In   In 影像报告元素清单.最小长度%Type,
    最大长度_In   In 影像报告元素清单.最大长度%Type,
    最小小数位_In In 影像报告元素清单.最小小数位%Type,
    最大小数位_In In 影像报告元素清单.最大小数位%Type,
    计量单位_In   In 影像报告元素清单.计量单位%Type,
    扩展描述_In   In Varchar2,
    值域ID_In      In 影像报告元素清单.值域ID%Type,
    值域种类_In   In 影像报告元素清单.值域种类%Type
	) As
  Begin
    Insert Into 影像报告元素清单
      (ID, 分类ID, 编码, 名称,  前缀, 后缀, 说明,  数据类型,  数值形态,  最小长度,  最大长度,
       最小小数位,  最大小数位,  计量单位,   扩展描述, 值域ID,  值域种类, 最后编辑时间)
    Values
      (ID_In, 分类ID_In, 编码_In, 名称_In, 前缀_In, 后缀_In, 说明_In, 数据类型_In, 数值形态_In, 最小长度_In, 最大长度_In,
       最小小数位_In, 最大小数位_In, 计量单位_In, 扩展描述_In, 值域ID_In, 值域种类_In, Sysdate);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddElement;

  --13.功 能：修改影像报告元素信息
  Procedure p_EditElement(
    ID_In           In 影像报告元素清单.ID%Type,
    分类ID_In     In 影像报告元素清单.分类ID%Type,
    编码_In        In 影像报告元素清单.编码%Type,
    名称_In        In 影像报告元素清单.名称%Type,
    前缀_In        In 影像报告元素清单.前缀%Type,
    后缀_In        In 影像报告元素清单.后缀%Type,
    说明_In        In 影像报告元素清单.说明%Type,
    数据类型_In   In 影像报告元素清单.数据类型%Type,
    数值形态_In   In 影像报告元素清单.数值形态%Type,
    最小长度_In   In 影像报告元素清单.最小长度%Type,
    最大长度_In   In 影像报告元素清单.最大长度%Type,
    最小小数位_In In 影像报告元素清单.最小小数位%Type,
    最大小数位_In In 影像报告元素清单.最大小数位%Type,
    计量单位_In    In 影像报告元素清单.计量单位%Type,
    扩展描述_In    In Varchar2,
    值域ID_In      In 影像报告元素清单.值域ID%Type,
    值域种类_In   In 影像报告元素清单.值域种类%Type
	) As
  Begin

    Update 影像报告元素清单
       Set 分类ID       = 分类ID_In,
           编码         = 编码_In,
           名称         = 名称_In,
           前缀         = 前缀_In,
           后缀         = 后缀_In,
           说明         = 说明_In,
           数据类型     = 数据类型_In,
           数值形态     = 数值形态_In,
           最小长度     = 最小长度_In,
           最大长度     = 最大长度_In,
           最小小数位   = 最小小数位_In,
           最大小数位   = 最大小数位_In,
           计量单位     = 计量单位_In,
           扩展描述     = 扩展描述_In,
           值域ID       = 值域ID_In,
           值域种类     = 值域种类_In,
           最后编辑时间 = Sysdate
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditElement;

  --14.功 能：删除影像报告元素信息
  Procedure p_DelElement(
    ID_In 影像报告元素清单.ID%Type
	) As
  Begin
    Delete From 影像报告元素清单 Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelElement;

  --15.功 能：通过ID获取影像报告分类信息
  Procedure p_GetElementClassByID(
    Val           Out t_Refcur,
	ID_In In 影像报告元素分类.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID, 编码, 名称, 说明, 最后编辑时间,RawToHex(上级id) 上级ID
        From 影像报告元素分类 t
       Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetElementClassByID;

  --16.功  能：获取元素的下一个编码
  Procedure p_Get_ElementNextCode(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select b_pacs_rptpublic.f_Get_Nextcode('影像报告元素清单') As 编码
        From dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_ElementNextCode;

  --17.功  能：获取元素分类的下一个编码
  Procedure p_Get_ElementClassNextCode(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select b_pacs_rptpublic.f_Get_Nextcode('影像报告元素分类') As 编码
        From dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_ElementClassNextCode;

  --18.功  能：获取对应的值域类型所在的元素类别
  Procedure p_Get_ElementClassByKind(
    Val           Out t_Refcur,
	值域种类_In In 影像报告值域清单.值域种类%Type
	) As
  Begin
    Open Val For
      Select Distinct 分类名称, 分类ID
        From (Select RawToHex(分类ID) 分类ID,
                     (Select a.名称
                        From 影像报告元素分类 A
                       Where a.Id = t.分类id) As 分类名称
                From 影像报告值域清单 T
               Where t.值域种类 = 值域种类_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_ElementClassByKind;

  --19.功  能：获取值域类型对应的值域信息
  Procedure p_Get_RangeByKind(
    Val           Out t_Refcur,
	值域种类_In In 影像报告值域清单.值域种类%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             编码,
             RawToHex(分类ID) 分类ID,
             名称,
             数据类型,
             (Select a.名称 From 影像报告元素分类 A Where a.Id = t.分类id) As 分类名称,
             值域种类
        From 影像报告值域清单 T
       Where t.值域种类 = 值域种类_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_RangeByKind;

  --20.功  能：获取对应的值域类型和数据类型所在的元素类别
  Procedure p_Get_ElementClassByKindType(
    Val           Out t_Refcur,
	值域种类_In In 影像报告值域清单.值域种类%Type,
	数据类型_In In 影像报告值域清单.数据类型%Type
	) As
  Begin
    Open Val For
      Select Distinct 分类名称, 分类ID
        From (Select RawToHex(分类id) 分类ID,
                     (Select a.名称
                        From 影像报告元素分类 A
                       Where a.Id = t.分类id) As 分类名称
                From 影像报告值域清单 T
               Where t.值域种类 = 值域种类_In
                 and t.数据类型 = 数据类型_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_ElementClassByKindType;

  --21.功  能：获取值域类型和数据类型对应的值域信息
  Procedure p_Get_RangeByKindAndType(
    Val           Out t_Refcur,
	值域种类_In In 影像报告值域清单.值域种类%Type,
	数据类型_In In 影像报告值域清单.数据类型%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             编码,
             RawToHex(分类ID) 分类ID,
             名称,
             数据类型,
             (Select a.名称 From 影像报告元素分类 A Where a.Id = t.分类id) As 分类名称,
             值域种类
        From 影像报告值域清单 T
       Where t.值域种类 = 值域种类_In
         and t.数据类型 = 数据类型_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_RangeByKindAndType;

  --22.功 能：获取编辑人对应的最后修改影像报告原型分类信息ID
  Procedure p_GetElementClassLastID(
    Val Out t_Refcur
	) AS
  Begin
    Open Val For
      Select RawToHex(ID) ID, 最后编辑时间
        From 影像报告元素分类 t1
       Where Not Exists (Select 1
                From 影像报告元素分类
               Where 最后编辑时间 > t1.最后编辑时间);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetElementClassLastID;

  --23.功 能：获取编辑人对应的最后修改影像报告元素信息ID
  Procedure p_GetElementLastID(
    Val Out t_Refcur
	) AS
  Begin
    Open Val For
      Select RawToHex(ID) ID, 最后编辑时间
        From 影像报告元素清单 t1
       Where Not Exists (Select 1
                From 影像报告元素清单
               Where 最后编辑时间 > t1.最后编辑时间);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetElementLastID;

  --24.功 能：获取最后修改影像报告值域信息ID
  Procedure p_GetRangeLastID(
    Val Out t_Refcur
	) AS
  Begin
    Open Val For
      Select RawToHex(ID) ID, 最后编辑时间
        From 影像报告值域清单 t1
       Where Not Exists (Select 1
                From 影像报告值域清单
               Where 最后编辑时间 > t1.最后编辑时间);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetRangeLastID;

  --25.功 能：添加计量单位信息
  Procedure p_AddMasure_Unit(
    编码_In 影像报告计量单位.编码%Type,
    名称_In 影像报告计量单位.名称%Type,
    说明_In 影像报告计量单位.说明%Type,
    前缀_In 影像报告计量单位.前缀%Type
	) As
  Begin
    Insert Into 影像报告计量单位
      (编码, 名称, 说明, 前缀)
    Values
      (编码_In, 名称_In, 说明_In, 前缀_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddMasure_Unit;

  --26.功  能：修改计量单位
  Procedure p_EditMasure_Unit(
    原编码_In 影像报告计量单位.编码%Type,
    编码_In   影像报告计量单位.编码%Type,
    名称_In   影像报告计量单位.名称%Type,
    说明_In   影像报告计量单位.说明%Type,
    前缀_In   影像报告计量单位.前缀%Type
	) As
  Begin
    Update 影像报告计量单位
       Set 编码 = 编码_In, 名称 = 名称_In, 说明 = 说明_In, 前缀 = 前缀_In
     Where 编码 = 原编码_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditMasure_Unit;

  --27.功  能：删除计量单位
  Procedure p_DelMasure_Unit(
    编码_In 影像报告计量单位.编码%Type
	) As
  Begin
    Delete from 影像报告计量单位 Where 编码 = 编码_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelMasure_Unit;

  --28.功  能：判断计量单位的编码是否已存在
  Procedure p_If_Exist_Masure_Unit(
    Val           Out t_Refcur,
	编码_In 影像报告计量单位.编码%Type
	) As
  Begin
    Open Val For
      Select Count(t.编码) 数量
        From 影像报告计量单位 t
       Where t.编码 = 编码_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_If_Exist_Masure_Unit;

  --29.功  能： 判断元素编码是否已存在
  Procedure p_If_Exist_ElementCode(
    Val           Out t_Refcur,
	ID_In   影像报告元素清单.ID%Type,
	编码_In 影像报告元素清单.编码%Type
	) As
  Begin
    Open Val For
      Select Count(t.ID) 数量
        From 影像报告元素清单 t
       Where t.编码 = 编码_In
         and t.id <> ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_If_Exist_ElementCode;

  --30.功  能： 判断元素名称是否已存在
  Procedure p_If_Exist_ElementName(
    Val           Out t_Refcur,
	ID_In   影像报告元素清单.ID%Type,
	名称_In 影像报告元素清单.名称%Type
	) As
  Begin
    Open Val For
      Select Count(t.ID) 数量
        From 影像报告元素清单 t
       Where t.名称 = 名称_In
         and t.id <> ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_If_Exist_ElementName;

  --31.功  能： 判断值域编码是否已存在
  Procedure p_If_Exist_RangeCode(
    Val           Out t_Refcur,
	ID_In   影像报告值域清单.ID%Type,
	编码_In 影像报告值域清单.编码%Type
	) As
  Begin
    Open Val For
      Select Count(t.ID) 数量
        From 影像报告值域清单 t
       Where t.编码 = 编码_In
         and t.id <> ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_If_Exist_RangeCode;

  --32.功  能： 判断值域名称是否已存在
  Procedure p_If_Exist_RangeName(
    Val           Out t_Refcur,
	ID_In   影像报告值域清单.ID%Type,
	名称_In 影像报告值域清单.名称%Type
	) As
  Begin
    Open Val For
      Select Count(t.ID) 数量
        From 影像报告值域清单 t
       Where t.名称 = 名称_In
         and t.id <> ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_If_Exist_RangeName;

  --33.获得元素列表
  Procedure p_GetElementList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(t.ID) ID,
             RawToHex(t.分类ID) 分类ID,
             t.编码,
             t.名称,
             t.前缀,
             t.后缀,
             t.说明,
             t.数据类型,
             t.数值形态,
             t.最小长度,
             t.最大长度,
             t.最小小数位,
             t.最大小数位,
             t.计量单位,
             (Nvl(t.扩展描述, XmlType('<NULL/>'))).GetClobVal() As 扩展描述,
             RawToHex(t.值域ID) 值域ID,
             (select a.名称 from 影像报告值域清单 a Where a.id=t.值域ID)as 值域名称,
             t.值域种类,
             t.最后编辑时间
        From 影像报告元素清单 t;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetElementList;

  --34.获得值域列表
  Procedure p_GetRangeList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(分类ID) 分类ID,
             编码,
             名称,
             说明,
             数据类型,
             值域种类,
             (Nvl(t.值域描述, XmlType('<NULL/>'))).GetClobVal() As 值域描述,
             最后编辑时间
        From 影像报告值域清单 t;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetRangeList;

  --35.判断元素分类的标题和编码是否存在
  Procedure p_If_Exist_ElementClass(
    Val           Out t_Refcur,
    ID_In   影像报告元素分类.ID%Type,
	编码_In 影像报告元素分类.编码%Type,
	名称_In 影像报告元素分类.名称%Type
	) As
  Begin
    Open Val For
      Select Count(t.ID) 数量
        From 影像报告元素分类 t
       Where (t.名称 = 名称_In or t.编码 = 编码_In)
         and t.id <> ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_If_Exist_ElementClass;

  --36.判断该元素分类下面是否有值域或者元素
  Procedure p_Is_CanDel_ElementClass(
    Val           Out t_Refcur,
	ID_In 影像报告元素分类.ID%Type
	) As
    ElementCount int;
    RangeCount   int;
    ElementClassCout int;
  Begin
    Select Count(*)
      into ElementCount
      From 影像报告元素清单 a
     Where a.分类id = ID_In;
    Select Count(*)
      into RangeCount
      From 影像报告值域清单 b
     Where b.分类id = ID_In;
     Select Count(*)
      into ElementClassCout
      From 影像报告元素分类 b
     Where b.上级id = ID_In;
    ElementCount := ElementCount + RangeCount+ElementClassCout;
    Open Val For
      Select ElementCount Count From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Is_CanDel_ElementClass;

End b_PACS_RptElement;
/




--影像报告组句管理(---定义部分---)***************************************************
--影像报告组句管理(---定义部分---)***************************************************
CREATE OR REPLACE Package b_PACS_RptCombo Is
  --Create By Hwei;
  --2014/11/25
  Type t_Refcur Is Ref Cursor;

  --1.功  能：获得影像报告组句列表
  Procedure p_GetComboList(
    Val Out t_Refcur
	);
  --2.功  能：添加影像报告组句信息
  Procedure p_AddComboInfo(
    ID_In     In 影像报告组句清单.ID%Type,
    编码_In   In 影像报告组句清单.编码%Type,
    名称_In   In 影像报告组句清单.名称%Type,
    说明_In   In 影像报告组句清单.说明%Type,
    分组_In   In 影像报告组句清单.分组%Type,
    多组_In   In 影像报告组句清单.多组%Type,
    组成_In   In 影像报告组句清单.组成%Type,
    编辑人_In In 影像报告组句清单.编辑人%Type
	);
  --3.功  能;修改影像报告组句信息
  Procedure p_EditComboInfo(
    ID_In     In 影像报告组句清单.ID%Type,
    编码_In   In 影像报告组句清单.编码%Type,
    名称_In   In 影像报告组句清单.名称%Type,
    说明_In   In 影像报告组句清单.说明%Type,
    分组_In   In 影像报告组句清单.分组%Type,
    多组_In   In 影像报告组句清单.多组%Type,
    组成_In   In 影像报告组句清单.组成%Type,
    编辑人_In In 影像报告组句清单.编辑人%Type
	);
  --4.功  能：通过ID删除影像报告组句信息
  Procedure p_DelComboInfo(
    ID_In In 影像报告组句清单.ID%Type
	);
  --5.功  能：根据ID获得影像报告组句信息
  Procedure p_GetComboInfoByID(
	Val           Out t_Refcur,
	ID_In In 影像报告组句清单.ID%Type
	);
  --6.功  能：获得影像报告组句的所有分组信息
  Procedure p_GetComboAllGroup(
    Val Out t_Refcur
	);
  --7.功  能：获得ID对应的影像报告组句的短语信息
  Procedure p_GetComboContent(
	Val           Out t_Refcur,
	ID_In In 影像报告组句清单.ID%Type
	);
  --8.功  能：更新ID对应的影像报告组句的短语信息
  Procedure p_EditComboContent(
	ID_In   In 影像报告组句清单.ID%Type,
	组成_In  In 影像报告组句清单.组成%Type
	);
  --9.功 能：获取编辑人对应的最后修改影像报告组句信息
  Procedure p_GetComboInfoByEditor(
	Val           Out t_Refcur,
	编辑人_In In 影像报告组句清单.编辑人%Type
	);
  --10.功  能：新增片段到组合句
  Procedure p_Append_Fragment_Tocombo(
    Text_In In XmlType,
    Id_In   In 影像报告组句清单.ID%Type
	);

  --11.功  能：修改片段到组合句
  Procedure p_Update_Combo_Fragment(
    Text_In In XmlType,
    Id_In   In 影像报告组句清单.ID%Type,
    Pid_In  In Varchar2
	);
  --12.功  能：根据分类ID查询词句
  Procedure p_Get_Fragment_By_Typeid(
	Val           Out t_Refcur,
	Id_In In 影像报告组句清单.ID%Type
	);
  --13.功  能：获取下一个编码
  Procedure p_Get_ComboNextCode(
    Val Out t_Refcur
	);
end b_PACS_RptCombo;
/

--影像报告组句管理(---实现部分---)***************************************************
CREATE OR REPLACE Package Body b_PACS_RptCombo Is
  --Create By Hwei;
  --2014/11/25

  --1.功  能：获得影像报告组句列表
  Procedure p_GetComboList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             编码,
             名称,
             说明,
             分组,
             多组,
             (Nvl(t.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
             编辑人,
             最后编辑时间
        From 影像报告组句清单 t;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboList;

  --2.功  能：添加影像报告组句信息
  Procedure p_AddComboInfo(
    ID_In     In 影像报告组句清单.ID%Type,
    编码_In   In 影像报告组句清单.编码%Type,
    名称_In   In 影像报告组句清单.名称%Type,
    说明_In   In 影像报告组句清单.说明%Type,
    分组_In   In 影像报告组句清单.分组%Type,
    多组_In   In 影像报告组句清单.多组%Type,
    组成_In   In 影像报告组句清单.组成%Type,
    编辑人_In In 影像报告组句清单.编辑人%Type
	) As
  Begin
    Insert Into 影像报告组句清单
      (ID, 编码, 名称, 说明, 分组, 多组, 组成, 编辑人, 最后编辑时间)
    Values
      (ID_In, 编码_In, 名称_In, 说明_In, 分组_In, 多组_In, 组成_In, 编辑人_In, Sysdate);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddComboInfo;

  --3.功  能;修改影像报告组句信息
  Procedure p_EditComboInfo(
    ID_In     In 影像报告组句清单.ID%Type,
    编码_In   In 影像报告组句清单.编码%Type,
    名称_In   In 影像报告组句清单.名称%Type,
    说明_In   In 影像报告组句清单.说明%Type,
    分组_In   In 影像报告组句清单.分组%Type,
    多组_In   In 影像报告组句清单.多组%Type,
    组成_In   In 影像报告组句清单.组成%Type,
    编辑人_In In 影像报告组句清单.编辑人%Type
	) As
  Begin
    Update 影像报告组句清单
       set 编码         = 编码_In,
           名称         = 名称_In,
           说明         = 说明_In,
           分组         = 分组_In,
           多组         = 多组_In,
           组成         = 组成_In,
           编辑人       = 编辑人_In,
           最后编辑时间 = SysDate
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditComboInfo;

  --4.功  能：通过ID删除影像报告组句信息
  Procedure p_DelComboInfo(
    ID_In In 影像报告组句清单.ID%Type
	) As
  Begin
    Delete From 影像报告组句清单 Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelComboInfo;

  --5.功  能：根据ID获得影像报告组句信息
  Procedure p_GetComboInfoByID(
	Val           Out t_Refcur,
	ID_In In 影像报告组句清单.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             编码,
             名称,
             说明,
             分组,
             多组,
             (Nvl(t.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
             编辑人,
             最后编辑时间
        From 影像报告组句清单 t
       Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboInfoByID;

  --6.功  能：获得影像报告组句的所有分组信息
  Procedure p_GetComboAllGroup(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Distinct 分组 From 影像报告组句清单;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboAllGroup;

  --7.功  能：获得ID对应的影像报告组句的短语信息
  Procedure p_GetComboContent(
	Val           Out t_Refcur,
	ID_In In 影像报告组句清单.ID%Type
	) As
  Begin
    Open Val For
      Select (Nvl(t.组成, XmlType('<NULL/>'))).GetClobVal() As 组成
        From 影像报告组句清单 t
       Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboContent;

  --8.功  能：更新ID对应的影像报告组句的短语信息
  Procedure p_EditComboContent(
    ID_In   In 影像报告组句清单.ID%Type,
    组成_In In 影像报告组句清单.组成%Type
	) As
  Begin
    Update 影像报告组句清单 Set 组成 = 组成_In Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditComboContent;

  --9.功 能：获取编辑人对应的最后修改影像报告组句信息
  Procedure p_GetComboInfoByEditor(
	Val           Out t_Refcur,
	编辑人_In In 影像报告组句清单.编辑人%Type
	) AS
  Begin
    Open Val For
      Select RawToHex(ID) ID, 编辑人, 最后编辑时间
        From 影像报告组句清单 t1
       Where Not Exists (Select 1
                From 影像报告组句清单
               Where 最后编辑时间 > t1.最后编辑时间)
         And 编辑人 = 编辑人_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboInfoByEditor;

  --10.功  能：新增片段到组合句
  Procedure p_Append_Fragment_Tocombo(
    Text_In In XmlType,
	Id_In   In 影像报告组句清单.ID%Type
	) As
  Begin
    Update 影像报告组句清单 A
       Set a.组成 = Appendchildxml(a.组成, '/root', Text_In)
     Where a.ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Append_Fragment_Tocombo;

  --11.功  能：修改片段到组合句
  Procedure p_Update_Combo_Fragment(
    Text_In In XmlType,
    Id_In   In 影像报告组句清单.ID%Type,
    Pid_In  In Varchar2
	) As
  Begin
    Update 影像报告组句清单 A
       Set a.组成 = Updatexml(a.组成,
                            '/root/sentence[@sid="' || Pid_In || '"]',
                            Text_In)
     Where a.ID = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Combo_Fragment;

  --12.功  能：根据分类ID查询词句
  Procedure p_Get_Fragment_By_Typeid(
	Val           Out t_Refcur,
	Id_In In 影像报告组句清单.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             上级id,
             编码,
             名称,
             说明,
             节点类型,
             (Nvl(a.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
             学科,
             标签,
             是否私有,
             作者,
			 (Nvl(a.适应条件, XmlType('<NULL/>'))).GetClobVal() As 适应条件, 
             最后编辑时间
        From 影像报告片段清单 A
       Where a.上级id = Id_In
         And a.节点类型 <> 0;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Fragment_By_Typeid;

  --13.功  能：获取下一个编码
  Procedure p_Get_ComboNextCode(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select b_pacs_rptpublic.f_Get_Nextcode('影像报告组句清单') As 编码
        From dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_ComboNextCode;
End b_PACS_RptCombo;
/



CREATE OR REPLACE Package b_PACS_RptFragments Is
  Type t_Refcur Is Ref Cursor;


  --功能：获取所有预备提纲
  Procedure p_Get_All_Phr_Onlines(
    Val Out t_Refcur
	);
  --功能：获取所有短语分类
  Procedure p_Get_All_Fragment_Class(
    Val Out t_Refcur
	);
  --功能：获取当前用户学科所有短语包括父节点
  Procedure p_Get_All_Fragment(
    Val           Out t_Refcur,
    Subjects_In 影像报告片段清单.学科%Type
	) ;


  --功能：根据分类ID查找短语
  Procedure p_Get_Fragment_By_Typeid(
    Val           Out t_Refcur,
    Id_In 影像报告片段清单.ID%Type
	) ;


   Procedure p_Get_Label_By_Typeid(
     Val           Out t_Refcur,
	 Id_In 影像报告片段清单.ID%Type
	 ) ;

  --功能：新增短语分类
  Procedure p_Add_Fragmenttype(
    Id_In     影像报告片段清单.ID%Type,
    Pid_In    影像报告片段清单.上级ID%Type,
    Code_In   影像报告片段清单.编码%Type,
    Title_In  影像报告片段清单.名称%Type,
    Note_In   影像报告片段清单.说明%Type,
    Leaf_In   影像报告片段清单.节点类型%Type,
    Author_In 影像报告片段清单.作者%Type
    ) ;

  --功能：修改短语分类
  Procedure p_Edit_Fragmenttype(
    Id_In     影像报告片段清单.ID%Type,
    Pid_In    影像报告片段清单.上级ID%Type,
    Code_In   影像报告片段清单.编码%Type,
    Title_In  影像报告片段清单.名称%Type,
    Note_In   影像报告片段清单.说明%Type,
    Leaf_In   影像报告片段清单.节点类型%Type,
    Author_In 影像报告片段清单.作者%Type
    ) ;

  --功能：删除短语分类
   Procedure p_Del_Fragmenttype(
     Id_In 影像报告片段清单.ID%Type
	 );

    --功能：添加短语
  Procedure p_Add_Fragment(
     Id_In      影像报告片段清单.ID%Type,
    Pid_In      影像报告片段清单.上级ID%Type,
    Code_In     影像报告片段清单.编码%Type,
    Title_In    影像报告片段清单.名称%Type,
    Note_In     影像报告片段清单.说明%Type,
    Leaf_In     影像报告片段清单.节点类型%Type,
    Content_In  影像报告片段清单.组成%Type,
    Subjects_In 影像报告片段清单.学科%Type,
    Label_In    影像报告片段清单.标签%Type,
    Private_In  影像报告片段清单.是否私有%Type,
    Author_In   影像报告片段清单.作者%Type
    ) ;

   --功能：修改短语
  Procedure p_Edit_Fragment(
    Id_In       影像报告片段清单.ID%Type,
    Pid_In      影像报告片段清单.上级ID%Type,
    Code_In     影像报告片段清单.编码%Type,
    Title_In    影像报告片段清单.名称%Type,
    Note_In     影像报告片段清单.说明%Type,
    Leaf_In     影像报告片段清单.节点类型%Type,
    Content_In  影像报告片段清单.组成%Type,
    Subjects_In 影像报告片段清单.学科%Type,
    Label_In    影像报告片段清单.标签%Type,
    Private_In  影像报告片段清单.是否私有%Type,
    Author_In   影像报告片段清单.作者%Type
    );
   --功能：删除短语
  Procedure p_Del_Fragment(
    Id_In 影像报告片段清单.ID%Type
	);

  procedure p_Get_All_Fragment_List(
    Val Out t_Refcur
	);

  --功能：导入短语
  Procedure p_Import_Fragment(
    Id_In       影像报告片段清单.ID%Type,
    Pid_In      影像报告片段清单.上级ID%Type,
    Code_In     影像报告片段清单.编码%Type,
    Title_In    影像报告片段清单.名称%Type,
    Note_In     影像报告片段清单.说明%Type,
    Leaf_In     影像报告片段清单.节点类型%Type,
    Content_In  影像报告片段清单.组成%Type,
    Subjects_In 影像报告片段清单.学科%Type,
    Label_In    影像报告片段清单.标签%Type,
    Private_In  影像报告片段清单.是否私有%Type,
    Author_In   影像报告片段清单.作者%Type
    ) ;

procedure p_Get_Data_Last_Edit_Time(
  Val           Out t_Refcur,
  Table_Name_In varchar2
  );

   --功能：判断片段分类能否删除
  Procedure p_IsCanDel_FragmentType(
    Val           Out t_Refcur,
	Id_In 影像报告片段清单.Id%Type
	);

  --功能：根据片段ID，设置当前片段的适应条件
  Procedure p_Edit_FragmentConditionById
  (
    ID_In      In 影像报告片段清单.ID%Type,
    适应条件_In In 影像报告片段清单.适应条件%Type
  );
  
  --功能：根据片段的父ID，设置整个目录或子目录片段的适应条件
  Procedure p_Edit_FragmentConditionByPid
  (
    上级ID_In      In 影像报告片段清单.ID%Type,
    适应条件_In    In 影像报告片段清单.适应条件%Type
  );

  --功能：获取当前检查的片段适应条件
  Procedure p_Get_FraConditionByOrderId
  (
    Val           Out t_Refcur,
	医嘱ID_In    影像检查记录.医嘱ID%Type
  );

  --功能：获取影像检查类别
  Procedure p_Get_CheckLueKind
  (
    Val           Out t_Refcur
  );
  
  --功能：根据类别获取诊疗检查部位
  Procedure p_Get_CheckPartList
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  );
  
  --功能：根据类别获取影像检查项目
  Procedure p_Get_CheckRadListByKind
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  );
  
  --功能：根据诊疗编码获取影像检查项目
  Procedure p_Get_CheckRadListByCode
  (
    Val           Out t_Refcur,
    Code_In       Varchar2
  );

  --判断是否有相同的代码
  Procedure p_Get_HasSameCode
  (
  Val      Out t_Refcur,
  ID_In      In 影像报告片段清单.ID%Type,
  Code_In  影像报告片段清单.编码%Type
  );

  --判断是否有相同的名称
  Procedure p_Get_HasSameName
  (
  Val      Out t_Refcur,
  ID_In    In 影像报告片段清单.ID%Type,
  PID_In    In 影像报告片段清单.上级ID%Type,
  Name_In  In 影像报告片段清单.名称%Type,
  Author_In In  影像报告片段清单.作者%Type
  );

  End  b_PACS_RptFragments;
/
CREATE OR REPLACE Package Body b_PACS_RptFragments Is

  ------------------------------------------------------------------------
  --片段模块
  ------------------------------------------------------------------------

  --功能：获取所有预备提纲
  Procedure p_Get_All_Phr_Onlines(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(ID) ID, 编码, 名称 From 影像报告预备提纲 Order By 编码;
  End p_Get_All_Phr_Onlines;

  --功能：获取所有短语分类
  Procedure p_Get_All_Fragment_Class(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(a.Id) As ID, Rawtohex(a.上级id) As 上级id, a.编码, a.名称, a.说明, a.节点类型
      From 影像报告片段清单 A
      Where a.节点类型 = 0
      Start With 上级id Is Null
      Connect By Prior ID = 上级id
      Order By 上级ID Desc, 编码;
  End p_Get_All_Fragment_Class;

  --功能：获取当前用户学科所有短语包括父节点
  Procedure p_Get_All_Fragment(
    Val           Out t_Refcur,
    Subjects_In 影像报告片段清单.学科%Type
	) As
  Begin
    If Subjects_In <> '' Then
      Open Val For
        Select Rawtohex(a.Id) As ID, Rawtohex(a.上级id) As 上级id, a.编码, a.名称, a.说明, a.节点类型, (Nvl(a.组成, XmlType('<NULL/>'))).GetClobVal() As 组成, 
			a.学科, a.标签, a.是否私有, a.作者, (Nvl(a.适应条件, XmlType('<NULL/>'))).GetClobVal() As 适应条件,a.最后编辑时间, a.节点类型 As Image
        From 影像报告片段清单 A
        Where (a.学科 In (Select /*+rule*/
                         Column_Value As Lable
                        From Table(b_Pacs_Common.f_Str2list(Subjects_In, ','))
                        Intersect
                        Select /*+rule*/
                         Column_Value As Lable
                        From Table(b_Pacs_Common.f_Str2list(a.学科, ','))) And a.节点类型 <> 0) Or a.节点类型 = 0 Or a.学科 Is Null
        Order By 编码, 上级id;
    Else
      Open Val For
        Select Rawtohex(a.Id) As ID, Rawtohex(a.上级id) As 上级id, a.编码, a.名称, a.说明, a.节点类型, (Nvl(a.组成, XmlType('<NULL/>'))).GetClobVal() As 组成, 
			a.学科, a.标签, a.是否私有, a.作者, (Nvl(a.适应条件, XmlType('<NULL/>'))).GetClobVal() As 适应条件,a.最后编辑时间, a.节点类型 As Image
        From 影像报告片段清单 A
        Order By 上级id, 节点类型, 编码, 名称;
    End If;
  End p_Get_All_Fragment;

  --功能：根据分类ID查找短语
  Procedure p_Get_Fragment_By_Typeid(
    Val           Out t_Refcur,
    Id_In 影像报告片段清单.Id%Type
  ) As
  Begin
    Open Val For
      Select Rawtohex(a.Id) As ID, a.上级ID,a.编码, a.名称, a.说明, a.节点类型, (Nvl(a.组成, XmlType('<NULL/>'))).GetClobVal() As 组成, 
				a.学科, a.标签, a.是否私有, a.作者, (Nvl(a.适应条件, XmlType('<NULL/>'))).GetClobVal() As 适应条件, a.最后编辑时间,a.节点类型 As Image
      From 影像报告片段清单 A
      Where a.上级id = Hextoraw(Id_In) And a.节点类型 <> 0;
  End p_Get_Fragment_By_Typeid;

  --功能：查找某分类下所有短语标签
  Procedure p_Get_Label_By_Typeid(
    Val           Out t_Refcur,
    Id_In 影像报告片段清单.Id%Type
    ) As
  Begin
    Open Val For
      Select Distinct 标签 From 影像报告片段清单 Where 上级id = Hextoraw(Id_In) And 标签 Is Not Null;
  End p_Get_Label_By_Typeid;

  --功能：新增短语分类
  Procedure p_Add_Fragmenttype(
    Id_In     影像报告片段清单.Id%Type,
    Pid_In    影像报告片段清单.上级id%Type,
    Code_In   影像报告片段清单.编码%Type,
    Title_In  影像报告片段清单.名称%Type,
    Note_In   影像报告片段清单.说明%Type,
    Leaf_In   影像报告片段清单.节点类型%Type,
    Author_In 影像报告片段清单.作者%Type
    ) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From 影像报告片段清单
    Where 编码 = Code_In Or 名称 = Title_In And 节点类型 = 0 And 上级id = Hextoraw(Pid_In);

    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]分类名称或编码已经存在！[ZLSOFT]';
      Raise Err_Item;
    Else
      Insert Into 影像报告片段清单
        (ID, 上级id, 编码, 名称, 说明, 节点类型, 作者, 最后编辑时间)
      Values
        (Hextoraw(Id_In), Hextoraw(Pid_In), Code_In, Title_In, Note_In, Leaf_In, Author_In, Sysdate);
    End If;

  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Add_Fragmenttype;

  --功能：修改短语分类
  Procedure p_Edit_Fragmenttype(
    Id_In     影像报告片段清单.Id%Type,
    Pid_In    影像报告片段清单.上级id%Type,
    Code_In   影像报告片段清单.编码%Type,
    Title_In  影像报告片段清单.名称%Type,
    Note_In   影像报告片段清单.说明%Type,
    Leaf_In   影像报告片段清单.节点类型%Type,
    Author_In 影像报告片段清单.作者%Type
    ) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From 影像报告片段清单
    Where (编码 = Code_In Or 名称 = Title_In) And 节点类型 = 0 And 上级id = Hextoraw(Pid_In) And ID <> Hextoraw(Id_In);

	If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]分类名称或编码已经存在！[ZLSOFT]';
      Raise Err_Item;
    Else
      Update 影像报告片段清单
      Set 上级id = Hextoraw(Pid_In), 编码 = Code_In, 名称 = Title_In, 说明 = Note_In, 节点类型 = Leaf_In, 作者 = Author_In,
          最后编辑时间 = Sysdate
      Where ID = Hextoraw(Id_In);
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_Fragmenttype;

  --功能：删除短语分类
  Procedure p_Del_Fragmenttype(
    Id_In 影像报告片段清单.Id%Type
	) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From 影像报告片段清单
    Where 节点类型 <> 0 And
          ID In (Select ID From 影像报告片段清单 Connect By Prior ID = 上级id Start With ID = Hextoraw(Id_In));

	If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]该分类下存在短语，暂不能删除！[ZLSOFT]';
      Raise Err_Item;
    Else
      Delete 影像报告片段清单 Where ID = Hextoraw(Id_In);
    End If;

  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Del_Fragmenttype;

  --功能：添加短语
  Procedure p_Add_Fragment(
    Id_In       影像报告片段清单.Id%Type,
    Pid_In      影像报告片段清单.上级id%Type,
    Code_In     影像报告片段清单.编码%Type,
    Title_In    影像报告片段清单.名称%Type,
    Note_In     影像报告片段清单.说明%Type,
    Leaf_In     影像报告片段清单.节点类型%Type,
    Content_In  影像报告片段清单.组成%Type,
    Subjects_In 影像报告片段清单.学科%Type,
    Label_In    影像报告片段清单.标签%Type,
    Private_In  影像报告片段清单.是否私有%Type,
    Author_In   影像报告片段清单.作者%Type
    ) As
  Begin

      Insert Into 影像报告片段清单
        (ID, 上级id, 编码, 名称, 说明, 节点类型, 组成, 学科, 标签, 是否私有, 作者, 最后编辑时间)
      Values
        (Hextoraw(Id_In), Hextoraw(Pid_In), Code_In, Title_In, Note_In, Leaf_In, Content_In, Subjects_In, Label_In,
         Private_In, Author_In, Sysdate);

  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Add_Fragment;

  --功能：修改短语
  Procedure p_Edit_Fragment(
    Id_In       影像报告片段清单.Id%Type,
    Pid_In      影像报告片段清单.上级id%Type,
    Code_In     影像报告片段清单.编码%Type,
    Title_In    影像报告片段清单.名称%Type,
    Note_In     影像报告片段清单.说明%Type,
    Leaf_In     影像报告片段清单.节点类型%Type,
    Content_In  影像报告片段清单.组成%Type,
    Subjects_In 影像报告片段清单.学科%Type,
    Label_In    影像报告片段清单.标签%Type,
    Private_In  影像报告片段清单.是否私有%Type,
    Author_In   影像报告片段清单.作者%Type
    ) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From 影像报告片段清单
    Where (编码 = Code_In Or 名称 = Title_In) And 节点类型 <> 0 And 上级id = Hextoraw(Pid_In) And ID <> Hextoraw(Id_In);

	If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]短语的名称或编码已经存在！[ZLSOFT]';
      Raise Err_Item;
    Else
      Update 影像报告片段清单
      Set 上级id = Hextoraw(Pid_In), 编码 = Code_In, 名称 = Title_In, 说明 = Note_In, 节点类型 = Leaf_In, 组成 = Content_In,
          学科 = Subjects_In, 标签 = Label_In, 是否私有 = Private_In, 作者 = Author_In, 最后编辑时间 = Sysdate
      Where ID = Hextoraw(Id_In);
    End If;

  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_Fragment;

  --
  Procedure p_Get_All_Fragment_List(Val Out t_Refcur) As
  Begin
    Open Val For
      Select Rawtohex(t.Id) As ID, Rawtohex(t.上级id) As 上级id, t.编码, t.名称, t.说明, t.节点类型, (Nvl(t.组成, XmlType('<NULL/>'))).GetClobVal() As 组成, t.学科, t.标签, t.是否私有, t.作者,
             t.最后编辑时间
      From 影像报告片段清单 T;
  End p_Get_All_Fragment_List;

  --功能：删除短语
  Procedure p_Del_Fragment(
    Id_In 影像报告片段清单.Id%Type
	) As
  Begin
    Delete 影像报告片段清单 Where ID = Hextoraw(Id_In);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Del_Fragment;

  --功能：导入短语
  Procedure p_Import_Fragment(
    Id_In       影像报告片段清单.Id%Type,
    Pid_In      影像报告片段清单.上级id%Type,
    Code_In     影像报告片段清单.编码%Type,
    Title_In    影像报告片段清单.名称%Type,
    Note_In     影像报告片段清单.说明%Type,
    Leaf_In     影像报告片段清单.节点类型%Type,
    Content_In  影像报告片段清单.组成%Type,
    Subjects_In 影像报告片段清单.学科%Type,
    Label_In    影像报告片段清单.标签%Type,
    Private_In  影像报告片段清单.是否私有%Type,
    Author_In   影像报告片段清单.作者%Type
    ) As
    v_Num Number(2);
  Begin
    Select Count(ID)
    Into v_Num
    From 影像报告片段清单
    Where ((编码 = Code_In Or 名称 = Title_In) And 上级id = Hextoraw(Pid_In)) Or
          (上级id Is Null And (编码 = Code_In Or 名称 = Title_In));

    If v_Num > 0 Then
      Update 影像报告片段清单
      Set 组成 = Content_In, 最后编辑时间 = Sysdate, 是否私有 = 0
      Where ((编码 = Code_In Or 名称 = Title_In) And 上级id = Hextoraw(Pid_In)) Or
            (上级id Is Null And (编码 = Code_In Or 名称 = Title_In));
    Else
      Insert Into 影像报告片段清单
        (ID, 上级id, 编码, 名称, 说明, 节点类型, 组成, 学科, 标签, 是否私有, 作者, 最后编辑时间)
      Values
        (Hextoraw(Id_In), Hextoraw(Pid_In), Code_In, Title_In, Note_In, Leaf_In, Content_In, Subjects_In, Label_In,
         Private_In, Author_In, Sysdate);
    End If;

  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Import_Fragment;

  --
  Procedure p_Get_Data_Last_Edit_Time(
    Val           Out t_Refcur,
    Table_Name_In Varchar2
    ) As
    v_Sql Varchar2(4000);
  Begin
    v_Sql := 'select max(最后编辑时间) maxvalue from ' || Table_Name_In;
    Open Val For v_Sql;
  End p_Get_Data_Last_Edit_Time;
  
   --功能：判断片段分类能否删除
  Procedure p_IsCanDel_FragmentType(
    Val           Out t_Refcur,
	Id_In 影像报告片段清单.Id%Type
	) As
  Begin
    Open Val For
      Select Count(t.id) Count
        From 影像报告片段清单 t
       Where 上级id = Hextoraw(Id_In);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_IsCanDel_FragmentType;
  
  --功能：根据片段ID，设置当前片段的适应条件
  Procedure p_Edit_FragmentConditionById
  (
    ID_In      In 影像报告片段清单.ID%Type,
    适应条件_In In 影像报告片段清单.适应条件%Type
  )As
  Begin
    Update 影像报告片段清单 Set 适应条件 = 适应条件_In Where ID = Hextoraw(ID_In) And 节点类型 != 0;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_FragmentConditionById;
  
  --功能：根据片段的父ID，设置整个目录或子目录片段的适应条件
  Procedure p_Edit_FragmentConditionByPid
  (
    上级ID_In      In 影像报告片段清单.ID%Type,
    适应条件_In In 影像报告片段清单.适应条件%Type
  )As
  Begin
    Update 影像报告片段清单 Set 适应条件 = 适应条件_In Where 上级ID = Hextoraw(上级ID_In) And 节点类型 != 0;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_FragmentConditionByPid;

  --功能：获取当前检查的片段适应条件
  Procedure p_Get_FraConditionByOrderId(
    Val           Out t_Refcur,
	  医嘱ID_In    影像检查记录.医嘱ID%Type
	) As
  Begin
    Open Val For
	  Select a.id, a.性别,c.影像类别, d.编码||' - '||d.名称 检查类别, c.影像类别||' - '||e.编码||' - '||e.名称 检查项目, A.医嘱内容
      From 病人医嘱记录 a, 病人医嘱发送 b, 影像检查记录 c, 影像检查类别 d, 诊疗项目目录 e
      Where a.id = b.医嘱id and b.医嘱id=c.医嘱id and c.影像类别 = d.编码 and a.诊疗项目id = e.id and a.id = 医嘱ID_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_FraConditionByOrderId;

  --功能：获取影像检查类别
  Procedure p_Get_CheckLueKind
  (
    Val           Out t_Refcur
  ) As
  Begin
    Open Val For
      Select 编码||' - '||名称 检查类别 From 影像检查类别;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckLueKind;
  
  --功能：根据类别获取诊疗检查部位
  Procedure p_Get_CheckPartList
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  ) As
  Begin
    Open Val For
      Select Distinct 类型||分组 IID, '' 上级ID, 类型||' - '||分组 诊疗部位 From 诊疗检查部位 a,
      Table(Cast(f_Str2list(''||Kind_In||'') As zlTools.t_Strlist)) b Where a.类型 = b.Column_Value
      Union Select 类型||分组||名称 IID, 类型||分组 上级ID, 类型||' - '||名称 诊疗部位 From 诊疗检查部位 c,
      Table(Cast(f_Str2list(''||Kind_In||'') As zlTools.t_Strlist)) d Where c.类型 = d.Column_Value;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckPartList;
  
  --功能：根据类别获取影像检查项目
  Procedure p_Get_CheckRadListByKind
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  ) As
  Begin
    Open Val For
      Select I.编码, r.影像类别||' - '||I.编码||' - '||I.名称 检查项目
      From 诊疗项目目录 I, 影像检查项目 R, Table(Cast(f_Str2list(''||Kind_In||'') As zlTools.t_Strlist)) M
      Where I.ID = R.诊疗项目id And R.影像类别=M.Column_Value;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckRadListByKind;
  
  --功能：根据诊疗编码获取影像检查项目
  Procedure p_Get_CheckRadListByCode
  (
    Val           Out t_Refcur,
    Code_In       Varchar2
  ) As
  Begin
    Open Val For
      Select I.编码, r.影像类别||' - '||I.编码||' - '||I.名称 检查项目
      From 诊疗项目目录 I, 影像检查项目 R, Table(Cast(f_Str2list(''||Code_In||'') As zlTools.t_Strlist)) M
      Where I.ID = R.诊疗项目id And I.编码=M.Column_Value;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckRadListByCode;

  --判断是否有相同的代码
  Procedure p_Get_HasSameCode
  (
  Val      Out t_Refcur,
  ID_In      In 影像报告片段清单.ID%Type,
  Code_In  影像报告片段清单.编码%Type
  ) As
  Begin
  Open Val For
    Select Count(1) From 影像报告片段清单 Where ID<>ID_In And 编码=Code_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_HasSameCode;

  --判断是否有相同的名称
  Procedure p_Get_HasSameName
  (
  Val      Out t_Refcur,
  ID_In    In 影像报告片段清单.ID%Type,
  PID_In    In 影像报告片段清单.上级ID%Type,
  Name_In  In 影像报告片段清单.名称%Type,
  Author_In In  影像报告片段清单.作者%Type
  ) As
  Begin
  Open Val For
    Select Count(1) From 影像报告片段清单 Where 上级ID=PID_In And 作者=Author_In And ID<>ID_In And 名称=Name_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_HasSameName;

End  b_PACS_RptFragments;
/





--影像报告原型管理(---定义部分---)***************************************************
CREATE OR REPLACE Package b_PACS_RptAntetype Is
  --Create By Hwei;
  --2014/11/25
  Type t_Refcur Is Ref Cursor;

  --1.获取文件原型类别
  Procedure p_Get_Antetypelistkind(
    Val Out t_Refcur
	);
  --2.根据文档类型获取文档信息
  Procedure p_Get_Antetypelis_By_Kind(
	Val           Out t_Refcur,
	种类_In      影像报告原型清单.种类%Type,
	Stop_Flag    Number,
	Condition_In Varchar2
	);
  --3.添加一个文档原型
  Procedure p_Add_Antetypelist(
    Id_In           影像报告原型清单.ID%Type,
	种类_In         影像报告原型清单.种类%Type,
	编码_In         影像报告原型清单.编码%Type,
	名称_In         影像报告原型清单.名称%Type,
    设备号_In		影像设备目录.设备号%Type,
	说明_In         影像报告原型清单.说明%Type,
	可否重置页面_In 影像报告原型清单.可否重置页面%Type,
	可否重置格式_In 影像报告原型清单.可否重置格式%Type,
    可否书写多份_In 影像报告原型清单.可否书写多份%Type,
	是否禁用_In     影像报告原型清单.是否禁用%Type,
	创建人_In       影像报告原型清单.创建人%Type,
	内容_In         影像报告原型清单.内容%Type,
	控制选项_In     影像报告原型清单.控制选项%Type,
	词句加载时机_In 影像报告原型清单.词句加载时机%Type,
	插件加载时机_In 影像报告原型清单.插件加载时机%Type,
	专用插件_In     影像报告原型清单.专用插件%Type,
	Copy_Id_In      影像报告原型清单.ID%Type,
	Only_Head_In    Varchar2,
	分组_In         影像报告原型清单.分组%Type
	);
  --4.修改一个文档原型
  Procedure p_Edit_Antetypelist(
    Id_In           影像报告原型清单.ID%Type,
    种类_In         影像报告原型清单.种类%Type,
    编码_In         影像报告原型清单.编码%Type,
    名称_In         影像报告原型清单.名称%Type,
    设备号_In		影像设备目录.设备号%Type,
    说明_In         影像报告原型清单.说明%Type,
    可否重置页面_In 影像报告原型清单.可否重置页面%Type,
    可否重置格式_In 影像报告原型清单.可否重置格式%Type,
    可否书写多份_In 影像报告原型清单.可否书写多份%Type,
    是否禁用_In     影像报告原型清单.是否禁用%Type,
    修改人_In       影像报告原型清单.修改人%Type,
    内容_In         影像报告原型清单.内容%Type,
    控制选项_In     影像报告原型清单.控制选项%Type,
	词句加载时机_In 影像报告原型清单.词句加载时机%Type,
	插件加载时机_In 影像报告原型清单.插件加载时机%Type,
    专用插件_In     影像报告原型清单.专用插件%Type,
    Copy_Id_In      影像报告原型清单.ID%Type,
    Only_Head_In    Varchar2,
    分组_In         影像报告原型清单.分组%Type
	);
  --5.删除一个文件原型
  Procedure p_Del_Antetypelist(
    Id_In 影像报告原型清单.Id%Type
	);
  --6.根据ID获取文件原型
  Procedure p_Get_Antetypelist_By_Id(
	Val           Out t_Refcur,
	Id_In 影像报告原型清单.Id%Type
	);
  --7.获取原型XML内容
  Procedure p_Get_Antetypelist_Content(
	Val           Out t_Refcur,
	Id_In 影像报告原型清单.Id%Type
	);
  --8.停用或启用文件原型
  Procedure p_Stop_Antetypelist(
    Id_In 影像报告原型清单.Id%Type
	);

  --9.新增文档种类信息
  Procedure p_Add_Doc_Kind(
    编码_In 影像报告种类.编码%Type,
    名称_In 影像报告种类.名称%Type,
    说明_In 影像报告种类.说明%Type
	);
  --10.删除文档种类信息
  Procedure p_Del_Doc_Kind;
  --11.获取预备提纲信息
  Procedure p_Get_Pre_Outline(
    Val Out t_Refcur
	);
  --12.添加预备提纲信息
  Procedure p_Add_Pre_Outline(
    ID_In   影像报告预备提纲.ID%Type,
	编码_In 影像报告预备提纲.编码%Type,
	名称_In 影像报告预备提纲.名称%Type,
	说明_In 影像报告预备提纲.说明%Type
	);
  --13.删除预备提纲信息
  Procedure p_Del_Pre_Outline;
  --14.获取导出的文档原型信息
  Procedure p_Output_Antetypelist(
    Val Out t_Refcur
	);
  --15.添加原型片段
  Procedure p_Add_Antetype_Fragments(
    原型ID_In 影像报告原型片段.原型ID%Type,
	片段ID_In 影像报告原型片段.片段ID%Type
	);
  --16.删除原型片段
  Procedure p_Del_Antetype_Fragments(
    原型ID_In 影像报告原型片段.原型ID%Type
	);
  --17.获取原型片段
  Procedure p_Get_Antetype_Fragments(
    Val           Out t_Refcur,
	原型ID_In 影像报告原型片段.原型ID%Type
	);
  --18.获取某个原型关联的某个片段分类
  Procedure p_Get_Antetype_f_Byaidfid(
    Val           Out t_Refcur,		
	原型ID_In 影像报告原型片段.原型ID%Type,
    片段ID_In 影像报告原型片段.片段ID%Type
	);
  --19.插入文档原型XML内容
  Procedure p_Edit_Antetypelist_Content(
    Id_In     影像报告原型清单.Id%Type,
	内容_In   影像报告原型清单.内容%Type,
	修改人_In 影像报告原型清单.修改人%Type
	);
  --20.获取所有原型
  Procedure p_Get_All_Antetype_Lists(
    Val Out t_Refcur
	);
  --21.获取已经设置了关联的原型片段类别的信息

  Procedure p_Get_Antetype_Fragments_Info(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型片段.原型ID%Type
	);
  --22.获取选择的类别下面的短语名称
  Procedure p_Get_Selected_Fragments(
	Val           Out t_Refcur,
	原型id_In Varchar2
	);
  --23.获取能复制的原型名称

  Procedure p_Get_Copy_Antetype(
	Val           Out t_Refcur,
	种类_In 影像报告原型清单.种类%Type
	);
  --24.获取原型的分组信息
  Procedure p_Get_Antetype_Category(
	Val           Out t_Refcur,
	种类_In 影像报告原型清单.种类%Type
	);
  --25.根据原型同步范文提纲
  Procedure p_Synchronous_Sample(
    原型id_In 影像报告原型清单.Id%Type
	);
  --26.获取导出的原型列表
  Procedure p_Get_Out_Antetypelist(
    Val Out t_Refcur
	);
  --27.通过编码获取原型种类信息
  Procedure p_Get_Antetype_Kind_By_Code(
	Val           Out t_Refcur,
	编码_In 影像报告种类.编码%Type
	);
  --28.获取事件信息，不包含固定事件
  Procedure p_Get_Doc_Event(
    Val Out t_Refcur
	);
  --29.获取关于原型导出的重复信息

  Procedure p_Get_Antetypelist_Same_Info(
	Val           Out t_Refcur,
	Tablename_In Varchar2,
    Id_In        影像报告原型清单.Id%Type,
    编码_In      Varchar2,
    名称_In      Varchar2
	);
  --30.获取事件重复的信息
  Procedure p_Event_Same_Info(
	Val           Out t_Refcur,	
	Id_In      影像报告事件.Id%Type,
    原型ID_In  影像报告事件.原型ID%Type,
    元素IID_In 影像报告事件.元素IID%Type,
    种类_In    影像报告事件.种类%Type,
    名称_In    影像报告事件.名称%Type,
    编号_In    影像报告事件.编号%Type
	);
  --31.获取原型校验的类别集合
  Procedure p_Get_Process_Kind(
    Val Out t_Refcur
	);

  ----32.获取元素或者提纲的名称集合
  --Procedure p_Get_Antetype_Ele_Section(
  --原型ID_In  影像报告原型清单.Id%Type,
  --Val     Out t_Refcur);

  --33.获取指定原型的文档处理
  Procedure p_Get_Doc_Process_Of_Antetype(
	Val           Out t_Refcur,
	原型id_In 影像报告动作.原型id%Type
	);

  --34. 根据字典名称获取相应子项
  Procedure p_Get_Dictitems_By_Title(
	Val           Out t_Refcur,
	名称_In 影像字典清单.名称%Type
	);
  --35.获得所有的预备提纲
  Procedure p_Get_All_Phr_Onlines(
    Val Out t_Refcur
	);
  --36.获取所有词句信息
  Procedure p_Get_All_Fragment(
	Val           Out t_Refcur,
	学科_In Varchar2
	);

  --37.获取词句信息
  Procedure p_Get_Fragment_Filter(
	Val           Out t_Refcur,
	原型id_In 影像报告原型片段.原型ID%Type,
    作者_In   影像报告片段清单.作者%Type,
    学科_In   影像报告片段清单.学科%Type,
    Type_In   Varchar2
	);
  --38.根据原型获取关联的片段标签值
  Procedure p_Get_Label_By_Aid(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型片段.原型ID%Type
	);
  --39.获取所有词句分类
  Procedure p_Get_All_Fragment_Class(
    Val Out t_Refcur
	);
  --40.获取表名对应的最后编辑时间
  Procedure p_Get_Data_Last_Edit_Time(
	Val           Out t_Refcur,
	Table_Name_In Varchar2
	);
  --41.添加文档事件
  Procedure p_Add_Doc_Event(
    ID_In       影像报告事件.ID%Type,
    种类_In     影像报告事件.种类%Type,
    原型ID_In   影像报告事件.原型ID%Type,
    编号_In     影像报告事件.编号%Type,
    名称_In     影像报告事件.名称%Type,
    说明_In     影像报告事件.说明%Type,
    元素IID_In  影像报告事件.元素IID%Type,
    扩展标记_In 影像报告事件.扩展标记%Type);
  --42.修改文档事件
  Procedure p_Update_Doc_Event(
    Id_In       影像报告事件.Id%Type,
    种类_In     影像报告事件.种类%Type,
    名称_In     影像报告事件.名称%Type,
    说明_In     影像报告事件.说明%Type,
    元素IID_In  影像报告事件.元素IID%Type,
    扩展标记_In 影像报告事件.扩展标记%Type);
  --43.删除文档事件
  Procedure p_Delete_Doc_Event(
    Id_In 影像报告事件.Id%Type
	);
  --44.删除所有未被使用的文档事件
  Procedure p_Delete_Unused_Doc_Events(
    Count_Out Out Number
	);
  --45.获取指定原型的文档事件
  Procedure p_Get_Doc_Event_Of_Antetype(
	Val           Out t_Refcur,
	原型ID_In       影像报告事件.原型ID%Type,
	Include_Base_In Number
	);
  --46.修改文档处理编号
  Procedure p_Update_Doc_Process_Seqnum(
    Id_In   影像报告动作.Id%Type,
	序号_In 影像报告动作.序号%Type
	);
  --47.添加文档处理
  Procedure p_Add_Doc_Process(
    Id_In           影像报告动作.Id%Type,
    原型ID_In       影像报告动作.原型ID%Type,
    事件ID_In       影像报告动作.事件ID%Type,
    动作类型_In     影像报告动作.动作类型%Type,
    名称_In         影像报告动作.名称%Type,
    说明_In         影像报告动作.说明%Type,
    可否手工执行_In 影像报告动作.可否手工执行%Type,
    序号_In         影像报告动作.序号%Type,
    内容_In         影像报告动作.内容%Type
	);
  --48.修改文档处理
  Procedure p_Update_Doc_Process(
    Id_In           影像报告动作.Id%Type,
    事件ID_In       影像报告动作.事件ID%Type,
    动作类型_In     影像报告动作.动作类型%Type,
    名称_In         影像报告动作.名称%Type,
    说明_In         影像报告动作.说明%Type,
    可否手工执行_In 影像报告动作.可否手工执行%Type,
    内容_In         影像报告动作.内容%Type
	);
  --49.获取元素或者提纲的名称集合
  Procedure p_Get_Antetype_Ele_Section(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型清单.Id%Type,
	Type_In   Varchar2
	);
  --50.删除文档处理
  Procedure p_Del_Doc_Process(
    Id_In        影像报告动作.ID%Type,
	Del_Event_In Number
	);

  --51.查询元素值域类别的覆盖情况
  Procedure p_Get_Ele_Same_Info(
	Val           Out t_Refcur,
	Id_In    影像报告值域清单.Id%Type,
	Code_In  Varchar2,
	Title_In Varchar2,
	Flag_In  Varchar2
	);
  --52.获得所有的插件信息
  Procedure p_Get_DocPluginList(
    Val Out t_Refcur
	);
  --53.该ID的插件是否被原型使用过
  Procedure p_IsExit_DocPluginByID(
	Val           Out t_Refcur,
	ID_In Varchar2
	);
  --54.新增报告插件信息
  Procedure p_AddDocPlugin(
    ID_In       影像报告插件.ID%Type,
    编码_In     影像报告插件.编码%Type,
    名称_In     影像报告插件.名称%Type,
    说明_In     影像报告插件.说明%Type,
    显示样式_In 影像报告插件.显示样式%Type,
    种类_In     影像报告插件.种类%Type,
    类名_In     影像报告插件.类名%Type,
    库名_In     影像报告插件.库名%Type,
    是否禁用_In 影像报告插件.是否禁用%Type
	);
  --55.修改报告插件信息
  Procedure p_EditDocPlugin(
    ID_In       影像报告插件.ID%Type,
    编码_In     影像报告插件.编码%Type,
    名称_In     影像报告插件.名称%Type,
    说明_In     影像报告插件.说明%Type,
    显示样式_In 影像报告插件.显示样式%Type,
    种类_In     影像报告插件.种类%Type,
    类名_In     影像报告插件.类名%Type,
    库名_In     影像报告插件.库名%Type,
    是否禁用_In 影像报告插件.是否禁用%Type
	);
  --56.删除报告插件信息
  Procedure p_DelDocPlugin(
    ID_In 影像报告插件.ID%Type
	);
  --57.改变插件的可用状态
  Procedure p_IsEnableDocPlugin(
    ID_In 影像报告插件.ID%Type
	);
  --58.通过ID获得对应的插件信息
  Procedure p_GetDocPluginByID(
	Val           Out t_Refcur,
	ID_In 影像报告插件.ID%type
	);
  --59.判断编码和名称是否已存在
  Procedure p_IsExitDocPlugin(
	Val           Out t_Refcur,
	ID_In   影像报告插件.ID%Type,
    编码_In 影像报告插件.编码%Type,
    名称_In 影像报告插件.名称%Type
	);
  --60.通过ID获得对应的专用插件信息
  Procedure p_GetDocSpecPluginByID(
	Val           Out t_Refcur,
	ID_In 影像报告插件.ID%Type
	);
  --61.获得诊疗列表信息
  Procedure p_GetDiagnosisList(
	Val           Out t_Refcur,
	类别_In Varchar2,
    条件_In Varchar2
	);
  --62.获得诊疗类别列表
  Procedure p_GetDiagnosisClass(
    Val Out t_Refcur
	);
  --63.添加影像报告原型应用信息
  Procedure p_AddMedicalAntetype(
    诊疗项目ID_In 影像报告原型应用.诊疗项目ID%Type,
	应用场合_In   影像报告原型应用.应用场合%Type,
	报告原型ID_In 影像报告原型应用.报告原型ID%Type
	);
  --64.删除原型ID对应的病历单据应用信息
  Procedure p_DelMedicalAntetype(
    报告原型ID_In 影像报告原型应用.报告原型ID%Type
	);
  --65.通过原型ID获得对应的病历单据应用信息
  Procedure p_GetMedicalByAID(
	Val           Out t_Refcur,
	报告原型ID_In 影像报告原型应用.报告原型ID%Type
	);
  --66.根据原型ID删除动作信息
  Procedure p_DelDocProcessByAid(
    报告原型ID_In 影像报告原型应用.报告原型ID%Type
	);
  --67.获取ID对应的原型的树形结构
  Procedure p_GetAntetypeTreeByID(
	Val           Out t_Refcur,
	ID_In 影像报告原型清单.ID%Type
	);
  --68.原型是否存在对应的编码或名称
  procedure p_IsExitAntetype(
	Val           Out t_Refcur,
	编码_In 影像报告原型清单.编码%Type,
    名称_In 影像报告原型清单.名称%Type,
    ID_In  影像报告原型清单.ID%Type
	);

  --69  获取影像存储设备
  Procedure p_GetStorageDevice(
		Val           Out t_Refcur);

End b_PACS_RptAntetype;
/

--影像报告原型管理(---实现部分---)***************************************************
CREATE OR REPLACE Package Body b_PACS_RptAntetype Is
  --Create By Hwei;
  --2014/11/25

  --1.获取文件原型类别
  Procedure p_Get_Antetypelistkind(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select 编码, a.名称, a.编码 || '-' || a.名称 As 标题
        From 影像报告种类 A
       Order By 编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetypelistkind;

  --2.根据文档类型获取文档信息
  Procedure p_Get_Antetypelis_By_Kind(
	Val           Out t_Refcur,
	种类_In      影像报告原型清单.种类%Type,
    Stop_Flag    Number,
    Condition_In Varchar2
	) As
  Begin
    Open Val For
      Select ID, 编码, 名称, 标题, 分组, 是否禁用, 说明, Imageindex
        From (Select Distinct 分组 As ID,
                              (Select Min(b.编码)
                                 From 影像报告原型清单 B
                                Where b.分组 = a.分组) As 编码,
                              a.分组 As 名称,
                              a.分组 As 标题,
                              null As 分组,
                              0 As 是否禁用,
                              null As 说明,
                              0 As Imageindex
                From 影像报告原型清单 A
               Where a.种类 = 种类_In
                 And ((a.是否禁用 <> 1 And Stop_Flag = 1) Or (Stop_Flag = 0))
                 And ((a.名称 Like '%' || Condition_In || '%' And
                     Condition_In Is Not Null) Or Condition_In Is Null)
                 And a.分组 Is Not Null
              Union
              Select RawtoHex(ID) ID,
                     a.编码,
                     名称 As 名称,
                     编码 || '-' || 名称 As 标题,
                     分组,
                     a.是否禁用,
                     a.说明,
                     Decode(a.是否禁用, 1, 2, 1) Imageindex
                From 影像报告原型清单 A
               Where a.
               种类 = 种类_In
                 And ((a.是否禁用 <> 1 And Stop_Flag = 1) Or (Stop_Flag = 0))
                 And ((a.名称 Like '%' || Condition_In || '%' And
                     Condition_In Is Not Null) Or Condition_In Is Null)) A
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetypelis_By_Kind;

  --3.添加一个文档原型
  Procedure p_Add_Antetypelist(
    ID_In           影像报告原型清单.ID%Type,
    种类_In         影像报告原型清单.种类%Type,
    编码_In         影像报告原型清单.编码%Type,
    名称_In         影像报告原型清单.名称%Type,
	设备号_In		影像设备目录.设备号%Type,
    说明_In         影像报告原型清单.说明%Type,
    可否重置页面_In 影像报告原型清单.可否重置页面%Type,
    可否重置格式_In 影像报告原型清单.可否重置格式%Type,
	可否书写多份_In 影像报告原型清单.可否书写多份%Type,
    是否禁用_In     影像报告原型清单.是否禁用%Type,
    创建人_In       影像报告原型清单.创建人%Type,
    内容_In         影像报告原型清单.内容%Type,
    控制选项_In     影像报告原型清单.控制选项%Type,
	词句加载时机_In 影像报告原型清单.词句加载时机%Type,
	插件加载时机_In 影像报告原型清单.插件加载时机%Type,
    专用插件_In     影像报告原型清单.专用插件%Type,
    Copy_ID_In      影像报告原型清单.ID%Type,
    Only_Head_In    Varchar2,
    分组_In         影像报告原型清单.分组%Type
	) As
    x_Str Xmltype;
  Begin
    Begin
      If Copy_ID_In Is Null or Copy_ID_In = 0 Then
        x_Str := 内容_In;
      Else
        Select Decode(Only_Head_In,
                      1,
                      Deletexml(a.内容, '/zlxml/document/node()'),
                      a.内容)
          Into x_Str
          From 影像报告原型清单 A
         Where a.id = Copy_ID_In;
      End If;
    Exception
      When Others Then
        x_Str := 内容_In;
    End;
  
    Insert Into 影像报告原型清单
      (ID,
       种类,
       编码,
       名称,
	   设备号,
       说明,
       可否重置页面,
       可否重置格式,
	   可否书写多份,
       是否禁用,
       创建人,
       创建时间,
       内容,
       控制选项,
	   词句加载时机,
	   插件加载时机,
       专用插件,
       分组)
    Values
      (ID_In,
       种类_In,
       编码_In,
       名称_In,
	   设备号_In,
       说明_In,
       可否重置页面_In,
       可否重置格式_In,
	   可否书写多份_In,
       是否禁用_In,
       创建人_In,
       sysdate,
       x_Str,
       控制选项_In,
	   词句加载时机_In,
	   插件加载时机_In,
       专用插件_In,
       分组_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Antetypelist;

  --4.修改一个文档原型
  Procedure p_Edit_Antetypelist(
    ID_In           影像报告原型清单.ID%Type,
    种类_In         影像报告原型清单.种类%Type,
    编码_In         影像报告原型清单.编码%Type,
    名称_In         影像报告原型清单.名称%Type,
	设备号_In		影像设备目录.设备号%Type,
    说明_In         影像报告原型清单.说明%Type,
    可否重置页面_In 影像报告原型清单.可否重置页面%Type,
    可否重置格式_In 影像报告原型清单.可否重置格式%Type,
	可否书写多份_In 影像报告原型清单.可否书写多份%Type,
    是否禁用_In     影像报告原型清单.是否禁用%Type,
    修改人_In       影像报告原型清单.修改人%Type,
    内容_In         影像报告原型清单.内容%Type,
    控制选项_In     影像报告原型清单.控制选项%Type,
	词句加载时机_In 影像报告原型清单.词句加载时机%Type,
	插件加载时机_In 影像报告原型清单.插件加载时机%Type,
    专用插件_In     影像报告原型清单.专用插件%Type,
    Copy_ID_In      影像报告原型清单.ID%Type,
    Only_Head_In    Varchar2,
    分组_In         影像报告原型清单.分组%Type
	) As
    x_Str     Xmltype;
    n_Num     Number;
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    Select Count(a.Id)
      Into n_Num
      From 影像报告原型清单 A
     Where (a.编码 = 编码_In Or a.名称 = 名称_In)
       And ID <> ID_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]存在相同的文档编码或者名称，请重新填写！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If Copy_ID_In Is Null or Copy_ID_In = 0 Then
      x_Str := 内容_In;
    Else
      Select Decode(Only_Head_In,
                    1,
                    Deletexml(a.内容, '/zlxml/document/node()'),
                    a.内容)
        Into x_Str
        From 影像报告原型清单 A
       Where a.id = Copy_ID_In;
    End If;
  
    Update 影像报告原型清单
       Set 种类         = 种类_In,
           编码         = 编码_In,
           名称         = 名称_In,
		   设备号		= 设备号_In,
           说明         = 说明_In,
           可否重置页面 = 可否重置页面_In,
           可否重置格式 = 可否重置格式_In,
		   可否书写多份 = 可否书写多份_In,
           是否禁用     = NVL(是否禁用_In, 是否禁用),
           修改人       = 修改人_In,
           修改时间     = sysdate,
           内容         = x_Str,
           控制选项     = 控制选项_In,
		   词句加载时机 =词句加载时机_In,
		   插件加载时机 =插件加载时机_In,
           专用插件     = 专用插件_In,
           分组         = 分组_In
     Where ID = ID_In;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Antetypelist;

  --5.删除一个文件原型
  Procedure p_Del_Antetypelist(
    ID_In 影像报告原型清单.Id%Type
	) As
    n_Num     Number;
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    Select Count(ID) Into n_Num From 影像报告记录 A Where a.原型id = ID_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]该原型已经被文档使用，不允许删除！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Select Count(原型ID)
      Into n_Num
      From 影像报告原型片段
     Where 影像报告原型片段.原型ID = ID_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]该文档下存在词句关联，不允许删除！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Select Count(ID)
      Into n_Num
      From 影像报告范文清单
     Where 影像报告范文清单.原型ID = ID_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]存在以此原型建立的范文信息，不允许删除！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Delete From 影像报告原型清单 C Where c.Id = ID_In;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Antetypelist;

  --6.根据ID获取文件原型
  Procedure p_Get_Antetypelist_By_Id(
	Val           Out t_Refcur,
	ID_In 影像报告原型清单.Id%Type
	) As
  Begin
    Open Val For
      Select rawtohex(a.ID) ID,
             a.种类,
             a.编码,
             a.名称,
			 a.设备号,
             a.说明,
             a.可否重置页面,
             a.可否重置格式,
			 a.可否书写多份,
             Extractvalue(b.Column_Value, '/root/print_hf_mode') Printhfmode,
             Extractvalue(b.Column_Value, '/root/print_follow_pages') Printfollowpages,
             Nvl(Extractvalue(b.Column_Value, '/root/print_limit'), 0) Printlimit,
             (Nvl(a.控制选项, XmlType('<NULL/>'))).GetClobVal() as 控制选项,
			 a.词句加载时机,
			 a.插件加载时机,
             a.是否禁用,
             (Nvl(a.专用插件, XmlType('<NULL/>'))).GetClobVal() as 专用插件,
             a.创建人,
             a.创建时间,
             a.修改人,
             a.修改时间,
             a.分组
        From 影像报告原型清单 A,
             Table(Xmlsequence(Extract(a.控制选项, '/root'))) B
       Where a.Id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetypelist_By_Id;

  --7.获取原型XML内容
  Procedure p_Get_Antetypelist_Content(
	Val           Out t_Refcur,
	ID_In 影像报告原型清单.Id%Type
	) As
  Begin
    Open Val For
      Select (Nvl(a.内容, XmlType('<ZLXML/>'))).GetClobVal() As 内容
        From 影像报告原型清单 A
       Where a.Id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetypelist_Content;

  --8.停用或启用文件原型
  Procedure p_Stop_Antetypelist(
    ID_In 影像报告原型清单.Id%Type
	) As
  Begin
    Update 影像报告原型清单
       Set 是否禁用 = Decode(是否禁用, 1, 0, 0, 1)
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Stop_Antetypelist;

  --9.新增文档种类信息
  Procedure p_Add_Doc_Kind(
    编码_In 影像报告种类.编码%Type,
    名称_In 影像报告种类.名称%Type,
    说明_In 影像报告种类.说明%Type
	) As
    n_Num     Number;
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    Select Count(a.编码)
      Into n_Num
      From 影像报告种类 A
     Where a.编码 = 编码_In
        Or a.名称 = 名称_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]种类的编码或者名称不能相同！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If 编码_In Is Null Or 编码_In Is Null Then
      v_Err_Msg := '[ZLSOFT]种类的编码或者名称不能为空！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Insert Into 影像报告种类
      (编码, 名称, 说明)
    Values
      (编码_In, 名称_In, 说明_In);
  
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Doc_Kind;

  --10.删除文档种类信息
  Procedure p_Del_Doc_Kind As
  Begin
    Delete From 影像报告种类;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Doc_Kind;

  --11.获取预备提纲信息
  Procedure p_Get_Pre_Outline(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawtoHex(ID) ID, 编码, 名称, 说明, 最后编辑时间
        From 影像报告预备提纲 A
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Pre_Outline;

  --12.添加预备提纲信息
  Procedure p_Add_Pre_Outline(
    ID_In   影像报告预备提纲.ID%Type,
    编码_In 影像报告预备提纲.编码%Type,
    名称_In 影像报告预备提纲.名称%Type,
    说明_In 影像报告预备提纲.说明%Type
	) As
    n_Num     Number;
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    Select Count(a.编码)
      Into n_Num
      From 影像报告预备提纲 A
     Where a.编码 = 编码_In
        Or a.名称 = 名称_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]种类的编码或者名称不能相同！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If 编码_In Is Null Or 名称_In Is Null Then
      v_Err_Msg := '[ZLSOFT]种类的编码或者名称不能为空！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Insert Into 影像报告预备提纲
      (ID, 编码, 名称, 说明, 最后编辑时间)
    Values
      (ID_In, 编码_In, 名称_In, 说明_In, sysdate);
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Pre_Outline;

  --13.删除预备提纲信息
  Procedure p_Del_Pre_Outline As
  Begin
    Delete From 影像报告预备提纲;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Pre_Outline;

  --14.获取导出的文档原型信息
  Procedure p_Output_Antetypelist(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select '类别' As 类别,
             b.编码 As ID,
             Null As 种类,
             b.名称 As 种类名称,
             b.编码 As 编码,
             b.名称 As 名称,
             b.说明 As 说明,
             Null As 可否重置页面,
             Null As 可否重置格式,
             Null As 是否禁用,
             Null As 创建人,
             Null As 创建时间,
             Null As 修改人,
             Null As 修改时间,
             Null As 内容
        From 影像报告种类 B
      Union All
      Select '原型' 类别,
             RawToHex(a.Id) ID,
             a.种类,
             b.名称 种类名称,
             a.编码,
             a.名称,
             a.说明,
             a.可否重置页面,
             a.可否重置格式,
             a.是否禁用,
             a.创建人,
             a.创建时间,
             a.修改人,
             a.修改时间,
             Null As 内容
        From 影像报告原型清单 A, 影像报告种类 B
       Where a.种类 = b.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Output_Antetypelist;

  --15.添加原型片段
  Procedure p_Add_Antetype_Fragments(
    原型ID_In 影像报告原型片段.原型ID%Type,
	片段ID_In 影像报告原型片段.片段ID%Type) As
  Begin
    Insert Into 影像报告原型片段
      (原型ID, 片段ID)
    Values
      (原型ID_In, 片段ID_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Antetype_Fragments;

  --16.删除原型片段
  Procedure p_Del_Antetype_Fragments(
    原型ID_In 影像报告原型片段.原型ID%Type
	) As
  Begin
    Delete From 影像报告原型片段 Where 原型ID = 原型ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Antetype_Fragments;

  --17.获取原型片段
  Procedure p_Get_Antetype_Fragments(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型片段.原型ID%Type
	) As
  Begin
    Open Val For
      Select RawtoHex(片段ID) 片段ID
        From 影像报告原型片段 A
       Where a.原型ID = 原型ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_Fragments;

  --18.获取某个原型关联的某个片段分类
  Procedure p_Get_Antetype_f_Byaidfid(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型片段.原型ID%Type,
	片段ID_In 影像报告原型片段.片段ID%Type
	) As
  Begin
    Open Val For
      Select RawtoHex(片段ID) 片段ID
        From 影像报告原型片段 A
       Where a.原型ID = 原型ID_In
         And a.片段ID = 片段ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_f_Byaidfid;

  --19.插入文档原型XML内容
  Procedure p_Edit_Antetypelist_Content(
    ID_In     影像报告原型清单.Id%Type,
	内容_In   影像报告原型清单.内容%Type,
	修改人_In 影像报告原型清单.修改人%Type
	) As
  Begin
    Update 影像报告原型清单
       Set 内容 = 内容_In, 修改人 = 修改人_In, 修改时间 = Sysdate
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Antetypelist_Content;

  --20.获取所有原型
  Procedure p_Get_All_Antetype_Lists(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawtoHex(ID) ID,
             a.编码,
             编码 || '-' || 名称 As 名称,
             分组,
             a.种类,
             a.是否禁用,
             a.说明,
             Decode(a.是否禁用, 1, 2, 1) Imageindex,
             (Nvl(a.内容, XmlType('<ZLXML/>'))).GetClobVal() As 内容
        From 影像报告原型清单 A
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_All_Antetype_Lists;

  --21.获取已经设置了关联的原型片段类别的信息
  Procedure p_Get_Antetype_Fragments_Info(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型片段.原型ID%Type
	) As
  Begin
    Open Val For
      Select RawtoHex(ID) ID,
             a.编码,
             a.名称,
             a.编码 || '-' || a.名称 标题,
             a.说明
        From 影像报告片段清单 A
       Where a.Id In (Select b.片段id
                        From 影像报告原型片段 B
                       Where b.原型id = 原型ID_In)
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_Fragments_Info;

  --22.获取选择的类别下面的短语名称
  Procedure p_Get_Selected_Fragments(
	Val           Out t_Refcur,
	原型ID_In Varchar2
	) As
    v_Sql  Varchar2(4000);
    v_Aids Varchar2(4000);
    v_Msg  Varchar2(4000);
    Err Exception;
  Begin
    For Myrow In (Select RawtoHex(a.片段id) ID
                    From 影像报告原型片段 A
                   Where a.原型id = 原型ID_In) Loop
      If v_Aids Is Null Then
        v_Aids := '''' || Myrow.Id || '''';
      Else
        v_Aids := v_Aids || ',''' || Myrow.Id || '''';
      End If;
    End Loop;
  
    If v_Aids Is Null Then
      If Substr(原型ID_In, 0, 1) <> '''' Then
        v_Aids := '''' || 原型ID_In || '''';
      Else
        v_Aids := 原型ID_In;
      End If;
    End If;
  
    v_Sql := 'Select Distinct  RawtoHex(a.id) ID,  RawtoHex(a.上级ID) 上级ID , a.编码, a.编码 || ''-'' || a.名称 标题,Decode(a.节点类型, 0, 0, 1) 节点类型
      From 影像报告片段清单 A
      Start With a.Id In (' || v_Aids || ')
      Connect By Prior a.Id = a.上级ID
      Order By a.编码';
  
    Open Val For v_Sql;
  Exception
    When Err Then
      Raise_Application_Error(-20101, v_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Selected_Fragments;

  --23.获取能复制的原型名称
  Procedure p_Get_Copy_Antetype(
	Val           Out t_Refcur,
	种类_In 影像报告原型清单.种类%Type
	) As
  Begin
    Open Val For
      Select RawtoHex(a.id) ID, a.编码 || '-' || a.名称 标题
        From 影像报告原型清单 A
       Where a.种类 = 种类_In
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Copy_Antetype;

  --24.获取原型的分组信息
  Procedure p_Get_Antetype_Category(
	Val           Out t_Refcur,
	种类_In 影像报告原型清单.种类%Type
	) As
  Begin
    Open Val For
      Select Distinct a.分组 As 分组
        From 影像报告原型清单 A
       Where a.种类 = 种类_In
         and a.分组 Is Not Null;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_Category;

  --25.根据原型同步范文提纲
  Procedure p_Synchronous_Sample(
    原型ID_In 影像报告原型清单.Id%Type
	) As
    x_Content Xmltype;
    x_Result  Xmltype;
    Cursor c_Antetype Is
      Select Extractvalue(c.Column_Value, '/section/@iid') Iid,
             Extractvalue(c.Column_Value, '/section/@title') Title,
             c.Column_Value As Content
        From 影像报告原型清单 A,
             Table(Xmlsequence(Extract(a.内容, '/zlxml//section'))) C
       Where a.Id = 原型ID_In;
    n_i               Number;
    n_j               Number;
    n_Count           Number;
    x_Subdocuments    Xmltype;
    x_Docparameters   Xmltype;
    x_Antetypecontent Xmltype;
    v_Textstyleno     Varchar2(10);
    v_Parastyleno     Varchar2(10);
    x_Acontent        Xmltype;
  Begin
    For Mysample In (Select b.id, b.内容
                       From 影像报告范文清单 B
                      Where b.原型id = 原型ID_In) Loop
      x_Content := Mysample.内容;
      n_i       := 1;
      If x_Content Is Null Then
        Select a.内容
          Into x_Result
          From 影像报告原型清单 A
         Where a.Id = 原型ID_In;
      Else
        Begin
          Select Extractvalue(c.Column_Value, '/section/@textstyleno') Textstyleno,
                 Extractvalue(c.Column_Value, '/section/@parastyleno') Parastyleno
            Into v_Textstyleno, v_Parastyleno
            From Table(Xmlsequence(Extract(x_Content, '/zlxml//section'))) C
           Where Rownum = 1;
        Exception
          When Others Then
            v_Textstyleno := '1';
            v_Parastyleno := '1';
        End;
      
        For Myantetype In c_Antetype Loop
          For I In 1 .. 1 Loop
            If n_i <> 1 Or n_Count <> 0 Or n_Count Is Null Then
              Select Count(*)
                Into n_Count
                From Table(Xmlsequence(Extract(x_Content, '/zlxml//section'))) C;
            End If;
            If n_Count < n_i Then
              Select Updatexml(Myantetype.Content,
                               '//section/@textstyleno',
                               v_Textstyleno)
                Into x_Acontent
                From Dual;
              Select Updatexml(x_Acontent,
                               '//section/@parastyleno',
                               v_Parastyleno)
                Into x_Acontent
                From Dual;
              Select Appendchildxml(x_Content,
                                    '/zlxml/document',
                                    x_Acontent)
                Into x_Content
                From Dual;
              Exit;
            End If;
            n_j := 1;
            For Mysample In (Select Extractvalue(c.Column_Value,
                                                 '/section/@iid') Iid,
                                    Extractvalue(c.Column_Value,
                                                 '/section/@title') Title
                               From Table(Xmlsequence(Extract(x_Content,
                                                              '/zlxml//section'))) C) Loop
              If n_i = n_j Then
                If Myantetype.Iid <> Mysample.Iid Then
                  Select Updatexml(Myantetype.Content,
                                   '//section/@textstyleno',
                                   v_Textstyleno)
                    Into x_Acontent
                    From Dual;
                  Select Updatexml(x_Acontent,
                                   '//section/@parastyleno',
                                   v_Parastyleno)
                    Into x_Acontent
                    From Dual;
                  Select Deletexml(x_Content,
                                   '//section[@iid="' || Myantetype.Iid || '"]')
                    Into x_Content
                    From Dual;
                  Select Insertxmlbefore(x_Content,
                                         '//section[@iid="' || Mysample.Iid || '"]',
                                         x_Acontent)
                    Into x_Content
                    From Dual;
                  n_j := n_j + 1;
                  Exit;
                Else
                  n_j := n_j + 1;
                  Exit;
                End If;
              End If;
              n_j := n_j + 1;
            End Loop;
            n_i := n_i + 1;
          End Loop;
        End Loop;
        x_Result := x_Content;
        For Mysample2 In (Select Iid
                            From (Select Extractvalue(c.Column_Value,
                                                      '/section/@iid') Iid
                                    From Table(Xmlsequence(Extract(x_Content,
                                                                   '/zlxml//section'))) C) C
                           Where c.Iid Not In
                                 (Select Extractvalue(c.Column_Value,
                                                      '/section/@iid') Iid
                                    From 影像报告原型清单 A,
                                         Table(Xmlsequence(Extract(a.内容,
                                                                   '/zlxml//section'))) C
                                   Where a.Id = 原型ID_In)) Loop
          Select Deletexml(x_Result,
                           '//section[@iid="' || Mysample2.Iid || '"]')
            Into x_Result
            From Dual;
        End Loop;
      End If;
    
      Update 影像报告范文清单 X
         Set x.内容 = x_Result
       Where x.Id = Mysample.Id;
    End Loop;
  
    Select a.内容
      Into x_Antetypecontent
      From 影像报告原型清单 A
     Where a.Id = 原型ID_In;
    Select Extract(x_Antetypecontent, 'zlxml/subdocuments')
      Into x_Subdocuments
      From Dual;
    Select Extract(x_Antetypecontent, 'zlxml/docparameters')
      Into x_Docparameters
      From Dual;
    Update 影像报告范文清单
       Set 内容 = Updatexml(内容, '/zlxml/subdocuments', x_Subdocuments)
     Where 原型ID = 原型ID_In;
    Update 影像报告范文清单
       Set 内容 = Updatexml(内容, '/zlxml/docparameters', x_Docparameters)
     Where 原型ID = 原型ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Synchronous_Sample;

  --26.获取导出的原型列表
  Procedure p_Get_Out_Antetypelist(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select ID,
             编码,
             标题,
             Parentid,
             种类,
             是否禁用,
             说明,
             Imageindex,
             名称
        From (Select a.编码 As ID,
                     a.编码 As 编码,
                     a.名称 As 标题,
                     '' As Parentid,
                     '-1' As 种类,
                     0 As 是否禁用,
                     a.说明 As 说明,
                     4 As Imageindex,
                     a.名称 名称
                From 影像报告种类 A
              Union
              Select Distinct a.种类 || '-' || a.分组 As ID,
                              (Select Min(编码)
                                 From 影像报告原型清单 B
                                Where b.分组 = a.分组) As 编码,
                              Max(a.分组) As 名称,
                              a.种类 As Parentid,
                              '0' As 种类,
                              0 As 是否禁用,
                              '' As 说明,
                              4 As Imageindex,
                              a.分组
                From 影像报告原型清单 A
               Where a.分组 Is Not Null
               Group By a.种类, a.分组
              Union
              Select RawTohex(ID),
                     a.编码,
                     编码 || '-' || 名称 As 标题,
                     Decode(a.分组, Null, a.种类, a.种类 || '-' || a.分组) Parentid,
                     a.种类 As 种类,
                     a.是否禁用,
                     a.说明,
                     Decode(a.是否禁用, 1, 1, 0, 2),
                     a.名称
                From 影像报告原型清单 A) A
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Out_Antetypelist;

  --27.通过编码获取原型种类信息
  Procedure p_Get_Antetype_Kind_By_Code(
	Val           Out t_Refcur,
	编码_In 影像报告种类.编码%Type
	) As
  Begin
    Open Val For
      Select a.编码, a.名称, a.说明
        From 影像报告种类 A
       Where a.编码 = 编码_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_Kind_By_Code;
  --28.获取事件信息，不包含固定事件
  Procedure p_Get_Doc_Event(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawtoHex(a.id) ID,
             a.种类,
             a.原型id,
             a.编号,
             a.名称,
             a.说明,
             a.元素iid,
             a.扩展标记
        From 影像报告事件 A
       Where a.种类 <> 1;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Doc_Event;

  --29.获取关于原型导出的重复信息
  Procedure p_Get_Antetypelist_Same_Info(
	Val           Out t_Refcur,
	Tablename_In Varchar2,
	ID_In        影像报告原型清单.Id%Type,
	编码_In      Varchar2,
	名称_In      Varchar2
	) As
    n_Num    Number;
    v_Result Varchar2(100);
    v_Sql    Varchar2(4000);
  Begin
    If ID_In Is Not Null Then
      v_Sql := 'select count(*) from ' || Tablename_In || ' where id=' ||
               ID_In;
      Execute Immediate v_Sql
        Into n_Num;
      If n_Num > 0 Then
        v_Result := 'ID重复';
      End If;
    End If;
    If 编码_In Is Not Null Then
      v_Sql := 'select count(*) from ' || Tablename_In || ' where 编码=''' ||
               编码_In || '''';
      Execute Immediate v_Sql
        Into n_Num;
      If n_Num > 0 Then
        If v_Result Is Not Null Then
          v_Result := v_Result || ',编码重复';
        Else
          v_Result := '编码重复';
        End If;
      End If;
    End If;
    If 名称_In Is Not Null Then
      v_Sql := 'select count(*) from ' || Tablename_In || ' where 名称=''' ||
               名称_In || '''';
      Execute Immediate v_Sql
        Into n_Num;
      If n_Num > 0 Then
        If v_Result Is Not Null Then
          v_Result := v_Result || ',名称重复';
        Else
          v_Result := '名称重复';
        End If;
      End If;
    End If;
    Open Val For
      Select v_Result Result From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetypelist_Same_Info;

  --30.获取事件重复的信息
  Procedure p_Event_Same_Info(
	Val           Out t_Refcur,
	ID_In      影像报告事件.Id%Type,
    原型ID_In  影像报告事件.原型ID%Type,
    元素IID_In 影像报告事件.元素IID%Type,
    种类_In    影像报告事件.种类%Type,
    名称_In    影像报告事件.名称%Type,
    编号_In    影像报告事件.编号%Type
	) As
    v_Same_Antetype Varchar2(50);
    n_Same_Id       Number;
    n_Same_Title    Number;
    n_Same_Seqnum   Number;
    n_Maxnum        Number;
  Begin
    Select Count(*)
      Into n_Same_Title
      From 影像报告事件 A
     Where a.原型ID = 原型ID_In
       And a.种类 = 种类_In
       And a.名称 = 名称_In;
    Select Count(*)
      Into n_Same_Seqnum
      From 影像报告事件 A
     Where a.原型ID = 原型ID_In
       And a.种类 = 种类_In
       And a.编号 = 编号_In;
    Begin
      Select a.Id
        Into v_Same_Antetype
        From 影像报告事件 A
       Where a.原型ID = 原型ID_In
         And a.元素IID = 元素IID_In;
    Exception
      When Others Then
        v_Same_Antetype := '';
    End;
  
    Select Count(*) Into n_Same_Id From 影像报告事件 A Where a.Id = ID_In;
    Select Max(a.编号) Into n_Maxnum From 影像报告事件 A;
  
    Open Val For
      Select v_Same_Antetype As Sameaid,
             n_Same_Id       As Sameid,
             n_Same_Title    As Sametitle,
             n_Same_Seqnum   As Sameseqnum,
             n_Maxnum        As Maxnum
        From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Event_Same_Info;

  --31.获取原型校验的类别集合
  Procedure p_Get_Process_Kind(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Distinct 动作类型
        From (Select Extractvalue(c.Column_Value, '/step/kind') As 动作类型
                From 影像报告动作 A,
                     Table(Xmlsequence(Extract(a.内容, '/root/step'))) C) B
       Where b.动作类型 Is Not Null;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Process_Kind;


  --33.获取指定原型的文档处理
  Procedure p_Get_Doc_Process_Of_Antetype(
	Val           Out t_Refcur,
	原型ID_In 影像报告动作.原型id%Type
	) As
  Begin
    Open Val For
      Select RawtoHex(p.id) ID,
             p.名称,
             p.动作类型,
             p.序号,
             p.说明,
             p.可否手工执行,
             (Nvl(p.内容, XmlType('<NULL/>'))).GetClobVal() As 内容, --Nvl(p.内容,'<NULL/>') As 内容,
             RawtoHex(p.事件ID) 事件ID,
             0 Is_Event
        From 影像报告动作 P
       Where p.原型ID = 原型ID_In
      Union All
      Select RawtoHex(e.id) ID,
             e.名称,
             e.种类,
             e.编号,
             e.说明,
             Null,
             (XmlType('<Null/>')).GetClobVal() As 内容, --(Null,'<NULL/>') As 内容,
             Null,
             1
        From 影像报告事件 E
       Where e.Id In (Select RawtoHex(事件ID) 事件ID
                        From 影像报告动作
                       Where 原型ID = 原型ID_In)
       Order By Is_Event, 动作类型, 序号;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Doc_Process_Of_Antetype;

  --34. 根据字典名称获取相应子项
  Procedure p_Get_Dictitems_By_Title(
	Val           Out t_Refcur,
	名称_In 影像字典清单.名称%Type
	) As
  Begin
    Open Val For
      Select a.编号, a.名称, Rawtohex(a.字典id) As 字典ID
        From 影像字典内容 A
       Where a.字典id In (Select id From 影像字典清单 b Where b.名称 = 名称_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Dictitems_By_Title;

  --35.获得所有的预备提纲
  Procedure p_Get_All_Phr_Onlines(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(ID) ID, a.编码, a.名称
        From 影像报告预备提纲 a
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_All_Phr_Onlines;

  --36.获取所有词句信息
  Procedure p_Get_All_Fragment(
	Val           Out t_Refcur,
	学科_In Varchar2
	) As
  Begin
    If 学科_In <> '' Then
      Open Val For
        Select RawToHex(a.id) ID,
               RawToHex(a.上级id) 上级id,
               a.编码,
               a.名称,
               a.说明,
               a.节点类型,
               (Nvl(a.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
               a.学科,
               a.标签,
               a.是否私有,
               a.作者
          From 影像报告片段清单 A
         Where (a.学科 In
               (Select /*+rule*/
                  Column_Value As Lable
                   From Table(b_PACS_RptPublic.f_Str2list(学科_In, ','))
                 Intersect
                 Select /*+rule*/
                  Column_Value As Lable
                   From Table(b_PACS_RptPublic.f_Str2list(a.学科, ','))) And
               a.节点类型 <> 0)
            Or a.节点类型 = 0
            Or a.学科 Is Null
         Order By a.编码, a.上级id;
    Else
      Open Val For
        Select RawToHex(a.id) ID,
               RawToHex(a.上级id) 上级id,
               a.编码,
               a.名称,
               a.说明,
               a.节点类型,
               (Nvl(a.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
               a.学科,
               a.标签,
               a.是否私有,
               a.作者
          From 影像报告片段清单 A
         Order By a.上级id, a.节点类型, a.编码, a.名称;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_All_Fragment;

  --37. 获取词句信息
  Procedure p_Get_Fragment_Filter(
	Val           Out t_Refcur,
	原型id_In 影像报告原型片段.原型ID%Type,
    作者_In   影像报告片段清单.作者%Type,
    学科_In   影像报告片段清单.学科%Type,
    Type_In   Varchar2
	) As
  Begin
    If Type_In = '1' Then
      Open Val For
        Select Rawtohex(b.Id) ID,
               Rawtohex(b.上级id) 上级id,
               b.编码,
               b.名称,
               b.说明,
               b.节点类型,
               (Nvl(b.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
               b.学科,
               b.标签,
               b.是否私有,
               b.作者,
               b.最后编辑时间
          From 影像报告原型片段 A, 影像报告片段清单 B
         Where a.片段id = b.id
           And a.原型id = 原型id_In;
    Else
      Open Val For
        Select /*+ rule*/
         Rawtohex(b.Id) ID,
         Rawtohex(b.上级id) 上级id,
         b.编码,
         b.名称,
         b.说明,
         b.节点类型,
         (Nvl(b.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
         b.学科,
         b.标签,
         b.是否私有,
         b.作者,
         b.最后编辑时间
          From 影像报告片段清单 B
         Where b.上级id = 原型id_In
           And (b.是否私有 = 0 Or (b.是否私有 = 1 And b.作者 = 作者_In))
           And (b.学科 Is Null Or
               (b.学科 Is Not Null And
               b_PACS_RptPublic.f_If_Intersect(b.学科, 学科_In) > 0));
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Fragment_Filter;

  --38.根据原型获取关联的片段标签值
  Procedure p_Get_Label_By_Aid(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型片段.原型ID%Type
	) As
  Begin
    Open Val For
      Select Distinct b.标签
        From 影像报告片段清单 B
       Start With b.上级id In (Select a.片段id
                               From 影像报告原型片段 A
                              Where a.原型id = 原型ID_In)
      Connect By Prior b.Id = b.上级id;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Label_By_Aid;

  --39.获取所有词句分类
  Procedure p_Get_All_Fragment_Class(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(a.Id) ID,
             Rawtohex(a.上级id) 上级id,
             a.编码,
             a.名称,
             a.说明,
             a.节点类型
        From 影像报告片段清单 A
       Where a.节点类型 = 0
       Start With 上级id Is Null
      Connect By Prior id = 上级id
       Order By 编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_All_Fragment_Class;

  --40.获取表名对应的最后编辑时间
  Procedure p_Get_Data_Last_Edit_Time(
	Val           Out t_Refcur,
	Table_Name_In Varchar2
	) As
    v_sql Varchar2(4000);
  Begin
    v_sql := 'select max(最后编辑时间) maxvalue from ' || Table_Name_In;
    Open val For v_sql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Data_Last_Edit_Time;

  --41.添加文档事件
  Procedure p_Add_Doc_Event(
    ID_In       影像报告事件.ID%Type,
    种类_In     影像报告事件.种类%Type,
    原型ID_In   影像报告事件.原型ID%Type,
    编号_In     影像报告事件.编号%Type,
    名称_In     影像报告事件.名称%Type,
    说明_In     影像报告事件.说明%Type,
    元素IID_In  影像报告事件.元素IID%Type,
    扩展标记_In 影像报告事件.扩展标记%Type
	) As
    n_Seq_Num  影像报告事件.编号%Type;
    n_Is_Exist Number(1) := 0;
    v_Err_Msg  Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(*)
      Into n_Is_Exist
      From 影像报告事件
     Where 原型ID = 原型ID_In
       And 种类 = 种类_In
       And 名称 = 名称_In;
  
    If n_Is_Exist > 0 Then
      v_Err_Msg := '[ZLSOFT]原型上已存在相同命名的事件[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If (编号_In Is Null Or 编号_In = 0) Then
      Select Nvl(Max(编号), 0) + 1 Into n_Seq_Num From 影像报告事件;
    Else
      Select Count(*)
        Into n_Is_Exist
        From 影像报告事件
       Where 原型ID = 原型ID_In
         And 种类 = 种类_In
         And 编号 = 编号_In;
      If n_Is_Exist > 0 Then
        v_Err_Msg := '[ZLSOFT]原型上已存在相同编号的事件[ZLSOFT]';
        Raise Err_Item;
      End If;
      n_Seq_Num := 编号_In;
    End If;
  
    Insert Into 影像报告事件
      (ID, 种类, 原型ID, 编号, 名称, 说明, 元素IID, 扩展标记)
    Values
      (ID_In,
       种类_In,
       原型ID_In,
       n_Seq_Num,
       名称_In,
       说明_In,
       元素IID_In,
       扩展标记_In);
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Doc_Event;

  --42.修改文档事件
  Procedure p_Update_Doc_Event(
    Id_In       影像报告事件.Id%Type,
    种类_In     影像报告事件.种类%Type,
    名称_In     影像报告事件.名称%Type,
    说明_In     影像报告事件.说明%Type,
    元素IID_In  影像报告事件.元素IID%Type,
    扩展标记_In 影像报告事件.扩展标记%Type
	) As
    r_Aid      影像报告事件.原型ID%Type;
    n_Is_Exist Number(1) := 0;
    v_Err_Msg  Varchar2(100);
    Err_Item Exception;
  Begin
    Select 原型ID Into r_Aid From 影像报告事件 Where ID = Id_In;
  
    Select Count(*)
      Into n_Is_Exist
      From 影像报告事件
     Where 原型ID = r_Aid
       And 种类 = 种类_In
       And 名称 = 名称_In
       And ID <> Id_In;
  
    If n_Is_Exist > 0 Then
      v_Err_Msg := '[ZLSOFT]原型上存在相同命名的事件[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Update 影像报告事件
       Set 种类     = 种类_In,
           名称     = 名称_In,
           说明     = 说明_In,
           元素IID  = 元素IID_In,
           扩展标记 = 扩展标记_In
     Where ID = Id_In;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Doc_Event;

  --43.删除文档事件
  Procedure p_Delete_Doc_Event(
    Id_In 影像报告事件.Id%Type
	) As
    n_Kind     影像报告事件.种类%Type;
    n_Is_Exist Number(1) := 0;
    v_Err_Msg  Varchar2(100);
    Err_Item Exception;
  Begin
    Select 种类 Into n_Kind From 影像报告事件 Where ID = Id_In;
  
    If n_Kind = 1 Then
      v_Err_Msg := '[ZLSOFT]不允许删除固定事件[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Select Count(*) Into n_Is_Exist From 影像报告动作 Where 事件ID = Id_In;
  
    If n_Is_Exist > 0 Then
      v_Err_Msg := '[ZLSOFT]事件已经被使用,不能被删除！[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Delete From 影像报告事件 Where ID = Id_In;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Delete_Doc_Event;

  --44.删除所有未被使用的文档事件
  Procedure p_Delete_Unused_Doc_Events(
    Count_Out Out Number
	) As
  Begin
    Delete From 影像报告事件
     Where 种类 <> 1
       And ID Not In
           (Select 事件ID From 影像报告动作 Where 事件ID Is Not Null);
    Count_Out := Sql%RowCount;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Delete_Unused_Doc_Events;

  --45.获取指定原型的文档事件
  Procedure p_Get_Doc_Event_Of_Antetype(
	Val           Out t_Refcur,
	原型ID_In       影像报告事件.原型ID%Type,
	Include_Base_In Number
	) As
  Begin
    If Include_Base_In = 1 Then
      Open Val For
        Select Rawtohex(t.Id) ID,
               t.种类,
               t.名称,
               t.说明,
               t.元素iid,
               t.扩展标记,
               Nvl(p.Used_Count, 0) Used_Count
          From 影像报告事件 T,
               (Select Count(*) Used_Count, Max(事件ID) 事件ID
                  From 影像报告动作
                 Where 事件ID Is Not Null
                 Group By 事件ID) P
         Where (t.种类 = 1 Or t.原型id = 原型ID_In)
           And t.Id = p.事件ID(+)
         Order By t.编号;
    Else
      Open Val For
        Select Rawtohex(t.Id) ID,
               t.种类,
               t.名称,
               t.说明,
               t.元素iid,
               t.扩展标记,
               Nvl(p.Used_Count, 0) Used_Count
          From 影像报告事件 T,
               (Select Count(*) Used_Count, Max(事件ID) 事件ID
                  From 影像报告动作
                 Where 事件ID Is Not Null
                 Group By 事件ID) P
         Where t.原型id = 原型ID_In
           And t.种类 <> 1
           And t.Id = p.事件ID(+)
         Order By t.编号;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Doc_Event_Of_Antetype;

  --46.修改文档处理编号
  Procedure p_Update_Doc_Process_Seqnum(
    Id_In   影像报告动作.Id%Type,
	序号_In 影像报告动作.序号%Type) As
  Begin
    Update 影像报告动作 Set 序号 = 序号_In Where ID = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Doc_Process_Seqnum;

  --47.添加文档处理
  Procedure p_Add_Doc_Process(
    Id_In           影像报告动作.Id%Type,
    原型ID_In       影像报告动作.原型ID%Type,
    事件ID_In       影像报告动作.事件ID%Type,
    动作类型_In     影像报告动作.动作类型%Type,
    名称_In         影像报告动作.名称%Type,
    说明_In         影像报告动作.说明%Type,
    可否手工执行_In 影像报告动作.可否手工执行%Type,
    序号_In         影像报告动作.序号%Type,
    内容_In         影像报告动作.内容%Type
	) As
    n_Seq_Num  影像报告动作.序号%Type;
    n_Is_Exist Number(1) := 0;
    v_Err_Msg  Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(*)
      Into n_Is_Exist
      From 影像报告动作
     Where 原型ID = 原型ID_In
       And 名称 = 名称_In;
    If (序号_In Is Null Or 序号_In = 0) Then
      If (事件ID_In Is Null) Then
        Select Nvl(Max(序号), 0) + 1
          Into n_Seq_Num
          From 影像报告动作
         Where 原型ID = 原型ID_In
           And 事件ID Is Null;
      Else
        Select Nvl(Max(序号), 0) + 1
          Into n_Seq_Num
          From 影像报告动作
         Where 原型ID = 原型ID_In
           And 事件ID = 事件ID_In;
      End If;
    Else
      n_Seq_Num := 序号_In;
    End If;
    If n_Is_Exist > 0 Then
      v_Err_Msg := '[ZLSOFT]原型上存在相同命名的动作[ZLSOFT]';
      Raise Err_Item;
    End If;
    Insert Into 影像报告动作
      (ID, 原型ID, 事件ID, 动作类型, 名称, 说明, 可否手工执行, 序号, 内容)
    Values
      (Id_In,
       原型ID_In,
       事件ID_In,
       动作类型_In,
       名称_In,
       说明_In,
       可否手工执行_In,
       n_Seq_Num,
       内容_In);
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Doc_Process;

  --48.修改文档处理
  Procedure p_Update_Doc_Process(
    Id_In           影像报告动作.Id%Type,
    事件ID_In       影像报告动作.事件ID%Type,
    动作类型_In     影像报告动作.动作类型%Type,
    名称_In         影像报告动作.名称%Type,
    说明_In         影像报告动作.说明%Type,
    可否手工执行_In 影像报告动作.可否手工执行%Type,
    内容_In         影像报告动作.内容%Type
	) As
    r_Aid          影像报告事件.原型ID%Type;
    r_Old_Event_Id 影像报告动作.事件ID%Type;
    n_Seq_Num      影像报告事件.编号%Type;
    n_Is_Exist     Number(1) := 0;
    v_Err_Msg      Varchar2(100);
    Err_Item Exception;
  Begin
    Select 原型ID Into r_Aid From 影像报告动作 Where ID = Id_In;
    If (事件ID_In Is Not Null) Then
      Select Count(*)
        Into n_Is_Exist
        From 影像报告事件
       Where (原型ID Is Null Or 原型ID = r_Aid)
         And ID = 事件ID_In;
    
      If n_Is_Exist = 0 Then
        v_Err_Msg := '[ZLSOFT]关联的事件不存在[ZLSOFT]';
        Raise Err_Item;
      End If;
    
    End If;
  
    Select Count(*)
      Into n_Is_Exist
      From 影像报告动作
     Where 原型ID = r_Aid
       And 名称 = 名称_In
       And ID <> Id_In;
  
    If n_Is_Exist > 0 Then
      v_Err_Msg := '[ZLSOFT]原型上存在相同命名的动作[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If (r_Old_Event_Id <> 事件ID_In Or
       (事件ID_In Is Null And r_Old_Event_Id Is Not Null)) Then
      If (事件ID_In Is Null) Then
        Select Nvl(Max(序号), 0) + 1
          Into n_Seq_Num
          From 影像报告动作
         Where 原型ID = r_Aid
           And 事件ID Is Null;
      Else
        Select Nvl(Max(序号), 0) + 1
          Into n_Seq_Num
          From 影像报告动作
         Where 原型ID = r_Aid
           And 事件ID = 事件ID_In;
      End If;
    Else
      n_Seq_Num := 0;
    End If;
  
    If n_Seq_Num > 0 Then
      Update 影像报告动作
         Set 事件id       = 事件ID_In,
             动作类型     = 动作类型_In,
             名称         = 名称_In,
             说明         = 说明_In,
             可否手工执行 = 可否手工执行_In,
             内容         = 内容_In,
             序号         = n_Seq_Num
       Where ID = Id_In;
    Else
      Update 影像报告动作
         Set 事件id       = 事件ID_In,
             动作类型     = 动作类型_In,
             名称         = 名称_In,
             说明         = 说明_In,
             可否手工执行 = 可否手工执行_In,
             内容         = 内容_In
       Where ID = Id_In;
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Doc_Process;

  --49.获取元素或者提纲的名称集合
  Procedure p_Get_Antetype_Ele_Section(
	Val           Out t_Refcur,
	原型ID_In 影像报告原型清单.Id%Type,
	Type_In   Varchar2
	) As
    c_Content Clob;
  Begin
    /*Select To_Clob(a.内容)*/
    Select a.内容.getclobval()
      Into c_Content
      From 影像报告原型清单 A
     Where a.Id = 原型ID_In;
  
    If Type_In = '1' Then
      Open Val For
        Select Distinct Name
          From (Select Extractvalue(c.Column_Value, '/*/@title') As Name
                  From Table(Xmlsequence(Extract(Xmltype(c_Content),
                                                 '/zlxml/document//element[@sid and @title]|/zlxml/document//e_list[@sid and @title]|/zlxml/document//e_enum[@sid and @title]|/zlxml/document//e_etree[@sid and @title]|/zlxml/document//e_utree[@sid and @title]'))) C) A
         Where a.Name Is Not Null;
    Else
      Open Val For
        Select Distinct Name
          From (Select Extractvalue(c.Column_Value, '/section/@title') As Name
                  From Table(Xmlsequence(Extract(Xmltype(c_Content),
                                                 '//section'))) C) A
         Where a.Name Is Not Null;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_Ele_Section;

  --50.删除文档处理
  Procedure p_Del_Doc_Process(Id_In        影像报告动作.ID%Type,
                              Del_Event_In Number) As
    r_Event_Id   影像报告动作.事件ID%Type := Null;
    n_Event_Kind 影像报告事件.种类%Type;
    n_Is_Exist   Number(1) := 0;
  Begin
    If Del_Event_In = 1 Then
      Select Max(e.Id), Max(e.种类)
        Into r_Event_Id, n_Event_Kind
        From 影像报告动作 P, 影像报告事件 E
       Where p.Id = Id_In
         And p.事件id = e.Id;
    End If;
  
    Delete From 影像报告动作 Where ID = Id_In;
  
    If Del_Event_In = 1 Then
      If (r_Event_Id Is Not Null And n_Event_Kind <> 1) Then
        Select Count(*)
          Into n_Is_Exist
          From 影像报告动作
         Where 事件id = r_Event_Id;
        If n_Is_Exist = 0 Then
          Delete From 影像报告事件
           Where ID = r_Event_Id
             And 种类 <> 1;
        End If;
      End If;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Doc_Process;

  --51.查询元素值域类别的覆盖情况
  Procedure p_Get_Ele_Same_Info(
	Val           Out t_Refcur,
	Id_In    影像报告值域清单.Id%Type,
	Code_In  Varchar2,
	Title_In Varchar2,
	Flag_In  Varchar2
	) As
    v_Result  Varchar2(50);
    v_Id      Varchar2(50);
    v_Code_Id Varchar2(50);
    n_Num     Number;
  Begin
    If Flag_In = 1 Then
      Select Count(ID)
        Into n_Num
        From 影像报告元素分类 A
       Where a.Id = Id_In;
    
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From 影像报告元素分类 A
         Where a.Id = Id_In;
        If v_Id Is Not Null Then
          v_Result := 'ID重复';
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From 影像报告元素分类 A
       Where a.编码 = Code_In;
    
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Code_Id
          From 影像报告元素分类 A
         Where a.编码 = Code_In;
        If v_Code_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',编码重复';
          Else
            v_Result := '编码重复';
          End If;
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From 影像报告元素分类 A
       Where a.名称 = Title_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From 影像报告元素分类 A
         Where a.名称 = Title_In;
        If v_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',名称重复';
          Else
            v_Result := '名称重复';
          End If;
        End If;
      End If;
    
    End If;
  
    If Flag_In = 2 Then
      Select Count(ID)
        Into n_Num
        From 影像报告元素清单 A
       Where a.Id = Id_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From 影像报告元素清单 A
         Where a.Id = Id_In;
        If v_Id Is Not Null Then
          v_Result := 'ID重复';
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From 影像报告元素清单 A
       Where a.编码 = Code_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Code_Id
          From 影像报告元素清单 A
         Where a.编码 = Code_In;
        If v_Code_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',编码重复';
          Else
            v_Result := '编码重复';
          End If;
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From 影像报告元素清单 A
       Where a.名称 = Title_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From 影像报告元素清单 A
         Where a.名称 = Title_In;
        If v_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',名称重复';
          Else
            v_Result := '名称重复';
          End If;
        End If;
      End If;
    
    End If;
  
    If Flag_In = 3 Then
      Select Count(ID)
        Into n_Num
        From 影像报告值域清单 A
       Where a.Id = Id_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From 影像报告值域清单 A
         Where a.Id = Id_In;
        If v_Id Is Not Null Then
          v_Result := 'ID重复';
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From 影像报告值域清单 A
       Where a.编码 = Code_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Code_Id
          From 影像报告值域清单 A
         Where a.编码 = Code_In;
        If v_Code_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',编码重复';
          Else
            v_Result := '编码重复';
          End If;
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From 影像报告值域清单 A
       Where a.名称 = Title_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From 影像报告值域清单 A
         Where a.名称 = Title_In;
        If v_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',名称重复';
          Else
            v_Result := '名称重复';
          End If;
        End If;
      End If;
    
    End If;
  
    Open Val For
      Select v_Result As Result, v_Id As ID, v_Code_Id As Codesameid
        From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Ele_Same_Info;

  --52.获得所有的插件信息
  Procedure p_Get_DocPluginList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(Id) ID,
             编码,
             名称,
             说明,
             显示样式,
             种类,
             Decode(显示样式, '1', '嵌入式', '弹出式') 显示样式II,
             Decode(种类, '1', '专用插件', '共享插件') 种类II,
             类名,
             库名,
             是否禁用,
             Decode(是否禁用, '1', '停用', '启用') IsEnable
        From 影像报告插件;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_DocPluginList;

  --53.该ID的插件是否被原型使用过
  Procedure p_IsExit_DocPluginByID(
    Val           Out t_Refcur,
	ID_In Varchar2
	) As
    CURSOR C_EVENT Is
      Select t.专用插件.getclobval() 专用插件 From 影像报告原型清单 t;
    anum Int := 0;
    sult Varchar2(6666);
  Begin
    For temp In C_EVENT Loop
      If instr(temp.专用插件, ID_In) > 0 Then
        anum := anum + 1;
      End If;
    End Loop;
    Open Val For
      Select anum From dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_IsExit_DocPluginByID;

  --54.新增报告插件信息
  Procedure p_AddDocPlugin(
    ID_In       影像报告插件.ID%Type,
    编码_In     影像报告插件.编码%Type,
    名称_In     影像报告插件.名称%Type,
    说明_In     影像报告插件.说明%Type,
    显示样式_In 影像报告插件.显示样式%Type,
    种类_In     影像报告插件.种类%Type,
    类名_In     影像报告插件.类名%Type,
    库名_In     影像报告插件.库名%Type,
    是否禁用_In 影像报告插件.是否禁用%Type
	) As
  Begin
    Insert Into 影像报告插件
      (ID, 编码, 名称, 说明, 显示样式, 种类, 类名, 库名, 是否禁用)
    Values
      (ID_In,
       编码_In,
       名称_In,
       说明_In,
       显示样式_In,
       种类_In,
       类名_In,
       库名_In,
       是否禁用_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddDocPlugin;

  --55.修改报告插件信息
  Procedure p_EditDocPlugin(
    ID_In       影像报告插件.ID%Type,
    编码_In     影像报告插件.编码%Type,
    名称_In     影像报告插件.名称%Type,
    说明_In     影像报告插件.说明%Type,
    显示样式_In 影像报告插件.显示样式%Type,
    种类_In     影像报告插件.种类%Type,
    类名_In     影像报告插件.类名%Type,
    库名_In     影像报告插件.库名%Type,
    是否禁用_In 影像报告插件.是否禁用%Type
	) As
  Begin
    Update 影像报告插件
       Set 编码     = 编码_In,
           名称     = 名称_In,
           说明     = 说明_In,
           显示样式 = 显示样式_In,
           种类     = 种类_In,
           类名     = 类名_In,
           库名     = 库名_In,
           是否禁用 = 是否禁用_In
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditDocPlugin;

  --56.删除报告插件信息
  Procedure p_DelDocPlugin(
    ID_In 影像报告插件.ID%Type
	) As
  Begin
    Delete From 影像报告插件 Where id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelDocPlugin;

  --57.改变插件的可用状态
  Procedure p_IsEnableDocPlugin(
    ID_In 影像报告插件.ID%Type
	) As
  Begin
    Update 影像报告插件 a
       Set 是否禁用 = Decode(a.是否禁用, 1, 0, 1)
     Where id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_IsEnableDocPlugin;

  --58.通过ID获得对应的插件信息
  Procedure p_GetDocPluginByID(
	Val           Out t_Refcur,
	ID_In 影像报告插件.ID%type
	) As
  Begin
    Open Val For
      Select Rawtohex(Id) ID,
             编码,
             名称,
             说明,
             显示样式,
             种类,
             类名,
             库名,
             是否禁用
        From 影像报告插件
       Where id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDocPluginByID;

  --59.判断编码和名称是否已存在
  Procedure p_IsExitDocPlugin(
	Val           Out t_Refcur,
	ID_In   影像报告插件.ID%Type,
	编码_In 影像报告插件.编码%Type,
	名称_In 影像报告插件.名称%Type
	) As
  Begin
    Open Val For
      Select Count(id)
        From 影像报告插件 a
       Where (a.编码 = 编码_In Or a.名称 = 名称_In)
         and a.id <> ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_IsExitDocPlugin;

  --60.通过ID获得对应的专用插件信息
  Procedure p_GetDocSpecPluginByID(
	Val           Out t_Refcur,
	ID_In 影像报告插件.ID%Type
	) As
  Begin
    Open Val For
      Select Rawtohex(Id) ID,
             编码,
             名称,
             说明,
             显示样式,
             种类,
             类名,
             库名,
             是否禁用
        From 影像报告插件
       Where id = ID_In
         And 种类 = 1;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDocSpecPluginByID;

  --61.获得诊疗列表信息
  Procedure p_GetDiagnosisList(
	Val           Out t_Refcur,
	类别_In Varchar2,
	条件_In Varchar2
	) As
  Begin
    Open Val For
      Select to_char(a.id) ID,
             a.编码,
             a.名称,
             (Select b.名称 From 诊疗项目类别 b Where b.编码 = a.类别) 类别
        From 诊疗项目目录 a
       Where (a.id In (Select t.诊疗项目id From 影像检查项目 t) And a.类别 = 类别_In)
         And (a.编码 Like 条件_In || '%' Or a.名称 Like 条件_In || '%');
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDiagnosisList;

  --62.获得诊疗类别列表
  Procedure p_GetDiagnosisClass(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select t.编码, t.名称, t.简码 From 诊疗项目类别 t;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDiagnosisClass;

  --63.添加影像报告原型应用信息
  Procedure p_AddMedicalAntetype(
    诊疗项目ID_In 影像报告原型应用.诊疗项目ID%Type,
    应用场合_In   影像报告原型应用.应用场合%Type,
    报告原型ID_In 影像报告原型应用.报告原型ID%Type
	) As
  Begin
    Insert Into 影像报告原型应用
      (诊疗项目ID, 应用场合, 报告原型ID)
    Values
      (诊疗项目ID_In, 应用场合_In, 报告原型ID_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddMedicalAntetype;

  --64.删除原型ID对应的病历单据应用信息
  Procedure p_DelMedicalAntetype(
    报告原型ID_In 影像报告原型应用.报告原型ID%Type
	) As
  Begin
    Delete From 影像报告原型应用 Where 报告原型ID = 报告原型ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelMedicalAntetype;

  --65.通过原型ID获得对应的病历单据应用信息
  Procedure p_GetMedicalByAID(
	Val           Out t_Refcur,
	报告原型ID_In 影像报告原型应用.报告原型ID%Type
	) As
  Begin
    Open Val For
      Select id,
             x.编码,
             x.名称,
             x.类别,
             Sum(x.门诊) 门诊,
             Sum(x.住院) 住院,
             Sum(x.外诊) 外诊,
             Sum(x.体检) 体检
        From (Select id,
                     编码,
                     名称,
                     类别,
                     Decode(应用场合, '1', 1, 0) as 门诊,
                     Decode(应用场合, '2', 1, 0) as 住院,
                     Decode(应用场合, '3', 1, 0) as 外诊,
                     Decode(应用场合, '4', 1, 0) as 体检
                From (Select to_Char(a.诊疗项目id) ID,
                             (Select b.编码
                                From 诊疗项目目录 b
                               Where b.id = a.诊疗项目id) as 编码,
                             (Select b.名称
                                From 诊疗项目目录 b
                               Where b.id = a.诊疗项目id) as 名称,
                             (Select c.名称
                                From 诊疗项目类别 c
                               Where c.编码 = (Select b.类别
                                               From 诊疗项目目录 b
                                              Where b.id = a.诊疗项目id)) As 类别,
                             a.应用场合
                        From 影像报告原型应用 a
                       Where a.报告原型id = 报告原型ID_In)) x
       Group By x.id, x.编码, x.名称, x.类别;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetMedicalByAID;

  --66.根据原型ID删除动作信息
  Procedure p_DelDocProcessByAid(
    报告原型ID_In 影像报告原型应用.报告原型ID%Type
	) As
  Begin
    Delete From 影像报告动作 t Where t.原型id = 报告原型ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;
  --67.获取ID对应的原型的树形结构
  Procedure p_GetAntetypeTreeByID(
	Val           Out t_Refcur,
	ID_In 影像报告原型清单.ID%Type
	) As
  Begin
    Open Val For
      Select ID, 编码, 名称, 标题, 分组, 是否禁用, 说明, Imageindex
        From (Select Distinct 分组 As ID,
                              (Select Min(b.编码)
                                 From 影像报告原型清单 B
                                Where b.分组 = a.分组) As 编码,
                              a.分组 As 名称,
                              a.分组 As 标题,
                              null As 分组,
                              0 As 是否禁用,
                              null As 说明,
                              0 As Imageindex
                From 影像报告原型清单 A
               Where a.id = ID_In
              Union
              Select RawtoHex(ID) ID,
                     a.编码,
                     a.名称 As 名称,
                     编码 || '-' || 名称 As 标题,
                     分组,
                     a.是否禁用,
                     a.说明,
                     Decode(a.是否禁用, 1, 2, 1) Imageindex
                From 影像报告原型清单 A
               Where a.id = ID_In) A
       Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetAntetypeTreeByID;

  --68.原型是否存在对应的编码或名称
  procedure p_IsExitAntetype(
	Val           Out t_Refcur,
	编码_In 影像报告原型清单.编码%Type,
	名称_In 影像报告原型清单.名称%Type,
	ID_In  影像报告原型清单.ID%Type
	) As
  begin
    Open Val For
      Select Count(*) AS num
        From 影像报告原型清单 t
       where (t.编码 = 编码_In
          or t.名称 = 名称_In) and t.id<>ID_In;
  End p_IsExitAntetype;

  --69. 获取影像存储设备
  Procedure p_GetStorageDevice(
	Val           Out t_Refcur
	) Is 
  Begin 
	Open Val For
		Select 设备号||' - '||设备名 As 存储设备, 设备号, IP地址, FTP目录, FTP用户名, FTP密码, 共享目录用户名, 共享目录密码, 共享目录  
		From 影像设备目录 Where 类型 = 1;
	Exception
	  When Others Then
	  Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetStorageDevice;
End b_PACS_RptAntetype;
/





CREATE OR REPLACE Package b_Pacs_RptSampleList Is
  Type t_Refcur Is Ref Cursor;

  -- Author  : SEEKING
  -- Created : 2014/10/30 10:05:38
  -- Purpose : 范文管理

  --查找范文的原型类别

  Procedure p_Get_Sample_List_Type(
    Val Out t_Refcur,
    Type_In Varchar2,
    Kind_In varchar2
	);

  --新增文档原型
  Procedure p_Add_Sample_List(
    Id_In       影像报告范文清单.Id%Type,
	Aid_In      影像报告范文清单.原型id%Type,
	Seq_Num_In  影像报告范文清单.编号%Type,
	Title_In    影像报告范文清单.名称%Type,
	Note_In     影像报告范文清单.说明%Type,
	Content_In  影像报告范文清单.内容%Type,
	Subject_In  影像报告范文清单.学科%Type,
	Label_In    影像报告范文清单.标签%Type,
	Private_In  影像报告范文清单.是否私有%Type,
	Author_In   影像报告范文清单.作者%Type,
	Lasttime_In 影像报告范文清单.最后编辑时间%Type
	);

  --编辑范文信息
  Procedure p_Edit_Sample_List(
    Id_In       影像报告范文清单.Id%Type,
    Aid_In      影像报告范文清单.原型id%Type,
    Seq_Num_In  影像报告范文清单.编号%Type,
    Title_In    影像报告范文清单.名称%Type,
    Note_In     影像报告范文清单.说明%Type,
    Content_In  影像报告范文清单.内容%Type,
    Subject_In  影像报告范文清单.学科%Type,
    Label_In    影像报告范文清单.标签%Type,
    Private_In  影像报告范文清单.是否私有%Type,
    Author_In   影像报告范文清单.作者%Type,
    Lasttime_In 影像报告范文清单.最后编辑时间%Type
	);
  --删除文档范文
  Procedure p_Del_Sample_List(
    Id_In 影像报告范文清单.Id%Type
	);

  --通过原型ID获得相应的范文信息
  Procedure p_Get_Samplelist_By_Aid(
    Val Out t_Refcur,
    Antetypelist_Id_In 影像报告范文清单.原型id%Type,
    Author_In          影像报告范文清单.作者%Type,
    Subjects_In        Varchar2
	);

  --通过种类id获取范文树
  Procedure p_Get_Samplelist_By_Kind(
    Val Out t_Refcur,
    Kind_In      Varchar2,
    Condition_In Varchar2,
    Author_In    影像报告范文清单.作者%Type,
    Subjects_In  Varchar2
	);

  --通过ID查找相应的范文信息
  Procedure p_Get_Samplelist_By_Id(
    Val Out t_Refcur,                                   
    Id_In 影像报告范文清单.Id%Type
	);

  --查询相应原型下的最大序号
  Procedure p_Get_Samplelist_Maxseqnum(
    Val Out t_Refcur,
	Aid_In 影像报告范文清单.原型id%Type
	);

  --获取范文XML信息
  Procedure p_Get_Samplexml(
    Val Out t_Refcur,     
	Id_In 影像报告范文清单.Id%Type
	);

  --修改范文XML信息
  Procedure p_Edit_Samplexml(
    Id_In      影像报告范文清单.Id%Type,
	Content_In 影像报告范文清单.内容%Type
	);

  --导出的范文列表
  Procedure p_Output_Samplelist(
    Val Out t_Refcur
	);

  --是否存在相应的原型类别
  Procedure p_If_Exist_Antetypelist(
    Val Out t_Refcur,
	Title_In 影像报告原型清单.名称%Type
	);

  --同一个类别下是否存在相同名称的范文
  Procedure p_If_Exist_Samplelist(
    Val Out t_Refcur,
    Type_In  影像报告原型清单.名称%Type,
    Title_In 影像报告范文清单.名称%Type
	);

  --通过范文ID获得范文对应的树形结构
  Procedure p_Get_SamplelistTree_By_Id(
    Val Out t_Refcur,
    Id_In 影像报告范文清单.Id%Type
	);

End b_Pacs_RptSampleList;
/

CREATE OR REPLACE Package Body b_Pacs_RptSampleList Is

  ------------------------------------------------------------------------
  --范文管理
  ------------------------------------------------------------------------

  --查找范文的原型类别

  Procedure p_Get_Sample_List_Type(
    Val Out t_Refcur,
    Type_In Varchar2,
    Kind_In Varchar2
	) As
  Begin
    If Type_In = '1' Then
      Open Val For
        Select Rawtohex(a.Id) ID, a.编码 || '-' || a.名称 名称
          From 影像报告原型清单 A
         Where a.Id In (Select Distinct b.原型id From 影像报告范文清单 B)
           and a.种类 = Kind_In;
    Else
      Open Val For
        Select Rawtohex(a.Id) ID, a.编码 || '-' || a.名称 名称
          From 影像报告原型清单 A
         where a.种类 = Kind_In
         Order By a.编码;
    End If;
  
  End p_Get_Sample_List_Type;
  --新增文档原型
  Procedure p_Add_Sample_List(
    Id_In       影像报告范文清单.Id%Type,
	Aid_In      影像报告范文清单.原型id%Type,
	Seq_Num_In  影像报告范文清单.编号%Type,
	Title_In    影像报告范文清单.名称%Type,
	Note_In     影像报告范文清单.说明%Type,
	Content_In  影像报告范文清单.内容%Type,
	Subject_In  影像报告范文清单.学科%Type,
	Label_In    影像报告范文清单.标签%Type,
	Private_In  影像报告范文清单.是否私有%Type,
	Author_In   影像报告范文清单.作者%Type,
	Lasttime_In 影像报告范文清单.最后编辑时间%Type
	) As
    n_Num Number;
    v_Msg Varchar2(200);
    Err Exception;
  Begin
    Select Count(a.Id)
      Into n_Num
      From 影像报告范文清单 A
     Where a.原型id = Hextoraw(Aid_In)
       And a.名称 = Title_In;
  
    If n_Num > 0 Then
      v_Msg := '[ZLSOFT]在同一个原型下的范文名称不能相同！[ZLSOFT]';
      Raise Err;
    End If;
  
    Insert Into 影像报告范文清单
      (ID,
       原型id,
       编号,
       名称,
       说明,
       内容,
       学科,
       标签,
       是否私有,
       作者,
       最后编辑时间)
    Values
      (Hextoraw(Id_In),
       Hextoraw(Aid_In),
       Seq_Num_In,
       Title_In,
       Note_In,
       Content_In,
       Subject_In,
       Label_In,
       Private_In,
       Author_In,
       Sysdate);
  
    --这里添加对于该范文的处理,页眉页脚，页面设置
    Update 影像报告范文清单
       Set 内容 = Updatexml(内容,
                          '/zlxml/subdocuments',
                          (Select Extract(内容, 'zlxml/subdocuments')
                             From 影像报告原型清单
                            Where ID = Hextoraw(Aid_In)))
     Where ID = Hextoraw(Id_In); --处理页眉页脚
  
    Update 影像报告范文清单
       Set 内容 = Updatexml(内容,
                          '/zlxml/docparameters',
                          (Select Extract(内容, 'zlxml/docparameters')
                             From 影像报告原型清单
                            Where ID = Hextoraw(Aid_In)))
     Where ID = Hextoraw(Id_In); --处理页面设置
  
  Exception
    When Err Then
      Raise_Application_Error(-20101, v_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Add_Sample_List;

  --编辑范文信息
  Procedure p_Edit_Sample_List(
    Id_In       影像报告范文清单.Id%Type,
    Aid_In      影像报告范文清单.原型id%Type,
    Seq_Num_In  影像报告范文清单.编号%Type,
    Title_In    影像报告范文清单.名称%Type,
    Note_In     影像报告范文清单.说明%Type,
    Content_In  影像报告范文清单.内容%Type,
    Subject_In  影像报告范文清单.学科%Type,
    Label_In    影像报告范文清单.标签%Type,
    Private_In  影像报告范文清单.是否私有%Type,
    Author_In   影像报告范文清单.作者%Type,
    Lasttime_In 影像报告范文清单.最后编辑时间%Type
	) As
    n_Num Number;
    v_Msg Varchar2(200);
    Err Exception;
  Begin
    Select Count(a.Id)
      Into n_Num
      From 影像报告范文清单 A
     Where a.原型id = Hextoraw(Aid_In)
       And a.名称 = Title_In
       And a.Id <> Hextoraw(Id_In);
  
    If n_Num > 0 Then
      v_Msg := '[ZLSOFT]在同一个原型下的范文名称不能相同！[ZLSOFT]';
      Raise Err;
    End If;
    Update 影像报告范文清单
       Set 原型id       = Hextoraw(Aid_In),
           编号         = Decode(Seq_Num_In, 0, 编号, Seq_Num_In),
           名称         = Title_In,
           说明         = Note_In,
           内容         = Content_In,
           学科         = Subject_In,
           标签         = Label_In,
           是否私有     = Private_In,
           作者         = Author_In,
           最后编辑时间 = Sysdate
     Where ID = Hextoraw(Id_In);
  
    --这里添加对于该范文的处理,页眉页脚，页面设置
    Update 影像报告范文清单
       Set 内容 = Updatexml(内容,
                          '/zlxml/subdocuments',
                          (Select Extract(内容, 'zlxml/subdocuments')
                             From 影像报告原型清单
                            Where ID = Hextoraw(Aid_In)))
     Where ID = Hextoraw(Id_In); --处理页眉页脚
  
    Update 影像报告范文清单
       Set 内容 = Updatexml(内容,
                          '/zlxml/docparameters',
                          (Select Extract(内容, 'zlxml/docparameters')
                             From 影像报告原型清单
                            Where ID = Hextoraw(Aid_In)))
     Where ID = Hextoraw(Id_In); --处理页面设置
  
  Exception
    When Err Then
      Raise_Application_Error(-20101, v_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_Sample_List;

  --删除文档范文
  Procedure p_Del_Sample_List(
    Id_In 影像报告范文清单.Id%Type
	) As
  Begin
    Delete From 影像报告范文清单
     Where 影像报告范文清单.Id = Hextoraw(Id_In);
  End p_Del_Sample_List;

  --通过原型ID获取该原型下的范文列表
  Procedure p_Get_Samplelist_By_Aid(
    Val Out t_Refcur,
    Antetypelist_Id_In 影像报告范文清单.原型id%Type,
    Author_In          影像报告范文清单.作者%Type,
    Subjects_In        Varchar2
	) As
  Begin
    Open Val For
      Select /*+rule*/
       Rawtohex(a.Id) As ID,
       a.名称,
       a.作者,
       a.说明,
       a.学科,
       a.编号 Seqnum,
       a.标签,
       a.是否私有
        From 影像报告范文清单 A
       Where a.原型id = Hextoraw(Antetypelist_Id_In)
         And (a.作者 = Author_In Or (a.学科 Is Null And a.是否私有 = 0) Or
             Subjects_In Is Null Or
             (a.学科 Is Not Null And
             b_pacs_rptpublic.f_If_Intersect(a.学科, Subjects_In) > 0 And
             a.是否私有 = 0));
  End p_Get_Samplelist_By_Aid;

  --通过种类id获取范文树
  Procedure p_Get_Samplelist_By_Kind(
    Val Out t_Refcur,
    Kind_In      Varchar2,
    Condition_In Varchar2,
    Author_In    影像报告范文清单.作者%Type,
    Subjects_In  Varchar2
	) As
  Begin
  
    --获得一个存在原型信息的范文树形结构
    Open Val For
      Select a.分组 As ID,
             a.分组 As 名称,
             '' 说明,
             '' As 上级id,
             ' category' As Type,
             '' As 作者,
             '' As 学科,
             Null 修改时间,
             '' As 标签,
             0 As Private,
             0 As Imgindex
        From 影像报告原型清单 A
       Where a.种类 = Kind_In
         And Exists
       (Select ID From 影像报告范文清单 C Where c.原型id = a.Id)
         And a.分组 Is Not Null
      Union
      Select m.*
        From (Select Rawtohex(b.Id) As ID,
                     b.名称,
                     b.说明,
                     b.分组 上级id,
                     'antetype' As Type,
                     '' As 作者,
                     '' As 学科,
                     Null 修改时间,
                     '' As 标签l,
                     0 As Private,
                     0 As Imgindex
                From 影像报告原型清单 B
               Where b.种类 = Kind_In
                 And Exists
               (Select ID From 影像报告范文清单 C Where c.原型id = b.Id)
               Order By b.编码) M
      Union All
      Select n.*
        From (Select /*+rule*/
               Rawtohex(a.Id) As ID,
               a.名称,
               a.说明,
               Rawtohex(a.原型id) As 上级id,
               'sample' As Type,
               a.作者,
               a.学科,
               a.最后编辑时间 As 修改时间,
               a.标签,
               a.是否私有 As Private,
               Decode(a.是否私有, 1, 2, 1) As Imgindex
                From 影像报告范文清单 A, 影像报告原型清单 C
               Where a.原型id = c.Id
                 And c.种类 = Kind_In
                 And ((a.名称 Like '%' || Condition_In || '%' And
                     Condition_In Is Not Null) Or Condition_In Is Null)
                 And (a.作者 = Author_In Or (a.学科 Is Null And a.是否私有 = 0) Or
                     Subjects_In Is Null Or
                     (a.学科 Is Not Null And
                     b_pacs_rptpublic.f_If_Intersect(a.学科, Subjects_In) > 0 And
                     a.是否私有 = 0))
               Order By a.编号, a.名称) N;
  
  End p_Get_Samplelist_By_Kind;

  --通过ID查找相应的范文信息
  Procedure p_Get_Samplelist_By_Id(
    Val Out t_Refcur,
	Id_In 影像报告范文清单.Id%Type
	) As
  Begin
    Open Val For
      Select Rawtohex(a.Id) As ID,
             Rawtohex(a.原型id) As 原型id,
             a.编号,
             a.名称,
             a.说明,
             a.学科,
             a.标签,
             a.是否私有,
             a.作者,
             a.最后编辑时间 Lasttime
        From 影像报告范文清单 A
       Where a.Id = Hextoraw(Id_In);
  
  End p_Get_Samplelist_By_Id;

  --查询相应原型下的最大序号
  Procedure p_Get_Samplelist_Maxseqnum(
    Val Out t_Refcur,
    Aid_In 影像报告范文清单.原型id%Type
	) As
  Begin
    Open Val For
      Select Nvl(Max(a.编号), 0) + 1 As Num
        From 影像报告范文清单 A
       Where a.原型id = Hextoraw(Aid_In);
  End p_Get_Samplelist_Maxseqnum;

  --获取范文XML信息
  Procedure p_Get_Samplexml(
    Val Out t_Refcur,
	Id_In 影像报告范文清单.Id%Type
	) As
  Begin
    Open Val For
      Select A.内容.getclobval() 内容 From 影像报告范文清单 A Where a.Id = Id_In;
  End p_Get_Samplexml;

  --修改范文XML信息
  Procedure p_Edit_Samplexml(
    Id_In      影像报告范文清单.Id%Type,
	Content_In 影像报告范文清单.内容%Type
	) As
  Begin
    Update 影像报告范文清单
       Set 内容 = Content_In
     Where ID = Hextoraw(Id_In);
  End p_Edit_Samplexml;

  --导出的范文列表
  Procedure p_Output_Samplelist(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select *
        From (Select Rawtohex(a.Id) As ID,
                     Null As Aid,
                     a.名称 As Antetypename,
                     Null As Code,
                     a.名称 As Title,
                     a.说明 As Note,
                     Null As Content,
                     Null As Subject,
                     Null As Label,
                     Null As Private,
                     Null As Author,
                     Null As Lasttime,
                     '' As Flag,
                     1 As Image,
                     'antetype' As Type
                From 影像报告原型清单 A
               Where Exists (Select b.Id
                        From 影像报告范文清单 B
                       Where b.原型id = a.Id)
               Order By a.编码)
      Union All
      Select Rawtohex(c.Id) As ID,
             Rawtohex(c.原型id) As Aid,
             d.名称,
             c.编号,
             c.名称,
             c.说明,
             '',
             c.学科,
             c.标签,
             c.是否私有,
             c.作者,
             c.最后编辑时间,
             '',
             0,
             'sample'
        From 影像报告范文清单 C, 影像报告原型清单 D
       Where c.原型id = d.Id;
  
  End p_Output_Samplelist;

  --是否存在相应的原型类别
  Procedure p_If_Exist_Antetypelist(
    Val Out t_Refcur,
	Title_In 影像报告原型清单.名称%Type
	) As
  Begin
    Open Val For
      Select Count(a.Id) Num
        From 影像报告原型清单 A
       Where a.名称 = Title_In;
  End p_If_Exist_Antetypelist;

  --同一个类别下是否存在相同名称的范文
  Procedure p_If_Exist_Samplelist(
    Val Out t_Refcur,
    Type_In  影像报告原型清单.名称%Type,
    Title_In 影像报告范文清单.名称%Type
	) As
  Begin
    Open Val For
      Select Count(a.Id) Num, Max(a.Id) ID
        From 影像报告范文清单 A, 影像报告原型清单 B
       Where a.原型id = b.Id
         And a.名称 = Title_In
         And b.名称 = Type_In;
  End p_If_Exist_Samplelist;
  
  --通过范文ID获得范文对应的树形结构
  Procedure p_Get_SamplelistTree_By_Id(
    Val Out t_Refcur,
    Id_In 影像报告范文清单.Id%Type
	) As
  Begin
    --'EE7CD4A510B045A9BBE6D8CC7DB6EE30'
    Open Val For
      Select RawToHex(t.ID) as ID,
             t.名称,
             T.说明,
             RawToHex(T.原型ID) 原型ID,
             'sample' 类型,
             t.作者,
             t.学科,
             t.最后编辑时间,
             t.标签,
             t.是否私有,
             2 IMGINDEX
        From 影像报告范文清单 t
       Where t.id = Id_In
      Union All
      Select RawToHex(x.id) as ID,
             x.名称,
             x.说明,
             x.分组 原型ID,
             'antetype' 类型,
             null 作者,
             null 学科,
             null 最后编辑时间,
             null 标签,
             0 是否私有,
             0 IMGINDEX
        From 影像报告原型清单 x
       Where x.id = (Select t.原型id
                       From 影像报告范文清单 t
                      where t.id = Id_In
                        and rownum <= 1)
      Union All
      Select x.分组 as ID,
             x.分组 名称,
             null 说明,
             null 原型ID,
             'category' 类型,
             null 作者,
             null 学科,
             null 最后编辑时间,
             null 标签,
             0 是否私有,
             0 IMGINDEX
        From 影像报告原型清单 x
       Where x.ID = (Select t.原型id
                       From 影像报告范文清单 t
                      where t.id = Id_In
                        and rownum <= 1);
  End p_Get_SamplelistTree_By_Id;
End b_Pacs_RptSampleList;
/



--影像报告业务(---定义部分---)***************************************************
Create Or Replace Package b_Pacs_Rptmanage Is
  Type t_Refcur Is Ref Cursor;

  --1、锁定报告人
  Procedure p_Edit_Doc_Lockinfo
  (
    报告_Id_In 影像报告记录.Id%Type,
    锁定人_In  影像报告记录.锁定人%Type
  );

  --2、评定报告质量
  Procedure p_Edit_Doc_Evaluatrptquality
  (
    报告id_In   影像报告记录.Id%Type,
    质量等级_In 影像报告记录.报告质量%Type
  );

  --3、评定阴阳性
  Procedure p_Edit_Doc_Evaluatresult
  (
    报告id_In   影像报告记录.Id%Type,
    检查结果_In 影像报告记录.结果阳性%Type
  );

  --4、报告发放/回收
  Procedure p_Edit_Doc_Reportrelease
  (
    报告id_In     影像报告记录.Id%Type,
    当前操作人_In 影像报告记录.报告发放人%Type
  );

  --5、新增，修改报告
  Procedure p_影像报告记录_新增
  (
    原型id_In     影像报告记录.原型id%Type,
    报告内容_In   影像报告记录.报告内容%Type,
    记录人_In     影像报告记录.记录人%Type,
    最后编辑人_In 影像报告记录.最后编辑人%Type,
    Id_In         影像报告记录.Id%Type,
    医嘱id_In     影像报告记录.医嘱id%Type
  );

  --6、获取书写的文档内容
  Procedure p_Get_Doc_Content
  (
    Val      Out t_Refcur,
    Docid_In 影像报告记录.Id%Type
  );

  --7、设置报告打印作废信息
  Procedure p_Checkrejectsignature
  (
    Signdate_In Date,
    报告id_In   影像报告操作记录.报告id%Type,
    作废人_In   影像报告操作记录.作废人%Type,
    作废说明_In 影像报告操作记录.作废说明%Type,
    Val         Out Sys_Refcursor
  );

  --8、查询相应原型下的最大序号
  Procedure p_Get_Samplelist_Maxseqnum
  (
    Val       Out t_Refcur,
    原型id_In 影像报告范文清单.原型id%Type
  );

  --9、删除文档范文
  Procedure p_Del_影像报告范文清单(Id_In 影像报告范文清单.Id%Type);

  --10、添加文档的操作日志
  Procedure p_影像报告操作记录_Add
  (
    Id_In       影像报告操作记录.Id%Type,
    报告id_In   影像报告操作记录.报告id%Type,
    操作人_In   影像报告操作记录.操作人%Type,
    操作类型_In 影像报告操作记录.操作类型%Type
  );

  --11、删除报告
  Procedure p_影像报告记录_删除(报告_Id_In 影像报告记录.Id%Type);

  --12、获取签名类型
  Procedure p_Get_Sysconfigsignature
  (
    Val       Out t_Refcur,
    科室id_In In 部门表.Id%Type
  );

  --13、获取账户签名印章
  Procedure p_Get_Personsignimg
  (
    Val   Out t_Refcur,
    Id_In In 人员表.Id%Type
  );

  --14、获取签名的证书信息
  Procedure p_Get_Signcertinfo
  (
    Val       Out t_Refcur,
    证书id_In 人员证书记录.Id%Type
  );

  --15、更新报告状态
  Procedure p_Update_Reportstate
  (
    报告id_In   影像报告记录.Id%Type,
    报告状态_In 影像报告记录.报告状态%Type,
    审核人_In   影像报告记录.最后审核人%Type
  );

  --16、获取报告状态
  Procedure p_Get_Reportstate
  (
    Val       Out t_Refcur,
    报告id_In 影像报告记录.Id%Type
  );

  --17、报告驳回
  Procedure p_Reject_Report
  (
    医嘱id_In   影像报告驳回.医嘱id%Type,
    报告id_In   影像报告驳回.检查报告id%Type,
    驳回理由_In 影像报告驳回.驳回理由%Type,
    驳回时间_In 影像报告驳回.驳回时间%Type,
    驳回人_In   影像报告驳回.驳回人%Type,
    待处理人_In 影像报告记录.待处理人%Type,
    报告状态_In 影像报告记录.报告状态%Type
  );

  --17.1、撤销报告驳回
  Procedure p_Reject_Cancel
  (
    Id_In       影像报告驳回.Id%Type,
    报告id_In   影像报告驳回.检查报告id%Type,
    报告状态_In 影像报告记录.报告状态%Type
  );

  --18、获取报告驳回信息
  Procedure p_Get_Rejectinfo
  (
    Val       Out t_Refcur,
    报告id_In 影像报告驳回.检查报告id%Type
  );

  --19、获取原型动作
  Procedure p_Get_Doc_Process
  (
    Val       Out t_Refcur,
    原型id_In 影像报告动作.原型id%Type
  );

  --20、通过学科筛选获得相应的范文信息
  Procedure p_Get_Samplelist_By_Conditions
  (
    Val          Out t_Refcur,
    原型id_In    Varchar2,
    学科_In      Varchar2,
    Condition_In Varchar2, --过滤筛选
    作者_In      Varchar2
  );

  --21、通过部门ID获取部门名称
  Procedure p_Get_部门名称_By_Id
  (
    Val   Out t_Refcur,
    Id_In 部门表.Id%Type
  );

  --22、提取所有预备提纲
  Procedure p_Get_Allpreoutlines(Val Out t_Refcur);

  --23、提取文档标题
  Procedure p_Get_Reporttitle_By_Id
  (
    Val   Out t_Refcur,
    Id_In 影像报告记录.Id%Type
  );

  --24、提取报告锁定人
  Procedure p_Get_报告锁定人_By_Id
  (
    Val   Out t_Refcur,
    Id_In 影像报告记录.Id%Type
  );

  --25、通过医嘱ID获取报告列表
  Procedure p_Get_影像报告记录_By_医嘱id
  (
    Val       Out t_Refcur,
    医嘱id_In 影像报告记录.医嘱id%Type
  );

  --26、查询影像流程参数值
  Procedure p_Get_影像流程参数值
  (
    Val       Out t_Refcur,
    科室id_In 影像流程参数.科室id%Type
  );

  --27、根据医嘱ID，查询对应的原型列表
  Procedure p_Get_影像原型列表_By_医嘱id
  (
    Val     Out t_Refcur,
    医嘱_In 影像检查记录.医嘱id%Type
  );

  --28、根据报告ID查询打印记录
  Procedure p_Get_Reportprintlog_By_报告id
  (
    Val     Out Sys_Refcursor,
    报告_In 影像报告操作记录.报告id%Type
  );

  --29、根据医嘱ID查询报告发放列表
  Procedure p_Get_Reportreleaselist
  (
    Val     Out t_Refcur,
    医嘱_In 影像报告记录.医嘱id%Type
  );

  --30、根据报告ID查询驳回记录数量
  Procedure p_Get_Rejectedcount
  (
    Val     Out t_Refcur,
    报告_In 影像报告驳回.检查报告id%Type
  );

  --31、根据医嘱ID查询报告动作需要的一些ID们
  Procedure p_Get_Docprocess_Ids
  (
    Val     Out t_Refcur,
    医嘱_In 病人医嘱记录.Id%Type
  );

  --32、根据医嘱ID和报告ID查询报告的一些参数
  Procedure p_Get_Docinfo
  (
    Val       Out t_Refcur,
    医嘱id_In 影像检查记录.医嘱id%Type,
    报告id_In 影像报告记录.Id%Type
  );

  --33、查询一个检查中相同原型ID的报告数量
  Procedure p_Get_Sameantetypedoccounts
  (
    Val       Out t_Refcur,
    医嘱id_In 影像报告记录.医嘱id%Type,
    原型id_In 影像报告记录.原型id%Type
  );

  --34、提取报告图存储信息
  Procedure p_Get_Docimagesaveinof_By_Id
  (
    Val   Out t_Refcur,
    Id_In 影像报告记录.Id%Type
  );

  --35、修改原型使用次数
  Procedure p_Update_Antetypeusecount(Id_In 影像报告原型清单.Id%Type);

  --36、更新影像检查图像的报告图标记
  Procedure p_Update_Rptimage
  (
    Uid_In        影像检查图象.图像uid%Type,
    Actiontype_In Number
  );

  --37、提取打印控制信息
  Procedure p_Get_Printcontrol
  (
    Val       Out t_Refcur,
    报告id_In 影像报告记录.Id%Type
  );

End b_Pacs_Rptmanage;

/

--影像报告业务(---实现部分---)***************************************************

Create Or Replace Package Body b_Pacs_Rptmanage Is

  --1、锁定报告人
  Procedure p_Edit_Doc_Lockinfo
  (
    报告_Id_In 影像报告记录.Id%Type,
    锁定人_In  影像报告记录.锁定人%Type
  ) Is
  Begin
  
    --  报告ID为空，则清空所有“锁定人_In”正在锁定的标记
    If 报告_Id_In Is Null Then
      Update 影像报告记录 a Set a.锁定人 = '' Where a.锁定人 = 锁定人_In;
    Else
      Update 影像报告记录 a Set a.锁定人 = 锁定人_In Where a.Id = 报告_Id_In;
    End If;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_Lockinfo;

  --2、评定报告质量
  Procedure p_Edit_Doc_Evaluatrptquality
  (
    报告id_In   影像报告记录.Id%Type,
    质量等级_In 影像报告记录.报告质量%Type
  ) Is
  Begin
    Update 影像报告记录 Set 报告质量 = 质量等级_In Where Id = 报告id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_Evaluatrptquality;

  --3、评定阴阳性
  Procedure p_Edit_Doc_Evaluatresult
  (
    报告id_In   影像报告记录.Id%Type,
    检查结果_In 影像报告记录.结果阳性%Type
  ) Is
  Begin
    Update 影像报告记录 Set 结果阳性 = 检查结果_In Where Id = 报告id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_Evaluatresult;

  --4、报告发放/回收
  Procedure p_Edit_Doc_Reportrelease
  (
    报告id_In     影像报告记录.Id%Type,
    当前操作人_In 影像报告记录.报告发放人%Type
  ) Is
    v_报告发放 影像报告记录.报告发放%Type;
  Begin
  
    Begin
      Select Nvl(报告发放, 0) Into v_报告发放 From 影像报告记录 Where Id = 报告id_In;
    Exception
      When Others Then
        v_报告发放 := 0;
    End;
  
    Update 影像报告记录
    Set 报告发放 = Decode(v_报告发放, 0, 1, 0), 报告发放人 = Decode(v_报告发放, 0, 当前操作人_In, '')
    Where Id = 报告id_In;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_Reportrelease;

  --5、新增，修改报告
  Procedure p_影像报告记录_新增
  (
    原型id_In     影像报告记录.原型id%Type,
    报告内容_In   影像报告记录.报告内容%Type,
    记录人_In     影像报告记录.记录人%Type,
    最后编辑人_In 影像报告记录.最后编辑人%Type,
    Id_In         影像报告记录.Id%Type,
    医嘱id_In     影像报告记录.医嘱id%Type
  ) As
    --原型ID_In 原型ID
    --保存文档书写记录
    --1 处理匿名数据
    --2 保存文档书写记录、状态
    --3 处理编辑日志
    --4 更新文档任务
    v_报告id    影像报告记录.Id%Type;
    v_原型名称  影像报告原型清单.名称%Type;
    v_设备号    影像报告原型清单.设备号%Type;
    v_报告序号  Number;
    x_Editlog   Xmltype;
    Cur_Time    Date;
    To_Editlist t_Editlist;
    Tn_Editlist t_Editlist;
    v_Msg       Varchar2(200);
    v_New       Number;
    Err_Custom Exception;
    v_Result 影像报告记录.诊断意见%Type;
    v_操作id 影像报告操作记录.Id%Type;
  
    Function Elist_Filter(Source_t t_Editlist) Return t_Editlist Is
      Target_t t_Editlist := t_Editlist();
    Begin
    
      --对独立文档来说，这个函数只是将 Source_t按照编辑时间排序后输出
      For Rs In (Select /*+rule*/
                  *
                 From Table(Cast(Source_t As t_Editlist)) a
                 Order By a.编辑时间) Loop
        Target_t.Extend;
        Target_t(Target_t.Count) := t_Edits(Rs.编辑人, Rs.编辑时间, Rs.签名, Rs.审订签名);
      End Loop;
      Return Target_t;
    End;
  
    Function Build_Editlog
    (
      Tn_Edit t_Editlist,
      To_Edit t_Editlist,
      v_Did   影像报告记录.Id%Type
    ) Return Xmltype Is
      --Tn_Edit 本次保存的新编辑记录；To_Edit上次保存的旧编辑记录
      --将两次编辑记录，组合成一个编辑记录
    
      x_Return Xmltype;
      r_Saveid Raw(16);
      n_Class  Number;
      --n_Class 编辑日志中的操作类别： 1-创建、2-删除、3-编辑、4-签名、5-审订、6-审签、7-撤签
      v_Signor  影像报告记录.创建人%Type;
      v_Adjunct 影像报告记录.创建人%Type;
      Tns_Edit  t_Editlist;
      Tos_Edit  t_Editlist;
    
      Function Atitle(原型id 影像报告原型清单.Id%Type) Return Varchar2 Is
        v_原型名称 影像报告原型清单.名称%Type;
      Begin
        --根据原型ID，返回原型名称
        If 原型id Is Null Then
          Return Null;
        Else
          Select 名称 Into v_原型名称 From 影像报告原型清单 Where Id = 原型id;
          Return v_原型名称;
        End If;
      End;
    
    Begin
      x_Return := Xmltype('<root></root>');
      If v_Did Is Null Then
        --表明是新增文档，新增文档传null进来
        Select Sys_Guid() Into r_Saveid From Dual;
      
        --PACS报告没有子文档，但是下面构造XML的语句保留成跟EMR相同，这里的v_Subiid赋值为空
        Tns_Edit := Elist_Filter(Tn_Edit);
        Select Decode(Tns_Edit(Tns_Edit.Count).签名, 0, 1, 4) Into n_Class From Dual;
        Select Appendchildxml(x_Return,
                               '/root',
                               Xmlelement("operate",
                                          Xmlforest(r_Saveid As "saving_id",
                                                    n_Class As "class",
                                                    To_Char(Cur_Time, 'yyyy-mm-dd hh24:mi:ss') As "cur_time",
                                                    最后编辑人_In As "operator",
                                                    Decode(n_Class, 4, Tns_Edit(Tns_Edit.Count).编辑人, '') As "signer",
                                                    '' As Adjunct)))
        Into x_Return
        From Dual;
      Else
        --不是新增的文档？
        Select Sys_Guid() Into r_Saveid From Dual;
      
        v_Signor  := '';
        v_Adjunct := '';
        Tns_Edit  := Elist_Filter(Tn_Edit);
        Tos_Edit  := Elist_Filter(To_Edit);
        If Tns_Edit(Tns_Edit.Count).签名 = 1 And Tns_Edit(Tns_Edit.Count).审订签名 = 0 Then
          --最近一次是签名
          If Tos_Edit.Count = 0 Then
            --新增子文档直接签名
            n_Class  := 4;
            v_Signor := Tns_Edit(Tns_Edit.Count).编辑人;
          Elsif Tos_Edit(Tos_Edit.Count).编辑人 Is Null Then
            --之前没签名
            n_Class  := 4;
            v_Signor := Tns_Edit(Tns_Edit.Count).编辑人;
          Elsif Tos_Edit(Tos_Edit.Count).签名 = 1 And Tns_Edit(Tns_Edit.Count).编辑时间 > Tos_Edit(Tos_Edit.Count).编辑时间 Then
            --多次普通签名
            n_Class  := 4;
            v_Signor := Tns_Edit(Tns_Edit.Count).编辑人;
          Elsif Tos_Edit(Tos_Edit.Count).签名 = 1 And Tns_Edit(Tns_Edit.Count).编辑时间 < Tos_Edit(Tos_Edit.Count).编辑时间 Then
            --撤消多次签名
            n_Class   := 7;
            v_Adjunct := Tos_Edit(Tos_Edit.Count).编辑人;
          Elsif Tos_Edit(Tos_Edit.Count).签名 = 1 And Tns_Edit(Tns_Edit.Count).编辑时间 = Tos_Edit(Tos_Edit.Count).编辑时间 Then
            --无变化
            n_Class := -1;
          End If;
        Elsif Tns_Edit(Tns_Edit.Count).签名 = 1 And Tns_Edit(Tns_Edit.Count).审订签名 = 1 Then
          --审订签名
          If Tos_Edit(Tos_Edit.Count).审订签名 = 0 Then
            --之前没审签，可能是已签名或已审订
            n_Class  := 6;
            v_Signor := Tns_Edit(Tns_Edit.Count).编辑人;
          Elsif Tos_Edit(Tos_Edit.Count).审订签名 = 1 And Tns_Edit(Tns_Edit.Count).编辑时间 > Tos_Edit(Tos_Edit.Count).编辑时间 Then
            --多次审签
            n_Class  := 6;
            v_Signor := Tns_Edit(Tns_Edit.Count).编辑人;
          Elsif Tos_Edit(Tos_Edit.Count).审订签名 = 1 And Tns_Edit(Tns_Edit.Count).编辑时间 < Tos_Edit(Tos_Edit.Count).编辑时间 Then
            --撤消多次审签
            n_Class   := 7;
            v_Adjunct := Tos_Edit(Tos_Edit.Count).编辑人;
          Elsif Tos_Edit(Tos_Edit.Count).审订签名 = 1 And Tns_Edit(Tns_Edit.Count).编辑时间 = Tos_Edit(Tos_Edit.Count).编辑时间 Then
            --无变化
            n_Class := -1;
          End If;
        Elsif Tns_Edit(Tns_Edit.Count).编辑人 Is Null And Tos_Edit.Count = 0 Then
          n_Class := 1;
        Elsif Tns_Edit(Tns_Edit.Count).编辑人 Is Null And Tos_Edit(Tos_Edit.Count).签名 = 1 Then
          n_Class   := 7;
          v_Adjunct := Tos_Edit(Tos_Edit.Count).编辑人;
        Elsif Tns_Edit(Tns_Edit.Count).编辑人 Is Null And Tos_Edit(Tos_Edit.Count).编辑人 Is Null Then
          n_Class := 3;
        Elsif Tns_Edit(Tns_Edit.Count).审订签名 = 0 And Tos_Edit(Tos_Edit.Count).审订签名 = 0 Then
          n_Class := 5;
        Elsif Tns_Edit(Tns_Edit.Count).审订签名 = 0 And Tos_Edit(Tos_Edit.Count).审订签名 = 1 Then
          n_Class   := 7;
          v_Adjunct := Tos_Edit(Tos_Edit.Count).编辑人;
        End If;
      
        If n_Class <> -1 Then
          Select Appendchildxml(x_Return,
                                 '/root',
                                 Xmlelement("operate",
                                            Xmlforest(r_Saveid As "saving_id",
                                                      n_Class As "class",
                                                      To_Char(Cur_Time, 'yyyy-mm-dd hh24:mi:ss') As "cur_time",
                                                      最后编辑人_In As "operator",
                                                      Decode(n_Class, 4, v_Signor, 6, v_Signor, '') As "signer",
                                                      v_Adjunct As Adjunct)))
          Into x_Return
          From Dual;
        End If;
      
      End If;
      Return x_Return;
    End Build_Editlog;
  
    Function Get_Nextrptnum
    (
      Antetypename 影像报告原型清单.名称%Type,
      Order_Id     影像报告记录.医嘱id%Type
    ) Return Number Is
      v_序号  Number;
      v_Count Number;
      v_Num   Number;
    Begin
    
      v_Count := 0;
      v_Num   := 1;
      Loop
        Select Count(*) + v_Num Into v_序号 From 影像报告记录 Where 医嘱id = Order_Id;
        Select Count(*)
        Into v_Count
        From 影像报告记录
        Where 医嘱id = Order_Id
        And 文档标题 = Antetypename || '_' || v_序号;
      
        If v_Count = 0 Then
          Exit;
        End If;
      
        v_Num := v_Num + 1;
      End Loop;
    
      Return v_序号;
    End;
  
  Begin
  
    Select 名称, 设备号, Sysdate Into v_原型名称, v_设备号, Cur_Time From 影像报告原型清单 Where Id = 原型id_In;
  
    --------------------1 保存文档书写记录、状态--------------------
    --提取文档的签名和编辑（新增、修改）记录
    Tn_Editlist := b_Pacs_Rptpublic.f_Geteditlist(报告内容_In);
  
    --------------------2 处理编辑日志--------------------
    Select Count(*) Into v_New From 影像报告记录 Where Id = Id_In;
  
    v_报告id := Id_In;
    Select Zlpub_Pacs_取提纲内容byxml(报告内容_In, '诊断意见') Into v_Result From Dual;
    If v_New = 0 Then
      --新增报告
      To_Editlist := t_Editlist();
      x_Editlog   := Build_Editlog(Tn_Editlist, To_Editlist, Null);
    
      --取报告序号
      v_报告序号 := Get_Nextrptnum(v_原型名称, 医嘱id_In);
    
      Insert Into 影像报告记录
        (Id, 原型id, 文档标题, 报告内容, 创建时间, 创建人, 报告状态, 最后编辑时间, 最后编辑人, 编辑日志, 医嘱id, 记录人, 诊断意见, 设备号)
      Values
        (v_报告id, 原型id_In, v_原型名称 || '_' || v_报告序号, 报告内容_In, Cur_Time, 最后编辑人_In, 1, Cur_Time, 最后编辑人_In, x_Editlog,
         医嘱id_In, 记录人_In, v_Result, v_设备号);
      Insert Into 病人医嘱报告 (医嘱id, 检查报告id) Values (医嘱id_In, v_报告id);
    
      Select Sys_Guid() Into v_操作id From Dual;
      Insert Into 影像报告操作记录
        (Id, 报告id, 医嘱id, 文档标题, 操作人, 操作时间, 操作类型)
      Values
        (v_操作id, v_报告id, 医嘱id_In, v_原型名称 || '_' || v_报告序号, 最后编辑人_In, Sysdate, 6);
    
    Else
      --提取文件原始编辑记录,必需在更新之前提取
      Select b_Pacs_Rptpublic.f_Geteditlist(报告内容) Into To_Editlist From 影像报告记录 Where Id = v_报告id;
    
      x_Editlog := Build_Editlog(Tn_Editlist, To_Editlist, v_报告id);
      Select Appendchildxml(编辑日志, '/root', Extract(x_Editlog, '/root/*'))
      Into x_Editlog
      From 影像报告记录
      Where Id = v_报告id;
    
      Update 影像报告记录
      Set 报告内容 = 报告内容_In, 最后编辑时间 = Cur_Time, 最后编辑人 = 最后编辑人_In, 编辑日志 = x_Editlog, 记录人 = 记录人_In, 诊断意见 = v_Result
      Where Id = v_报告id;
    End If;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, v_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_影像报告记录_新增;

  --6、获取书写的文档内容
  Procedure p_Get_Doc_Content
  (
    Val      Out t_Refcur,
    Docid_In 影像报告记录.Id%Type
  ) As
  Begin
    Open Val For
      Select (Nvl(a.报告内容, Xmltype('<ZLXML/>'))).Getclobval() As 报告内容 From 影像报告记录 a Where a.Id = Docid_In;
  End;

  --7、设置报告打印作废信息
  Procedure p_Checkrejectsignature
  (
    Signdate_In Date,
    报告id_In   影像报告操作记录.报告id%Type,
    作废人_In   影像报告操作记录.作废人%Type,
    作废说明_In 影像报告操作记录.作废说明%Type,
    Val         Out Sys_Refcursor
  ) As
  Begin
    Open Val For
      Select 操作人, 操作时间
      From 影像报告操作记录
      Where 报告id = 报告id_In
      And 操作类型 = 1
      And 操作时间 >= Signdate_In
      And 作废时间 Is Null
      Order By 操作时间 Asc;
    --作废打印记录
    Update 影像报告操作记录 b
    Set 作废人 = 作废人_In, 作废时间 = Sysdate, b.作废说明 = 作废说明_In
    Where 报告id = 报告id_In
    And 操作类型 = 1
    And 操作时间 >= Signdate_In;
  
  End p_Checkrejectsignature;

  --8、查询相应原型下的最大序号
  Procedure p_Get_Samplelist_Maxseqnum
  (
    Val       Out t_Refcur,
    原型id_In 影像报告范文清单.原型id%Type
  ) As
  Begin
    Open Val For
      Select Nvl(Max(a.编号), 0) + 1 As Num From 影像报告范文清单 a Where a.原型id = 原型id_In;
  End;

  --9、删除文档范文
  Procedure p_Del_影像报告范文清单(Id_In 影像报告范文清单.Id%Type) As
  Begin
    Delete From 影像报告范文清单 Where Id = Id_In;
  End;

  --10、添加文档的操作日志
  Procedure p_影像报告操作记录_Add
  (
    Id_In       影像报告操作记录.Id%Type,
    报告id_In   影像报告操作记录.报告id%Type,
    操作人_In   影像报告操作记录.操作人%Type,
    操作类型_In 影像报告操作记录.操作类型%Type
  ) As
    n_医嘱id   影像报告操作记录.医嘱id%Type;
    n_文档标题 影像报告记录.文档标题%Type;
  Begin
  
    Begin
      Select 医嘱id, 文档标题 Into n_医嘱id, n_文档标题 From 影像报告记录 Where Id = 报告id_In;
    Exception
      When Others Then
        Null;
    End;
    If n_医嘱id Is Not Null Then
      Insert Into 影像报告操作记录
        (Id, 报告id, 医嘱id, 文档标题, 操作人, 操作时间, 操作类型)
      Values
        (Id_In, 报告id_In, n_医嘱id, n_文档标题, 操作人_In, Sysdate, 操作类型_In);
      If 操作类型_In = 1 Then
        Update 影像报告记录 Set 报告打印 = 1 Where Id = 报告id_In;
      End If;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --11、删除报告
  Procedure p_影像报告记录_删除(报告_Id_In 影像报告记录.Id%Type) As
  Begin
  
    Delete From 影像报告记录 Where 影像报告记录.Id = Hextoraw(报告_Id_In);
  
    Delete From 病人医嘱报告 Where 检查报告id = Hextoraw(报告_Id_In);
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_影像报告记录_删除;

  --12、获取签名类型
  Procedure p_Get_Sysconfigsignature
  (
    Val       Out t_Refcur,
    科室id_In In 部门表.Id%Type
  ) Is
  Begin
    --返回用户, 模块号,功能
    Open Val For
      Select Zl_Fun_Getsignpar(7, 科室id_In) As 签名类型 From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --13、获取账户签名印章
  Procedure p_Get_Personsignimg
  (
    Val   Out t_Refcur,
    Id_In In 人员表.Id%Type
  ) Is
    --v_sql Varchar2(1000);
    --n_count Number(5);
  Begin
    --Select Count(*) Into n_Count From user_tables Where table_name =Upper('影像签名图片');
  
    --If n_Count > 0 Then
    --   v_sql := 'Truncate Table 影像签名图片';
    --   Execute Immediate v_sql;
  
    --   v_sql := 'Insert Into 影像签名图片 Select a.id, to_lob(a.签名图片) as 签名图片 From 人员表 a Where a.ID=' || ID_In;
    --   Execute Immediate v_sql;
    --Else
    --   v_sql := 'Create GLOBAL TEMPORARY TABLE 影像签名图片 ON COMMIT PRESERVE ROWS AS Select a.id, to_lob(a.签名图片) as 签名图片 From 人员表 a Where a.ID=' || ID_In;
    --   Execute Immediate v_sql;
    --End If;
  
    --v_sql := 'Select 签名图片 From 影像签名图片 Where Id=:ID';
  
    ----返回用户, 模块号,功能
    --Open  Val For v_sql Using ID_In;
  
    Open Val For
      Select 签名图片 From 人员表 Where Id = Id_In;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --14、获取签名的证书信息
  Procedure p_Get_Signcertinfo
  (
    Val       Out t_Refcur,
    证书id_In 人员证书记录.Id%Type
  ) Is
  Begin
    Open Val For
      Select Id, Certdn, Certsn, Signcert, Enccert From 人员证书记录 Where Id = 证书id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --15、更新报告状态
  Procedure p_Update_Reportstate
  (
    报告id_In   影像报告记录.Id%Type,
    报告状态_In 影像报告记录.报告状态%Type,
    审核人_In   影像报告记录.最后审核人%Type
  ) Is
  Begin
    --报告状态1-未签名；2-已诊断；3-已审核；4-已终审；5-诊断驳回；6-审核驳回
    --如果报告状态是1-未签名；2-已诊断;5-诊断驳回，此时是没有审核人的
    If (报告状态_In = 1) Or (报告状态_In = 2) Or (报告状态_In = 5) Then
      Update 影像报告记录 Set 报告状态 = 报告状态_In, 最后审核人 = Null, 最后审核时间 = Null Where Id = 报告id_In;
    Elsif (报告状态_In = 3) Or (报告状态_In = 4) Then
      Update 影像报告记录
      Set 报告状态 = 报告状态_In, 最后审核人 = 审核人_In, 最后审核时间 = Sysdate
      Where Id = 报告id_In;
    Else
      Update 影像报告记录 Set 报告状态 = 报告状态_In Where Id = 报告id_In;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --16、获取报告状态
  Procedure p_Get_Reportstate
  (
    Val       Out t_Refcur,
    报告id_In 影像报告记录.Id%Type
  ) Is
  Begin
    Open Val For
      Select 报告状态 From 影像报告记录 Where Id = 报告id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --17、报告驳回
  Procedure p_Reject_Report
  (
    医嘱id_In   影像报告驳回.医嘱id%Type,
    报告id_In   影像报告驳回.检查报告id%Type,
    驳回理由_In 影像报告驳回.驳回理由%Type,
    驳回时间_In 影像报告驳回.驳回时间%Type,
    驳回人_In   影像报告驳回.驳回人%Type,
    待处理人_In 影像报告记录.待处理人%Type,
    报告状态_In 影像报告记录.报告状态%Type
  ) Is
  Begin
    Insert Into 影像报告驳回
      (Id, 医嘱id, 检查报告id, 驳回理由, 驳回时间, 驳回人)
    Values
      (影像报告驳回_Id.Nextval, 医嘱id_In, 报告id_In, 驳回理由_In, 驳回时间_In, 驳回人_In);
  
    Update 影像报告记录 Set 报告状态 = 报告状态_In, 待处理人 = 待处理人_In Where Id = 报告id_In;
  
    --Update 病人医嘱发送 Set 执行过程=-1 Where 医嘱ID= 医嘱ID_IN;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --17.1、撤销报告驳回
  Procedure p_Reject_Cancel
  (
    Id_In       影像报告驳回.Id%Type,
    报告id_In   影像报告驳回.检查报告id%Type,
    报告状态_In 影像报告记录.报告状态%Type
  ) Is
  Begin
    Update 影像报告驳回 Set 是否撤销 = 1 Where Id = Id_In;
    Update 影像报告记录 Set 报告状态 = 报告状态_In, 待处理人 = '' Where Id = 报告id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --18、获取报告驳回信息
  Procedure p_Get_Rejectinfo
  (
    Val       Out t_Refcur,
    报告id_In 影像报告驳回.检查报告id%Type
  ) Is
  Begin
    Open Val For
      Select a.Id, a.驳回理由, a.驳回时间, a.驳回人, Nvl(a.是否撤销, 0) As 驳回状态, b.报告状态
      From 影像报告驳回 a, 影像报告记录 b
      Where a.检查报告id = 报告id_In
      And a.检查报告id = b.Id
      Order By 驳回时间;
  End;

  --19、获取原型动作
  Procedure p_Get_Doc_Process
  (
    Val       Out t_Refcur,
    原型id_In 影像报告动作.原型id%Type
  ) As
  Begin
    Open Val For
      Select Rawtohex(p.Id) Id, p.名称 As 动作名称, e.名称 As 事件名称, e.种类 As 事件种类, e.元素iid As 元素iid, p.动作类型, p.序号, p.说明, p.可否手工执行,
             (Nvl(p.内容, Xmltype('<NULL/>'))).Getclobval() As 内容, Rawtohex(p.事件id) 事件id
      From 影像报告动作 p, 影像报告事件 e
      Where p.事件id = e.Id(+)
      And p.原型id = 原型id_In
      Order By 动作类型, 序号;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Doc_Process;

  --20、通过学科筛选获得相应的范文信息
  Procedure p_Get_Samplelist_By_Conditions
  (
    Val          Out t_Refcur,
    原型id_In    Varchar2,
    学科_In      Varchar2,
    Condition_In Varchar2, --过滤筛选
    作者_In      Varchar2
  ) As
  Begin
  
    Open Val For
      Select /*+ rule*/
       Rawtohex(a.Id) Id, a.名称, a.作者, a.说明, Nvl2(a.说明, a.说明 || '作者:' || a.作者, '作者:' || a.作者) Content, a.标签, a.学科
      From 影像报告范文清单 a
      Where a.原型id = Hextoraw(原型id_In)
      And ((a.学科 Is Null And a.是否私有 = 0) Or 学科_In Is Null Or a.作者 = 作者_In Or
            (a.学科 Is Not Null And b_Pacs_Rptpublic.f_If_Intersect(a.学科, 学科_In) > 0 And a.是否私有 = 0))
      And (Condition_In Is Null Or
            (a.标签 Is Not Null And Condition_In Is Not Null And b_Pacs_Rptpublic.f_If_Intersect(a.标签, Condition_In) > 0))
      Order By a.编号;
  
  End p_Get_Samplelist_By_Conditions;

  --21、通过部门ID获取部门名称
  Procedure p_Get_部门名称_By_Id
  (
    Val   Out t_Refcur,
    Id_In 部门表.Id%Type
  ) Is
  Begin
    Open Val For
      Select 名称 From 部门表 Where Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_部门名称_By_Id;

  --22、提取所有预备提纲
  Procedure p_Get_Allpreoutlines(Val Out t_Refcur) Is
  Begin
    Open Val For
      Select Rawtohex(Id) Id, a.编码, a.名称 From 影像报告预备提纲 a Order By a.编码;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Allpreoutlines;

  --23、提取文档标题
  Procedure p_Get_Reporttitle_By_Id
  (
    Val   Out t_Refcur,
    Id_In 影像报告记录.Id%Type
  ) Is
  Begin
    Open Val For
      Select 文档标题 From 影像报告记录 Where Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Reporttitle_By_Id;

  --24、提取报告锁定人
  Procedure p_Get_报告锁定人_By_Id
  (
    Val   Out t_Refcur,
    Id_In 影像报告记录.Id%Type
  ) Is
  Begin
    Open Val For
      Select 锁定人 From 影像报告记录 Where Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_报告锁定人_By_Id;

  --25、通过医嘱ID获取报告列表
  Procedure p_Get_影像报告记录_By_医嘱id
  (
    Val       Out t_Refcur,
    医嘱id_In 影像报告记录.医嘱id%Type
  ) Is
  Begin
    Open Val For
      Select Rawtohex(Id) As Reportid, Rawtohex(原型id) As Antetypeid, 医嘱id As Orderid, 文档标题 As Reportname,
             创建时间 As Reportdate,
             Decode(Nvl(报告状态, 0), 1, '编辑中', 2, '已诊断', 3, '已审核', 4, '已终审', 5, '诊断驳回', '审核驳回') As Reportstate,
             创建人 As Createuser, 最后审核时间 As Examineydate, 最后审核人 As Examineyuser,
             Decode(Nvl(结果阳性, 0), 1, '阳性', '') As Resultpositive, Nvl(报告质量, 0) As Innerquality, ' ' As Reportquality,
             Decode(Nvl(报告打印, 0), 0, '未打印', '已打印') As Reportprint,
             Decode(Nvl(报告发放, 0), 0, '未发放', '已发放') As Reportrelease, 记录人 As Recdoctor, 锁定人 As RecLocker, ' ' As Locked
      From 影像报告记录
      Where 医嘱id = 医嘱id_In
      Order By Reportdate Desc;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_影像报告记录_By_医嘱id;

  --26、查询影像流程参数值
  Procedure p_Get_影像流程参数值
  (
    Val       Out t_Refcur,
    科室id_In 影像流程参数.科室id%Type
  ) Is
  Begin
    Open Val For
      Select 参数名, 参数值 From 影像流程参数 Where 科室id = 科室id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_影像流程参数值;

  --27、根据医嘱ID，查询对应的原型列表
  Procedure p_Get_影像原型列表_By_医嘱id
  (
    Val     Out t_Refcur,
    医嘱_In 影像检查记录.医嘱id%Type
  ) Is
  Begin
    Open Val For
      Select Rawtohex(c.Id) As Antetypeid, c.名称 As Antetypename, c.说明
      From 病人医嘱记录 a, 影像报告原型应用 b, 影像报告原型清单 c
      Where a.Id = 医嘱_In
      And a.诊疗项目id = b.诊疗项目id
      And b.报告原型id = c.Id
      And a.病人来源 = b.应用场合
      Order By c.使用次数 Desc;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_影像原型列表_By_医嘱id;

  --28、根据报告ID查询打印记录
  Procedure p_Get_Reportprintlog_By_报告id
  (
    Val     Out Sys_Refcursor,
    报告_In 影像报告操作记录.报告id%Type
  ) Is
  Begin
    Open Val For
      Select c.文档标题, b.操作人, To_Char(b.操作时间, 'yyyy-MM-dd HH24:mi') 打印时间, b.作废人,
             To_Char(b.作废时间, 'yyyy-MM-dd HH24:mi') 作废时间, b.作废说明
      From 影像报告操作记录 b, 影像报告记录 c
      Where c.Id = 报告_In
      And b.报告id = c.Id
      And 操作类型 = 1
      Order By b.操作时间;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Reportprintlog_By_报告id;

  --29、根据医嘱ID查询报告发放列表
  Procedure p_Get_Reportreleaselist
  (
    Val     Out t_Refcur,
    医嘱_In 影像报告记录.医嘱id%Type
  ) Is
  Begin
    Open Val For
      Select Rawtohex(Id) As 报告id, 文档标题 As 报告名称, 最后编辑时间 As 报告日期, Decode(Nvl(报告发放, 0), 0, '未发放', '已发放') As 报告发放
      From 影像报告记录
      Where 报告状态 Between 2 And 4
      And 医嘱id = 医嘱_In;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Reportreleaselist;

  --30、根据报告ID查询驳回记录数量
  Procedure p_Get_Rejectedcount
  (
    Val     Out t_Refcur,
    报告_In 影像报告驳回.检查报告id%Type
  ) Is
  Begin
    Open Val For
      Select Count(*) As 驳回数量 From 影像报告驳回 Where 检查报告id = 报告_In;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Rejectedcount;

  --31、根据医嘱ID查询报告动作需要的一些ID们
  Procedure p_Get_Docprocess_Ids
  (
    Val     Out t_Refcur,
    医嘱_In 病人医嘱记录.Id%Type
  ) Is
  Begin
    Open Val For
      Select Id As 医嘱id, 主页id, 挂号单 From 病人医嘱记录 Where Id = 医嘱_In;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Docprocess_Ids;

  --32、根据医嘱ID和报告ID查询报告的一些参数
  Procedure p_Get_Docinfo
  (
    Val       Out t_Refcur,
    医嘱id_In 影像检查记录.医嘱id%Type,
    报告id_In 影像报告记录.Id%Type
  ) Is
  Begin
    If 报告id_In Is Null Then
      Open Val For
        Select 执行科室id, '创建人' As 创建人 From 影像检查记录 Where 医嘱id = 医嘱id_In;
    Else
      Open Val For
        Select 执行科室id, 创建人
        From 影像检查记录 a, 影像报告记录 b
        Where a.医嘱id = b.医嘱id
        And b.Id = 报告id_In;
    End If;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Docinfo;

  --33、查询一个检查中相同原型ID的报告数量
  Procedure p_Get_Sameantetypedoccounts
  (
    Val       Out t_Refcur,
    医嘱id_In 影像报告记录.医嘱id%Type,
    原型id_In 影像报告记录.原型id%Type
  ) Is
  Begin
    Open Val For
      Select Count(Id) As Doccounts
      From 影像报告记录
      Where 医嘱id = 医嘱id_In
      And 原型id = 原型id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Sameantetypedoccounts;

  --34、提取报告图存储信息
  Procedure p_Get_Docimagesaveinof_By_Id
  (
    Val   Out t_Refcur,
    Id_In 影像报告记录.Id%Type
  ) Is
  Begin
    Open Val For
      Select 设备号, 创建时间 From 影像报告记录 Where Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Docimagesaveinof_By_Id;

  --35、修改原型使用次数
  Procedure p_Update_Antetypeusecount(Id_In 影像报告原型清单.Id%Type) Is
  Begin
    Update 影像报告原型清单 Set 使用次数 = 使用次数 + 1 Where Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Antetypeusecount;

  --36、更新影像检查图像的报告图标记
  Procedure p_Update_Rptimage
  (
    Uid_In        影像检查图象.图像uid%Type,
    Actiontype_In Number
  ) Is
    v_Sql Varchar2(4000);
    No_Column Exception;
    Pragma Exception_Init(No_Column, -00904);
  Begin
    If Actiontype_In = 1 Then
      v_Sql := 'Update 影像检查图象 Set 报告图 = Nvl(报告图, 0) + 1 Where 图像uid = :1';
    Else
      v_Sql := 'Update 影像检查图象
      Set 报告图 = Decode(报告图, Null, Null, Decode(Nvl(报告图, 0) - 1, 0, Null, Nvl(报告图, 0) - 1))
      Where 图像uid = :1';
    End If;
    Execute Immediate v_Sql
      Using Uid_In;
  Exception
    When No_Column Then
      --兼容处理，10.36新增加 报告图 字段，问题号 103996
      Null;
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Rptimage;

  --37、提取打印控制信息
  Procedure p_Get_Printcontrol
  (
    Val       Out t_Refcur,
    报告id_In 影像报告记录.Id%Type
  ) Is
    v_紧急     Number;
    v_打印控制 Number;
  Begin
  
    Select Nvl(Decode(a.急诊, 1, 1, b.紧急标志), 0) As 紧急
    Into v_紧急
    From 病人挂号记录 a, 病人医嘱记录 b
    Where a.No(+) = b.挂号单
    And b.Id = (Select c.医嘱id From 影像报告记录 c Where c.Id = 报告id_In);
  
    Select Nvl(Extractvalue(b.Column_Value, '/root/print_limit'), 0) Printlimit
    Into v_打印控制
    From 影像报告原型清单 a, Table(Xmlsequence(Extract(a.控制选项, '/root'))) b, 影像报告记录 c
    Where a.Id = c.原型id
    And c.Id = 报告id_In;
  
    Open Val For
      Select v_紧急 As 紧急, v_打印控制 As 打印控制 From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Printcontrol;

End b_Pacs_Rptmanage;

/


--影像报告插件管理(---定义部分---)***************************************************
Create Or Replace Package b_Pacs_Rptpluginoriginal Is
  Type t_Refcur Is Ref Cursor;

  -- 1、功    能：获取历史报告记录
  Procedure p_Getreporthistory
  (
    Val                   Out t_Refcur,
    医嘱id_In             In 病人医嘱记录.Id%Type,
    人员id_In             In 部门人员.人员id%Type,
    当前科室id_In         In 部门人员.部门id%Type,
    查看其他科历史报告_In In Number := 0
  );

  --2、功    能：获取对应报告内容
  Procedure p_Getreportcontent
  (
    Val           Out t_Refcur,
    报告id_In     In Varchar2,
    Editortype_In Number := 0 --0:PACS报告编辑器，1--电子病历编辑器，2--报告文档编辑器
  );

  --3、功    能：根据医嘱ID获取检查信息
  Procedure p_Getstudyinfobyadviceid
  (
    Val       Out t_Refcur,
    医嘱id_In In 影像检查记录.医嘱id%Type
  );

  --4、功    能：获取报告图像总数
  Procedure p_Getreportimagecount
  (
    Val         Out t_Refcur,
    查询条件_In In Varchar2
  );

  --5、功    能：获取报告图像数据
  Procedure p_Getreportimagedata
  (
    Val         Out t_Refcur,
    查询条件_In In Varchar2,
    开始位置_In In Number,
    结束位置_In In Number
  );

  --6、功    能：获取预览图像总数
  Procedure p_Getstudyimagecount
  (
    Val         Out t_Refcur,
    查询条件_In In Varchar2,
    是否临时_In In Number := 0
  );

  --7、功    能：获取预览图像数据
  Procedure p_Getstudyimagedata
  (
    Val         Out t_Refcur,
    查询方式_In In Varchar2,
    查询条件_In In Varchar2,
    开始位置_In In Number,
    结束位置_In In Number,
    是否临时_In In Number
  );

  --8、功能：获取临时图像序列
  Procedure p_Get_Tempimageseries
  (
    Val         Out t_Refcur,
    时间范围_In In Number,
    姓名_In     In 影像临时记录.姓名%Type := Null
  );

  --9、功能;获取图像备注
  Procedure p_Get_Normalnote(Val Out t_Refcur);

  --10、功能：插入常用图像备注
  Procedure p_Insert_Normalnote
  (
    Note_In In 影像字典内容.名称%Type,
    Code_In 影像字典内容.简码%Type
  );

  --11、功能：修改常用图像备注
  Procedure p_Edit_Normalnote
  (
    Note_In In 影像字典内容.名称%Type,
    Num_In  影像字典内容.编号%Type
  );

  --12、功能：删除常用图像备注
  Procedure p_Del_Normalnote(Num_In 影像字典内容.编号%Type);

  --13、功能：获取备注的下一个编码
  Procedure p_Get_Normalnum(Val Out t_Refcur);
  --14、功能：获取插件ID
  Procedure p_Get_Plugid
  (
    Val     Out t_Refcur,
    类名_In In 影像报告插件.类名%Type
  );

  --15、功能：插入编辑器字体参数
  Procedure p_Setfontparam
  (
    Font_In Nvarchar2,
    User_In Nvarchar2
  );

  --16、功能：获取编辑器字体参数
  Procedure p_Getfontparam
  (
    Val     Out t_Refcur,
    User_In Nvarchar2
  );

  --17、功能：插入编辑器窗体参数
  Procedure p_Setformparam
  (
    Form_In Nvarchar2,
    User_In Nvarchar2
  );

  --18、功能：获取编辑器字体参数
  Procedure p_Getformparam
  (
    Val     Out t_Refcur,
    User_In Nvarchar2
  );

  --19、功能：根据图像UID获取检查信息
  Procedure p_Getstudyinfobyimageuid
  (
    Val        Out t_Refcur,
    医嘱id_In  In 影像检查记录.医嘱id%Type,
    图像uid_In In 影像检查图象.图像uid%Type
  );

  --20、功能：根据检查UID获取FTP信息
  Procedure p_Getftpinfobystudyuid
  (
    Val        Out t_Refcur,
    检查uid_In In 影像检查记录.检查uid%Type
  );

  --21、功能：根据科室ID获取FTP信息
  Procedure p_Getftpinfobydeptid
  (
    Val       Out t_Refcur,
    科室id_In In 影像流程参数.科室id%Type
  );

  --22、功能：根据医嘱ID获取FTP信息
  Procedure p_Getftpinfobyadvicetid
  (
    Val       Out t_Refcur,
    医嘱id_In In 影像检查记录.医嘱id%Type
  );

  --23、功能：获取检查UID
  Procedure p_Getstudyuid
  (
    Val        Out t_Refcur,
    检查uid_In In 影像检查记录.检查uid%Type
  );

  --24、功能：获取序列UID
  Procedure p_Getseriesuid
  (
    Val        Out t_Refcur,
    序列uid_In In 影像检查序列.序列uid%Type
  );

  --25、功能：根据设备号获取设备信息
  Procedure p_Getdeviceinfo
  (
    Val       Out t_Refcur,
    设备号_In In 影像设备目录.设备号%Type
  );

  --26、获取医技站存储设备号
  Procedure p_Getdeviceidbyadviceid
  (
    Val       Out t_Refcur,
    医嘱id_In In 病人医嘱发送.医嘱id%Type
  );
End b_Pacs_Rptpluginoriginal;

/

--影像报告范文管理(---实现部分---)***************************************************
Create Or Replace Package Body b_Pacs_Rptpluginoriginal Is

  --1、功    能：获取历史报告记录
  Procedure p_Getreporthistory
  (
    Val                   Out t_Refcur,
    医嘱id_In             In 病人医嘱记录.Id%Type,
    人员id_In             In 部门人员.人员id%Type,
    当前科室id_In         In 部门人员.部门id%Type,
    查看其他科历史报告_In In Number := 0
  ) Is
    Strsql     Varchar2(4000);
    Strsqlback Varchar2(4000);
    Strfilter  Varchar2(400);
  Begin
    If 查看其他科历史报告_In = 1 Then
      Strfilter := ' ';
    Else
      Strfilter := ' And c.执行科室id+0 in (select 部门id from 部门人员 where 人员id = ' || 人员id_In ||
                   ' union all select to_Number(' || 当前科室id_In || ') from dual) ';
    End If;
  
    Strsql := 'Select 2 as 报告类型, f.编码' || '||''-''||' || 'f.名称 As 科室名称, c.Id As 医嘱id, a.影像类别 as 类别,b.创建人 as 报告人,' ||
              'to_char(b.创建时间,''yyyy-mm-dd hh24:mi:ss'') as 创建时间,b.文档标题 报告名称, c.医嘱内容, TO_CHAR(RAWTOHEX(b.id)) 报告ID ' ||
              'From 影像检查记录 A, 影像报告记录 B, 病人医嘱记录 C, 影像检查记录 D, 病人医嘱记录 E, 部门表 F ' ||
              'Where a.医嘱id = b.医嘱id And d.医嘱id = e.Id And e.Id =' || 医嘱id_In ||
              ' And c.执行科室ID = F.ID And b.医嘱id = c.Id And ' ||
              '(c.病人id = e.病人id Or a.关联id = d.关联id) And c.相关id Is Null ' || Strfilter || ' union all ' ||
              'Select 1 as 报告类型, g.编码' || '||''-''||' || 'g.名称 As 科室名称, c.Id As 医嘱id, a.影像类别 as 类别, a.报告人, ' ||
              'to_char(f.创建时间,''yyyy-mm-dd hh24:mi:ss'') as 创建时间, a.影像类别||''报告'' 报告名称, c.医嘱内容,TO_CHAR( b.病历id) as 报告ID ' ||
              'From 影像检查记录 A, 病人医嘱报告 B, 病人医嘱记录 C, 影像检查记录 D, 病人医嘱记录 E, 电子病历记录 F, 部门表 G ' ||
              'Where a.医嘱id = b.医嘱id And d.医嘱id = e.Id And e.Id = ' || 医嘱id_In ||
              ' And c.执行科室ID = g.ID And b.医嘱id = c.Id And b.病历ID Is Not Null And ' ||
              '(c.病人id = e.病人id Or a.关联id = d.关联id) And c.相关id Is Null And b.病历id = f.id ' || Strfilter;
  
    Strsqlback := Strsql;
    Strsqlback := Replace(Strsqlback, '影像检查记录', 'H影像检查记录');
    Strsqlback := Replace(Strsqlback, '病人医嘱报告', 'H病人医嘱报告');
    Strsqlback := Replace(Strsqlback, '病人医嘱记录', 'H病人医嘱记录');
  
    Strsql := Strsql || ' UNION ALL ' || Strsqlback || ' Order By 创建时间 Asc';
  
    Open Val For Strsql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getreporthistory;

  --2、功    能：获取对应报告内容
  Procedure p_Getreportcontent
  (
    Val           Out t_Refcur,
    报告id_In     Varchar2,
    Editortype_In Number := 0 --0:电子病历编辑器，1--PACS报告编辑器，2--报告文档编辑器
  ) Is
    Strsql Varchar2(1000);
  Begin
    If Editortype_In = 1 Then
      Strsql := 'Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文,b.开始版 as 版本 From 电子病历内容 a,电子病历内容 b ' || 'Where a.文件id = ' ||
                报告id_In || ' And a.对象类型 = 3 And a.Id = b.父ID And b.对象类型 = 2 and b.终止版=0 ';
    Elsif Editortype_In = 0 Then
      Strsql := 'select 内容 from 电子病历格式 where 文件ID=' || 报告id_In;
    Else
      Strsql := 'Select 报告内容 As 内容 From 影像报告记录 Where ID=HexToRaw(''' || 报告id_In || ''')';
    End If;
  
    Open Val For Strsql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getreportcontent;

  --3、功    能：根据医嘱ID获取检查信息
  Procedure p_Getstudyinfobyadviceid
  (
    Val       Out t_Refcur,
    医嘱id_In In 影像检查记录.医嘱id%Type
  ) Is
    Strsql Varchar2(100);
  Begin
    Strsql := 'Select 检查UID,报告图象,接收日期,检查号,姓名,性别,年龄 from 影像检查记录 where 医嘱ID =' || 医嘱id_In;
    Open Val For Strsql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getstudyinfobyadviceid;

  --4、功    能：获取报告图像总数
  Procedure p_Getreportimagecount
  (
    Val         Out t_Refcur,
    查询条件_In In Varchar2
  ) Is
  Begin
    Open Val For
      Select Count(b.Column_Value) 返回值
      From 影像检查记录 a, Table(Cast(f_Str2list(Replace(a.报告图象, ';', ',')) As Zltools.t_Strlist)) b
      Where 医嘱id = 查询条件_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getreportimagecount;

  --5、功    能：获取报告图像数据
  Procedure p_Getreportimagedata
  (
    Val         Out t_Refcur,
    查询条件_In In Varchar2,
    开始位置_In In Number,
    结束位置_In In Number
  ) Is
  Begin
    Open Val For
      Select *
      From (Select Rownum As 顺序号, Rownum As 图像号, b.Ftp用户名 As User1, b.Ftp密码 As Pwd1, b.Ip地址 As Host1,
                    '/' || b.Ftp目录 || '/' As Root1,
                    Decode(a.接收日期, Null, '', To_Char(a.接收日期, 'YYYYMMDD') || '/') || a.检查uid || '/' ||
                     Replace(Trim(d.Column_Value), '.jpg', '') As Url, b.设备号 As 设备号1, c.Ftp用户名 As User2, c.Ftp密码 As Pwd2,
                    c.Ip地址 As Host2, '/' || c.Ftp目录 || '/' As Root2, c.设备号 As 设备号2,
                    Replace(Trim(d.Column_Value), '.jpg', '') As 图像uid, a.检查uid, '' 序列uid, 0 动态图, '' 编码名称, '' 采集时间,
                    '' 录制长度, '' 报告图
             From 影像检查记录 a, 影像设备目录 b, 影像设备目录 c, Table(Cast(f_Str2list(Replace(a.报告图象, ';', ',')) As Zltools.t_Strlist)) d
             Where a.位置一 = b.设备号(+)
             And a.位置二 = c.设备号(+)
             And a.医嘱id = 查询条件_In)
      Where 顺序号 >= 开始位置_In
      And 顺序号 <= 结束位置_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getreportimagedata;

  --6、功    能：获取预览图像总数
  Procedure p_Getstudyimagecount
  (
    Val         Out t_Refcur,
    查询条件_In In Varchar2,
    是否临时_In In Number := 0
  ) Is
    Strsql Varchar2(2000);
  Begin
    If 是否临时_In = 0 Then
      Strsql := 'select T1.返回值+T2.返回值 as 返回值 from ' || '(select count(1) as 返回值 from 影像检查图象 a, 影像检查序列 b, 影像检查记录 c ' ||
                'where a.序列UID=b.序列UID and b.检查UID=c.检查UID and c.医嘱ID=''' || 查询条件_In || ''') T1,' ||
                '(select count(1) as 返回值 from H影像检查图象 a, H影像检查序列 b, 影像检查记录 c ' ||
                'where a.序列UID=b.序列UID and b.检查UID=c.检查UID and c.医嘱ID=''' || 查询条件_In || ''') T2';
    Else
      Strsql := 'select count(1)  as 返回值 from 影像临时图象  where  序列UID=''' || 查询条件_In || '''';
    End If;
  
    Open Val For Strsql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getstudyimagecount;

  --7、功    能：获取预览图像数据
  Procedure p_Getstudyimagedata
  (
    Val         Out t_Refcur,
    查询方式_In In Varchar2,
    查询条件_In In Varchar2,
    开始位置_In In Number,
    结束位置_In In Number,
    是否临时_In In Number
  ) Is
    Strsql    Varchar2(2000);
    Strsql2   Varchar2(2000);
    Strfilter Varchar2(100);
    No_Column Exception;
    Pragma Exception_Init(No_Column, -00904);
  Begin
    If 查询方式_In = 0 Then
      Strfilter := 'and c.医嘱ID=''' || 查询条件_In || '''';
    Elsif 查询方式_In = 1 Then
      Strfilter := 'and B.序列UID=''' || 查询条件_In || '''';
    Else
      Strfilter := 'and A.图像UID=''' || 查询条件_In || '''';
    End If;
  
    Strsql := 'Select * from (Select rownum as 顺序号, T.* from(' ||
              'Select A.图像号,D.FTP用户名 As User1,D.FTP密码 As Pwd1,D.IP地址 As Host1,''/''||D.Ftp目录||''/'' As Root1,' ||
              'Decode(C.接收日期,Null,'''',to_Char(C.接收日期,''YYYYMMDD'')||''/'')||C.检查UID||''/''||A.图像UID As URL,d.设备号 as 设备号1,' ||
              'E.FTP用户名 As User2,E.FTP密码 As Pwd2,E.IP地址 As Host2,''/''||E.Ftp目录||''/'' As Root2,' ||
              'e.设备号 as 设备号2, A.图像UID,C.检查UID,B.序列UID,A.动态图,A.编码名称,A.采集时间, A.录制长度,A.报告图 ' ||
              'From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E ' ||
              'Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) ' || Strfilter || ' ' ||
              'Order by 序列UID, 图像号) T ) ' || 'Where 顺序号>=' || 开始位置_In || ' and 顺序号<=' || 结束位置_In || '';
  
    Strsql2 := 'Select * from (Select rownum as 顺序号, T.* from(' ||
               'Select A.图像号,D.FTP用户名 As User1,D.FTP密码 As Pwd1,D.IP地址 As Host1,''/''||D.Ftp目录||''/'' As Root1,' ||
               'Decode(C.接收日期,Null,'''',to_Char(C.接收日期,''YYYYMMDD'')||''/'')||C.检查UID||''/''||A.图像UID As URL,d.设备号 as 设备号1,' ||
               'E.FTP用户名 As User2,E.FTP密码 As Pwd2,E.IP地址 As Host2,''/''||E.Ftp目录||''/'' As Root2,' ||
               'e.设备号 as 设备号2, A.图像UID,C.检查UID,B.序列UID,A.动态图,A.编码名称,A.采集时间, A.录制长度 ' ||
               'From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E ' ||
               'Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) ' || Strfilter || ' ' ||
               'Order by 序列UID, 图像号) T ) ' || 'Where 顺序号>=' || 开始位置_In || ' and 顺序号<=' || 结束位置_In || '';
  
    If 是否临时_In = 1 Then
      Strsql  := Replace(Strsql, '影像检查', '影像临时');
      Strsql2 := Replace(Strsql2, '影像检查', '影像临时');
    End If;
  
    Begin
      Open Val For Strsql;
    Exception
      When No_Column Then
        --兼容处理，10.36新增加 报告图 字段，问题号 103996
        Open Val For Strsql2;
    End;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getstudyimagedata;

  --8、功能：获取临时图像序列
  Procedure p_Get_Tempimageseries
  (
    Val         Out t_Refcur,
    时间范围_In In Number,
    姓名_In     In 影像临时记录.姓名%Type := Null
  ) As
  Begin
    If 姓名_In Is Null Then
      Open Val For
        Select b.序列uid, a.姓名, a.检查号 As 序号, a.接收日期
        From 影像临时记录 a, 影像临时序列 b
        Where a.检查uid = b.检查uid
        And a.接收日期 Between Sysdate - 时间范围_In And Sysdate
        Order By 序号;
    Else
      Open Val For
        Select b.序列uid, a.姓名, a.检查号 As 序号, a.接收日期
        From 影像临时记录 a, 影像临时序列 b
        Where a.检查uid = b.检查uid
        And a.接收日期 Between Sysdate - 时间范围_In And Sysdate
        And a.姓名 = 姓名_In
        Order By 序号;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --9、功能：获取图像备注
  Procedure p_Get_Normalnote(Val Out t_Refcur) As
  Begin
    Open Val For
      Select b.编号 As 编号, b.名称 As 名称
      From 影像字典清单 a, 影像字典内容 b
      Where a.Id = b.字典id
      And a.名称 = '影像图像备注';
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --10、功能：插入常用图像备注
  Procedure p_Insert_Normalnote
  (
    Note_In In 影像字典内容.名称%Type,
    Code_In 影像字典内容.简码%Type
  ) As
    n_Num         Number;
    Dictionary_Id Varchar2(36);
  Begin
    Select Id Into Dictionary_Id From 影像字典清单 Where 说明 = '影像图像备注';
    Select Decode(Max(To_Number(编号)), Null, 0, Max(To_Number(编号)))
    Into n_Num
    From 影像字典内容
    Where 字典id = Dictionary_Id;
    n_Num := n_Num + 1;
    Insert Into 影像字典内容
      (字典id, 编号, 名称, 说明)
    Values
      (Dictionary_Id, To_Char(n_Num), Note_In, '影像图像备注');
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Insert_Normalnote;

  --11、功能：修改常用图像备注
  Procedure p_Edit_Normalnote
  (
    Note_In In 影像字典内容.名称%Type,
    Num_In  影像字典内容.编号%Type
  ) As
    Dictionary_Id Varchar2(36);
  Begin
    Select Id Into Dictionary_Id From 影像字典清单 Where 说明 = '影像图像备注';
    Update 影像字典内容 t
    Set t.名称 = Note_In
    Where t.字典id = Dictionary_Id
    And t.编号 = Num_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Normalnote;

  --12、功能：删除常用图像备注
  Procedure p_Del_Normalnote(Num_In 影像字典内容.编号%Type) As
    Dictionary_Id Varchar2(36);
  Begin
    Select Id Into Dictionary_Id From 影像字典清单 Where 说明 = '影像图像备注';
    Delete 影像字典内容 t
    Where t.字典id = Dictionary_Id
    And t.编号 = Num_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Normalnote;

  --13、功能：获取备注的下一个编码
  Procedure p_Get_Normalnum(Val Out t_Refcur) As
    Dictionary_Id Varchar2(36);
  Begin
    Select Id Into Dictionary_Id From 影像字典清单 Where 说明 = '影像图像备注';
    Open Val For
      Select Decode(Max(To_Number(编号)), Null, 1, Max(To_Number(编号) + 1)) 编号
      From 影像字典内容 t
      Where t.字典id = Dictionary_Id;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Normalnum;

  --14、功能：获取插件ID
  Procedure p_Get_Plugid
  (
    Val     Out t_Refcur,
    类名_In In 影像报告插件.类名%Type
  ) Is
  Begin
    Open Val For
      Select Rawtohex(Id) Id From 影像报告插件 Where 类名 = 类名_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Plugid;

  --15、功能：插入编辑器字体参数
  Procedure p_Setfontparam
  (
    Font_In Nvarchar2,
    User_In Nvarchar2
  ) As
    m_Id     Nvarchar2(36);
    Numcount Int;
  Begin
    Select Rawtohex(Id)
    Into m_Id
    From 影像参数说明
    Where 模块 = 'ImageEditor'
    And 参数名 = '字体设置';
    Select Count(*)
    Into Numcount
    From 影像参数取值 t
    Where t.参数id = m_Id
    And t.参数标识 = User_In;
    If Numcount > 0 Then
      Update 影像参数取值 a
      Set a.参数值 = Font_In
      Where a.参数标识 = User_In
      And a.参数id = m_Id;
    Else
      Insert Into 影像参数取值 a (Id, 参数id, 参数标识, 参数值) Values (Sys_Guid(), m_Id, User_In, Font_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Setfontparam;

  --16、功能：获取编辑器字体参数
  Procedure p_Getfontparam
  (
    Val     Out t_Refcur,
    User_In Nvarchar2
  ) As
    m_Id Nvarchar2(36);
  Begin
    Select Rawtohex(Id)
    Into m_Id
    From 影像参数说明
    Where 模块 = 'ImageEditor'
    And 参数名 = '字体设置';
    Open Val For
      Select a.参数值
      From 影像参数取值 a
      Where a.参数id = m_Id
      And a.参数标识 = User_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getfontparam;

  --17、功能：插入编辑器窗体参数
  Procedure p_Setformparam
  (
    Form_In Nvarchar2,
    User_In Nvarchar2
  ) As
    m_Id     Nvarchar2(36);
    Numcount Int;
  Begin
    Select Rawtohex(Id)
    Into m_Id
    From 影像参数说明
    Where 模块 = 'ImageEditor'
    And 参数名 = '窗口设置';
    Select Count(*)
    Into Numcount
    From 影像参数取值 t
    Where t.参数id = m_Id
    And t.参数标识 = User_In;
    If Numcount > 0 Then
      Update 影像参数取值 a
      Set a.参数值 = Form_In
      Where a.参数标识 = User_In
      And a.参数id = m_Id;
    Else
      Insert Into 影像参数取值 a (Id, 参数id, 参数标识, 参数值) Values (Sys_Guid(), m_Id, User_In, Form_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Setformparam;

  --18、功能：获取编辑器字体参数
  Procedure p_Getformparam
  (
    Val     Out t_Refcur,
    User_In Nvarchar2
  ) As
    m_Id Nvarchar2(36);
  Begin
    Select Rawtohex(Id)
    Into m_Id
    From 影像参数说明
    Where 模块 = 'ImageEditor'
    And 参数名 = '窗口设置';
    Open Val For
      Select a.参数值
      From 影像参数取值 a
      Where a.参数id = m_Id
      And a.参数标识 = User_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getformparam;

  --19、功能：根据图像UID获取检查信息
  Procedure p_Getstudyinfobyimageuid
  (
    Val        Out t_Refcur,
    医嘱id_In  In 影像检查记录.医嘱id%Type,
    图像uid_In In 影像检查图象.图像uid%Type
  ) As
  Begin
    Open Val For
      Select d.检查uid
      From 影像检查图象 a, 影像检查序列 b, 影像检查记录 c, 影像临时序列 d
      Where c.医嘱id = 医嘱id_In
      And a.图像uid = 图像uid_In
      And a.序列uid = b.序列uid
      And b.检查uid = c.检查uid
      And a.序列uid = d.序列uid;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getstudyinfobyimageuid;

  --20、功能：根据检查UID获取FTP信息
  Procedure p_Getftpinfobystudyuid
  (
    Val        Out t_Refcur,
    检查uid_In In 影像检查记录.检查uid%Type
  ) As
  Begin
    Open Val For
      Select d.Ftp用户名 As Ftpuser, d.Ftp密码 As Ftppwd, c.位置一, c.位置二, c.位置三, c.接收日期, d.Ip地址 As Host,
             '/' || d.Ftp目录 || '/' As Root,
             Decode(c.接收日期, Null, '', To_Char(c.接收日期, 'YYYYMMDD') || '/') || c.检查uid As Url
      From 影像检查记录 c, 影像设备目录 d
      Where Decode(c.位置一, Null, c.位置二, c.位置一) = d.设备号(+)
      And c.检查uid = 检查uid_In
      Union All
      Select d.Ftp用户名 As Ftpuser, d.Ftp密码 As Ftppwd, c.位置一, c.位置二, c.位置三, c.接收日期, d.Ip地址 As Host,
             '/' || d.Ftp目录 || '/' As Root,
             Decode(c.接收日期, Null, '', To_Char(c.接收日期, 'YYYYMMDD') || '/') || c.检查uid As Url
      From 影像临时记录 c, 影像设备目录 d
      Where Decode(c.位置一, Null, c.位置二, c.位置一) = d.设备号(+)
      And c.检查uid = 检查uid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getftpinfobystudyuid;

  --21、功能：根据科室ID获取FTP信息
  Procedure p_Getftpinfobydeptid
  (
    Val       Out t_Refcur,
    科室id_In In 影像流程参数.科室id%Type
  ) As
  Begin
    Open Val For
      Select a.设备号, a.Ip地址, a.Ftp用户名, a.Ftp密码
      From 影像设备目录 a, 影像流程参数 b
      Where a.设备号 = b.参数值
      And b.参数名 = '存储设备号'
      And b.科室id = 科室id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getftpinfobydeptid;

  --22、功能：根据医嘱ID获取FTP信息
  Procedure p_Getftpinfobyadvicetid
  (
    Val       Out t_Refcur,
    医嘱id_In In 影像检查记录.医嘱id%Type
  ) As
  Begin
    Open Val For
      Select a.设备号, a.Ip地址, a.Ftp用户名, a.Ftp密码
      From 影像设备目录 a, 影像检查记录 b
      Where b.位置一 = a.设备号(+)
      And b.医嘱id = 医嘱id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getftpinfobyadvicetid;

  --23、功能：获取检查UID
  Procedure p_Getstudyuid
  (
    Val        Out t_Refcur,
    检查uid_In In 影像检查记录.检查uid%Type
  ) As
  Begin
    Open Val For
      Select 检查uid
      From 影像检查记录
      Where 检查uid = 检查uid_In
      Union All
      Select 检查uid
      From 影像临时记录
      Where 检查uid = 检查uid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getstudyuid;

  --24、功能：获取序列UID
  Procedure p_Getseriesuid
  (
    Val        Out t_Refcur,
    序列uid_In In 影像检查序列.序列uid%Type
  ) As
  Begin
    Open Val For
      Select 序列uid
      From 影像检查序列
      Where 序列uid = 序列uid_In
      Union All
      Select 序列uid
      From 影像临时序列
      Where 序列uid = 序列uid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getseriesuid;

  --25、功能：根据设备号获取设备信息
  Procedure p_Getdeviceinfo
  (
    Val       Out t_Refcur,
    设备号_In In 影像设备目录.设备号%Type
  ) As
  Begin
    Open Val For
      Select 设备号, 设备名, '/' || Decode(Ftp目录, Null, '', Ftp目录 || '/') As Url, Ftp用户名, Ftp密码, Ip地址
      From 影像设备目录
      Where 类型 = 1
      And 设备号 = 设备号_In
      And Nvl(状态, 0) = 1;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getdeviceinfo;

  --26、获取医技站存储设备号
  Procedure p_Getdeviceidbyadviceid
  (
    Val       Out t_Refcur,
    医嘱id_In In 病人医嘱发送.医嘱id%Type
  ) As
  Begin
    Open Val For
      Select d.参数值
      From 医技执行房间 a, 病人医嘱发送 b, 影像dicom服务对 c, 影像dicom服务参数 d
      Where a.科室id = b.执行部门id
      And a.执行间 = b.执行间
      And a.检查设备 = c.设备号
      And c.服务功能 = '图像接收'
      And c.服务id = d.服务id
      And d.参数名称 = '存储设备'
      And b.医嘱id = 医嘱id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getdeviceidbyadviceid;
End b_Pacs_Rptpluginoriginal;

/



--影像报告范文管理(---定义部分---)***************************************************
Create Or Replace Package b_PACS_RptPluginCustom Is
  Type t_Refcur Is Ref Cursor;
-- 功    能：该方法只用于演示...
  Procedure Demo1;

end b_PACS_RptPluginCustom ;
/

--影像报告范文管理(---实现部分---)***************************************************
Create Or Replace Package Body b_PACS_RptPluginCustom  Is
-- 功    能：该方法只用于演示...
  Procedure Demo1
  Is
  Begin
      --TODO:
      Null;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Demo1;

End b_PACS_RptPluginCustom;
/


--XWPACS接口包
Create Or Replace Package b_XINWANGInterface Is
  Type t_Refcur Is Ref Cursor;
  -- 1 PACS状态改变信息
  Procedure PacsStatusChange
  (
    状态ID_In   In Number,
    医嘱ID_In   In 影像检查记录.医嘱ID%Type,
    影像类别_In In 影像检查记录.影像类别%Type,
    检查号_In   In 影像检查记录.检查号%Type,
    处理时间    In Date,
    执行人      In Varchar2,
    胶片大小    In Varchar2,
    检查UID_In  In 影像检查记录.检查UID%Type := Null
  );
  -- 2 取消图像关联
  Procedure PacsUnmatchImage(医嘱ID_In In 影像检查记录.医嘱ID%Type);
  -- 3 填写报告图的存储设备
  Procedure PacsSetFTPDeviceNo
  (
    医嘱ID_In In 影像检查记录.医嘱ID%Type,
    设备号_In In 影像检查记录.位置一%Type
  );
  -- 4 更新图像数
  Procedure UpdateImgCount
  (
    医嘱ID_In 影像检查记录.医嘱ID%Type,
    图像数_In Number
  );

End b_XINWANGInterface;

/

Create Or Replace Package Body b_XINWANGInterface Is

  -- 1 PACS状态改变信息
  Procedure PacsStatusChange
  (
    状态ID_In   In Number,
    医嘱ID_In   In 影像检查记录.医嘱ID%Type,
    影像类别_In In 影像检查记录.影像类别%Type,
    检查号_In   In 影像检查记录.检查号%Type,
    处理时间    In Date,
    执行人      In Varchar2,
    胶片大小    In Varchar2,
    检查UID_In  In 影像检查记录.检查UID%Type := Null
  ) Is
    Strsql     Varchar2(2000);
    v_StudyUID 影像检查记录.检查UID%Type;
  
    Cursor c_Advice Is
      Select ID From 病人医嘱记录 Where ID = 医嘱id_In Or (相关id = 医嘱id_In And 诊疗类别 In ('F', 'G', 'D'));
  
  Begin
    --状态ID_In:1-匹配成功;2-匹配失败;3-新检查（收到第一幅图像）;4-收到每一幅图像;
    --     5-删除检查;6-胶片打印成功；7-更新电子胶片状态；8-图像转移到云平台；9-图像从云平台下载
  
    If 检查UID_In Is Null Then
      v_StudyUID := 医嘱ID_In;
    Else
      v_StudyUID := 检查UID_In;
    End If;
  
    If 状态ID_In = 1 Then
      --图象匹配成功
    
      --填写影像检查记录表的 检查UID，接收日期等，但是不填写序列级别的表,检查UID填写检查UID_In（StudyUID）
      Update 影像检查记录
      Set 检查UID = v_StudyUID, 接收日期 = Decode(处理时间, Null, Sysdate, 处理时间), 图像位置 = 1
      Where 医嘱ID = 医嘱ID_In;
    
      --设置医嘱执行状态
      For r_Advice In c_Advice Loop
        Update 病人医嘱发送
        Set 执行状态 = 3, 执行过程 = Decode(Sign(执行过程 - 2), 1, 执行过程, 3)
        Where 医嘱id = r_Advice.id;
      End Loop;
    Elsif 状态ID_In = 2 Then
      Strsql := 'dd';
    Elsif 状态ID_In = 3 Then
      -- 3-新检查（收到第一幅图像），暂时不处理
      Strsql := 'dd';
    Elsif 状态ID_In = 4 Then
      --  4-收到每一幅图像 ，暂时不处理
      Strsql := 'dd';
    Elsif 状态ID_In = 5 Then
      -- 5-删除检查
      -- 删除影像检查记录表中对应的检查UID，接收日期等
      Update 影像检查记录
      Set 检查UID = Null, 位置一 = Null, 位置二 = Null, 位置三 = Null, 报告图象 = Null, 接收日期 = Null
      Where 医嘱ID = 医嘱ID_IN;
    Elsif 状态ID_In = 6 Then
      -- 6-胶片打印成功
      --记录胶片大小，打印人等
    
      --一个医嘱打印一张或者多张胶片的情况，每张胶片调用一过程，相关ID为空
      Insert Into 胶片打印记录
        (ID, 相关id, 医嘱id, 胶片大小, 打印人, 打印时间)
      Values
        (胶片打印记录_Id.Nextval, Null, 医嘱ID_In, 胶片大小, 执行人, Decode(处理时间, Null, Sysdate, 处理时间));
      Update 影像检查记录 Set 是否打印 = 1 Where 医嘱ID = 医嘱ID_In;
    Elsif 状态ID_In = 7 Then
      --更新电子胶片状态
      Update 影像检查记录 Set 是否电子胶片 = 1 Where 医嘱ID = 医嘱ID_In;
    Elsif 状态ID_In = 8 Then
      -- 图像转移到云平台
      Update 影像检查记录 Set 检查UID = 检查UID_In, 图像位置 = 2 Where 医嘱ID = 医嘱ID_In;
    Elsif 状态ID_In = 9 Then
      -- 图像从云平台下载
      Update 影像检查记录 Set 图像位置 = 1 Where 医嘱ID = 医嘱ID_In;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End PacsStatusChange;

  -- 2 PACS图像取消关联
  Procedure PacsUnmatchImage(医嘱ID_In In 影像检查记录.医嘱ID%Type) Is
    v_执行过程 病人医嘱发送.执行过程%Type;
    v_发送号   病人医嘱发送.发送号%Type;
  Begin
    --设置影像检查记录表的状态
    Update 影像检查记录 Set 检查UID = Null, 接收日期 = Null, 图像位置 = Null, 位置一 = Null Where 医嘱ID = 医嘱ID_In;
  
    --调用 Zl_影像检查_State 改变检查过程的状态
    Select 执行过程, 发送号 Into v_执行过程, v_发送号 From 病人医嘱发送 Where 医嘱ID = 医嘱ID_In;
  
    --如果执行过程是3，则将过程修改成2
    If v_执行过程 = 3 Then
      Zl_影像检查_State(医嘱ID_In, v_发送号, 2);
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End PacsUnmatchImage;

  -- 3 填写报告图的存储设备
  Procedure PacsSetFTPDeviceNo
  (
    医嘱ID_In In 影像检查记录.医嘱ID%Type,
    设备号_In In 影像检查记录.位置一%Type
  ) Is
  Begin
    --设置影像检查记录表的状态
    Update 影像检查记录 Set 位置一 = 设备号_In Where 医嘱ID = 医嘱ID_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End PacsSetFTPDeviceNo;

  -- 4 更新图像数量
  Procedure UpdateImgCount
  (
    医嘱ID_IN 影像检查记录.医嘱ID%Type,
    图像数_In Number
  ) Is
  Begin
    Update 影像检查记录 Set 图像数量 = 图像数_In Where 医嘱ID = 医嘱ID_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End UpdateImgCount;

End b_XINWANGInterface;
/

Create Or Replace Package b_Emergency_Rating Is
  --疼痛分级方法
  --入参数量1①描述文本
  --调用形式：b_emergency_rating.Is_Pain_num_rating(4,5)
  --返回值类型：varchar2   返回结果形式 病人等级:分数：描述  例如：2:9:重度疼痛
  Function Is_Pain_Num_Rating(Describe Varchar2) Return Varchar2;
  --昏迷评分分级方法
  --入参数量3②睁眼反应指标id：指标结果描述 ③ 语言反应指标id：指标结果描述 ④活动反应指标id：指标结果描述
  --调用形式： b_emergency_rating.Is_coma_rating('1:声音刺激','2:定向良好','3:痛刺激屈曲')
  --返回值类型：varchar2   返回结果形式 病人等级:总分数：描述  例如：3:11:中度意识障碍
  Function Is_Coma_Rating
  (
    Open_Reaction     Varchar2,
    Language_Reaction Varchar2,
    Activity_Reaction Varchar2
  ) Return Varchar2;
  --判断客观评估为儿童还是成人函数
  Function Is_Judgement_Function
  (
    Agenum  Number,
    Ageunit Varchar2
  ) Return Varchar2;
  --客观评价分级方法成人和儿童方法
  --入参数量3①年龄 ②年龄单位 ③ 指标id：指标结果描述（可多个）
  --调用形式： b_emergency_rating.Is_objective_rating(5,'岁','11:9,6:100,4:20')
  --返回值类型：varchar2   返回结果形式 病人等级 1
  Function Is_Objective_Rating
  (
    Agenum           Number,
    Ageunit          Varchar2,
    Indexid_Describe Varchar2
  ) Return Varchar2;
End b_Emergency_Rating;
/
Create Or Replace Package Body b_Emergency_Rating Is

  Function Is_Pain_Num_Rating(Describe Varchar2) --疼痛等级规则
   Return Varchar2 As
    State_Level  Varchar2(10); --病人级别
    Score_Result Varchar2(100); --评分结果描述
    Score        Number; --分数
  Begin
    Select 指标结果分值 Into Score From 急诊评分方法规则 Where 指标结果描述 = Describe;
  
    Select Min(病情级别), Min(评分结果描述)
    Into State_Level, Score_Result
    From 急诊评分方法分级
    Where 运算符 = 2 And Score > 分值上限 And 方法id = 4 Or 运算符 = 3 And 分值下限 < Score And 方法id = 4 Or
          运算符 = 6 And Score Between 分值下限 And 分值上限 And 方法id = 4 Or 运算符 = 1 And 分值上限 = Score And 方法id = 4 Or
          运算符 = 4 And 分值上限 >= Score And 方法id = 4 Or 运算符 = 5 And Score <= 分值下限 And 方法id = 4;
    Return State_Level || ':' || Score || ':' || Score_Result;
  Exception
    --异常处理语句段
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Is_Pain_Num_Rating;

  Function Is_Coma_Rating( --昏迷等级规则
                          Open_Reaction     Varchar2,
                          Language_Reaction Varchar2,
                          Activity_Reaction Varchar2) Return Varchar2 As
  
    Coma_Score_All Number; --昏迷总分数
    Coma_Level     Varchar2(10); --昏迷等级
    Score_Result   Varchar2(100); --评分结果描述
  
    Coma_Id1    Varchar2(10); --昏迷-睁眼指标ID
    Coma_Text1  Varchar2(100); --昏迷-睁眼描述
    Coma_Score1 Number; --昏迷-睁眼分数
  
    Coma_Id2    Varchar2(10); --昏迷-语言指标ID
    Coma_Text2  Varchar2(100); --昏迷-语言描述
    Coma_Score2 Number; --昏迷-语言分数
  
    Coma_Id3    Varchar2(10); --昏迷-活动指标ID
    Coma_Text3  Varchar2(100); --昏迷-活动描述
    Coma_Score3 Number; --昏迷-活动分数
  Begin
    Select C1, C2 Into Coma_Id1, Coma_Text1 From Table(f_Str2list2(Open_Reaction));
    Select C1, C2 Into Coma_Id2, Coma_Text2 From Table(f_Str2list2(Language_Reaction));
    Select C1, C2 Into Coma_Id3, Coma_Text3 From Table(f_Str2list2(Activity_Reaction));
  
    Select 指标结果分值
    Into Coma_Score1
    From 急诊评分方法规则
    Where 方法id = 3 And 指标结果描述 = Coma_Text1 And 指标id = Coma_Id1;
  
    Select 指标结果分值
    Into Coma_Score2
    From 急诊评分方法规则
    Where 方法id = 3 And 指标结果描述 = Coma_Text2 And 指标id = Coma_Id2;
  
    Select 指标结果分值
    Into Coma_Score3
    From 急诊评分方法规则
    Where 方法id = 3 And 指标结果描述 = Coma_Text3 And 指标id = Coma_Id3;
    Coma_Score_All := Coma_Score1 + Coma_Score2 + Coma_Score3;
  
    Select Min(病情级别), Min(评分结果描述)
    Into Coma_Level, Score_Result
    From 急诊评分方法分级
    Where 运算符 = 2 And Coma_Score_All > 分值上限 And 方法id = 3 Or 运算符 = 3 And 分值下限 < Coma_Score_All And 方法id = 3 Or
          运算符 = 6 And Coma_Score_All Between 分值下限 And 分值上限 And 方法id = 3 Or
          运算符 = 1 And 分值上限 = Coma_Score_All And 方法id = 3 Or 运算符 = 4 And 分值上限 >= Coma_Score_All And 方法id = 3 Or
          运算符 = 5 And Coma_Score_All <= 分值下限 And 方法id = 3;
    Return Coma_Level || ':' || Coma_Score_All || ':' || Score_Result;
  
  Exception
    --异常处理语句段
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Is_Coma_Rating;

  Function Is_Judgement_Function
  (
    Agenum  Number,
    Ageunit Varchar2
  ) Return Varchar2 As
    --判断成人或儿童规则
    Children_Age Varchar2(100);
  Begin
    If Ageunit Is Null Then
      Return '1'; --成人
    End If;
  
    If Ageunit = '岁' Then
      Select 参数值 Into Children_Age From zlParameters Where 参数名 = '儿童年龄界定上限';
    
      If Agenum <= To_Number(Children_Age) Then
        Return '2'; --儿童
      Else
        Return '1'; --成人
      End If;
    End If;
    Return '2'; --年龄单位不为岁返回儿童0-1岁
  Exception
    --异常处理语句段
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Is_Judgement_Function;

  Function Is_Objective_Rating --客观评价评分规则
  (
    Agenum           Number,
    Ageunit          Varchar2,
    Indexid_Describe Varchar2
  ) Return Varchar2 As
    Person        Varchar2(2); --儿童或者成人
    o_Indexid     t_Numlist; --指标ID
    o_Describe    t_Numlist; --传入指标参数
    Level_Max     Number; --病情最大值
    Illness_Level Number; --病情级别
    Age_Id        Number; --儿童年龄id
  Begin
    Select b_Emergency_Rating.Is_Judgement_Function(Agenum, Ageunit) Into Person From Dual;
    If Person = '1' Then
      --成人的规则
      Select Max(病情级别) Into Level_Max From 急诊评分方法规则;
      Select C1, C2 Bulk Collect Into o_Indexid, o_Describe From Table(f_Str2list2(Indexid_Describe));
      For I In 1 .. o_Indexid.Count Loop
        Select Min(病情级别)
        Into Illness_Level
        From 急诊评分方法规则
        Where 运算符 = 2 And o_Describe(I) > 指标值上限 And 指标id = o_Indexid(I) And 方法id = 1 Or
              运算符 = 3 And o_Describe(I) < 指标值下限 And 方法id = 1 And 指标id = o_Indexid(I) Or
              运算符 = 6 And o_Describe(I) >= 指标值下限 And o_Describe(I) <= 指标值上限 And 方法id = 1 And 指标id = o_Indexid(I) Or
              运算符 = 1 And 指标值上限 = o_Describe(I) And 方法id = 1 And 指标id = o_Indexid(I) Or
              运算符 = 4 And o_Describe(I) >= 指标值上限 And 方法id = 1 And 指标id = o_Indexid(I) Or
              运算符 = 5 And o_Describe(I) <= 指标值下限 And 方法id = 1 And 指标id = o_Indexid(I);
        If Illness_Level < Level_Max Then
          Level_Max := Illness_Level;
        End If;
      End Loop;
      Return Level_Max;
    End If;
  
    If Person = '2' Then
      --儿童规则
      Select Max(病情级别) Into Level_Max From 急诊评分方法规则;
      --程序逻辑根据传入年龄和单位和指标id找到相应的年龄id，根据年龄id和指标id和指标值找到级别
      --如果没有找到，抛弃年龄条件寻找没有年龄值的级别，如果级别还为空将找到的最小级别赋给它
      If Ageunit = '岁' Then
        Select C1, C2 Bulk Collect Into o_Indexid, o_Describe From Table(f_Str2list2(Indexid_Describe));
        For I In 1 .. o_Indexid.Count Loop
          Select Max(ID)
          Into Age_Id
          From 急诊评分指标年龄
          Where 运算符 = 2 And Agenum > 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '岁' Or
                运算符 = 3 And Agenum < 年龄下限 And 指标id = o_Indexid(I) And 年龄单位 = '岁' And 年龄单位 = '岁' Or
                运算符 = 6 And Agenum Between 年龄下限 And 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '岁' Or
                运算符 = 1 And 年龄上限 = Agenum And 指标id = o_Indexid(I) And 年龄单位 = '岁' Or
                运算符 = 4 And Agenum >= 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '岁' Or
                运算符 = 5 And Agenum <= 年龄下限 And 指标id = o_Indexid(I) And 年龄单位 = '岁';
        
          Select Min(病情级别)
          Into Illness_Level
          From 急诊评分方法规则
          Where 运算符 = 2 And o_Describe(I) > 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 3 And o_Describe(I) < 指标值下限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 6 And o_Describe(I) Between 指标值下限 And 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And
                方法id = 2 Or 运算符 = 1 And 指标值上限 = o_Describe(I) And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 4 And o_Describe(I) >= 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 5 And o_Describe(I) <= 指标值下限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2;
          If Illness_Level Is Null Then
            Select Nvl(Min(病情级别), Level_Max)
            Into Illness_Level
            From 急诊评分方法规则
            Where 运算符 = 2 And o_Describe(I) > 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 3 And o_Describe(I) < 指标值下限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 6 And o_Describe(I) Between 指标值下限 And 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And
                  指标年龄id Is Null Or
                  运算符 = 1 And 指标值上限 = o_Describe(I) And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 4 And o_Describe(I) >= 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 5 And o_Describe(I) <= 指标值下限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null;
          
          End If;
          If Illness_Level < Level_Max Then
            Level_Max := Illness_Level;
          End If;
        
        End Loop;
        Return Level_Max;
      End If;
    
      If Ageunit = '月' Then
        Select C1, C2 Bulk Collect Into o_Indexid, o_Describe From Table(f_Str2list2(Indexid_Describe));
        For I In 1 .. o_Indexid.Count Loop
          Select Max(ID)
          Into Age_Id
          From 急诊评分指标年龄
          Where 运算符 = 2 And Agenum > 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '月' Or
                运算符 = 3 And Agenum < 年龄下限 And 指标id = o_Indexid(I) And 年龄单位 = '月' And 年龄单位 = '月' Or
                运算符 = 6 And Agenum Between 年龄下限 And 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '月' Or
                运算符 = 1 And 年龄上限 = Agenum And 指标id = o_Indexid(I) And 年龄单位 = '月' Or
                运算符 = 4 And Agenum >= 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '月' Or
                运算符 = 5 And Agenum <= 年龄下限 And 指标id = o_Indexid(I) And 年龄单位 = '月';
          Select Min(病情级别)
          Into Illness_Level
          From 急诊评分方法规则
          Where 运算符 = 2 And o_Describe(I) > 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 3 And o_Describe(I) < 指标值下限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 6 And o_Describe(I) Between 指标值下限 And 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And
                方法id = 2 Or 运算符 = 1 And 指标值上限 = o_Describe(I) And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 4 And o_Describe(I) >= 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 5 And o_Describe(I) <= 指标值下限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2;
        
          If Illness_Level Is Null Then
            Select Nvl(Min(病情级别), Level_Max)
            Into Illness_Level
            From 急诊评分方法规则
            Where 运算符 = 2 And o_Describe(I) > 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 3 And o_Describe(I) < 指标值下限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 6 And o_Describe(I) Between 指标值下限 And 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And
                  指标年龄id Is Null Or
                  运算符 = 1 And 指标值上限 = o_Describe(I) And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 4 And o_Describe(I) >= 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 5 And o_Describe(I) <= 指标值下限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null;
          
          End If;
          If Illness_Level < Level_Max Then
            Level_Max := Illness_Level;
          End If;
        
        End Loop;
        Return Level_Max;
      End If;
    
      If Ageunit = '天' Then
        Select C1, C2 Bulk Collect Into o_Indexid, o_Describe From Table(f_Str2list2(Indexid_Describe));
        For I In 1 .. o_Indexid.Count Loop
          Select Max(ID)
          Into Age_Id
          From 急诊评分指标年龄
          Where 运算符 = 2 And Agenum > 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '天' Or
                运算符 = 3 And Agenum < 年龄下限 And 指标id = o_Indexid(I) And 年龄单位 = '天' And 年龄单位 = '天' Or
                运算符 = 6 And Agenum Between 年龄下限 And 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '天' Or
                运算符 = 1 And 年龄上限 = Agenum And 指标id = o_Indexid(I) And 年龄单位 = '天' Or
                运算符 = 4 And Agenum >= 年龄上限 And 指标id = o_Indexid(I) And 年龄单位 = '天' Or
                运算符 = 5 And Agenum <= 年龄下限 And 指标id = o_Indexid(I) And 年龄单位 = '天';
          Select Min(病情级别)
          Into Illness_Level
          From 急诊评分方法规则
          Where 运算符 = 2 And o_Describe(I) > 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 3 And o_Describe(I) < 指标值下限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 6 And o_Describe(I) Between 指标值下限 And 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And
                方法id = 2 Or 运算符 = 1 And 指标值上限 = o_Describe(I) And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 4 And o_Describe(I) >= 指标值上限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2 Or
                运算符 = 5 And o_Describe(I) <= 指标值下限 And 指标id = o_Indexid(I) And 指标年龄id = Age_Id And 方法id = 2;
          If Illness_Level Is Null Then
            Select Nvl(Min(病情级别), Level_Max)
            Into Illness_Level
            From 急诊评分方法规则
            Where 运算符 = 2 And o_Describe(I) > 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 3 And o_Describe(I) < 指标值下限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 6 And o_Describe(I) Between 指标值下限 And 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And
                  指标年龄id Is Null Or
                  运算符 = 1 And 指标值上限 = o_Describe(I) And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 4 And o_Describe(I) >= 指标值上限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null Or
                  运算符 = 5 And o_Describe(I) <= 指标值下限 And 指标id = o_Indexid(I) And 方法id = 2 And 指标年龄id Is Null;
          End If;
        
          If Illness_Level < Level_Max Then
            Level_Max := Illness_Level;
          End If;
        
        End Loop;
        Return Level_Max;
      End If;
    
    End If;
    Return Level_Max;
  Exception
    --异常处理语句段
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Is_Objective_Rating;

End b_Emergency_Rating;
/

Create Or Replace Package Pkg_Pretriage_Dql As
  -----------------------------------------------------
  --获取国籍基础数据
  -----------------------------------------------------
  Procedure Get_Nationality
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取报表列表
  -----------------------------------------------------
  Procedure Get_Reportlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --检查身份证录入是否正确
  -----------------------------------------------------
  Procedure Checkidcard
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );

  -----------------------------------------------------
  --检查年龄录入是否正确
  -----------------------------------------------------
  Procedure Checkage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );
  -----------------------------------------------------
  --通过医保号读取病人信息列表清单
  -----------------------------------------------------
  Procedure Get_Patlistbymedical
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --通过身份证号读取病人信息列表清单
  -----------------------------------------------------
  Procedure Get_Patlistbyidcard
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --通过输入姓名匹配病人信息列表清单
  -----------------------------------------------------
  Procedure Get_Patlistbyname
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取最新的就诊状态
  -----------------------------------------------------
  Procedure Getvisitstate
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取病人分诊评分信息
  -----------------------------------------------------
  Procedure Load_Levelinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取病人分诊指标信息
  -----------------------------------------------------
  Procedure Load_Rulesinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取病人分诊记录内容
  -----------------------------------------------------
  Procedure Load_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取单个病人分诊信息
  -----------------------------------------------------
  Procedure Get_Patidetail
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --获取病人列表清单
  -----------------------------------------------------
  Procedure Get_Patlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --疼痛分级方法
  -----------------------------------------------------
  Procedure Get_Pain_Num_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --昏迷评分分级方法
  -----------------------------------------------------
  Procedure Get_Coma_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --客观评价分级方法成人和儿童方法
  -----------------------------------------------------
  Procedure Get_Objective_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取儿童年龄上限
  -----------------------------------------------------
  Procedure Get_Childmaxage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );
  -----------------------------------------------------
  --获取急诊等级
  -----------------------------------------------------
  Procedure Get_Level
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取急诊科室
  -----------------------------------------------------
  Procedure Get_Dept
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取人工评估规则
  -----------------------------------------------------
  Procedure Get_Rules
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --根据出生日期返回年龄
  -----------------------------------------------------
  Procedure Get_Datetoage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );
  -----------------------------------------------------
  --获取性别基础数据
  -----------------------------------------------------
  Procedure Get_Sexbase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取民族基础数据
  -----------------------------------------------------
  Procedure Get_Nationbase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取急诊评分指标
  -----------------------------------------------------
  Procedure Get_Scorebase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --获取急诊主诉
  -----------------------------------------------------
  Procedure Get_Paticc
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取病人来源
  -----------------------------------------------------
  Procedure Get_Patifrom
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取急诊意识状态
  -----------------------------------------------------
  Procedure Get_Patistate
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取急诊陪同人员
  -----------------------------------------------------
  Procedure Get_Entourage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取急诊常见既往史
  -----------------------------------------------------
  Procedure Get_Dishistory
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --获取数据库系统时间
  -----------------------------------------------------
  Procedure Get_Now_Time
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );

End Pkg_Pretriage_Dql;
/
Create Or Replace Package Body Pkg_Pretriage_Dql As
  -----------------------------------------------------
  --获取国籍基础数据
  -----------------------------------------------------
  Procedure Get_Nationality
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 编码, 名称, 简码, Nvl(缺省标志, 0) As 缺省 From 国籍 Order By Nvl(缺省标志, 0) Desc, 编码;
  End Get_Nationality;

  -----------------------------------------------------
  --获取报表列表
  -----------------------------------------------------
  Procedure Get_Reportlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 标志, 系统, 编号, 名称, Nvl(是否停用, 0) 是否停用
      From (Select 1 As 标志, a.系统, a.编号, a.名称, a.是否停用
             From zlReports A, zlPrograms B
             Where a.系统 = b.系统 And a.程序id = b.序号 And Not Upper(a.编号) Like '%BILL%' And Upper(b.部件) <> Upper('zl9Report') And
                   b.系统 = 100 And b.序号 = 1244
             Union All
             Select Decode(a.系统, Null, 2, 1) As 标志, a.系统, a.编号, a.名称, a.是否停用
             From zlReports A, zlRPTPuts B, zlPrograms C
             Where a.Id = b.报表id And b.系统 = c.系统 And b.程序id = c.序号 And (Not Upper(a.编号) Like '%BILL%' Or a.系统 Is Null) And
                   c.系统 = 100 And c.序号 = 1244)
      Where Instr(',ZL1_REPORT_1244_1,ZL1_REPORT_1244_2,', ',' || 编号 || ',') = 0 And Nvl(是否停用, 0) = 0
      Order By 标志, 编号;
  End Get_Reportlist;

  -----------------------------------------------------
  --检查身份证录入是否正确
  -----------------------------------------------------
  Procedure Checkidcard
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
    --入参：Json_In:格式
    --  input
    --    idcard        C 1 录入的身份证号
    -- 返回值：固定格式XML串
    --<OUTPUT>
    --       <BIRTHDAY></BIRTHDAY>                //出生日期
    --       <SEX></SEX>                  //性别
    --       <AGE></AGE>                //年龄
    --     <MSG></MSG>         //空串-身份证号有效(可从身份证号中获取出生日期和性别)，非空串-返回错误信息
    --</OUTPUT>
  
    Jsonobj  Pljson;
    j_In     Pljson;
    v_录入项 Varchar2(50);
  Begin
    j_In       := Pljson(Input_In);
    Jsonobj    := j_In.Get_Pljson('input');
    v_录入项   := Jsonobj.Get_String('idcard');
    Output_Out := Zl_Fun_Checkidcard(v_录入项);
  End Checkidcard;

  -----------------------------------------------------
  --检查年龄录入是否正确
  -----------------------------------------------------
  Procedure Checkage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
    --入参：Json_In:格式
    --  input
    --    age        C 1 年龄
    Jsonobj Pljson;
    j_In    Pljson;
    v_年龄  Varchar2(50);
  Begin
    j_In       := Pljson(Input_In);
    Jsonobj    := j_In.Get_Pljson('input');
    v_年龄     := Jsonobj.Get_String('age');
    Output_Out := Zl_Age_Check(v_年龄);
  End Checkage;

  -----------------------------------------------------
  --通过医保号读取病人信息列表清单
  -----------------------------------------------------
  Procedure Get_Patlistbymedical
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    Jsonobj    Pljson;
    j_In       Pljson;
    v_医保号   Varchar2(200);
    v_医保类型 Varchar2(200);
  Begin
    j_In       := Pljson(Input_In);
    Jsonobj    := j_In.Get_Pljson('input');
    v_医保号   := Jsonobj.Get_String('医保号');
    v_医保类型 := Jsonobj.Get_String('医保类型');
    Open Output_Out For
      Select /*+Rule */
      Distinct a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, To_Char(a.出生日期, 'yyyy-MM-dd') As 出生日期, a.国籍, a.民族, a.身份证号, a.手机号, a.医保号,
               b.名称 As 保险类别, a.家庭地址, Nvl(就诊时间, 登记时间) As 就诊时间
      From 病人信息 A, 保险类别 B
      Where a.险类 = b.序号(+) And a.停用时间 Is Null And a.医保号 = v_医保号 And b.名称 = v_医保类型
      Order By 病人id Desc;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Patlistbymedical;

  -----------------------------------------------------
  --通过身份证号读取病人信息列表清单
  -----------------------------------------------------
  Procedure Get_Patlistbyidcard
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    Jsonobj    Pljson;
    j_In    Pljson;
    v_身份证号 Varchar2(200);
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    v_身份证号 := Jsonobj.Get_String('身份证号');
    Open Output_Out For
      Select /*+Rule */
      Distinct a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, To_Char(a.出生日期, 'yyyy-MM-dd') As 出生日期, a.国籍, a.民族, a.身份证号, a.手机号, a.医保号,
               b.名称 As 保险类别, a.家庭地址, Nvl(就诊时间, 登记时间) As 就诊时间
      From 病人信息 A, 保险类别 B
      Where a.险类 = b.序号(+) And a.停用时间 Is Null And a.身份证号 = v_身份证号
      Order By 病人id Desc;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Patlistbyidcard;

  -----------------------------------------------------
  --通过输入姓名匹配病人信息列表清单
  -----------------------------------------------------
  Procedure Get_Patlistbyname
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    Jsonobj    Pljson;
    j_In    Pljson;
    v_姓名     Varchar2(200);
    n_录入项   Varchar2(200);
    n_精确查找 Number;
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    v_姓名  := Jsonobj.Get_String('姓名输入');
  
    If v_姓名 Is Not Null Then
      If Substr(v_姓名, 1, 1) = '-' Then
        n_精确查找 := 1;
        v_姓名     := Substr(v_姓名, 2);
        n_录入项   := Zl_To_Number(v_姓名);
      End If;
    End If;
    If n_精确查找 = 1 Then
      Open Output_Out For
        Select /*+Rule */
         1 As 排序id, a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, To_Char(a.出生日期, 'yyyy-MM-dd') As 出生日期, a.国籍, a.民族, a.身份证号, a.手机号,
         a.医保号, b.名称 As 保险类别, a.家庭地址, Nvl(就诊时间, 登记时间) As 就诊时间
        From 病人信息 A, 保险类别 B
        Where a.险类 = b.序号(+) And a.停用时间 Is Null And a.身份证号 = v_姓名
        Union All
        Select 1 /*+Rule */ As 排序id, a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, To_Char(a.出生日期, 'yyyy-MM-dd') As 出生日期, a.国籍, a.民族,
               a.身份证号, a.手机号, a.医保号, b.名称 As 保险类别, a.家庭地址, Nvl(就诊时间, 登记时间) As 就诊时间
        From 病人信息 A, 保险类别 B
        Where a.险类 = b.序号(+) And a.停用时间 Is Null And a.病人id = n_录入项
        Union All
        Select /*+Rule */
         1 As 排序id, a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, To_Char(a.出生日期, 'yyyy-MM-dd') As 出生日期, a.国籍, a.民族, a.身份证号, a.手机号,
         a.医保号, b.名称 As 保险类别, a.家庭地址, Nvl(就诊时间, 登记时间) As 就诊时间
        From 病人信息 A, 保险类别 B
        Where a.险类 = b.序号(+) And a.停用时间 Is Null And a.门诊号 = n_录入项
        Union All
        Select /*+Rule */
         1 As 排序id, a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, To_Char(a.出生日期, 'yyyy-MM-dd') As 出生日期, a.国籍, a.民族, a.身份证号, a.手机号,
         a.医保号, b.名称 As 保险类别, a.家庭地址, Nvl(就诊时间, 登记时间) As 就诊时间
        From 病人信息 A, 保险类别 B
        Where a.险类 = b.序号(+) And a.停用时间 Is Null And a.住院号 = n_录入项
        Union All
        Select /*+Rule */
         1 As 排序id, a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, To_Char(a.出生日期, 'yyyy-MM-dd') As 出生日期, a.国籍, a.民族, a.身份证号, a.手机号,
         a.医保号, b.名称 As 保险类别, a.家庭地址, Nvl(就诊时间, 登记时间) As 就诊时间
        From 病人信息 A, 保险类别 B
        Where a.险类 = b.序号(+) And a.停用时间 Is Null And a.医保号 = v_姓名
        Union All
        Select 0 As 排序id, -null, '[新病人]', Null, Null, -null, Null, Null, Null, Null, Null, Null, Null, Null,
               To_Date(Null)
        From Dual
        Order By 排序id;
    Else
      Open Output_Out For
        Select 1 As 排序id, a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, To_Char(a.出生日期, 'yyyy-MM-dd') As 出生日期, a.国籍, a.民族, a.身份证号,
               a.手机号, a.医保号, b.名称 As 保险类别, a.家庭地址, Nvl(就诊时间, 登记时间) As 就诊时间
        From 病人信息 A, 保险类别 B
        Where a.险类 = b.序号(+) And a.停用时间 Is Null And (a.身份证号 = v_姓名)
        Union All
        Select 1 As 排序id, a.*
        From (Select /*+Rule */
               Distinct a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, To_Char(a.出生日期, 'yyyy-MM-dd') As 出生日期, a.国籍, a.民族, a.身份证号, a.手机号,
                        a.医保号, b.名称 As 保险类别, a.家庭地址, Nvl(就诊时间, 登记时间) As 就诊时间
               From 病人信息 A, 保险类别 B
               Where a.险类 = b.序号(+) And a.停用时间 Is Null And a.姓名 Like v_姓名 || '%'
               Order By 就诊时间 Desc) A
        Where Rownum < 101
        Union All
        Select 0 As 排序id, -null, '[新病人]', Null, Null, -null, Null, Null, Null, Null, Null, Null, Null, Null,
               To_Date(Null)
        From Dual
        Order By 排序id;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Patlistbyname;

  -----------------------------------------------------
  --获取最新的就诊状态
  -----------------------------------------------------
  Procedure Getvisitstate
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj Pljson;
    j_In    Pljson;
    n_Id    急诊分诊记录.Id%Type; --分诊ID
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    n_Id    := To_Number(Jsonobj.Get_String('id'));
    Open Output_Out For
      Select Max(Decode(Nvl(c.执行状态, 0), 0, 0, 1)) As 就诊状态
      From 急诊就诊记录 A, 急诊分诊记录 B, 病人挂号记录 C
      Where a.Id = b.就诊id And a.挂号id = c.Id(+) And b.Id = n_Id;
  End Getvisitstate;

  -----------------------------------------------------
  --获取病人分诊评分信息
  -----------------------------------------------------
  Procedure Load_Levelinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj Pljson;
    j_In    Pljson;
    n_Id    急诊分诊记录.Id%Type; --分诊ID
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    n_Id    := To_Number(Jsonobj.Get_String('id'));
    Open Output_Out For
      Select ID, 分诊id, 方法id, 评分方法分值, 评分结果描述, 病情级别 From 急诊病人评分 Where 分诊id = n_Id;
  End Load_Levelinfo;

  -----------------------------------------------------
  --获取病人分诊指标信息
  -----------------------------------------------------
  Procedure Load_Rulesinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj Pljson;
    j_In    Pljson;
    n_Id    急诊分诊记录.Id%Type; --分诊ID
  Begin
    j_In := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    n_Id    := To_Number(Jsonobj.Get_String('id'));
    Open Output_Out For
      Select a.评分id, b.方法id, a.指标id, a.指标结果文本
      From 急诊病人评分指标 A, 急诊病人评分 B
      Where a.评分id = b.Id And b.分诊id = n_Id;
  End Load_Rulesinfo;

  -----------------------------------------------------
  --获取病人分诊记录内容
  -----------------------------------------------------
  Procedure Load_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj Pljson;
    j_In    Pljson;
    n_Id    急诊分诊记录.Id%Type; --分诊ID
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    n_Id    := To_Number(Jsonobj.Get_String('id'));
    Open Output_Out For
      Select a.病人id, Nvl(e.姓名, a.姓名) As 姓名, Nvl(e.性别, a.性别) As 性别, Nvl(e.年龄, a.年龄) As 年龄,
             To_Char(a.出生日期, 'yyyy-MM-dd') As 出生日期, a.身份证号, a.国籍, a.民族, a.家庭地址, d.名称 As 保险类别, a.医保号, a.手机号,
             b. ID As 就诊id, b. 病人id, b. 病人年龄, b. 年龄数值, b. 年龄单位, b. 挂号id, b. 病情级别,
             To_Char(b. 到院时间, 'yyyy-MM-dd HH24:mi') As 到院时间, b. 主诉, b. 是否三无人员, b. 陪同人员, b. 病人来源, b. 既往病史, b. 意识状态,
             b. 是否成批就诊, b. 成批就诊人数, b. 是否复合伤, b. 备注, b. 登记人 As 就诊登记人, b. 登记时间 As 就诊登记时间, c.修改说明, c.Id As 分诊id, c.分诊次数,
             c.自动病情级别, c.分诊科室id, c.分诊科室名称, c.收缩压, c.舒张压, c.心率, c.指氧饱和度, c.体温, c.血糖, c.血钾,
             To_Char(c.体征测量时间, 'yyyy-MM-dd HH24:mi') As 体征测量时间, c.登记人, c.登记时间, c.人工病情级别, c.人工评级说明, c.呼吸频率, b. 是否绿色通道
      From 病人信息 A, 急诊就诊记录 B, 急诊分诊记录 C, 保险类别 D, 病人挂号记录 E
      Where a.病人id = b.病人id And b.Id = c.就诊id And b.挂号id = e.Id(+) And a.险类 = d.序号(+) And c.Id = n_Id;
  End Load_Pretriage;

  -----------------------------------------------------
  --获取单个病人分诊信息
  -----------------------------------------------------
  Procedure Get_Patidetail
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    Jsonobj   Pljson;
    j_In    Pljson;
    n_Id      急诊分诊记录.就诊id%Type;
    n_Max序号 急诊分诊记录.分诊次数%Type;
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    n_Id    := To_Number(Jsonobj.Get_String('id'));
  
    Select Max(分诊次数) Into n_Max序号 From 急诊分诊记录 Where 就诊id = n_Id;
  
    Open Output_Out For
      Select a.Id 分诊id, a.分诊次数, a.自动病情级别 || '级' As 自动病情级别, a.人工病情级别 || '级' As 人工病情级别,
             '第' || a.分诊次数 || '次分诊    自动评级（' || a.自动病情级别 || '级）' ||
              Decode(a.人工病情级别, '', '', '    人工评级（' || a.人工病情级别 || '级）') ||
              Decode(n_Max序号, a.分诊次数,
                     Decode(Nvl(b.病情级别 || '', '0'), Nvl(b.分诊病情级别 || '', '0'), '',
                             '    修订病情级别（' || Nvl(b.病情级别 || '', '0') || '级）')) || '    分诊时间：' ||
              To_Char(a.登记时间, 'yyyy-MM-dd HH24:mi') As 病情情况
      From 急诊分诊记录 A, 急诊就诊记录 B
      Where a.就诊id = b.Id And 就诊id = n_Id
      Order By 分诊次数 Desc;
  End Get_Patidetail;

  -----------------------------------------------------
  --获取病人列表清单
  -----------------------------------------------------
  Procedure Get_Patlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    Jsonobj    Pljson;
    j_In    Pljson;
    d_开始时间 急诊就诊记录.登记时间%Type;
    d_结束时间 急诊就诊记录.登记时间%Type;
    v_分诊状态 Varchar2(10);
    n_已超时   Number(2); -- =1 仅过滤已超时病人
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    d_开始时间 := To_Date(Jsonobj.Get_String('begin'), 'yyyy-mm-dd hh24:mi:ss');
    d_结束时间 := To_Date(Jsonobj.Get_String('end'), 'yyyy-mm-dd hh24:mi:ss');
    v_分诊状态 := Jsonobj.Get_String('state');
    n_已超时   := Nvl(To_Number(Jsonobj.Get_String('timeout')), 0);
  
    If n_已超时 = 1 Then
    
      Open Output_Out For
        Select b.病人id, b.Id 就诊序号, Nvl(d.姓名, e.姓名) As 姓名, Nvl(d.性别, e.性别) As 性别, Nvl(d.年龄, e.年龄) As 年龄,
               To_Char(b.登记时间, 'yyyy-MM-dd HH24:mi') As 登记时间, b.登记人 分诊护士, b.病情级别 || '级' As 病情级别,
               Decode(Nvl(d.执行状态, 0), 0, 0, 1) As 就诊状态, b.是否绿色通道
        From 急诊就诊记录 B, 急诊病情级别 C, 病人挂号记录 D, 病人信息 E
        Where b.病人id = e.病人id And b.挂号id = d.Id(+) And b.病情级别 = c.序号 And b.登记时间 >= d_开始时间 And b.登记时间 < d_结束时间 And
              Decode(Nvl(d.执行状态, 0), 0, 0, 1) In
              (Select Column_Value From Table(Cast(f_Str2list(v_分诊状态) As t_Strlist))) And
              (c.再次评估时限 Is Not Null And (b.登记时间 + (Nvl(c.再次评估时限, 0) / 24 / 60)) < Sysdate);
    Else
      Open Output_Out For
        Select b.病人id, b.Id 就诊序号, Nvl(d.姓名, e.姓名) As 姓名, Nvl(d.性别, e.性别) As 性别, Nvl(d.年龄, e.年龄) As 年龄,
               To_Char(b.登记时间, 'yyyy-MM-dd HH24:mi') As 登记时间, b.登记人 分诊护士, b.病情级别 || '级' As 病情级别,
               Decode(Nvl(d.执行状态, 0), 0, 0, 1) As 就诊状态, b.是否绿色通道
        From 急诊就诊记录 B, 病人挂号记录 D, 病人信息 E
        Where b.病人id = e.病人id And b.挂号id = d.Id(+) And b.登记时间 >= d_开始时间 And b.登记时间 < d_结束时间 And
              Decode(Nvl(d.执行状态, 0), 0, 0, 1) In
              (Select Column_Value From Table(Cast(f_Str2list(v_分诊状态) As t_Strlist)));
    End If;
  End Get_Patlist;
  -----------------------------------------------------
  --疼痛分级方法
  -----------------------------------------------------
  Procedure Get_Pain_Num_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    --  input
    --    pain        C 1 疼痛等级
    Jsonobj    Pljson;
    j_In    Pljson;
    v_疼痛等级 Varchar2(200);
    v_Out      Varchar2(200);
    v_病人等级 Varchar2(200);
    v_病人分数 Varchar2(200);
    v_描述     Varchar2(200);
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    v_疼痛等级 := Jsonobj.Get_String('pain');
  
    Select b_Emergency_Rating.Is_Pain_Num_Rating(v_疼痛等级) Into v_Out From Dual;
  
    v_病人等级 := Substr(v_Out, 1, Instr(v_Out, ':', 1, 1) - 1);
    v_病人分数 := Substr(v_Out, Instr(v_Out, ':', 1, 1) + 1, Instr(v_Out, ':', 1, 2) - Instr(v_Out, ':', 1, 1) - 1);
    v_描述     := Substr(v_Out, Instr(v_Out, ':', 1, 2) + 1);
  
    Open Output_Out For
      Select v_病人等级 As 病人等级, v_病人分数 As 病人分数, v_描述 As 描述 From Dual;
  End Get_Pain_Num_Rating;

  -----------------------------------------------------
  --昏迷评分分级方法
  -----------------------------------------------------
  Procedure Get_Coma_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    --  input
    --    open_reaction            C 1 睁眼反应指标id：指标结果描述
    --    language_reaction        C 1 语言反应指标id：指标结果描述
    --    activity_reaction        C 1 活动反应指标id：指标结果描述
    Jsonobj    Pljson;
    j_In    Pljson;
    v_睁眼反应 Varchar2(200);
    v_语言反应 Varchar2(200);
    v_活动反应 Varchar2(200);
    v_Out      Varchar2(200);
  
    v_病人等级 Varchar2(200);
    v_病人分数 Varchar2(200);
    v_描述     Varchar2(200);
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    v_睁眼反应 := Jsonobj.Get_String('open_reaction');
    v_语言反应 := Jsonobj.Get_String('language_reaction');
    v_活动反应 := Jsonobj.Get_String('activity_reaction');
  
    Select b_Emergency_Rating.Is_Coma_Rating(v_睁眼反应, v_语言反应, v_活动反应) Into v_Out From Dual;
  
    v_病人等级 := Substr(v_Out, 1, Instr(v_Out, ':', 1, 1) - 1);
    v_病人分数 := Substr(v_Out, Instr(v_Out, ':', 1, 1) + 1, Instr(v_Out, ':', 1, 2) - Instr(v_Out, ':', 1, 1) - 1);
    v_描述     := Substr(v_Out, Instr(v_Out, ':', 1, 2) + 1);
  
    Open Output_Out For
      Select v_病人等级 As 病人等级, v_病人分数 As 病人分数, v_描述 As 描述 From Dual;
  End Get_Coma_Rating;

  -----------------------------------------------------
  --客观评价分级方法成人和儿童方法
  -----------------------------------------------------
  Procedure Get_Objective_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    --  input
    --    agenum                   C 1 年龄
    --    ageunit                  C 1 年龄单位
    --    indexid_describe         C 1 指标id：指标结果描述（可多个）
    Jsonobj    Pljson;
    j_In    Pljson;
    n_年龄     Number;
    v_年龄单位 Varchar2(200);
    v_指标信息 Varchar2(200);
    v_Out      Varchar2(200);
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    n_年龄     := Nvl(Zl_To_Number(Jsonobj.Get_String('agenum')), 0);
    v_年龄单位 := Jsonobj.Get_String('ageunit');
    v_指标信息 := Jsonobj.Get_String('indexid_describe');
  
    Select b_Emergency_Rating.Is_Objective_Rating(n_年龄, v_年龄单位, v_指标信息) Into v_Out From Dual;
  
    Open Output_Out For
      Select v_Out As 病人等级 From Dual;
  End Get_Objective_Rating;
  -----------------------------------------------------
  --获取儿童年龄上限
  -----------------------------------------------------
  Procedure Get_Childmaxage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
  Begin
    Output_Out := Nvl(zl_GetSysParameter('儿童年龄界定上限'), 0);
  End Get_Childmaxage;

  -----------------------------------------------------
  --获取急诊等级
  -----------------------------------------------------
  Procedure Get_Level
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select a.序号, a.名称, a.严重程度, a.再次评估时限, a.患者标识颜色, Null As 缺省
      From 急诊病情级别 A
      Order By a.序号;
  End Get_Level;

  -----------------------------------------------------
  --获取急诊科室
  -----------------------------------------------------
  Procedure Get_Dept
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select a.Id, a.编码, a.名称, a.简码, Null As 缺省
      From 部门表 A, 临床部门 B
      Where a.Id = b.部门id And b.工作性质 = '20' And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)
      Order By a.编码;
  End Get_Dept;

  -----------------------------------------------------
  --获取人工评估规则
  -----------------------------------------------------
  Procedure Get_Rules
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select a.Id, a.分类, a.指标名称, a.适用人群, a.病情级别 From 急诊人工评定规则 A Order By ID, 病情级别;
  End Get_Rules;

  -----------------------------------------------------
  --根据出生日期返回年龄
  -----------------------------------------------------
  Procedure Get_Datetoage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
    --入参：Json_In:格式
    --  input
    --    birthday        C 1 出生日期 yyyy-mm-dd
    Jsonobj    Pljson;
    j_In    Pljson;
    d_出生日期 Date;
    v_年龄     Varchar2(50);
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    d_出生日期 := To_Date(Jsonobj.Get_String('birthday'), 'yyyy-mm-dd hh24:mi:ss');
    Select Zl_Age_Calc(0, d_出生日期, Sysdate) Into v_年龄 From Dual;
  
    Output_Out := v_年龄;
  End Get_Datetoage;

  -----------------------------------------------------
  --获取性别基础数据
  -----------------------------------------------------
  Procedure Get_Sexbase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 编码, 名称, 简码, Nvl(缺省标志, 0) As 缺省 From 性别 Order By Nvl(缺省标志, 0) Desc, 编码;
  End Get_Sexbase;

  -----------------------------------------------------
  --获取民族基础数据
  -----------------------------------------------------
  Procedure Get_Nationbase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 编码, 名称, 简码, Nvl(缺省标志, 0) As 缺省 From 民族 Order By Nvl(缺省标志, 0) Desc, 编码;
  End Get_Nationbase;

  -----------------------------------------------------
  --获取急诊评分指标
  -----------------------------------------------------
  Procedure Get_Scorebase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select ID, 指标名称, 值域范围, 方法id, 值域单位 From 急诊评分指标 Order By ID;
  End Get_Scorebase;

  -----------------------------------------------------
  --获取急诊主诉
  -----------------------------------------------------
  Procedure Get_Paticc
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select b.名称 分类, a.编码, a.名称, a.简码
      From 急诊常用主诉 A, 急诊常用主诉 B
      Where a.上级 = b.编码 And a.上级 Is Not Null And b.上级 Is Null
      Order By b.编码;
  End Get_Paticc;

  -----------------------------------------------------
  --获取病人来源
  -----------------------------------------------------
  Procedure Get_Patifrom
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 编码, 名称, 缺省标志 As 缺省 From 急诊病人来源 Order By Nvl(缺省标志, 0) Desc, 名称;
  End Get_Patifrom;

  -----------------------------------------------------
  --获取急诊意识状态
  -----------------------------------------------------
  Procedure Get_Patistate
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 编码, 名称, 缺省标志 As 缺省 From 急诊意识状态 Order By Nvl(缺省标志, 0) Desc, 名称;
  End Get_Patistate;

  -----------------------------------------------------
  --获取急诊陪同人员
  -----------------------------------------------------
  Procedure Get_Entourage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 编码, 名称, 缺省标志 As 缺省 From 急诊陪同人员 Order By Nvl(缺省标志, 0) Desc, 名称;
  End Get_Entourage;

  -----------------------------------------------------
  --获取急诊常见既往史
  -----------------------------------------------------
  Procedure Get_Dishistory
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select 编码, 名称, 0 As 缺省 From 急诊常见既往史 Order By 名称;
  End Get_Dishistory;

  -----------------------------------------------------
  --获取数据库系统时间
  -----------------------------------------------------
  Procedure Get_Now_Time
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
  Begin
    Output_Out := To_Char(Sysdate, 'yyyy-MM-dd HH24:mi');
  End Get_Now_Time;
End Pkg_Pretriage_Dql;
/



--145003:蒋廷中,2019-10-15,新增模块急诊预检分诊工作站
Create Or Replace Package Pkg_Pretriage_Dml As

  -----------------------------------------------------
  --变更病人就诊记录的绿色通道状态
  -----------------------------------------------------
  Procedure Change_Greenchannel
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );
  -----------------------------------------------------
  --删除病人就诊记录
  -----------------------------------------------------
  Procedure Del_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );

  -----------------------------------------------------
  --清除挂号事务锁定
  -----------------------------------------------------
  Procedure Register_Unlock
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );

  -----------------------------------------------------
  --更新最新的挂号安排
  -----------------------------------------------------
  Procedure Register_Update
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );

  -----------------------------------------------------
  --保存病人分诊信息
  -----------------------------------------------------
  Procedure Save_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
End Pkg_Pretriage_Dml;
/
Create Or Replace Package Body Pkg_Pretriage_Dml As
  -----------------------------------------------------
  --变更病人就诊记录的绿色通道状态
  -----------------------------------------------------
  Procedure Change_Greenchannel
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) As
    --功能：标记或取消急诊绿色通道
    Jsonobj Pljson;
    j_In    Pljson;
  
    n_Id           急诊就诊记录.Id%Type; --就诊ID
    n_是否绿色通道 急诊就诊记录.是否绿色通道%Type; --是否绿色通道
    n_挂号id       急诊就诊记录.挂号id %Type;
  Begin
    j_In           := Pljson(Input_In);
    Jsonobj        := j_In.Get_Pljson('input');
    n_Id           := To_Number(Jsonobj.Get_String('id'));
    n_是否绿色通道 := To_Number(Jsonobj.Get_String('是否绿色通道'));
  
    Select Max(挂号id) Into n_挂号id From 急诊就诊记录 Where ID = n_Id;
  
    Zl_急诊绿色通道_Edit(n_挂号id, n_是否绿色通道);
    Output_Out := '成功';
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Change_Greenchannel;

  -----------------------------------------------------
  --删除病人就诊记录
  -----------------------------------------------------
  Procedure Del_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
    Jsonobj  Pljson;
    j_In     Pljson;
    n_Id     急诊就诊记录.Id%Type; --就诊ID
    n_病人id 急诊就诊记录.病人id%Type;
    n_挂号id 急诊就诊记录.挂号id%Type;
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    n_Id    := To_Number(Jsonobj.Get_String('id'));
  
    Delete From 急诊就诊记录 Where ID = n_Id Return 病人id, 挂号id Into n_病人id, n_挂号id;
  
    Zl_Emergencyregistdel(n_挂号id);
  
    Delete From 病人信息从表
    Where 病人id = n_病人id And 就诊id = n_挂号id And 信息名 In ('体温', '呼吸', '脉搏', '收缩压', '舒张压', '血糖');
  
    Output_Out := '成功';
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Del_Pretriage;

  -----------------------------------------------------
  --清除挂号事务锁定
  -----------------------------------------------------
  Procedure Register_Unlock
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
    v_人员姓名 Varchar2(200);
    v_Temp     Varchar2(4000);
  Begin
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    Zl_挂号序号状态_Lock(2, v_人员姓名);
    Output_Out := '成功';
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Register_Unlock;

  -----------------------------------------------------
  --更新最新的挂号安排
  -----------------------------------------------------
  Procedure Register_Update
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
  Begin
    Zl_挂号安排_Autoupdate();
    Output_Out := '成功';
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Register_Update;

  -----------------------------------------------------
  --保存病人分诊信息
  -----------------------------------------------------
  Procedure Save_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --入参：Json_In:格式
    Jsonobj Pljson;
    j_In    Pljson;
    n_Type  Number; --1  新增,2  修改
  
    --病人信息
    v_姓名     病人信息.姓名%Type;
    v_性别     病人信息.性别%Type;
    d_出生日期 病人信息.出生日期%Type;
    v_身份证号 病人信息.身份证号%Type;
    v_联系电话 病人信息.联系人电话%Type;
    v_民族     病人信息.民族%Type;
    v_国籍     病人信息.国籍%Type;
    v_医保卡号 病人信息.医保号%Type;
    v_保险类别 保险类别.名称%Type;
    v_家庭地址 病人信息.家庭地址%Type;
  
    n_就诊id 急诊就诊记录.Id%Type;
    n_病人id 急诊就诊记录.病人id%Type;
    n_挂号id 急诊就诊记录.挂号id%Type;
    n_分诊id 急诊分诊记录.Id%Type;
  
    --就诊记录
    v_病人年龄 急诊就诊记录.病人年龄%Type;
    n_年龄数值 急诊就诊记录.年龄数值%Type;
    v_年龄单位 急诊就诊记录.年龄单位%Type;
  
    d_到院时间     急诊就诊记录.到院时间%Type;
    n_是否三无人员 急诊就诊记录.是否三无人员%Type;
    n_是否复合伤   急诊就诊记录.是否复合伤%Type;
    n_是否绿色通道 急诊就诊记录.是否绿色通道%Type;
  
    n_是否成批就诊 急诊就诊记录.是否成批就诊%Type;
    n_成批就诊人数 急诊就诊记录.成批就诊人数%Type;
    v_病人来源     急诊就诊记录.病人来源%Type;
    v_陪同人员     急诊就诊记录.陪同人员%Type;
    v_意识状态     急诊就诊记录.意识状态%Type;
    v_既往病史     急诊就诊记录.既往病史%Type;
    v_主诉         急诊就诊记录.主诉%Type;
    n_病情级别     急诊就诊记录.病情级别%Type;
    v_登记人       急诊就诊记录.登记人%Type;
    d_登记时间     急诊就诊记录.登记时间%Type;
    v_备注         急诊就诊记录.备注%Type;
  
    --分诊记录
    n_分诊次数 急诊分诊记录.分诊次数%Type;
  
    n_分诊科室id   急诊分诊记录.分诊科室id%Type;
    v_分诊科室名称 急诊分诊记录.分诊科室名称%Type;
  
    d_体征测量时间 急诊分诊记录.体征测量时间%Type;
    n_舒张压       急诊分诊记录.舒张压%Type;
    n_收缩压       急诊分诊记录.收缩压%Type;
    n_血糖         急诊分诊记录.血糖%Type;
    n_指氧饱和度   急诊分诊记录.指氧饱和度%Type;
    n_心率         急诊分诊记录.心率%Type;
    n_血钾         急诊分诊记录.血钾%Type;
    n_体温         急诊分诊记录.体温%Type;
    n_呼吸频率     急诊分诊记录.呼吸频率%Type;
  
    n_自动病情级别  急诊分诊记录.自动病情级别%Type;
    n_人工病情级别  急诊分诊记录.人工病情级别%Type;
    v_人工评级说明  急诊分诊记录.人工评级说明%Type;
    v_修改说明      急诊分诊记录.修改说明%Type;
    v_站点          Varchar2(10);
    n_分诊科室idold 急诊分诊记录.分诊科室id%Type;
  
    d_Now Date;
  
    n_门诊号     Number(18);
    n_险类       Number(5);
    v_登记人编号 Varchar2(6);
    n_Count      Number(5);
  
    n_评分id       Number(18);
    n_方法id       Number(18);
    n_评分方法分值 Number(5);
    v_评分结果描述 Varchar2(100);
    n_评分等级     Number(1);
  
    Jsonlist评分指标 Pljson_List;
    Jsonlist病人评分 Pljson_List;
    Jsonlistitem     Pljson;
    Jsonlistitem指标 Pljson;
  
    n_Edittmp Number(5); --0  新增  1  修改
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
  
    n_Type           := Jsonobj.Get_String('type');
    n_就诊id         := Jsonobj.Get_String('就诊id');
    n_病人id         := Nvl(To_Number(Jsonobj.Get_String('病人id')), 0);
    n_门诊号         := Nvl(To_Number(Jsonobj.Get_String('门诊号')), 0);
    v_姓名           := Jsonobj.Get_String('姓名');
    v_性别           := Jsonobj.Get_String('性别');
    d_出生日期       := To_Date(Jsonobj.Get_String('出生日期'), 'yyyy-mm-dd');
    v_身份证号       := Jsonobj.Get_String('身份证号');
    v_联系电话       := Jsonobj.Get_String('联系电话');
    v_民族           := Jsonobj.Get_String('民族');
    v_医保卡号       := Jsonobj.Get_String('医保卡号');
    v_保险类别       := Jsonobj.Get_String('保险类别');
    v_家庭地址       := Jsonobj.Get_String('家庭地址');
    v_病人年龄       := Jsonobj.Get_String('病人年龄');
    n_年龄数值       := To_Number(Jsonobj.Get_String('年龄数值'));
    v_年龄单位       := Jsonobj.Get_String('年龄单位');
    d_到院时间       := To_Date(Jsonobj.Get_String('到院时间'), 'yyyy-mm-dd hh24:mi:ss');
    n_是否三无人员   := To_Number(Jsonobj.Get_String('是否三无人员'));
    n_是否复合伤     := To_Number(Jsonobj.Get_String('是否复合伤'));
    n_是否绿色通道   := To_Number(Jsonobj.Get_String('是否绿色通道'));
    n_是否成批就诊   := To_Number(Jsonobj.Get_String('是否成批就诊'));
    n_成批就诊人数   := To_Number(Jsonobj.Get_String('成批就诊人数'));
    v_病人来源       := Jsonobj.Get_String('病人来源');
    v_陪同人员       := Jsonobj.Get_String('陪同人员');
    v_意识状态       := Jsonobj.Get_String('意识状态');
    v_既往病史       := Jsonobj.Get_String('既往病史');
    v_主诉           := Jsonobj.Get_String('主诉');
    n_病情级别       := To_Number(Jsonobj.Get_String('病情级别'));
    v_登记人         := Jsonobj.Get_String('登记人');
    v_备注           := Jsonobj.Get_String('备注');
    n_分诊科室id     := To_Number(Jsonobj.Get_String('分诊科室id'));
    v_分诊科室名称   := Jsonobj.Get_String('分诊科室名称');
    d_体征测量时间   := To_Date(Jsonobj.Get_String('体征测量时间'), 'yyyy-mm-dd hh24:mi:ss');
    n_舒张压         := To_Number(Jsonobj.Get_String('舒张压'));
    n_收缩压         := To_Number(Jsonobj.Get_String('收缩压'));
    n_血糖           := To_Number(Jsonobj.Get_String('血糖'));
    n_指氧饱和度     := To_Number(Jsonobj.Get_String('指氧饱和度'));
    n_心率           := To_Number(Jsonobj.Get_String('心率'));
    n_血钾           := To_Number(Jsonobj.Get_String('血钾'));
    n_体温           := To_Number(Jsonobj.Get_String('体温'));
    n_呼吸频率       := To_Number(Jsonobj.Get_String('呼吸频率'));
    n_自动病情级别   := To_Number(Jsonobj.Get_String('自动病情级别'));
    n_人工病情级别   := To_Number(Jsonobj.Get_String('人工病情级别'));
    v_人工评级说明   := Jsonobj.Get_String('人工评级说明');
    v_修改说明       := Jsonobj.Get_String('修改说明');
    v_登记人编号     := Jsonobj.Get_String('登记人编号');
    v_站点           := Jsonobj.Get_String('站点');
    v_国籍           := Jsonobj.Get_String('国籍');
    Jsonlist评分指标 := Jsonobj.Get_Pljson_List('评分指标');
    Jsonlist病人评分 := Jsonobj.Get_Pljson_List('病人评分');
  
    n_Edittmp := 0;
  
    --获取登记人编号
    If v_登记人编号 Is Null Then
      Select Max(编号) Into v_登记人编号 From 人员表 Where 姓名 = v_登记人;
    End If;
    --获取保险类别
    If v_保险类别 Is Not Null Then
      Select Max(序号) Into n_险类 From 保险类别 Where 名称 = v_保险类别;
    End If;
  
    Select Sysdate Into d_Now From Dual;
    d_登记时间 := d_Now;
  
    --新增时重新产生
    If n_Type = 1 Then
      Select 急诊就诊记录_Id.Nextval Into n_就诊id From Dual;
    End If;
  
    --分诊ID都是重新产生
    Select 急诊分诊记录_Id.Nextval Into n_分诊id From Dual;
  
    --产生门诊号
    --等待处理身份信息
    If n_Type = 1 Then
      If n_病人id > 0 Then
        n_Edittmp := 1;
      Else
        If v_身份证号 Is Not Null And v_国籍 = '中国' Then
          n_Count := Nvl(zl_GetSysParameter(279), 0);
          If n_Count = 1 Then
            Select Max(病人id) Into n_病人id From 病人信息 Where 身份证号 = v_身份证号;
            If n_病人id > 0 Then
              n_Edittmp := 1;
            End If;
          End If;
        End If;
      End If;
    
      If n_Edittmp = 0 Then
        Select 病人信息_Id.Nextval Into n_病人id From Dual;
        n_门诊号 := Nextno(3);
        Zl_病人信息_Insert(n_病人id, n_门诊号, Null, Null, v_姓名, v_性别, v_病人年龄, d_出生日期, Null, v_身份证号, Null, Null, v_民族, v_国籍,
                       Null, Null, v_家庭地址, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null,
                       Null, n_险类, Sysdate, Null, Null, v_登记人编号, v_登记人, v_医保卡号, Null, Null, Null, Null, Null, Null, Null,
                       v_联系电话);
      Else
        If n_门诊号 = 0 Then
          Select Nvl(Max(门诊号), 0) Into n_门诊号 From 病人信息 Where 病人id = n_病人id;
          If n_门诊号 = 0 Then
            n_门诊号 := Nextno(3);
          End If;
        End If;
        Update 病人信息
        Set 门诊号 = n_门诊号, 姓名 = Nvl(v_姓名, 姓名), 性别 = Nvl(v_性别, 性别), 年龄 = Nvl(v_病人年龄, 年龄), 出生日期 = Nvl(d_出生日期, 出生日期),
            身份证号 = Nvl(v_身份证号, 身份证号), 民族 = Nvl(v_民族, 民族), 家庭地址 = Nvl(v_家庭地址, 家庭地址), 险类 = Nvl(n_险类, 险类),
            医保号 = Nvl(v_医保卡号, 医保号), 手机号 = Nvl(v_联系电话, 手机号), 国籍 = Nvl(v_国籍, 国籍)
        Where 病人id = n_病人id;
      End If;
    Else
      If n_就诊id Is Not Null Then
        Select Max(Nvl(病人id, 0)), Max(Nvl(挂号id, 0)), Max(Nvl(分诊科室id, 0))
        Into n_病人id, n_挂号id, n_分诊科室idold
        From 急诊就诊记录
        Where ID = n_就诊id;
      
        --修改不处理病人信息
        /*Select Max(门诊号) Into n_门诊号 From 病人信息 Where 病人id = n_病人id;
        Update 病人信息
        Set 门诊号 = n_门诊号, 姓名 = Nvl(v_姓名, 姓名), 性别 = Nvl(v_性别, 性别), 年龄 = Nvl(v_病人年龄, 年龄), 出生日期 = Nvl(d_出生日期, 出生日期),
            身份证号 = Nvl(v_身份证号, 身份证号), 民族 = Nvl(v_民族, 民族), 家庭地址 = Nvl(v_家庭地址, 家庭地址), 险类 = Nvl(n_险类, 险类),
            医保号 = Nvl(v_医保卡号, 医保号), 手机号 = Nvl(v_联系电话, 手机号)
        Where 病人id = n_病人id;*/
      End If;
    End If;
  
    If n_Type = 1 Then
      --处理挂号id
    
      n_挂号id := Zl_Emergencyregist(n_病人id, n_分诊科室id, v_站点, n_是否绿色通道);
    
      Insert Into 急诊就诊记录
        (ID, 病人id, 病人年龄, 年龄数值, 年龄单位, 挂号id, 病情级别, 到院时间, 主诉, 是否三无人员, 陪同人员, 病人来源, 既往病史, 意识状态, 是否成批就诊, 成批就诊人数, 是否复合伤, 备注,
         登记人, 登记时间, 分诊病情级别, 是否绿色通道, 分诊科室id)
      Values
        (n_就诊id, n_病人id, v_病人年龄, n_年龄数值, v_年龄单位, n_挂号id, n_病情级别, d_到院时间, v_主诉, n_是否三无人员, v_陪同人员, v_病人来源, v_既往病史, v_意识状态,
         n_是否成批就诊, n_成批就诊人数, n_是否复合伤, v_备注, v_登记人, d_登记时间, n_病情级别, n_是否绿色通道, n_分诊科室id);
    Else
      If n_分诊科室idold <> n_分诊科室id Then
        Zl_Emergencyregistredo(n_挂号id, n_分诊科室id, v_站点);
      End If;
      Update 急诊就诊记录
      Set 病人年龄 = v_病人年龄, 年龄数值 = n_年龄数值, 年龄单位 = v_年龄单位, 挂号id = n_挂号id, 病情级别 = n_病情级别, 到院时间 = d_到院时间, 主诉 = v_主诉,
          是否三无人员 = n_是否三无人员, 陪同人员 = v_陪同人员, 病人来源 = v_病人来源, 既往病史 = v_既往病史, 意识状态 = v_意识状态, 是否成批就诊 = n_是否成批就诊,
          成批就诊人数 = n_成批就诊人数, 是否复合伤 = n_是否复合伤, 备注 = v_备注, 分诊病情级别 = n_病情级别, 是否绿色通道 = n_是否绿色通道, 登记时间 = d_登记时间,
          分诊科室id = n_分诊科室id
      Where ID = n_就诊id;
    End If;
  
    If n_Type = 1 Then
      n_分诊次数 := 1;
    Else
      Select Max(分诊次数) + 1 Into n_分诊次数 From 急诊分诊记录 Where 就诊id = n_就诊id;
    End If;
  
    Insert Into 急诊分诊记录
      (ID, 就诊id, 分诊次数, 自动病情级别, 分诊科室id, 分诊科室名称, 收缩压, 舒张压, 心率, 指氧饱和度, 体温, 血糖, 血钾, 体征测量时间, 登记人, 登记时间, 人工病情级别, 人工评级说明, 呼吸频率,
       修改说明)
    Values
      (n_分诊id, n_就诊id, n_分诊次数, n_自动病情级别, n_分诊科室id, v_分诊科室名称, n_收缩压, n_舒张压, n_心率, n_指氧饱和度, n_体温, n_血糖, n_血钾, d_体征测量时间,
       v_登记人, d_登记时间, n_人工病情级别, v_人工评级说明, n_呼吸频率, v_修改说明);
  
    Delete From 病人信息从表
    Where 病人id = n_病人id And 就诊id = n_挂号id And 信息名 In ('体温', '呼吸', '脉搏', '收缩压', '舒张压', '血糖');
  
    If n_体温 Is Not Null Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
        Select n_病人id, n_挂号id, '体温', To_Char(n_体温) From Dual;
    End If;
  
    If n_呼吸频率 Is Not Null Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
        Select n_病人id, n_挂号id, '呼吸', To_Char(n_呼吸频率) From Dual;
    End If;
  
    If n_心率 Is Not Null Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
        Select n_病人id, n_挂号id, '脉搏', To_Char(n_心率) From Dual;
    End If;
  
    If n_收缩压 Is Not Null Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
        Select n_病人id, n_挂号id, '收缩压', To_Char(n_收缩压) From Dual;
    End If;
  
    If n_舒张压 Is Not Null Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
        Select n_病人id, n_挂号id, '舒张压', To_Char(n_舒张压) From Dual;
    End If;
  
    If n_血糖 Is Not Null Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
        Select n_病人id, n_挂号id, '血糖', To_Char(n_血糖) From Dual;
    End If;
  
    For I In 1 .. Jsonlist病人评分.Count Loop
      Jsonlistitem   := Pljson();
      Jsonlistitem   := Pljson(Jsonlist病人评分.Get(I));
      n_方法id       := To_Number(Jsonlistitem.Get_String('方法ID'));
      n_评分方法分值 := To_Number(Jsonlistitem.Get_String('评分方法分值'));
      v_评分结果描述 := Jsonlistitem.Get_String('评分结果描述');
      n_评分等级     := To_Number(Jsonlistitem.Get_String('评分等级'));
      Select 急诊病人评分_Id.Nextval Into n_评分id From Dual;
    
      Insert Into 急诊病人评分
        (ID, 分诊id, 方法id, 评分方法分值, 评分结果描述, 病情级别)
      Values
        (n_评分id, n_分诊id, n_方法id, n_评分方法分值, v_评分结果描述, n_评分等级);
    
      For I In 1 .. Jsonlist评分指标.Count Loop
        Jsonlistitem指标 := Pljson();
        Jsonlistitem指标 := Pljson(Jsonlist评分指标.Get(I));
        If n_方法id = To_Number(Jsonlistitem指标.Get_String('方法ID')) Then
          Insert Into 急诊病人评分指标
            (评分id, 指标id, 指标结果文本)
          Values
            (n_评分id, To_Number(Jsonlistitem指标.Get_String('指标ID')), Jsonlistitem指标.Get_String('指标结果文本'));
        End If;
      End Loop;
    End Loop;
  
    Open Output_Out For
      Select n_病人id As 病人id, n_就诊id As 就诊id, n_分诊id As 分诊id From Dual;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Save_Pretriage;

End Pkg_Pretriage_Dml;
/
