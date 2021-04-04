----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------



------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--130409:焦博,2018-09-17,挂号管理模块增加系统参数病人同一号源限挂N个号
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1111, 0, 0, 0, 0, 0, 0, 74, '病人同一号源限挂N个号', Null, '0',
         '控制同一病人在当天内对同一号源允许挂号（含预约数)的数量', '0-不限制；>=1表示限制数量', '', '适用于个别医院需要对病人在同一号源的挂号数量控制', Null
  From Dual;

-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--130409:焦博,2018-09-19,挂号管理模块增加系统参数病人同一号源限挂N个号
Create Or Replace Function Zl_Fun_病人挂号记录_Check
(
  操作方式_In   Integer,
  病人id_In     门诊费用记录.病人id%Type,
  号码_In       挂号安排.号码%Type,
  出诊记录id_In 临床出诊记录.Id%Type := Null,
  发生时间_In   门诊费用记录.发生时间%Type,
  专家号_In     Number := 0
) Return Varchar2 As
  --功能：挂号有效性检查(包含预约;预约挂号不扣款;预约挂号扣款)
  --入参:操作方式_IN:0-挂号(包含收款预约),1-预约,2-预约接收
  --     是否加号_In:是否加号调用，0-非加号调用，1-加号调用
  --返回:0-检查通过
  --     1-特定检查项检查失败，同时返回错误提示文本
  --     2-其他错误导致的检查失败，同时返回错误提示文本
  --说明:
  --    当出诊记录id_In=0是,为不当班医生挂号,不需要进行挂号有效性检查,直接退出
  Err_Item Exception;
  n_病人预约科室数 Number(18);
  n_已约科室       Number(18);
  v_Temp           Varchar2(500);
  v_加入原因       特殊病人.加入原因%Type;
  n_同科限号数     Number;
  n_同科限约数     Number;
  n_科室id         挂号安排.科室id%Type;
  n_Count          Number(18);
  n_病人挂号科室数 Number;
  n_专家号挂号限制 Number;
  n_专家号预约限制 Number;
  n_专家号         Number;
  n_同源号限号数   Number;

  Cursor c_Pati(n_病人id 病人信息.病人id%Type) Is
    Select a.病人id, a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, a.费别, a.险类, c.编码 As 付款方式
    From 病人信息 A, 医疗付款方式 C
    Where a.病人id = n_病人id And a.医疗付款方式 = c.名称(+);

  r_Pati c_Pati%RowType;

  Function Zl_操作员
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-获取缺省部门ID;1-获取操作员编号;2-获取操作员姓名
    -- SplitStr:格式为:部门ID,部门名称;人员ID,人员编号,人员姓名(用Zl_Identity获取的)
  Begin
    If Type_In = 0 Then
      --缺省部门
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --操作员编码
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --操作员姓名
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      Return v_Sub;
    End If;
  End;

Begin
  --检测病人相关
  If 出诊记录id_In = 0 Then
    Return '0-号源正常';
  End If;

  Open c_Pati(病人id_In);
  n_Count := 0;
  Begin
    Fetch c_Pati
      Into r_Pati;
    n_Count := 1;
  Exception
    When Others Then
      n_Count := -1;
  End;
  If n_Count <= 0 Then
    Return '1|病人未找到，不能继续。';
  End If;
  --预约检测黑名单
  If 操作方式_In = 1 Then
    Begin
      Select 加入原因 Into v_加入原因 From 特殊病人 Where 撤消时间 Is Null And 病人id = 病人id_In And Rownum = 1;
      Return '1|此病人在特殊病人名单中，原因：【' || v_加入原因 || '】不能继续！';
    Exception
      When Others Then
        Null;
    End;
  End If;

  --检测挂号时间
  If Trunc(Sysdate) > Trunc(发生时间_In) Then
    Return '1|不能挂以前的号(' || To_Char(发生时间_In, 'yyyy-mm-dd') || ')。';
  End If;

  --部门ID,部门名称;人员ID,人员编号,人员姓名
  v_Temp := Zl_Identity(0);
  If Nvl(v_Temp, ' ') = ' ' Then
    Return '1|当前操作人员未设置对应的人员关系,不能继续。';
  End If;

  n_专家号 := 专家号_In;
  If 出诊记录id_In Is Null Then
    Select Max(科室id) Into n_科室id From 挂号安排 Where 号码 = 号码_In;
  Else
    Select Max(科室id) Into n_科室id From 临床出诊记录 Where ID = 出诊记录id_In;
  End If;

  --检测系统参数
  v_Temp           := Nvl(zl_GetSysParameter('病人同科限挂N个号', 1111), '0|0') || '|';
  n_同科限号数     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
  n_同科限约数     := To_Number(Nvl(zl_GetSysParameter('病人同科限约N个号', 1111), '0'));
  n_病人预约科室数 := To_Number(Nvl(zl_GetSysParameter('病人预约科室数', 1111), '0'));
  n_病人挂号科室数 := To_Number(Nvl(zl_GetSysParameter('病人挂号科室限制', 1111), '0'));
  n_专家号挂号限制 := To_Number(Nvl(zl_GetSysParameter('专家号挂号限制'), '0'));
  n_专家号预约限制 := To_Number(Nvl(zl_GetSysParameter('专家号预约限制'), '0'));
  n_同源号限号数   := To_Number(Nvl(zl_GetSysParameter('病人同一号源限挂N个号', '1111'), '0'));
  --对参数控制进行检查
  If 操作方式_In = 1 Then
    If Nvl(n_同科限约数, 0) <> 0 Or Nvl(n_病人预约科室数, 0) <> 0 Then
      n_已约科室 := 0;
      For c_Chkitem In (Select Distinct 执行部门id
                        From 病人挂号记录
                        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
                              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> n_科室id) Loop
        n_已约科室 := n_已约科室 + 1;
      End Loop;
      If n_已约科室 >= Nvl(n_病人预约科室数, 0) And Nvl(n_病人预约科室数, 0) > 0 Then
        Return '1|同一病人最多同时能预约[' || Nvl(n_病人预约科室数, 0) || ']个科室,不能再预约！';
      End If;
    
      Select Count(1)
      Into n_Count
      From 病人挂号记录
      Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
            Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id = n_科室id;
      If n_Count >= Nvl(n_同科限约数, 0) And Nvl(n_同科限约数, 0) > 0 Then
        Return '1|该病人已经在该科室预约了' || n_Count || '次,不能再预约！';
      End If;
    End If;
    If Nvl(n_专家号预约限制, 0) <> 0 And n_专家号 = 1 Then
      If 出诊记录id_In Is Null Then
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 号别 = 号码_In;
      Else
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 出诊记录id = 出诊记录id_In;
      End If;
      If n_Count >= Nvl(n_专家号预约限制, 0) And Nvl(n_专家号预约限制, 0) > 0 Then
        Return '1|该病人已经超过本号预约限制,不能再预约！';
      End If;
    End If;
  Else
    If (Nvl(n_同科限号数, 0) <> 0 Or Nvl(n_病人挂号科室数, 0) <> 0) And 操作方式_In = 0 Then
      n_已约科室 := 0;
      For c_Chkitem In (Select Distinct 执行部门id
                        From 病人挂号记录
                        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And
                              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> n_科室id) Loop
        n_已约科室 := n_已约科室 + 1;
      End Loop;
      If n_已约科室 >= Nvl(n_病人挂号科室数, 0) And Nvl(n_病人挂号科室数, 0) > 0 Then
        Return '1|同一病人最多同时能挂号[' || Nvl(n_病人挂号科室数, 0) || ']个科室,不能再挂号！';
      End If;
    
      Select Count(1)
      Into n_Count
      From 病人挂号记录
      Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And
            Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id = n_科室id;
      If n_Count >= Nvl(n_同科限号数, 0) And Nvl(n_同科限号数, 0) > 0 Then
        Return '1|该病人已经在该科室挂号了' || n_Count || '次,不能再挂号！';
      End If;
    End If;
  
    If Nvl(n_专家号挂号限制, 0) <> 0 And n_专家号 = 1 Then
      If 出诊记录id_In Is Null Then
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 号别 = 号码_In;
      Else
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 出诊记录id = 出诊记录id_In;
      End If;
      If n_Count >= Nvl(n_专家号挂号限制, 0) And Nvl(n_专家号挂号限制, 0) > 0 Then
        Return '1|该病人已经超过本号挂号限制,不能再挂号！';
      End If;
    End If;
  End If;
  If 操作方式_In <> 2 And Nvl(n_同源号限号数, 0) <> 0 Then
    If 出诊记录id_In Is Null Then
      Select Count(1)
      Into n_Count
      From 病人挂号记录
      Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 In (1, 2) And 发生时间 Between Trunc(发生时间_In) And
            Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 号别 = 号码_In;
    Else
      Select Count(1)
      Into n_Count
      From 病人挂号记录
      Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 In (1, 2) And 出诊记录id = 出诊记录id_In;
    End If;
    If n_Count >= Nvl(n_同源号限号数, 0) And Nvl(n_同源号限号数, 0) > 0 Then
      Return '1|同一病人最多能同时挂(预约)[' || Nvl(n_同源号限号数, 0) || ']个相同号别的号,不能再挂号(预约)！';
    End If;
  End If;
  Return '0-号源正常';

Exception
  When Others Then
    Return '2-' || SQLErrM;
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_病人挂号记录_Check;
/

Create Or Replace Procedure Zl_Third_Registercheck
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:HIS挂号检查
  --入参:Xml_In:
  --<IN>
  --  <BRID>1</BRID>                    //病人ID
  --  <XM>姓名</XM>                     //姓名
  --  <SFZH>510221197008184710</SFZH>   //身份证号
  --  <HM>0100</HM>                     //号码
  --  <CZJLID>100</CZJLID>              //出诊记录ID,计划排班模式可以不传
  --  <GHSJ>2016-08-10 09:52:00</GHSJ>  //挂号时间
  --  <KSID>1</KSID>                    //科室ID
  --  <YSXM>张震</YSXM>                 //医生姓名
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  -- <ERROR><MSG></MSG></ERROR> //为空表示检查成功
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  n_病人id         病人信息.病人id%Type;
  v_姓名           病人信息.姓名%Type;
  v_身份证号       病人信息.身份证号%Type;
  v_号码           挂号安排.号码%Type;
  n_出诊记录id     Number(18);
  d_发生时间       病人挂号记录.发生时间%Type;
  v_Para           Varchar2(500);
  d_启用时间       Date;
  n_挂号模式       Number(3);
  n_同科限号数     Number;
  n_同科限约数     Number;
  n_同源限号数     Number;
  n_病人挂号科室数 Number;
  n_病人预约科室数 Number;
  n_专家号挂号限制 Number;
  n_专家号预约限制 Number;
  n_Exists         Number;
  n_Count          Number;
  n_科室id         病人挂号记录.执行部门id%Type;
  v_医生姓名       病人挂号记录.执行人%Type;
  v_性别           病人信息.性别%Type;
  v_年龄           病人信息.年龄%Type;
  n_已约科室       Number;
  v_Checkresult    Varchar2(500);
  v_Temp           Varchar2(32767); --临时XML
  x_Templet        Xmltype; --模板XML
  v_Err_Msg        Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/BRID'), Extractvalue(Value(A), 'IN/HM'), Extractvalue(Value(A), 'IN/CZJLID'),
         To_Date(Extractvalue(Value(A), 'IN/GHSJ'), 'yyyy-mm-dd hh24:mi:ss'),
         To_Number(Extractvalue(Value(A), 'IN/KSID')), Extractvalue(Value(A), 'IN/YSXM'),
         Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM')
  Into n_病人id, v_号码, n_出诊记录id, d_发生时间, n_科室id, v_医生姓名, v_身份证号, v_姓名
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
  End If;
  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '无法确定病人信息,请检查';
    Raise Err_Item;
  End If;

  v_Para := zl_GetSysParameter(256);
  If v_Para Is Not Null Then
    n_挂号模式 := Substr(v_Para, 1, 1);
    Begin
      d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_启用时间 := Null;
    End;
  
    If Sysdate - 10 > Nvl(d_启用时间, Sysdate - 30) Then
      If n_挂号模式 = 1 And Nvl(d_发生时间, Sysdate) > Nvl(d_启用时间, Sysdate - 30) And n_出诊记录id Is Null Then
        v_Temp := '系统当前处于出诊表排班模式，传入的参数无法确定挂号安排，请重试！';
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End If;
    Else
      If n_挂号模式 = 1 And Nvl(d_发生时间, Sysdate) > Nvl(d_启用时间, Sysdate - 30) And n_出诊记录id Is Null Then
        Begin
          Select a.Id
          Into n_出诊记录id
          From 临床出诊记录 A, 临床出诊号源 B
          Where a.号源id = b.Id And b.号码 = v_号码 And Nvl(d_发生时间, Sysdate) Between a.开始时间 And a.终止时间;
        Exception
          When Others Then
            v_Temp := '系统当前处于出诊表排班模式，传入的参数无法确定挂号安排，请重试！';
            v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
            Xml_Out := x_Templet;
            Return;
        End;
      End If;
    End If;
  End If;

  If n_出诊记录id Is Not Null Then
    Select 性别, 年龄 Into v_性别, v_年龄 From 病人信息 Where 病人id = n_病人id And Rownum < 2;
    v_Checkresult := Zl_临床出诊限制_Check(n_出诊记录id, v_年龄, v_性别);
    If Substr(Nvl(v_Checkresult, '0'), 1, 1) <> '0' Then
      v_Temp := '病人不适用该本号别,请检查！';
      v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      Xml_Out := x_Templet;
      Return;
    End If;
  End If;

  If Trunc(Sysdate) > Trunc(d_发生时间) Then
    v_Temp := '不能挂以前的号(' || To_Char(d_发生时间, 'yyyy-mm-dd') || ')。';
    v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Xml_Out := x_Templet;
    Return;
  End If;

  v_Temp := Zl_Identity(0);
  If Nvl(v_Temp, ' ') = ' ' Then
    v_Temp := '当前操作人员未设置对应的人员关系,不能继续。';
    v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Xml_Out := x_Templet;
    Return;
  End If;

  v_Temp           := Nvl(zl_GetSysParameter('病人同科限挂N个号', 1111), '0|0') || '|';
  n_同科限号数     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
  n_同科限约数     := To_Number(Nvl(zl_GetSysParameter('病人同科限约N个号', 1111), '0'));
  n_病人预约科室数 := To_Number(Nvl(zl_GetSysParameter('病人预约科室数', 1111), '0'));
  n_病人挂号科室数 := To_Number(Nvl(zl_GetSysParameter('病人挂号科室限制', 1111), '0'));
  n_专家号挂号限制 := To_Number(Nvl(zl_GetSysParameter('专家号挂号限制'), '0'));
  n_专家号预约限制 := To_Number(Nvl(zl_GetSysParameter('专家号预约限制'), '0'));
  n_同源限号数     := To_Number(Nvl(zl_GetSysParameter('病人同一号源限挂N个号', 1111), '0'));

  If Trunc(Sysdate) <> Trunc(d_发生时间) Then
    If Nvl(n_病人预约科室数, 0) <> 0 Then
      n_已约科室 := 0;
      For c_Chkitem In (Select Distinct 执行部门id
                        From 病人挂号记录
                        Where 病人id = n_病人id And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(d_发生时间) And
                              Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> n_科室id) Loop
        n_已约科室 := n_已约科室 + 1;
      End Loop;
      If n_已约科室 >= Nvl(n_病人预约科室数, 0) And Nvl(n_病人预约科室数, 0) > 0 Then
        v_Temp := '同一病人最多同时能预约[' || Nvl(n_病人预约科室数, 0) || ']个科室,不能再预约！';
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End If;
    End If;
    If Nvl(n_同科限约数, 0) <> 0 Then
      Select Count(1)
      Into n_Count
      From 病人挂号记录
      Where 病人id = n_病人id And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(d_发生时间) And
            Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And 执行部门id = n_科室id;
      If n_Count >= Nvl(n_同科限约数, 0) And Nvl(n_同科限约数, 0) > 0 Then
        v_Temp := '该病人已经在该科室预约了' || n_Count || '次,不能再预约！';
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End If;
    End If;
  Else
    If Nvl(n_病人挂号科室数, 0) <> 0 Then
      n_已约科室 := 0;
      For c_Chkitem In (Select Distinct 执行部门id
                        From 病人挂号记录
                        Where 病人id = n_病人id And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(d_发生时间) And
                              Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> n_科室id) Loop
        n_已约科室 := n_已约科室 + 1;
      End Loop;
      If n_已约科室 >= Nvl(n_病人挂号科室数, 0) And Nvl(n_病人挂号科室数, 0) > 0 Then
        v_Temp := '同一病人最多同时能挂号[' || Nvl(n_病人挂号科室数, 0) || ']个科室,不能再挂号！';
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End If;
    End If;
    If Nvl(n_同科限号数, 0) <> 0 Then
      Select Count(1)
      Into n_Count
      From 病人挂号记录
      Where 病人id = n_病人id And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(d_发生时间) And
            Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And 执行部门id = n_科室id;
      If n_Count >= Nvl(n_同科限号数, 0) And Nvl(n_同科限号数, 0) > 0 Then
        v_Temp := '该病人已经在该科室挂号了' || n_Count || '次,不能再挂号！';
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End If;
    End If;
  End If;

  If Nvl(n_同源限号数, 0) <> 0 Then
    If n_出诊记录id Is Null Then
      Select Count(1)
      Into n_Count
      From 病人挂号记录
      Where 病人id = n_病人id And 记录状态 = 1 And 记录性质 In (1, 2) And 发生时间 Between Trunc(d_发生时间) And
            Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And 号别 = v_号码;
    Else
      Select Count(1)
      Into n_Count
      From 病人挂号记录
      Where 病人id = n_病人id And 记录状态 = 1 And 记录性质 In (1, 2) And 出诊记录id = n_出诊记录id;
    End If;
    If n_Count >= Nvl(n_同源限号数, 0) And Nvl(n_同源限号数, 0) > 0 Then
      v_Temp := '同一病人最多能同时挂(预约)[' || Nvl(n_同源限号数, 0) || ']个相同号别的号,不能再挂号(预约)！';
      v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      Xml_Out := x_Templet;
      Return;
    End If;
  End If;

  If Trunc(Sysdate) = Trunc(d_发生时间) Then
    --挂号
    If Nvl(n_专家号挂号限制, 0) <> 0 And v_医生姓名 Is Not Null Then
      If n_出诊记录id Is Null Then
        --无出诊记录对应
        Begin
          Select Count(1)
          Into n_Exists
          From 病人挂号记录
          Where 病人id = n_病人id And 号别 = v_号码 And 发生时间 Between Trunc(d_发生时间) And Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And
                记录状态 = 1 And 记录性质 = 1;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists >= n_专家号挂号限制 Then
          v_Temp := '该病人已经超过本号挂号限制,不能再次挂号！';
          v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
          Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
          Xml_Out := x_Templet;
          Return;
        End If;
      Else
        --对应出诊记录
        Begin
          Select Count(1)
          Into n_Exists
          From 病人挂号记录
          Where 病人id = n_病人id And 出诊记录id = n_出诊记录id And 记录状态 = 1 And 记录性质 = 1;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists >= n_专家号挂号限制 Then
          v_Temp := '该病人已经超过本号挂号限制,不能再次挂号！';
          v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
          Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
          Xml_Out := x_Templet;
          Return;
        End If;
      End If;
    End If;
  Else
    --预约
    If Nvl(n_专家号预约限制, 0) <> 0 And v_医生姓名 Is Not Null Then
      If n_出诊记录id Is Null Then
        --无出诊记录对应
        Begin
          Select Count(1)
          Into n_Exists
          From 病人挂号记录
          Where 病人id = n_病人id And 号别 = v_号码 And 发生时间 Between Trunc(d_发生时间) And Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And
                记录状态 = 1 And 记录性质 = 2;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists >= n_专家号预约限制 Then
          v_Temp := '该病人已经超过本号预约限制,不能再次预约！';
          v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
          Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
          Xml_Out := x_Templet;
          Return;
        End If;
      Else
        --对应出诊记录
        Begin
          Select Count(1)
          Into n_Exists
          From 病人挂号记录
          Where 病人id = n_病人id And 出诊记录id = n_出诊记录id And 记录状态 = 1 And 记录性质 = 2;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists >= n_专家号预约限制 Then
          v_Temp := '该病人已经超过本号预约限制,不能再次预约！';
          v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
          Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
          Xml_Out := x_Templet;
          Return;
        End If;
      End If;
    End If;
  End If;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Registercheck;
/

Create Or Replace Procedure Zl_三方机构挂号_Insert
(
  操作方式_In      Integer,
  病人id_In        门诊费用记录.病人id%Type,
  号码_In          挂号安排.号码%Type,
  号序_In          挂号序号状态.序号%Type,
  单据号_In        门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  结算方式_In      Varchar2,
  摘要_In          门诊费用记录.摘要%Type, --预约挂号摘要信息
  发生时间_In      门诊费用记录.发生时间%Type,
  登记时间_In      门诊费用记录.登记时间%Type,
  合作单位_In      挂号合作单位.名称%Type,
  挂号金额合计_In  门诊费用记录.实收金额%Type,
  领用id_In        票据使用明细.领用id%Type,
  收费票据_In      Number := 0, --挂号是否使用收费票据
  交易流水号_In    病人预交记录.交易流水号%Type,
  交易说明_In      病人预交记录.交易说明%Type,
  预约方式_In      预约方式.名称%Type := Null,
  预交id_In        病人预交记录.Id%Type := Null,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  加入序号状态_In  Number := 0,
  是否自助设备_In  Number := 0,
  结帐id_In        门诊费用记录.结帐id%Type := Null,
  锁定类型_In      Number := 0,
  保险结算_In      Varchar2 := Null,
  冲预交_In        Number := Null,
  支付卡号_In      病人预交记录.卡号%Type := Null,
  退号重用_In      Number := 1,
  费别_In          门诊费用记录.费别%Type := Null,
  冲预交病人ids_In Varchar2 := Null,
  机器名_In        挂号序号状态.机器名%Type := Null,
  更新年龄_In      Number := 0,
  购买病历_In      Number := 0,
  出诊记录id_In    临床出诊记录.Id%Type := Null,
  记帐费用_In      Number := 0,
  付款方式_In      医疗付款方式.名称%Type := Null
) As
  --功能：三方机构进行挂号(包含预约;预约挂号不扣款;预约挂号扣款)
  --入参:操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
  --      结算方式_IN:支持多种结算方式,多种结算方式时，传入格式如下:结算方式名称1,金额,结算号码,三方卡标志|结算方式名称2,金额,结算号码,三方卡标志|...
  --      加入序号状态_In:1-表示强制加入挂号序号状态表中;否则根据启用序号或启用时段时才加入.
  --      是否自助设备_In:1-表示是医院的自助设备进行此函数的调用,自助设备调用此函数 允许加号,否则不允许
  --      锁定类型_In :0-产生正常数据 1-表示对单据进行锁定,产生未生效的单据信息;2-对锁定的记录进行解锁-生成正常数据:未生效的单据在银行扣款完成后进行解锁
  --      保险结算_IN:格式="结算方式|结算金额||....."
  --      冲预交病人ids_In:多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
  Err_Item Exception;
  Err_Special Exception;
  v_Err_Msg            Varchar2(255);
  n_打印id             票据打印内容.Id%Type;
  n_返回值             病人预交记录.金额%Type;
  v_排队号码           Varchar2(20);
  v_队列名称           排队叫号队列.队列名称%Type;
  n_预交id             病人预交记录.Id%Type;
  n_挂号id             病人挂号记录.Id%Type;
  v_结算内容           Varchar2(3000);
  v_当前结算           Varchar2(150);
  d_发生时间           Date;
  v_结算方式           病人预交记录.结算方式%Type;
  n_结算金额           病人预交记录.冲预交%Type;
  n_结算合计           Number(16, 5);
  n_预交金额           病人预交记录.冲预交%Type;
  n_组id               财务缴款分组.Id%Type;
  d_排队时间           Date;
  n_锁定               Number;
  n_病人预约科室数     Number(18);
  n_已约科室           Number(18);
  n_合作单位限制       Number(18);
  n_是否开放           Number(1);
  n_Count              Number(18);
  n_行号               Number(18);
  n_序号               病人挂号记录.号序%Type;
  n_费用id             门诊费用记录.Id%Type;
  n_价格父号           Number(18);
  n_原项目id           收费项目目录.Id%Type;
  n_原收入项目id       收费项目目录.Id%Type;
  v_诊室               病人挂号记录.诊室%Type;
  n_安排id             挂号安排.Id%Type;
  n_实收金额合计       门诊费用记录.实收金额%Type;
  n_开单部门id         门诊费用记录.开单部门id%Type;
  n_实收金额           门诊费用记录.实收金额%Type;
  n_应收金额           门诊费用记录.实收金额%Type;
  n_结帐id             病人结帐记录.Id%Type;
  v_Temp               Varchar2(500);
  n_预约时段序号       Number;
  n_预约总数           Number;
  n_Exists             Number;
  n_分时点显示         Number;
  d_时段开始时间       Date;
  v_冲预交病人ids      Varchar2(4000);
  v_收费项目ids        Varchar2(300);
  n_预约数量           合作单位挂号汇总.已约数%Type;
  n_号序               病人挂号记录.号序%Type;
  d_登记时间           Date;
  v_操作员编号         人员表.编号%Type;
  v_操作员姓名         人员表.姓名%Type;
  n_急诊               病人挂号记录.急诊%Type;
  n_预约               Integer;
  v_星期               挂号安排时段.星期%Type;
  n_启用分时段         Integer;
  n_已挂数             病人挂号汇总.已挂数%Type;
  n_已约数             病人挂号汇总.已约数%Type;
  n_其中已接收         病人挂号汇总.已约数%Type;
  n_预约生成队列       Number;
  d_Date               Date;
  n_挂号序号           Number;
  v_排队序号           排队叫号队列.排队序号%Type;
  v_机器名             挂号序号状态.机器名%Type;
  v_序号操作员         挂号序号状态.操作员姓名%Type;
  v_序号机器名         挂号序号状态.机器名%Type;
  n_序号锁定           Number := 0;
  n_病历费id           收费特定项目.收费细目id%Type;
  v_付款方式           病人挂号记录.医疗付款方式%Type;
  v_费别               门诊费用记录.费别%Type;
  n_屏蔽费别           Number(3) := 0;
  n_Tmp安排id          挂号安排.Id%Type;
  n_计划id             挂号安排计划.Id%Type;
  v_年龄               病人信息.年龄%Type;
  n_合作单位限数量模式 Number;
  n_出诊记录id         临床出诊记录.Id%Type;
  n_挂号模式           Number(3);
  n_同科限号数         Number;
  n_同科限约数         Number;
  n_同源限号数         Number;
  n_病人挂号科室数     Number;
  d_启用时间           Date;
  v_Para               Varchar2(2000);
  n_专家号挂号限制     Number;
  n_专家号预约限制     Number;
  v_站点               部门表.站点%Type;
  v_普通等级           Varchar2(100);
  v_Pricegrade         Varchar2(500);
  v_时间段             时间段.时间段%Type;
  d_检查开始时间       时间段.开始时间%Type;
  d_检查结束时间       时间段.终止时间%Type;
  v_传入               Varchar2(100);
  n_更新项目id         挂号安排.项目id%Type;
  n_项目id             挂号安排.项目id%Type;

  Cursor c_Pati(n_病人id 病人信息.病人id%Type) Is
    Select a.病人id, a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, a.费别, a.险类, c.编码 As 付款方式, a.出生日期, a.身份证号
    From 病人信息 A, 医疗付款方式 C
    Where a.病人id = n_病人id And a.医疗付款方式 = c.名称(+);

  r_Pati c_Pati%RowType;

  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select 病人id, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1
     Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By NO, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 结帐id, NO;

  Cursor c_安排
  (
    v_号码        挂号安排.号码%Type,
    d_发生时间_In Date
  ) Is
    Select *
    From (With 安排时间段 As (Select 时间段
                         From (Select 时间段,
                                       To_Date(Decode(Sign(开始时间 - 终止时间), 1, '3000-01-09 ' || To_Char(开始时间, 'HH24:MI:SS'),
                                                       '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS')), 'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                       To_Date(Decode(Sign(开始时间 - 终止时间), 1, '3000-01-11 ' || To_Char(终止时间, 'HH24:MI:SS'),
                                                       '3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS')), 'yyyy-mm-dd hh24:mi:ss') As 终止时间,
                                       To_Date('3000-01-10 ' || To_Char(d_发生时间_In, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As 当前时间,
                                       To_Date('3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间1,
                                       To_Date('3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As 终止时间1
                                From 时间段)
                         Where 当前时间 Between 开始时间 And 终止时间1 Or 当前时间 Between 开始时间1 And 终止时间)
           Select Distinct p.Id, p.号类, p.号码, p.科室id, b.编码 As 科室编码, b.名称 As 科室名称, p.项目id, c.编码 As 项目编码, c.名称 As 项目名称,
                           p.医生id, d.编号 As 医生编号, p.医生姓名, p.限号数, p.限约数, p.周日 As 日, p.周一 As 一, p.周二 As 二, p.周三 As 三,
                           p.周四 As 四, p.周五 As 五, p.周六 As 六, p.序号控制, p.计划id
           From (Select p.Id, p.号码, p.号类, p.科室id, p.项目id, p.医生id, p.医生姓名, b.限号数, b.限约数, Nvl(p.病案必须, 0) As 病案必须, p.周日, p.周一,
                         p.周二, p.周三, p.周四, p.周五, p.周六, p.分诊方式, p.序号控制,
                         Decode(To_Char(d_发生时间_In, 'D'), '1', p.周日, '2', p.周一, '3', p.周二, '4', p.周三, '5', p.周四, '6', p.周五,
                                 '7', p.周六, Null) As 排班, Null As 计划id
                  From 挂号安排 P, 挂号安排限制 B
                  Where p.停用日期 Is Null And p.Id = b.安排id(+) And
                        b.限制项目(+) = Decode(To_Char(d_发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四',
                                           '6', '周五', '7', '周六', Null) And
                        d_发生时间_In Between Nvl(p.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                        Nvl(p.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Not Exists
                   (Select 1
                         From 挂号安排计划
                         Where 安排id = p.Id And (d_发生时间_In Between 生效时间 And 失效时间) And 审核时间 Is Not Null) And Not Exists
                   (Select 1
                         From 挂号安排停用状态
                         Where 安排id = p.Id And d_发生时间_In Between 开始停止时间 And 结束停止时间) And p.号码 = v_号码
                  Union All
                  Select c.Id, c.号码, c.号类, c.科室id, p.项目id, p.医生id, p.医生姓名, b.限号数, b.限约数, Nvl(c.病案必须, 0) As 病案必须, p.周日, p.周一,
                         p.周二, p.周三, p.周四, p.周五, p.周六, p.分诊方式, p.序号控制,
                         Decode(To_Char(d_发生时间_In, 'D'), '1', p.周日, '2', p.周一, '3', p.周二, '4', p.周三, '5', p.周四, '6', p.周五,
                                 '7', p.周六, Null) As 排班, p.Id As 计划id
                  From 挂号安排计划 P, 挂号安排 C, 挂号计划限制 B,
                       (Select Max(a.生效时间) As 生效, 安排id
                         From 挂号安排计划 A, 挂号安排 B
                         Where a.安排id = b.Id And a.审核时间 Is Not Null And
                               发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                               Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And b.号码 = 号码_In
                         Group By 安排id) E
                  Where p.安排id = c.Id And p.Id = b.计划id(+) And p.生效时间 = e.生效 And p.安排id = e.安排id And
                        Nvl(p.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) = Nvl(e.生效, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                        b.限制项目(+) = Decode(To_Char(d_发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四',
                                           '6', '周五', '7', '周六', Null) And (d_发生时间_In Between p.生效时间 And p.失效时间) And
                        p.审核时间 Is Not Null And Not Exists
                   (Select 1
                         From 挂号安排停用状态
                         Where 安排id = c.Id And d_发生时间_In Between 开始停止时间 And 结束停止时间) And p.号码 = v_号码) P, 部门表 B, 收费项目目录 C,
                人员表 D
           Where p.科室id = b.Id And p.医生id = d.Id(+) And p.项目id = c.Id And
                 (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                 (Nvl(p.医生id, 0) = 0 Or Exists
                  (Select 1
                   From 人员表 Q
                   Where p.医生id = q.Id And (q.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or q.撤档时间 Is Null))) And Exists
            (Select 1 From 安排时间段 Where 时间段 = p.排班))
           Order By 号码;


  r_安排 c_安排%RowType;

  Function Zl_诊室(号码_In 挂号安排.号码%Type) Return Varchar2 As
    n_分诊方式 挂号安排.分诊方式%Type;
    n_安排id   挂号安排.Id%Type;
    v_诊室     病人挂号记录.诊室%Type;
    v_Rowid    Varchar2(500);
    n_Next     Integer;
    n_First    Integer;
  Begin
  
    If 锁定类型_In = 2 Then
      --对单据进行解锁,首先检查是否存在锁定
      Select Count(Rowid) Into n_锁定 From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
      If n_锁定 = 0 Then
        v_Err_Msg := '单据号为(' || 单据号_In || ')的单据,不存在或者已经被解锁!';
        Raise Err_Item;
      End If;
      Select Max(号序) Into n_号序 From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
    End If;
  
    Begin
      Select ID, Nvl(分诊方式, 0) Into n_安排id, n_分诊方式 From 挂号安排 Where 号码 = 号码_In;
    Exception
      When Others Then
        n_安排id := -1;
    End;
  
    If n_安排id = -1 Then
      v_Err_Msg := '号码(' || 号码_In || ')未找到!';
      Raise Err_Item;
    End If;
    --0-不分诊、1-指定诊室、2-动态分诊、3-平均分诊,对应门诊诊室设置
    v_诊室 := Null;
    If n_分诊方式 = 1 Then
      --1-指定诊室
      Begin
        Select 门诊诊室 Into v_诊室 From 挂号安排诊室 Where 号表id = n_安排id;
      Exception
        When Others Then
          v_诊室 := Null;
      End;
    End If;
    If n_分诊方式 = 2 Then
      --2-动态分诊:该个号别当天挂号未诊数最少的诊室
      For c_诊室 In (Select 门诊诊室, Sum(Num) As Num
                   From (Select 门诊诊室, 0 As Num
                          From 挂号安排诊室
                          Where 号表id = n_安排id
                          Union All
                          Select 诊室, Count(诊室) As Num
                          From 病人挂号记录
                          Where Nvl(执行状态, 0) = 0 And 发生时间 Between Trunc(Sysdate) And Sysdate And 号别 = 号码_In And
                                诊室 In (Select 门诊诊室 From 挂号安排诊室 Where 号表id = n_安排id)
                          Group By 诊室)
                   Group By 门诊诊室
                   Order By Num) Loop
        v_诊室 := c_诊室.门诊诊室;
        Exit;
      End Loop;
    End If;
    If n_分诊方式 = 3 Then
    
      --平均分诊：当前分配=1表示下次应取的当前诊室
      n_Next  := 0;
      n_First := 1;
      For c_诊室 In (Select Rowid As Rid, 号表id, 门诊诊室, 当前分配 From 挂号安排诊室 Where 号表id = n_安排id) Loop
        If n_First = 1 Then
          v_Rowid := c_诊室.Rid;
        End If;
        If n_Next = 1 Then
          v_诊室 := c_诊室.门诊诊室;
          Update 挂号安排诊室 Set 当前分配 = 1 Where Rowid = c_诊室.Rid;
          Exit;
        End If;
        If Nvl(c_诊室.当前分配, 0) = 1 Then
          Update 挂号安排诊室 Set 当前分配 = 0 Where Rowid = c_诊室.Rid;
          n_Next := 1;
        End If;
      End Loop;
      If v_诊室 Is Null Then
        Update 挂号安排诊室 Set 当前分配 = 1 Where Rowid = v_Rowid Returning 门诊诊室 Into v_诊室;
      End If;
    End If;
  
    Return v_诊室;
  End;

  Function Zl_操作员
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-获取缺省部门ID;1-获取操作员编号;2-获取操作员姓名
    -- SplitStr:格式为:部门ID,部门名称;人员ID,人员编号,人员姓名(用Zl_Identity获取的)
  Begin
    If Type_In = 0 Then
      --缺省部门
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --操作员编码
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --操作员姓名
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      Return v_Sub;
    End If;
  End;

  Procedure Zl_三方机构挂号_出诊_Insert
  (
    记录id_In        临床出诊记录.Id%Type,
    操作方式_In      Integer,
    病人id_In        门诊费用记录.病人id%Type,
    号码_In          挂号安排.号码%Type,
    号序_In          挂号序号状态.序号%Type,
    单据号_In        门诊费用记录.No%Type,
    票据号_In        门诊费用记录.实际票号%Type,
    结算方式_In      Varchar2,
    摘要_In          门诊费用记录.摘要%Type, --预约挂号摘要信息
    发生时间_In      门诊费用记录.发生时间%Type,
    登记时间_In      门诊费用记录.登记时间%Type,
    合作单位_In      挂号合作单位.名称%Type,
    挂号金额合计_In  门诊费用记录.实收金额%Type,
    领用id_In        票据使用明细.领用id%Type,
    收费票据_In      Number := 0, --挂号是否使用收费票据
    交易流水号_In    病人预交记录.交易流水号%Type,
    交易说明_In      病人预交记录.交易说明%Type,
    预约方式_In      预约方式.名称%Type := Null,
    预交id_In        病人预交记录.Id%Type := Null,
    卡类别id_In      病人预交记录.卡类别id%Type := Null,
    加入序号状态_In  Number := 0,
    是否自助设备_In  Number := 0,
    结帐id_In        门诊费用记录.结帐id%Type := Null,
    锁定类型_In      Number := 0,
    保险结算_In      Varchar2 := Null,
    冲预交_In        Number := Null,
    支付卡号_In      病人预交记录.卡号%Type := Null,
    费别_In          门诊费用记录.费别%Type := Null,
    冲预交病人ids_In Varchar2 := Null,
    机器名_In        挂号序号状态.机器名%Type := Null,
    更新年龄_In      Number := 0,
    购买病历_In      Number := 0,
    记帐费用_In      Number := 0,
    付款方式_In      医疗付款方式.名称%Type := Null
  ) As
    --功能：三方机构进行挂号(包含预约;预约挂号不扣款;预约挂号扣款),出诊表排班模式下使用
    --入参: 操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
    --      加入序号状态_In:1-表示强制加入挂号序号状态表中;否则根据启用序号或启用时段时才加入.
    --      是否自助设备_In:1-表示是医院的自助设备进行此函数的调用,自助设备调用此函数 允许加号,否则不允许
    --      锁定类型_In :0-产生正常数据 1-表示对单据进行锁定,产生未生效的单据信息;2-对锁定的记录进行解锁-生成正常数据:未生效的单据在银行扣款完成后进行解锁
    --      保险结算_IN:格式="结算方式|结算金额||....."
    --      冲预交病人ids_In:多个用逗分分离,冲预交时有效(冲预交类别与业务操作保持一致),主要是使用家属的预交款
    Err_Item Exception;
    Err_Special Exception;
    v_Err_Msg  Varchar2(255);
    n_打印id   票据打印内容.Id%Type;
    n_返回值   病人预交记录.金额%Type;
    v_排队号码 Varchar2(20);
    v_队列名称 排队叫号队列.队列名称%Type;
    n_预交id   病人预交记录.Id%Type;
    n_挂号id   病人挂号记录.Id%Type;
    v_结算内容 Varchar2(3000);
    v_当前结算 Varchar2(150);
  
    v_结算方式           病人预交记录.结算方式%Type;
    n_结算金额           病人预交记录.冲预交%Type;
    n_结算合计           Number(16, 5);
    n_预交金额           病人预交记录.冲预交%Type;
    n_组id               财务缴款分组.Id%Type;
    d_排队时间           Date;
    n_锁定               Number;
    n_病人预约科室数     Number(18);
    n_已约科室           Number(18);
    d_发生时间           Date;
    n_合作单位限制       Number(18);
    n_是否开放           Number(1);
    n_Count              Number(18);
    n_行号               Number(18);
    n_费用id             门诊费用记录.Id%Type;
    n_价格父号           Number(18);
    n_原项目id           收费项目目录.Id%Type;
    n_原收入项目id       收费项目目录.Id%Type;
    v_诊室               病人挂号记录.诊室%Type;
    n_实收金额合计       门诊费用记录.实收金额%Type;
    n_开单部门id         门诊费用记录.开单部门id%Type;
    n_实收金额           门诊费用记录.实收金额%Type;
    n_应收金额           门诊费用记录.实收金额%Type;
    n_急诊               病人挂号记录.急诊%Type;
    n_结帐id             病人结帐记录.Id%Type;
    v_Temp               Varchar2(500);
    v_结算方式记录       Varchar2(1000);
    n_预约时段序号       Number;
    n_序号控制           临床出诊记录.是否序号控制%Type;
    n_限约数             临床出诊记录.限约数%Type;
    n_项目id             临床出诊记录.项目id%Type;
    n_科室id             临床出诊记录.科室id%Type;
    d_终止时间           临床出诊记录.终止时间%Type;
    v_医生姓名           临床出诊记录.医生姓名%Type;
    n_医生id             临床出诊记录.医生id%Type;
    n_预约顺序号         临床出诊序号控制.预约顺序号%Type;
    n_预约总数           Number;
    d_时段开始时间       Date;
    d_时段终止时间       Date;
    v_收费项目ids        Varchar2(300);
    n_三方卡标志         Number;
    n_号序               病人挂号记录.号序%Type;
    d_登记时间           Date;
    n_单笔金额           病人预交记录.冲预交%Type;
    v_结算号码           病人预交记录.结算号码%Type;
    v_操作员编号         人员表.编号%Type;
    v_操作员姓名         人员表.姓名%Type;
    n_预约               Integer;
    n_分时点显示         Number;
    v_现金               病人预交记录.结算方式%Type;
    n_启用分时段         Integer;
    n_已挂数             病人挂号汇总.已挂数%Type;
    n_已约数             病人挂号汇总.已约数%Type;
    n_其中已接收         病人挂号汇总.已约数%Type;
    n_预约生成队列       Number;
    n_限号数             临床出诊记录.限号数%Type;
    d_Date               Date;
    n_挂号序号           Number;
    v_排队序号           排队叫号队列.排队序号%Type;
    v_机器名             挂号序号状态.机器名%Type;
    v_序号操作员         挂号序号状态.操作员姓名%Type;
    v_序号机器名         挂号序号状态.机器名%Type;
    n_序号锁定           Number := 0;
    n_病历费id           收费特定项目.收费细目id%Type;
    v_付款方式           病人挂号记录.医疗付款方式%Type;
    v_费别               门诊费用记录.费别%Type;
    n_屏蔽费别           Number(3) := 0;
    v_年龄               病人信息.年龄%Type;
    n_合作单位限数量模式 Number;
    n_同科限号数         Number;
    n_同科限约数         Number;
    n_同源限号数         Number;
    n_病人挂号科室数     Number;
    n_Exists             Number(5);
    v_Exists             Varchar2(4000);
    v_冲预交病人ids      Varchar2(4000);
    n_替诊医生id         临床出诊记录.替诊医生id%Type;
    v_替诊医生姓名       临床出诊记录.替诊医生姓名%Type;
    d_替诊开始时间       临床出诊记录.替诊开始时间%Type;
    d_替诊终止时间       临床出诊记录.替诊终止时间%Type;
    n_专家号挂号限制     Number;
    n_专家号预约限制     Number;
    v_站点               部门表.站点%Type;
    v_普通等级           Varchar2(100);
    v_Pricegrade         Varchar2(500);
    v_传入               Varchar2(100);
    n_更新项目id         挂号安排.项目id%Type;
  
    Cursor c_Pati(n_病人id 病人信息.病人id%Type) Is
      Select a.病人id, a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, a.费别, a.险类, c.编码 As 付款方式, a.出生日期, a.身份证号
      From 病人信息 A, 医疗付款方式 C
      Where a.病人id = n_病人id And a.医疗付款方式 = c.名称(+);
  
    r_Pati c_Pati%RowType;
  
    --该游标用于收费冲预交的可用预交列表
    --以ID排序，优先冲上次未冲完的。
    Cursor c_Deposit
    (
      v_病人id        病人信息.病人id%Type,
      v_冲预交病人ids Varchar2
    ) Is
      Select 病人id, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
             Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id
      From 病人预交记录
      Where 记录性质 In (1, 11) And 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1
       Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
      Group By NO, 病人id
      Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 结帐id, NO;
  
    Function Zl_诊室(记录id_In 临床出诊记录.Id%Type) Return Varchar2 As
      n_分诊方式 临床出诊记录.分诊方式%Type;
      v_诊室     病人挂号记录.诊室%Type;
      v_Rowid    Varchar2(500);
      n_Next     Integer;
      n_First    Integer;
    Begin
    
      If 锁定类型_In = 2 Then
        --对单据进行解锁,首先检查是否存在锁定
        Select Count(Rowid)
        Into n_锁定
        From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
        If n_锁定 = 0 Then
          v_Err_Msg := '单据号为(' || 单据号_In || ')的单据,不存在或者已经被解锁!';
          Raise Err_Item;
        End If;
        Select Max(号序) Into n_号序 From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
      End If;
    
      Begin
        Select Nvl(分诊方式, 0) Into n_分诊方式 From 临床出诊记录 Where ID = 记录id_In;
      Exception
        When Others Then
          v_Err_Msg := '出诊记录(' || 记录id_In || ')未找到!';
          Raise Err_Item;
      End;
    
      --0-不分诊、1-指定诊室、2-动态分诊、3-平均分诊,对应门诊诊室设置
      v_诊室 := Null;
      If n_分诊方式 = 1 Then
        --1-指定诊室
        Begin
          Select b.名称 Into v_诊室 From 临床出诊诊室记录 A, 门诊诊室 B Where a.诊室id = b.Id And a.记录id = 记录id_In;
        Exception
          When Others Then
            v_诊室 := Null;
        End;
      End If;
      If n_分诊方式 = 2 Then
        --2-动态分诊:该个号别当天挂号未诊数最少的诊室
        For c_诊室 In (Select 门诊诊室, Sum(Num) As Num
                     From (Select b.名称 As 门诊诊室, 0 As Num
                            From 临床出诊诊室记录 A, 门诊诊室 B
                            Where a.诊室id = b.Id And a.记录id = 记录id_In
                            Union All
                            Select 诊室, Count(诊室) As Num
                            From 病人挂号记录
                            Where Nvl(执行状态, 0) = 0 And 发生时间 Between Trunc(Sysdate) And Sysdate And 号别 = 号码_In And
                                  诊室 In (Select d.名称
                                         From 临床出诊诊室记录 C, 门诊诊室 D
                                         Where c.诊室id = d.Id And c.记录id = 记录id_In)
                            Group By 诊室)
                     Group By 门诊诊室
                     Order By Num) Loop
          v_诊室 := c_诊室.门诊诊室;
          Exit;
        End Loop;
      End If;
      If n_分诊方式 = 3 Then
        --平均分诊：当前分配=1表示下次应取的当前诊室
        n_Next  := 0;
        n_First := 1;
        For c_诊室 In (Select a.Rowid As Rid, b.名称 As 门诊诊室, a.当前分配
                     From 临床出诊诊室记录 A, 门诊诊室 B
                     Where a.诊室id = b.Id And a.记录id = 记录id_In) Loop
          If n_First = 1 Then
            v_Rowid := c_诊室.Rid;
          End If;
          If n_Next = 1 Then
            v_诊室 := c_诊室.门诊诊室;
            Update 临床出诊诊室记录 Set 当前分配 = 1 Where Rowid = c_诊室.Rid;
            Exit;
          End If;
          If Nvl(c_诊室.当前分配, 0) = 1 Then
            Update 临床出诊诊室记录 Set 当前分配 = 0 Where Rowid = c_诊室.Rid;
            n_Next := 1;
          End If;
        End Loop;
        If v_诊室 Is Null Then
          Update 临床出诊诊室记录 Set 当前分配 = 1 Where Rowid = v_Rowid Returning 诊室id Into v_诊室;
          Select 名称 Into v_诊室 From 门诊诊室 Where ID = v_诊室;
        End If;
      End If;
      Return v_诊室;
    End;
  
    Function Zl_操作员
    (
      Type_In     Integer,
      Splitstr_In Varchar2
    ) Return Varchar2 As
      n_Step Number(18);
      v_Sub  Varchar2(1000);
      --Type_In:0-获取缺省部门ID;1-获取操作员编号;2-获取操作员姓名
      -- SplitStr:格式为:部门ID,部门名称;人员ID,人员编号,人员姓名(用Zl_Identity获取的)
    Begin
      If Type_In = 0 Then
        --缺省部门
        n_Step := Instr(Splitstr_In, ',');
        v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 1 Then
        --操作员编码
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 2 Then
        --操作员姓名
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        Return v_Sub;
      End If;
    End;
  
  Begin
    d_发生时间 := 发生时间_In;
  
    If d_发生时间 Is Null Then
      d_发生时间 := Sysdate;
    End If;
  
    If 付款方式_In Is Null Then
      Select Max(名称) Into v_付款方式 From 医疗付款方式 Where 缺省标志 = 1;
    Else
      Select Max(名称) Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In;
      If v_付款方式 Is Null Then
        v_付款方式 := 付款方式_In;
      End If;
    End If;
    v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
    Begin
      Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
    Exception
      When Others Then
        v_现金 := '现金';
    End;
  
    If 费别_In Is Null Then
      Select Zl_Custom_Getpatifeetype(1, 病人id_In) Into v_费别 From Dual;
    Else
      v_费别 := 费别_In;
    End If;
    If v_费别 Is Null Then
      n_屏蔽费别 := 1;
      Select 名称 Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
    End If;
    Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
  
    If 更新年龄_In = 1 Then
      Select Zl_Age_Calc(病人id_In) Into v_年龄 From Dual;
      If v_年龄 Is Not Null Then
        Update 病人信息 Set 年龄 = v_年龄 Where 病人id = 病人id_In;
      End If;
    End If;
    --获取当前机器名称
    If 机器名_In Is Not Null Then
      v_机器名 := 机器名_In;
    Else
      Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
    End If;
    n_实收金额合计 := 0;
  
    Select Count(*) + 1
    Into n_挂号序号
    From 病人挂号记录
    Where 出诊记录id = 记录id_In And 登记时间 Between Trunc(发生时间_In) And Trunc(发生时间_In + 1) - 1 / 24 / 60 / 60;
  
    If 登记时间_In Is Null Then
      d_登记时间 := Sysdate;
    Else
      d_登记时间 := 登记时间_In;
    End If;
    If Trunc(Sysdate) > Trunc(发生时间_In) Then
      v_Err_Msg := '不能挂以前的号(' || To_Char(发生时间_In, 'yyyy-mm-dd') || ')。';
      Raise Err_Item;
    End If;
  
    v_Temp           := Nvl(zl_GetSysParameter('病人同科限挂N个号', 1111), '0|0') || '|';
    n_同科限号数     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    n_同科限约数     := To_Number(Nvl(zl_GetSysParameter('病人同科限约N个号', 1111), '0'));
    n_病人预约科室数 := To_Number(Nvl(zl_GetSysParameter('病人预约科室数', 1111), '0'));
    n_病人挂号科室数 := To_Number(Nvl(zl_GetSysParameter('病人挂号科室限制', 1111), '0'));
    n_专家号挂号限制 := To_Number(Nvl(zl_GetSysParameter('专家号挂号限制'), '0'));
    n_专家号预约限制 := To_Number(Nvl(zl_GetSysParameter('专家号预约限制'), '0'));
    n_同源限号数     := To_Number(Nvl(zl_GetSysParameter('病人同一号源限挂N个号', 1111), '0'));
  
    --部门ID,部门名称;人员ID,人员编号,人员姓名
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Err_Msg := '当前操作人员未设置对应的人员关系,不能继续。';
      Raise Err_Item;
    End If;
    n_开单部门id := To_Number(Zl_操作员(0, v_Temp));
    v_操作员编号 := Zl_操作员(1, v_Temp);
    v_操作员姓名 := Zl_操作员(2, v_Temp);
    n_组id       := Zl_Get组id(v_操作员姓名);
  
    --支付宝并发提交检查
    Select Nvl(Max(1), 0)
    Into n_Exists
    From 病人挂号记录
    Where 病人id = 病人id_In And 号别 = 号码_In And 号序 = 号序_In And 操作员姓名 = v_操作员姓名 And Nvl(记录id_In, 0) = Nvl(出诊记录id, 0) And
          登记时间 > Sysdate - 0.01 And 记录状态 = 1 And 发生时间 = 发生时间_In;
    If n_Exists = 1 Then
      v_Err_Msg := '病人已经挂号,不能重复挂相同的号！';
      Raise Err_Special;
    End If;
  
    If 操作方式_In <> 1 Then
      --预约检查是否添加合作单位控制
      --如果设置了合作单位控制 则
      Begin
        Select 1
        Into n_合作单位限制
        From 临床出诊挂号控制记录
        Where 记录id = 记录id_In And 类型 = 1 And 性质 = 1 And 控制方式 <> 4 And Rownum < 2;
      Exception
        When Others Then
          n_合作单位限制 := 0;
      End;
    End If;
  
    If 操作方式_In <> 2 Then
      v_诊室 := Zl_诊室(记录id_In);
    End If;
  
    --因为现在有按日编单据号规则,日挂号量不能超过10000次,所以要检查唯一约束。
    Select Count(*) Into n_Count From 门诊费用记录 Where 记录性质 = 4 And 记录状态 In (1, 3) And NO = 单据号_In;
    If n_Count <> 0 Then
      v_Err_Msg := '挂号单据号重复,不能保存！' || Chr(13) || '如果使用了按日顺序编号,当日挂号量不能超过10000人次。';
      Raise Err_Item;
    End If;
  
    Open c_Pati(病人id_In);
    n_Count := 0;
    Begin
      Fetch c_Pati
        Into r_Pati;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '病人未找到，不能继续。';
      Raise Err_Item;
    End If;
  
    Begin
      Select Nvl(是否分时段, 0), 限号数, 已挂数, 其中已接收, 已约数, 是否序号控制, 限约数, 项目id, 科室id, 医生id, 医生姓名, 替诊医生id, 替诊医生姓名, 替诊开始时间, 替诊终止时间
      Into n_启用分时段, n_限号数, n_已挂数, n_其中已接收, n_已约数, n_序号控制, n_限约数, n_项目id, n_科室id, n_医生id, v_医生姓名, n_替诊医生id, v_替诊医生姓名,
           d_替诊开始时间, d_替诊终止时间
      From 临床出诊记录
      Where ID = 记录id_In And Nvl(是否锁定, 0) = 0;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '该号别没有在' || To_Char(发生时间_In, 'yyyy-mm-dd hh24:mi:ss') || '中进行安排。';
      Raise Err_Item;
    End If;
  
    Select Min(站点) Into v_站点 From 部门表 Where ID = n_科室id;
    v_Pricegrade := Zl_Get_Pricegrade(v_站点, 病人id_In, Null, v_付款方式);
    v_普通等级   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);
  
    If 发生时间_In Between Nvl(d_替诊开始时间, Sysdate) And Nvl(d_替诊终止时间, Sysdate - 1) And v_替诊医生姓名 Is Not Null Then
      n_医生id   := n_替诊医生id;
      v_医生姓名 := v_替诊医生姓名;
    End If;
  
    --对参数控制进行检查
    --仅在预约不扣款时进行检查
    If 操作方式_In = 2 Then
      If Nvl(n_同科限约数, 0) <> 0 Or Nvl(n_病人预约科室数, 0) <> 0 Then
        n_已约科室 := 0;
        For c_Chkitem In (Select Distinct 执行部门id
                          From 病人挂号记录
                          Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
                                Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> n_科室id) Loop
          n_已约科室 := n_已约科室 + 1;
        End Loop;
        If n_已约科室 >= Nvl(n_病人预约科室数, 0) And Nvl(n_病人预约科室数, 0) > 0 Then
          v_Err_Msg := '同一病人最多同时能预约[' || Nvl(n_病人预约科室数, 0) || ']个科室,不能再预约！';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id = n_科室id;
        If n_Count >= Nvl(n_同科限约数, 0) And Nvl(n_同科限约数, 0) > 0 Then
          v_Err_Msg := '该病人已经在该科室预约了' || n_Count || '次,不能再预约！';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_专家号预约限制, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 出诊记录id = 记录id_In;
        If n_Count >= Nvl(n_专家号预约限制, 0) And Nvl(n_专家号预约限制, 0) > 0 Then
          v_Err_Msg := '该病人已经超过本号预约限制,不能再预约！';
          Raise Err_Item;
        End If;
      End If;
    Else
      If Nvl(n_同科限号数, 0) <> 0 Or Nvl(n_病人挂号科室数, 0) <> 0 Then
        n_已约科室 := 0;
        For c_Chkitem In (Select Distinct 执行部门id
                          From 病人挂号记录
                          Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And
                                Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> n_科室id) Loop
          n_已约科室 := n_已约科室 + 1;
        End Loop;
        If n_已约科室 >= Nvl(n_病人挂号科室数, 0) And Nvl(n_病人挂号科室数, 0) > 0 Then
          v_Err_Msg := '同一病人最多同时能挂号[' || Nvl(n_病人挂号科室数, 0) || ']个科室,不能再挂号！';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id = n_科室id;
        If n_Count >= Nvl(n_同科限号数, 0) And Nvl(n_同科限号数, 0) > 0 Then
          v_Err_Msg := '该病人已经在该科室挂号了' || n_Count || '次,不能再挂号！';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_专家号挂号限制, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 出诊记录id = 记录id_In;
        If n_Count >= Nvl(n_专家号挂号限制, 0) And Nvl(n_专家号挂号限制, 0) > 0 Then
          v_Err_Msg := '该病人已经超过本号挂号限制,不能再挂号！';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    If Nvl(n_同源限号数, 0) <> 0 Then
      If 出诊记录id_In Is Null Then
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 In (1, 2) And 发生时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 号别 = 号码_In;
      Else
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 In (1, 2) And 出诊记录id = 出诊记录id_In;
      End If;
      If n_Count >= Nvl(n_同源限号数, 0) And Nvl(n_同源限号数, 0) > 0 Then
        v_Err_Msg := '同一病人最多能同时挂(预约)[' || Nvl(n_同源限号数, 0) || ']个相同号别的号,不能再挂号(预约)！';
        Raise Err_Item;
      End If;
    End If;
  
    d_Date         := Null;
    d_时段开始时间 := Null;
  
    If Nvl(n_限号数, 0) >= 0 Or n_限号数 Is Null Then
      If n_启用分时段 = 1 Then
        If Nvl(n_序号控制, 0) = 1 Then
          If Nvl(是否自助设备_In, 0) = 0 Then
            Select Count(*), Max(开始时间)
            Into n_Count, d_时段开始时间
            From 临床出诊序号控制
            Where 记录id = 记录id_In And 序号 = Nvl(号序_In, 0);
          
            v_Temp := '挂号';
            If 操作方式_In > 1 Then
              v_Temp := '预约挂号';
            End If;
          
            If n_Count = 0 Then
              v_Err_Msg := '号别为' || 号码_In || '的挂号安排中不存在序号为' || Nvl(号序_In, 0) || '的安排,不能再' || v_Temp || '！';
              Raise Err_Item;
            End If;
          End If;
        
          --过点的,不能选择挂号
          If Trunc(Sysdate) = Trunc(发生时间_In) Then
            --挂当天的号
            v_Temp := To_Char(Sysdate, 'yyyy-mm-dd') || ' ';
            For v_时段 In (Select To_Date(v_Temp || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                To_Date(To_Char(Sysdate + Decode(Sign(开始时间 - 终止时间), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                         To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 终止时间, 数量, 是否预约
                         From 临床出诊序号控制
                         Where 记录id = 记录id_In And 序号 = Nvl(号序_In, 0)) Loop
              If Sysdate > v_时段.终止时间 Then
                v_Err_Msg := '号别为' || 号码_In || '及号序为' || Nvl(号序_In, 0) || '的安排,已经超过时点,不能再' || v_Temp || '！';
                Raise Err_Item;
              End If;
            End Loop;
          End If;
        Elsif 操作方式_In > 1 Then
          --未启用序号的,需要检查预约的情况
          n_Count := 0;
          For v_时段 In (Select 序号, 开始时间, 终止时间, 数量, 是否预约
                       From 临床出诊序号控制
                       Where 记录id = 记录id_In And
                             (('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                             Decode(Sign(开始时间 - 终止时间 - 1 / 24 / 60 / 60), 1,
                                      '3000-01-09 ' || To_Char(开始时间, 'HH24:MI:SS'),
                                      '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS')) And
                             '3000-01-10 ' || To_Char(终止时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                             ('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                             '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS') And
                             Decode(Sign(开始时间 - 终止时间 - 1 / 24 / 60 / 60), 1,
                                      '3000-01-11 ' || To_Char(终止时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                      '3000-01-10 ' || To_Char(终止时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
            n_预约时段序号 := v_时段.序号;
            d_时段开始时间 := v_时段.开始时间;
            d_时段终止时间 := v_时段.终止时间;
          
            Select Count(*), Max(序号), Max(预约顺序号) + 1
            Into n_Count, n_预约总数, n_预约顺序号
            From 临床出诊序号控制
            Where 记录id = 记录id_In And Nvl(挂号状态, 0) Not In (0, 4, 5);
          
            If Nvl(n_Count, 0) > Nvl(v_时段.数量, 0) And 锁定类型_In <> 2 Then
              v_Err_Msg := '号别为' || 号码_In || '的挂号安排中在' || To_Char(v_时段.开始时间, 'hh24:mi:ss') || '至' ||
                           To_Char(v_时段.终止时间, 'hh24:mi:ss') || '最多只能预约' || Nvl(v_时段.数量, 0) || '人,不能再进行预约挂号！';
              Raise Err_Item;
            End If;
            n_Count := 1;
          End Loop;
        
          If n_Count = 0 Then
            v_Err_Msg := '号别为' || 号码_In || '的挂号安排中没有相关的安排计划(' || To_Char(发生时间_In, 'YYYY-mm-dd HH24:MI:SS') ||
                         '),不能进行预约挂号！';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End If;
  
    If 操作方式_In = 1 And 锁定类型_In <> 2 Then
      --挂号规则:
      --  已挂数不能大于限号数
      If n_已挂数 >= Nvl(n_限号数, 0) And n_限号数 Is Not Null Then
        v_Err_Msg := '该号别今天已达到限号数 ' || Nvl(n_限号数, 0) || '不能再挂号！';
        Raise Err_Item;
      End If;
    End If;
  
    If 操作方式_In > 1 Then
      --预约的相关检查
      --规则:
      --   1.已限约不能超过限约数
      --   2.检查是否启用时段的
      If n_已约数 >= Nvl(n_限约数, 0) And Nvl(n_限约数, 0) <> 0 And n_限约数 Is Not Null And 锁定类型_In <> 2 Then
        v_Err_Msg := '该号别已达到限约数 ' || Nvl(n_限约数, 0) || '不能再预约挂号！';
        Raise Err_Item;
      End If;
      If 预约方式_In Is Not Null Then
        Select Zl_Fun_Get临床出诊预约状态(记录id_In, 发生时间_In, 号序_In, 预约方式_In, Null, 0, v_操作员姓名, v_机器名)
        Into v_Exists
        From Dual;
        If To_Number(Substr(v_Exists, 1, 1)) <> 0 Then
          v_Err_Msg := '传入的预约方式' || 预约方式_In || '不可用,原因:' || Substr(v_Exists, 3);
          Raise Err_Item;
        End If;
      End If;
    End If;
    If n_合作单位限制 > 0 And 操作方式_In <> 1 And 合作单位_In Is Not Null Then
      If Nvl(n_序号控制, 0) = 1 And Nvl(号序_In, 0) = 0 Then
        v_Err_Msg := '当前安排使用了序号控制,请确认所需要预约的序号,不能继续。';
        Raise Err_Item;
      End If; --Nvl(r_安排.序号控制, 0) =0
    
      --合作单位控制模式
      Begin
        Select Nvl(控制方式, 0)
        Into n_合作单位限数量模式
        From 临床出诊挂号控制记录
        Where 记录id = 记录id_In And 名称 = 合作单位_In And 类型 = 1 And 性质 = 1 And Rownum < 2;
      Exception
        When Others Then
          n_合作单位限数量模式 := 4;
      End;
    
      If n_合作单位限数量模式 = 0 Then
        v_Err_Msg := '当前号码(' || Nvl(号码_In, 0) || '未开放' || 合作单位_In || '的预约,不能继续。';
        Raise Err_Item;
      End If;
      If n_合作单位限数量模式 = 1 Or n_合作单位限数量模式 = 2 Then
        Select 数量
        Into n_Count
        From 临床出诊挂号控制记录
        Where 记录id = 记录id_In And 名称 = 合作单位_In And 类型 = 1 And 性质 = 1;
        If n_合作单位限数量模式 = 1 Then
          n_Count := Round(Nvl(n_限约数, n_限号数) * n_Count / 100);
        End If;
        Select Count(1)
        Into n_Exists
        From 病人挂号记录
        Where 记录状态 = 1 And 出诊记录id = 记录id_In And 合作单位 = 合作单位_In;
        If n_Exists >= n_Count Then
          v_Err_Msg := '当前号码(' || Nvl(号码_In, 0) || '已经超过' || 合作单位_In || '的预约限制,不能继续。';
          Raise Err_Item;
        End If;
      End If;
      --开放序号检查
      If n_合作单位限数量模式 = 3 Then
        For c_合作单位 In (Select 序号, 数量
                       From 临床出诊挂号控制记录
                       Where 记录id = 记录id_In And 名称 = 合作单位_In And 类型 = 1 And 性质 = 1 And 序号 = 号序_In) Loop
          If n_序号控制 = 1 Then
            Begin
              Select 1
              Into n_Count
              From 临床出诊序号控制
              Where 记录id = 记录id_In And 序号 = 号序_In And Nvl(挂号状态, 0) = 0;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count = 1 Then
              n_是否开放 := 1;
            Else
              v_Err_Msg := '当前序号(' || Nvl(号序_In, 0) || '已经超过' || 合作单位_In || '的预约限制,不能继续。';
              Raise Err_Item;
            End If;
          Else
            Select Count(1)
            Into n_Count
            From 临床出诊序号控制
            Where 记录id = 记录id_In And 序号 = 号序_In And 预约顺序号 Is Not Null And Nvl(挂号状态, 0) <> 0;
            If n_Count >= c_合作单位.数量 Then
              v_Err_Msg := '当前序号(' || Nvl(号序_In, 0) || '已经超过' || 合作单位_In || '的预约限制,不能继续。';
              Raise Err_Item;
            Else
              n_是否开放 := 1;
            End If;
          End If;
        End Loop;
      
        If Nvl(n_是否开放, 0) = 0 Then
          v_Err_Msg := '当前序号(' || Nvl(号序_In, 0) || '未开放,不能继续。';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    --检查限号数和限约数
    n_行号         := 1;
    n_原项目id     := 0;
    n_原收入项目id := 0;
    n_实收金额合计 := 0;
    If 锁定类型_In <> 1 Then
      If 操作方式_In <> 2 Then
        If Nvl(结帐id_In, 0) = 0 Then
          --这里应该程序传入
          Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
        Else
          n_结帐id := 结帐id_In;
        End If;
      Else
        n_结帐id := Null;
      End If;
    End If;
  
    If Nvl(记录id_In, 0) <> 0 Then
      v_传入 := '2|' || 记录id_In;
    End If;
    If v_传入 Is Null Then
      v_传入 := '3|' || 号码_In;
    End If;
  
    n_更新项目id := Zl_Custom_Getregeventitem(r_Pati.病人id, r_Pati.姓名, r_Pati.身份证号, r_Pati.出生日期, r_Pati.性别, r_Pati.年龄, v_传入);
    If Nvl(n_更新项目id, 0) <> 0 Then
      n_项目id := n_更新项目id;
    End If;
    If Nvl(购买病历_In, 0) = 1 Then
      Begin
        Select 收费细目id Into n_病历费id From 收费特定项目 Where 特定项目 = '病历费';
        v_收费项目ids := n_项目id || ',' || n_病历费id;
      Exception
        When Others Then
          v_Err_Msg := '不能确定病历费,挂号失败!';
          Raise Err_Item;
      End;
    Else
      v_收费项目ids := n_项目id;
    End If;
  
    For c_Item In (Select 1 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, Null As 从属父号,
                          Nvl(a.项目特性, 0) As 急诊
                   From 收费项目目录 A, 收费价目 B, 收入项目 C
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = n_项目id And d_发生时间 Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.价格等级 = v_普通等级 Or
                         (b.价格等级 Is Null And Not Exists
                          (Select 1
                            From 收费价目
                            Where b.收费细目id = 收费细目id And 价格等级 = v_普通等级 And d_发生时间 Between 执行日期 And
                                  Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Union All
                   Select 2 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, Null As 从属父号, 0 As 急诊
                   From 收费项目目录 A, 收费价目 B, 收入项目 C
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = n_病历费id And d_发生时间 Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.价格等级 = v_普通等级 Or
                         (b.价格等级 Is Null And Not Exists
                          (Select 1
                            From 收费价目
                            Where b.收费细目id = 收费细目id And 价格等级 = v_普通等级 And d_发生时间 Between 执行日期 And
                                  Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Union All
                   Select 3 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, d.从项数次 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, 1 As 从属父号, 0 As 急诊
                   From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And
                         d.主项id In (Select Column_Value From Table(f_Str2list(v_收费项目ids))) And d_发生时间 Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.价格等级 = v_普通等级 Or
                         (b.价格等级 Is Null And Not Exists
                          (Select 1
                            From 收费价目
                            Where b.收费细目id = 收费细目id And 价格等级 = v_普通等级 And d_发生时间 Between 执行日期 And
                                  Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Order By 性质, 项目编码, 收入编码) Loop
      If c_Item.性质 = 1 Then
        n_急诊 := Nvl(c_Item.急诊, 0);
      End If;
      n_价格父号 := Null;
      If n_原项目id = c_Item.项目id Then
        If n_原收入项目id <> c_Item.收入项目id Then
          n_价格父号 := n_行号;
        End If;
        n_原收入项目id := c_Item.收入项目id;
      End If;
      n_原项目id := c_Item.项目id;
      n_应收金额 := Round(c_Item.数次 * c_Item.单价, 5);
      n_实收金额 := n_应收金额;
      If Nvl(c_Item.屏蔽费别, 0) <> 1 And n_屏蔽费别 = 0 Then
        --打折:
        v_Temp     := Zl_Actualmoney(r_Pati.费别, c_Item.项目id, c_Item.收入项目id, n_应收金额);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ':') + 1);
        n_实收金额 := Zl_To_Number(v_Temp);
      End If;
      n_实收金额合计 := Nvl(n_实收金额合计, 0) + n_实收金额;
    
      --锁定单据不产生费用
      If 锁定类型_In <> 1 Then
        --产生病人挂号费用(可能单独是或包括病历费用)
        Select 病人费用记录_Id.Nextval Into n_费用id From Dual;
        --:操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
        Insert Into 门诊费用记录
          (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id,
           收费类别, 计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人,
           操作员编号, 操作员姓名, 发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
        Values
          (n_费用id, 4, Decode(操作方式_In, 2, 0, 1), n_行号, n_价格父号, c_Item.从属父号, 单据号_In, 票据号_In, 1, n_急诊, Null,
           Decode(操作方式_In, 2, To_Char(号序_In), v_诊室), r_Pati.病人id, r_Pati.门诊号, r_Pati.付款方式, r_Pati.姓名, r_Pati.性别,
           r_Pati.年龄, r_Pati.费别, n_科室id, c_Item.类别, 号码_In, c_Item.项目id, c_Item.收入项目id, c_Item.收据费目, 1, c_Item.数次,
           c_Item.单价, n_应收金额, n_实收金额, Decode(操作方式_In, 2, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, n_实收金额)),
           Decode(Nvl(记帐费用_In, 0), 1, Null, n_结帐id), Decode(Nvl(记帐费用_In, 0), 1, 1, 0), n_开单部门id, v_操作员姓名,
           Decode(操作方式_In, 2, v_操作员姓名, Null), n_科室id, v_医生姓名, v_操作员编号, v_操作员姓名, 发生时间_In, d_登记时间, Null, 0, Null, Null,
           摘要_In, 预约方式_In, Decode(操作方式_In, 2, Null, n_组id));
      End If;
      n_行号 := n_行号 + 1;
    
    End Loop;
  
    If Round(Nvl(挂号金额合计_In, 0), 5) <> Round(Nvl(n_实收金额合计, 0), 5) Then
      v_Err_Msg := '本次挂号金额不正确,可能是因为医院调整了价格,请重新获取挂号收费项目的价格,不能继续。';
      Raise Err_Item;
    End If;
  
    If n_启用分时段 = 1 Then
      d_Date := To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(d_时段开始时间, 'hh24:mi:ss'),
                        'yyyy-mm-dd hh24:mi:ss');
    Else
      d_Date := Trunc(发生时间_In);
    End If;
  
    --更新挂号序号状态
    If 锁定类型_In <> 2 Then
      n_号序 := 号序_In;
    End If;
    Begin
      Select 1
      Into n_Count
      From 临床出诊序号控制
      Where 记录id = 记录id_In And 序号 = n_号序 And Nvl(挂号状态, 0) Not In (0, 5);
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count = 1 Then
      If n_启用分时段 = 0 And Nvl(n_序号控制, 0) = 1 Then
        n_号序 := Null;
      End If;
      If n_启用分时段 = 1 And Nvl(n_序号控制, 0) = 1 Then
        v_Err_Msg := '当前序号已被使用，请重新选择一个序号！';
        Raise Err_Item;
      End If;
    End If;
  
    If n_启用分时段 = 0 And Nvl(n_序号控制, 0) = 1 And n_号序 Is Null And 锁定类型_In <> 2 Then
      Select Nvl(Min(序号), 0)
      Into n_号序
      From 临床出诊序号控制
      Where 记录id = 记录id_In And Nvl(挂号状态, 0) = 5 And 操作员姓名 = v_操作员姓名 And 工作站名称 = v_机器名;
      If n_号序 = 0 Then
        Select Nvl(Min(序号), 0) Into n_号序 From 临床出诊序号控制 Where 记录id = 记录id_In And Nvl(挂号状态, 0) = 0;
        If n_号序 = 0 Then
          Select Nvl(Max(序号), 0) + 1 Into n_号序 From 临床出诊序号控制 Where 记录id = 记录id_In;
        End If;
      End If;
    End If;
  
    If n_启用分时段 = 1 And 锁定类型_In <> 2 Then
      If 操作方式_In > 1 And Nvl(n_序号控制, 0) = 0 Then
        --规则:预约时段序号||预约数
        If Nvl(n_预约总数, 0) = 0 Then
          v_Temp := Nvl(n_限约数, 0);
          v_Temp := LTrim(RTrim(v_Temp));
          v_Temp := LPad(Nvl(n_预约总数, 0) + 1, Length(v_Temp), '0');
          v_Temp := n_预约时段序号 || v_Temp;
          n_号序 := To_Number(v_Temp);
        Else
          n_号序 := n_预约总数 + 1;
        End If;
      End If;
    End If;
  
    If Nvl(n_序号控制, 0) = 1 Or (操作方式_In > 1 And n_启用分时段 = 1) Or 加入序号状态_In = 1 Then
      --锁定序号的处理
      Begin
        Select 操作员姓名, 工作站名称
        Into v_序号操作员, v_序号机器名
        From 临床出诊序号控制
        Where 挂号状态 = 5 And 记录id = 记录id_In And 序号 = n_号序;
        n_序号锁定 := 1;
      Exception
        When Others Then
          v_序号操作员 := Null;
          v_序号机器名 := Null;
          n_序号锁定   := 0;
      End;
      If n_序号锁定 = 0 Then
        If n_启用分时段 = 1 And n_序号控制 = 0 Then
          Insert Into 临床出诊序号控制
            (记录id, 序号, 预约顺序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 类型, 名称, 操作员姓名, 备注)
            Select 记录id_In, n_预约时段序号, n_预约顺序号, d_时段开始时间, d_时段终止时间, 1, Decode(操作方式_In, 1, 0, 1), Decode(操作方式_In, 2, 2, 1),
                   1, 合作单位_In, v_操作员姓名, n_号序
            From Dual;
        Else
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(操作方式_In, 2, 2, 1), 操作员姓名 = v_操作员姓名
          Where 记录id = 记录id_In And 序号 = n_号序;
        End If;
        If Sql%RowCount = 0 Then
          Begin
            If n_启用分时段 = 1 Then
              --分时段
              If n_序号控制 = 1 Then
                --序号控制
                Select Max(终止时间) Into d_终止时间 From 临床出诊序号控制 Where 记录id = 记录id_In;
                If Sysdate > d_终止时间 Then
                  d_终止时间 := Sysdate;
                End If;
                Insert Into 临床出诊序号控制
                  (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 类型, 名称, 操作员姓名)
                  Select 记录id_In, n_号序, d_终止时间, d_终止时间, 1, Decode(操作方式_In, 1, 0, 1), Decode(操作方式_In, 2, 2, 1), 1,
                         合作单位_In, v_操作员姓名
                  From Dual;
              Else
                --分时段,非序号控制
                Null;
              End If;
            Else
              --不分时段
              Insert Into 临床出诊序号控制
                (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 类型, 名称, 操作员姓名)
                Select 记录id_In, n_号序, 开始时间, 终止时间, 1, Decode(操作方式_In, 1, 0, 1), Decode(操作方式_In, 2, 2, 1), 1, 合作单位_In,
                       v_操作员姓名
                From 临床出诊序号控制
                Where 记录id = 记录id_In And 序号 = 1;
            End If;
          Exception
            When Others Then
              v_Err_Msg := '序号' || n_号序 || '已被使用,请重新选择一个序号.';
              Raise Err_Item;
          End;
        End If;
      Else
        If v_操作员姓名 <> v_序号操作员 Or v_机器名 <> v_序号机器名 Then
          v_Err_Msg := '序号' || n_号序 || '已被机器' || v_机器名 || '锁定,请重新选择一个序号.';
          Raise Err_Item;
        Else
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(操作方式_In, 2, 2, 1), 锁号时间 = Null
          Where 记录id = 记录id_In And 序号 = n_号序 And 挂号状态 = 5 And 操作员姓名 = v_操作员姓名 And 工作站名称 = v_机器名;
        End If;
      End If;
    End If;
  
    --锁定单据不产生任何 费用
    If 操作方式_In <> 2 And 锁定类型_In <> 1 And Nvl(记帐费用_In, 0) = 0 Then
      --挂号,预约挂号已经扣款部分
      n_预交id := 预交id_In;
      If Nvl(n_预交id, 0) = 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      End If;
      n_结算合计 := 0;
      If 保险结算_In Is Not Null Then
        --各个保险结算
        v_结算内容 := 保险结算_In || '||';
        n_结算合计 := 0;
        While v_结算内容 Is Not Null Loop
          v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
          v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
          n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
          If Nvl(n_结算金额, 0) <> 0 Then
            Insert Into 病人预交记录
              (ID, 记录性质, NO, 记录状态, 病人id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 结算性质)
            Values
              (n_预交id, 4, 单据号_In, 1, Decode(病人id_In, 0, Null, 病人id_In), '保险结算', v_结算方式, d_登记时间, v_操作员编号, v_操作员姓名,
               n_结算金额, n_结帐id, n_组id, n_结帐id, 4);
            Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          End If;
          n_结算合计 := Nvl(n_结算合计, 0) + Nvl(n_结算金额, 0);
          v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
        End Loop;
      End If;
    
      If Nvl(冲预交_In, 0) <> 0 Then
        --处理总预交
        n_结算合计 := n_结算合计 + Nvl(冲预交_In, 0);
        n_预交金额 := 冲预交_In;
        For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
          n_结算金额 := Case
                      When r_Deposit.金额 - n_预交金额 < 0 Then
                       r_Deposit.金额
                      Else
                       n_预交金额
                    End;
          If r_Deposit.结帐id = 0 Then
            --第一次冲预交(填上结帐ID,金额为0)
            Update 病人预交记录 Set 冲预交 = 0, 结帐id = n_结帐id, 结算性质 = 4 Where ID = r_Deposit.原预交id;
          End If;
          --冲上次剩余额
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名,
             操作员编号, 冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行,
                   单位帐号, d_登记时间, v_操作员姓名, v_操作员编号, n_结算金额, n_结帐id, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, n_结帐id, 4
            From 病人预交记录
            Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
        
          --更新病人预交余额
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) - n_结算金额
          Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2)
          Returning 预交余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 病人余额 (病人id, 预交余额, 性质, 类型) Values (r_Deposit.病人id, -1 * n_结算金额, 1, 1);
            n_返回值 := -1 * n_结算金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 病人余额
            Where 病人id = r_Deposit.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;
        
          --检查是否已经处理完
          If r_Deposit.金额 <= n_结算金额 Then
            n_预交金额 := n_预交金额 - r_Deposit.金额;
          Else
            n_预交金额 := 0;
          End If;
          If n_预交金额 = 0 Then
            Exit;
          End If;
        End Loop;
        If n_预交金额 > 0 Then
          v_Err_Msg := '预交余额不够支付本次支付金额,不能继续操作！';
          Raise Err_Item;
        End If;
      End If;
      --剩余款项,用指定结算方支付
      n_结算金额 := Nvl(n_实收金额合计, 0) - Nvl(n_结算合计, 0);
      If Nvl(n_结算金额, 0) < 0 Then
        v_Err_Msg := '挂号的相关结算金额超出了当前实结金额,不能继续操作！';
        Raise Err_Item;
      End If;
    
      If Nvl(n_结算金额, 0) <> 0 Or (Nvl(n_结算金额, 0) = 0 And Nvl(冲预交_In, 0) = 0) Then
        If 结算方式_In Is Null Then
          v_Err_Msg := '未传入指定的结算方式,不能继续操作！';
          Raise Err_Item;
        End If;
      
        If Nvl(预交id_In, 0) <> 0 Then
          --传入的预交ID_In主要是为了解决三方交易,如果医保结算站用了该ID,需要用新的ID进行更新,三方交易用转入的ID
          Update 病人预交记录 Set ID = n_预交id Where ID = Nvl(预交id_In, 0);
          n_预交id := Nvl(预交id_In, 0);
        End If;
        If Instr(结算方式_In, ',') = 0 Then
          --只传入一种结算方式的
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(结算方式_In, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             v_操作员编号, v_操作员姓名, n_结帐id, '挂号收费', n_组id, 卡类别id_In, Null, 支付卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4, v_结算号码);
        Else
          v_结算内容     := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
          n_Exists       := 0;
          v_结算方式记录 := '';
          While v_结算内容 Is Not Null Loop
            v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
            v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
          
            v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
            n_单笔金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
          
            v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
            v_结算号码 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
          
            v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
            n_三方卡标志 := To_Number(v_当前结算);
          
            If Instr('|' || v_结算方式记录 || '|', '|' || Nvl(v_结算方式, v_现金) || '|') <> 0 Then
              v_Err_Msg := '使用了重复的结算方式,请检查!';
              Raise Err_Item;
            Else
              v_结算方式记录 := v_结算方式记录 || '|' || Nvl(v_结算方式, v_现金);
            End If;
          
            If n_三方卡标志 = 0 Then
              Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
              Insert Into 病人预交记录
                (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号,
                 交易说明, 合作单位, 结算性质, 结算号码)
              Values
                (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_单笔金额, 0), 登记时间_In,
                 v_操作员编号, v_操作员姓名, n_结帐id, '挂号收费', n_组id, Null, Null, Null, Null, Null, 合作单位_In, 4, v_结算号码);
            Else
              If n_Exists = 1 Then
                v_Err_Msg := '目前挂号仅支持一种三方结算方式,不能继续操作！';
                Raise Err_Item;
              End If;
              Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
              Insert Into 病人预交记录
                (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号,
                 交易说明, 合作单位, 结算性质, 结算号码)
              Values
                (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_单笔金额, 0), 登记时间_In,
                 v_操作员编号, v_操作员姓名, n_结帐id, '挂号收费', n_组id, 卡类别id_In, Null, 支付卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4, v_结算号码);
              n_Exists := 1;
            End If;
          
            v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
          End Loop;
        End If;
      End If;
    
      --更新人员缴款数据
    
      For v_缴款 In (Select 结算方式, Sum(Nvl(a.冲预交, 0)) As 冲预交
                   From 病人预交记录 A
                   Where a.结帐id = n_结帐id And Mod(a.记录性质, 10) <> 1 And Nvl(病人id, 0) = Nvl(病人id_In, 0)
                   Group By 结算方式) Loop
      
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + Nvl(v_缴款.冲预交, 0)
        Where 收款员 = v_操作员姓名 And 性质 = 1 And 结算方式 = v_缴款.结算方式
        Returning 余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (v_操作员姓名, v_缴款.结算方式, 1, Nvl(v_缴款.冲预交, 0));
          n_返回值 := Nvl(v_缴款.冲预交, 0);
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = v_操作员姓名 And 结算方式 = v_缴款.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
        End If;
      End Loop;
    End If;
  
    --处理挂号记录
    If 锁定类型_In = 2 Then
      Begin
        Select ID Into n_挂号id From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
      Exception
        When Others Then
          Null;
      End;
    Else
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
    End If;
  
    Update 病人挂号记录
    Set 记录性质 = Decode(操作方式_In, 2, 2, 1), 记录状态 = Decode(锁定类型_In, 1, 0, 1), 门诊号 = r_Pati.门诊号, 操作员姓名 = v_操作员姓名,
        操作员编号 = v_操作员编号, 预约 = Decode(操作方式_In, 1, 0, 1),
        接收人 = Decode(锁定类型_In, 1, Null, Decode(操作方式_In, 2, Null, v_操作员姓名)),
        接收时间 = Case 锁定类型_In
                  When 1 Then
                   Null
                  Else
                   Case 操作方式_In
                     When 2 Then
                      Null
                     Else
                      d_登记时间
                   End
                End, 交易流水号 = Nvl(交易流水号_In, 交易流水号), 交易说明 = Nvl(交易说明_In, 交易说明), 合作单位 = Nvl(合作单位_In, 合作单位),
        预约操作员 = Decode(操作方式_In, 1, Nvl(预约操作员, Null), Nvl(预约操作员, v_操作员姓名)),
        预约操作员编号 = Decode(操作方式_In, 1, Nvl(预约操作员编号, Null), Nvl(预约操作员编号, v_操作员编号)), 出诊记录id = 记录id_In
    Where ID = n_挂号id;
    If Sql%NotFound Then
      Begin
        Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = r_Pati.付款方式 And Rownum < 2;
      Exception
        When Others Then
          v_付款方式 := Null;
      End;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 预约时间, 操作员编号,
         操作员姓名, 复诊, 号序, 预约, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式, 预约操作员, 预约操作员编号, 出诊记录id)
      Values
        (n_挂号id, 单据号_In, Decode(操作方式_In, 2, 2, 1), Decode(锁定类型_In, 1, 0, 1), r_Pati.病人id, r_Pati.门诊号, r_Pati.姓名,
         r_Pati.性别, r_Pati.年龄, 号码_In, n_急诊, v_诊室, Null, n_科室id, v_医生姓名, 0, Null, d_登记时间, 发生时间_In,
         Case When(Nvl(操作方式_In, 0)) = 1 Then Null Else 发生时间_In End, v_操作员编号, v_操作员姓名, 0, n_号序, Decode(操作方式_In, 1, 0, 1),
         Decode(操作方式_In, 2, Null, v_操作员姓名), Decode(操作方式_In, 2, To_Date(Null), d_登记时间), 交易流水号_In, 交易说明_In, 合作单位_In,
         v_付款方式, Decode(操作方式_In, 1, Null, v_操作员姓名), Decode(操作方式_In, 1, Null, v_操作员编号), 记录id_In);
    End If;
    --锁定单据不能产生队列
    If 锁定类型_In <> 1 Then
      n_预约生成队列 := 0;
      If 操作方式_In > 1 Then
        n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
      End If;
      --挂号和收费的预约都直接进入队列(收费预约缺少接收过程,所以直接和挂号一样直接进入队列)
      If 操作方式_In <> 2 Or n_预约生成队列 = 1 Then
        If Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113)) <> 0 Then
          --排队叫号模式:-0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
          If Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(n_科室id, 0))) = 0 Or n_预约生成队列 = 1 Then
            n_分时点显示 := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
            If Nvl(操作方式_In, 0) > 1 And n_分时点显示 = 1 And n_启用分时段 = 1 Then
              n_分时点显示 := 1;
            Else
              n_分时点显示 := Null;
            End If;
            --产生队列
            --.按”执行部门” 的方式生成队列
            v_队列名称 := n_科室id;
            v_排队号码 := Zlgetnextqueue(n_科室id, n_挂号id, 号码_In || '|' || 号序_In);
            --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
            v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
            --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
            d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号码_In, 号序_In, d_Date);
            --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,v_排队标记,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
            Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, n_科室id, v_排队号码, Null, r_Pati.姓名, r_Pati.病人id, v_诊室, v_医生姓名, d_排队时间,
                             预约方式_In, n_分时点显示, v_排队序号);
          End If;
        End If;
      End If;
    
      If Nvl(操作方式_In, 0) = 1 And Nvl(记帐费用_In, 0) = 0 Then
        --处理票据使用情况
        If 票据号_In Is Not Null Then
          Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
          --发出票据
          Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 4, 单据号_In);
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
          Values
            (票据使用明细_Id.Nextval, Decode(收费票据_In, 1, 1, 4), 票据号_In, 1, 1, 领用id_In, n_打印id, d_登记时间, v_操作员姓名, 挂号金额合计_In);
          --状态改动
          Update 票据领用记录
          Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
          Where ID = Nvl(领用id_In, 0);
        End If;
        --病人本次就诊(以发生时间为准)
        If Nvl(r_Pati.病人id, 0) <> 0 Then
          Update 病人信息 Set 就诊时间 = 发生时间_In, 就诊状态 = 1, 就诊诊室 = v_诊室 Where 病人id = r_Pati.病人id;
        End If;
      End If;
    End If;
    --病人挂号汇总
    --解锁单据时不用再对汇总单据进行统计了 在锁定单据时已经进行了汇总
    If 锁定类型_In <> 2 Then
      --操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
      --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收;3-收费预约
      --调用zl_third_lockno进行锁号，不建议使用本过程锁号
      n_预约 := Case
                When Nvl(操作方式_In, 0) = 1 Then
                 0
                When Nvl(操作方式_In, 0) = 2 Then
                 1
                When Nvl(操作方式_In, 0) = 3 Then
                 3
                Else
                 0
              End;
      Zl_病人挂号汇总_Update(v_医生姓名, n_医生id, n_项目id, n_科室id, 发生时间_In, n_预约, 号码_In, 0, 记录id_In);
    End If;
  
    If 锁定类型_In <> 1 Then
      --消息推送,锁号时不发送信息
      Begin
        Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
          Using 1, n_挂号id;
      Exception
        When Others Then
          Null;
      End;
      b_Message.Zlhis_Regist_001(n_挂号id, 单据号_In);
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
    When Err_Special Then
      Raise_Application_Error(-20105, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

Begin
  n_出诊记录id := 出诊记录id_In;
  v_Para       := zl_GetSysParameter(256);
  n_挂号模式   := Substr(v_Para, 1, 1);
  Begin
    d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_启用时间 := Null;
  End;

  d_发生时间 := 发生时间_In;
  If d_发生时间 Is Null Then
    d_发生时间 := Sysdate;
  End If;

  If 付款方式_In Is Null Then
    Select Max(名称) Into v_付款方式 From 医疗付款方式 Where 缺省标志 = 1;
  Else
    Select Max(名称) Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In;
    If v_付款方式 Is Null Then
      v_付款方式 := 付款方式_In;
    End If;
  End If;

  If Sysdate - 10 > Nvl(d_启用时间, Sysdate - 30) Then
    If n_挂号模式 = 1 And Nvl(发生时间_In, Sysdate) > Nvl(d_启用时间, Sysdate - 30) And n_出诊记录id Is Null Then
      v_Err_Msg := '系统当前处于出诊表排班模式，传入的参数无法确定挂号安排，请重试！';
      Raise Err_Item;
    End If;
  Else
    If n_挂号模式 = 1 And Nvl(发生时间_In, Sysdate) > Nvl(d_启用时间, Sysdate - 30) And n_出诊记录id Is Null Then
      Begin
        Select a.Id
        Into n_出诊记录id
        From 临床出诊记录 A, 临床出诊号源 B
        Where a.号源id = b.Id And b.号码 = 号码_In And Nvl(发生时间_In, Sysdate) Between a.开始时间 And a.终止时间;
      Exception
        When Others Then
          v_Err_Msg := '系统当前处于出诊表排班模式，传入的参数无法确定挂号安排，请重试！';
          Raise Err_Item;
      End;
    End If;
  End If;

  If n_出诊记录id Is Not Null Then
    --出诊表排班模式
    Zl_三方机构挂号_出诊_Insert(n_出诊记录id, 操作方式_In, 病人id_In, 号码_In, 号序_In, 单据号_In, 票据号_In, 结算方式_In, 摘要_In, 发生时间_In, 登记时间_In,
                        合作单位_In, 挂号金额合计_In, 领用id_In, 收费票据_In, 交易流水号_In, 交易说明_In, 预约方式_In, 预交id_In, 卡类别id_In, 加入序号状态_In,
                        是否自助设备_In, 结帐id_In, 锁定类型_In, 保险结算_In, 冲预交_In, 支付卡号_In, 费别_In, 冲预交病人ids_In, 机器名_In, 更新年龄_In,
                        购买病历_In, 记帐费用_In, 付款方式_In);
  Else
    v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
    v_Temp          := zl_GetSysParameter(256);
    If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
      Null;
    Else
      Begin
        d_启用时间 := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
      Exception
        When Others Then
          Null;
      End;
      If 发生时间_In > d_启用时间 Then
        v_Err_Msg := '当前挂号的发生时间' || To_Char(发生时间_In, 'yyyy-mm-dd hh24:mi:ss') || '已经启用了出诊表排班模式,不能再使用计划排班模式挂号!';
        Raise Err_Item;
      End If;
    End If;
    If 费别_In Is Null Then
      Select Zl_Custom_Getpatifeetype(1, 病人id_In) Into v_费别 From Dual;
    Else
      v_费别 := 费别_In;
    End If;
    If v_费别 Is Null Then
      n_屏蔽费别 := 1;
      Select 名称 Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
    End If;
    Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
  
    If 更新年龄_In = 1 Then
      Select Zl_Age_Calc(病人id_In) Into v_年龄 From Dual;
      If v_年龄 Is Not Null Then
        Update 病人信息 Set 年龄 = v_年龄 Where 病人id = 病人id_In;
      End If;
    End If;
    --获取当前机器名称
    If 机器名_In Is Not Null Then
      v_机器名 := 机器名_In;
    Else
      Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
    End If;
    n_实收金额合计 := 0;
    Select Count(*) + 1
    Into n_挂号序号
    From 病人挂号记录
    Where 号别 = 号码_In And 登记时间 Between Trunc(发生时间_In) And Trunc(发生时间_In + 1) - 1 / 24 / 60 / 60;
    --Begin
    v_Temp           := Nvl(zl_GetSysParameter('病人同科限挂N个号', 1111), '0|0') || '|';
    n_同科限号数     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    n_同科限约数     := To_Number(Nvl(zl_GetSysParameter('病人同科限约N个号', 1111), '0'));
    n_病人预约科室数 := To_Number(Nvl(zl_GetSysParameter('病人预约科室数', 1111), '0'));
    n_病人挂号科室数 := To_Number(Nvl(zl_GetSysParameter('病人挂号科室限制', 1111), '0'));
    n_专家号挂号限制 := To_Number(Nvl(zl_GetSysParameter('专家号挂号限制'), '0'));
    n_专家号预约限制 := To_Number(Nvl(zl_GetSysParameter('专家号预约限制'), '0'));
    n_同源限号数     := To_Number(Nvl(zl_GetSysParameter('病人同一号源限挂N个号', 1111), '0'));
    --部门ID,部门名称;人员ID,人员编号,人员姓名
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Err_Msg := '当前操作人员未设置对应的人员关系,不能继续。';
      Raise Err_Item;
    End If;
  
    If 登记时间_In Is Null Then
      d_登记时间 := Sysdate;
    Else
      d_登记时间 := 登记时间_In;
    End If;
    If Trunc(Sysdate) > Trunc(发生时间_In) Then
      v_Err_Msg := '不能挂以前的号(' || To_Char(发生时间_In, 'yyyy-mm-dd') || ')。';
      Raise Err_Item;
    End If;
    n_开单部门id := To_Number(Zl_操作员(0, v_Temp));
    v_操作员编号 := Zl_操作员(1, v_Temp);
    v_操作员姓名 := Zl_操作员(2, v_Temp);
    n_组id       := Zl_Get组id(v_操作员姓名);
  
    --支付宝并发提交检查
    Select Nvl(Max(1), 0)
    Into n_Exists
    From 病人挂号记录
    Where 病人id = 病人id_In And 号别 = 号码_In And 号序 = 号序_In And 操作员姓名 = v_操作员姓名 And Nvl(n_出诊记录id, 0) = Nvl(出诊记录id, 0) And
          登记时间 > Sysdate - 0.01 And 记录状态 = 1 And 发生时间 = 发生时间_In;
    If n_Exists = 1 Then
      v_Err_Msg := '病人已经挂号,不能重复挂相同的号！';
      Raise Err_Special;
    End If;
  
    If 操作方式_In <> 1 Then
      --预约检查是否添加合作单位控制
      --如果设置了合作单位控制 则
      Begin
        Select ID
        Into n_计划id
        From 挂号安排计划
        Where 号码 = 号码_In And 发生时间_In Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
              Nvl(失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Rownum < 2
        Order By 生效时间 Desc;
      Exception
        When Others Then
          Select ID Into n_Tmp安排id From 挂号安排 Where 号码 = 号码_In;
      End;
      If Nvl(n_计划id, 0) <> 0 Then
        Select Count(0)
        Into n_合作单位限制
        From 合作单位计划控制
        Where 合作单位 = 合作单位_In And 计划id = n_计划id And Rownum < 2;
      Else
        Select Count(0)
        Into n_合作单位限制
        From 合作单位安排控制
        Where 合作单位 = 合作单位_In And 安排id = n_Tmp安排id And Rownum < 2;
      End If;
    End If;
  
    If 操作方式_In <> 2 Then
      v_诊室 := Zl_诊室(号码_In);
    End If;
    If 操作方式_In <> 2 And 结算方式_In Is Not Null Then
      --检查结算方式是否完备
      Select Count(*) Into n_Count From 结算方式 Where 名称 = Nvl(结算方式_In, 'Lxh') And 性质 In (2, 7, 8);
      If Nvl(卡类别id_In, 0) <> 0 And n_Count = 0 Then
        Select Count(1)
        Into n_Count
        From 医疗卡类别
        Where ID = Nvl(卡类别id_In, 0) And 结算方式 = Nvl(结算方式_In, 'lxh');
      End If;
      If n_Count = 0 Then
        v_Err_Msg := '结算方式(' || 结算方式_In || ')未设置,请在结算方式管理中设置。';
        Raise Err_Item;
      End If;
    End If;
  
    --因为现在有按日编单据号规则,日挂号量不能超过10000次,所以要检查唯一约束。
    Select Count(*) Into n_Count From 门诊费用记录 Where 记录性质 = 4 And 记录状态 In (1, 3) And NO = 单据号_In;
    If n_Count <> 0 Then
      v_Err_Msg := '挂号单据号重复,不能保存！' || Chr(13) || '如果使用了按日顺序编号,当日挂号量不能超过10000人次。';
      Raise Err_Item;
    End If;
  
    Open c_Pati(病人id_In);
    n_Count := 0;
    Begin
      Fetch c_Pati
        Into r_Pati;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '病人未找到，不能继续。';
      Raise Err_Item;
    End If;
  
    Open c_安排(号码_In, 发生时间_In);
    Begin
      Fetch c_安排
        Into r_安排;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '该号别没有在' || To_Char(发生时间_In, 'yyyy-mm-dd hh24:mi:ss') || '中进行安排。';
      Raise Err_Item;
    End If;
  
    Select Min(站点) Into v_站点 From 部门表 Where ID = r_安排.科室id;
    v_Pricegrade := Zl_Get_Pricegrade(v_站点, 病人id_In, Null, v_付款方式);
    v_普通等级   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);
  
    Select Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六',
                   '周日')
    Into v_星期
    From Dual;
    Begin
      If r_安排.计划id Is Null Then
        Select Max(1) Into n_启用分时段 From 挂号安排时段 Where 安排id = r_安排.Id And 星期 = v_星期 And Rownum < 2;
        Select Decode(To_Char(发生时间_In, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
        Into v_时间段
        From 挂号安排
        Where ID = r_安排.Id;
      Else
        Select Max(1)
        Into n_启用分时段
        From 挂号计划时段
        Where 计划id = r_安排.计划id And 星期 = v_星期 And Rownum < 2;
        Select Decode(To_Char(发生时间_In, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
        Into v_时间段
        From 挂号安排计划
        Where ID = r_安排.计划id;
      End If;
    Exception
      When Others Then
        n_启用分时段 := 0;
    End;
  
    If v_时间段 Is Not Null And d_启用时间 Is Not Null Then
      --检查是否跨模式挂号安排
      Select To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
      Into d_检查开始时间, d_检查结束时间
      From 时间段
      Where 时间段 = v_时间段 And 站点 Is Null And 号类 Is Null;
      If d_检查开始时间 > d_检查结束时间 Then
        d_检查结束时间 := d_检查结束时间 + 1;
      End If;
      If d_检查结束时间 > d_启用时间 Then
        --获取出诊记录id
        Begin
          Select a.Id
          Into n_出诊记录id
          From 临床出诊记录 A, 临床出诊号源 B
          Where a.号源id = b.Id And b.号码 = 号码_In And 上班时段 = v_时间段 And 发生时间_In Between 开始时间 And 终止时间;
        Exception
          When Others Then
            n_出诊记录id := Null;
        End;
      End If;
    End If;
  
    --对参数控制进行检查
    --仅在预约不扣款时进行检查
    If 操作方式_In = 2 Then
      If Nvl(n_同科限约数, 0) <> 0 Or Nvl(n_病人预约科室数, 0) <> 0 Then
        n_已约科室 := 0;
        For c_Chkitem In (Select Distinct 执行部门id
                          From 病人挂号记录
                          Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
                                Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> r_安排.科室id) Loop
          n_已约科室 := n_已约科室 + 1;
        End Loop;
        If n_已约科室 >= Nvl(n_病人预约科室数, 0) And Nvl(n_病人预约科室数, 0) > 0 Then
          v_Err_Msg := '同一病人最多同时能预约[' || Nvl(n_病人预约科室数, 0) || ']个科室,不能再预约！';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id = r_安排.科室id;
        If n_Count >= Nvl(n_同科限约数, 0) And Nvl(n_同科限约数, 0) > 0 Then
          v_Err_Msg := '该病人已经在该科室预约了' || n_Count || '次,不能再预约！';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_专家号预约限制, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 号别 = r_安排.号码;
        If n_Count >= Nvl(n_专家号预约限制, 0) And Nvl(n_专家号预约限制, 0) > 0 Then
          v_Err_Msg := '该病人已经超过本号预约限制,不能再预约！';
          Raise Err_Item;
        End If;
      End If;
    Else
      If Nvl(n_同科限号数, 0) <> 0 Or Nvl(n_病人挂号科室数, 0) <> 0 Then
        n_已约科室 := 0;
        For c_Chkitem In (Select Distinct 执行部门id
                          From 病人挂号记录
                          Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And
                                Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> r_安排.科室id) Loop
          n_已约科室 := n_已约科室 + 1;
        End Loop;
        If n_已约科室 >= Nvl(n_病人挂号科室数, 0) And Nvl(n_病人挂号科室数, 0) > 0 Then
          v_Err_Msg := '同一病人最多同时能挂号[' || Nvl(n_病人挂号科室数, 0) || ']个科室,不能再挂号！';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id = r_安排.科室id;
        If n_Count >= Nvl(n_同科限号数, 0) And Nvl(n_同科限号数, 0) > 0 Then
          v_Err_Msg := '该病人已经在该科室挂号了' || n_Count || '次,不能再挂号！';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_专家号挂号限制, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 预约时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 号别 = r_安排.号码;
        If n_Count >= Nvl(n_专家号挂号限制, 0) And Nvl(n_专家号挂号限制, 0) > 0 Then
          v_Err_Msg := '该病人已经超过本号挂号限制,不能再挂号！';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    If Nvl(n_同源限号数, 0) <> 0 Then
      If 出诊记录id_In Is Null Then
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 In (1, 2) And 发生时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 号别 = 号码_In;
      Else
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 In (1, 2) And 出诊记录id = 出诊记录id_In;
      End If;
      If n_Count >= Nvl(n_同源限号数, 0) And Nvl(n_同源限号数, 0) > 0 Then
        v_Err_Msg := '同一病人最多能同时挂(预约)[' || Nvl(n_同源限号数, 0) || ']个相同号别的号,不能再挂号(预约)！';
        Raise Err_Item;
      End If;
    End If;
  
    d_Date         := Null;
    d_时段开始时间 := Null;
  
    If Nvl(r_安排.限号数, 0) >= 0 Or r_安排.限号数 Is Null Then
    
      Select Nvl(Sum(Nvl(b.已挂数, 0)), 0), Nvl(Sum(Nvl(b.其中已接收, 0)), 0), Nvl(Sum(Nvl(b.已约数, 0)), 0)
      Into n_已挂数, n_其中已接收, n_已约数
      From 挂号安排 A, 病人挂号汇总 B
      Where a.科室id = b.科室id And a.项目id = b.项目id And a.号码 = 号码_In And b.日期 Between Trunc(发生时间_In) And
            Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And (a.号码 = b.号码 Or b.号码 Is Null) And Nvl(a.医生id, 0) = Nvl(b.医生id, 0) And
            Nvl(a.医生姓名, '医生') = Nvl(b.医生姓名, '医生');
    
      If n_启用分时段 = 1 Then
        If Nvl(r_安排.序号控制, 0) = 1 Then
          If Nvl(是否自助设备_In, 0) = 0 Then
            If r_安排.计划id Is Null Then
              Select Count(*), Max(开始时间)
              Into n_Count, d_时段开始时间
              From 挂号安排时段
              Where 安排id = r_安排.Id And 星期 = v_星期 And 序号 = Nvl(号序_In, 0);
            Else
              Select Count(*), Max(开始时间)
              Into n_Count, d_时段开始时间
              From 挂号计划时段
              Where 计划id = r_安排.计划id And 星期 = v_星期 And 序号 = Nvl(号序_In, 0);
            End If;
            v_Temp := '挂号';
            If 操作方式_In > 1 Then
              v_Temp := '预约挂号';
            End If;
          
            If n_Count = 0 Then
              v_Err_Msg := '号别为' || 号码_In || '的挂号安排中不存在序号为' || Nvl(号序_In, 0) || '的安排,不能再' || v_Temp || '！';
              Raise Err_Item;
            End If;
          End If;
          --过点的,不能选择挂号
          If Trunc(Sysdate) = Trunc(发生时间_In) Then
            --挂当天的号
            v_Temp := To_Char(Sysdate, 'yyyy-mm-dd') || ' ';
            If r_安排.计划id Is Null Then
              For v_时段 In (Select To_Date(v_Temp || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                  To_Date(To_Char(Sysdate + Decode(Sign(开始时间 - 结束时间), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                           To_Char(结束时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 结束时间, 限制数量, 是否预约
                           From 挂号安排时段
                           Where 安排id = r_安排.Id And 星期 = v_星期 And 序号 = Nvl(号序_In, 0)) Loop
                If Sysdate > v_时段.结束时间 Then
                  v_Err_Msg := '号别为' || 号码_In || '及号序为' || Nvl(号序_In, 0) || '的安排,已经超过时点,不能再' || v_Temp || '！';
                  Raise Err_Item;
                End If;
              End Loop;
            Else
              For v_时段 In (Select To_Date(v_Temp || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                  To_Date(To_Char(Sysdate + Decode(Sign(开始时间 - 结束时间), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                           To_Char(结束时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 结束时间, 限制数量, 是否预约
                           From 挂号计划时段
                           Where 计划id = r_安排.计划id And 星期 = v_星期 And 序号 = Nvl(号序_In, 0)) Loop
                If Sysdate > v_时段.结束时间 Then
                  v_Err_Msg := '号别为' || 号码_In || '及号序为' || Nvl(号序_In, 0) || '的安排,已经超过时点,不能再' || v_Temp || '！';
                  Raise Err_Item;
                End If;
              End Loop;
            End If;
          End If;
        Elsif 操作方式_In > 1 Then
          --未启用序号的,需要检查预约的情况
          n_Count := 0;
          If r_安排.计划id Is Null Then
            For v_时段 In (Select 序号, 开始时间, 结束时间, 限制数量, 是否预约
                         From 挂号安排时段
                         Where 安排id = r_安排.Id And 星期 = v_星期 And
                               (('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                               Decode(Sign(开始时间 - 结束时间 - 1 / 24 / 60 / 60), 1,
                                        '3000-01-09 ' || To_Char(开始时间, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS')) And
                               '3000-01-10 ' || To_Char(结束时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                               ('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                               '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS') And
                               Decode(Sign(开始时间 - 结束时间 - 1 / 24 / 60 / 60), 1,
                                        '3000-01-11 ' || To_Char(结束时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(结束时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
              n_预约时段序号 := v_时段.序号;
              d_时段开始时间 := v_时段.开始时间;
            
              Select Count(*), Max(序号)
              Into n_Count, n_预约总数
              From 挂号序号状态
              Where 号码 = 号码_In And 日期 = 发生时间_In And 状态 Not In (4, 5);
            
              If Nvl(n_Count, 0) > Nvl(v_时段.限制数量, 0) And 锁定类型_In <> 2 Then
                v_Err_Msg := '号别为' || 号码_In || '的挂号安排中在' || To_Char(v_时段.开始时间, 'hh24:mi:ss') || '至' ||
                             To_Char(v_时段.结束时间, 'hh24:mi:ss') || '最多只能预约' || Nvl(v_时段.限制数量, 0) || '人,不能再进行预约挂号！';
                Raise Err_Item;
              End If;
              n_Count := 1;
            End Loop;
          Else
            For v_时段 In (Select 序号, 开始时间, 结束时间, 限制数量, 是否预约
                         From 挂号计划时段
                         Where 计划id = r_安排.计划id And 星期 = v_星期 And
                               (('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                               Decode(Sign(开始时间 - 结束时间 - 1 / 24 / 60 / 60), 1,
                                        '3000-01-09 ' || To_Char(开始时间, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS')) And
                               '3000-01-10 ' || To_Char(结束时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                               ('3000-01-10 ' || To_Char(发生时间_In, 'HH24:MI:SS') Between
                               '3000-01-10 ' || To_Char(开始时间, 'HH24:MI:SS') And
                               Decode(Sign(开始时间 - 结束时间 - 1 / 24 / 60 / 60), 1,
                                        '3000-01-11 ' || To_Char(结束时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(结束时间 - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
              n_预约时段序号 := v_时段.序号;
              d_时段开始时间 := v_时段.开始时间;
            
              Select Count(*), Max(序号)
              Into n_Count, n_预约总数
              From 挂号序号状态
              Where 号码 = 号码_In And 日期 = 发生时间_In And 状态 Not In (4, 5);
            
              If Nvl(n_Count, 0) > Nvl(v_时段.限制数量, 0) And 锁定类型_In <> 2 Then
                v_Err_Msg := '号别为' || 号码_In || '的挂号安排中在' || To_Char(v_时段.开始时间, 'hh24:mi:ss') || '至' ||
                             To_Char(v_时段.结束时间, 'hh24:mi:ss') || '最多只能预约' || Nvl(v_时段.限制数量, 0) || '人,不能再进行预约挂号！';
                Raise Err_Item;
              End If;
              n_Count := 1;
            End Loop;
          End If;
        
          If n_Count = 0 Then
            v_Err_Msg := '号别为' || 号码_In || '的挂号安排中没有相关的安排计划(' || To_Char(发生时间_In, 'YYYY-mm-dd HH24:MI:SS') ||
                         '),不能进行预约挂号！';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End If;
  
    If 操作方式_In = 1 And 锁定类型_In <> 2 Then
      --挂号规则:
      --  已挂数不能大于限号数
      If n_已挂数 >= Nvl(r_安排.限号数, 0) And r_安排.限号数 Is Not Null Then
        v_Err_Msg := '该号别今天已达到限号数 ' || Nvl(r_安排.限号数, 0) || '不能再挂号！';
        Raise Err_Item;
      End If;
    End If;
  
    If 操作方式_In > 1 Then
      --预约的相关检查
      --规则:
      --   1.已限约不能超过限约数
      --   2.检查是否启用时段的
      If n_已约数 >= Nvl(r_安排.限约数, 0) And Nvl(r_安排.限约数, 0) <> 0 And r_安排.限约数 Is Not Null And 锁定类型_In <> 2 Then
        v_Err_Msg := '该号别已达到限约数 ' || Nvl(r_安排.限约数, 0) || '不能再预约挂号！';
        Raise Err_Item;
      End If;
    End If;
    If n_合作单位限制 > 0 And 操作方式_In <> 1 And 合作单位_In Is Not Null Then
    
      If Nvl(r_安排.序号控制, 0) = 1 And Nvl(号序_In, 0) = 0 Then
        v_Err_Msg := '当前安排使用了序号控制,请确认所需要预约的序号,不能继续。';
        Raise Err_Item;
      End If; --Nvl(r_安排.序号控制, 0) =0
    
      n_序号 := Case
                When Nvl(r_安排.序号控制, 0) = 1 Or n_启用分时段 = 1 And 操作方式_In > 1 Then
                 Nvl(号序_In, 0)
                Else
                 0
              End;
    
      --合作单位限数量模式
      Begin
        If Nvl(n_计划id, 0) <> 0 Then
          Select 0
          Into n_序号
          From 合作单位计划控制
          Where 合作单位 = 合作单位_In And 计划id = n_计划id And
                限制项目 = Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                              '7', '周六', Null) And 数量 <> 0 And 序号 = 0 And Rownum < 2;
        Else
          Select 0
          Into n_序号
          From 合作单位安排控制
          Where 合作单位 = 合作单位_In And 安排id = n_Tmp安排id And
                限制项目 = Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                              '7', '周六', Null) And 数量 <> 0 And 序号 = 0 And Rownum < 2;
        End If;
        n_合作单位限数量模式 := 1;
      Exception
        When Others Then
          n_合作单位限数量模式 := 0;
      End;
      --开放序号检查
      For c_合作单位 In (Select c.序号, 数量
                     From 挂号安排 A, 合作单位安排控制 C
                     Where a.号码 = 号码_In And Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                   '周四', '6', '周五', '7', '周六', Null) = c.限制项目(+) And a.Id = c.安排id And
                           c.合作单位 = 合作单位_In And c.序号 = n_序号 And Not Exists
                      (Select 1
                            From 挂号安排计划 D
                            Where d.安排id = a.Id And d.审核时间 Is Not Null And
                                  发生时间_In Between Nvl(d.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                  Nvl(d.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')))
                     Union All
                     Select c.序号, 数量
                     From 挂号安排计划 A, 挂号安排 D, 合作单位计划控制 C,
                          (Select Max(a.生效时间) As 生效, 安排id
                            From 挂号安排计划 A, 挂号安排 B
                            Where a.安排id = b.Id And a.审核时间 Is Not Null And
                                  发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                  Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And b.号码 = 号码_In
                            Group By 安排id) E
                     Where a.安排id = d.Id And a.审核时间 Is Not Null And d.号码 = 号码_In And a.安排id = e.安排id And
                           Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) =
                           Nvl(e.生效, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                           Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null) = c.限制项目(+) And a.Id = c.计划id And c.合作单位 = 合作单位_In And c.序号 = n_序号 And
                           发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                           Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd'))) Loop
      
        If Nvl(r_安排.序号控制, 0) = 1 And c_合作单位.序号 = n_序号 And n_合作单位限数量模式 = 0 Then
          n_是否开放 := 1;
          Exit;
        Elsif (Nvl(r_安排.序号控制, 0) = 0 And c_合作单位.序号 = n_序号) Or n_合作单位限数量模式 = 1 Then
          Begin
            Select Nvl(已约数, 0)
            Into n_预约数量
            From 合作单位挂号汇总
            Where 合作单位 = 合作单位_In And 日期 = Trunc(发生时间_In) And 号码 = 号码_In;
          Exception
            When Others Then
              n_预约数量 := 0;
          End;
          If c_合作单位.数量 <= n_预约数量 And Nvl(c_合作单位.数量, 0) > 0 And 锁定类型_In <> 2 Then
            v_Err_Msg := '该号别已达到限约数 ' || Nvl(c_合作单位.数量, 0) || '不能再预约挂号！';
            Raise Err_Item;
          End If;
          n_是否开放 := 1;
          Exit;
        End If;
      
      End Loop;
    
      If Nvl(n_是否开放, 0) = 0 Then
        v_Err_Msg := '当前序号(' || Nvl(号序_In, 0) || '未开放,不能继续。';
        Raise Err_Item;
      End If;
    End If;
  
    --检查限号数和限约数
    n_行号         := 1;
    n_原项目id     := 0;
    n_原收入项目id := 0;
    n_实收金额合计 := 0;
    If 锁定类型_In <> 1 Then
      If 操作方式_In <> 2 Then
        If Nvl(结帐id_In, 0) = 0 Then
          --这里应该程序传入
          Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
        Else
          n_结帐id := 结帐id_In;
        End If;
      Else
        n_结帐id := Null;
      End If;
    End If;
    n_项目id := r_安排.项目id;
    If Nvl(n_计划id, 0) <> 0 Then
      v_传入 := '1|' || n_计划id;
    Else
      If Nvl(r_安排.Id, 0) <> 0 Then
        v_传入 := '0|' || r_安排.Id;
      End If;
    End If;
    If v_传入 Is Null Then
      v_传入 := '3|' || 号码_In;
    End If;
  
    n_更新项目id := Zl_Custom_Getregeventitem(r_Pati.病人id, r_Pati.姓名, r_Pati.身份证号, r_Pati.出生日期, r_Pati.性别, r_Pati.年龄, v_传入);
    If Nvl(n_更新项目id, 0) <> 0 Then
      n_项目id := n_更新项目id;
    End If;
  
    If Nvl(购买病历_In, 0) = 1 Then
      Begin
        Select 收费细目id Into n_病历费id From 收费特定项目 Where 特定项目 = '病历费';
        v_收费项目ids := n_项目id || ',' || n_病历费id;
      Exception
        When Others Then
          v_Err_Msg := '不能确定病历费,挂号失败!';
          Raise Err_Item;
      End;
    Else
      v_收费项目ids := n_项目id;
    End If;
  
    For c_Item In (Select 1 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, Null As 从属父号,
                          Nvl(a.项目特性, 0) As 急诊
                   From 收费项目目录 A, 收费价目 B, 收入项目 C
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = r_安排.项目id And d_发生时间 Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.价格等级 = v_普通等级 Or
                         (b.价格等级 Is Null And Not Exists
                          (Select 1
                            From 收费价目
                            Where b.收费细目id = 收费细目id And 价格等级 = v_普通等级 And d_发生时间 Between 执行日期 And
                                  Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Union All
                   Select 2 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, 1 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, Null As 从属父号, 0 As 急诊
                   From 收费项目目录 A, 收费价目 B, 收入项目 C
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = n_病历费id And d_发生时间 Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.价格等级 = v_普通等级 Or
                         (b.价格等级 Is Null And Not Exists
                          (Select 1
                            From 收费价目
                            Where b.收费细目id = 收费细目id And 价格等级 = v_普通等级 And d_发生时间 Between 执行日期 And
                                  Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Union All
                   Select 3 As 性质, a.类别, a.Id As 项目id, a.名称 As 项目名称, a.编码 As 项目编码, a.计算单位, a.屏蔽费别, d.从项数次 As 数次,
                          c.Id As 收入项目id, c.名称 As 收入项目, c.编码 As 收入编码, c.收据费目, b.现价 As 单价, 1 As 从属父号, 0 As 急诊
                   From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费从属项目 D
                   Where b.收费细目id = a.Id And b.收入项目id = c.Id And a.Id = d.从项id And
                         d.主项id In (Select Column_Value From Table(f_Str2list(v_收费项目ids))) And d_发生时间 Between b.执行日期 And
                         Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.价格等级 = v_普通等级 Or
                         (b.价格等级 Is Null And Not Exists
                          (Select 1
                            From 收费价目
                            Where b.收费细目id = 收费细目id And 价格等级 = v_普通等级 And d_发生时间 Between 执行日期 And
                                  Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Order By 性质, 项目编码, 收入编码) Loop
      If c_Item.性质 = 1 Then
        n_急诊 := Nvl(c_Item.急诊, 0);
      End If;
      n_价格父号 := Null;
      If n_原项目id = c_Item.项目id Then
        If n_原收入项目id <> c_Item.收入项目id Then
          n_价格父号 := n_行号;
        End If;
        n_原收入项目id := c_Item.收入项目id;
      End If;
      n_原项目id := c_Item.项目id;
      n_应收金额 := Round(c_Item.数次 * c_Item.单价, 5);
      n_实收金额 := n_应收金额;
      If Nvl(c_Item.屏蔽费别, 0) <> 1 And n_屏蔽费别 = 0 Then
        --打折:
        v_Temp     := Zl_Actualmoney(r_Pati.费别, c_Item.项目id, c_Item.收入项目id, n_应收金额);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ':') + 1);
        n_实收金额 := Zl_To_Number(v_Temp);
      End If;
      n_实收金额合计 := Nvl(n_实收金额合计, 0) + n_实收金额;
    
      --锁定单据不产生费用
      If 锁定类型_In <> 1 Then
        --产生病人挂号费用(可能单独是或包括病历费用)
        Select 病人费用记录_Id.Nextval Into n_费用id From Dual;
        --:操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
        Insert Into 门诊费用记录
          (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id,
           收费类别, 计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人,
           操作员编号, 操作员姓名, 发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
        Values
          (n_费用id, 4, Decode(操作方式_In, 2, 0, 1), n_行号, n_价格父号, c_Item.从属父号, 单据号_In, 票据号_In, 1, n_急诊, Null,
           Decode(操作方式_In, 2, To_Char(号序_In), v_诊室), r_Pati.病人id, r_Pati.门诊号, r_Pati.付款方式, r_Pati.姓名, r_Pati.性别,
           r_Pati.年龄, r_Pati.费别, r_安排.科室id, c_Item.类别, 号码_In, c_Item.项目id, c_Item.收入项目id, c_Item.收据费目, 1, c_Item.数次,
           c_Item.单价, n_应收金额, n_实收金额, Decode(操作方式_In, 2, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, n_实收金额)),
           Decode(Nvl(记帐费用_In, 0), 1, Null, n_结帐id), Decode(Nvl(记帐费用_In, 0), 1, 1, 0), n_开单部门id, v_操作员姓名,
           Decode(操作方式_In, 2, v_操作员姓名, Null), r_安排.科室id, r_安排.医生姓名, v_操作员编号, v_操作员姓名, 发生时间_In, d_登记时间, Null, 0, Null,
           Null, 摘要_In, 预约方式_In, Decode(操作方式_In, 2, Null, n_组id));
      End If;
      n_行号 := n_行号 + 1;
    
    End Loop;
  
    If Round(Nvl(挂号金额合计_In, 0), 5) <> Round(Nvl(n_实收金额合计, 0), 5) Then
      v_Err_Msg := '本次挂号金额不正确,可能是因为医院调整了价格,请重新获取挂号收费项目的价格,不能继续。';
      Raise Err_Item;
    End If;
  
    If n_启用分时段 = 1 Then
      d_Date := To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(d_时段开始时间, 'hh24:mi:ss'),
                        'yyyy-mm-dd hh24:mi:ss');
    Else
      d_Date := Trunc(发生时间_In);
    End If;
  
    --更新挂号序号状态
    If 锁定类型_In <> 2 Then
      n_号序 := 号序_In;
    End If;
    Begin
      Select 1
      Into n_Count
      From 挂号序号状态
      Where Trunc(日期) = Trunc(发生时间_In) And 号码 = 号码_In And 序号 = n_号序 And 状态 <> 5;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count = 1 Then
      If n_启用分时段 = 0 And Nvl(r_安排.序号控制, 0) = 1 Then
        n_号序 := Null;
      End If;
      If n_启用分时段 = 1 And Nvl(r_安排.序号控制, 0) = 1 Then
        v_Err_Msg := '当前序号已被使用，请重新选择一个序号！';
        Raise Err_Item;
      End If;
    End If;
    n_Count := 0;
    If n_启用分时段 = 0 And Nvl(r_安排.序号控制, 0) = 1 And n_号序 Is Null And 锁定类型_In <> 2 Then
      If 退号重用_In = 1 Then
        Select Nvl(Max(序号), 0) + 1
        Into n_号序
        From 挂号序号状态
        Where 日期 = Trunc(发生时间_In) And 号码 = r_安排.号码 And 状态 Not In (4, 5);
      Else
        Select Nvl(Max(序号), 0) + 1
        Into n_号序
        From 挂号序号状态
        Where 日期 = Trunc(发生时间_In) And 号码 = r_安排.号码 And 状态 <> 5;
      End If;
    End If;
    If n_启用分时段 = 1 And 锁定类型_In <> 2 Then
    
      If 操作方式_In > 1 And Nvl(r_安排.序号控制, 0) = 0 Then
        --规则:预约时段序号||预约数
        If Nvl(n_预约总数, 0) = 0 Then
          v_Temp := Nvl(r_安排.限约数, 0);
          v_Temp := LTrim(RTrim(v_Temp));
          v_Temp := LPad(Nvl(n_预约总数, 0) + 1, Length(v_Temp), '0');
          v_Temp := n_预约时段序号 || v_Temp;
          n_号序 := To_Number(v_Temp);
        Else
          n_号序 := n_预约总数 + 1;
        End If;
      End If;
    End If;
  
    If Nvl(r_安排.序号控制, 0) = 1 Or (操作方式_In > 1 And n_启用分时段 = 1) Or 加入序号状态_In = 1 Then
      --锁定序号的处理
      Begin
        Select 操作员姓名, 机器名
        Into v_序号操作员, v_序号机器名
        From 挂号序号状态
        Where 状态 = 5 And 号码 = 号码_In And Trunc(日期) = Trunc(d_Date) And 序号 = n_号序;
        n_序号锁定 := 1;
      Exception
        When Others Then
          v_序号操作员 := Null;
          v_序号机器名 := Null;
          n_序号锁定   := 0;
      End;
      If n_序号锁定 = 0 Then
        Update 挂号序号状态
        Set 状态 = Decode(操作方式_In, 2, 2, 1), 预约 = Decode(操作方式_In, 1, 0, 1), 登记时间 = Sysdate
        Where 号码 = 号码_In And 日期 = d_Date And 序号 = n_号序 And 操作员姓名 = v_操作员姓名;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 预约, 登记时间)
            Values
              (号码_In, d_Date, n_号序, Decode(操作方式_In, 2, 2, 1), v_操作员姓名, Decode(操作方式_In, 1, 0, 1), Sysdate);
          
            If n_合作单位限制 > 0 And 操作方式_In > 1 And Nvl(n_是否开放, 0) = 1 Then
              Update 合作单位挂号汇总
              Set 已约数 = 已约数 + Decode(操作方式_In, 2, 1, 0), 已接数 = 已接数 + Decode(操作方式_In, 3, 1, 0)
              Where 号码 = 号码_In And 日期 = d_Date And 序号 = n_号序 And 合作单位 = 合作单位_In;
              If Sql%NotFound Then
                Insert Into 合作单位挂号汇总
                  (号码, 日期, 序号, 合作单位, 已约数, 已接数)
                Values
                  (号码_In, d_Date, n_号序, 合作单位_In, Decode(操作方式_In, 1, 0, 1), Decode(操作方式_In, 3, 1, 0));
              End If;
            End If;
          Exception
            When Others Then
              v_Err_Msg := '序号' || n_号序 || '已被使用,请重新选择一个序号.';
              Raise Err_Item;
          End;
        End If;
      Else
        If v_操作员姓名 <> v_序号操作员 Or v_机器名 <> v_序号机器名 Then
          v_Err_Msg := '序号' || n_号序 || '已被自助机' || v_机器名 || '锁定,请重新选择一个序号.';
          Raise Err_Item;
        Else
          Update 挂号序号状态
          Set 状态 = Decode(操作方式_In, 2, 2, 1), 预约 = Decode(操作方式_In, 1, 0, 1), 登记时间 = Sysdate
          Where 号码 = 号码_In And Trunc(日期) = Trunc(d_Date) And 序号 = n_号序 And 状态 = 5 And 操作员姓名 = v_操作员姓名 And 机器名 = v_机器名;
        End If;
      End If;
    End If;
  
    If n_出诊记录id Is Not Null Then
      Update 临床出诊序号控制
      Set 挂号状态 = Decode(操作方式_In, 2, 2, 1), 操作员姓名 = v_操作员姓名
      Where 记录id = n_出诊记录id And 序号 = n_序号;
      If 操作方式_In = 2 Then
        Update 临床出诊记录 Set 已约数 = 已约数 + 1 Where ID = n_出诊记录id;
      Else
        If 操作方式_In <> 1 Then
          Update 临床出诊记录
          Set 已约数 = 已约数 + 1, 已挂数 = 已挂数 + 1, 其中已接收 = 其中已接收 + 1
          Where ID = n_出诊记录id;
        Else
          Update 临床出诊记录 Set 已挂数 = 已挂数 + 1 Where ID = n_出诊记录id;
        End If;
      End If;
    End If;
  
    --锁定单据不产生任何 费用
    If 操作方式_In <> 2 And 锁定类型_In <> 1 And Nvl(记帐费用_In, 0) = 0 Then
      --挂号,预约挂号已经扣款部分
      n_预交id := 预交id_In;
      If Nvl(n_预交id, 0) = 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      End If;
      n_结算合计 := 0;
      If 保险结算_In Is Not Null Then
        --各个保险结算
        v_结算内容 := 保险结算_In || '||';
        n_结算合计 := 0;
        While v_结算内容 Is Not Null Loop
          v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
          v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
          n_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
          If Nvl(n_结算金额, 0) <> 0 Then
            Insert Into 病人预交记录
              (ID, 记录性质, NO, 记录状态, 病人id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 结算性质)
            Values
              (n_预交id, 4, 单据号_In, 1, Decode(病人id_In, 0, Null, 病人id_In), '保险结算', v_结算方式, d_登记时间, v_操作员编号, v_操作员姓名,
               n_结算金额, n_结帐id, n_组id, n_结帐id, 4);
            Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          End If;
          n_结算合计 := Nvl(n_结算合计, 0) + Nvl(n_结算金额, 0);
          v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
        End Loop;
      End If;
    
      If Nvl(冲预交_In, 0) <> 0 Then
        --处理总预交
        n_结算合计 := n_结算合计 + Nvl(冲预交_In, 0);
        n_预交金额 := 冲预交_In;
        For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
          n_结算金额 := Case
                      When r_Deposit.金额 - n_预交金额 < 0 Then
                       r_Deposit.金额
                      Else
                       n_预交金额
                    End;
          If r_Deposit.结帐id = 0 Then
            --第一次冲预交(填上结帐ID,金额为0)
            Update 病人预交记录 Set 冲预交 = 0, 结帐id = n_结帐id, 结算性质 = 4 Where ID = r_Deposit.原预交id;
          End If;
          --冲上次剩余额
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名,
             操作员编号, 冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行,
                   单位帐号, d_登记时间, v_操作员姓名, v_操作员编号, n_结算金额, n_结帐id, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, n_结帐id, 4
            From 病人预交记录
            Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
        
          --更新病人预交余额
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) - n_结算金额
          Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2)
          Returning 预交余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 病人余额 (病人id, 预交余额, 性质, 类型) Values (r_Deposit.病人id, -1 * n_结算金额, 1, 1);
            n_返回值 := -1 * n_结算金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 病人余额
            Where 病人id = r_Deposit.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;
        
          --检查是否已经处理完
          If r_Deposit.金额 <= n_结算金额 Then
            n_预交金额 := n_预交金额 - r_Deposit.金额;
          Else
            n_预交金额 := 0;
          End If;
          If n_预交金额 = 0 Then
            Exit;
          End If;
        End Loop;
        If n_预交金额 > 0 Then
          v_Err_Msg := '预交余额不够支付本次支付金额,不能继续操作！';
          Raise Err_Item;
        End If;
      End If;
      --剩余款项,用指定结算方支付
      n_结算金额 := Nvl(n_实收金额合计, 0) - Nvl(n_结算合计, 0);
      If Nvl(n_结算金额, 0) < 0 Then
        v_Err_Msg := '挂号的相关结算金额超出了当前实结金额,不能继续操作！';
        Raise Err_Item;
      End If;
      If Nvl(n_结算金额, 0) <> 0 Or (Nvl(n_结算金额, 0) = 0 And Nvl(冲预交_In, 0) = 0) Then
        If 结算方式_In Is Null Then
          v_Err_Msg := '未传入指定的结算方式,不能继续操作！';
          Raise Err_Item;
        End If;
      
        If Nvl(预交id_In, 0) <> 0 Then
          --传入的预交ID_In主要是为了解决三方交易,如果医保结算站用了该ID,需要用新的ID进行更新,三方交易用转入的ID
          Update 病人预交记录 Set ID = n_预交id Where ID = Nvl(预交id_In, 0);
          n_预交id := Nvl(预交id_In, 0);
        End If;
      
        Insert Into 病人预交记录
          (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 交易流水号, 交易说明, 结算序号, 合作单位, 卡类别id, 卡号,
           结算性质)
        Values
          (n_预交id, 4, 1, 单据号_In, r_Pati.病人id, 结算方式_In, Nvl(n_结算金额, 0), d_登记时间, v_操作员编号, v_操作员姓名, n_结帐id,
           合作单位_In || '缴款', n_组id, 交易流水号_In, 交易说明_In, n_结帐id, 合作单位_In, 卡类别id_In, 支付卡号_In, 4);
      End If;
    
      --更新人员缴款数据
    
      For v_缴款 In (Select 结算方式, Sum(Nvl(a.冲预交, 0)) As 冲预交
                   From 病人预交记录 A
                   Where a.结帐id = n_结帐id And Mod(a.记录性质, 10) <> 1 And Nvl(病人id, 0) = Nvl(病人id_In, 0)
                   Group By 结算方式) Loop
      
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + Nvl(v_缴款.冲预交, 0)
        Where 收款员 = v_操作员姓名 And 性质 = 1 And 结算方式 = v_缴款.结算方式
        Returning 余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (v_操作员姓名, v_缴款.结算方式, 1, Nvl(v_缴款.冲预交, 0));
          n_返回值 := Nvl(v_缴款.冲预交, 0);
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = v_操作员姓名 And 结算方式 = 结算方式_In And 性质 = 1 And Nvl(余额, 0) = 0;
        End If;
      
      End Loop;
    
    End If;
  
    --处理挂号记录
    If 锁定类型_In = 2 Then
      Begin
        Select ID Into n_挂号id From 病人挂号记录 Where　记录状态 = 0 And NO = 单据号_In And 病人id = 病人id_In;
      Exception
        When Others Then
          Null;
      End;
    Else
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
    End If;
  
    Update 病人挂号记录
    Set 记录性质 = Decode(操作方式_In, 2, 2, 1), 记录状态 = Decode(锁定类型_In, 1, 0, 1), 门诊号 = r_Pati.门诊号, 操作员姓名 = v_操作员姓名,
        操作员编号 = v_操作员编号, 预约 = Decode(操作方式_In, 1, 0, 1),
        接收人 = Decode(锁定类型_In, 1, Null, Decode(操作方式_In, 2, Null, v_操作员姓名)),
        接收时间 = Case 锁定类型_In
                  When 1 Then
                   Null
                  Else
                   Case 操作方式_In
                     When 2 Then
                      Null
                     Else
                      d_登记时间
                   End
                End, 交易流水号 = Nvl(交易流水号_In, 交易流水号), 交易说明 = Nvl(交易说明_In, 交易说明), 合作单位 = Nvl(合作单位_In, 合作单位),
        预约操作员 = Decode(操作方式_In, 1, Nvl(预约操作员, Null), Nvl(预约操作员, v_操作员姓名)),
        预约操作员编号 = Decode(操作方式_In, 1, Nvl(预约操作员编号, Null), Nvl(预约操作员编号, v_操作员编号))
    Where ID = n_挂号id;
    If Sql%NotFound Then
      Begin
        Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = r_Pati.付款方式 And Rownum < 2;
      Exception
        When Others Then
          v_付款方式 := Null;
      End;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 预约时间, 操作员编号,
         操作员姓名, 复诊, 号序, 预约, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式, 预约操作员, 预约操作员编号)
      Values
        (n_挂号id, 单据号_In, Decode(操作方式_In, 2, 2, 1), Decode(锁定类型_In, 1, 0, 1), r_Pati.病人id, r_Pati.门诊号, r_Pati.姓名,
         r_Pati.性别, r_Pati.年龄, 号码_In, n_急诊, v_诊室, Null, r_安排.科室id, r_安排.医生姓名, 0, Null, d_登记时间, 发生时间_In,
         Case When(Nvl(操作方式_In, 0)) = 1 Then Null Else 发生时间_In End, v_操作员编号, v_操作员姓名, 0, n_号序, Decode(操作方式_In, 1, 0, 1),
         Decode(操作方式_In, 2, Null, v_操作员姓名), Decode(操作方式_In, 2, To_Date(Null), d_登记时间), 交易流水号_In, 交易说明_In, 合作单位_In,
         v_付款方式, Decode(操作方式_In, 1, Null, v_操作员姓名), Decode(操作方式_In, 1, Null, v_操作员编号));
    End If;
    --锁定单据不能产生队列
    If 锁定类型_In <> 1 Then
      n_预约生成队列 := 0;
      If 操作方式_In > 1 Then
        n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
      End If;
      --挂号和收费的预约都直接进入队列(收费预约缺少接收过程,所以直接和挂号一样直接进入队列)
      If 操作方式_In <> 2 Or n_预约生成队列 = 1 Then
        If Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113)) <> 0 Then
          --排队叫号模式:-0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
          If Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(r_安排.科室id, 0))) = 0 Or n_预约生成队列 = 1 Then
            n_分时点显示 := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
            If Nvl(操作方式_In, 0) > 1 And n_分时点显示 = 1 And n_启用分时段 = 1 Then
              n_分时点显示 := 1;
            Else
              n_分时点显示 := Null;
            End If;
            --产生队列
            --.按”执行部门” 的方式生成队列
            v_队列名称 := r_安排.科室id;
            v_排队号码 := Zlgetnextqueue(r_安排.科室id, n_挂号id, 号码_In || '|' || 号序_In);
            --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
            v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
            --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
            d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号码_In, 号序_In, d_Date);
            --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,v_排队标记,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
            Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, r_安排.科室id, v_排队号码, Null, r_Pati.姓名, r_Pati.病人id, v_诊室, r_安排.医生姓名,
                             d_排队时间, 预约方式_In, n_分时点显示, v_排队序号);
          End If;
        End If;
      End If;
    
      If Nvl(操作方式_In, 0) = 1 Then
        --处理票据使用情况
        If 票据号_In Is Not Null Then
          Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
          --发出票据
          Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 4, 单据号_In);
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
          Values
            (票据使用明细_Id.Nextval, Decode(收费票据_In, 1, 1, 4), 票据号_In, 1, 1, 领用id_In, n_打印id, d_登记时间, v_操作员姓名, 挂号金额合计_In);
          --状态改动
          Update 票据领用记录
          Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
          Where ID = Nvl(领用id_In, 0);
        End If;
        --病人本次就诊(以发生时间为准)
        If Nvl(r_Pati.病人id, 0) <> 0 Then
          Update 病人信息 Set 就诊时间 = 发生时间_In, 就诊状态 = 1, 就诊诊室 = v_诊室 Where 病人id = r_Pati.病人id;
        End If;
      End If;
    End If;
    --病人挂号汇总
    --解锁单据时不用再对汇总单据进行统计了 在锁定单据时已经进行了汇总
    If 锁定类型_In <> 2 Then
      --操作方式_IN:1-表示挂号,2-表示预约挂号不扣款,3-表示预约挂号,扣款
      --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收;3-收费预约
      --调用zl_third_lockno进行锁号，不建议使用本过程锁号
      n_预约 := Case
                When Nvl(操作方式_In, 0) = 1 Then
                 0
                When Nvl(操作方式_In, 0) = 2 Then
                 1
                When Nvl(操作方式_In, 0) = 3 Then
                 3
                Else
                 0
              End;
      Zl_病人挂号汇总_Update(r_安排.医生姓名, r_安排.医生id, r_安排.项目id, r_安排.科室id, 发生时间_In, n_预约, 号码_In);
    End If;
  
    If 锁定类型_In <> 1 Then
      --消息推送,锁号时不发送信息
      Begin
        Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
          Using 1, n_挂号id;
      Exception
        When Others Then
          Null;
      End;
      b_Message.Zlhis_Regist_001(n_挂号id, 单据号_In);
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_三方机构挂号_Insert;
/

--131243:余伟节,2018-09-17,身份证校验增加台湾地区码
Create Or Replace Function Zl_Fun_Checkidcard
(
  Idcard_In   In Varchar2,
  Calcdate_In In Date := Null
) Return Varchar2 Is
  -------------------------------------------------------------------------------
  --功能：身份证号码合法性校验,并返回身份证号的出生日期、性别、年龄
  --参数说明:
  -- 入参 IDcard_In:身份证号码
  --    Calcdate_In:计算日期,缺省时按系统时间
  -- 返回值：固定格式XML串
  --<OUTPUT>
  --       <BIRTHDAY></BIRTHDAY>                //出生日期
  --       <SEX></SEX>                  //性别
  --       <AGE></AGE>                //年龄
  --     <MSG></MSG>         //空串-身份证号有效(可从身份证号中获取出生日期和性别)，非空串-返回错误信息
  --</OUTPUT>
  -------------------------------------------------------------------------------
  n_Count     Number(5);
  n_Sum       Number(5);
  v_校验位    Varchar2(50);
  v_Pattern   Varchar2(500);
  v_Err_Msg   Varchar2(2000);
  v_性别      Varchar2(100);
  v_年龄      Varchar2(100);
  d_Curr_Time Date;
  d_出生日期  Date;
  v_Temp      Varchar2(20);

Begin
  Select Sysdate Into d_Curr_Time From Dual;

  If Idcard_In Is Null Then
    v_Err_Msg := '传入身份证号为空!';
    Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
  Else
    --身份证合法验证
    v_Pattern := '11,12,13,14,15,21,22,23,31,32,33,34,35,36,37,41,42,43,44,45,46,50,51,52,53,54,61,62,63,64,65,71,81,82,83,91';
    --地区检验
    If Instr(v_Pattern, Substr(Idcard_In, 1, 2)) = 0 Then
      v_Err_Msg := '身份证前两位地区码不正确!';
      Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
    End If;
    --身份证长度检查
    If Length(Idcard_In) = 15 Then
      --检查身份证号:15位身份证号要求全部为数字
      v_Pattern := '^\d{15}$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like(Idcard_In, v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '身份证中包含非法字符，请检查!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      End If;
      --获取性别
      If Mod(To_Number(Substr(Idcard_In, 15, 1)), 2) = 1 Then
        v_性别 := '男';
      Else
        v_性别 := '女';
      End If;
      --出生日期的合法性检查
      v_Pattern := '^19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like('19' || Substr(Idcard_In, 7, 6), v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '身份证中的出生日期无效，请检查!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      Else
        --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
        If Instr(',0229,0230,', ',' || Substr(Idcard_In, 9, 4) || ',') > 0 Then
          v_Temp     := '19' || Substr(Idcard_In, 7, 2) || '0301';
          d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
        Else
          d_出生日期 := To_Date('19' || Substr(Idcard_In, 7, 6), 'yyyy-mm-dd');
        End If;
        If d_出生日期 > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
          v_Err_Msg := '身份证中的出生日期无效，请检查!';
          Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
        End If;
      End If;
    Elsif Length(Idcard_In) = 18 Then
      -- 18 位身份证号前17 位全部为数字，最后1位可为数字或x
      v_Pattern := '^\d{17}[0-9Xx]$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like(Idcard_In, v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '身份证中包含非法字符!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      End If;
      --获取性别
      If Mod(To_Number(Substr(Idcard_In, 17, 1)), 2) = 1 Then
        v_性别 := '男';
      Else
        v_性别 := '女';
      End If;
      --出生日期的合法性检查
      v_Pattern := '^(1[6-9]|[2-9][0-9])[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like(Substr(Idcard_In, 7, 8), v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '身份证中的出生日期无效，请检查!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      Else
        --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
        If Instr(',0229,0230,', ',' || Substr(Idcard_In, 11, 4) || ',') > 0 Then
          v_Temp     := Substr(Idcard_In, 7, 4) || '0301';
          d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
        Else
          d_出生日期 := To_Date(Substr(Idcard_In, 7, 8), 'yyyy-mm-dd');
        End If;
        If d_出生日期 > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
          v_Err_Msg := '身份证中的出生日期无效，请检查!';
          Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
        End If;
        --计算校验位
        n_Sum     := (To_Number(Substr(Idcard_In, 1, 1)) + To_Number(Substr(Idcard_In, 11, 1))) * 7 +
                     (To_Number(Substr(Idcard_In, 2, 1)) + To_Number(Substr(Idcard_In, 12, 1))) * 9 +
                     (To_Number(Substr(Idcard_In, 3, 1)) + To_Number(Substr(Idcard_In, 13, 1))) * 10 +
                     (To_Number(Substr(Idcard_In, 4, 1)) + To_Number(Substr(Idcard_In, 14, 1))) * 5 +
                     (To_Number(Substr(Idcard_In, 5, 1)) + To_Number(Substr(Idcard_In, 15, 1))) * 8 +
                     (To_Number(Substr(Idcard_In, 6, 1)) + To_Number(Substr(Idcard_In, 16, 1))) * 4 +
                     (To_Number(Substr(Idcard_In, 7, 1)) + To_Number(Substr(Idcard_In, 17, 1))) * 2 +
                     To_Number(Substr(Idcard_In, 8, 1)) * 1 + To_Number(Substr(Idcard_In, 9, 1)) * 6 +
                     To_Number(Substr(Idcard_In, 10, 1)) * 3;
        n_Count   := Mod(n_Sum, 11);
        v_Pattern := '10X98765432';
        v_校验位  := Substr(v_Pattern, n_Count + 1, 1);
        If v_校验位 <> Upper(Substr(Idcard_In, 18, 1)) Then
          v_Err_Msg := '身份证号码不正确，请检查。';
          Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
        End If;
      End If;
    Else
      v_Err_Msg := '身份证长度不对,请检查。';
      Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
    End If;
    v_年龄 := Zl_Age_Calc(0, d_出生日期, Calcdate_In);
  End If;

  Return '<OUTPUT><BIRTHDAY>' || To_Char(d_出生日期, 'YYYY-MM-DD') || '</BIRTHDAY><SEX>' || v_性别 || '</SEX><AGE>' || v_年龄 || '</AGE><MSG></MSG></OUTPUT>';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Checkidcard;
/
--131243:余伟节,2018-09-17,身份证校验增加台湾地区码

Create Or Replace Procedure Zl_Third_Buildpatient
(
  Patiinfo_In  In Xmltype,
  Patiinfo_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------
  --参数说明:
  -- 入参 Patiinfo_In:
  --<IN>
  --  <ZJH></ZJH>                 //证件号，目前仅支持身份证号
  --  <ZJLX></ZJLX>                       //证件类型(目前仅支持身份证,为空时默认为身份证)
  --  <XM></XM>                       //姓名
  --  <SJH></SJH>                      //手机号
  --</IN>

  --出参 Patiinfo_Out：
  --<OUTPUT>
  --       <BRID></BRID>                //病人ID
  --       <MZH></MZH>                  //门诊号
  --     <ERROR></ERROR>         //如果有错误返回该节点
  --</OUTPUT>
  -------------------------------------------------------------------------------
  n_Pati_Id      病人信息.病人id%Type;
  n_Card_Type_Id 医疗卡类别.Id%Type;
  n_Count        Number(5);
  n_Sum          Number(5);
  v_校验位       Varchar2(50);

  v_姓名         病人信息.姓名%Type;
  v_身份证号     病人信息.身份证号%Type;
  v_手机号       病人信息.家庭电话%Type;
  v_性别         病人信息.性别%Type;
  v_年龄         病人信息.年龄%Type;
  v_操作员       人员表.姓名%Type;
  v_医疗付款方式 病人信息.医疗付款方式%Type;
  n_门诊号       病人信息.门诊号%Type;
  v_证件类型     医疗卡类别.名称%Type;
  v_证件号       病人医疗卡信息.卡号%Type;

  v_Pattern Varchar2(500);
  v_Temp    Varchar2(32767); --临时XML
  v_Err_Msg Varchar2(2000);
  n_存在    Number(2);

  d_出生日期  病人信息.出生日期%Type;
  d_Curr_Time Date;

  Err_Item Exception;
Begin
  Patiinfo_Out := Xmltype('<OUTPUT></OUTPUT>');
  Select Sysdate Into d_Curr_Time From Dual;

  --新建病人：姓名、身份证号、手机号（存在家庭电话中）、出生日期、性别、年龄(后面三项可从身份证中获取)。
  Select Extractvalue(Value(I), 'IN/XM'), Extractvalue(Value(I), 'IN/ZJH'), Extractvalue(Value(I), 'IN/SJH'),
         Extractvalue(Value(I), 'IN/ZJLX')
  Into v_姓名, v_证件号, v_手机号, v_证件类型
  From Table(Xmlsequence(Extract(Patiinfo_In, 'IN'))) I;

  Begin
    If v_证件类型 Is Null Then
      Select 病人id
      Into n_Pati_Id
      From 病人医疗卡信息
      Where 卡号 = v_证件号 And 卡类别id In (Select ID From 医疗卡类别 Where 名称 Like '%身份证%') And Rownum < 2;
    Else
      Select 病人id
      Into n_Pati_Id
      From 病人医疗卡信息
      Where 卡号 = v_证件号 And 卡类别id In (Select ID From 医疗卡类别 Where 名称 = v_证件类型) And Rownum < 2;
    End If;
    n_存在 := 1;
  Exception
    When Others Then
      n_存在 := 0;
  End;

  If Nvl(n_存在, 0) = 1 Then
    v_Temp := '<BRID>' || n_Pati_Id || '</BRID>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
    Select 门诊号 Into n_门诊号 From 病人信息 Where 病人id = n_Pati_Id;
    If n_门诊号 Is Null Then
      n_门诊号 := Nextno(3);
      Update 病人信息 Set 门诊号 = n_门诊号 Where 病人id = n_Pati_Id;
    End If;
    v_Temp := '<MZH>' || n_门诊号 || '</MZH>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
  Else
    If v_姓名 Is Null Then
      v_Err_Msg := '传入姓名为空!';
      Raise Err_Item;
    End If;
    If v_证件类型 Like '%身份证%' Or v_证件类型 Is Null Then
      v_身份证号 := v_证件号;
    Else
      v_Err_Msg := '目前不支持身份证以外的方式建档！';
      Raise Err_Item;
    End If;
  
    If v_身份证号 Is Null Then
      v_Err_Msg := '传入身份证号为空!';
      Raise Err_Item;
    Else
      --身份证合法验证
      v_Pattern := '11,12,13,14,15,21,22,23,31,32,33,34,35,36,37,41,42,43,44,45,46,50,51,52,53,54,61,62,63,64,65,71,81,82,83,91';
    
      --地区检验
      If Instr(v_Pattern, Substr(v_身份证号, 1, 2)) = 0 Then
        v_Err_Msg := '身份证前两位地区码不正确!';
        Raise Err_Item;
      End If;
      --身份证长度检查
      If Length(v_身份证号) = 15 Then
        --检查身份证号:15位身份证号要求全部为数字
        v_Pattern := '^\d{15}$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(v_身份证号, v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中包含非法字符，请检查!';
          Raise Err_Item;
        End If;
        --获取性别
        If Mod(To_Number(Substr(v_身份证号, 15, 1)), 2) = 1 Then
          v_性别 := '男';
        Else
          v_性别 := '女';
        End If;
        --出生日期的合法性检查
      
        v_Pattern := '^19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like('19' || Substr(v_身份证号, 7, 6), v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中的出生日期无效，请检查!';
          Raise Err_Item;
        Else
          --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
          If Instr(',0229,0230,', ',' || Substr(v_身份证号, 9, 4) || ',') > 0 Then
            v_Temp     := '19' || Substr(v_身份证号, 7, 2) || '0301';
            d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
          Else
            d_出生日期 := To_Date('19' || Substr(v_身份证号, 7, 6), 'yyyy-mm-dd');
          End If;
          If d_出生日期 > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
            v_Err_Msg := '身份证中的出生日期无效，请检查!';
            Raise Err_Item;
          End If;
        End If;
      Elsif Length(v_身份证号) = 18 Then
        -- 18 位身份证号前17 位全部为数字，最后1位可为数字或x
        v_Pattern := '^\d{17}[0-9Xx]$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(v_身份证号, v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中包含非法字符!';
          Raise Err_Item;
        End If;
        --获取性别
        If Mod(To_Number(Substr(v_身份证号, 17, 1)), 2) = 1 Then
          v_性别 := '男';
        Else
          v_性别 := '女';
        End If;
        --出生日期的合法性检查
        v_Pattern := '^(1[6-9]|[2-9][0-9])[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(Substr(v_身份证号, 7, 8), v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中的出生日期无效，请检查!';
          Raise Err_Item;
        Else
          --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
          If Instr(',0229,0230,', ',' || Substr(v_身份证号, 11, 4) || ',') > 0 Then
            v_Temp     := Substr(v_身份证号, 7, 4) || '0301';
            d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
          Else
            d_出生日期 := To_Date(Substr(v_身份证号, 7, 8), 'yyyy-mm-dd');
          End If;
          If d_出生日期 > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
            v_Err_Msg := '身份证中的出生日期无效，请检查!';
            Raise Err_Item;
          End If;
          --计算校验位
          n_Sum     := (To_Number(Substr(v_身份证号, 1, 1)) + To_Number(Substr(v_身份证号, 11, 1))) * 7 +
                       (To_Number(Substr(v_身份证号, 2, 1)) + To_Number(Substr(v_身份证号, 12, 1))) * 9 +
                       (To_Number(Substr(v_身份证号, 3, 1)) + To_Number(Substr(v_身份证号, 13, 1))) * 10 +
                       (To_Number(Substr(v_身份证号, 4, 1)) + To_Number(Substr(v_身份证号, 14, 1))) * 5 +
                       (To_Number(Substr(v_身份证号, 5, 1)) + To_Number(Substr(v_身份证号, 15, 1))) * 8 +
                       (To_Number(Substr(v_身份证号, 6, 1)) + To_Number(Substr(v_身份证号, 16, 1))) * 4 +
                       (To_Number(Substr(v_身份证号, 7, 1)) + To_Number(Substr(v_身份证号, 17, 1))) * 2 +
                       To_Number(Substr(v_身份证号, 8, 1)) * 1 + To_Number(Substr(v_身份证号, 9, 1)) * 6 +
                       To_Number(Substr(v_身份证号, 10, 1)) * 3;
          n_Count   := Mod(n_Sum, 11);
          v_Pattern := '10X98765432';
          v_校验位  := Substr(v_Pattern, n_Count + 1, 1);
          If v_校验位 <> Upper(Substr(v_身份证号, 18, 1)) Then
            v_Err_Msg := '身份证号码不正确，请检查。';
            Raise Err_Item;
          End If;
        End If;
      Else
        v_Err_Msg := '身份证长度不对,请检查。';
        Raise Err_Item;
      End If;
    
      If Nvl(v_年龄, '_') = '_' Then
        v_年龄 := Zl_Age_Calc(0, d_出生日期, d_Curr_Time);
      End If;
    End If;
  
    Select 名称 Into v_医疗付款方式 From 医疗付款方式 Where 缺省标志 = 1;
    n_Pati_Id := Nextno(1);
    n_门诊号  := Nextno(3);
    Insert Into 病人信息
      (病人id, 姓名, 身份证号, 家庭电话, 出生日期, 性别, 年龄, 登记时间, 门诊号, 医疗付款方式, 手机号)
      Select n_Pati_Id, v_姓名, v_身份证号, v_手机号, d_出生日期, v_性别, v_年龄, d_Curr_Time, n_门诊号, v_医疗付款方式, v_手机号



      From Dual;
    --病人信息保存完后，完成医疗卡绑定（二代身份证卡类别的绑定）
    Begin
      If v_证件类型 Is Null Then
        Select ID Into n_Card_Type_Id From 医疗卡类别 Where 名称 Like '%身份证%' And Rownum < 2;
      Else
        Select ID Into n_Card_Type_Id From 医疗卡类别 Where 名称 = v_证件类型 And Rownum < 2;
      End If;
    Exception
      When No_Data_Found Then
        v_Err_Msg := '身份证卡类别不存在！';
        Raise Err_Item;
    End;
    Select b.姓名 Into v_操作员 From 上机人员表 A, 人员表 B Where a.人员id = b.Id And a.用户名 = User;
  
    Zl_医疗卡变动_Insert(11, n_Pati_Id, n_Card_Type_Id, Null, v_身份证号, '创建虚拟卡', Null, v_操作员, d_Curr_Time);
  
    v_Temp := '<BRID>' || n_Pati_Id || '</BRID>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
    v_Temp := '<MZH>' || n_门诊号 || '</MZH>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
  End If;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Buildpatient;
/



------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0030' Where 编号=&n_System;
Commit;