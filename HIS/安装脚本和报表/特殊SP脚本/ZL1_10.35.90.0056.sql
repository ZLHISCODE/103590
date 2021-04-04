----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--139433:胡俊勇,2019-04-09,三方微生物报告修改
alter table 医嘱报告内容 add 是否禁止打印 number(1);

--120692:陈刘,2019-04-03,护理记录支持检验项目导入
create table 护理内容导入定义
(
类别 number(1),
名称 varchar2(100),
格式 varchar2(500)
)tablespace zl9BaseItem;
alter table 护理内容导入定义 add constraint 护理内容导入定义_PK primary key (类别) using index tablespace zl9Indexhis;

--139063:冉俊明,2019-04-01,门诊留观病人按门诊流程就诊
Alter Table 门诊费用记录 Add 病人病区id Number(18);


------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--120692:陈刘,2019-04-03,护理记录支持检验项目导入
Insert into zlTables(系统,表名,表空间,分类) Values(&n_System,'护理内容导入定义','ZL9BASEITEM','A2');

--110283:焦博,2019-04-02,增加系统参数指定发料部门时不显示无库存卫材来控制是否显示无库存卫材
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, -Null, 0, 0, 0, 0, 0, 0, 316, '指定发料部门时不显示无库存卫材', Null, '0',
         '在门诊收费（划价、记帐)或住院记帐、划价或医技站补费等录入卫生材料选择时，如果设置了缺省发料部门的，则在选择器不显示无库存的卫生材料。', '1-指定了缺省库房时显示有库存的卫材;0-不限制 ',
         '本参数需要配合参数"缺省发料部门"配合使用，设置了“缺省发料部门”时，本参数才有效。', '适用于当前部门无库存则不显示卫材的情况。', Null
  From Dual;

--115787:冉俊明,2019-04-01,增加一个私有模块参数包含门诊费用来控制查询病人费用信息时是否读取门诊费用
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1139, 1, 0, 0, 0, 0, 1, 21, '包含门诊费用', Null, '1',
         '调用病人费用查询模块查看病人费用信息时，是否包含病人门诊费用', '0-不包含门诊费用；1-包含门诊费用', '', '适用于调用病人费用查询模块查看病人费用信息时只查看住院费用信息的情况', Null
  From Dual;


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--120692:陈刘,2019-04-03,护理记录支持检验项目导入
Insert Into Zlprogprivs(系统, 序号, 功能, 所有者, 对象, 权限)Values(&n_System, 1255, '基本', User, '护理内容导入定义', 'SELECT');
Insert Into Zlprogprivs(系统, 序号, 功能, 所有者, 对象, 权限) Values (&n_System, 1255, '护理记录登记', User, 'Zl_护理内容导入定义_Update', 'EXECUTE');


-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--137062:冉俊明,2019-04-08,获取HIS结帐数据，返回结算明细
Create Or Replace Procedure Zl_Third_Getsettlement
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:获取HIS结帐数据
  --入参:Xml_In:
  --<IN>
  -- <BRID></BRID>       //病人ID
  -- <XM></XM>          //姓名
  -- <SFZH></SFZH>       //身份证号
  -- <ZYID></ZYID>         //主页ID
  -- <JSLX></JSLX>       //结算类型。1-门诊,2-住院。固定传2
  -- <JSKLB></JSKLB>       //结算卡类别
  --</IN>
  --出参:Xml_Out
  --<OUTPUT>
  --<JBXX>              //基本信息
  --   <XM></XM>           //姓名
  --   <XB></XB>           //性别
  --   <NL></NL>         //年龄
  --   <ZYH></ ZYH>        //住院号
  --   <ZYKS></ ZYKS>          //住院科室
  --   <KSID></KSID>         //科室ID
  --   <ZZYS></ ZZYS>          //主治医生
  --   <RYSJ></ RYSJ>          //入院时间
  --   <CYSJ></ CYSJ >         //出院时间
  --   <JZSJ></JZSJ>         //结帐时间(未结帐为空)
  --   <DJH></DJH>         //单据号(未结帐为空)
  --   <JSZFY></JSZFY>         //结算总费用
  --</JBXX>
  --<YJKLIST>              //冲抵预缴款集合
  --   <ITEM>
  --     <DJH><DJH>        //预交款单据号
  --     <JSFS></JSFS>     //结算方式（为名称，返回什么就取什么）
  --     <JE></JE>           //预缴款金额
  --     <JYLSH></JYLSH>       //交易流水号（便于冲销使用）
  --     <JYSM></JYSM>        //交易说明
  --     <SFJSK></SFJSK>       //是否结算卡，1-是，0-否。如果是由传入的卡类别缴费，返回1，否则返回0
  --     <ZFZT></ZFZT>        //支付状态：0-已支付，1-正在支付
  --   </ITEM>
  --</YJKLIST >
  --<TBQK>               //退补情况
  --   <TBLX></TBLX>         //退补类型(1:个人补款，2:医院退款)
  --   <TBJE></TBJE>         //退补金额
  --</TBQK>
  --<JSMX>                 //结算明细
  --  <ITEM>
  --    <JSFS></JSFS>         //结算方式
  --    <JSJE></JSJE>         //结算金额
  --    <SFYB></SFYB>         //是否医保结算方式,1-是，0-否
  --    <SFYJK></SFYJK>         //是否预交款,1-是，0-否
  --  </ITEM>
  --</JSMX>
  -- <ERROR><MSG></MSG></ERROR>    //出现错误时返回具体原因，error节点为空表示成功
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  n_病人id     病人信息.病人id%Type;
  v_姓名       病人信息.姓名%Type;
  v_身份证号   病人信息.身份证号%Type;
  n_主页id     病人信息.主页id%Type;
  n_结算类型   Number(3);
  n_卡类别id   医疗卡类别.Id%Type;
  v_结算卡类别 Varchar2(200);
  n_是否结清   Number(3); -- 1-未结清,0-结清
  n_结帐金额   住院费用记录.结帐金额%Type;
  v_Temp       Varchar2(32767); --临时XML
  v_Subtemp    Varchar2(32767);
  n_退补金额   病人预交记录.冲预交%Type;
  n_病人余额   病人预交记录.金额%Type;
  n_结帐id     病人预交记录.结帐id%Type;

  n_Number  Number(2);
  x_Templet Xmltype; --模板XML
  x_Temp    Xmltype;

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/BRID'), Extractvalue(Value(A), 'IN/ZYID'), Nvl(Extractvalue(Value(A), 'IN/JSLX'), 2),
         Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM')
  Into n_病人id, n_主页id, n_结算类型, v_结算卡类别, v_身份证号, v_姓名
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If n_结算类型 = 1 And Nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
  End If;
  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '无法确定病人信息,请检查!';
    Raise Err_Item;
  End If;

  Select Decode(Translate(Nvl(v_结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0) Into n_Number From Dual;
  If Nvl(n_Number, 0) = 1 Then
    Select Max(ID) Into n_卡类别id From 医疗卡类别 Where ID = To_Number(v_结算卡类别);
  Else
    Select Max(ID) Into n_卡类别id From 医疗卡类别 Where 名称 = v_结算卡类别;
  End If;
  If Nvl(n_卡类别id, 0) = 0 Then
    v_Err_Msg := '无法确认传入的结算卡,请检查!';
    Raise Err_Item;
  End If;

  If n_结算类型 = 2 Then
    Select Count(1)
    Into n_是否结清
    From (Select 1
           From 住院费用记录
           Where 病人id = n_病人id And 记录状态 <> 0 And 主页id = n_主页id And 记帐费用 = 1
           Group By 病人id, 主页id
           Having Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0)
    Where Rownum < 2;
  
    If n_是否结清 = 0 Then
      --结清,读取结帐数据
      For r_结帐 In (Select 姓名, 性别, 年龄, 住院号, 住院科室, 科室id, 主治医生, To_Char(入院时间, 'yyyy-mm-dd') As 入院时间,
                          To_Char(出院时间, 'yyyy-mm-dd') As 出院时间, To_Char(结帐时间, 'yyyy-mm-dd') As 结帐时间, 单据号, 结算总费用, 结帐id
                   From (Select c.姓名, c.性别, c.年龄, c.住院号, d.名称 As 住院科室, c.入院科室id As 科室id, c.住院医师 As 主治医生, c.入院日期 As 入院时间,
                                 c.出院日期 As 出院时间, a.收费时间 As 结帐时间, a.No As 单据号, a.结帐金额 As 结算总费用, a.Id As 结帐id
                          From 病人结帐记录 A, 病案主页 C, 部门表 D
                          Where a.记录状态 = 1 And Nvl(a.结算状态, 0) In (0, 2) And a.病人id = c.病人id And a.病人id = n_病人id And
                                a.主页id = n_主页id And a.主页id = c.主页id And c.入院科室id = d.Id(+) And Exists
                           (Select 1 From 病人预交记录 Where 结帐id = a.Id And 卡类别id = n_卡类别id)
                          Order By 结帐时间 Desc)
                   Where Rownum < 2) Loop
        v_Temp := '<XM>' || r_结帐.姓名 || '</XM>';
        v_Temp := v_Temp || '<XB>' || r_结帐.性别 || '</XB>';
        v_Temp := v_Temp || '<NL>' || r_结帐.年龄 || '</NL>';
        v_Temp := v_Temp || '<ZYH>' || r_结帐.住院号 || '</ZYH>';
        v_Temp := v_Temp || '<ZYKS>' || r_结帐.住院科室 || '</ZYKS>';
        v_Temp := v_Temp || '<KSID>' || r_结帐.科室id || '</KSID>';
        v_Temp := v_Temp || '<ZZYS>' || r_结帐.主治医生 || '</ZZYS>';
        v_Temp := v_Temp || '<RYSJ>' || r_结帐.入院时间 || '</RYSJ>';
        v_Temp := v_Temp || '<CYSJ>' || r_结帐.出院时间 || '</CYSJ>';
        v_Temp := v_Temp || '<JZSJ>' || r_结帐.结帐时间 || '</JZSJ>';
        v_Temp := v_Temp || '<DJH>' || r_结帐.单据号 || '</DJH>';
        v_Temp := v_Temp || '<JSZFY>' || r_结帐.结算总费用 || '</JSZFY>';
        v_Temp := '<JBXX>' || v_Temp || '</JBXX>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        n_结帐id := r_结帐.结帐id;
      End Loop;
    
      If n_结帐id Is Null Then
        v_Err_Msg := '该病人没有结帐数据!';
        Raise Err_Item;
      End If;
    
      --冲抵预缴款集合
      v_Temp := '';
      For r_预交 In (Select NO As 单据号, 结算方式, Sum(冲预交) As 金额, 卡类别id, 交易流水号, 交易说明, Decode(Nvl(校对标志, 0), 0, 0, 1) As 支付状态
                   From 病人预交记录
                   Where 结帐id = n_结帐id And Mod(记录性质, 10) = 1
                   Group By NO, 结算方式, 卡类别id, 交易流水号, 交易说明, Nvl(校对标志, 0)
                   Order By 单据号 Desc) Loop
        v_Temp := '<DJH>' || r_预交.单据号 || '</DJH>';
        v_Temp := v_Temp || '<JSFS>' || r_预交.结算方式 || '</JSFS>';
        v_Temp := v_Temp || '<JE>' || r_预交.金额 || '</JE>';
        v_Temp := v_Temp || '<JYLSH>' || r_预交.交易流水号 || '</JYLSH>';
        v_Temp := v_Temp || '<JYSM>' || r_预交.交易说明 || '</JYSM>';
        If n_卡类别id = r_预交.卡类别id Then
          v_Temp := v_Temp || '<SFJSK>' || 1 || '</SFJSK>';
        Else
          v_Temp := v_Temp || '<SFJSK>' || 0 || '</SFJSK>';
        End If;
        v_Temp    := v_Temp || '<ZFZT>' || r_预交.支付状态 || '</ZFZT>';
        v_Temp    := '<ITEM>' || v_Temp || '</ITEM>';
        v_Subtemp := v_Subtemp || v_Temp;
      End Loop;
      v_Subtemp := '<YJKLIST>' || v_Subtemp || '</YJKLIST>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Subtemp)) Into x_Templet From Dual;
    
      --退补情况
      Select Nvl(Sum(冲预交), 0)
      Into n_退补金额
      From 病人预交记录
      Where 结帐id = n_结帐id And Mod(记录性质, 10) = 2 And Nvl(校对标志, 0) = 0;
      If n_退补金额 < 0 Then
        v_Temp := '<TBLX>' || 2 || '</TBLX>';
      Else
        v_Temp := '<TBLX>' || 1 || '</TBLX>';
      End If;
      v_Temp := v_Temp || '<TBJE>' || Abs(n_退补金额) || '</TBJE>';
      v_Temp := '<TBQK>' || v_Temp || '</TBQK>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      --结算明细
      Select Xmlelement("JSMX",
                         Xmlagg(Xmlelement("ITEM",
                                            Xmlforest(结算方式 As "JSFS", 结算金额 As "JSJE", 是否医保 As "SFYB", 是否预交款 As "SFYJK"))))
      Into x_Temp
      From (Select a.结算方式, Sum(a.冲预交) As 结算金额, Decode(Mod(a.记录性质, 10), 1, 1, 0) As 是否预交款,
                    Max(Decode(Nvl(b.性质, 0), 3, 1, 4, 1, 0)) As 是否医保
             From 病人预交记录 A, 结算方式 B
             Where a.结算方式 = b.名称(+) And a.结帐id = n_结帐id
             Group By Decode(Mod(a.记录性质, 10), 1, 1, 0), a.结算方式);
      Select Appendchildxml(x_Templet, '/OUTPUT', x_Temp) Into x_Templet From Dual;
    Else
      --未结清，读取未结数据
      For r_Info In (Select c.姓名, c.性别, c.年龄, c.住院号, d.名称 As 住院科室, c.入院科室id As 科室id, c.住院医师 As 主治医生,
                            To_Char(c.入院日期, 'yyyy-mm-dd') As 入院时间, To_Char(c.出院日期, 'yyyy-mm-dd') As 出院时间
                     From 病案主页 C, 部门表 D
                     Where c.病人id = n_病人id And c.入院科室id = d.Id(+) And c.主页id = n_主页id And Rownum < 2) Loop
        v_Temp := '<XM>' || r_Info.姓名 || '</XM>';
        v_Temp := v_Temp || '<XB>' || r_Info.性别 || '</XB>';
        v_Temp := v_Temp || '<NL>' || r_Info.年龄 || '</NL>';
        v_Temp := v_Temp || '<ZYH>' || r_Info.住院号 || '</ZYH>';
        v_Temp := v_Temp || '<ZYKS>' || r_Info.住院科室 || '</ZYKS>';
        v_Temp := v_Temp || '<KSID>' || r_Info.科室id || '</KSID>';
        v_Temp := v_Temp || '<ZZYS>' || r_Info.主治医生 || '</ZZYS>';
        v_Temp := v_Temp || '<RYSJ>' || r_Info.入院时间 || '</RYSJ>';
        v_Temp := v_Temp || '<CYSJ>' || r_Info.出院时间 || '</CYSJ>';
        v_Temp := v_Temp || '<JZSJ>' || '' || '</JZSJ>';
        v_Temp := v_Temp || '<DJH>' || '' || '</DJH>';
      End Loop;
    
      Begin
        Select Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0))
        Into n_结帐金额
        From 住院费用记录
        Where 病人id = n_病人id And 记录状态 <> 0 And 主页id = n_主页id And 记帐费用 = 1;
      Exception
        When Others Then
          n_结帐金额 := 0;
      End;
      v_Temp := v_Temp || '<JSZFY>' || n_结帐金额 || '</JSZFY>';
      v_Temp := '<JBXX>' || v_Temp || '</JBXX>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      --冲抵预缴款集合
      v_Subtemp := '';
      For r_预交 In (Select a.No As 单据号, a.结算方式, Sum(Nvl(a.金额, 0)) - Sum(Nvl(a.冲预交, 0)) As 金额, a.卡类别id, a.交易流水号, a.交易说明,
                          Decode(Nvl(a.校对标志, 0), 0, 0, 1) As 支付状态
                   From 病人预交记录 A, 结算方式 B
                   Where a.结算方式 = b.名称(+) And a.病人id = n_病人id And Mod(a.记录性质, 10) = 1 And Nvl(a.预交类别, 2) = 2 And
                         (a.主页id = n_主页id Or a.主页id Is Null) And Nvl(b.性质, 0) <> 5
                   Group By a.No, a.结算方式, a.卡类别id, a.交易流水号, a.交易说明, Nvl(a.校对标志, 0)
                   Having Sum(Nvl(a.金额, 0)) - Sum(Nvl(a.冲预交, 0)) <> 0
                   Order By 单据号) Loop
        v_Temp := '<DJH>' || r_预交.单据号 || '</DJH>';
        v_Temp := v_Temp || '<JSFS>' || r_预交.结算方式 || '</JSFS>';
        v_Temp := v_Temp || '<JE>' || r_预交.金额 || '</JE>';
        v_Temp := v_Temp || '<JYLSH>' || r_预交.交易流水号 || '</JYLSH>';
        v_Temp := v_Temp || '<JYSM>' || r_预交.交易说明 || '</JYSM>';
        If n_卡类别id = r_预交.卡类别id Then
          v_Temp := v_Temp || '<SFJSK>' || 1 || '</SFJSK>';
        Else
          v_Temp := v_Temp || '<SFJSK>' || 0 || '</SFJSK>';
        End If;
        v_Temp     := v_Temp || '<ZFZT>' || r_预交.支付状态 || '</ZFZT>';
        v_Temp     := '<ITEM>' || v_Temp || '</ITEM>';
        v_Subtemp  := v_Subtemp || v_Temp;
        n_病人余额 := Nvl(n_病人余额, 0) + Nvl(r_预交.金额, 0);
      End Loop;
      v_Subtemp := '<YJKLIST>' || v_Subtemp || '</YJKLIST>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Subtemp)) Into x_Templet From Dual;
    
      --退补情况
      If Nvl(n_病人余额, 0) - Nvl(n_结帐金额, 0) > 0 Then
        v_Temp := '<TBLX>' || 2 || '</TBLX>';
      Else
        v_Temp := '<TBLX>' || 1 || '</TBLX>';
      End If;
      v_Temp := v_Temp || '<TBJE>' || Abs(Nvl(n_病人余额, 0) - Nvl(n_结帐金额, 0)) || '</TBJE>';
      v_Temp := '<TBQK>' || v_Temp || '</TBQK>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    End If;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getsettlement;
/

--139063:冉俊明,2019-04-08,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_门诊转住院_补结算转出
(
  No_In         费用补充记录.No%Type,
  费用冲销id_In 病人预交记录.结帐id%Type,
  结算冲销id_In 病人预交记录.结帐id%Type,
  结算序号_In   病人预交记录.结算序号%Type,
  退费时间_In   住院费用记录.发生时间%Type,
  操作员编号_In 住院费用记录.操作员编号%Type,
  操作员姓名_In 住院费用记录.操作员姓名%Type,
  主页id_In     病人预交记录.主页id%Type,
  入院科室id_In 病人预交记录.科室id%Type,
  结算方式_In   病人预交记录.结算方式%Type := Null,
  误差费_In     病人预交记录.冲预交%Type := Null
) As
  --功能：对费用补充结算的门诊费用进行转住院费用处理
  --入参：
  --  结算方式_In 不为空，表示所有除预交款的非医保金额全部退为指定的结算方式；
  --              为空，表示所有除预交款的非医保金额全部转为住院预交款
  Err_Item Exception;
  v_Err_Msg Varchar2(200);
  n_返回值  病人预交记录.冲预交%Type;

  n_组id   财务缴款分组.Id%Type;
  v_误差费 结算方式.名称%Type;
  n_误差费 病人预交记录.冲预交%Type;
  n_Dec    Number; --金额小数位数 

  v_Nos    Varchar2(4000);
  n_病人id 病人预交记录.病人id%Type;

  n_已退金额 病人预交记录.冲预交%Type;
  n_未退金额 病人预交记录.冲预交%Type;
  n_冲预交   病人预交记录.冲预交%Type;
  v_结算方式 Varchar2(4000);
  v_预交no   病人预交记录.No%Type;

  --保存预交款单据
  Procedure 病人预交记录_Insert
  (
    病人id_In     病人预交记录.病人id%Type,
    金额_In       病人预交记录.金额%Type,
    结算方式_In   病人预交记录.结算方式%Type,
    收款时间_In   病人预交记录.收款时间%Type,
    结算号码_In   病人预交记录.结算号码%Type,
    卡类别id_In   病人预交记录.卡类别id%Type := Null,
    卡号_In       病人预交记录.卡号%Type := Null,
    交易流水号_In 病人预交记录.交易流水号%Type := Null,
    交易说明_In   病人预交记录.交易说明%Type := Null
  ) As
    v_预交no 病人预交记录.No%Type;
    n_返回值 病人预交记录.金额%Type;
  Begin
    If Nvl(金额_In, 0) = 0 Or 结算方式_In Is Null Then
      Return;
    End If;
  
    --一卡通，每一笔都生成一条预交款记录
    --其它，同一种结算方式只生成一条预交款记录
    Update 病人预交记录
    Set 金额 = Nvl(金额, 0) + 金额_In
    Where 记录性质 = 1 And 记录状态 = 1 And 收款时间 = 收款时间_In And 病人id + 0 = 病人id_In And 结算方式 = 结算方式_In And Nvl(卡类别id, 0) = 0;
    If Sql%RowCount = 0 Or Nvl(卡类别id_In, 0) <> 0 Then
      v_预交no := Nextno(11);
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id, 预交类别,
         卡类别id, 卡号, 交易说明, 交易流水号, 结算号码)
      Values
        (病人预交记录_Id.Nextval, v_预交no, Null, 1, 1, 病人id_In, 主页id_In, 入院科室id_In, 金额_In, 结算方式_In, 收款时间_In, Null, Null, Null,
         操作员编号_In, 操作员姓名_In, '门诊转住院预交', n_组id, 2, 卡类别id_In, 卡号_In, 交易说明_In, 交易流水号_In, 结算号码_In);
    End If;
  
    Update 病人余额
    Set 预交余额 = Nvl(预交余额, 0) + 金额_In
    Where 性质 = 1 And 病人id = 病人id_In And 类型 = 2
    Returning 预交余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (病人id_In, 1, 2, 金额_In, 0);
      n_返回值 := 金额_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
    End If;
  End;
Begin
  n_组id := Zl_Get组id(操作员姓名_In);
  --误差费
  Begin
    Select 名称 Into v_误差费 From 结算方式 Where 性质 = 9 And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := '没有发现误差结算方式，请检查是否正确设置！';
      Raise Err_Item;
  End;
  n_误差费 := Nvl(误差费_In, 0);

  --金额小数位数 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  Select f_List2str(Cast(Collect(a.No) As t_Strlist), ',', 1), Max(a.病人id)
  Into v_Nos, n_病人id
  From 门诊费用记录 A, 费用补充记录 B
  Where a.结帐id = b.收费结帐id And b.记录性质 = 1 And b.附加标志 = 0 And b.No = No_In;
  If v_Nos Is Null Then
    v_Err_Msg := '未找到原医保补结算数据，费用转出失败!';
    Raise Err_Item;
  End If;

  --1.更新费用审核记录 
  Update 费用审核记录
  Set 记录状态 = 2
  Where 性质 = 1 And 费用id In (Select /*+cardinality(b,10)*/
                             a.Id
                            From 门诊费用记录 A, (Select Column_Value As NO From Table(f_Str2list(v_Nos))) B
                            Where a.No = b.No And Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3));

  --2.作废门诊费用记录 
  Update 门诊费用记录
  Set 记录状态 = 3
  Where Mod(记录性质, 10) = 1 And 记录状态 = 1 And NO In (Select Column_Value As NO From Table(f_Str2list(v_Nos)));

  For c_费用 In (Select /*+cardinality(b,10)*/
                a.No, a.序号, a.从属父号, a.价格父号, a.病人id, a.医嘱序号, a.门诊标志, a.姓名, a.性别, a.年龄, a.标识号, a.付款方式, a.病人科室id, a.费别,
                a.收费类别, a.收费细目id, a.计算单位, a.发药窗口, Sum(Nvl(a.付数, 1) * a.数次) As 数次, a.加班标志, a.附加标志, a.婴儿费, a.收入项目id, a.收据费目,
                a.标准单价, Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额, a.划价人, a.开单部门id, a.开单人, a.发生时间, a.执行部门id, a.执行人,
                Min(Decode(a.记录状态, 2, a.执行状态, 0)) - 1 As 执行状态, a.结论, Sum(a.结帐金额) As 结帐金额, Max(保险大类id) As 保险大类id,
                Max(保险项目否) As 保险项目否, Max(保险编码) As 保险编码, Max(费用类型) As 费用类型, Sum(a.统筹金额) As 统筹金额, Max(是否上传) As 是否上传, 是否急诊,
                a.挂号id, a.主页id, a.病人病区id
               From 门诊费用记录 A, (Select Column_Value As NO From Table(f_Str2list(v_Nos))) B
               Where a.No = b.No And a.记录性质 In (1, 11)
               Group By a.No, a.序号, a.从属父号, a.价格父号, a.病人id, a.医嘱序号, a.门诊标志, a.姓名, a.性别, a.年龄, a.标识号, a.付款方式, a.病人科室id,
                        a.费别, a.收费类别, a.收费细目id, a.计算单位, a.发药窗口, a.加班标志, a.附加标志, a.婴儿费, a.收入项目id, a.收据费目, a.标准单价, a.划价人,
                        a.开单部门id, a.开单人, a.发生时间, a.执行部门id, a.执行人, a.结论, 是否急诊, a.挂号id, a.主页id, a.病人病区id
               Having Nvl(Sum(Nvl(a.付数, 1) * a.数次), 0) <> 0) Loop
  
    Insert Into 门诊费用记录
      (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数,
       发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间,
       结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 缴款组id, 费用状态, 挂号id, 主页id, 病人病区id)
    Values
      (病人费用记录_Id.Nextval, 1, c_费用.No, 2, c_费用.序号, c_费用.从属父号, c_费用.价格父号, c_费用.病人id, c_费用.医嘱序号, c_费用.门诊标志, c_费用.姓名,
       c_费用.性别, c_费用.年龄, c_费用.标识号, c_费用.付款方式, c_费用.病人科室id, c_费用.费别, c_费用.收费类别, c_费用.收费细目id, c_费用.计算单位, 1, c_费用.发药窗口,
       -1 * c_费用.数次, c_费用.加班标志, c_费用.附加标志, c_费用.婴儿费, c_费用.收入项目id, c_费用.收据费目, c_费用.标准单价, -1 * c_费用.应收金额, -1 * c_费用.实收金额,
       c_费用.划价人, c_费用.开单部门id, c_费用.开单人, c_费用.发生时间, 退费时间_In, c_费用.执行部门id, c_费用.执行人, c_费用.执行状态, Null, c_费用.结论, 操作员编号_In,
       操作员姓名_In, 费用冲销id_In, -1 * c_费用.结帐金额, c_费用.保险大类id, c_费用.保险项目否, c_费用.保险编码, c_费用.费用类型, -1 * c_费用.统筹金额, c_费用.是否上传, '',
       c_费用.是否急诊, n_组id, 0, c_费用.挂号id, c_费用.主页id, c_费用.病人病区id);
  End Loop;
  Zl_门诊退费结算_Modify(1, n_病人id, 费用冲销id_In, Null);

  --3.作废补充结算记录（同时已进行了票据回收和医保原样退）
  Zl_费用补充记录_Delete(No_In, 结算冲销id_In, Null, 结算序号_In, 费用冲销id_In, 操作员编号_In, 操作员姓名_In, 退费时间_In);
  Update 费用补充记录 Set 费用状态 = 0 Where 结算序号 = 结算序号_In;
  --处理为医保接口已调用成功
  Update 病人预交记录
  Set 校对标志 = 2
  Where 记录性质 = 6 And 结帐id = 结算冲销id_In And 结算方式 In (Select 名称 From 结算方式 Where 性质 In (3, 4));

  --4.结算数据处理
  Select -1 * Nvl(Sum(a.冲预交), 0)
  Into n_未退金额
  From 病人预交记录 A
  Where a.结算序号 = 结算序号_In And a.结算方式 Is Null;
  If Nvl(n_误差费, 0) = 0 Then
    n_误差费 := Round(n_未退金额, n_Dec) - n_未退金额;
  End If;
  n_未退金额 := n_未退金额 - n_误差费;

  For r_预交 In (Select Case
                        When Mod(a.记录性质, 10) = 1 Then
                         1
                        When Nvl(a.卡类别id, 0) <> 0 Then
                         2
                        Else
                         0
                      End As 类型, a.结帐id, Nvl(a.冲预交, 0) As 冲预交, a.No, a.病人id, a.结算方式, a.卡类别id, a.卡号, a.交易流水号, a.交易说明,
                      a.结算号码
               From 病人预交记录 A, 结算方式 B
               Where a.结算方式 = b.名称 And a.记录状态 In (1, 3) And b.性质 Not In (3, 4, 9) And
                     a.结帐id In (Select 收费结帐id From 费用补充记录 Where 记录性质 = 1 And 附加标志 = 0 And NO = No_In)) Loop
  
    --都是单种结算方式
    If r_预交.类型 = 1 Then
      --预交款
      Zl_费用补充结算_完成退费(结算冲销id_In, Null, Null, Null, Null, Null, n_误差费, 0, 0, -1 * n_未退金额);
      Exit;
    Elsif r_预交.类型 = 2 Then
      --一卡通
      Select Nvl(Sum(金额), 0) Into n_已退金额 From 三方退款信息 Where 记录id = r_预交.结帐id;
      If r_预交.冲预交 - n_已退金额 > 0 Then
        If r_预交.冲预交 - n_已退金额 > n_未退金额 Then
          n_冲预交 := n_未退金额;
        Else
          n_冲预交 := r_预交.冲预交 - n_已退金额;
        End If;
      
        v_结算方式 := r_预交.结算方式 || '|' || -1 * n_冲预交 || '| | ';
        Zl_费用补充结算_完成退费(结算冲销id_In, v_结算方式, r_预交.卡类别id, r_预交.卡号, r_预交.交易流水号, r_预交.交易说明, n_误差费, 0, 1);
        Zl_三方退款信息_Insert(结算序号_In, r_预交.结帐id, n_冲预交, r_预交.卡号, r_预交.交易流水号, r_预交.交易说明);
      
        --转为住院预交款
        病人预交记录_Insert(r_预交.病人id, n_冲预交, r_预交.结算方式, 退费时间_In, r_预交.结算号码, r_预交.卡类别id, r_预交.卡号, r_预交.交易流水号, r_预交.交易说明);
      
        n_未退金额 := n_未退金额 - n_冲预交;
        n_误差费   := 0;
      End If;
      If n_未退金额 = 0 Then
        Exit;
      End If;
    Else
      --其它非医保结算方式
      --结算方式|结算金额|结算号码|结算摘要
      v_结算方式 := r_预交.结算方式 || '|' || n_未退金额 || '| | ';
      Zl_费用补充结算_完成退费(结算冲销id_In, v_结算方式, Null, Null, Null, Null, n_误差费, 0);
    
      --转为住院预交款
      病人预交记录_Insert(r_预交.病人id, n_未退金额, r_预交.结算方式, 退费时间_In, r_预交.结算号码);
      Exit;
    End If;
  End Loop;

  --5.转出完成处理   
  Delete From 病人预交记录 Where 结帐id = 结算冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '还存在未缴款的数据，不能完成结算！';
    Raise Err_Item;
  End If;
  Delete From 病人预交记录 Where 结帐id = 费用冲销id_In And 结算方式 Is Null And Nvl(冲预交, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '还存在未缴款的数据，不能完成结算！';
    Raise Err_Item;
  End If;
  Update 病人预交记录 Set 校对标志 = 0, 会话号 = Null Where 结算序号 = 结算序号_In;

  --人员缴款余额（主要是医保）
  For c_预交 In (Select a.结算方式, a.操作员姓名, Nvl(Sum(a.冲预交), 0) As 冲预交
               From 病人预交记录 A, 结算方式 B
               Where a.结算方式 = b.名称 And b.性质 In (3, 4) And a.结算序号 = 结算序号_In
               Group By a.结算方式, a.操作员姓名) Loop
  
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + c_预交.冲预交
    Where 收款员 = c_预交.操作员姓名 And 性质 = 1 And 结算方式 = c_预交.结算方式
    Returning 余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (c_预交.操作员姓名, c_预交.结算方式, 1, c_预交.冲预交);
      n_返回值 := c_预交.冲预交;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = c_预交.操作员姓名 And 性质 = 1 And 结算方式 = c_预交.结算方式 And Nvl(余额, 0) = 0;
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊转住院_补结算转出;
/

--139063:冉俊明,2019-04-08,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_门诊转住院_收费转出
(
  No_In         住院费用记录.No%Type,
  操作员编号_In 住院费用记录.操作员编号%Type,
  操作员姓名_In 住院费用记录.操作员姓名%Type,
  退费时间_In   住院费用记录.发生时间%Type,
  门诊退费_In   Number := 0,
  入院科室id_In 住院费用记录.开单部门id%Type := Null,
  主页id_In     住院费用记录.主页id%Type := Null,
  结算方式_In   病人预交记录.结算方式%Type := Null,
  结帐id_In     病人预交记录.结帐id%Type := Null,
  原结帐id_In   病人预交记录.结帐id%Type := Null,
  误差费_In     病人预交记录.冲预交%Type := Null
) As
  --门诊退费_In:0-门诊转住院立即销帐;1-门诊退费模式
  -- 门诊退费_In为1时:入院科室id_In和主页ID_IN可以不传入
  n_Count      Number(5);
  n_原结帐id   住院费用记录.结帐id%Type;
  n_实收金额   门诊费用记录.实收金额%Type;
  n_实际冲销   病人预交记录.冲预交%Type;
  n_组id       财务缴款分组.Id%Type;
  n_病人id     病人信息.病人id%Type;
  v_预交no     病人预交记录.No%Type;
  n_预交金额   病人预交记录.冲预交%Type;
  n_打印id     票据使用明细.打印id%Type;
  n_开单部门id 住院费用记录.开单部门id%Type;
  v_开单人     门诊费用记录.开单人%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  v_误差费     结算方式.名称%Type;
  n_返回值     病人余额.费用余额%Type;
  v_结算方式   结算方式.名称%Type;
  v_Nos        Varchar2(3000);
  v_结帐ids    Varchar2(3000);
  v_原结帐ids  Varchar2(3000);
  n_Tempid     病人预交记录.Id%Type;
  n_医保       Number;
  n_存在       Number;
  n_退现       Number;
  n_部分退费   Number;
  n_退费条数   Number;
  n_费用状态   门诊费用记录.费用状态%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(200);
Begin
  n_组id := Zl_Get组id(操作员姓名_In);
  --误差费
  Begin
    Select 名称 Into v_误差费 From 结算方式 Where 性质 = 9 And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := '没有发现误差结算方式，请检查是否正确设置！';
      Raise Err_Item;
  End;

  If 原结帐id_In Is Null Then
  
    Select Count(NO), Sum(实收金额)
    Into n_Count, n_实收金额
    From 门诊费用记录
    Where NO = No_In And Mod(记录性质, 10) = 1;
    If n_Count = 0 Or n_实收金额 = 0 Then
      v_Err_Msg := '单据' || No_In || '不是收费单据或因并发原因他人操作了该单据,不能转为住院费用.';
      Raise Err_Item;
    End If;
  
    Select 结帐id, 病人id, 开单部门id, 开单人
    Into n_原结帐id, n_病人id, n_开单部门id, v_开单人
    From 门诊费用记录
    Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And Rownum < 2;
  
    --1.1作废费用记录
    If 结帐id_In Is Null Then
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    Else
      n_结帐id := 结帐id_In;
    End If;
  
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id,
       计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态, 执行时间,
       操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id, 费用状态, 主页id, 病人病区id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id,
             收费类别, 收费细目id, 计算单位, 付数, 发药窗口, -1 * 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, -1 * 应收金额, -1 * 实收金额, 开单部门id,
             开单人, 执行部门id, 划价人, 执行人, -1, 执行时间, 操作员编号_In, 操作员姓名_In, 发生时间, 退费时间_In, n_结帐id, -1 * 结帐金额, 保险项目否, 保险大类id, 统筹金额,
             摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id, 0, 主页id, 病人病区id
      From 门诊费用记录
      Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 = 1;
  
    --Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And 记录性质 = 1 And 记录状态 = 1;
  
    --1.2作废预交记录
    --作废冲预交部分
    For r_结账id In (Select Distinct 结帐id
                   From 门诊费用记录
                   Where NO In (Select Distinct NO
                                From 门诊费用记录
                                Where 结帐id In (Select 结帐id
                                               From 病人预交记录
                                               Where 结算序号 In (Select b.结算序号
                                                              From 门诊费用记录 A, 病人预交记录 B
                                                              Where a.No = No_In And b.结算序号 < 0 And Mod(a.记录性质, 10) = 1 And
                                                                    a.记录状态 <> 0 And a.结帐id = b.结帐id))) And
                         Mod(记录性质, 10) = 1 And 记录状态 <> 0
                   Union
                   Select Distinct 结帐id
                   From 门诊费用记录
                   Where NO In (Select Distinct NO
                                From 门诊费用记录
                                Where 结帐id In (Select a.结帐id
                                               From 门诊费用记录 A, 病人预交记录 B
                                               Where a.No = No_In And b.结算序号 > 0 And Mod(a.记录性质, 10) = 1 And a.记录状态 <> 0 And
                                                     a.结帐id = b.结帐id)) And Mod(记录性质, 10) = 1 And 记录状态 <> 0) Loop
      v_原结帐ids := v_原结帐ids || ',' || r_结账id.结帐id;
    End Loop;
    v_原结帐ids := Substr(v_原结帐ids, 2);
  
    Begin
      Select 1
      Into n_医保
      From 保险结算记录
      Where 记录id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And Rownum < 2;
    Exception
      When Others Then
        n_医保 := 0;
    End;
  
    If n_医保 = 1 Then
      Begin
        Select 1
        Into n_存在
        From 医保结算明细
        Where NO = No_In And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And Rownum < 2;
      Exception
        When Others Then
          v_Err_Msg := '当前单据' || No_In || '不存在医保结算明细,无法进行门诊转住院!';
          Raise Err_Item;
      End;
    End If;
  
    --医保退款
    For r_医保 In (Select 结帐id, NO, 结算方式, 金额, 备注
                 From 医保结算明细
                 Where NO = No_In And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids)))) Loop
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) - r_医保.金额
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_医保.结算方式
      Returning 余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (操作员姓名_In, r_医保.结算方式, 1, -1 * r_医保.金额);
        n_返回值 := r_医保.金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_医保.结算方式 And Nvl(余额, 0) = 0;
      End If;
    
      Update 病人预交记录
      Set 冲预交 = 冲预交 + (-1 * r_医保.金额)
      Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = r_医保.结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
           缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
        Values
          (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_医保.金额, r_医保.结算方式, Null, 退费时间_In,
           Null, Null, Null, 操作员编号_In, 操作员姓名_In, r_医保.备注, n_组id, Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id,
           0, 3);
      End If;
    
      Update 病人预交记录
      Set 记录状态 = 3
      Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
            结算方式 = r_医保.结算方式;
    
      Update 医保结算明细
      Set 金额 = 金额 + (-1 * r_医保.金额)
      Where NO = No_In And 结帐id = n_结帐id And 结算方式 = r_医保.结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 医保结算明细
          (结帐id, NO, 结算方式, 金额)
        Values
          (n_结帐id, No_In, r_医保.结算方式, -1 * r_医保.金额);
      End If;
      n_实收金额 := n_实收金额 - r_医保.金额;
    End Loop;
  
    Begin
      Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
    Exception
      When Others Then
        Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
    End;
  
    If n_实收金额 <> 0 Then
      For r_Prepay In (Select NO, 实际票号, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, Sum(冲预交) As 冲预交, 卡类别id, 结算卡序号,
                              卡号, 交易流水号, 交易说明, 合作单位
                       From 病人预交记录 A
                       Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids)))
                       Group By n_Tempid, NO, 实际票号, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 卡类别id, 结算卡序号, 卡号,
                                交易流水号, 交易说明, 合作单位) Loop
        If n_实收金额 <> 0 Then
          If r_Prepay.冲预交 >= n_实收金额 Then
            Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结算序号, 缴款组id)
              Select n_Tempid, r_Prepay.No, r_Prepay.实际票号, 11, 1, r_Prepay.病人id, r_Prepay.主页id, r_Prepay.科室id, Null,
                     r_Prepay.结算方式, r_Prepay.结算号码, Null, r_Prepay.缴款单位, r_Prepay.单位开户行, r_Prepay.单位帐号, 退费时间_In, 操作员姓名_In,
                     操作员编号_In, -1 * n_实收金额, n_结帐id, r_Prepay.卡类别id, r_Prepay.结算卡序号, r_Prepay.卡号, r_Prepay.交易流水号,
                     r_Prepay.交易说明, r_Prepay.合作单位, 1, -1 * n_结帐id, n_组id
              From Dual;
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + Nvl(n_实收金额, 0)
            Where 病人id = n_病人id And 类型 = 1 And 性质 = 1
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (n_病人id, 1, n_实收金额, 1);
              n_返回值 := n_实收金额;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
            n_实收金额 := 0;
          Else
            Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结算序号, 缴款组id)
              Select n_Tempid, r_Prepay.No, r_Prepay.实际票号, 11, 1, r_Prepay.病人id, r_Prepay.主页id, r_Prepay.科室id, Null,
                     r_Prepay.结算方式, r_Prepay.结算号码, Null, r_Prepay.缴款单位, r_Prepay.单位开户行, r_Prepay.单位帐号, 退费时间_In, 操作员姓名_In,
                     操作员编号_In, -1 * r_Prepay.冲预交, n_结帐id, r_Prepay.卡类别id, r_Prepay.结算卡序号, r_Prepay.卡号, r_Prepay.交易流水号,
                     r_Prepay.交易说明, r_Prepay.合作单位, 1, -1 * n_结帐id, n_组id
              From Dual;
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + Nvl(r_Prepay.冲预交, 0)
            Where 病人id = n_病人id And 类型 = 1 And 性质 = 1
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (n_病人id, 1, r_Prepay.冲预交, 1);
              n_返回值 := r_Prepay.冲预交;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
            n_实收金额 := n_实收金额 - r_Prepay.冲预交;
          End If;
        End If;
      End Loop;
    End If;
    --2.票据收回
    --可能以前没有打印,无收回
    Select Nvl(Max(ID), 0)
    Into n_打印id
    From (Select b.Id
           From 票据使用明细 A, 票据打印内容 B
           Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 1 And b.No = No_In
           Order By a.使用时间 Desc)
    Where Rownum < 2;
    If n_打印id > 0 Then
      --多张单据循环调用时只能收回一次
      Select Count(打印id) Into n_Count From 票据使用明细 Where 票种 = 1 And 性质 = 2 And 打印id = n_打印id;
      If n_Count = 0 Then
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
          Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, 退费时间_In, 操作员姓名_In, 票据金额
          From 票据使用明细
          Where 打印id = n_打印id And 票种 = 1 And 性质 = 1;
      End If;
    End If;
  
    --3.缴款数据处理(
    --   现有两种情况:
    --    1. 转出过程直接销帐的,则缴款数据不增加;
    --    2. 先转出,再到门诊退款退票,则需要进行缴款数据处理
    If Nvl(门诊退费_In, 0) = 1 Then
      For c_预交 In (Select a.结算方式, Sum(a.冲预交) As 冲预交, 2 As 预交类别, a.卡类别id, a.结算卡序号, a.卡号, Min(a.交易流水号) As 交易流水号,
                          Min(a.交易说明) As 交易说明, Min(a.合作单位) As 合作单位, b.性质
                   From 病人预交记录 A, 结算方式 B
                   Where a.记录性质 = 3 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                         a.结算方式 = b.名称 And b.性质 In (1, 2, 7, 8) And a.结算方式 Is Not Null
                   Group By a.结算方式, 预交类别, a.卡类别id, a.结算卡序号, a.卡号, b.性质
                   Having Sum(a.冲预交) <> 0
                   Order By a.卡类别id, 性质 Desc) Loop
        If n_实收金额 <> 0 Then
          Begin
            Select 是否退现 Into n_退现 From 医疗卡类别 Where ID = c_预交.卡类别id;
          Exception
            When Others Then
              n_退现 := 0;
          End;
          If (c_预交.性质 = 7 Or (c_预交.性质 = 8 And c_预交.卡类别id Is Not Null)) And n_退现 = 0 Then
            If c_预交.冲预交 > n_实收金额 Then
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * n_实收金额), 摘要 = 摘要 || '1' || ',' || c_预交.卡类别id || ',' || -1 * n_实收金额 || '|'
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实收金额, Null, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || c_预交.卡类别id || ',' || -1 * n_实收金额 || '|', n_组id,
                   Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
              End If;
              Update 病人预交记录
              Set 记录状态 = 3
              Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                    结算方式 = c_预交.结算方式;
              n_费用状态 := 1;
              n_实收金额 := 0;
            Else
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * c_预交.冲预交), 摘要 = 摘要 || '1' || ',' || c_预交.卡类别id || ',' || -1 * c_预交.冲预交 || '|'
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * c_预交.冲预交, Null, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || c_预交.卡类别id || ',' || -1 * c_预交.冲预交 || '|', n_组id,
                   Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
              End If;
            
              Update 病人预交记录
              Set 记录状态 = 3
              Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                    结算方式 = c_预交.结算方式;
              n_费用状态 := 1;
              n_实收金额 := n_实收金额 - c_预交.冲预交;
            End If;
          Else
            n_实际冲销 := 0;
            If c_预交.性质 In (3, 4) Or (c_预交.性质 = 8 And c_预交.结算卡序号 Is Not Null) Then
              v_结算方式 := c_预交.结算方式;
            Else
              If 结算方式_In Is Null Then
                Begin
                  Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
                Exception
                  When Others Then
                    Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
                End;
              Else
                v_结算方式 := 结算方式_In;
              End If;
            End If;
          
            If c_预交.性质 = 8 And c_预交.结算卡序号 Is Not Null Then
              If n_实收金额 >= c_预交.冲预交 Then
                --Zl_Square_Update(v_原结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, c_预交.冲预交, c_预交.结算卡序号);
                Update 病人预交记录
                Set 冲预交 = 冲预交 + (-1 * c_预交.冲预交), 摘要 = 摘要 || '0' || ',' || c_预交.结算卡序号 || ',' || -1 * c_预交.冲预交 || '|'
                Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into 病人预交记录
                    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                     摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                  Values
                    (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * c_预交.冲预交, Null, Null,
                     退费时间_In, Null, Null, Null, 操作员编号_In, 操作员姓名_In,
                     '0' || ',' || c_预交.结算卡序号 || ',' || -1 * c_预交.冲预交 || '|', n_组id, Null, Null, Null, Null, Null, Null,
                     n_结帐id, -1 * n_结帐id, 3, 1);
                End If;
                n_费用状态 := 1;
                n_实际冲销 := c_预交.冲预交;
              Else
                --Zl_Square_Update(v_原结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, n_实收金额, c_预交.结算卡序号);
                Update 病人预交记录
                Set 冲预交 = 冲预交 + (-1 * n_实收金额), 摘要 = 摘要 || '0' || ',' || c_预交.结算卡序号 || ',' || -1 * n_实收金额 || '|'
                Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into 病人预交记录
                    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                     摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                  Values
                    (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实收金额, Null, Null, 退费时间_In,
                     Null, Null, Null, 操作员编号_In, 操作员姓名_In, '0' || ',' || c_预交.结算卡序号 || ',' || -1 * n_实收金额 || '|', n_组id,
                     Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
                End If;
                n_费用状态 := 1;
                n_实际冲销 := n_实收金额;
              End If;
            Else
              If c_预交.冲预交 > n_实收金额 Then
                n_实际冲销 := n_实收金额;
              Else
                n_实际冲销 := c_预交.冲预交;
              End If;
            End If;
          
            If c_预交.结算卡序号 Is Null Then
              Update 人员缴款余额
              Set 余额 = Nvl(余额, 0) - n_实际冲销
              Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
              Returning 余额 Into n_返回值;
              If Sql%RowCount = 0 Then
                Insert Into 人员缴款余额
                  (收款员, 结算方式, 性质, 余额)
                Values
                  (操作员姓名_In, v_结算方式, 1, -1 * n_实际冲销);
                n_返回值 := n_实际冲销;
              End If;
              If Nvl(n_返回值, 0) = 0 Then
                Delete From 人员缴款余额
                Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式 And Nvl(余额, 0) = 0;
              End If;
            
              --退原预交记录
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * n_实际冲销)
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实际冲销, v_结算方式, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, c_预交.合作单位, n_结帐id,
                   -1 * n_结帐id, 0, 3);
              End If;
            End If;
            Update 病人预交记录
            Set 记录状态 = 3
            Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                  结算方式 = c_预交.结算方式;
            n_实收金额 := n_实收金额 - n_实际冲销;
          End If;
        End If;
      End Loop;
    
      --更新费用审核记录
      Update 费用审核记录
      Set 记录状态 = 2
      Where 费用id In (Select ID From 门诊费用记录 Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3)) And 性质 = 1;
      --作废门诊记录
      Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 = 1;
      For r_Clinic In (Select 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型,
                              发药窗口, 付数, Sum(数次) As 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, Sum(应收金额) As 应收金额,
                              Sum(实收金额) As 实收金额, Sum(统筹金额) As 统筹金额, 开单部门id, 开单人, 执行部门id, 划价人, Max(记帐单id) As 记帐单id, 发生时间,
                              实际票号
                       From 门诊费用记录
                       Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 In (2, 3) And Nvl(附加标志, 0) Not In (8, 9)
                       Group By 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码,
                                费用类型, 发药窗口, 付数, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 开单部门id, 开单人, 执行部门id, 划价人, 发生时间, 实际票号
                       Having Sum(数次) <> 0) Loop
        Insert Into 门诊费用记录
          (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否,
           保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id, 开单人, 发生时间,
           登记时间, 执行部门id, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 缴款组id, 结帐id, 结帐金额, 费用状态)
        Values
          (病人费用记录_Id.Nextval, 1, No_In, r_Clinic.实际票号, 2, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号, 1, r_Clinic.病人id,
           '', r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, r_Clinic.病人科室id, r_Clinic.费别, r_Clinic.收费类别, r_Clinic.收费细目id,
           r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id, r_Clinic.保险编码, r_Clinic.费用类型, r_Clinic.发药窗口, r_Clinic.付数,
           -1 * r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志, r_Clinic.收入项目id, r_Clinic.收据费目, r_Clinic.标准单价,
           -1 * r_Clinic.应收金额, -1 * r_Clinic.实收金额, -1 * r_Clinic.统筹金额, 0, r_Clinic.开单部门id, r_Clinic.开单人, r_Clinic.发生时间,
           退费时间_In, r_Clinic.执行部门id, r_Clinic.划价人, 操作员编号_In, 操作员姓名_In, r_Clinic.记帐单id, '', n_组id, n_结帐id,
           -1 * r_Clinic.实收金额, 0);
      End Loop;
    Else
      --4.退款转预交(不产生票据,由操作员通过重打进行)
      For r_Pay In (Select Min(a.Id) As 预交id, a.结算方式, Sum(a.冲预交) As 冲预交, 2 As 预交类别, a.卡类别id, a.结算卡序号, a.卡号, a.交易流水号,
                           a.交易说明, a.合作单位, b.性质
                    From 病人预交记录 A, 结算方式 B
                    Where a.记录性质 = 3 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                          a.结算方式 = b.名称 And (b.性质 In (1, 2, 7, 8)) And a.结算方式 Is Not Null
                    Group By a.结算方式, 预交类别, a.卡类别id, a.结算卡序号, a.卡号, b.性质, a.交易流水号, a.交易说明, a.合作单位

                    
                    Having Sum(a.冲预交) <> 0
                    Order By a.卡类别id, 性质 Desc) Loop
        --4.1产生预交款单据 (不存在部分退费的情况)
        --所有单据,按规则生成预交款单据
        --因为收款后立即缴款,所以人员缴款余额无变化
        If n_实收金额 <> 0 Then
          If r_Pay.性质 = 7 Or (r_Pay.性质 = 8 And r_Pay.卡类别id Is Not Null) Then
            If r_Pay.冲预交 > n_实收金额 Then
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * n_实收金额), 摘要 = 摘要 || '1' || ',' || r_Pay.卡类别id || ',' || -1 * n_实收金额 || '|'
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实收金额, Null, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || r_Pay.卡类别id || ',' || -1 * n_实收金额 || '|', n_组id,
                   Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
              End If;
            
              Update 病人预交记录
              Set 记录状态 = 3
              Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                    结算方式 = r_Pay.结算方式;
              n_费用状态 := 1;
              n_实收金额 := 0;
            Else
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * r_Pay.冲预交), 摘要 = 摘要 || '1' || ',' || r_Pay.卡类别id || ',' || -1 * r_Pay.冲预交 || '|'
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_Pay.冲预交, Null, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || r_Pay.卡类别id || ',' || -1 * r_Pay.冲预交 || '|',
                   n_组id, Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
              End If;
            
              Update 病人预交记录
              Set 记录状态 = 3
              Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                    结算方式 = r_Pay.结算方式;
              n_费用状态 := 1;
              n_实收金额 := n_实收金额 - r_Pay.冲预交;
            End If;
          Else
            n_实际冲销 := 0;
            If r_Pay.性质 In (3, 4) Or (r_Pay.性质 = 8 And r_Pay.结算卡序号 Is Not Null) Then
              v_结算方式 := r_Pay.结算方式;
            Else
              If 结算方式_In Is Null Then
                Begin
                  Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
                Exception
                  When Others Then
                    Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
                End;
              Else
                v_结算方式 := 结算方式_In;
              End If;
            End If;
          
            If r_Pay.性质 = 8 And r_Pay.结算卡序号 Is Not Null Then
              If n_实收金额 >= r_Pay.冲预交 Then
                --Zl_Square_Update(v_原结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, r_Pay.冲预交, r_Pay.结算卡序号);
                Update 病人预交记录
                Set 冲预交 = 冲预交 + (-1 * r_Pay.冲预交), 摘要 = 摘要 || '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * r_Pay.冲预交 || '|'
                Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into 病人预交记录
                    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                     摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                  Values
                    (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_Pay.冲预交, Null, Null,
                     退费时间_In, Null, Null, Null, 操作员编号_In, 操作员姓名_In,
                     '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * r_Pay.冲预交 || '|', n_组id, Null, Null, Null, Null, Null,
                     Null, n_结帐id, -1 * n_结帐id, 3, 1);
                End If;
                n_费用状态 := 1;
                n_实际冲销 := r_Pay.冲预交;
              Else
                --Zl_Square_Update(v_原结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, n_实收金额, r_Pay.结算卡序号);
                Update 病人预交记录
                Set 冲预交 = 冲预交 + (-1 * n_实收金额), 摘要 = 摘要 || '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * n_实收金额 || '|'
                Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into 病人预交记录
                    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                     摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                  Values
                    (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实收金额, Null, Null, 退费时间_In,
                     Null, Null, Null, 操作员编号_In, 操作员姓名_In, '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * n_实收金额 || '|', n_组id,
                     Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
                End If;
                n_费用状态 := 1;
                n_实际冲销 := n_实收金额;
              End If;
            Else
              If r_Pay.冲预交 > n_实收金额 Then
                n_实际冲销 := n_实收金额;
              Else
                n_实际冲销 := r_Pay.冲预交;
              End If;
            End If;
          
            If r_Pay.性质 Not In (3, 4, 7, 8) Then
              Update 病人预交记录
              Set 金额 = 金额 + n_实际冲销
              Where 记录性质 = 1 And 记录状态 = 1 And 收款时间 = 退费时间_In And 病人id + 0 = n_病人id And 结算方式 = v_结算方式;
              If Sql%RowCount = 0 Then
                v_预交no := Nextno(11);
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 预交类别)
                Values
                  (病人预交记录_Id.Nextval, v_预交no, Null, 1, 1, n_病人id, 主页id_In, 入院科室id_In, n_实际冲销, v_结算方式, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '门诊转住院预交', n_组id, r_Pay.预交类别);
              End If;
            
              --病人余额
              Update 病人余额
              Set 预交余额 = Nvl(预交余额, 0) + n_实际冲销
              Where 性质 = 1 And 病人id = n_病人id And 类型 = 2
              Returning 预交余额 Into n_返回值;
              If Sql%RowCount = 0 Then
                Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (n_病人id, 1, 2, n_实际冲销, 0);
                n_返回值 := n_实际冲销;
              End If;
              If Nvl(n_返回值, 0) = 0 Then
                Delete From 病人余额
                Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
              End If;
            End If;
            --4.2缴款数据处理
            --   因为没有实际收病人的钱,所以不处理
            --部分退费情况，退原预交记录
            If r_Pay.性质 In (3, 4) Then
              Update 人员缴款余额
              Set 余额 = Nvl(余额, 0) - n_实际冲销
              Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Pay.结算方式
              Returning 余额 Into n_返回值;
              If Sql%RowCount = 0 Then
                Insert Into 人员缴款余额
                  (收款员, 结算方式, 性质, 余额)
                Values
                  (操作员姓名_In, r_Pay.结算方式, 1, -1 * n_实际冲销);
                n_返回值 := n_实际冲销;
              End If;
              If Nvl(n_返回值, 0) = 0 Then
                Delete From 人员缴款余额
                Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Pay.结算方式 And Nvl(余额, 0) = 0;
              End If;
            End If;
          
            If r_Pay.性质 <> 8 Then
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * n_实际冲销)
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实际冲销, v_结算方式, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, r_Pay.卡类别id, r_Pay.结算卡序号, r_Pay.卡号, r_Pay.交易流水号,
                   r_Pay.交易说明, r_Pay.合作单位, n_结帐id, -1 * n_结帐id, 0, 3);
              End If;
            End If;
          
            Update 病人预交记录
            Set 记录状态 = 3
            Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                  结算方式 = r_Pay.结算方式;
            n_实收金额 := n_实收金额 - n_实际冲销;
          
          End If;
        End If;
      End Loop;
    End If;
  
    If 误差费_In Is Not Null Then
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id,
         卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
      Values
        (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, 误差费_In, v_误差费, Null, 退费时间_In, Null, Null,
         Null, 操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 0, 3);
    End If;
    Delete From 病人预交记录
    Where 结帐id = n_结帐id And 记录性质 = 3 And 记录状态 = 2 And 冲预交 = 0 And 结算方式 Is Not Null;
    Delete From 病人预交记录 Where 结帐id = n_原结帐id And 摘要 = '预交临时记录' And 记录性质 = 3;
    Update 门诊费用记录 Set 费用状态 = Nvl(n_费用状态, 0) Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 = 2;
  Else
    --医保按结算转出
    For r_Nos In (Select Distinct a.No
                  From 门诊费用记录 A
                  Where Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And a.结帐id = 原结帐id_In) Loop
      v_Nos := v_Nos || ',' || r_Nos.No;
    End Loop;
    v_Nos := Substr(v_Nos, 2);
  
    For r_结帐ids In (Select Distinct a.结帐id
                    From 门诊费用记录 A
                    Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.记录性质, 10) = 1 And
                          a.记录状态 <> 0) Loop
      v_结帐ids := v_结帐ids || ',' || r_结帐ids.结帐id;
    End Loop;
    v_结帐ids := Substr(v_结帐ids, 2);
    Select Count(a.No), Sum(a.实收金额)
    Into n_Count, n_实收金额
    From 门诊费用记录 A
    Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.记录性质, 10) = 1;
    If n_Count = 0 Or n_实收金额 = 0 Then
      v_Err_Msg := '本次结算不是收费或因并发原因他人操作了该结算,不能转为住院费用.';
      Raise Err_Item;
    End If;
  
    Select 结帐id, 病人id, 开单部门id, 开单人
    Into n_原结帐id, n_病人id, n_开单部门id, v_开单人
    From 门诊费用记录
    Where 结帐id = 原结帐id_In And Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And Rownum < 2;
  
    Begin
      Select 1
      Into n_部分退费
      From 门诊费用记录 A
      Where Mod(a.记录性质, 10) = 1 And a.记录状态 = 2 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
            Rownum < 2;
    Exception
      When Others Then
        n_部分退费 := 0;
    End;
  
    Begin
      Select 0
      Into n_部分退费
      From 门诊费用记录 A
      Where 记录性质 = 11 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select Count(Avg(1))
      Into n_退费条数
      From 病人预交记录 A
      Where a.记录性质 = 3 And a.记录状态 <> 0 And 结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids)))
      Group By a.结算方式;
    Exception
      When Others Then
        n_退费条数 := 0;
    End;
    --1.1作废费用记录
    If 结帐id_In Is Null Then
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    Else
      n_结帐id := 结帐id_In;
    End If;
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id,
       计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态, 执行时间,
       操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id, 费用状态)
      Select 病人费用记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.序号, a.从属父号, a.价格父号, a.病人id, a.医嘱序号, a.门诊标志, a.姓名, a.性别, a.年龄,
             a.标识号, a.付款方式, a.费别, a.病人科室id, a.收费类别, a.收费细目id, a.计算单位, a.付数, a.发药窗口, -1 * a.数次, a.加班标志, a.附加标志, a.收入项目id,
             a.收据费目, a.记帐费用, a.标准单价, -1 * a.应收金额, -1 * a.实收金额, a.开单部门id, a.开单人, a.执行部门id, a.划价人, a.执行人, -1, a.执行时间,
             操作员编号_In, 操作员姓名_In, a.发生时间, 退费时间_In, n_结帐id, -1 * a.结帐金额, a.保险项目否, a.保险大类id, a.统筹金额, a.摘要,
             Decode(Nvl(a.附加标志, 0), 9, 1, 0), a.保险编码, a.费用类型, n_组id, 0
      From 门诊费用记录 A
      Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.记录性质, 10) = 1 And a.记录状态 = 1;
  
    --作废医保
    For r_医保 In (Select 结帐id, NO, 结算方式, 金额, 备注
                 From 医保结算明细
                 Where NO In (Select Column_Value From Table(f_Str2list(v_Nos))) And
                       结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids)))) Loop
      Update 医保结算明细
      Set 金额 = 金额 + (-1 * r_医保.金额)
      Where NO = r_医保.No And 结帐id = r_医保.结帐id And 结算方式 = r_医保.结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 医保结算明细
          (结帐id, NO, 结算方式, 金额)
        Values
          (r_医保.结帐id, r_医保.No, r_医保.结算方式, -1 * r_医保.金额);
      End If;
    End Loop;
  
    --Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And 记录性质 = 1 And 记录状态 = 1;
    --1.2作废预交记录
    --作废冲预交部分
    If n_部分退费 = 0 And Nvl(门诊退费_In, 0) = 0 Then
      For r_Prepay In (Select NO, 实际票号, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 收款时间, -1 * Sum(冲预交) As 冲预交,
                              卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质
                       From 病人预交记录 A
                       Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
                             Nvl(冲预交, 0) <> 0
                       Group By n_Tempid, NO, 实际票号, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 收款时间, 卡类别id, 结算卡序号,
                                卡号, 交易流水号, 交易说明, 合作单位, 结算性质) Loop
        Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
           结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 预交类别, 结算性质)
          Select n_Tempid, r_Prepay.No, r_Prepay.实际票号, 11, 1, r_Prepay.病人id, r_Prepay.主页id, r_Prepay.科室id, Null,
                 r_Prepay.结算方式, r_Prepay.结算号码, Null, r_Prepay.缴款单位, r_Prepay.单位开户行, r_Prepay.单位帐号, 退费时间_In, 操作员姓名_In,
                 操作员编号_In, r_Prepay.冲预交, n_结帐id, n_组id, r_Prepay.卡类别id, r_Prepay.结算卡序号, r_Prepay.卡号, r_Prepay.交易流水号,
                 r_Prepay.交易说明, r_Prepay.合作单位, -1 * n_结帐id, 1, r_Prepay.结算性质
          From Dual;
      End Loop;
    
      For v_预交 In (Select 病人id, Nvl(预交类别, 2) As 预交类别, Nvl(Sum(Nvl(冲预交, 0)), 0) As 预交金额
                   From 病人预交记录 A
                   Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
                         a.结帐id <> n_结帐id
                   Group By 病人id, Nvl(预交类别, 2)
                   Having Sum(Nvl(冲预交, 0)) <> 0) Loop
      
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) + Nvl(v_预交.预交金额, 0)
        Where 病人id = v_预交.病人id And 类型 = Nvl(v_预交.预交类别, 2) And 性质 = 1
        Returning 预交余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 病人余额
            (病人id, 类型, 预交余额, 性质)
          Values
            (v_预交.病人id, Nvl(v_预交.预交类别, 2), v_预交.预交金额, 1);
          n_返回值 := v_预交.预交金额;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 病人余额
          Where 病人id = v_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
        End If;
      End Loop;
    Else
      If n_退费条数 = 0 And Nvl(门诊退费_In, 0) = 0 Then
        --只使用了预交，原样退回预交
        For r_Prepay In (Select NO, 实际票号, 病人id, 主页id, 科室id, Max(结算方式) As 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 收款时间,
                                -1 * Sum(冲预交) As 冲预交, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质
                         From 病人预交记录 A
                         Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
                               Nvl(冲预交, 0) <> 0
                         Group By n_Tempid, NO, 实际票号, 病人id, 主页id, 科室id, 结算号码, 缴款单位, 单位开户行, 单位帐号, 收款时间, 卡类别id, 结算卡序号, 卡号,
                                  交易流水号, 交易说明, 合作单位, 结算性质) Loop
          Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
             结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 预交类别, 结算性质)
            Select n_Tempid, r_Prepay.No, r_Prepay.实际票号, 11, 1, r_Prepay.病人id, r_Prepay.主页id, r_Prepay.科室id, Null,
                   r_Prepay.结算方式, r_Prepay.结算号码, Null, r_Prepay.缴款单位, r_Prepay.单位开户行, r_Prepay.单位帐号, 退费时间_In, 操作员姓名_In,
                   操作员编号_In, r_Prepay.冲预交, n_结帐id, n_组id, r_Prepay.卡类别id, r_Prepay.结算卡序号, r_Prepay.卡号, r_Prepay.交易流水号,
                   r_Prepay.交易说明, r_Prepay.合作单位, -1 * n_结帐id, 1, r_Prepay.结算性质
            From Dual;
          Select -1 * 冲预交 Into n_预交金额 From 病人预交记录 Where ID = n_Tempid;
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) + Nvl(n_预交金额, 0)
          Where 病人id = r_Prepay.病人id And 类型 = 1 And 性质 = 1
          Returning 预交余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (n_病人id, 1, n_预交金额, 1);
            n_返回值 := n_预交金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 病人余额
            Where 病人id = r_Prepay.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
          End If;
        End Loop;
      Else
        Begin
          Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
        Exception
          When Others Then
            Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
        End;
        Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
           结帐id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
          Select n_Tempid, Max(NO), Max(实际票号), 3, 3, 病人id, 主页id, 科室id, Null, v_结算方式, Max(结算号码), '预交临时记录', Null, Null,
                 Null, Max(收款时间), 操作员姓名_In, 操作员编号_In, Sum(冲预交), n_原结帐id, Null, Null, Null, Null, Null, Null,
                 -1 * n_原结帐id, 3
          From 病人预交记录 A
          Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
                Nvl(冲预交, 0) <> 0
          Group By n_Tempid, 3, 3, 病人id, 主页id, 科室id, Null, v_结算方式, '预交临时记录', 操作员姓名_In, 操作员编号_In, n_原结帐id;
      End If;
    End If;
  
    --作废门诊缴费及医保部分
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
       卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
      Select 病人预交记录_Id.Nextval, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, 退费时间_In, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
             0, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, -1 * n_结帐id, 结算性质
      From 病人预交记录 A, 结算方式 B
      Where a.记录性质 = 3 And a.记录状态 = 1 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
            a.结算方式 = b.名称 And b.性质 Not In (7, 8);
  
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
       卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质, 校对标志)
      Select 病人预交记录_Id.Nextval, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, 退费时间_In, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
             0, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, -1 * n_结帐id, 结算性质, 1
      From 病人预交记录 A, 结算方式 B
      Where a.记录性质 = 3 And a.记录状态 = 1 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
            a.结算方式 = b.名称 And b.性质 = 7;
    If Sql%RowCount <> 0 Then
      n_费用状态 := 1;
    End If;
  
    Update 病人预交记录
    Set 记录状态 = 3
    Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids)));
  
    --2.票据收回
    --可能以前没有打印,无收回
    For r_Nos In (Select Distinct a.No
                  From 门诊费用记录 A
                  Where Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And
                        a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids)))) Loop
    
      Select Nvl(Max(ID), 0)
      Into n_打印id
      From (Select b.Id
             From 票据使用明细 A, 票据打印内容 B
             Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 1 And b.No = r_Nos.No
             Order By a.使用时间 Desc)
      Where Rownum < 2;
      If n_打印id > 0 Then
        --多张单据循环调用时只能收回一次
        Select Count(打印id) Into n_Count From 票据使用明细 Where 票种 = 1 And 性质 = 2 And 打印id = n_打印id;
        If n_Count = 0 Then
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, 退费时间_In, 操作员姓名_In, 票据金额
            From 票据使用明细
            Where 打印id = n_打印id And 票种 = 1 And 性质 = 1;
        End If;
      End If;
    End Loop;
  
    --3.缴款数据处理(
    --   现有两种情况:
    --    1. 转出过程直接销帐的,则缴款数据不增加;
    --    2. 先转出,再到门诊退款退票,则需要进行缴款数据处理
    If Nvl(门诊退费_In, 0) = 1 Then
      For c_预交 In (Select a.结算方式, Sum(a.冲预交) As 冲预交, 2 As 预交类别, a.卡类别id, a.结算卡序号, a.卡号, Min(a.交易流水号) As 交易流水号,
                          Min(a.交易说明) As 交易说明, Min(a.合作单位) As 合作单位, b.性质
                   From 病人预交记录 A, 结算方式 B
                   Where a.记录性质 = 3 And a.记录状态 In (2, 3) And
                         a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And a.结算方式 = b.名称 And
                         b.性质 In (1, 2, 3, 4, 7, 8) And a.结算方式 Is Not Null
                   Group By a.结算方式, 预交类别, a.卡类别id, a.结算卡序号, a.卡号, b.性质
                   Having Sum(a.冲预交) <> 0) Loop
        Begin
          Select 是否退现 Into n_退现 From 医疗卡类别 Where ID = c_预交.卡类别id;
        Exception
          When Others Then
            n_退现 := 0;
        End;
        If (c_预交.性质 = 7 Or (c_预交.性质 = 8 And c_预交.卡类别id Is Not Null)) And n_退现 = 0 Then
          Update 病人预交记录
          Set 冲预交 = 冲预交 + (-1 * c_预交.冲预交), 摘要 = 摘要 || '1' || ',' || c_预交.卡类别id || ',' || -1 * c_预交.冲预交 || '|'
          Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
               缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
            Values
              (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * c_预交.冲预交, Null, Null, 退费时间_In,
               Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || c_预交.卡类别id || ',' || -1 * c_预交.冲预交 || '|', n_组id,
               Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
          End If;
          n_费用状态 := 1;
        Else
          If c_预交.性质 In (3, 4) Or (c_预交.性质 = 8 And c_预交.结算卡序号 Is Not Null) Then
            v_结算方式 := c_预交.结算方式;
          Else
            If 结算方式_In Is Null Then
              Begin
                Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
              Exception
                When Others Then
                  Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
              End;
            Else
              v_结算方式 := 结算方式_In;
            End If;
          End If;
        
          If c_预交.性质 = 8 And c_预交.结算卡序号 Is Not Null Then
            --Zl_Square_Update(v_结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, c_预交.冲预交, c_预交.结算卡序号);
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * c_预交.冲预交), 摘要 = 摘要 || '0' || ',' || c_预交.结算卡序号 || ',' || -1 * c_预交.冲预交 || '|'
            Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
              Values
                (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * c_预交.冲预交, Null, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '0' || ',' || c_预交.结算卡序号 || ',' || -1 * c_预交.冲预交 || '|', n_组id,
                 Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
            End If;
            n_费用状态 := 1;
          End If;
          If c_预交.结算卡序号 Is Null Then
            Update 人员缴款余额
            Set 余额 = Nvl(余额, 0) - c_预交.冲预交
            Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
            Returning 余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 人员缴款余额
                (收款员, 结算方式, 性质, 余额)
              Values
                (操作员姓名_In, v_结算方式, 1, -1 * c_预交.冲预交);
              n_返回值 := c_预交.冲预交;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 人员缴款余额
              Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式 And Nvl(余额, 0) = 0;
            End If;
            --部分退费情况，退原预交记录
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * c_预交.冲预交)
            Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
              Values
                (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * c_预交.冲预交, v_结算方式, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, c_预交.合作单位, n_结帐id,
                 -1 * n_结帐id, 0, 3);
            End If;
          End If;
        End If;
      End Loop;
    
      --更新费用审核记录
      Update 费用审核记录
      Set 记录状态 = 2
      Where 费用id In (Select a.Id
                     From 门诊费用记录 A
                     Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.记录性质, 10) = 1 And
                           a.记录状态 In (1, 3)) And 性质 = 1;
      --作废门诊记录
      For r_Nos In (Select Distinct NO
                    From 门诊费用记录
                    Where Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And
                          结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids)))) Loop
        Update 门诊费用记录 Set 记录状态 = 3 Where NO = r_Nos.No And Mod(记录性质, 10) = 1 And 记录状态 = 1;
      End Loop;
      For r_Clinic In (Select Min(a.记录性质) As 记录性质, a.No, a.序号, a.从属父号, a.价格父号, a.病人id, a.姓名, a.性别, a.年龄, a.病人科室id, a.费别,
                              a.收费类别, a.收费细目id, a.计算单位, a.保险项目否, a.保险大类id, a.保险编码, a.费用类型, a.发药窗口, a.付数, Sum(a.数次) As 数次,
                              a.加班标志, a.附加标志, a.收入项目id, a.收据费目, a.标准单价, Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额,
                              Sum(a.统筹金额) As 统筹金额, a.开单部门id, a.开单人, a.执行部门id, a.划价人, Max(a.记帐单id) As 记帐单id,
                              Max(a.是否急诊) As 是否急诊, a.发生时间, Min(a.实际票号) As 实际票号
                       From 门诊费用记录 A
                       Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.记录性质, 10) = 1 And
                             a.记录状态 In (2, 3) And Nvl(a.附加标志, 0) Not In (8, 9)
                       Group By a.No, a.序号, a.从属父号, a.价格父号, a.病人id, a.姓名, a.性别, a.年龄, a.病人科室id, a.费别, a.收费类别, a.收费细目id,
                                a.计算单位, a.保险项目否, a.保险大类id, a.保险编码, a.费用类型, a.发药窗口, a.付数, a.加班标志, a.附加标志, a.收入项目id, a.收据费目,
                                a.标准单价, a.开单部门id, a.开单人, a.执行部门id, a.划价人, a.发生时间
                       Having Sum(a.数次) <> 0) Loop
        Insert Into 门诊费用记录
          (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否,
           保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id, 开单人, 发生时间,
           登记时间, 执行部门id, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 是否急诊, 缴款组id, 结帐id, 结帐金额, 执行状态, 费用状态)
        Values
          (病人费用记录_Id.Nextval, r_Clinic.记录性质, r_Clinic.No, r_Clinic.实际票号, 2, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号,
           1, r_Clinic.病人id, '', r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, r_Clinic.病人科室id, r_Clinic.费别, r_Clinic.收费类别,
           r_Clinic.收费细目id, r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id, r_Clinic.保险编码, r_Clinic.费用类型, r_Clinic.发药窗口,
           r_Clinic.付数, -1 * r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志, r_Clinic.收入项目id, r_Clinic.收据费目, r_Clinic.标准单价,
           -1 * r_Clinic.应收金额, -1 * r_Clinic.实收金额, -1 * r_Clinic.统筹金额, 0, r_Clinic.开单部门id, r_Clinic.开单人, r_Clinic.发生时间,
           退费时间_In, r_Clinic.执行部门id, r_Clinic.划价人, 操作员编号_In, 操作员姓名_In, r_Clinic.记帐单id, '', r_Clinic.是否急诊, n_组id, n_结帐id,
           -1 * r_Clinic.实收金额, -1, 0);
      End Loop;
    Else
      --4.退款转预交(不产生票据,由操作员通过重打进行)
    
      For r_Pay In (Select Min(a.Id) As 预交id, a.结算方式, Sum(a.冲预交) As 冲预交, 2 As 预交类别, a.卡类别id, a.结算卡序号, a.卡号, a.交易流水号,
                           a.交易说明, a.合作单位, b.性质
                    From 病人预交记录 A, 结算方式 B
                    Where a.记录性质 = 3 And a.记录状态 In (2, 3) And
                          a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And a.结算方式 = b.名称 And
                          b.性质 In (1, 2, 3, 4, 7, 8) And a.结算方式 Is Not Null
                    Group By a.结算方式, 预交类别, a.卡类别id, a.结算卡序号, a.卡号, b.性质, a.交易流水号, a.交易说明, a.合作单位

                    
                    Having Sum(a.冲预交) <> 0) Loop
        --4.1产生预交款单据 (不存在部分退费的情况)
        --所有单据,按规则生成预交款单据
        --因为收款后立即缴款,所以人员缴款余额无变化
        If r_Pay.性质 = 7 Or (r_Pay.性质 = 8 And r_Pay.卡类别id Is Not Null) Then
          Update 病人预交记录
          Set 冲预交 = 冲预交 + (-1 * r_Pay.冲预交), 摘要 = 摘要 || '1' || ',' || r_Pay.卡类别id || ',' || -1 * r_Pay.冲预交 || '|'
          Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
               缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
            Values
              (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_Pay.冲预交, Null, Null, 退费时间_In,
               Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || r_Pay.卡类别id || ',' || -1 * r_Pay.冲预交 || '|', n_组id,
               Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
          End If;
          n_费用状态 := 1;
        Else
          If r_Pay.性质 In (3, 4) Or (r_Pay.性质 = 8 And r_Pay.结算卡序号 Is Not Null) Then
            v_结算方式 := r_Pay.结算方式;
          Else
            Begin
              Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
            Exception
              When Others Then
                Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
            End;
          End If;
        
          If r_Pay.性质 = 8 Then
            --Zl_Square_Update(v_结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, r_Pay.冲预交, r_Pay.结算卡序号);
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * r_Pay.冲预交), 摘要 = 摘要 || '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * r_Pay.冲预交 || '|'
            Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
              Values
                (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_Pay.冲预交, Null, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * r_Pay.冲预交 || '|', n_组id,
                 Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
            End If;
            n_费用状态 := 1;
          End If;
          If r_Pay.性质 Not In (3, 4, 7, 8) Then
            Update 病人预交记录
            Set 金额 = 金额 + r_Pay.冲预交
            Where 记录性质 = 1 And 记录状态 = 1 And 收款时间 = 退费时间_In And 病人id + 0 = n_病人id And 结算方式 = v_结算方式;
            If Sql%RowCount = 0 Then
              v_预交no := Nextno(11);
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 预交类别)
              Values
                (病人预交记录_Id.Nextval, v_预交no, Null, 1, 1, n_病人id, 主页id_In, 入院科室id_In, r_Pay.冲预交, v_结算方式, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '门诊转住院预交', n_组id, r_Pay.预交类别);
            End If;
          
            --病人余额
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + r_Pay.冲预交
            Where 性质 = 1 And 病人id = n_病人id And 类型 = 2
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (n_病人id, 1, 2, r_Pay.冲预交, 0);
              n_返回值 := r_Pay.冲预交;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
          End If;
          --4.2缴款数据处理
          --   因为没有实际收病人的钱,所以不处理
          --部分退费情况，退原预交记录
          If r_Pay.性质 In (3, 4) Then
            Update 人员缴款余额
            Set 余额 = Nvl(余额, 0) - r_Pay.冲预交
            Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Pay.结算方式
            Returning 余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 人员缴款余额
                (收款员, 结算方式, 性质, 余额)
              Values
                (操作员姓名_In, r_Pay.结算方式, 1, -1 * r_Pay.冲预交);
              n_返回值 := r_Pay.冲预交;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 人员缴款余额
              Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Pay.结算方式 And Nvl(余额, 0) = 0;
            End If;
          End If;
        
          If r_Pay.结算卡序号 Is Null Then
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * r_Pay.冲预交)
            Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
              Values
                (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_Pay.冲预交, v_结算方式, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, r_Pay.卡类别id, r_Pay.结算卡序号, r_Pay.卡号, r_Pay.交易流水号,
                 r_Pay.交易说明, r_Pay.合作单位, n_结帐id, -1 * n_结帐id, 0, 3);
            End If;
          End If;
        End If;
      End Loop;
    End If;
    If 误差费_In Is Not Null Then
      Begin
        Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
      Exception
        When Others Then
          Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
      End;
      Update 病人预交记录
      Set 冲预交 = 冲预交 - 误差费_In
      Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
      Update 病人预交记录
      Set 冲预交 = 冲预交 + 误差费_In
      Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_误差费;
      If Sql%RowCount = 0 Then
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
           缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
        Values
          (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, 误差费_In, v_误差费, Null, 退费时间_In, Null, Null,
           Null, 操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 0, 3);
      End If;
    End If;
    Delete From 病人预交记录 Where 结帐id = n_原结帐id And 摘要 = '预交临时记录' And 记录性质 = 3;
    Delete From 病人预交记录
    Where 结帐id = n_结帐id And 记录性质 = 3 And 记录状态 = 2 And 冲预交 = 0 And 结算方式 Is Not Null;
    Update 门诊费用记录
    Set 费用状态 = Nvl(n_费用状态, 0)
    Where NO In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(记录性质, 10) = 1 And 记录状态 = 2;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊转住院_收费转出;
/

--139063:冉俊明,2019-04-08,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_门诊转住院_记帐转出
(
  No_In         住院费用记录.No%Type,
  操作员编号_In 住院费用记录.操作员编号%Type,
  操作员姓名_In 住院费用记录.操作员姓名%Type,
  退费时间_In   住院费用记录.发生时间%Type,
  门诊销帐_In   Number := 0
) As
  --门诊销帐_In:0-门诊转住院立即销帐;1-门诊记帐退费模式
  n_Count      Number(5);
  n_实收金额   住院费用记录.实收金额%Type;
  n_病人id     住院费用记录.病人id%Type;
  n_开单部门id 住院费用记录.开单部门id%Type;
  v_开单人     门诊费用记录.开单人%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(200);
Begin

  Select Count(NO), Sum(实收金额) Into n_Count, n_实收金额 From 门诊费用记录 Where NO = No_In And 记录性质 = 2;
  If n_Count = 0 Then
    v_Err_Msg := '单据' || No_In || '不是记帐单据或因并发原因他人操作了该单据,不能转为住院费用.';
    Raise Err_Item;
  End If;

  Select 病人id, 开单部门id, 开单人
  Into n_病人id, n_开单部门id, v_开单人
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 2 And 记录状态 In (1, 3) And Rownum = 1;

  --处理病人余额
  Begin
    Select Nvl(Sum(实收金额), 0)
    Into n_实收金额
    From 门诊费用记录
    Where NO = No_In And 记录性质 = 2 And 记录状态 In (1, 2, 3) And Nvl(门诊标志, 0) <> 4 And 结帐id Is Null
    Group By NO, 记录性质;
  Exception
    When Others Then
      n_实收金额 := 0;
  End;

  Update 病人余额 Set 费用余额 = Nvl(费用余额, 0) - n_实收金额 Where 病人id = n_病人id And 类型 = 1 And 性质 = 1;
  If Sql%RowCount = 0 And n_实收金额 <> 0 Then
    Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (n_病人id, 1, 1, -1 * n_实收金额, 0);
  End If;

  --处理未结费用
  For v_未结 In (Select 开单部门id, 病人id, 病人科室id, 执行部门id, 收入项目id, 门诊标志, -1 * Nvl(Sum(实收金额), 0) As 实收金额, 主页id, 病人病区id
               From 门诊费用记录
               Where NO = No_In And 记录性质 = 2 And 记录状态 In (1, 2, 3)
               Group By 开单部门id, 病人id, 病人科室id, 执行部门id, 收入项目id, 门诊标志, 主页id, 病人病区id) Loop
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + v_未结.实收金额
    Where 病人id = v_未结.病人id And Nvl(主页id, 0) = Nvl(v_未结.主页id, 0) And Nvl(病人病区id, 0) = Nvl(v_未结.病人病区id, 0) And
          Nvl(病人科室id, 0) = Nvl(v_未结.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(v_未结.开单部门id, 0) And
          Nvl(执行部门id, 0) = Nvl(v_未结.执行部门id, 0) And 收入项目id + 0 = v_未结.收入项目id And 来源途径 + 0 = v_未结.门诊标志;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (v_未结.病人id, v_未结.主页id, v_未结.病人病区id, v_未结.病人科室id, v_未结.开单部门id, v_未结.执行部门id, v_未结.收入项目id, v_未结.门诊标志, v_未结.实收金额);
    End If;
  End Loop;

  --作废费用记录
  Insert Into 门诊费用记录
    (ID, NO, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 婴儿费, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id, 计算单位,
     付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态, 执行时间, 操作员编号,
     操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 记帐单id, 摘要, 保险编码, 是否急诊, 结论, 主页id, 病人病区id)
    Select 病人费用记录_Id.Nextval, NO, 记录性质, 2, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 婴儿费, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id,
           收费类别, 收费细目id, 计算单位, 付数, 发药窗口, -1 * 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, -1 * 应收金额, -1 * 实收金额, 开单部门id,
           开单人, 执行部门id, 划价人, 执行人, -1, 执行时间, 操作员编号_In, 操作员姓名_In, 发生时间, 退费时间_In, 保险项目否, 保险大类id, -1 * 统筹金额, 记帐单id, 摘要, 保险编码,
           是否急诊, 结论, 主页id, 病人病区id
    From 门诊费用记录
    Where NO = No_In And 记录性质 = 2 And 记录状态 = 1;

  --Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And 记录性质 = 2 And 记录状态 = 1;

  --药品处理(未处理,主要是因为直接转换成相关的药房即可.)
  If Nvl(门诊销帐_In, 0) = 1 Then
    Update 费用审核记录
    Set 记录状态 = 2
    Where 费用id In (Select ID From 门诊费用记录 Where NO = No_In And 记录性质 = 2 And 记录状态 In (1, 3)) And 性质 = 1;
    --作废门诊记录
    Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And 记录性质 = 2 And 记录状态 = 1;
    For r_Clinic In (Select 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型,
                            发药窗口, 付数, Sum(数次) As 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, Sum(应收金额) As 应收金额,
                            Sum(实收金额) As 实收金额, Sum(统筹金额) As 统筹金额, 开单部门id, 开单人, 执行部门id, 划价人, 记帐单id, 是否急诊, 缴款组id, 发生时间,
                            实际票号, 主页id, 病人病区id
                     From 门诊费用记录
                     Where NO = No_In And 记录性质 = 2 And 记录状态 In (2, 3) And 附加标志 Not In (8, 9)
                     Group By 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型,
                              发药窗口, 付数, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 开单部门id, 开单人, 执行部门id, 划价人, 记帐单id, 是否急诊, 缴款组id,
                              发生时间, 实际票号, 主页id, 病人病区id
                     Having Sum(数次) <> 0) Loop
      Insert Into 门诊费用记录
        (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否,
         保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id, 开单人,
         发生时间, 登记时间, 执行部门id, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 是否急诊, 缴款组id, 主页id, 病人病区id)
      Values
        (病人费用记录_Id.Nextval, 2, No_In, r_Clinic.实际票号, 2, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号, 1, r_Clinic.病人id, '',
         r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, r_Clinic.病人科室id, r_Clinic.费别, r_Clinic.收费类别, r_Clinic.收费细目id,
         r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id, r_Clinic.保险编码, r_Clinic.费用类型, r_Clinic.发药窗口, r_Clinic.付数,
         -1 * r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志, r_Clinic.婴儿费, r_Clinic.收入项目id, r_Clinic.收据费目, r_Clinic.标准单价,
         -1 * r_Clinic.应收金额, -1 * r_Clinic.实收金额, -1 * r_Clinic.统筹金额, 1, r_Clinic.开单部门id, r_Clinic.开单人, r_Clinic.发生时间,
         退费时间_In, r_Clinic.执行部门id, r_Clinic.划价人, 操作员编号_In, 操作员姓名_In, r_Clinic.记帐单id, '', r_Clinic.是否急诊, r_Clinic.缴款组id,
         r_Clinic.主页id, r_Clinic.病人病区id);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊转住院_记帐转出;
/

--139063:冉俊明,2019-04-08,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_病人结帐记录_Cancel
(
  No_In         病人结帐记录.No%Type,
  冲销id_In     病人结帐记录.Id%Type,
  操作员编号_In 病人结帐记录.操作员编号%Type,
  操作员姓名_In 病人结帐记录.操作员姓名%Type,
  作废时间_In   病人结帐记录.收费时间%Type := Null,
  票据号_In     病人结帐记录.实际票号%Type := Null,
  领用id_In     票据领用记录.Id%Type := Null,
  票种_In       票据使用明细.票种%Type := Null
) As
  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  --该游标用于预交记录相关信息

  --该游标用于处理费用相关汇总表
  Cursor c_Money(v_Id 病人预交记录.结帐id%Type) Is
    Select NO, 开单部门id, 病人科室id, 执行部门id, 病人病区id, 病人id, 主页id, 收入项目id, 门诊标志, 结帐金额
    From 住院费用记录
    Where 结帐id = v_Id
    Union All
    Select NO, 开单部门id, 病人科室id, 执行部门id, 0 As 病人病区id, 病人id, 0 As 主页id, 收入项目id, 门诊标志, 结帐金额
    From 门诊费用记录
    Where 结帐id = v_Id;

  r_Moneyrow c_Money%RowType;

  --该游标包含病人的相关信息
  Cursor c_Pati(n_病人id 病人信息.病人id%Type) Is
    Select a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, b.主页id, b.出院病床, b.当前病区id, b.出院科室id, Nvl(b.费别, a.费别) As 费别, a.险类, c.编码 As 付款方式
    From 病人信息 A, 病案主页 B, 医疗付款方式 C
    Where a.病人id = n_病人id And a.病人id = b.病人id(+) And Nvl(a.主页id, 0) = b.主页id(+) And a.医疗付款方式 = c.名称(+);
  r_Pati c_Pati%RowType;

  --过程变量
  v_实际票号 病人预交记录.实际票号%Type;
  n_预交id   病人预交记录.Id%Type;
  n_病人id   病人信息.病人id%Type;

  n_原id    病人结帐记录.Id%Type;
  n_结帐id  病人结帐记录.Id%Type;
  v_打印ids Varchar2(5000);
  v_打印id  票据打印内容.Id%Type;

  n_来源     Number; --1-门诊;2-住院;3-门诊和住院
  n_返回值   病人余额.预交余额%Type;
  n_组id     财务缴款分组.Id%Type;
  n_预交类别 Number;
  d_Date     Date;

Begin
  n_组id := Zl_Get组id(操作员姓名_In);
  Begin
    Select ID, 病人id, 实际票号 Into n_原id, n_病人id, v_实际票号 From 病人结帐记录 Where 记录状态 = 1 And NO = No_In;
    --打印的内容
    Begin
      Select ID
      Into v_打印ids
      From (Select b.Id
             From 票据使用明细 A, 票据打印内容 B
             Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 3 And b.No = No_In
             Order By a.使用时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  Exception
    When Others Then
      Begin
        v_Err_Msg := '没有发现要作废的结帐单据,可能已经作废！';
        Raise Err_Item;
      End;
  End;

  If 票据号_In Is Not Null Then
    Select 票据打印内容_Id.Nextval Into v_打印id From Dual;
  
    --发出票据
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (v_打印id, 3, No_In);
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
    Values
      (票据使用明细_Id.Nextval, 票种_In, 票据号_In, 1, 6, 领用id_In, v_打印id, 作废时间_In, 操作员姓名_In);
  
    --状态改动
    Update 票据领用记录
    Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
    Where ID = Nvl(领用id_In, 0);
  End If;

  Open c_Pati(n_病人id);
  Fetch c_Pati
    Into r_Pati; --体检系统调用此过程,团体结帐时没有病人信息
  d_Date := 作废时间_In;
  If d_Date Is Null Then
    Select Sysdate Into d_Date From Dual;
  End If;
  n_结帐id := 冲销id_In;
  If Nvl(n_结帐id, 0) = 0 Then
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  End If;

  --病人结帐记录
  Insert Into 病人结帐记录
    (ID, NO, 实际票号, 记录状态, 中途结帐, 病人id, 操作员编号, 操作员姓名, 开始日期, 结束日期, 收费时间, 备注, 原因, 缴款组id, 结帐类型, 结算状态, 主页id, 住院次数, 结帐金额)
    Select n_结帐id, NO, 实际票号, 2, 中途结帐, 病人id, 操作员编号_In, 操作员姓名_In, 开始日期, 结束日期, d_Date, 备注, 原因, n_组id, 结帐类型, 1, 主页id, 住院次数,
           -1 * 结帐金额
    From 病人结帐记录
    Where ID = n_原id;

  Update 病人结帐记录 Set 记录状态 = 3 Where ID = n_原id;

  --作废收回票据(可能以前没有使用票据,无法收回)
  If v_打印ids Is Not Null Then
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
      Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In, 票据金额
      From 票据使用明细
      Where 打印id In (Select Column_Value From Table(f_Str2list(v_打印ids))) And 票种 In (1, 3) And 性质 = 1;
  End If;

  Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
  --插入结算方式为NULL的结算方式
  Insert Into 病人预交记录
    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交, 结帐id,
     缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质, 校对标志)
    Select n_预交id, No_In, v_实际票号, 12, 1, n_病人id, Max(Decode(Mod(记录性质, 10), 1, Null, 主页id)),
           Max(Decode(Mod(记录性质, 10), 1, Null, 科室id)), Null, Null, Null, Null, Null, Null, Null, d_Date, 操作员姓名_In,
           操作员编号_In, -1 * Sum(冲预交), n_结帐id, n_组id, Null, Null, Null, Null As 卡号, Null As 交易流水号, Null As 交易说明,
           Null As 合作单位, 2, 1
    From 病人预交记录
    Where 结帐id = n_原id;

  --确定结帐的费用记录来源
  Begin
    Select Case
             When Nvl(Max(住院), 0) = 1 And Nvl(Max(门诊), 0) = 1 Then
              3
             When Nvl(Max(住院), 0) = 1 Then
              2
             Else
              1
           End
    Into n_来源
    From (Select 1 As 住院, 0 As 门诊
           From 住院费用记录
           Where 结帐id = n_原id And Rownum = 1
           Union All
           Select 0 As 住院, 1 As 门诊
           From 门诊费用记录
           Where 结帐id = n_原id And Rownum = 1);
  
  Exception
    When Others Then
      n_来源 := 3;
  End;

  If n_来源 = 2 Or n_来源 = 3 Then
    --作废结帐对应的费用记录:不包含原始结帐产生的误差项目
    Insert Into 住院费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id,
       病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id,
       开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额, 结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要,
       缴款组id, 医疗小组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, To_Number('1' || Substr(记录性质, Length(记录性质), 1)), 记录状态, 序号, 从属父号, 价格父号, 多病人单,
             记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次,
             加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, Null, Null, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人,
             执行时间, 操作员姓名, 操作员编号, -1 * 结帐金额, n_结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要, 缴款组id, 医疗小组id
      From 住院费用记录
      Where 结帐id = n_原id;
  End If;

  If n_来源 = 1 Or n_来源 = 3 Then
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别,
       收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间,
       执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额, 结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要, 缴款组id, 主页id, 病人病区id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, To_Number('1' || Substr(记录性质, Length(记录性质), 1)), 记录状态, 序号, 从属父号, 价格父号, 记帐单id,
             病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费,
             记帐费用, 收入项目id, 收据费目, 标准单价, Null, Null, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号,
             -1 * 结帐金额, n_结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要, 缴款组id, 主页id, 病人病区id
      From 门诊费用记录
      Where 结帐id = n_原id;
  End If;

  For r_Moneyrow In c_Money(n_结帐id) Loop
    --病人余额 ,所以不需要更新这两个汇总表
  
    If Nvl(r_Moneyrow.门诊标志, 0) = 1 Or Nvl(r_Moneyrow.门诊标志, 0) = 2 Then
      n_预交类别 := r_Moneyrow.门诊标志;
    Elsif Nvl(r_Moneyrow.主页id, 0) = 0 Or Nvl(r_Moneyrow.门诊标志, 0) = 4 Then
      --体检:门诊病人
      n_预交类别 := 1;
    Else
      n_预交类别 := 2;
    End If;
  
    If Nvl(r_Moneyrow.门诊标志, 0) <> 4 Then
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) - r_Moneyrow.结帐金额 --注:新的结帐ID产生的是负数金额
      Where 病人id = r_Moneyrow.病人id And 类型 = n_预交类别 And 性质 = 1
      Returning 费用余额 Into n_返回值;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 预交余额, 费用余额)
        Values
          (r_Moneyrow.病人id, 1, n_预交类别, 0, -1 * r_Moneyrow.结帐金额);
        n_返回值 := -1 * r_Moneyrow.结帐金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete 病人余额
        Where 病人id = r_Moneyrow.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
      End If;
    End If;
  
    --病人未结费用
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) - r_Moneyrow.结帐金额
    Where 病人id = r_Moneyrow.病人id And Nvl(主页id, 0) = Nvl(r_Moneyrow.主页id, 0) And
          Nvl(病人病区id, 0) = Nvl(r_Moneyrow.病人病区id, 0) And Nvl(病人科室id, 0) = Nvl(r_Moneyrow.病人科室id, 0) And
          Nvl(开单部门id, 0) = Nvl(r_Moneyrow.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(r_Moneyrow.执行部门id, 0) And
          收入项目id + 0 = r_Moneyrow.收入项目id And 来源途径 + 0 = r_Moneyrow.门诊标志;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (r_Moneyrow.病人id, Decode(r_Moneyrow.主页id, Null, Null, 0, Null, r_Moneyrow.主页id),
         Decode(r_Moneyrow.病人病区id, Null, Null, 0, Null, r_Moneyrow.病人病区id), r_Moneyrow.病人科室id, r_Moneyrow.开单部门id,
         r_Moneyrow.执行部门id, r_Moneyrow.收入项目id, r_Moneyrow.门诊标志, -1 * r_Moneyrow.结帐金额);
    End If;
  
  End Loop;
  Close c_Pati;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人结帐记录_Cancel;
/

--139063:冉俊明,2019-04-08,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_病人结帐记录_Delete
(
  No_In           病人结帐记录.No%Type,
  操作员编号_In   病人结帐记录.操作员编号%Type,
  操作员姓名_In   病人结帐记录.操作员姓名%Type,
  误差金额_In     病人预交记录.冲预交%Type := 0, --医保或预交退现金产生的误差
  结帐作废结算_In Varchar2 := Null, --结算方式|结算金额|结算号码||......
  预交退现金_In   Number := 0, --当预交款退现金时，结算方式及金额通过参数结帐作废结算_In传入
  冲销id_In       病人预交记录.结帐id%Type := Null,
  冲销时间_In     Date := Null,
  缴预交id_In     病人预交记录.Id%Type := Null, --在作废时将相关的金额充值到预交款时填写
  票据号_In       病人结帐记录.实际票号%Type := Null,
  领用id_In       票据领用记录.Id%Type := Null,
  票种_In         票据使用明细.票种%Type := Null
) As
  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  --该游标用于预交记录相关信息
  Cursor c_Deposit(v_Id 病人预交记录.结帐id%Type) Is
    Select 病人id, 记录性质, 结算方式, 冲预交, 预交类别 From 病人预交记录 Where 结帐id = v_Id;
  r_Depositrow c_Deposit%RowType;

  --该游标用于处理费用相关汇总表
  Cursor c_Money(v_Id 病人预交记录.结帐id%Type) Is
    Select NO, 开单部门id, 病人科室id, 执行部门id, 病人病区id, 病人id, 主页id, 收入项目id, 门诊标志, 结帐金额
    From 住院费用记录
    Where 结帐id = v_Id
    Union All
    Select NO, 开单部门id, 病人科室id, 执行部门id, 0 As 病人病区id, 病人id, 0 As 主页id, 收入项目id, 门诊标志, 结帐金额
    From 门诊费用记录
    Where 结帐id = v_Id;

  r_Moneyrow c_Money%RowType;

  --该游标包含病人的相关信息
  Cursor c_Pati(n_病人id 病人信息.病人id%Type) Is
    Select a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, b.主页id, b.出院病床, b.当前病区id, b.出院科室id, Nvl(b.费别, a.费别) As 费别, a.险类, c.编码 As 付款方式
    From 病人信息 A, 病案主页 B, 医疗付款方式 C
    Where a.病人id = n_病人id And a.病人id = b.病人id(+) And Nvl(a.主页id, 0) = b.主页id(+) And a.医疗付款方式 = c.名称(+);
  r_Pati c_Pati%RowType;

  --过程变量
  v_结算内容 Varchar2(500);
  v_当前结算 Varchar2(50);
  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  v_结算号码 病人预交记录.结算号码%Type;
  v_实际票号 病人预交记录.实际票号%Type;
  v_误差no   住院费用记录.No%Type;
  v_误差     结算方式.名称%Type;
  n_病人id   病人信息.病人id%Type;

  n_原id   病人结帐记录.Id%Type;
  n_结帐id 病人结帐记录.Id%Type;
  n_打印id 票据打印内容.Id%Type;

  n_来源     Number; --1-门诊;2-住院;3-门诊和住院
  n_返回值   病人余额.预交余额%Type;
  n_组id     财务缴款分组.Id%Type;
  n_预交类别 Number;
  d_Date     Date;
  v_打印id   票据打印内容.Id%Type;
Begin
  n_组id := Zl_Get组id(操作员姓名_In);

  Select 名称 Into v_误差 From 结算方式 Where 性质 = 9 And Rownum = 1;

  Begin
    Select ID, 病人id, 实际票号 Into n_原id, n_病人id, v_实际票号 From 病人结帐记录 Where 记录状态 = 1 And NO = No_In;
    --最后一次打印的内容
    Select Max(ID)
    Into n_打印id
    From (Select b.Id
           From 票据使用明细 A, 票据打印内容 B
           Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 3 And b.No = No_In
           Order By a.使用时间 Desc)
    Where Rownum < 2;
  Exception
    When Others Then
      Begin
        v_Err_Msg := '没有发现要作废的结帐单据,可能已经作废！';
        Raise Err_Item;
      End;
  End;

  Open c_Pati(n_病人id);
  Fetch c_Pati
    Into r_Pati; --体检系统调用此过程,团体结帐时没有病人信息

  d_Date := 冲销时间_In;
  If d_Date Is Null Then
    Select Sysdate Into d_Date From Dual;
  End If;
  n_结帐id := 冲销id_In;
  If Nvl(n_结帐id, 0) = 0 Then
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  End If;

  If 票据号_In Is Not Null Then
    Select 票据打印内容_Id.Nextval Into v_打印id From Dual;
  
    --发出票据
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (v_打印id, 3, No_In);
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
    Values
      (票据使用明细_Id.Nextval, 票种_In, 票据号_In, 1, 6, 领用id_In, v_打印id, d_Date, 操作员姓名_In);
  
    --状态改动
    Update 票据领用记录
    Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
    Where ID = Nvl(领用id_In, 0);
  End If;

  --病人结帐记录
  Insert Into 病人结帐记录
    (ID, NO, 实际票号, 记录状态, 中途结帐, 病人id, 操作员编号, 操作员姓名, 开始日期, 结束日期, 收费时间, 备注, 原因, 缴款组id, 结帐类型)
    Select n_结帐id, NO, 实际票号, 2, 中途结帐, 病人id, 操作员编号_In, 操作员姓名_In, 开始日期, 结束日期, d_Date, 备注, 原因, n_组id, 结帐类型
    From 病人结帐记录
    Where ID = n_原id;

  Update 病人结帐记录 Set 记录状态 = 3 Where ID = n_原id;

  --作废收回票据(可能以前没有使用票据,无法收回)
  If n_打印id Is Not Null Then
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
      Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In, 票据金额
      From 票据使用明细
      Where 打印id = n_打印id And 票种 In (1, 3) And 性质 = 1;
  End If;

  --病人预交记录(冲预交及缴款)
  If 结帐作废结算_In Is Null Then
    For c_预交 In (Select 病人预交记录_Id.Nextval As 预交id, ID, NO, 实际票号, To_Number('1' || Substr(记录性质, Length(记录性质), 1)) As 记录性质,
                        记录状态, 病人id, 主页id, 科室id, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 冲预交, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号,
                        交易说明, 合作单位
                 From 病人预交记录
                 Where 结帐id = n_原id And (记录性质 In (1, 11) And Nvl(冲预交, 0) <> 0 Or 记录性质 Not In (1, 11))) Loop
    
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
      Values
        (c_预交.预交id, c_预交.No, c_预交.实际票号, c_预交.记录性质, c_预交.记录状态, c_预交.病人id, c_预交.主页id, c_预交.科室id, Null, c_预交.结算方式,
         c_预交.结算号码, c_预交.摘要, c_预交.缴款单位, c_预交.单位开户行, c_预交.单位帐号, d_Date, 操作员姓名_In, 操作员编号_In, -1 * c_预交.冲预交, n_结帐id, n_组id,
         c_预交.预交类别, c_预交.卡类别id, c_预交.结算卡序号, c_预交.卡号, c_预交.交易流水号, c_预交.交易说明, c_预交.合作单位, 2);
    
      --消费卡处理
      For c_记录 In (Select c.接口编号, c.消费卡id, c.卡号, -1 * Sum(c.应收金额) As 结算金额
                   From 病人卡结算记录 C
                   Where c.结算id = c_预交.Id And c.记录状态 = 1
                   Group By c.接口编号, c.消费卡id, c.卡号) Loop
        Zl_病人卡结算记录_退款(c_记录.接口编号, c_记录.卡号, c_记录.消费卡id, c_记录.结算金额, c_预交.Id, c_预交.预交id, 操作员编号_In, 操作员姓名_In, d_Date);
      End Loop;
    End Loop;
  Else
    --1.先处理冲预交部分
    If 预交退现金_In = 0 Then
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, To_Number('1' || Substr(记录性质, Length(记录性质), 1)), 记录状态, 病人id, 主页id, 科室id,
               Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_Date, 操作员姓名_In, 操作员编号_In, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 2
        From 病人预交记录
        Where 结帐id = n_原id And 记录性质 In (1, 11) And Nvl(冲预交, 0) <> 0;
    End If;
  
    --2.再处理结帐结算,包括医保和非医保
    v_结算内容 := 结帐作废结算_In || ' ||'; --以空格分开以|结尾,没有结算号码的
    While v_结算内容 Is Not Null Loop
      v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '||') - 1);
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
      n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1));
      v_结算号码 := LTrim(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 结算性质)
      Values
        (病人预交记录_Id.Nextval, No_In, v_实际票号, 12, 1, n_病人id, r_Pati.主页id, r_Pati.出院科室id, Null, v_结算方式, v_结算号码, '结帐作废退款',
         Null, Null, Null, d_Date, 操作员姓名_In, 操作员编号_In, -1 * n_结算金额, n_结帐id, n_组id, 2);
    
      v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '||') + 2);
    End Loop;
  End If;
  --确定结帐的费用记录来源
  Begin
    Select Case
             When Nvl(Max(住院), 0) = 1 And Nvl(Max(门诊), 0) = 1 Then
              3
             When Nvl(Max(住院), 0) = 1 Then
              2
             Else
              1
           End
    Into n_来源
    From (Select 1 As 住院, 0 As 门诊
           From 住院费用记录
           Where 结帐id = n_原id And Rownum = 1
           Union All
           Select 0 As 住院, 1 As 门诊
           From 门诊费用记录
           Where 结帐id = n_原id And Rownum = 1);
  
  Exception
    When Others Then
      n_来源 := 3;
  End;

  If 误差金额_In <> 0 Then
    Update 病人预交记录
    Set 冲预交 = 冲预交 + 误差金额_In
    Where NO = No_In And 记录性质 = 12 And 记录状态 = 1 And 结帐id = n_结帐id;
    If Sql%RowCount = 0 Then
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 结算性质)
      Values
        (病人预交记录_Id.Nextval, No_In, v_实际票号, 12, 1, n_病人id, r_Pati.主页id, r_Pati.出院科室id, Null, v_误差, Null, '结帐作废退款', Null,
         Null, Null, d_Date, 操作员姓名_In, 操作员编号_In, 误差金额_In, n_结帐id, n_组id, 2);
    End If;
  End If;

  If n_来源 = 2 Or n_来源 = 3 Then
    --作废结帐对应的费用记录:不包含原始结帐产生的误差项目
    Insert Into 住院费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id,
       病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id,
       开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额, 结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要,
       缴款组id, 医疗小组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, To_Number('1' || Substr(记录性质, Length(记录性质), 1)), 记录状态, 序号, 从属父号, 价格父号, 多病人单,
             记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次,
             加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, Null, Null, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人,
             执行时间, 操作员姓名, 操作员编号, -1 * 结帐金额, n_结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要, 缴款组id, 医疗小组id
      From 住院费用记录
      Where 结帐id = n_原id And Nvl(附加标志, 0) <> 9;
  End If;

  If n_来源 = 1 Or n_来源 = 3 Then
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别,
       收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间,
       执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额, 结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要, 缴款组id, 主页id, 病人病区id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, To_Number('1' || Substr(记录性质, Length(记录性质), 1)), 记录状态, 序号, 从属父号, 价格父号, 记帐单id,
             病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费,
             记帐费用, 收入项目id, 收据费目, 标准单价, Null, Null, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号,
             -1 * 结帐金额, n_结帐id, 保险项目否, 保险大类id, 统筹金额, 是否急诊, 保险编码, 费用类型, 摘要, 缴款组id, 主页id, 病人病区id
      From 门诊费用记录
      Where 结帐id = n_原id And Nvl(附加标志, 0) <> 9;
  End If;
  --相关汇总表处理
  For r_Depositrow In c_Deposit(n_结帐id) Loop
    If r_Depositrow.记录性质 In (1, 11) Then
    
      --病人余额(预交)
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - r_Depositrow.冲预交 --注:新的结帐ID产生的是负数金额
      Where 病人id = r_Depositrow.病人id And 类型 = Nvl(r_Depositrow.预交类别, 2) And 性质 = 1
      Returning 预交余额 Into n_返回值;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 预交余额, 费用余额)
        Values
          (r_Depositrow.病人id, 1, Nvl(r_Depositrow.预交类别, 2), -1 * r_Depositrow.冲预交, 0);
        n_返回值 := -1 * r_Depositrow.冲预交;
      End If;
    
      If Nvl(n_返回值, 0) = 0 Then
        Delete 病人余额
        Where 性质 = 1 And 病人id = r_Depositrow.病人id And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
      End If;
    
    Else
      --人员缴款余额,医保不支持作废的结算方式在新的预交结算中已被处理为了退现金,
      --此处用加,表示收回退给病人的现金(结帐时,退款是负,作废时是正)
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + r_Depositrow.冲预交
      Where 收款员 = 操作员姓名_In And 结算方式 = r_Depositrow.结算方式 And 性质 = 1
      Returning 余额 Into n_返回值;
    
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (操作员姓名_In, r_Depositrow.结算方式, 1, r_Depositrow.冲预交);
        n_返回值 := -1 * r_Depositrow.冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 收款员 = 操作员姓名_In And 结算方式 = r_Depositrow.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
      End If;
    End If;
  End Loop;

  For r_Moneyrow In c_Money(n_结帐id) Loop
    --病人余额 ,误差项已结帐,所以不需要更新这两个汇总表
    If Nvl(v_误差no, 'sc') <> Nvl(r_Moneyrow.No, 'sc') Then
      If Nvl(r_Moneyrow.门诊标志, 0) = 1 Or Nvl(r_Moneyrow.门诊标志, 0) = 2 Then
        n_预交类别 := r_Moneyrow.门诊标志;
      Elsif Nvl(r_Moneyrow.主页id, 0) = 0 Or Nvl(r_Moneyrow.门诊标志, 0) = 4 Then
        --体检:门诊病人
        n_预交类别 := 1;
      Else
        n_预交类别 := 2;
      End If;
    
      If Nvl(r_Moneyrow.门诊标志, 0) <> 4 Then
        Update 病人余额
        Set 费用余额 = Nvl(费用余额, 0) - r_Moneyrow.结帐金额 --注:新的结帐ID产生的是负数金额
        Where 病人id = r_Moneyrow.病人id And 类型 = n_预交类别 And 性质 = 1
        Returning 费用余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 病人余额
            (病人id, 性质, 类型, 预交余额, 费用余额)
          Values
            (r_Moneyrow.病人id, 1, n_预交类别, 0, -1 * r_Moneyrow.结帐金额);
          n_返回值 := -1 * r_Moneyrow.结帐金额;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete 病人余额
          Where 病人id = r_Moneyrow.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
        End If;
      End If;
    
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) - r_Moneyrow.结帐金额
      Where 病人id = r_Moneyrow.病人id And Nvl(主页id, 0) = Nvl(r_Moneyrow.主页id, 0) And
            Nvl(病人病区id, 0) = Nvl(r_Moneyrow.病人病区id, 0) And Nvl(病人科室id, 0) = Nvl(r_Moneyrow.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(r_Moneyrow.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(r_Moneyrow.执行部门id, 0) And
            收入项目id + 0 = r_Moneyrow.收入项目id And 来源途径 + 0 = r_Moneyrow.门诊标志;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (r_Moneyrow.病人id, Decode(r_Moneyrow.主页id, Null, Null, 0, Null, r_Moneyrow.主页id),
           Decode(r_Moneyrow.病人病区id, Null, Null, 0, Null, r_Moneyrow.病人病区id), r_Moneyrow.病人科室id, r_Moneyrow.开单部门id,
           r_Moneyrow.执行部门id, r_Moneyrow.收入项目id, r_Moneyrow.门诊标志, -1 * r_Moneyrow.结帐金额);
      End If;
    End If;
  End Loop;

  If Nvl(缴预交id_In, 0) <> 0 Then
    --作废时将退款金额充值到预交款帐户,这里标明是本次结帐缴存的
    Update 病人预交记录 Set 结帐id = 冲销id_In Where ID = 缴预交id_In And 结帐id Is Null;
    If Sql%NotFound Then
      v_Err_Msg := '未找到对应的预交款记录！';
      Raise Err_Item;
    End If;
  End If;
  Close c_Pati;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人结帐记录_Delete;
/

--139063:冉俊明,2019-04-08,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_结帐费用记录_Insert
(
  Id_In       住院费用记录.Id%Type,
  No_In       住院费用记录.No%Type,
  记录性质_In 住院费用记录.记录性质%Type,
  记录状态_In 住院费用记录.记录状态%Type,
  执行状态_In 住院费用记录.执行状态%Type,
  序号_In     住院费用记录.序号%Type,
  结帐金额_In 住院费用记录.结帐金额%Type,
  结帐id_In   住院费用记录.结帐id%Type
) As
  n_Next_Id    住院费用记录.Id%Type;
  n_病人id     住院费用记录.病人id%Type;
  n_主页id     住院费用记录.主页id%Type;
  n_病人病区id 住院费用记录.病人病区id%Type;
  n_病人科室id 住院费用记录.病人科室id%Type;
  n_开单部门id 住院费用记录.开单部门id%Type;
  n_执行部门id 住院费用记录.执行部门id%Type;
  n_收入项目id 住院费用记录.收入项目id%Type;
  n_门诊标志   住院费用记录.门诊标志%Type;
  n_记帐费用   住院费用记录.记帐费用%Type;
  v_操作员     住院费用记录.操作员姓名%Type;
  v_操作员姓名 住院费用记录.操作员姓名%Type;

  n_结帐金额 住院费用记录.结帐金额%Type;
  n_实收金额 住院费用记录.实收金额%Type;
  n_返回值   病人余额.预交余额%Type;
  n_类别     Number(18);
  v_Temp     Varchar2(500);

  Err_Custom  Exception;
  Err_Special Exception;
  v_Error Varchar2(255);
  n_来源  Number;
Begin
  --人员id,人员编号,人员姓名
  v_Temp := Zl_Identity(1);
  If Not (Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_') Then
    v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_操作员姓名 := v_Temp;
  End If;

  If Id_In <> 0 Then
    Begin
      Select 2 Into n_来源 From 住院费用记录 Where ID = Id_In;
    Exception
      When Others Then
        n_来源 := 1;
    End;
  
    --第一次结帐但部分结
    If n_来源 = 1 Then
      Update 门诊费用记录 Set 结帐金额 = 结帐金额_In, 结帐id = 结帐id_In Where ID = Id_In And 结帐id Is Null;
    Else
      Update 住院费用记录 Set 结帐金额 = 结帐金额_In, 结帐id = 结帐id_In Where ID = Id_In And 结帐id Is Null;
    End If;
  
    If Sql%RowCount = 0 Then
      If n_来源 = 1 Then
        Select Max(b.操作员姓名)
        Into v_操作员
        From 门诊费用记录 A, 病人结帐记录 B
        Where a.Id = Id_In And b.Id = a.结帐id;
      Else
        Select Max(b.操作员姓名)
        Into v_操作员
        From 住院费用记录 A, 病人结帐记录 B
        Where a.Id = Id_In And b.Id = a.结帐id;
      End If;
      If v_操作员 Is Null Then
        v_Error := '未发现结帐的费用,当前结帐操作不能继续。';
        Raise Err_Custom;
      Else
        If v_操作员姓名 = v_操作员 Then
          v_Error := '发现已经被结帐的费用,当前结帐操作不能继续。';
          Raise Err_Special;
        Else
          v_Error := '发现已经被其他人结帐的费用,当前结帐操作不能继续。';
          Raise Err_Custom;
        End If;
      End If;
    End If;
  
    n_Next_Id := Id_In;
  Else
    --结以前的余帐
    Select 病人费用记录_Id.Nextval Into n_Next_Id From Dual;
  
    If Mod(记录性质_In, 10) = 3 Or Mod(记录性质_In, 10) = 5 Then
      --自动记帐或就诊卡;肯定是住院
      n_来源 := 2;
    Else
      Begin
        Select 2
        Into n_来源
        From 住院费用记录
        Where NO = No_In And 序号 = 序号_In And 记录状态 In (1, 2, 3) And Nvl(执行状态, 0) = Nvl(执行状态_In, 0) And
              Substr(记录性质, Length(记录性质), 1) = 记录性质_In And Rownum < 2;
      Exception
        When Others Then
          n_来源 := 1;
      End;
    End If;
  
    If n_来源 = 1 Then
      Insert Into 门诊费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别,
         收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间,
         执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额, 结帐id, 保险项目否, 保险大类id, 统筹金额, 保险编码, 费用类型, 是否急诊, 摘要, 主页id, 病人病区id)
        Select n_Next_Id, NO, 实际票号, To_Number('1' || 记录性质_In), 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄,
               标识号, 付款方式, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, Null,
               Null, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额_In, 结帐id_In, 保险项目否,
               保险大类id, 统筹金额, 保险编码, 费用类型, 是否急诊, 摘要, 主页id, 病人病区id
        From 门诊费用记录
        Where NO = No_In And 序号 = 序号_In And (记录状态 = 记录状态_In Or 记录状态 = Decode(记录状态_In, 1, 3, 记录状态_In)) And
              Nvl(执行状态, 0) = Nvl(执行状态_In, 0) And Substr(记录性质, Length(记录性质), 1) = 记录性质_In And Rownum < 2;
    
      --检查多次结帐后结帐金额是否高于原金额
      Select Nvl(Sum(实收金额), 0), Nvl(Sum(结帐金额), 0)
      Into n_实收金额, n_结帐金额
      From 门诊费用记录
      Where NO = No_In And 序号 = 序号_In And (记录状态 = 记录状态_In Or 记录状态 = Decode(记录状态_In, 1, 3, 记录状态_In)) And
            Substr(记录性质, Length(记录性质), 1) = 记录性质_In And Nvl(执行状态, 0) = 执行状态_In;
    Else
      Insert Into 住院费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id,
         病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人,
         开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额, 结帐id, 保险项目否, 保险大类id, 统筹金额, 保险编码, 费用类型,
         是否急诊, 摘要, 医疗小组id)
        Select n_Next_Id, NO, 实际票号, To_Number('1' || 记录性质_In), 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志,
               姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 记帐费用, 收入项目id,
               收据费目, 标准单价, Null, Null, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 执行人, 执行时间, 操作员姓名, 操作员编号, 结帐金额_In,
               结帐id_In, 保险项目否, 保险大类id, 统筹金额, 保险编码, 费用类型, 是否急诊, 摘要, 医疗小组id
        From 住院费用记录
        Where NO = No_In And 序号 = 序号_In And (记录状态 = 记录状态_In Or 记录状态 = Decode(记录状态_In, 1, 3, 记录状态_In)) And
              Nvl(执行状态, 0) = Nvl(执行状态_In, 0) And Substr(记录性质, Length(记录性质), 1) = 记录性质_In And Rownum < 2;
    
      --检查多次结帐后结帐金额是否高于原金额
      Select Nvl(Sum(实收金额), 0), Nvl(Sum(结帐金额), 0)
      Into n_实收金额, n_结帐金额
      From 住院费用记录
      Where NO = No_In And 序号 = 序号_In And (记录状态 = 记录状态_In Or 记录状态 = Decode(记录状态_In, 1, 3, 记录状态_In)) And
            Substr(记录性质, Length(记录性质), 1) = 记录性质_In And Nvl(执行状态, 0) = 执行状态_In;
    End If;
  
    If n_结帐金额 > n_实收金额 Then
      If n_来源 = 1 Then
        Select Max(b.操作员姓名)
        Into v_操作员
        From 门诊费用记录 A, 病人结帐记录 B
        Where a.Id = Id_In And b.Id = a.结帐id;
      Else
        Select Max(b.操作员姓名)
        Into v_操作员
        From 住院费用记录 A, 病人结帐记录 B
        Where a.Id = Id_In And b.Id = a.结帐id;
      End If;
    
      If v_操作员 Is Null Then
        v_Error := '未发现结帐的费用,当前结帐操作不能继续。';
        Raise Err_Custom;
      Else
        If v_操作员姓名 = v_操作员 Then
          v_Error := '发现已经被结帐的费用,当前结帐操作不能继续。';
          Raise Err_Special;
        Else
          v_Error := '发现已经被其他人结帐的费用,当前结帐操作不能继续。';
          Raise Err_Custom;
        End If;
      End If;
    End If;
  End If;
  If n_来源 = 1 Then
    Select 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 门诊标志, 记帐费用
    Into n_病人id, n_主页id, n_病人病区id, n_病人科室id, n_开单部门id, n_执行部门id, n_收入项目id, n_门诊标志, n_记帐费用
    From 门诊费用记录
    Where ID = n_Next_Id;
    n_类别 := 1;
  Else
    Select 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 门诊标志, 记帐费用
    Into n_病人id, n_主页id, n_病人病区id, n_病人科室id, n_开单部门id, n_执行部门id, n_收入项目id, n_门诊标志, n_记帐费用
    From 住院费用记录
    Where ID = n_Next_Id;
  
    If Nvl(n_门诊标志, 0) = 1 Or Nvl(n_门诊标志, 0) = 2 Then
      n_类别 := n_门诊标志;
    Elsif Nvl(n_主页id, 0) = 0 Or Nvl(n_门诊标志, 0) = 4 Then
      n_类别 := 1;
    Else
      n_类别 := 2;
    End If;
  End If;

  If Nvl(n_门诊标志, 0) <> 4 Then
    --病人余额
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) - 结帐金额_In
    Where 病人id = n_病人id And 性质 = 1 And 类型 = n_类别
    Returning 费用余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (n_病人id, 1, n_类别, 0, -1 * 结帐金额_In);
      n_返回值 := -1 * 结帐金额_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 病人余额 Where Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0 And 病人id = n_病人id;
    End If;
  End If;

  --病人未结费用
  Update 病人未结费用
  Set 金额 = Nvl(金额, 0) - 结帐金额_In
  Where 病人id = n_病人id And Nvl(主页id, 0) = Nvl(n_主页id, 0) And Nvl(病人病区id, 0) = Nvl(n_病人病区id, 0) And
        Nvl(病人科室id, 0) = Nvl(n_病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(n_开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(n_执行部门id, 0) And
        收入项目id + 0 = n_收入项目id And 来源途径 + 0 = n_门诊标志
  Returning 金额 Into n_返回值;
  If Sql%RowCount = 0 Then
    Insert Into 病人未结费用
      (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
    Values
      (n_病人id, Decode(n_主页id, 0, Null, n_主页id), Decode(n_病人病区id, 0, Null, n_病人病区id), n_病人科室id, n_开单部门id, n_执行部门id,
       n_收入项目id, n_门诊标志, -1 * 结帐金额_In);
    n_返回值 := -1 * 结帐金额_In;
  End If;
  If Nvl(n_返回值, 0) = 0 Then
    Delete From 病人未结费用 Where 病人id = n_病人id And Nvl(金额, 0) = 0;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_结帐费用记录_Insert;
/

--139063:冉俊明,2019-04-08,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_门诊收费记录_销帐
(
  No_In         门诊费用记录.No%Type,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type,
  序号_In       Varchar2 := Null,
  退费时间_In   门诊费用记录.登记时间%Type := Null,
  退费摘要_In   门诊费用记录.摘要%Type := Null,
  结帐id_In     病人预交记录.结帐id%Type := Null,
  回收票据_In   Number := 0
) As
  --功能：删除一张门诊收费单据
  --参数：
  --        序号_IN           =要退费的项目序号,格式为"1,3,5,6...",缺省NULL表示退"未退的"所有行。
  --        回收票据_In       =0:全退或最后一次全退时,收回票据。
  --                           1:部份退费不处理票据,通过重打调用单独处理。
  --该游标为要退费单据的所有原始记录

  --医保全退但某种结算退现金从而产生了新的误差时,排开此处的误差处理,执行完本过程后,界面程序中单独处理新误差
  Cursor c_Bill Is
    Select a.Id, a.No, a.附加标志, a.收费细目id, a.序号, a.价格父号, a.执行状态, a.收费类别, a.付数, a.数次, a.医嘱序号, j.诊疗类别, m.跟踪在用,
           Nvl(a.附加标志, 0) As 误差, Nvl(j.医嘱状态, 0) As 医嘱状态
    From 门诊费用记录 A, 病人医嘱记录 J, 材料特性 M
    Where a.医嘱序号 = j.Id(+) And a.No = No_In And a.记录性质 = 1 And a.记录状态 In (1, 3) And a.收费细目id + 0 = m.材料id(+)
    Order By a.收费细目id, a.序号;

  --:不管原始单据误差,都应该根据当前退费产生的误差项进行处理
  -- Decode(Sign(误差_In), 0, 999, 9)

  --该光标用于处理人员缴款余额中退的不同结算方式的金额

  n_结帐id 门诊费用记录.结帐id%Type;
  n_打印id 票据打印内容.Id%Type;

  --部分退费计算变量
  n_剩余数量 Number;
  n_剩余应收 Number;
  n_剩余实收 Number;
  n_剩余统筹 Number;
  n_准退数量 Number;
  n_退费次数 Number;

  n_应收金额 Number;
  n_实收金额 Number;
  n_统筹金额 Number;
  n_总金额   Number;
  n_组id     财务缴款分组.Id%Type;

  l_使用id   t_Numlist := t_Numlist();
  l_序号     t_Numlist := t_Numlist();
  l_执行状态 t_Numlist := t_Numlist();

  n_Dec   Number;
  d_Date  Date;
  n_Count Number;

  Err_Item Exception;
  v_Err_Msg  Varchar2(255);
  n_启用模式 Number(3);
  v_Para     Varchar2(1000);

Begin
  n_组id := Zl_Get组id(操作员姓名_In);
  --是否已经全部完全执行(只是该单据整张单据的检查)
  Select Nvl(Count(*), 0)
  Into n_Count
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Nvl(执行状态, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '该单据中的项目已经全部完全执行！';
    Raise Err_Item;
  End If;

  --未完全执行的项目是否有剩余数量(只是整张单据的检查)
  --执行状态在原始记录上判断
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select 序号, Sum(数量) As 剩余数量
         From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量, 结帐id
                From 门诊费用记录
                Where NO = No_In And Mod(记录性质, 10) = 1 And Nvl(附加标志, 0) <> 9 And
                      Nvl(价格父号, 序号) In
                      (Select Nvl(价格父号, 序号)
                       From 门诊费用记录
                       Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And Nvl(执行状态, 0) <> 1)
                Group By 记录状态, Nvl(价格父号, 序号), 结帐id)
         Group By 序号
         Having Sum(数量) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '该单据中未完全执行部分项目剩余数量为零,没有可以退费的费用！';
    Raise Err_Item;
  End If;

  ---------------------------------------------------------------------------------
  --公用变量
  If 退费时间_In Is Not Null Then
    d_Date := 退费时间_In;
  Else
    Select Sysdate Into d_Date From Dual;
  End If;

  If 结帐id_In Is Null Then
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  Else
    n_结帐id := 结帐id_In;
  End If;

  --金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --循环处理每行费用(收入项目行)
  n_总金额 := 0;
  For r_Bill In c_Bill Loop
    If Instr(',' || 序号_In || ',', ',' || Nvl(r_Bill.价格父号, r_Bill.序号) || ',') > 0 Or 序号_In Is Null Then
      If Nvl(r_Bill.执行状态, 0) <> 1 Then
        --求剩余数量,剩余应收,剩余实收
        Select Sum(Nvl(付数, 1) * 数次), Sum(应收金额), Sum(实收金额), Sum(统筹金额)
        Into n_剩余数量, n_剩余应收, n_剩余实收, n_剩余统筹
        From 门诊费用记录
        Where NO = No_In And Mod(记录性质, 10) = 1 And 序号 = r_Bill.序号;
      
        If n_剩余数量 = 0 Then
          If 序号_In Is Not Null Then
            v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经全部退费！';
            Raise Err_Item;
          End If;
        Else
          --准退数量(非药品项目为剩余数量,原始数量)
          If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Or (r_Bill.收费类别 = '4' And Nvl(r_Bill.跟踪在用, 0) = 0) Then
            --@@@
            --非药品部分(以具体医嘱执行为准进行检查)
            --: 1.存在医嘱执行计价的,则以医嘱执行为准(但不能包含:检查;检验;手术;麻醉及输血),已执行的不允许退费
            --: 2.不存在医嘱执行计价的,则以剩余数量为准
            --: 3.医嘱作废了的,则以剩余数量为准(病人医嘱记录.医嘱状态=4表示作废医嘱，会删除"病人医嘱发送",门诊药嘱先作废后退药)
            --: 4.病人医嘱发送.执行状态=1（完成执行）时，准退数为0，不再根据医嘱执行计价来统计准退数
            n_Count := 0;
            If Instr(',C,D,F,G,K,', ',' || r_Bill.诊疗类别 || ',') = 0 And r_Bill.诊疗类别 Is Not Null And r_Bill.医嘱状态 <> 4 Then
              Select Nvl(Sum(Decode(b.执行状态, 1, 0, 1) * Decode(c.执行状态, 0, 1, 0) * c.数量), 0), Count(1)
              Into n_准退数量, n_Count
              From 病人医嘱发送 B, 医嘱执行计价 C
              Where b.医嘱id = r_Bill.医嘱序号 And b.No = r_Bill.No And b.医嘱id = c.医嘱id And b.发送号 = c.发送号 And
                    c.收费细目id + 0 = r_Bill.收费细目id And b.记录性质 = 1;
            End If;
            If Nvl(n_Count, 0) <> 0 And n_准退数量 = 0 Then
              v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已执行，不允许退费！';
              Raise Err_Item;
            End If;
          
            If Nvl(n_Count, 0) = 0 Then
              If r_Bill.执行状态 = 2 Then
                --无医嘱执行计价的部分退费无法判断准退数量，不允许退费
                v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已部分执行，无法判断准退数量，不允许退费！';
                Raise Err_Item;
              Else
                n_准退数量 := n_剩余数量;
              End If;
            End If;
          
          Else
            Select Nvl(Sum(Nvl(付数, 1) * 实际数量), 0), Count(*)
            Into n_准退数量, n_Count
            From 药品收发记录
            Where NO = No_In And 单据 In (8, 24) And Mod(记录状态, 3) = 1 --@@@
                  And 审核人 Is Null And 费用id = r_Bill.Id;
          
            --有剩余数量无准退数量的有两种情况：
            --1.不跟踪在用的卫材无对应的收发记录,这时使用剩余数量
            --2.并发操作,此时已发药或发料
            If n_准退数量 = 0 Then
              If r_Bill.收费类别 = '4' Then
                If n_Count > 0 Then
                  v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发料,须退料后再退费！';
                  Raise Err_Item;
                Else
                  n_准退数量 := n_剩余数量;
                End If;
              Else
                v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发药,须退药后再退费！';
                Raise Err_Item;
              End If;
            End If;
          End If;
        
          If n_准退数量 > n_剩余数量 Then
            v_Err_Msg := '单据[' || No_In || '] 中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用的退费数量(' || n_准退数量 ||
                         ')大于了剩余数量(' || n_剩余数量 || ')，不允许退费！';
            Raise Err_Item;
          End If;
          --收费的时候是负数数量的不检查准退数量是否小于零
          If n_准退数量 < 0 And Nvl(r_Bill.付数, 1) * r_Bill.数次 > 0 Then
            v_Err_Msg := '单据[' || No_In || '] 中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用的退费数量(' || n_准退数量 ||
                         ')小于了零，不允许退费！';
            Raise Err_Item;
          End If;
        
          --该笔项目第几次退费
          Select Nvl(Max(Abs(执行状态)), 0) + 1
          Into n_退费次数
          From 门诊费用记录
          Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 = 2 And Nvl(执行状态, 0) < 0 And 序号 = r_Bill.序号;
        
          --金额=剩余金额*(准退数/剩余数)
          n_应收金额 := Round(n_剩余应收 * (n_准退数量 / n_剩余数量), n_Dec);
          n_实收金额 := Round(n_剩余实收 * (n_准退数量 / n_剩余数量), n_Dec);
          n_统筹金额 := Round(n_剩余统筹 * (n_准退数量 / n_剩余数量), n_Dec);
          n_总金额   := n_总金额 + n_实收金额;
        
          --插入退费记录
          Insert Into 门诊费用记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别,
             收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人,
             执行状态, 费用状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 结论,
             缴款组id, 挂号id, 主页id, 病人病区id)
            Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别,
                   病人科室id, 收费类别, 收费细目id, 计算单位, Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, 付数, 1), 发药窗口,
                   Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, -1 * 数次, -1 * n_准退数量), 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价,
                   -1 * n_应收金额, -1 * n_实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, -1 * n_退费次数, 1, 执行时间, 操作员编号_In, 操作员姓名_In,
                   发生时间, d_Date, n_结帐id, -1 * n_实收金额, 保险项目否, 保险大类id, -1 * n_统筹金额, Nvl(退费摘要_In, 摘要),
                   Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, 结论, n_组id, 挂号id, 主页id, 病人病区id
            From 门诊费用记录
            Where ID = r_Bill.Id;
        
          --标记原费用记录
          l_序号.Extend;
          l_序号(l_序号.Count) := r_Bill.序号;
          l_执行状态.Extend;
          l_执行状态(l_执行状态.Count) := Case
                                    When Sign(n_准退数量 - n_剩余数量) = 0 Then
                                     0
                                    Else
                                     1
                                  End;
        End If;
      Else
        If 序号_In Is Not Null Then
          v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经完全执行,不能退费！';
          Raise Err_Item;
        End If;
      End If;
    End If;
  End Loop;
  --标记原费用记录
  Forall I In 1 .. l_序号.Count
    Update 门诊费用记录
    Set 记录状态 = 3, 执行状态 = l_执行状态(I)
    Where Mod(记录性质, 10) = 1 And NO = No_In And 序号 = l_序号(I) And 记录状态 In (1, 3);

  l_序号.Delete;
  For c_结帐 In (Select Distinct b.结帐id
               From 门诊费用记录 A, 病人预交记录 B
               Where a.结帐id = b.结帐id And a.No = No_In And Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And
                     Nvl(b.记录状态, 0) = 1) Loop
    l_序号.Extend;
    l_序号(l_序号.Count) := c_结帐.结帐id;
  End Loop;

  Forall I In 1 .. l_序号.Count
    Update 病人预交记录 Set 记录状态 = 3 Where 结帐id = l_序号(I) And Mod(记录性质, 10) <> 1;

  ---------------------------------------------------------------------------------
  --退费票据回收(仅全退时才回退,部分退是在重打过程中回收)
  If 回收票据_In = 1 Then
  
    --启用标志||NO;执行科室(条数);收据费目(首页汇总,条数);收费细目(条数)
    v_Para     := Nvl(zl_GetSysParameter('票据分配规则', 1121), '0||0;0;0,0;0');
    n_启用模式 := Zl_To_Number(Substr(v_Para, 1, 1));
    If n_启用模式 <> 0 Then
      --收回票据
      Select 使用id
      Bulk Collect
      Into l_使用id
      From (Select Distinct b.使用id From 票据打印明细 B Where b.No = No_In And Nvl(b.票种, 0) = 1);
    
      n_启用模式 := l_使用id.Count;
      If l_使用id.Count <> 0 Then
        --插入回收记录
        Forall I In 1 .. l_使用id.Count
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, 操作员姓名_In, d_Date, 票据金额
            From 票据使用明细 A
            Where ID = l_使用id(I) And 性质 = 1 And Not Exists
             (Select 1 From 票据使用明细 Where 号码 = a.号码 And 票种 = a.票种 And Nvl(性质, 0) <> 1);
      
        Forall I In 1 .. l_使用id.Count
          Update 票据打印明细 Set 是否回收 = 1 Where 使用id = l_使用id(I) And Nvl(是否回收, 0) = 0;
      
      End If;
    End If;
    If n_启用模式 = 0 Then
      --获取单据最后一次的打印ID(可能是多张单据收费打印)
      Begin
        --性质=1，原因=6为退费打印票据(红票)，不回收
        Select ID
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 1 And b.No = No_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
      --可能以前没有打印,无收回
      If n_打印id Is Not Null Then
        --a.多张单据循环调用时只能收回一次
        Select Count(*) Into n_Count From 票据使用明细 Where 票种 = 1 And 性质 = 2 And 打印id = n_打印id;
        If n_Count = 0 Then
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In, 票据金额
            From 票据使用明细
            Where 打印id = n_打印id And 票种 = 1 And 性质 = 1;
        Else
          --b.部分退费多次收回时,最后一次全退收回要排开已收回的
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In, 票据金额
            From 票据使用明细 A
            Where 打印id = n_打印id And 票种 = 1 And 性质 = 1 And Not Exists
             (Select 1 From 票据使用明细 B Where a.号码 = b.号码 And 打印id = n_打印id And 票种 = 1 And 性质 = 2);
        End If;
      End If;
    End If;
  End If;

  ---------------------------------------------------------------------------------
  --药品卫材相关内容
  --必须按照“收费细目id”升序排序，防止并发锁“药品库存”表
  For r_Expenses In (Select ID
                     From 门诊费用记录
                     Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And 收费类别 In ('4', '5', '6', '7') And
                           (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null)
                     Order By 收费细目id) Loop
    Zl_药品收发记录_销售退费(r_Expenses.Id);
  End Loop;

  --医嘱处理
  --删除病人医嘱附费(最后一次删除时)
  For c_医嘱 In (Select Distinct 医嘱序号
               From 门诊费用记录
               Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 = 3 And 医嘱序号 Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select 序号, Sum(数量) As 剩余数量
           From (Select 记录状态, 执行状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                  From 门诊费用记录
                  Where Mod(记录性质, 10) = 1 And Nvl(附加标志, 0) <> 9 And 医嘱序号 + 0 = c_医嘱.医嘱序号 And NO = No_In
                  Group By 记录状态, 执行状态, Nvl(价格父号, 序号))
           Group By 序号
           Having Sum(数量) <> 0);
  
    If n_Count = 0 Then
      Delete From 病人医嘱附费 Where 医嘱id = c_医嘱.医嘱序号 And 记录性质 = 1 And NO = No_In;
    End If;
  End Loop;

  --调整医嘱执行计价.执行状态 NULL-历史数据；0-未执行；1-已执行；2-已退费
  For c_费用 In (Select Distinct a.医嘱序号 As 医嘱id, a.收费细目id, b.发送号
               From 门诊费用记录 A, 病人医嘱发送 B
               Where a.医嘱序号 = b.医嘱id And a.No = b.No And a.记录性质 = 1 And a.记录状态 In (1, 3) And a.No = No_In And
                     (Instr(',' || 序号_In || ',', ',' || a.序号 || ',') > 0 Or 序号_In Is Null) And a.价格父号 Is Null And
                     b.记录性质 = 1) Loop
    Update 医嘱执行计价
    Set 执行状态 = 2
    Where 医嘱id = c_费用.医嘱id And 发送号 = c_费用.发送号 And 收费细目id = c_费用.收费细目id And 执行状态 = 0;
  End Loop;

  --场合_In    Integer:=0, --0:门诊;1-住院
  --性质_In    Integer:=1, --1-收费单;2-记帐单
  --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
  --No_In      门诊费用记录.No%Type,
  --医嘱ids_In Varchar2 := Null
  Zl_医嘱发送_计费状态_Update(0, 1, 2, No_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊收费记录_销帐;
/

--139063:冉俊明,2019-04-08,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_门诊收费记录_重收
(
  原结帐id_In     门诊费用记录.结帐id%Type,
  冲销id_In       门诊费用记录.结帐id%Type,
  重收结帐id_In   门诊费用记录.结帐id%Type,
  排开医保结算_In Varchar2 := Null
) As
  --排开医保结算_IN:多个用逗号分离(只某些医保结算,允许退现金)
  Cursor c_Fee_Data Is
    Select ID
    From 门诊费用记录 A
    Where 结帐id = 原结帐id_In And Not Exists
     (Select 1
           From 门诊费用记录 B
           Where Mod(b.记录性质, 10) = 1 And a.No = b.No And a.序号 = b.序号 And 结帐id = 冲销id_In)
    Order By ID;

  v_操作员编号 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  d_登记时间   门诊费用记录.登记时间%Type;
  n_缴款组id   门诊费用记录.缴款组id%Type;
  n_病人id     门诊费用记录.病人id%Type;
  Err_Item Exception;
  v_Err_Msg    Varchar2(255);
  n_Array_Size Number := 200;
  t_费用id     t_Numlist;
  n_结算金额   门诊费用记录.实收金额%Type;
  n_冲销金额   病人预交记录.冲预交%Type;
  n_Count      Number(18);
Begin
  Begin
    Select 操作员编号, 操作员姓名, 登记时间, 缴款组id, 病人id
    Into v_操作员编号, v_操作员姓名, d_登记时间, n_缴款组id, n_病人id
    From 门诊费用记录
    Where 结帐id = 冲销id_In And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := 'NO';
  End;

  If Nvl(v_Err_Msg, '-') = 'NO' Then
    v_Err_Msg := '由于并发操作,该单据可能已经初他人退费或删除,不能再进行退费操作！';
    Raise Err_Item;
  End If;

  --1.处理界面选择的且是部分退或部分执行的
  Insert Into 门诊费用记录
    (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id, 计算单位,
     付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态, 费用状态, 执行时间,
     操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 结论, 缴款组id, 挂号id, 主页id, 病人病区id)
    Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id,
           收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人,
           执行人, 执行状态, 费用状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 结论,
           缴款组id, 挂号id, 主页id, 病人病区id
    From (Select NO, Max(实际票号) As 实际票号, 11 As 记录性质, 1 As 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式,
                  费别, 病人科室id, 收费类别, 收费细目id, 计算单位, 1 As 付数, Max(发药窗口) As 发药窗口, Sum(Nvl(付数, 1) * Nvl(数次, 0)) As 数次,
                  Max(加班标志) As 加班标志, Max(附加标志) As 附加标志, 收入项目id, 收据费目, 记帐费用, Avg(标准单价) As 标准单价, Sum(应收金额) As 应收金额,
                  Sum(实收金额) As 实收金额, 开单部门id, 开单人, 执行部门id, Max(划价人) As 划价人, Max(执行人) 执行人, Max(执行状态) As 执行状态, 1 As 费用状态,
                  Max(执行时间) 执行时间, v_操作员编号 As 操作员编号, v_操作员姓名 As 操作员姓名, 发生时间, d_登记时间 As 登记时间, 重收结帐id_In As 结帐id,
                  Sum(结帐金额) As 结帐金额, Max(保险项目否) As 保险项目否, 保险大类id, Sum(统筹金额) As 统筹金额,
                  Max(Decode(记录性质, 1, 摘要, 11, 摘要, Null)) As 摘要, 0 As 是否上传, Max(保险编码) As 保险编码, Max(费用类型) As 费用类型,
                  Max(Decode(记录性质, 1, 结论, 11, 结论, Null)) As 结论, n_缴款组id As 缴款组id, Max(挂号id) As 挂号id, Max(主页id) As 主页id,
                  Max(病人病区id) As 病人病区id
           From 门诊费用记录
           Where Mod(记录性质, 10) = 1 And (NO, 序号) In (Select NO, 序号 From 门诊费用记录 Where 结帐id = 冲销id_In)
           Group By NO, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id, 计算单位, 收入项目id,
                    收据费目, 记帐费用, 开单部门id, 开单人, 执行部门id, 发生时间, 保险大类id
           Having Sum(Nvl(付数, 1) * Nvl(数次, 0)) <> 0);

  For c_冲销 In (Select NO, 序号, 从属父号, 价格父号, 收入项目id, -1 * Sum(Nvl(付数, 1) * Nvl(数次, 0)) As 数次, Sum(标准单价) As 标准单价,
                      -1 * Sum(应收金额) As 应收金额, -1 * Sum(实收金额) As 实收金额, -1 * Sum(统筹金额) As 统筹金额, -1 * Sum(结帐金额) As 结帐金额
               From 门诊费用记录
               Where 记录性质 = 11 And 结帐id = 重收结帐id_In
               Group By NO, 序号, 从属父号, 价格父号, 收入项目id) Loop
    Update 门诊费用记录
    Set 数次 = Nvl(数次, 0) + Nvl(c_冲销.数次, 0), 实收金额 = Nvl(实收金额, 0) + Nvl(c_冲销.实收金额, 0),
        应收金额 = Nvl(应收金额, 0) + Nvl(c_冲销.应收金额, 0), 结帐金额 = Nvl(结帐金额, 0) + Nvl(c_冲销.结帐金额, 0),
        统筹金额 = Nvl(统筹金额, 0) + Nvl(c_冲销.统筹金额, 0)
    Where NO = c_冲销.No And 序号 = c_冲销.序号 And Nvl(从属父号, -1) = Nvl(c_冲销.从属父号, '-1') And
          Nvl(价格父号, -1) = Nvl(c_冲销.价格父号, '-1') And 收入项目id = c_冲销.收入项目id And 结帐id = 冲销id_In;
  End Loop;

  --2.处理界面未选退费部分,需要全退且产生11的重收记录
  Open c_Fee_Data;
  Loop
    Fetch c_Fee_Data Bulk Collect
      Into t_费用id Limit n_Array_Size;
    Exit When t_费用id.Count = 0;
  
    --退费记录
    Forall I In 1 .. t_费用id.Count
      Insert Into 门诊费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id,
         计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态, 费用状态,
         执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 结论, 缴款组id, 挂号id, 主页id,
         病人病区id)
        Select 病人费用记录_Id.Nextval, a.No, a.实际票号, 1, 2, a.序号, a.从属父号, a.价格父号, a.病人id, a.医嘱序号, a.门诊标志, a.姓名, a.性别, a.年龄,
               a.标识号, a.付款方式, a.费别, a.病人科室id, a.收费类别, a.收费细目id, a.计算单位, a.付数, a.发药窗口, -1 * a.数次, a.加班标志, a.附加标志,
               a.收入项目id, a.收据费目, a.记帐费用, a.标准单价, -1 * a.应收金额, -1 * a.实收金额, a.开单部门id, a.开单人, a.执行部门id, a.划价人, 执行人,
               Nvl(q.执行状态, -1) As 执行状态, 1, a.执行时间, v_操作员编号, v_操作员姓名, a.发生时间, d_登记时间, 冲销id_In, -1 * a.结帐金额, a.保险项目否,
               a.保险大类id, -1 * a.统筹金额, a.摘要, 0 As 是否上传, a.保险编码, a.费用类型, a.结论, n_缴款组id As 缴款组id, 挂号id, 主页id, 病人病区id
        From 门诊费用记录 A,
             (Select j.No, j.序号, Nvl(Max(j.执行状态), 0) - 1 As 执行状态
               From 门诊费用记录 M, 门诊费用记录 J
               Where m.Id = t_费用id(I) And m.No = j.No And m.序号 = j.序号 And Mod(j.记录性质, 10) = 1 And j.记录状态 = 2
               Group By j.No, j.序号) Q
        Where ID = t_费用id(I) And a.No = q.No(+) And a.序号 = q.序号(+);
  
    --将原记录状态由1变为3
    Forall I In 1 .. t_费用id.Count
      Update 门诊费用记录 Set 记录状态 = 3 Where ID = t_费用id(I) And 记录状态 = 1;
  
    --重新收费记录
    If Nvl(重收结帐id_In, 0) <> 0 Then
      Forall I In 1 .. t_费用id.Count
        Insert Into 门诊费用记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id,
           计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态,
           费用状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 结论, 缴款组id, 挂号id,
           主页id, 病人病区id)
          Select 病人费用记录_Id.Nextval, NO, 实际票号, 11, 1, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id,
                 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id,
                 划价人, 执行人, 执行状态, 1, 执行时间, v_操作员编号, v_操作员姓名, 发生时间, d_登记时间, 重收结帐id_In, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要,
                 0 As 是否上传, 保险编码, 费用类型, 结论, n_缴款组id As 缴款组id, 挂号id, 主页id, 病人病区id
          From 门诊费用记录
          Where ID = t_费用id(I);
    End If;
  End Loop;
  Close c_Fee_Data;

  Select Count(1) Into n_Count From 病人预交记录 Where 结帐id = 冲销id_In And 结算方式 Is Null;
  If n_Count = 0 Then
    --退费结算方式
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 结算性质)
      Select 病人预交记录_Id.Nextval, 3, Null, 2, n_病人id, 结算方式, d_登记时间, v_操作员编号, v_操作员姓名, -1 * 冲预交, 冲销id_In, n_缴款组id,
             -1 * 冲销id_In, 2, 3
      From 病人预交记录
      Where 结帐id = 原结帐id_In And 结算方式 In (Select 名称 From 结算方式 Where 性质 In (3, 4)) And
            Instr(',' || 排开医保结算_In || ',', ',' || 结算方式 || ',') = 0 And Mod(记录性质, 10) <> 1;
    --将原误差费全部退了
    --Insert Into 病人预交记录
    --  (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志)
    --  Select 病人预交记录_Id.Nextval, 3, Null, 2, n_病人id, 结算方式, d_登记时间, v_操作员编号, v_操作员姓名, -1 * 冲预交, 冲销id_In, n_缴款组id,
    --         -1 * 冲销id_In, 2
    --  From 病人预交记录
    --  Where 结帐id = 原结帐id_In And 结算方式 = v_误差费 And Mod(记录性质, 10) <> 1;
  
    Select Sum(冲预交) Into n_冲销金额 From 病人预交记录 Where 结帐id = 冲销id_In;
    Select Sum(结帐金额) Into n_结算金额 From 门诊费用记录 Where 结帐id = 冲销id_In;
  
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 3, Null, 2, n_病人id, Null, d_登记时间, v_操作员编号, v_操作员姓名, -1 * (Nvl(n_冲销金额, 0) - Nvl(n_结算金额, 0)),
       冲销id_In, n_缴款组id, -1 * 冲销id_In, 1, 3);
  
  End If;
  If Nvl(重收结帐id_In, 0) <> 0 Then
    Select Count(1) Into n_Count From 病人预交记录 Where 结帐id = 重收结帐id_In And 结算方式 Is Null;
    If n_Count = 0 Then
      Select Sum(结帐金额) Into n_结算金额 From 门诊费用记录 Where 结帐id = 重收结帐id_In;
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 校对标志, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 3, Null, 1, n_病人id, Null, d_登记时间, v_操作员编号, v_操作员姓名, n_结算金额, 重收结帐id_In, n_缴款组id,
         -1 * 冲销id_In, 1, 3);
    End If;
  
    --场合_In    Integer:=0, --0:门诊;1-住院
    --性质_In    Integer:=1, --1-收费单;2-记帐单
    --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
    --No_In      门诊费用记录.No%Type,
    --医嘱ids_In Varchar2 := Null
    For c_No In (Select Distinct NO From 门诊费用记录 Where 记录性质 = 11 And 结帐id = 重收结帐id_In) Loop
      Zl_医嘱发送_计费状态_Update(0, 1, 2, c_No.No);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊收费记录_重收;
/

--139063:冉俊明,2019-04-08,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_门诊收费记录_Delete
(
  No_In           门诊费用记录.No%Type,
  操作员编号_In   门诊费用记录.操作员编号%Type,
  操作员姓名_In   门诊费用记录.操作员姓名%Type,
  医保结算方式_In Varchar2 := Null,
  序号_In         Varchar2 := Null,
  结算方式_In     病人预交记录.结算方式%Type := Null,
  误差_In         门诊费用记录.实收金额%Type := 0,
  退费时间_In     门诊费用记录.登记时间%Type := Null,
  回收票据_In     Number := 0,
  退费摘要_In     门诊费用记录.摘要%Type := Null,
  校对标志_In     Number := 0,
  结帐id_In       病人预交记录.结帐id%Type := Null,
  结算序号_In     病人预交记录.结算序号%Type := Null,
  一卡通结算_In   Varchar2 := Null,
  退款操作_In     Number := 0,
  多单据全退_In   Number := 0
) As
  --功能：删除一张门诊收费单据
  --参数：
  --        医保结算方式_IN   =医保退费时,不支持结算作废的结算方式,如果为空表示非医保退费或医保退费全部结算允许作废。
  --        序号_IN           =要退费的项目序号,格式为"1,3,5,6...",缺省NULL表示退"未退的"所有行。
  --        结算方式_IN       =当为部分退费时,退费金额的结算方式。
  --        误差_IN           =指退费时新产生的误差金额,部份退费或医保全退但某种结算退现金时才会产生新的误差。
  --                           此时传入仅用于计算本次退费的结算金额,误差费用记录的处理在本过程执行完后调用Zl_门诊收费误差_Insert产生
  --        回收票据_In       =0:单张全退或多张一起全退时收回票据,注意,多张单据退费循环调本过程时只收回一次。
  --                           1:部份退费不处理票据,通过重打调用单独处理。
  --        校对标志_IN:0-不需要较对;1-需较对(不处理人员缴款余额,不回收票据,不处理预交余额)
  --        一卡通结算_In 全退时传入不原样退回的结算方式；医疗卡部分退费时，传入"结算方式|金额"
  --        退款操作_In:1-进行部分退(将退款方式退到指定的结算方式<结算方式_In>中,0-不指定退款方式)
  --        多单据全退_IN=1-多单据全退(多张单据全退,原样退);0-非原样退
  --该游标为要退费单据的所有原始记录

  --医保全退但某种结算退现金从而产生了新的误差时,排开此处的误差处理,执行完本过程后,界面程序中单独处理新误差
  Cursor c_Bill Is
    Select a.Id, a.No, a.附加标志, a.收费细目id, a.序号, a.价格父号, a.执行状态, a.收费类别, a.付数, a.数次, a.医嘱序号, j.诊疗类别, m.跟踪在用,
           Nvl(a.附加标志, 0) As 误差
    From 门诊费用记录 A, 病人医嘱记录 J, 材料特性 M
    Where a.医嘱序号 = j.Id(+) And a.No = No_In And a.记录性质 = 1 And a.记录状态 In (1, 3) And a.收费细目id + 0 = m.材料id(+) And
          Nvl(a.附加标志, 0) <> Decode(多单据全退_In, 1, 999, 9)
    Order By a.收费细目id, a.序号;
  --:不管原始单据误差,都应该根据当前退费产生的误差项进行处理
  -- Decode(Sign(误差_In), 0, 999, 9)

  --该光标用于处理人员缴款余额中退的不同结算方式的金额
  Cursor c_Money(冲销id_In 病人预交记录.结帐id%Type) Is
    Select 结算方式, 冲预交
    From 病人预交记录
    Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = 冲销id_In And 结算方式 Is Not Null And Nvl(冲预交, 0) <> 0 And Nvl(校对标志, 0) = 0;

  --该游标用于查找收费时使用过的冲预交款记录
  Cursor c_Deposit(V结帐id 病人预交记录.结帐id%Type) Is
    Select ID, 病人id, 冲预交 As 金额, 预交类别
    From 病人预交记录
    Where 记录性质 In (1, 11) And 记录状态 In (1, 3) And 结帐id = V结帐id And Nvl(冲预交, 0) <> 0
    Order By ID Desc;

  n_病人id   病人信息.病人id%Type;
  n_结帐id   门诊费用记录.结帐id%Type;
  n_结算序号 病人预交记录.结算序号%Type;
  n_打印id   票据打印内容.Id%Type;

  n_已退金额 病人预交记录.冲预交%Type;
  n_预交金额 病人预交记录.冲预交%Type;
  n_返回值   病人预交记录.冲预交%Type;
  n_原误差费 门诊费用记录.实收金额%Type;
  --部分退费计算变量
  n_剩余数量 Number;
  n_剩余应收 Number;
  n_剩余实收 Number;
  n_剩余统筹 Number;
  n_准退数量 Number;
  n_退费次数 Number;

  n_应收金额 Number;
  n_实收金额 Number;
  n_统筹金额 Number;
  n_总金额   Number;
  n_费用状态 门诊费用记录.费用状态%Type;
  n_正常退费 Number; --是否第一次退费且全部退费,在每行退费过程中判断得到。
  n_组id     财务缴款分组.Id%Type;

  v_退费结算 结算方式.名称%Type;
  v_结算内容 Varchar2(500);
  n_部分退   Number(2);
  v_当前结算 Varchar2(50);
  v_结算方式 病人预交记录.结算方式%Type;
  n_结算金额 病人预交记录.冲预交%Type;
  n_预交id   病人预交记录.Id%Type;

  l_使用id   t_Numlist := t_Numlist();
  n_Dec      Number;
  d_Date     Date;
  n_Count    Number;
  n_原结帐id Number;

  Err_Item Exception;
  v_Err_Msg      Varchar2(255);
  n_启用模式     Number(3);
  v_Para         Varchar2(1000);
  n_医属执行计价 Number;
  n_会话号       病人预交记录.会话号%Type; --格式：SID+'_'+SERIAL#

Begin
  n_组id := Zl_Get组id(操作员姓名_In);

  Begin
    Select Sid || '_' || Serial# Into n_会话号 From V$session Where Audsid = Userenv('sessionid');
  Exception
    When Others Then
      n_会话号 := Null;
  End;

  n_部分退 := 0;
  --是否已经全部完全执行(只是该单据整张单据的检查)
  Select Nvl(Count(*), 0)
  Into n_Count
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Nvl(执行状态, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '该单据中的项目已经全部完全执行！';
    Raise Err_Item;
  End If;

  --未完全执行的项目是否有剩余数量(只是整张单据的检查)
  --执行状态在原始记录上判断
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select 序号, Sum(数量) As 剩余数量
         From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                From 门诊费用记录
                Where NO = No_In And 记录性质 = 1 And Nvl(附加标志, 0) <> 9 And
                      Nvl(价格父号, 序号) In
                      (Select Nvl(价格父号, 序号)
                       From 门诊费用记录
                       Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Nvl(执行状态, 0) <> 1)
                Group By 记录状态, Nvl(价格父号, 序号))
         Group By 序号
         Having Sum(数量) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '该单据中未完全执行部分项目剩余数量为零,没有可以退费的费用！';
    Raise Err_Item;
  End If;
  --确定是否在医嘱执行计价中存在数据,如果存在数据,则根据医嘱执行计价进行退费,否则按旧方式进行处理
  Select Count(1)
  Into n_医属执行计价
  From 门诊费用记录 A, 医嘱执行计价 B
  Where a.医嘱序号 = b.医嘱id And a.记录性质 = 1 And a.No = No_In And a.记录状态 In (1, 3) And Rownum = 1;

  ---------------------------------------------------------------------------------
  --公用变量
  If 退费时间_In Is Not Null Then
    d_Date := 退费时间_In;
  Else
    Select Sysdate Into d_Date From Dual;
  End If;
  If 结帐id_In Is Null Then
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  Else
    n_结帐id := 结帐id_In;
  End If;
  n_结算序号 := 结算序号_In;
  If n_结算序号 Is Null Then
    n_结算序号 := 结帐id_In;
  End If;
  --金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --获取结算方式名称
  v_退费结算 := 结算方式_In;
  If v_退费结算 Is Null Then
    Begin
      Select 名称 Into v_退费结算 From 结算方式 Where 性质 = 1;
    Exception
      When Others Then
        v_退费结算 := '现金';
    End;
  End If;
  --循环处理每行费用(收入项目行)
  n_总金额   := 0;
  n_正常退费 := 1;
  For r_Bill In c_Bill Loop
    If Instr(',' || 序号_In || ',', ',' || Nvl(r_Bill.价格父号, r_Bill.序号) || ',') > 0 Or 序号_In Is Null Then
      If Nvl(r_Bill.执行状态, 0) <> 1 Then
        --求剩余数量,剩余应收,剩余实收
        Select Sum(Nvl(付数, 1) * 数次), Sum(应收金额), Sum(实收金额), Sum(统筹金额)
        Into n_剩余数量, n_剩余应收, n_剩余实收, n_剩余统筹
        From 门诊费用记录
        Where NO = No_In And 记录性质 = 1 And 序号 = r_Bill.序号;
      
        If n_剩余数量 = 0 Then
          If 序号_In Is Not Null Then
            v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经全部退费！';
            Raise Err_Item;
          End If;
          --情况：未限定行号,原始单据中的该笔已经全部退费(执行状态=0的一种可能)
          n_正常退费 := 0;
        Else
          --准退数量(非药品项目为剩余数量,原始数量)
          If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Or (r_Bill.收费类别 = '4' And Nvl(r_Bill.跟踪在用, 0) = 0) Then
            --@@@
            --非药品部分(以具体医嘱执行为准进行检查)
            --: 1.存在医嘱发送的,则以医嘱执行为准(但不能包含:检查;检验;手术;麻醉及输血)
            --: 2.不存在医嘱的,则以剩余数量为准
            n_Count := 0;
            If Instr(',C,D,F,G,K,', ',' || r_Bill.诊疗类别 || ',') = 0 And r_Bill.诊疗类别 Is Not Null Then
              If n_医属执行计价 = 1 Then
                Select Decode(Sign(Sum(数量)), -1, 0, Sum(数量)), Count(*)
                Into n_准退数量, n_Count
                From (Select Max(Decode(a.记录状态, 2, 0, a.Id)) As ID, Max(a.医嘱序号) As 医嘱id, Max(a.收费细目id) As 收费细目id,
                              Sum(Nvl(a.付数, 1) * Nvl(a.数次, 1)) As 数量,
                              Sum(Decode(a.记录状态, 2, 0, Nvl(a.付数, 1) * Nvl(a.数次, 1))) As 原始数量
                       From 门诊费用记录 A, 病人医嘱记录 M
                       Where a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And
                             Instr('5,6,7', a.收费类别) = 0 And a.No = No_In And a.序号 = r_Bill.序号 And a.记录性质 = 1 And
                             a.记录状态 In (1, 2, 3) And a.价格父号 Is Null
                       Group By a.序号
                       Union All
                       Select a.Id, a.医嘱序号 As 医嘱id, a.收费细目id, -1 * b.数量 As 已执行, 0 原始数量
                       From 门诊费用记录 A, 医嘱执行计价 B, 病人医嘱记录 M
                       Where a.医嘱序号 = b.医嘱id And a.收费细目id = b.收费细目id + 0 And a.医嘱序号 = m.Id And
                             Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Instr('5,6,7', a.收费类别) = 0 And
                             (Exists
                              (Select 1
                               From 病人医嘱执行
                               Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And b.要求时间 = 要求时间 And Nvl(执行结果, 0) = 1) Or Exists
                              (Select 1
                               From 病人医嘱发送
                               Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And Nvl(执行状态, 0) = 1)) And Not Exists
                        (Select 1
                              From 病人医嘱附费
                              Where a.医嘱序号 = 医嘱id And a.No = NO And Mod(a.记录性质, 10) = 记录性质) And a.No = No_In And
                             a.序号 = r_Bill.序号 And a.记录性质 = 1 And a.记录状态 In (1, 3) 　and a.价格父号 Is Null) Q1
                Where Not Exists (Select 1
                       From 药品收发记录
                       Where 费用id = Q1.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0) Having
                 Max(ID) <> 0;
              Else
              
                Select Nvl(Sum(数量), 0), Count(*)
                Into n_准退数量, n_Count
                From (Select a.医嘱id, a.收费细目id, Nvl(a.数量, 1) * Nvl(b.发送数次, 1) As 数量
                       From 病人医嘱计价 A, 病人医嘱发送 B, 门诊费用记录 J, 病人医嘱记录 M
                       Where a.医嘱id = b.医嘱id And a.医嘱id = m.Id And Nvl(b.执行状态, 0) <> 1 And a.医嘱id = j.医嘱序号 And
                             a.收费细目id = j.收费细目id And j.No = No_In And j.记录性质 = 1 And j.序号 = r_Bill.序号 And
                             j.记录状态 In (1, 3) And j.价格父号 Is Null And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And
                             Exists
                        (Select 1
                              From 病人医嘱计价 A
                              Where a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And Nvl(a.收费方式, 0) = 0) And Not Exists
                        (Select 1
                              From 药品收发记录
                              Where 费用id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0)
                       Union All
                       Select a.医嘱id, a.收费细目id, -1 * Nvl(a.数量, 1) * Nvl(c.本次数次, 1) As 数量
                       From 病人医嘱计价 A, 病人医嘱发送 B, 病人医嘱执行 C, 门诊费用记录 J, 病人医嘱记录 M
                       Where a.医嘱id = b.医嘱id And b.医嘱id = c.医嘱id And b.发送号 = c.发送号 And a.医嘱id = m.Id And
                             Nvl(c.执行结果, 1) = 1 And Nvl(b.执行状态, 0) <> 1 And a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And
                             j.No = No_In And j.记录性质 = 1 And Nvl(a.收费方式, 0) = 0 And j.序号 = r_Bill.序号 And j.记录状态 In (1, 3) And
                             j.价格父号 Is Null And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Not Exists
                        (Select 1
                              From 药品收发记录
                              Where 费用id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0) And Not Exists
                        (Select 1 From 材料特性 Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)
                       Union All
                       Select a.医嘱序号 As 医嘱id, a.收费细目id, Nvl(a.付数, 1) * a.数次 As 数量
                       From 门诊费用记录 A, 病人医嘱记录 M
                       Where a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And a.No = No_In And
                             a.记录性质 = 1 And a.序号 = r_Bill.序号 And a.记录状态 = 2 And a.价格父号 Is Null And Not Exists
                        (Select 1 From 药品收发记录 Where NO = No_In And 单据 In (8, 24) And 药品id = a.收费细目id));
              End If;
            End If;
            If Nvl(n_Count, 0) <> 0 And n_准退数量 = 0 Then
              v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已执行,不允许退费！';
              Raise Err_Item;
            End If;
          
            If Nvl(n_Count, 0) = 0 Then
              n_准退数量 := n_剩余数量;
            End If;
          
          Else
            Select Nvl(Sum(Nvl(付数, 1) * 实际数量), 0), Count(*)
            Into n_准退数量, n_Count
            From 药品收发记录
            Where NO = No_In And 单据 In (8, 24) And Mod(记录状态, 3) = 1 --@@@
                  And 审核人 Is Null And 费用id = r_Bill.Id;
          
            --有剩余数量无准退数量的有两种情况：
            --1.不跟踪在用的卫材无对应的收发记录,这时使用剩余数量
            --2.并发操作,此时已发药或发料
            If n_准退数量 = 0 Then
              If r_Bill.收费类别 = '4' Then
                If n_Count > 0 Then
                  v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发料,须退料后再退费！';
                  Raise Err_Item;
                Else
                  n_准退数量 := n_剩余数量;
                End If;
              Else
                v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发药,须退药后再退费！';
                Raise Err_Item;
              End If;
            End If;
          End If;
        
          --是否部分退费
          If r_Bill.执行状态 = 2 Or n_准退数量 <> Nvl(r_Bill.付数, 1) * r_Bill.数次 Then
            n_正常退费 := 0;
          End If;
        
          --处理门诊费用记录
          n_费用状态 := 0;
          --该笔项目第几次退费
          If Nvl(校对标志_In, 0) <> 0 Then
            n_退费次数 := -9; --先标明,固定为9
            n_费用状态 := 1;
          Else
            Select Nvl(Max(Abs(执行状态)), 0) + 1
            Into n_退费次数
            From 门诊费用记录
            Where NO = No_In And 记录性质 = 1 And 记录状态 = 2 And Nvl(执行状态, 0) < 0 And 序号 = r_Bill.序号;
          End If;
        
          --金额=剩余金额*(准退数/剩余数)
          If Nvl(r_Bill.误差, 0) = 9 Then
            --误差可以超过设置的小数位(比如:医保结算超过小数位后,误差就可能超过小数位
            n_应收金额 := Round(n_剩余应收 * (n_准退数量 / n_剩余数量), 5);
            n_实收金额 := Round(n_剩余实收 * (n_准退数量 / n_剩余数量), 5);
            n_统筹金额 := Round(n_剩余统筹 * (n_准退数量 / n_剩余数量), 5);
          Else
            n_应收金额 := Round(n_剩余应收 * (n_准退数量 / n_剩余数量), n_Dec);
            n_实收金额 := Round(n_剩余实收 * (n_准退数量 / n_剩余数量), n_Dec);
            n_统筹金额 := Round(n_剩余统筹 * (n_准退数量 / n_剩余数量), n_Dec);
          End If;
          n_总金额 := n_总金额 + n_实收金额;
        
          --插入退费记录
          Insert Into 门诊费用记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别,
             收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人,
             执行状态, 费用状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 结论,
             缴款组id, 主页id, 病人病区id)
            Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别,
                   病人科室id, 收费类别, 收费细目id, 计算单位, Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, 付数, 1), 发药窗口,
                   Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, -1 * 数次, -1 * n_准退数量), 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价,
                   -1 * n_应收金额, -1 * n_实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, -1 * n_退费次数, n_费用状态, 执行时间, 操作员编号_In,
                   操作员姓名_In, 发生时间, d_Date, n_结帐id, -1 * n_实收金额, 保险项目否, 保险大类id, -1 * n_统筹金额, Nvl(退费摘要_In, 摘要),
                   Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, 结论, n_组id, 主页id, 病人病区id
            From 门诊费用记录
            Where ID = r_Bill.Id;
        
          --标记原费用记录
          --执行状态:全部退完(准退数=剩余数)标记为0,否则标记为1,异常收费单,还是标明9
          Update 门诊费用记录
          Set 记录状态 = 3, 执行状态 = Decode(Nvl(执行状态, 0), 9, 9, Decode(Sign(n_准退数量 - n_剩余数量), 0, 0, 1))
          Where ID = r_Bill.Id;
        End If;
      Else
        If 序号_In Is Not Null Then
          v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经完全执行,不能退费！';
          Raise Err_Item;
        End If;
        --情况:没限定行号,原始单据中包括已经完全执行的
        n_正常退费 := 0;
      End If;
    Else
      n_正常退费 := 0; --未指定该笔,属于部分退费
    End If;
  End Loop;
  ---------------------------------------------------------------------------------
  --处理病人预交记录

  --原单据的结帐ID
  Select 结帐id, 病人id
  Into n_原结帐id, n_病人id
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Rownum = 1;

  If n_正常退费 = 1 And Nvl(退款操作_In, 0) = 0 Then
    --单据第一次退费且全部退完
    --冲预交部分记录
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交, 结帐id,
       缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质, 会话号)
      Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_Date,
             操作员姓名_In, 操作员编号_In, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
             Decode(校对标志_In, 1, 2, 校对标志_In), n_结算序号, 3, n_会话号
      From 病人预交记录
      Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
    If Nvl(校对标志_In, 0) = 0 Then
      --处理病人预交余额
      For v_预交 In (Select 预交类别, Nvl(Sum(Nvl(冲预交, 0)), 0) As 预交金额, 病人id
                   From 病人预交记录
                   Where 记录性质 In (1, 11) And 结帐id = n_原结帐id
                   Group By 预交类别, 病人id
                   Having Sum(Nvl(冲预交, 0)) <> 0) Loop
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) + Nvl(v_预交.预交金额, 0)
        Where 病人id = v_预交.病人id And 性质 = 1 And 类型 = Nvl(v_预交.预交类别, 2)
        Returning 预交余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 病人余额
            (病人id, 类型, 预交余额, 性质)
          Values
            (v_预交.病人id, Nvl(v_预交.预交类别, 2), Nvl(v_预交.预交金额, 0), 1);
          n_返回值 := n_预交金额;
        End If;
        If n_返回值 = 0 Then
          Delete From 病人余额
          Where 病人id = v_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
        End If;
      End Loop;
    End If;
    --非医保全退,和医保所有结算方式都允许回退,原样退回(冲预交在前面已处理)
    If 医保结算方式_In Is Null Then
      v_结算内容 := ',' || Nvl(一卡通结算_In, '-Lxh') || ',' || Nvl(一卡通结算_In, 'Lxh') || ',';
    
      --一卡通或消费卡或银行卡的相关数据需要特殊处理,需要最后较对.
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
         卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质, 会话号)
        Select 病人预交记录_Id.Nextval, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, d_Date, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
               -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
               Case
                 When Nvl(卡类别id, 0) <> 0 Then
                  Decode(校对标志_In, 1, 1, 0) * 1
                 When Nvl(结算卡序号, 0) <> 0 Then
                  Decode(校对标志_In, 1, 1, 0) * 1
                 When Nvl(q.预交id, 0) <> 0 Then
                  Decode(校对标志_In, 1, 1, 0) * 1
                 When Nvl(j.名称, '-') <> '-' Then
                  Decode(校对标志_In, 1, 1, 0)
                 Else
                  Decode(校对标志_In, 1, 2, 0)
               End As 校对标志, n_结算序号, 3, n_会话号
        From 病人预交记录 A, (Select 名称 From 结算方式 Where 性质 In (3, 4)) J,
             (Select m.Id As 预交id
               From 病人预交记录 M, 一卡通目录 C
               Where m.结帐id = n_原结帐id And m.结算方式 = c.结算方式 And m.记录性质 = 3 And m.记录状态 = 1) Q
        Where a.记录性质 = 3 And a.记录状态 = 1 And a.结帐id = n_原结帐id And a.Id = q.预交id(+) And a.结算方式 = j.名称(+) And
              Instr(v_结算内容, ',' || 结算方式 || ',') = 0 And
              (Not Exists (Select 1 From 病人卡结算记录 Where 结算id = a.Id) Or Nvl(a.结算卡序号, 0) = 0);
    
      --处理消费卡,结算卡在上面就已经处理了
      Select Count(1)
      Into n_Count
      From 病人预交记录 A, 病人卡结算记录 B
      Where a.Id = b.结算id And a.记录性质 = 3 And a.结帐id = n_原结帐id And Rownum < 2;
      If n_Count <> 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id,
           预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质, 会话号)
          Select n_预交id, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, d_Date, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
                 -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 2, n_结算序号, Mod(记录性质, 10), n_会话号
          From 病人预交记录 A
          Where a.记录性质 = 3 And a.结帐id = n_原结帐id And Exists
           (Select 1 From 病人卡结算记录 Where 结算id = a.Id) And Instr(Nvl(v_结算内容, '_LXH'), ',' || a.结算方式 || ',') = 0;
      
        --收费时可能使用了多张消费卡
        For c_记录 In (Select a.Id, c.接口编号, c.消费卡id, c.卡号, -1 * Sum(c.应收金额) As 结算金额
                     From 病人预交记录 A, 病人卡结算记录 C
                     Where a.Id = c.结算id And a.记录性质 = 3 And a.记录状态 In (1, 3) And a.结帐id = n_原结帐id And
                           Instr(Nvl(v_结算内容, '_LXH'), ',' || a.结算方式 || ',') = 0
                     Group By a.Id, c.接口编号, c.消费卡id, c.卡号) Loop
        
          Zl_病人卡结算记录_退款(c_记录.接口编号, c_记录.卡号, c_记录.消费卡id, c_记录.结算金额, c_记录.Id, n_预交id, 操作员编号_In, 操作员姓名_In, d_Date);
        End Loop;
      End If;
    
      --b.余下的就是三方接口支持的退现了,不允许作废的结算方式,加上到指定的结算方式上,加上误差(因为界面程序会在这之后退误差)
      If 一卡通结算_In Is Not Null Then
        Begin
          Select -1 * Nvl(Sum(冲预交), 0) Into n_已退金额 From 病人预交记录 Where 结帐id = n_结帐id;
        Exception
          When Others Then
            n_已退金额 := 0;
        End;
      
        If (n_总金额 - n_已退金额) <> 0 Then
          --此时的总金额还没有包含误差,因为界面程序中在调用本过程后才产生误差费用记录
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号,
             交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质, 会话号)
            Select 病人预交记录_Id.Nextval, 3, NO, 2, 病人id, 主页id, '门诊退费结算', v_退费结算, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * (n_总金额 - n_已退金额), n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
                   Decode(校对标志_In, 1, 2, 0), n_结算序号, 3, n_会话号
            From 病人预交记录
            Where 记录性质 = 3 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum = 1;
          n_部分退 := 1;
        End If;
      End If;
      --医保按允许作废的结算方式退,不允许的,退到指定的结算方式上
      --需要处理误差金额
    Else
      --a.原样退回
      v_结算内容 := ',' || 医保结算方式_In || ',' || Nvl(一卡通结算_In, '-Lxh') || ',' || v_退费结算 || ',';
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号,
         交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质, 会话号)
        Select 病人预交记录_Id.Nextval, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, d_Date, 操作员编号_In, 操作员姓名_In, -1 * 冲预交, n_结帐id,
               n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
               
               Case
                 When Nvl(卡类别id, 0) <> 0 Then
                  Decode(校对标志_In, 1, 1, 0) * 1
                 When Nvl(结算卡序号, 0) <> 0 Then
                  Decode(校对标志_In, 1, 1, 0) * 1
                 When Nvl(q.预交id, 0) <> 0 Then
                  Decode(校对标志_In, 1, 1, 0) * 1
                 When Nvl(j.名称, '-') <> '-' Then
                  Decode(校对标志_In, 1, 1, 0)
                 Else
                  Decode(校对标志_In, 1, 2, 0)
               End As 校对标志, n_结算序号, 3, n_会话号
        From 病人预交记录 A, (Select 名称 From 结算方式 Where 性质 In (3, 4)) J,
             (Select m.Id As 预交id
               From 病人预交记录 M, 一卡通目录 C
               Where m.结帐id = n_原结帐id And m.结算方式 = c.结算方式 And m.记录性质 = 3 And m.记录状态 = 1) Q
        Where a.记录性质 = 3 And a.记录状态 = 1 And a.结算方式 = j.名称(+) And a.结帐id = n_原结帐id And
              Instr(v_结算内容, ',' || a.结算方式 || ',') = 0 And a.Id = q.预交id(+) And
              (Not Exists (Select 1 From 病人卡结算记录 Where 结算id = a.Id) Or Nvl(a.结算卡序号, 0) = 0);
    
      --处理消费卡,结算卡在上面就已经处理了
      Select Count(1)
      Into n_Count
      From 病人预交记录 A, 病人卡结算记录 B
      Where a.Id = b.结算id And a.记录性质 = 3 And a.结帐id = n_原结帐id And Rownum < 2;
      If n_Count <> 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id,
           预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质, 会话号)
          Select n_预交id, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, d_Date, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
                 -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 2, n_结算序号, Mod(记录性质, 10), n_会话号
          From 病人预交记录 A
          Where a.记录性质 = 3 And a.结帐id = n_原结帐id And Exists
           (Select 1 From 病人卡结算记录 Where 结算id = a.Id) And Instr(Nvl(v_结算内容, '_LXH'), ',' || a.结算方式 || ',') = 0;
      
        --收费时可能使用了多张消费卡
        For c_记录 In (Select a.Id, c.接口编号, c.消费卡id, c.卡号, -1 * c.应收金额 As 结算金额
                     From 病人预交记录 A, 病人卡结算记录 C
                     Where a.Id = c.结算id And a.记录性质 = 3 And a.记录状态 In (1, 3) And a.结帐id = n_原结帐id And
                           Instr(Nvl(v_结算内容, '_LXH'), ',' || a.结算方式 || ',') = 0) Loop
        
          Zl_病人卡结算记录_退款(c_记录.接口编号, c_记录.卡号, c_记录.消费卡id, c_记录.结算金额, c_记录.Id, n_预交id, 操作员编号_In, 操作员姓名_In, d_Date);
        End Loop;
      End If;
    
      --b.余下的就是医保不允许作废的结算方式,加上到指定的结算方式上,加上误差(因为界面程序会在这之后退误差)
      Begin
        Select -1 * Nvl(Sum(冲预交), 0) Into n_已退金额 From 病人预交记录 Where 结帐id = n_结帐id;
      Exception
        When Others Then
          n_已退金额 := 0;
      End;
    
      If (n_总金额 - n_已退金额) <> 0 Then
        --此时的总金额还没有包含误差,因为界面程序中在调用本过程后才产生误差费用记录
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号,
           交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质, 会话号)
          Select 病人预交记录_Id.Nextval, 3, NO, 2, 病人id, 主页id, Decode(一卡通结算_In, Null, '门诊医保接口退费', '门诊医保接口和三方接口退费'), v_退费结算,
                 d_Date, 操作员编号_In, 操作员姓名_In, -1 * (n_总金额 - n_已退金额), n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
                 合作单位, Decode(校对标志_In, 1, 2, 0), n_结算序号, 3, n_会话号
          From 病人预交记录
          Where 记录性质 = 3 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum = 1;
        n_部分退 := 1;
      End If;
    
    End If;
  Else
    -------------------------------------------------
    --部分退费
    n_已退金额 := 0;
    --医疗卡部分退费时，传入:结算方式|金额
    If 一卡通结算_In Is Not Null Then
      If Instr(一卡通结算_In, '|') > 0 Then
        v_当前结算 := 一卡通结算_In;
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, '|') + 1);
        n_结算金额 := Nvl(To_Number(v_当前结算), 0);
        If Not Nvl(v_结算方式, 'TMP') = 'TMP' Then
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
             卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质, 会话号)
            Select 病人预交记录_Id.Nextval, 3, No_In, 2, 病人id, 主页id, '三方接口部分退费', v_结算方式, d_Date, Null, Null, Null, 操作员编号_In,
                   操作员姓名_In, -1 * (n_结算金额), n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
                   Decode(校对标志_In, 1, 1, 0), n_结算序号, 3, n_会话号
            From 病人预交记录
            Where 记录性质 = 3 And 记录状态 In (1, 3) And 结帐id = n_原结帐id And Rownum < 2;
        End If;
        n_已退金额 := n_结算金额;
      End If;
    End If;
    --其它直接退为指定结算方式
    If (n_总金额 - n_已退金额 + Nvl(误差_In, 0)) <> 0 Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
         结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质, 会话号)
        Select 病人预交记录_Id.Nextval, 3, No_In, 2, 病人id, 主页id, '部分退费结算', v_退费结算, d_Date, Null, Null, Null, 操作员编号_In,
               操作员姓名_In, -1 * (n_总金额 - n_已退金额 + Nvl(误差_In, 0)), n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
               Decode(校对标志, 1, 2, 0), n_结算序号, 3, n_会话号
        From 病人预交记录
        Where 记录性质 = 3 And 记录状态 In (1, 3) And 结帐id = n_原结帐id And Rownum = 1;
    End If;
  
    --如果收费时只使用了预交款,则要退预交,并且可能有多笔冲预交
    If Sql%RowCount = 0 And 一卡通结算_In Is Null Then
      n_预交金额 := n_总金额 - n_已退金额 + Nvl(误差_In, 0);
    
      For r_Deposit In c_Deposit(n_原结帐id) Loop
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
           结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质, 会话号)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                 d_Date, 操作员姓名_In, 操作员编号_In, Decode(Sign(r_Deposit.金额 - n_预交金额), -1, -1 * r_Deposit.金额, -1 * n_预交金额),
                 n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, Decode(校对标志_In, 1, 2, 0), n_结算序号, 3, n_会话号
          From 病人预交记录
          Where ID = r_Deposit.Id;
      
        If Nvl(校对标志_In, 0) = 0 Then
          --更新病人预交余额
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) + n_总金额 + Nvl(误差_In, 0)
          Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = 1
          Returning 预交余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 病人余额
              (病人id, 类型, 预交余额, 性质)
            Values
              (r_Deposit.病人id, 1, n_总金额 + Nvl(误差_In, 0), 1);
            n_返回值 := n_总金额 + Nvl(误差_In, 0);
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 病人余额
            Where 病人id = r_Deposit.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;
        End If;
      
        --检查是否已经处理完
        If r_Deposit.金额 < n_预交金额 Then
          n_预交金额 := n_预交金额 - r_Deposit.金额;
        Else
          n_预交金额 := 0;
        End If;
        If n_预交金额 = 0 Then
          Exit;
        End If;
      End Loop;
    End If;
  End If;

  --更新原记录
  Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 3 And 记录状态 In (1, 3) And 结帐id = n_原结帐id;

  If 多单据全退_In <> 1 Then
    --处理误差项,多单据全退时 ,按原样退无误差处理
    --将误差项的记录状态调整为3
    If Nvl(误差_In, 0) <> 0 Then
      n_Count := 1;
      If n_正常退费 = 1 And Nvl(退款操作_In, 0) = 0 Then
        n_原误差费 := 0;
        --原样退,但存在误差
        If n_部分退 = 0 Then
          Select -1 * Nvl(Sum(实收金额), 0)
          Into n_原误差费
          From 门诊费用记录 A
          Where NO = No_In And a.记录性质 = 1 And a.记录状态 In (1, 3) And Nvl(a.附加标志, 0) = 9;
        End If;
        If Nvl(n_原误差费, 0) <> 0 Or Nvl(误差_In, 0) <> 0 Then
          Update 病人预交记录
          Set 冲预交 = 冲预交 - n_原误差费 - Nvl(误差_In, 0)
          Where 结算方式 = v_退费结算 And 结帐id = n_结帐id;
          If Sql%NotFound Then
            Insert Into 病人预交记录
              (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
               卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质, 会话号)
              Select 病人预交记录_Id.Nextval, 3, No_In, 2, 病人id, 主页id, '误差费', v_退费结算, d_Date, Null, Null, Null, 操作员编号_In,
                     操作员姓名_In, -1 * n_原误差费 - Nvl(误差_In, 0), n_结帐id, n_组id, 预交类别, Null, Null, Null, Null, Null, Null, 0,
                     n_结算序号, 3, n_会话号
              From 病人预交记录
              Where 记录性质 = 3 And 记录状态 In (1, 3) And 结帐id = n_原结帐id And Rownum = 1;
          End If;
        End If;
      End If;
    Elsif n_正常退费 = 1 And Nvl(退款操作_In, 0) = 0 Then
      --原样退时,需要处理预交记录不足的情况
      Select Nvl(Sum(Nvl(结帐金额, 0)), 0) Into n_实收金额 From 门诊费用记录 Where 结帐id = n_结帐id;
      Select Nvl(Sum(Nvl(冲预交, 0)), 0) Into n_返回值 From 病人预交记录 Where 结帐id = n_结帐id;
      If Abs(n_实收金额) <> Abs(n_返回值) Then
        n_实收金额 := n_实收金额 - n_返回值;
        Update 病人预交记录 Set 冲预交 = 冲预交 + Nvl(n_实收金额, 0) Where 结算方式 = v_退费结算 And 结帐id = n_结帐id;
        If Sql%NotFound Then
          Insert Into 病人预交记录
            (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
             卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质, 会话号)
            Select 病人预交记录_Id.Nextval, 3, No_In, 2, 病人id, 主页id, '误差费', v_退费结算, d_Date, Null, Null, Null, 操作员编号_In,
                   操作员姓名_In, Nvl(n_实收金额, 0), n_结帐id, n_组id, 预交类别, Null, Null, Null, Null, Null, Null, 0, n_结算序号, 3,
                   n_会话号
            From 病人预交记录
            Where 记录性质 = 3 And 记录状态 In (1, 3) And 结帐id = n_原结帐id And Rownum = 1;
        End If;
      End If;
    End If;
  
    Select Nvl(Sum(Nvl(结帐金额, 0)), 0) Into n_实收金额 From 门诊费用记录 Where 结帐id = n_结帐id;
    Select Nvl(Sum(Nvl(冲预交, 0)), 0) Into n_返回值 From 病人预交记录 Where 结帐id = n_结帐id;
  
    n_实收金额 := n_实收金额 - n_返回值;
  
    If n_实收金额 <> 0 Then
      --未找到，新产生误差项
      Zl_门诊收费误差_Insert(No_In, n_实收金额, 1, 0);
    End If;
  End If;
  ---------------------------------------------------------------------------------
  --人员缴款余额(注意是预交记录处理后才处理，包括个人帐户等的结算金额,不含退冲预交款)
  --如果是需要校对的,暂不处理人员缴款余额
  If Nvl(校对标志_In, 0) = 0 Then
    For r_Moneyrow In c_Money(n_结帐id) Loop
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + r_Moneyrow.冲预交
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Moneyrow.结算方式
      Returning 余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (操作员姓名_In, r_Moneyrow.结算方式, 1, r_Moneyrow.冲预交);
        n_返回值 := r_Moneyrow.冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Moneyrow.结算方式 And Nvl(余额, 0) = 0;
      End If;
    End Loop;
  End If;

  ---------------------------------------------------------------------------------
  --退费票据回收(仅全退时才回退,部分退是在重打过程中回收)
  If 回收票据_In = 0 Then
  
    --启用标志||NO;执行科室(条数);收据费目(首页汇总,条数);收费细目(条数)
    v_Para     := Nvl(zl_GetSysParameter('票据分配规则', 1121), '0||0;0;0,0;0');
    n_启用模式 := Zl_To_Number(Substr(v_Para, 1, 1));
    If n_启用模式 <> 0 Then
      --收回票据
      Select 使用id
      Bulk Collect
      Into l_使用id
      From (Select Distinct b.使用id From 票据打印明细 B Where b.No = No_In And Nvl(b.票种, 0) = 1);
    
      n_启用模式 := l_使用id.Count;
      If l_使用id.Count <> 0 Then
        --插入回收记录
        Forall I In 1 .. l_使用id.Count
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, 操作员姓名_In, d_Date, 票据金额
            From 票据使用明细 A
            Where ID = l_使用id(I) And 性质 = 1 And Not Exists
             (Select 1 From 票据使用明细 Where 号码 = a.号码 And 票种 = a.票种 And Nvl(性质, 0) <> 1);
      
        Forall I In 1 .. l_使用id.Count
          Update 票据打印明细 Set 是否回收 = 1 Where 使用id = l_使用id(I) And Nvl(是否回收, 0) = 0;
      
      End If;
    End If;
    If n_启用模式 = 0 Then
      --获取单据最后一次的打印ID(可能是多张单据收费打印)
      Begin
        --性质=1，原因=6为退费打印票据(红票)，不回收
        Select ID
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 1 And b.No = No_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
      --可能以前没有打印,无收回
      If n_打印id Is Not Null Then
        --a.多张单据循环调用时只能收回一次
        Select Count(*) Into n_Count From 票据使用明细 Where 票种 = 1 And 性质 = 2 And 打印id = n_打印id;
        If n_Count = 0 Then
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In, 票据金额
            From 票据使用明细
            Where 打印id = n_打印id And 票种 = 1 And 性质 = 1;
        Else
          --b.部分退费多次收回时,最后一次全退收回要排开已收回的
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In, 票据金额
            From 票据使用明细 A
            Where 打印id = n_打印id And 票种 = 1 And 性质 = 1 And Not Exists
             (Select 1 From 票据使用明细 B Where a.号码 = b.号码 And 打印id = n_打印id And 票种 = 1 And 性质 = 2);
        End If;
      End If;
    End If;
  End If;

  ---------------------------------------------------------------------------------
  --药品卫材相关内容
  --必须按照“收费细目id”升序排序，防止并发锁“药品库存”表
  For r_Expenses In (Select ID
                     From 门诊费用记录
                     Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And 收费类别 In ('4', '5', '6', '7') And
                           (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null)
                     Order By 收费细目id) Loop
    Zl_药品收发记录_销售退费(r_Expenses.Id);
  End Loop;

  --医嘱处理
  --删除病人医嘱附费(最后一次删除时)
  For c_医嘱 In (Select Distinct 医嘱序号
               From 门诊费用记录
               Where NO = No_In And 记录性质 = 1 And 记录状态 = 3 And 医嘱序号 Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select 序号, Sum(数量) As 剩余数量
           From (Select 记录状态, 执行状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                  From 门诊费用记录
                  Where 记录性质 = 1 And Nvl(附加标志, 0) <> 9 And 医嘱序号 + 0 = c_医嘱.医嘱序号 And NO = No_In
                  Group By 记录状态, 执行状态, Nvl(价格父号, 序号))
           Group By 序号
           Having Sum(数量) <> 0);
  
    If n_Count = 0 Then
      Delete From 病人医嘱附费 Where 医嘱id = c_医嘱.医嘱序号 And 记录性质 = 1 And NO = No_In;
    End If;
  End Loop;

  --场合_In    Integer:=0, --0:门诊;1-住院
  --性质_In    Integer:=1, --1-收费单;2-记帐单
  --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
  --No_In      门诊费用记录.No%Type,
  --医嘱ids_In Varchar2 := Null
  Zl_医嘱发送_计费状态_Update(0, 1, 2, No_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊收费记录_Delete;
/

--139063:冉俊明,2019-04-03,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_病人费用销帐_Delete
(
  Ids_In    In Varchar2,
  配药id_In In 输液配药记录.Id%Type := Null
) As
  n_Id  病人费用销帐.费用id%Type;
  v_Ids Varchar2(4000);

  n_医嘱id   住院费用记录.Id%Type;
  v_No       住院费用记录.No%Type;
  v_操作人员 输液配药记录.操作人员%Type;
  d_操作时间 输液配药记录.操作时间%Type;
  n_操作类型 输液配药记录.操作状态%Type;
  n_销帐时间 输液配药记录.操作时间%Type;
Begin
  If 配药id_In Is Not Null Then
    Select 操作时间
    Into n_销帐时间
    From (Select 操作人员, 操作时间, 操作类型
           From 输液配药状态
           Where 配药id = 配药id_In And 操作类型 = 9
           Order By 操作时间 Desc)
    Where Rownum = 1;
  End If;

  v_Ids := Ids_In || ',';
  While v_Ids Is Not Null Loop
    n_Id  := To_Number(Substr(v_Ids, 1, Instr(v_Ids, ',') - 1));
    v_Ids := Substr(v_Ids, Instr(v_Ids, ',') + 1);
  
    If n_销帐时间 Is Null Then
      Delete 病人费用销帐 Where 费用id = n_Id And 状态 = 0;
    
      Select NO, 医嘱序号
      Into v_No, n_医嘱id
      From (Select a.No, a.医嘱序号
             From 住院费用记录 A
             Where a.Id = n_Id
             Union All
             Select a.No, a.医嘱序号
             From 门诊费用记录 A
             Where a.Id = n_Id);
      If Not n_医嘱id Is Null Then
        --暂未提供按配药批次取消的功能，所有已申请的批次一起取消
        For R In (Select d.Id
                  From 病人医嘱记录 A, 病人医嘱发送 B, 输液配药记录 D
                  Where a.Id = n_医嘱id And a.Id = b.医嘱id And b.No = v_No And a.相关id = d.医嘱id And b.发送号 = d.发送号 And
                        b.记录性质 = 2) Loop
          Select 操作人员, 操作时间, 操作类型
          Into v_操作人员, d_操作时间, n_操作类型
          From (Select 操作人员, 操作时间, 操作类型
                 From 输液配药状态
                 Where 配药id = r.Id And 操作类型 <> 9
                 Order By 操作时间 Desc, 操作类型 Desc)
          Where Rownum = 1;
          Update 输液配药记录 Set 操作人员 = v_操作人员, 操作时间 = d_操作时间, 操作状态 = n_操作类型 Where ID = r.Id;
        End Loop;
      End If;
    Else
      Delete 病人费用销帐 Where 费用id = n_Id And 状态 = 0 And 申请时间 = n_销帐时间;
      Select 操作人员, 操作时间, 操作类型
      Into v_操作人员, d_操作时间, n_操作类型
      From (Select 操作人员, 操作时间, 操作类型
             From 输液配药状态
             Where 配药id = 配药id_In And 操作类型 <> 9
             Order By 操作时间 Desc, 操作类型 Desc)
      Where Rownum = 1;
      Update 输液配药记录 Set 操作人员 = v_操作人员, 操作时间 = d_操作时间, 操作状态 = n_操作类型 Where ID = 配药id_In;
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人费用销帐_Delete;
/

--120692:陈刘,2019-04-03,护理记录支持检验项目导入
Create Or Replace Procedure Zl_护理内容导入定义_Update
(
  类别_In 护理内容导入定义.类别%Type,
  名称_In 护理内容导入定义.名称%Type,
  格式_In 护理内容导入定义.格式%Type
) Is
Begin
  Update 护理内容导入定义 Set 名称 = 名称_In, 格式 = 格式_In Where 类别 = 类别_In;
  If Sql%Rowcount = 0 Then
    Insert Into 护理内容导入定义 (类别, 名称, 格式) Values (类别_In, 名称_In, 格式_In);
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_护理内容导入定义_Update;
/

--139063:冉俊明,2019-04-01,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_病人费用销帐_Audit
(
  Id_In       病人费用销帐.费用id%Type,
  申请时间_In 病人费用销帐.申请时间%Type,
  审核人_In   病人费用销帐.审核人%Type,
  审核时间_In 病人费用销帐.审核时间%Type,
  状态_In     病人费用销帐.状态%Type,
  Int自动退料 Integer := 1,
  申请类别_In 病人费用销帐.申请类别%Type := 1
) As
  --功能：审核或取消审核销账申请
  --入参：
  --    状态_In 1-审核通过,2-审核未通过
  --    申请类别_In 对药品和卫材有效,缺省为已执行的药品或卫材 
  --说明：
  --    费用可能来自于住院费用记录，也可能来自于门诊费用记录
  n_执行状态       住院费用记录.执行状态%Type;
  n_申请类别       病人费用销帐.申请类别%Type;
  v_收费类别       住院费用记录.收费类别%Type;
  v_No             住院费用记录.No%Type;
  n_实际数量       药品收发记录.实际数量%Type;
  n_数量           病人费用销帐.数量%Type;
  n_收发id         药品收发记录.Id%Type;
  n_医嘱id         住院费用记录.医嘱序号%Type;
  v_跟踪在用       材料特性.跟踪在用%Type;
  n_收费细目id     住院费用记录.收费细目id%Type;
  n_审核部门id     病人费用销帐.审核部门id%Type;
  n_执行部门id     住院费用记录.执行部门id%Type;
  n_病人id         住院费用记录.病人id%Type;
  n_主页id         住院费用记录.主页id%Type;
  n_审核标志       病案主页.审核标志%Type;
  n_住院状态       病案主页.状态%Type;
  n_病人审核方式   Number(2);
  n_未入科禁止记账 Number(2);

  n_Count   Number(18);
  n_Temp    Number(18);
  v_Err_Msg Varchar2(300);
  Err_Item Exception;
Begin
  n_申请类别 := 0;
  Select a.执行状态, a.收费类别, a.收费细目id, a.执行部门id, a.No, Nvl(b.跟踪在用, 0), a.医嘱序号, 病人id, 主页id
  Into n_执行状态, v_收费类别, n_收费细目id, n_执行部门id, v_No, v_跟踪在用, n_医嘱id, n_病人id, n_主页id
  From (Select 收费类别, NO, 收费细目id, 病人病区id, 执行部门id, 医嘱序号, 病人id, 主页id, 执行状态
         From 住院费用记录
         Where ID = Id_In
         Union All
         Select 收费类别, NO, 收费细目id, 病人病区id, 执行部门id, 医嘱序号, 病人id, 主页id, 执行状态
         From 门诊费用记录
         Where ID = Id_In) A, 材料特性 B
  Where a.收费细目id = b.材料id(+);

  If Nvl(n_主页id, 0) <> 0 Then
    n_病人审核方式   := Nvl(zl_GetSysParameter(185), 0);
    n_未入科禁止记账 := Nvl(zl_GetSysParameter(215), 0);
    If n_病人审核方式 = 1 Or n_未入科禁止记账 = 1 Then
      Begin
        Select 审核标志, 状态
        Into n_审核标志, n_住院状态
        From 病案主页
        Where 病人id = Nvl(n_病人id, 0) And 主页id = Nvl(n_主页id, 0);
      Exception
        When Others Then
          n_审核标志 := 0;
          n_住院状态 := 0;
      End;
      If n_未入科禁止记账 = 1 And n_住院状态 = 1 Then
        v_Err_Msg := '病人未入科,禁止对病人相关费用的操作!';
        Raise Err_Item;
      End If;
    
      If n_病人审核方式 = 1 Then
        If Nvl(n_审核标志, 0) = 1 Then
          v_Err_Msg := '该病人目前正在审核费用,不能进行费用相关调整!';
          Raise Err_Item;
        End If;
        If Nvl(n_审核标志, 0) = 2 Then
          v_Err_Msg := '该病人目前已经完成了费用审核,不能进行费用相关调整!';
          Raise Err_Item;
        End If;
      End If;
    End If;
  End If;
  If Instr('567', v_收费类别) > 0 Or (v_收费类别 = '4' And Nvl(v_跟踪在用, 0) = 1) Then
    n_申请类别 := 申请类别_In;
  End If;

  Update 病人费用销帐
  Set 审核人 = 审核人_In, 审核时间 = 审核时间_In, 状态 = 状态_In
  Where 费用id = Id_In And 申请类别 = n_申请类别 And 申请时间 = 申请时间_In And 状态 = 0
  Returning 数量, 审核部门id Into n_数量, n_审核部门id;
  If Sql%RowCount = 0 Then
    v_Err_Msg := '销帐审核失败,当前操作的记录可能因为并发操作已经被他人处理,请先刷新信息!';
    Raise Err_Item;
  End If;

  If n_申请类别 = 0 And (Instr(',5,6,7', ',' || v_收费类别) > 0 Or (v_收费类别 = '4' And Nvl(v_跟踪在用, 0) = 1)) Then
    --需要检查未执行的数量必须全部申请,才会通过 
    Select Sum(Nvl(付数, 0) * Nvl(实际数量, 0))
    Into n_实际数量
    From 药品收发记录
    Where 审核日期 Is Null And 费用id = Id_In And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0;
    If Nvl(n_实际数量, 0) < Nvl(n_数量, 0) Then
      Select '在单据号<<' || v_No || '>>中' || Decode(v_收费类别, '4', '卫材', '药品') || '为:' || Chr(13) || 编码 || '-' || 名称 ||
              Chr(13) || '的申请数量(' || LTrim(To_Char(n_数量, '9999999990.99')) || ')大于了待发' || Decode(v_收费类别, '4', '料', '药') ||
              '数量(' || LTrim(To_Char(Nvl(n_实际数量, 0), '9999999990.99')) || '),不允许审核!'
      Into v_Err_Msg
      From 收费项目目录
      Where ID = n_收费细目id;
      Raise Err_Item;
    End If;
  
    If n_医嘱id <> 0 Then
      Select Nvl(Max(d.Id), 0)
      Into n_Count
      From 病人医嘱记录 A, 病人医嘱发送 B, 输液配药记录 D
      Where a.Id = n_医嘱id And a.Id = b.医嘱id And b.No = v_No And a.相关id = d.医嘱id And b.发送号 = d.发送号 And b.记录性质 = 2 And
            d.操作时间 = 申请时间_In And d.操作状态 = 9;
    
      If n_Count <> 0 Then
        Select Count(1)
        Into n_Temp
        From 输液配药状态
        Where 配药id = n_Count And 操作类型 = 10 And 操作时间 = 审核时间_In;
        If n_Temp = 0 Then
          Insert Into 输液配药状态 (配药id, 操作类型, 操作人员, 操作时间) Values (n_Count, 10, 审核人_In, 审核时间_In);
        End If;
        Update 输液配药记录 Set 操作人员 = 审核人_In, 操作时间 = 审核时间_In, 操作状态 = 10 Where ID = n_Count;
      End If;
    End If;
  End If;

  If n_执行状态 <> 0 Then
    If Instr(',5,6,7,', ',' || v_收费类别 || ',') > 0 And n_申请类别 = 1 Then
      If n_执行部门id <> n_审核部门id Then
        Begin
          Select '[' || 编码 || ']' || 名称 Into v_Err_Msg From 收费项目目录 Where ID = n_收费细目id;
        Exception
          When Others Then
            v_Err_Msg := '';
        End;
        v_Err_Msg := '在销帐审核时,药品为' || v_Err_Msg || ' 的已经被执行科室执行,不能再进行销帐审核,请取消审核!';
        Raise Err_Item;
      End If;
    End If;
  
    If v_收费类别 = '4' Then
      If v_跟踪在用 = 1 Then
        If n_执行部门id <> n_审核部门id And n_申请类别 = 1 And Int自动退料 <> 1 Then
          Begin
            Select '[' || 编码 || ']' || 名称 Into v_Err_Msg From 收费项目目录 Where ID = n_收费细目id;
          Exception
            When Others Then
              v_Err_Msg := '';
          End;
          v_Err_Msg := '在销帐审核时,卫材为' || v_Err_Msg || ' 的已经被执行科室执行,不能再进行销帐审核,请取消审核!';
          Raise Err_Item;
        End If;
      
        If n_申请类别 = 1 And Int自动退料 = 1 Then
          n_收发id := -1;
          --可能来自于多个批次 
          For c_收发记录 In (Select ID, 批号, Nvl(Sum(Nvl(付数, 1) * 实际数量), 0) As 数次
                         From 药品收发记录
                         Where 费用id = Id_In And 单据 In (25, 26) And (记录状态 = 1 Or Mod(记录状态, 3) = 0)
                         Group By ID, 批号) Loop
            n_收发id := c_收发记录.Id;
            If n_数量 = 0 Then
              Exit;
            End If;
          
            If n_数量 > c_收发记录.数次 Then
              n_Temp := c_收发记录.数次;
              n_数量 := n_数量 - c_收发记录.数次;
            Else
              n_Temp := n_数量;
              n_数量 := 0;
            End If;
            Zl_材料收发记录_部门退料(c_收发记录.Id, 审核人_In, 审核时间_In, c_收发记录.批号, Null, Null, n_Temp, 0);
          End Loop;
          If n_收发id = -1 Then
            v_Err_Msg := '在销帐审核时,卫材为' || v_Err_Msg || ' 的未找到相关的药品收发信息,可能是因为中途' || Chr(13) ||
                         '更改了卫材的跟踪属性,不能再进行销帐审核,请取消审核!';
            Raise Err_Item;
          End If;
        End If;
      Else
        --不是跟踪的卫材 
        Update 住院费用记录 Set 执行状态 = 0 Where ID = Id_In;
        If Sql%NotFound Then
          Update 门诊费用记录 Set 执行状态 = 0 Where ID = Id_In;
        End If;
      End If;
    Elsif Instr(',5,6,7,', ',' || v_收费类别 || ',') = 0 Then
      --可能存在部分消帐,所以先将非药品的处理成部分执行,再在销帐审核过程(ZL_住院记帐记录_Delete)中处理,处理规则如下: 
      --在调用本过程时: 
      --   1.如果是已经执行的,则改为部分执行(执行状态=2);再在销帐过程中处理这部分数据(ZL_住院记帐记录_Delete):即:如果执行状态=2,并且部分销帐的,则改为1(已执行) 
      --      原因是因为非药品类只能存在两种状态.已执行;2-未执行 
      --   2.如果是未执行的,则执行状态还是为0,而在销帐过程中记录状态保持不变 
    
      --非药品由于没有取消执行的操作,所以对已执行的要先改状态才能调销帐 
      Update 住院费用记录 Set 执行状态 = Decode(Nvl(执行状态, 0), 0, 0, 2) Where ID = Id_In;
      If Sql%NotFound Then
        Update 门诊费用记录 Set 执行状态 = Decode(Nvl(执行状态, 0), 0, 0, 2) Where ID = Id_In;
      End If;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人费用销帐_Audit;
/

--139063:冉俊明,2019-04-01,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_病人费用销帐_Insert
(
  Id_In         In 病人费用销帐.费用id%Type,
  收费细目id_In In 病人费用销帐.收费细目id%Type,
  申请部门id_In In 病人费用销帐.申请部门id%Type,
  数量_In       In 病人费用销帐.数量%Type,
  申请人_In     In 病人费用销帐.申请人%Type,
  申请时间_In   In 病人费用销帐.申请时间%Type,
  申请类别_In   In 病人费用销帐.申请类别%Type,
  删除标志_In   In Integer := 0,
  配药id_In     In Integer := 0,
  销帐原因_In   In 病人费用销帐.销帐原因%Type := Null,
  配液更新_In   In Number := 1
) As
  --功能：费用销账申请
  --入参：
  --     申请类别_In 对药品和卫材有效:0-未发药(料);1-已发药(料);其他为0
  --     删除标志_In 删除病人费用销帐时的条件:1-删除时不管申请类别,0-删除时,根据申请类别来进行删除(因为可能出现在申请销帐时,存在已执行和未执行两种状态)
  --     配液更新_In 是否 输液配药记录 状态字段。1-要更新，0-不更新
  --说明：
  --    费用可能来自于住院费用记录，也可能来自于门诊费用记录
  n_审核部门id   病人费用销帐.审核部门id%Type;
  n_申请类别     病人费用销帐.申请类别%Type;
  n_开单科室病区 病人费用销帐.审核部门id%Type;
  n_执行科室病区 病人费用销帐.审核部门id%Type;
  n_跟踪在用     材料特性.跟踪在用%Type;
  n_执行状态     住院费用记录.执行状态%Type;
  v_收费类别     住院费用记录.收费类别%Type;
  n_实际数量     药品收发记录.实际数量%Type;
  n_医嘱id       住院费用记录.医嘱序号%Type;
  n_主页id       住院费用记录.Id%Type;
  v_No           住院费用记录.No%Type;
  n_病人id       住院费用记录.病人id%Type;
  n_病人科室id   住院费用记录.病人科室id%Type;
  n_Icu科室id    住院费用记录.病人科室id%Type;
  n_已申请数量   药品收发记录.实际数量%Type;

  n_Temp    Number;
  n_Icu     Number;
  n_Count   Number;
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  n_审核标志       病案主页.审核标志%Type;
  n_住院状态       病案主页.状态%Type;
  n_病人审核方式   Number(2);
  n_未入科禁止记账 Number(2);

Begin
  Select Count(1)
  Into n_Count
  From (Select 1
         From 住院费用记录 A, 住院费用记录 B
         Where a.No = b.No And Mod(a.记录性质, 10) = Mod(b.记录性质, 10) And a.序号 = b.序号 And b.Id = Id_In Having
          Nvl(Sum(a.结帐金额), 0) <> 0
         Union All
         Select 1
         From 门诊费用记录 A, 门诊费用记录 B
         Where a.No = b.No And Mod(a.记录性质, 10) = Mod(b.记录性质, 10) And a.序号 = b.序号 And b.Id = Id_In Having
          Nvl(Sum(a.结帐金额), 0) <> 0);
  If Nvl(n_Count, 0) > 0 Then
    v_Err_Msg := '申请销帐的记录已被他人结帐';
    Raise Err_Item;
  End If;

  Select a.收费类别, a.No, Nvl(b.跟踪在用, 0), Decode(Nvl(申请类别_In, 0), 0, a.病人病区id, a.执行部门id), 医嘱序号, 病人id, Nvl(主页id, 0)
  Into v_收费类别, v_No, n_跟踪在用, n_审核部门id, n_医嘱id, n_病人id, n_主页id
  From (Select 收费类别, NO, 收费细目id, 病人病区id, 执行部门id, 医嘱序号, 病人id, 主页id
         From 住院费用记录
         Where ID = Id_In
         Union All
         Select 收费类别, NO, 收费细目id, 病人病区id, 执行部门id, 医嘱序号, 病人id, 主页id
         From 门诊费用记录
         Where ID = Id_In) A, 材料特性 B
  Where a.收费细目id = b.材料id(+);

  If Nvl(n_主页id, 0) <> 0 Then
    n_病人审核方式   := Nvl(zl_GetSysParameter(185), 0);
    n_未入科禁止记账 := Nvl(zl_GetSysParameter(215), 0);
    If n_病人审核方式 = 1 Or n_未入科禁止记账 = 1 Then
      Begin
        Select 审核标志, 状态
        Into n_审核标志, n_住院状态
        From 病案主页
        Where 病人id = Nvl(n_病人id, 0) And 主页id = Nvl(n_主页id, 0);
      Exception
        When Others Then
          n_审核标志 := 0;
          n_住院状态 := 0;
      End;
      If n_未入科禁止记账 = 1 And n_住院状态 = 1 Then
        v_Err_Msg := '病人未入科,禁止对病人相关费用的操作!';
        Raise Err_Item;
      End If;
    
      If n_病人审核方式 = 1 Then
        If Nvl(n_审核标志, 0) = 1 Then
          v_Err_Msg := '该病人目前正在审核费用,不能进行费用相关调整!';
          Raise Err_Item;
        End If;
        If Nvl(n_审核标志, 0) = 2 Then
          v_Err_Msg := '该病人目前已经完成了费用审核,不能进行费用相关调整!';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
  End If;

  Select Max(出院科室id)
  Into n_病人科室id
  From 病人信息 A, 病案主页 B
  Where a.病人id = b.病人id And a.主页id = b.主页id And a.病人id = n_病人id;

  Select Decode(Count(1), 0, 0, 1) Into n_Icu From 部门性质说明 B Where b.部门id = n_病人科室id And b.工作性质 = 'ICU';
  If n_Icu = 1 Then
    --检查是否是当前操作员属性于ICU
    Select Decode(Count(Distinct a.用户名), 0, 0, 1)
    Into n_Icu
    From 上机人员表 A, 部门性质说明 B, 部门人员 C
    Where a.用户名 = User And a.人员id = c.人员id And c.部门id = b.部门id And b.工作性质 = 'ICU';
  End If;

  If n_Icu = 1 Then
    n_Icu科室id := n_病人科室id;
    If Nvl(申请类别_In, 0) = 0 Then
      n_审核部门id := n_病人科室id;
    End If;
  End If;

  If Instr(',5,6,7', ',' || v_收费类别) > 0 Or v_收费类别 = '4' And Nvl(n_跟踪在用, 0) = 1 Then
    n_申请类别 := 申请类别_In;
  Else
    n_申请类别 := 0;
  End If;

  --取消以前申请的重新生成(按批次申请时，不能取消，因为费用id相同，每个批次可分别申请)
  If 配药id_In = 0 Then
    If Nvl(删除标志_In, 0) = 1 Then
      Delete 病人费用销帐 Where 费用id = Id_In And 状态 = 0;
    Else
      Delete 病人费用销帐 Where 费用id = Id_In And 申请类别 = n_申请类别 And 状态 = 0;
    End If;
  End If;
  If 数量_In <> 0 Then
    --审核科室
    --1.药品费用或跟踪在用的卫材:
    --    a. 如果未执行,则按病人病区作为审核部门;
    --    b. 如果已执行,则按执行部门ID作为审核部门
    --2.医技科室开单的费用(即开单科室<>病人科室)，销帐审核科室为开单科室,
    --  如果开单科室是属于病区的临床科室,则销帐科室为所属病区(即护士记病人在其它科室发生的费用)
    --  (如果执行科室x属于a、b两病区，则a、b两病区都可以作为销帐确认科室,取第一个，如果a病区同时是病人病区，则只有a病区能确认)。
    --3.病区产生的费用,没有经过划价审核的,销帐审核科室为病人病区(如果已经被执行，则为执行部门,否则为病人病区)
    --  经过划价审核的,销帐审核科室为执行科室。
    --  如果执行科室是属于病区的临床科室，则销帐审核科室为所属病区
    --  (如果执行科室x属于a、b两病区，则a、b两病区都可以作为销帐确认科室,取第一个，如果a病区同时是病人病区，则只有a病区能确认)。
    --4.如果当前操作员是属于ICU,并且病人当前科室也为ICU以及未执行的项目,由ICU科室来来进行审核.
  
    If Nvl(n_跟踪在用, 0) = 1 Then
      If Nvl(申请类别_In, 0) = 0 Then
        --要检查未执行的数量必须大于等于申请数量,才会通过
        Select Sum(Nvl(付数, 0) * Nvl(实际数量, 0))
        Into n_实际数量
        From 药品收发记录
        Where 审核日期 Is Null And 费用id = Id_In And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0;
      
        If Nvl(n_实际数量, 0) < Nvl(数量_In, 0) Then
          Select '在单据号<<' || v_No || '>>中卫材料为:' || Chr(13) || 编码 || '-' || 名称 || Chr(13) || '的申请数量(' ||
                  LTrim(To_Char(数量_In, '9999999990.99')) || ')大于了待发料数量(' ||
                  LTrim(To_Char(Nvl(n_实际数量, 0), '9999999990.99')) || '),不能进行申请销帐!'
          Into v_Err_Msg
          From 收费项目目录
          Where ID = 收费细目id_In;
          Raise Err_Item;
        End If;
      End If;
    Else
      --a.执行科室或其所属病区:：0:未执行;1:完全执行;2:部份执行
      Select Decode(b.病区id, Null, a.执行部门id, Nvl(c.病区id, b.病区id)), Decode(a.执行状态, 1, 1, 2, 1, 0)
      Into n_执行科室病区, n_执行状态
      From (Select 执行部门id, 病人病区id, 执行状态
             From 住院费用记录
             Where ID = Id_In
             Union All
             Select 执行部门id, 病人病区id, 执行状态
             From 门诊费用记录
             Where ID = Id_In) A, 病区科室对应 B, 病区科室对应 C
      Where a.执行部门id = b.科室id(+) And a.执行部门id = c.科室id(+) And a.病人病区id = c.病区id(+) And Rownum < 2;
    
      --b.开单科室或其所属病区
      Select Decode(b.病区id, Null, a.开单部门id, Nvl(c.病区id, b.病区id))
      Into n_开单科室病区
      From (Select 开单部门id, 病人病区id
             From 住院费用记录
             Where ID = Id_In
             Union All
             Select 开单部门id, 病人病区id
             From 门诊费用记录
             Where ID = Id_In) A, 病区科室对应 B, 病区科室对应 C
      Where a.开单部门id = b.科室id(+) And a.开单部门id = c.科室id(+) And a.病人病区id = c.病区id(+) And Rownum < 2;
    
      For v_费用 In (Select 收费类别, Nvl(执行状态, 0) As 执行状态, 病人病区id, 执行部门id, 开单部门id, 病人科室id, 划价人, 操作员姓名
                   From 住院费用记录
                   Where ID = Id_In
                   Union All
                   Select 收费类别, Nvl(执行状态, 0) As 执行状态, 病人病区id, 执行部门id, 开单部门id, 病人科室id, 划价人, 操作员姓名
                   From 门诊费用记录
                   Where ID = Id_In) Loop
      
        If Instr('567', v_费用.收费类别, 1) > 0 Then
          n_Temp       := Case
                            When 申请类别_In Is Null Then
                             Nvl(v_费用.执行状态, 0)
                            Else
                             申请类别_In
                          End;
          n_审核部门id := Case
                        When n_Temp = 0 Then
                         v_费用.病人病区id
                        Else
                         v_费用.执行部门id
                      End;
          If n_Temp = 0 And n_Icu = 1 Then
            --ICU为ICU科室
            n_审核部门id := n_Icu科室id;
          End If;
        Else
          If v_费用.开单部门id = v_费用.病人科室id Then
            --临床产生的费用
            If Nvl(v_费用.划价人, '-') = v_费用.操作员姓名 Or v_费用.划价人 Is Null Then
              --划价审核
              n_审核部门id := Case
                            When n_执行状态 = 1 Then
                             v_费用.执行部门id
                            Else
                             v_费用.病人病区id
                          End;
            Else
              n_审核部门id := n_执行科室病区;
            End If;
          Else
            n_审核部门id := n_执行科室病区;
          End If;
          If n_执行状态 = 0 And n_Icu = 1 Then
            --ICU为ICU科室
            n_审核部门id := n_Icu科室id;
          End If;
        End If;
      End Loop;
    
      If Instr(',5,6,7', ',' || v_收费类别) > 0 And Nvl(申请类别_In, Nvl(n_执行状态, 0)) = 0 Then
        --需要检查未执行的数量必须大于等于申请数量,才会通过
        Select Sum(Nvl(付数, 0) * Nvl(实际数量, 0))
        Into n_实际数量
        From 药品收发记录
        Where 审核日期 Is Null And 费用id = Id_In And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0;
        If Nvl(n_实际数量, 0) < Nvl(数量_In, 0) Then
          Select '在单据号<<' || v_No || '>>中药品为:' || Chr(13) || 编码 || '-' || 名称 || Chr(13) || '的申请数量(' ||
                  LTrim(To_Char(数量_In, '9999999990.99')) || ')不能大于待发药数量(' ||
                  LTrim(To_Char(Nvl(n_实际数量, 0), '9999999990.99')) || '),不能进行申请销帐!'
          Into v_Err_Msg
          From 收费项目目录
          Where ID = 收费细目id_In;
          Raise Err_Item;
        End If;
      End If;
    End If;
    --解决并发问题:当前申请数量+已经申请数量不能大于某笔申请数量
    Select Sum(Nvl(数量, 0)) Into n_已申请数量 From 病人费用销帐 Where 费用id = Id_In And Nvl(状态, 0) <> 2;
  
    Select Sum(Nvl(付数, 1) * Nvl(数次, 0))
    Into n_实际数量
    From (Select a.付数, a.数次
           From 住院费用记录 A, 住院费用记录 B
           Where a.No = b.No And Mod(a.记录性质, 10) = Mod(b.记录性质, 10) And a.记录状态 In (0, 1, 3) And a.序号 = b.序号 And
                 b.Id = Id_In
           Union All
           Select a.付数, a.数次
           From 门诊费用记录 A, 门诊费用记录 B
           Where a.No = b.No And Mod(a.记录性质, 10) = Mod(b.记录性质, 10) And a.记录状态 In (0, 1, 3) And a.序号 = b.序号 And
                 b.Id = Id_In);
  
    If Nvl(n_实际数量, 0) < Nvl(n_已申请数量, 0) + Nvl(数量_In, 0) Then
      Select '在单据号<<' || v_No || '>>收费项目:' || Chr(13) || 编码 || '-' || 名称 || Chr(13) || '的申请数量(' ||
              LTrim(To_Char(Nvl(n_已申请数量, 0) + Nvl(数量_In, 0), '9999999990.99')) || ')不能大于记帐数量(' ||
              LTrim(To_Char(Nvl(n_实际数量, 0), '9999999990.99')) || '),不能进行申请销帐!'
      Into v_Err_Msg
      From 收费项目目录
      Where ID = 收费细目id_In;
      Raise Err_Item;
    End If;
  
    If n_医嘱id <> 0 And 配药id_In <> 0 Then
      --如果是输液配药中心的，则更新相关表字段
      If Nvl(配液更新_In, 0) = 1 Then
        Select Count(1)
        Into n_Temp
        From 输液配药状态
        Where 配药id = 配药id_In And 操作类型 = 9 And 操作时间 = 申请时间_In;
        If n_Temp = 0 Then
          Insert Into 输液配药状态
            (配药id, 操作类型, 操作人员, 操作时间, 操作说明)
          Values
            (配药id_In, 9, 申请人_In, 申请时间_In, 销帐原因_In);
        End If;
        Update 输液配药记录
        Set 操作人员 = 申请人_In, 操作时间 = 申请时间_In, 操作状态 = 9
        Where ID = 配药id_In
        Returning 部门id Into n_审核部门id;
      Else
        Select 部门id Into n_审核部门id From 输液配药记录 Where ID = 配药id_In;
      End If;
    End If;
  
    Insert Into 病人费用销帐
      (费用id, 申请类别, 收费细目id, 审核部门id, 申请部门id, 数量, 申请人, 申请时间, 状态, 销帐原因)
    Values
      (Id_In, n_申请类别, 收费细目id_In, n_审核部门id, 申请部门id_In, 数量_In, 申请人_In, 申请时间_In, 0, 销帐原因_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人费用销帐_Insert;
/

--139063:冉俊明,2019-04-04,门诊留观病人按门诊流程就诊
CREATE OR REPLACE Procedure Zl_病人未结费用_Recalc
(
  病人id_In 住院费用记录.病人id%Type,
  主页id_In 住院费用记录.主页id%Type
) As
  v_费别     费别.名称%Type;
  n_病人性质 病案主页.病人性质%Type;

  v_No       住院费用记录.No%Type;
  n_实收金额 住院费用记录.实收金额%Type;
  n_费用余额 病人余额.费用余额%Type;
  n_小数位数 Number(2);
  v_Counter  Number(5);
  d_Sysdate  Date;
  v_Thisinfo Varchar(100);
  v_Lastinfo Varchar(100);

  Err_Custom Exception;
  v_Error Varchar2(255);
Begin
  Select 费别, 病人性质 Into v_费别, n_病人性质 From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;

  --条件判断 
  --a.当前不是按主从项汇总计算折扣模式 
  v_Counter := To_Number(Nvl(zl_GetSysParameter(93), 0));
  If v_Counter = 1 Then
    v_Error := '当前费别使用主从项汇总计算折扣模式,不支持费用重算!';
    Raise Err_Custom;
  End If;

  --b.当前费别不是使用药品按成本价加收打折的费别 
  v_Counter := 0;
  Select Count(费别) Into v_Counter From 费别明细 Where 费别 = v_费别 And 计算方法 = 1;
  If v_Counter > 0 Then
    v_Error := '当前费别使用药品按成本价加收打折模式,不支持费用重算!';
    Raise Err_Custom;
  End If;

  --门诊留观病人可能没有住院费用记录 
  If Nvl(n_病人性质, 0) <> 1 Then
    --c.没有未结费用 
    Begin
      Select 费用余额 Into n_费用余额 From 病人余额 Where 病人id = 病人id_In And 类型 = 2 And 性质 = 1;
    Exception
      When Others Then
        n_费用余额 := 0;
    End;
    --可能有未结费用，但不是本次住院发生的，在后面执行时再判断本次是否有未结明细 
    If n_费用余额 = 0 Then
      v_Counter := 0;
      --费用余额为0时，也可能有费用（所有费用都不收费） 
      Select Count(ID) Into v_Counter From 住院费用记录 Where 病人id = 病人id_In And 主页id = 主页id_In And Rownum < 2;
      If v_Counter = 0 Then
        v_Error := '病人不存在未结费用,不用进行费用重算!';
        Raise Err_Custom;
      End If;
    End If;
  End If;

  --d.不存在与本次住院费别不同的费用明细 
  v_Counter := 0;
  Select Count(ID)
  Into v_Counter
  From 住院费用记录
  Where 病人id = 病人id_In And 主页id = 主页id_In And 费别 <> v_费别 And Rownum < 2;
  If v_Counter = 0 And Nvl(n_病人性质, 0) <> 1 Then
    v_Error := '病人不存在与本次住院费别不同的费用明细 ,不用进行费用重算!';
    Raise Err_Custom;
  End If;

  --执行 
  If Nvl(v_Counter, 0) <> 0 Then
    v_Counter  := 0;
    d_Sysdate  := Sysdate;
    n_小数位数 := To_Number(Nvl(zl_GetSysParameter(9), 2));
    For r_Fee In (Select 病人id, 主页id, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 收费类别, 收费细目id, 计算单位, 加班标志, 附加标志, 婴儿费,
                         收入项目id, 收据费目, 开单部门id, 开单人, 执行部门id, 发生时间, 操作员编号, 操作员姓名, 医疗小组id, Nvl(Sum(应收金额), 0) 应收金额,
                         Nvl(Sum(实收金额), 0) 实收金额
                  From (Select ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别,
                                年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id,
                                收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号,
                                操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 医疗小组id
                         From 住院费用记录
                         Union All
                         Select ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别,
                                年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id,
                                收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号,
                                操作员姓名, 结帐id, 结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 医疗小组id
                         From H住院费用记录)
                  Where 病人id = 病人id_In And 主页id = 主页id_In And 记录状态 <> 0 And 记帐费用 = 1
                  Group By 病人id, 主页id, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 收费类别, 收费细目id, 计算单位, 加班标志, 附加标志,
                           婴儿费, 收入项目id, 收据费目, 开单部门id, 开单人, 执行部门id, 发生时间, 操作员编号, 操作员姓名, 医疗小组id
                  Having(Nvl(Sum(实收金额), 0) <> Nvl(Sum(结帐金额), 0) Or Nvl(Sum(结帐金额), 0) = 0) And Not(Nvl(Sum(应收金额), 0) = 0 And Nvl(Sum(实收金额), 0) = 0)
                  Order By 开单部门id, 开单人, 操作员姓名) Loop
      --          包括从未结的费用,费用明细部分结帐,以及结帐后作废,这些记录有可能已转入后备表 
      --          1.排开了已全部结帐的记录(Sum(应收金额)=Sum(应收金额)) 
      --          2.排开了无打折冲减的记帐后已销帐的记录(Sum(应收金额)=0,Sum(应收金额)=0) 
      --          3.不排开打折冲减后发生了单据销帐的记录，要将原冲减记录一并汇总重算(Sum(应收金额)=0,Sum(应收金额)<>0) 
      --          4.不排开打折冲减后产生的实收和结帐都为零的记录，因为改回原来的费别时，要重算回去 
      If r_Fee.应收金额 <> 0 Then
        Begin
          Select 实收金额
          Into n_实收金额
          From (Select Round(r_Fee.应收金额 * Nvl(实收比率, 0) / 100, n_小数位数) 实收金额
                 From 费别明细
                 Where 收费细目id = r_Fee.收费细目id And 费别 = v_费别 And Abs(r_Fee.应收金额) Between 应收段首值 And 应收段尾值 And
                       Nvl(计算方法, 0) = 0
                 Union All
                 Select Round(r_Fee.应收金额 * Nvl(实收比率, 0) / 100, n_小数位数) 实收金额
                 From 费别明细 A
                 Where 收入项目id = r_Fee.收入项目id And 费别 = v_费别 And Abs(r_Fee.应收金额) Between 应收段首值 And 应收段尾值 And
                       Nvl(计算方法, 0) = 0 And Not Exists
                  (Select 1 From 费别明细 B Where b.费别 = a.费别 And b.收费细目id = r_Fee.收费细目id));
        Exception
          When Others Then
            n_实收金额 := r_Fee.应收金额;
        End;
      Else
        n_实收金额 := 0;
      End If;
      --计算用来冲减原实收的差额 
      n_实收金额 := -1 * (r_Fee.实收金额 - n_实收金额);
    
      If n_实收金额 <> 0 Then
        --一张单据的开单部门id,开单人,操作员姓名,床号要求相同，如果其中之一变了则产生新单据，如果都没有变，一张单据最多100条明细 
        v_Thisinfo := r_Fee.开单部门id || r_Fee.开单人 || r_Fee.操作员姓名 || r_Fee.床号;
        If v_Counter = 0 Or v_Counter = 100 Or v_Thisinfo <> v_Lastinfo Then
          v_No       := Nextno(14);
          v_Counter  := 1;
          v_Lastinfo := v_Thisinfo;
        Else
          v_Counter := v_Counter + 1;
        End If;
      
        Insert Into 住院费用记录
          (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 门诊标志, 病人id, 主页id, 标识号, 床号, 姓名, 性别, 年龄, 病人病区id, 病人科室id, 费别,
           收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 付数, 数次, 发药窗口, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用,
           划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 结帐id, 结帐金额, 操作员编号, 操作员姓名, 摘要, 是否急诊, 医嘱序号, 医疗小组id)
        Values
          (病人费用记录_Id.Nextval, 2, v_No, 1, v_Counter, Null, Null, 0, Null, r_Fee.门诊标志, r_Fee.病人id, r_Fee.主页id, r_Fee.标识号,
           r_Fee.床号, r_Fee.姓名, r_Fee.性别, r_Fee.年龄, r_Fee.病人病区id, r_Fee.病人科室id, v_费别, r_Fee.收费类别, r_Fee.收费细目id,
           r_Fee.计算单位, Null, Null, 0, 0, Null, r_Fee.加班标志, r_Fee.附加标志, r_Fee.婴儿费, r_Fee.收入项目id, r_Fee.收据费目, 0, 0, n_实收金额,
           Null, 1, Null, r_Fee.开单部门id, r_Fee.开单人, r_Fee.发生时间, d_Sysdate, r_Fee.执行部门id, 0, Null, Null, r_Fee.操作员编号,
           r_Fee.操作员姓名, Decode(v_Counter, 1, '实收重算冲减', ''), 0, Null, r_Fee.医疗小组id);
      End If;
    End Loop;
  End If;

  If v_Counter = 0 Then
    If Nvl(n_病人性质, 0) <> 1 Then
      v_Error := '由于以下原因之一,没有进行费用重算:' || Chr(13) || Chr(13) || 'a.没有发现病人本次住院的未结费用.' || Chr(13) || 'b.所有未结费用已进行了费用重算.' ||
                 Chr(13) || 'c.按当前费别重算的实收冲减金额都为零.';
      Raise Err_Custom;
    End If;
  Else
    --病人余额 
    n_实收金额 := 0;
    Select Sum(实收金额)
    Into n_实收金额
    From 住院费用记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 2 And 登记时间 = d_Sysdate;
  
    Update 病人余额 Set 费用余额 = Nvl(费用余额, 0) + n_实收金额 Where 病人id = 病人id_In And 性质 = 1 And 类型 = 2;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 费用余额, 预交余额, 类型) Values (病人id_In, 1, n_实收金额, 0, 2);
    End If;
  
    --病人未结费用 
    For r_Fee In (Select 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Sum(实收金额) 实收金额
                  From 住院费用记录
                  Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 2 And 登记时间 = d_Sysdate
                  Group By 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id) Loop
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + r_Fee.实收金额
      Where 病人id = 病人id_In And Nvl(主页id, 0) = 主页id_In And Nvl(病人病区id, 0) = r_Fee.病人病区id And
            Nvl(病人科室id, 0) = r_Fee.病人科室id And Nvl(开单部门id, 0) = r_Fee.开单部门id And Nvl(执行部门id, 0) = r_Fee.执行部门id And
            收入项目id + 0 = r_Fee.收入项目id And 来源途径 + 0 = 2;
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, 主页id_In, r_Fee.病人病区id, r_Fee.病人科室id, r_Fee.开单部门id, r_Fee.执行部门id, r_Fee.收入项目id, 2, r_Fee.实收金额);
      End If;
    End Loop;
  End If;

  --门诊留观病人重算门诊费用
  If Nvl(n_病人性质, 0) = 1 Then
    Begin
      Zl_病人未结门诊费用_Recalc(病人id_In, 主页id_In);
    Exception
      When Others Then
        Null; --忽略
    End;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人未结费用_Recalc;
/

--139063:冉俊明,2019-04-01,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_病人未结门诊费用_Recalc
(
  病人id_In 住院费用记录.病人id%Type,
  主页id_In 病案主页.主页id%Type := Null
) As
  v_费别     费别.名称%Type;
  v_No       门诊费用记录.No%Type;
  n_实收金额 门诊费用记录.实收金额%Type;
  n_费用余额 病人余额.费用余额%Type;
  n_小数位数 Number(2);
  v_Counter  Number(5);
  d_Sysdate  Date;
  v_Thisinfo Varchar(100);
  v_Lastinfo Varchar(100);

  Err_Custom Exception;
  v_Error Varchar2(255);
Begin
  If Nvl(主页id_In, 0) = 0 Then
    Select 费别 Into v_费别 From 病人信息 Where 病人id = 病人id_In;
  Else
    Select Nvl(b.费别, a.费别)
    Into v_费别
    From 病人信息 A, 病案主页 B
    Where a.病人id = b.病人id And b.病人id = 病人id_In And b.主页id = 主页id_In;
  End If;

  --条件判断 
  --a.当前不是按主从项汇总计算折扣模式 
  v_Counter := To_Number(Nvl(zl_GetSysParameter(93), 0));
  If v_Counter = 1 Then
    v_Error := '当前费别使用主从项汇总计算折扣模式,不支持费用重算!';
    Raise Err_Custom;
  End If;

  --b.当前费别不是使用药品按成本价加收打折的费别 
  v_Counter := 0;
  Select Count(费别) Into v_Counter From 费别明细 Where 费别 = v_费别 And 计算方法 = 1;
  If v_Counter > 0 Then
    v_Error := '当前费别使用药品按成本价加收打折模式,不支持费用重算!';
    Raise Err_Custom;
  End If;

  --c.没有未结费用 
  Begin
    Select 费用余额 Into n_费用余额 From 病人余额 Where 病人id = 病人id_In And 类型 = 1 And 性质 = 1;
  Exception
    When Others Then
      n_费用余额 := 0;
  End;
  --可能有未结费用，但不是本次住院发生的，在后面执行时再判断本次是否有未结明细 
  If n_费用余额 = 0 Then
    v_Counter := 0;
    --费用余额为0时，也可能有费用（所有费用都不收费） 
    Select Count(ID) Into v_Counter From 门诊费用记录 Where 病人id = 病人id_In And Rownum < 2;
    If v_Counter = 0 Then
      v_Error := '病人不存在未结费用,不用进行费用重算!';
      Raise Err_Custom;
    End If;
  End If;

  --d.不存在与本次住院费别不同的费用明细 
  v_Counter := 0;
  Select Count(ID) Into v_Counter From 门诊费用记录 Where 病人id = 病人id_In And 费别 <> v_费别 And Rownum < 2;
  If v_Counter = 0 Then
    v_Error := '病人不存在与本次住院费别不同的费用明细 ,不用进行费用重算!';
    Raise Err_Custom;
  End If;

  --执行 
  v_Counter  := 0;
  d_Sysdate  := Sysdate;
  n_小数位数 := To_Number(Nvl(zl_GetSysParameter(9), 2));
  For r_Fee In (Select 病人id, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 病人病区id, 病人科室id, 收费类别, 收费细目id, 计算单位, 加班标志, 附加标志, 婴儿费, 收入项目id,
                       收据费目, 开单部门id, 开单人, 执行部门id, 发生时间, 操作员编号, 操作员姓名, Nvl(Sum(应收金额), 0) 应收金额, Nvl(Sum(实收金额), 0) 实收金额,
                       主页id, 挂号id
                From (Select ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号,
                              病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价,
                              应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id,
                              结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 主页id, 挂号id
                       From 门诊费用记录
                       Union All
                       Select ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号,
                              病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价,
                              应收金额, 实收金额, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id,
                              结帐金额, 保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊, 主页id, 挂号id
                       From H门诊费用记录)
                Where 病人id = 病人id_In And 记录状态 <> 0 And 记帐费用 = 1
                Group By 病人id, 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 病人病区id, 病人科室id, 收费类别, 收费细目id, 计算单位, 加班标志, 附加标志, 婴儿费, 收入项目id,
                         收据费目, 开单部门id, 开单人, 执行部门id, 发生时间, 操作员编号, 操作员姓名, 主页id, 挂号id
                Having(Nvl(Sum(实收金额), 0) <> Nvl(Sum(结帐金额), 0) Or Nvl(Sum(结帐金额), 0) = 0) And Not(Nvl(Sum(应收金额), 0) = 0 And Nvl(Sum(实收金额), 0) = 0)
                Order By 开单部门id, 开单人, 操作员姓名) Loop
    --          包括从未结的费用,费用明细部分结帐,以及结帐后作废,这些记录有可能已转入后备表 
    --          1.排开了已全部结帐的记录(Sum(应收金额)=Sum(应收金额)) 
    --          2.排开了无打折冲减的记帐后已销帐的记录(Sum(应收金额)=0,Sum(应收金额)=0) 
    --          3.不排开打折冲减后发生了单据销帐的记录，要将原冲减记录一并汇总重算(Sum(应收金额)=0,Sum(应收金额)<>0) 
    --          4.不排开打折冲减后产生的实收和结帐都为零的记录，因为改回原来的费别时，要重算回去 
    If r_Fee.应收金额 <> 0 Then
      Begin
        Select 实收金额
        Into n_实收金额
        From (Select Round(r_Fee.应收金额 * Nvl(实收比率, 0) / 100, n_小数位数) 实收金额
               From 费别明细
               Where 收费细目id = r_Fee.收费细目id And 费别 = v_费别 And Abs(r_Fee.应收金额) Between 应收段首值 And 应收段尾值 And Nvl(计算方法, 0) = 0
               Union All
               Select Round(r_Fee.应收金额 * Nvl(实收比率, 0) / 100, n_小数位数) 实收金额
               From 费别明细 A
               Where 收入项目id = r_Fee.收入项目id And 费别 = v_费别 And Abs(r_Fee.应收金额) Between 应收段首值 And 应收段尾值 And Nvl(计算方法, 0) = 0 And
                     Not Exists (Select 1 From 费别明细 B Where b.费别 = a.费别 And b.收费细目id = r_Fee.收费细目id));
      Exception
        When Others Then
          n_实收金额 := r_Fee.应收金额;
      End;
    Else
      n_实收金额 := 0;
    End If;
    --计算用来冲减原实收的差额 
    n_实收金额 := -1 * (r_Fee.实收金额 - n_实收金额);
  
    If n_实收金额 <> 0 Then
      --一张单据的开单部门id,开单人,操作员姓名,床号要求相同，如果其中之一变了则产生新单据，如果都没有变，一张单据最多100条明细 
      v_Thisinfo := r_Fee.开单部门id || r_Fee.开单人 || r_Fee.操作员姓名 || ' ';
      If v_Counter = 0 Or v_Counter = 100 Or v_Thisinfo <> v_Lastinfo Then
        v_No       := Nextno(14);
        v_Counter  := 1;
        v_Lastinfo := v_Thisinfo;
      Else
        v_Counter := v_Counter + 1;
      End If;
    
      Insert Into 门诊费用记录
        (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否,
         保险大类id, 付数, 数次, 发药窗口, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 划价人, 开单部门id, 开单人, 发生时间, 登记时间,
         执行部门id, 执行状态, 结帐id, 结帐金额, 操作员编号, 操作员姓名, 摘要, 是否急诊, 医嘱序号, 主页id, 挂号id, 病人病区id)
      Values
        (病人费用记录_Id.Nextval, 2, v_No, 1, v_Counter, Null, Null, Null, r_Fee.门诊标志, r_Fee.病人id, r_Fee.标识号, r_Fee.姓名,
         r_Fee.性别, r_Fee.年龄, r_Fee.病人科室id, v_费别, r_Fee.收费类别, r_Fee.收费细目id, r_Fee.计算单位, Null, Null, 0, 0, Null,
         r_Fee.加班标志, r_Fee.附加标志, r_Fee.婴儿费, r_Fee.收入项目id, r_Fee.收据费目, 0, 0, n_实收金额, Null, 1, Null, r_Fee.开单部门id,
         r_Fee.开单人, r_Fee.发生时间, d_Sysdate, r_Fee.执行部门id, 0, Null, Null, r_Fee.操作员编号, r_Fee.操作员姓名,
         Decode(v_Counter, 1, '实收重算冲减', ''), 0, Null, r_Fee.主页id, r_Fee.挂号id, r_Fee.病人病区id);
    End If;
  End Loop;

  If v_Counter = 0 Then
    v_Error := '由于以下原因之一,没有进行费用重算:' || Chr(13) || Chr(13) || 'a.没有发现病人本次住院的未结费用.' || Chr(13) || 'b.所有未结费用已进行了费用重算.' ||
               Chr(13) || 'c.按当前费别重算的实收冲减金额都为零.';
    Raise Err_Custom;
  Else
    --病人余额 
    n_实收金额 := 0;
    Select Sum(实收金额)
    Into n_实收金额
    From 门诊费用记录
    Where 病人id = 病人id_In And 记录性质 = 2 And Nvl(门诊标志, 0) <> 4 And 登记时间 = d_Sysdate;
    Update 病人余额 Set 费用余额 = Nvl(费用余额, 0) + n_实收金额 Where 病人id = 病人id_In And 性质 = 1 And 类型 = 1;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 费用余额, 预交余额, 类型) Values (病人id_In, 1, n_实收金额, 0, 1);
    End If;
  
    --病人未结费用 
    For r_Fee In (Select 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Sum(实收金额) 实收金额
                  From 门诊费用记录
                  Where 病人id = 病人id_In And 记录性质 = 2 And 登记时间 = d_Sysdate
                  Group By 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id) Loop
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + r_Fee.实收金额
      Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(r_Fee.主页id, 0) And Nvl(病人病区id, 0) = Nvl(r_Fee.病人病区id, 0) And
            Nvl(病人科室id, 0) = Nvl(r_Fee.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Fee.开单部门id, 0) And
            Nvl(执行部门id, 0) = Nvl(r_Fee.执行部门id, 0) And 收入项目id + 0 = Nvl(r_Fee.收入项目id, 0) And 来源途径 + 0 = 2;
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, r_Fee.主页id, r_Fee.病人病区id, r_Fee.病人科室id, r_Fee.开单部门id, r_Fee.执行部门id, r_Fee.收入项目id, 2, r_Fee.实收金额);
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人未结门诊费用_Recalc;
/

--139063:冉俊明,2019-04-01,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_门诊划价记录_Insert
(
  No_In         门诊费用记录.No%Type,
  序号_In       门诊费用记录.序号%Type,
  病人id_In     门诊费用记录.病人id%Type,
  主页id_In     住院费用记录.主页id%Type,
  标识号_In     门诊费用记录.标识号%Type,
  付款方式_In   门诊费用记录.付款方式%Type,
  姓名_In       门诊费用记录.姓名%Type,
  性别_In       门诊费用记录.性别%Type,
  年龄_In       门诊费用记录.年龄%Type,
  费别_In       门诊费用记录.费别%Type,
  加班标志_In   门诊费用记录.加班标志%Type,
  病人科室id_In 门诊费用记录.病人科室id%Type,
  开单部门id_In 门诊费用记录.开单部门id%Type,
  开单人_In     门诊费用记录.开单人%Type,
  从属父号_In   门诊费用记录.从属父号%Type,
  收费细目id_In 门诊费用记录.收费细目id%Type,
  收费类别_In   门诊费用记录.收费类别%Type,
  计算单位_In   门诊费用记录.计算单位%Type,
  发药窗口_In   门诊费用记录.发药窗口%Type,
  付数_In       门诊费用记录.付数%Type,
  数次_In       门诊费用记录.数次%Type,
  附加标志_In   门诊费用记录.附加标志%Type,
  执行部门id_In 门诊费用记录.执行部门id%Type,
  价格父号_In   门诊费用记录.价格父号%Type,
  收入项目id_In 门诊费用记录.收入项目id%Type,
  收据费目_In   门诊费用记录.收据费目%Type,
  标准单价_In   门诊费用记录.标准单价%Type,
  应收金额_In   门诊费用记录.应收金额%Type,
  实收金额_In   门诊费用记录.实收金额%Type,
  发生时间_In   门诊费用记录.发生时间%Type,
  登记时间_In   门诊费用记录.登记时间%Type,
  药品摘要_In   药品收发记录.摘要%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type,
  费用摘要_In   门诊费用记录.摘要%Type := Null,
  医嘱序号_In   门诊费用记录.医嘱序号%Type := Null,
  频次_In       药品收发记录.频次%Type := Null,
  单量_In       药品收发记录.单量%Type := Null,
  用法_In       药品收发记录.用法%Type := Null, --用法[|煎法]
  期效_In       药品收发记录.扣率%Type := Null,
  计价特性_In   药品收发记录.扣率%Type := Null,
  病人来源_In   Number := 1,
  保险编码_In   门诊费用记录.保险编码%Type := Null,
  费用类型_In   门诊费用记录.费用类型%Type := Null,
  保险项目否_In 门诊费用记录.保险项目否%Type := Null,
  保险大类id_In 门诊费用记录.保险大类id%Type := Null,
  中药形态_In   门诊费用记录.结论%Type := Null,
  备货材料_In   Number := 0,
  批次_In       药品收发记录.批次%Type := Null,
  执行人_In     门诊费用记录.执行人%Type := Null,
  病人病区id_In 门诊费用记录.病人病区id%Type := Null
) As
  --功能：新收一张门诊划价单据
  --参数：
  --   病人来源_IN:1-门诊病人,2-住院病人
  --     主页ID_IN:住院病人划价时用。
  --   药品摘要_IN:修改保存新单据时用。目前仅存放在药品收发记录的摘要中。
  --         新单据(记录状态=1)记录所修改的原单据号。
  v_费用id 门诊费用记录.Id%Type;
  n_急诊   病人挂号记录.急诊%Type;
  n_挂号id 病人挂号记录.Id%Type;
  n_主页id 门诊费用记录.主页id%Type;

  --临时变量
  v_用法       药品收发记录.用法%Type;
  v_煎法       药品收发记录.外观%Type;
  n_Dec        Number;
  v_付款方式   医疗付款方式.名称%Type;
  v_费别性质   费别.属性%Type;
  n_新病人模式 Number;
  n_单价小数   Number;
  v_Err_Msg    Varchar2(255);
  Err_Item Exception;
  v_Strtmpbefor Varchar2(4000);
  v_Msg         Varchar2(4000);
Begin
  --金额及单价小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))
  Into n_Dec, n_单价小数
  From Dual;

  n_主页id := 主页id_In;
  If Nvl(n_主页id, 0) = 0 Then
    Select Max(主页id) Into n_主页id From 病案主页 Where 病人id = 病人id_In And 病人性质 = 1 And 出院日期 Is Null;
  End If;

  --门诊费用记录
  Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
  --是否是急诊挂号单
  If Nvl(医嘱序号_In, 0) <> 0 Then
    Begin
      Select Nvl(Max(急诊), 0), Max(ID)
      Into n_急诊, n_挂号id
      From 病人挂号记录
      Where NO In (Select 挂号单 From 病人医嘱记录 Where ID = Nvl(医嘱序号_In, 0)) And 病人id = 病人id_In;
    Exception
      When Others Then
        n_急诊   := Null;
        n_挂号id := Null;
    End;
  End If;

  Insert Into 门诊费用记录
    (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 数次, 发药窗口,
     加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 记帐费用, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 摘要, 医嘱序号, 保险项目否, 保险编码,
     保险大类id, 费用类型, 结论, 是否急诊, 挂号id, 主页id, 付款方式, 执行人, 执行时间, 执行状态, 病人病区id)
  Values
    (v_费用id, 1, No_In, 0, 序号_In, Decode(从属父号_In, 0, Null, 从属父号_In), Decode(价格父号_In, 0, Null, 价格父号_In), Nvl(病人来源_In, 1),
     Decode(病人id_In, 0, Null, 病人id_In), Decode(标识号_In, 0, Null, 标识号_In), 姓名_In, 性别_In, 年龄_In, 病人科室id_In, 费别_In, 收费类别_In,
     收费细目id_In, 计算单位_In, 付数_In, 数次_In, 发药窗口_In, 加班标志_In, 附加标志_In, 收入项目id_In, 收据费目_In, 标准单价_In, 应收金额_In, 实收金额_In, 0,
     操作员姓名_In, 开单部门id_In, 开单人_In, 发生时间_In, 登记时间_In, 执行部门id_In, 费用摘要_In, 医嘱序号_In, 保险项目否_In, 保险编码_In, 保险大类id_In, 费用类型_In,
     中药形态_In, Nvl(n_急诊, 0), n_挂号id, n_主页id, 付款方式_In, 执行人_In, Decode(执行人_In, Null, Null, 登记时间_In),
     Decode(执行人_In, Null, 0, 2), 病人病区id_In);

  --药品和卫生材料部分
  If 收费类别_In In ('4', '5', '6', '7') Then
    --药品用法煎法分解
    If 用法_In Is Not Null Then
      If Instr(用法_In, '|') > 0 Then
        v_用法 := Substr(用法_In, 1, Instr(用法_In, '|') - 1);
        v_煎法 := Substr(用法_In, Instr(用法_In, '|') + 1);
      Else
        v_用法 := 用法_In;
      End If;
    End If;
    Zl_药品收发记录_销售出库(v_费用id, 药品摘要_In, 频次_In, 单量_In, v_用法, v_煎法, 期效_In, 计价特性_In, n_主页id, 备货材料_In, 批次_In);
  End If;

  --更新部份病人信息
  If 序号_In = 1 And 病人id_In Is Not Null Then
  
    If 付款方式_In Is Not Null And 病人来源_In = 1 Then
      Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In;
    End If;
  
    If 费别_In Is Not Null Then
      Select Max(属性) Into v_费别性质 From 费别 Where 名称 = 费别_In; --2-动态费别不更新
    End If;

    Update 病人信息
    Set 性别 = Decode(姓名, '新病人', Nvl(性别_In, 性别), 性别), 年龄 = Decode(姓名, '新病人', Nvl(年龄_In, 年龄), 年龄),
        姓名 = Decode(姓名, '新病人', 姓名_In, 姓名), 医疗付款方式 = Nvl(v_付款方式, 医疗付款方式), 费别 = Decode(v_费别性质, 1, 费别_In, 费别)
    Where 病人id = 病人id_In;

    Select Zl_To_Number(Nvl(zl_GetSysParameter('自动产生姓名', '1111'), '0')) Into n_新病人模式 From Dual;
    If n_新病人模式 = 1 Then
      Update 病人挂号记录
      Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In
      Where 病人id = 病人id_In And 姓名 = '新病人';
      Update 门诊费用记录
      Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In
      Where 病人id = 病人id_In And 姓名 = '新病人';
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊划价记录_Insert;
/

--139063:冉俊明,2019-04-01,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_门诊记帐记录_Delete
(
  No_In           门诊费用记录.No%Type,
  序号_In         Varchar2,
  操作员编号_In   门诊费用记录.操作员编号%Type,
  操作员姓名_In   门诊费用记录.操作员姓名%Type,
  输液配药检查_In Number := 1,
  登记时间_In     住院费用记录.登记时间%Type := Sysdate
) As
  --功能：冲销一张门诊记帐单据中指定序号行 
  --序号：格式如"1,3,5,7,8",或"1:2:33456,3:2,5:2,7:2,8:2",冒号前面的数字表示行号,中间的数字表示退的数量,后面的数字表示配药记录的ID,目前仅在销帐审核时才传入 
  --      为空表示冲销所有可冲销行 

  --该游标为要退费单据的所有原始记录
  Cursor c_Bill(n_标志 Number) Is
    Select a.Id, a.价格父号, a.序号, a.执行状态, a.收费类别, a.医嘱序号, a.病人id, a.主页id, a.收入项目id, a.开单部门id, a.执行部门id, a.病人病区id, a.病人科室id,
           a.实收金额, Decode(a.记录状态, 0, 1, 0) As 划价, j.诊疗类别, m.跟踪在用
    From 门诊费用记录 A, 病人医嘱记录 J, 材料特性 M
    Where a.医嘱序号 = j.Id(+) And a.收费细目id + 0 = m.材料id(+) And a.No = No_In And a.记录性质 = 2 And a.记录状态 In (0, 1, 3) And
          a.门诊标志 = n_标志
    Order By a.收费细目id, a.序号;

  --该游标用于处理费用记录序号
  Cursor c_Serial Is
    Select 序号, 价格父号 From 门诊费用记录 Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) Order By 序号;
  l_划价 t_Numlist := t_Numlist();

  v_医嘱ids  Varchar2(4000);
  n_父号     门诊费用记录.价格父号%Type;
  n_门诊标志 门诊费用记录.门诊标志%Type;

  --部分退费计算变量
  n_剩余数量 Number;
  n_剩余应收 Number;
  n_剩余实收 Number;
  n_剩余统筹 Number;

  n_准退数量 Number;
  n_退费次数 Number;
  n_退费数量 Number;
  n_部分销帐 Number;

  n_应收金额 Number;
  n_实收金额 Number;
  n_统筹金额 Number;

  v_序号   Varchar2(4000);
  v_配药id Varchar2(4000);
  v_Tmp    Varchar2(4000);

  n_Dec Number;

  n_Count   Number;
  d_Curdate Date;
  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  --销帐审核时,非药品会传入行号的销帐数量 
  If Not 序号_In Is Null Then
    If Instr(序号_In, ':') > 0 Then
      --格式：1:2:33456,3:2,5:2,7:2,8:2
      For c_序号 In (Select C1, C2 From Table(f_Str2list2(序号_In, ',', ':'))) Loop
        v_序号 := v_序号 || ',' || c_序号.C1;
        If Instr(c_序号.C2, ':') > 0 Then
          v_配药id := v_配药id || ',' || Substr(c_序号.C2, Instr(c_序号.C2, ':') + 1);
        End If;
      End Loop;
      v_序号   := Substr(v_序号, 2);
      v_配药id := Substr(v_配药id, 2);
    Else
      v_序号 := 序号_In;
    End If;
  End If;

  --是否已经全部完全执行(只是整张单据的检查)
  Select Nvl(Count(1), 0), Max(Nvl(门诊标志, 1))
  Into n_Count, n_门诊标志
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) And Nvl(执行状态, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '该单据中的项目已经全部完全执行！';
    Raise Err_Item;
  End If;

  If Nvl(n_门诊标志, 0) = 0 Then
    n_门诊标志 := 1;
  End If;

  --未完全执行的项目是否有剩余数量(只是整张单据的检查)
  Select Nvl(Count(1), 0)
  Into n_Count
  From (Select 序号, Sum(数量) As 剩余数量
         From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                From 门诊费用记录
                Where NO = No_In And 记录性质 = 2 And 门诊标志 = n_门诊标志 And
                      Nvl(价格父号, 序号) In
                      (Select Nvl(价格父号, 序号)
                       From 门诊费用记录
                       Where NO = No_In And 记录性质 = 2 And 门诊标志 = n_门诊标志 And 记录状态 In (0, 1, 3) And Nvl(执行状态, 0) <> 1)
                Group By 记录状态, Nvl(价格父号, 序号))
         Group By 序号
         Having Sum(数量) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '该单据中未完全执行部分项目剩余数量为零,没有可以销帐的费用！';
    Raise Err_Item;
  End If;

  ---------------------------------------------------------------------------------
  --公用变量
  Select Nvl(登记时间_In, Sysdate) Into d_Curdate From Dual;

  --金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --循环处理每行费用(收入项目行)
  For r_Bill In c_Bill(n_门诊标志) Loop
    If Instr(',' || v_序号 || ',', ',' || Nvl(r_Bill.价格父号, r_Bill.序号) || ',') > 0 Or v_序号 Is Null Then
      If Nvl(r_Bill.执行状态, 0) <> 1 Then
        --求剩余数量,剩余应收,剩余实收
        Select Sum(Nvl(付数, 1) * 数次), Sum(应收金额), Sum(实收金额), Sum(统筹金额)
        Into n_剩余数量, n_剩余应收, n_剩余实收, n_剩余统筹
        From 门诊费用记录
        Where NO = No_In And 记录性质 = 2 And 序号 = r_Bill.序号;
      
        n_部分销帐 := 0;
        n_退费数量 := 0;
        If n_剩余数量 = 0 Then
          If v_序号 Is Not Null Then
            v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经全部销帐！';
            Raise Err_Item;
          End If;
          --情况：未限定行号,原始单据中的该笔已经全部销帐(执行状态=0的一种可能)
        Else
          If Instr(序号_In, ':') > 0 Then
            Select Max(a.C2) Into v_Tmp From Table(f_Str2list2(序号_In, ',', ':')) A Where a.C1 = r_Bill.序号;
            If Instr(v_Tmp, ':') > 0 Then
              n_退费数量 := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
            Else
              n_退费数量 := To_Number(v_Tmp);
            End If;
            n_部分销帐 := 1;
          End If;
        
          --准销数量(非药品项目为剩余数量,原始数量)
          If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Or (r_Bill.收费类别 = '4' And Nvl(r_Bill.跟踪在用, 0) = 0) Then
            --@@@
            --非药品部分(以具体医嘱执行为准进行检查)
            --: 1.存在医嘱发送的,则以医嘱执行为准(但不能包含:检查;检验;手术;麻醉及输血)
            --: 2.对于病人医吃计价中的收费方式为:0-正常收取 的,才支持部分退;如果是其他的,则只能全退
            --: 3.不存在医嘱的,则以剩余数量为准
            n_Count := 0;
            If Instr(',C,D,F,G,K,', ',' || r_Bill.诊疗类别 || ',') = 0 And r_Bill.诊疗类别 Is Not Null Then
              Select Nvl(Sum(数量), 0), Count(*)
              Into n_准退数量, n_Count
              From (Select j.医嘱序号 As 医嘱id, j.收费细目id, Nvl(j.付数, 1) * Nvl(j.数次, 1) As 数量
                     From 门诊费用记录 J, 病人医嘱记录 M
                     Where j.医嘱序号 = m.Id And j.No = No_In And j.记录性质 = 2 And j.序号 = r_Bill.序号 And j.记录状态 In (1, 3) And
                           Exists
                      (Select 1
                            From 病人医嘱发送 A
                            Where a.医嘱id = j.医嘱序号 And Nvl(a.执行状态, 0) <> 1 And a.No || '' = No_In) And Exists
                      (Select 1
                            From 病人医嘱计价 A
                            Where a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And Nvl(a.收费方式, 0) = 0) And j.价格父号 Is Null And
                           Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And
                           (j.记录状态 In (1, 3) And Not Exists
                            (Select 1
                             From 药品收发记录
                             Where 费用id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0) Or
                            j.记录状态 = 2 And Not Exists
                            (Select 1 From 药品收发记录 Where NO = No_In And 单据 In (8, 24) And 药品id = j.收费细目id))
                     Union All
                     Select a.医嘱id, a.收费细目id, -1 * Nvl(a.数量, 1) * Nvl(c.本次数次, 1) As 数量
                     From 病人医嘱计价 A, 病人医嘱发送 B, 病人医嘱执行 C, 门诊费用记录 J, 病人医嘱记录 M
                     Where a.医嘱id = b.医嘱id And b.医嘱id = c.医嘱id And Nvl(a.收费方式, 0) = 0 And b.发送号 = c.发送号 And a.医嘱id = m.Id And
                           Nvl(c.执行结果, 1) = 1 And Nvl(b.执行状态, 0) <> 1 And a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And
                           j.No = No_In And j.记录性质 = 2 And j.序号 = r_Bill.序号 And j.记录状态 In (1, 3) And j.价格父号 Is Null And
                           Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Not Exists
                      (Select 1
                            From 药品收发记录
                            Where 费用id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0) And Not Exists
                      (Select 1 From 材料特性 Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)
                     Union All
                     Select a.医嘱id, a.收费细目id, 0 As 数量
                     From 病人医嘱计价 A, 门诊费用记录 J, 病人医嘱记录 M
                     Where a.医嘱id = m.Id And a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And Nvl(a.收费方式, 0) <> 0 And
                           j.No = No_In And j.记录性质 = 2 And Nvl(j.执行状态, 0) = 2 And Not Exists
                      (Select 1 From 材料特性 Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1) And
                           Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0);
            End If;
          
            If Nvl(n_Count, 0) = 0 Then
              n_准退数量 := n_剩余数量;
            End If;
          Else
            Select Sum(Nvl(付数, 1) * 实际数量)
            Into n_准退数量
            From 药品收发记录
            Where NO = No_In And 单据 In (9, 25) And Mod(记录状态, 3) = 1 And 审核人 Is Null And 费用id = r_Bill.Id;
          
            --不跟踪在用的卫生材料
            If r_Bill.收费类别 = '4' And Nvl(n_准退数量, 0) = 0 Then
              n_准退数量 := n_剩余数量;
            End If;
          End If;
        
          If Nvl(n_退费数量, 0) = 0 Then
            n_退费数量 := n_准退数量;
          Else
            If n_准退数量 < n_退费数量 Then
              v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用准退数量不足本次销帐数量！';
              Raise Err_Item;
            End If;
          End If;
        
          --金额=剩余金额*(准退数/剩余数)
          n_应收金额 := Round(n_剩余应收 * (n_退费数量 / n_剩余数量), n_Dec);
          n_实收金额 := Round(n_剩余实收 * (n_退费数量 / n_剩余数量), n_Dec);
          n_统筹金额 := Round(n_剩余统筹 * (n_退费数量 / n_剩余数量), n_Dec);
        
          If Nvl(r_Bill.划价, 0) = 0 Then
            --该笔项目第几次销帐
            Select Nvl(Max(Abs(执行状态)), 0) + 1
            Into n_退费次数
            From 门诊费用记录
            Where NO = No_In And 记录性质 = 2 And 记录状态 = 2 And 序号 = r_Bill.序号;
          
            --插入退费记录
            Insert Into 门诊费用记录
              (ID, NO, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 婴儿费, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别,
               收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人,
               执行人, 执行状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 记帐单id, 摘要, 保险编码, 是否急诊, 结论, 挂号id, 主页id,
               病人病区id)
              Select 病人费用记录_Id.Nextval, NO, 记录性质, 2, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 婴儿费, 姓名, 性别, 年龄, 标识号, 付款方式, 费别,
                     病人科室id, 收费类别, 收费细目id, 计算单位, Decode(Sign(n_退费数量 - Nvl(付数, 1) * 数次), 0, 付数, 1), 发药窗口,
                     Decode(Sign(n_退费数量 - Nvl(付数, 1) * 数次), 0, -1 * 数次, -1 * n_退费数量), 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用,
                     标准单价, -1 * n_应收金额, -1 * n_实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, -1 * n_退费次数, 执行时间, 操作员编号_In,
                     操作员姓名_In, 发生时间, d_Curdate, 保险项目否, 保险大类id, -1 * n_统筹金额, 记帐单id, 摘要, 保险编码, 是否急诊, 结论, 挂号id, 主页id,
                     病人病区id
              From 门诊费用记录
              Where ID = r_Bill.Id;
          
            --病人余额
            If n_门诊标志 <> 4 Then
              Update 病人余额
              Set 费用余额 = Nvl(费用余额, 0) - n_实收金额
              Where 病人id = r_Bill.病人id And 性质 = 1 And 类型 = 1;
              If Sql%RowCount = 0 Then
                Insert Into 病人余额
                  (病人id, 性质, 类型, 费用余额, 预交余额)
                Values
                  (r_Bill.病人id, 1, 1, -1 * n_实收金额, 0);
              End If;
            End If;
          
            --病人未结费用
            Update 病人未结费用
            Set 金额 = Nvl(金额, 0) - n_实收金额
            Where 病人id = r_Bill.病人id And Nvl(主页id, 0) = Nvl(r_Bill.主页id, 0) And Nvl(病人病区id, 0) = Nvl(r_Bill.病人病区id, 0) And
                  Nvl(病人科室id, 0) = Nvl(r_Bill.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Bill.开单部门id, 0) And
                  Nvl(执行部门id, 0) = Nvl(r_Bill.执行部门id, 0) And 收入项目id + 0 = r_Bill.收入项目id And 来源途径 + 0 = n_门诊标志;
            If Sql%RowCount = 0 Then
              Insert Into 病人未结费用
                (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
              Values
                (r_Bill.病人id, r_Bill.主页id, r_Bill.病人病区id, r_Bill.病人科室id, r_Bill.开单部门id, r_Bill.执行部门id, r_Bill.收入项目id,
                 n_门诊标志, -1 * n_实收金额);
            End If;
          
            --标记原费用记录
            --执行状态:全部退完(准退数=剩余数)标记为0,否则标记为1
            Update 门诊费用记录
            Set 记录状态 = 3, 执行状态 = Decode(Sign(n_退费数量 - n_剩余数量), 0, 0, 1)
            Where ID = r_Bill.Id;
          Else
            --划价记账单
            If Nvl(n_部分销帐, 0) = 0 Then
              l_划价.Extend;
              l_划价(l_划价.Count) := r_Bill.Id;
            Else
              --更新数量 
              --划价的,先将相关的数据处理在内部表集中
              Update 住院费用记录
              Set 付数 = 1, 数次 = Nvl(付数, 1) * 数次 - n_退费数量, 应收金额 = Nvl(应收金额, 0) - n_应收金额, 实收金额 = Nvl(实收金额, 0) - n_实收金额,
                  登记时间 = d_Curdate, 统筹金额 = Nvl(统筹金额, 0) - n_统筹金额
              Where ID = r_Bill.Id
              Returning 数次 Into n_剩余数量;
              If Nvl(n_剩余数量, 0) <= 0 Then
                l_划价.Extend;
                l_划价(l_划价.Count) := r_Bill.Id;
              End If;
            End If;
          
            If r_Bill.医嘱序号 Is Not Null Then
              If Instr(',' || Nvl(v_医嘱ids, '') || ',', ',' || r_Bill.医嘱序号 || ',') = 0 Then
                v_医嘱ids := Nvl(v_医嘱ids, '') || ',' || r_Bill.医嘱序号;
              End If;
            End If;
          End If;
        End If;
      Else
        If v_序号 Is Not Null Then
          v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经完全执行,不能销帐！';
          Raise Err_Item;
        End If;
        --情况:没限定行号,原始单据中包括已经完全执行的
      End If;
    End If;
  End Loop;

  --不存在配药ID,检查该药品是否在输液配药中心 
  If v_配药id Is Null And 输液配药检查_In = 1 Then
    For v_费用 In (Select ID
                 From 门诊费用记录
                 Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) And 收费类别 In ('4', '5', '6', '7') And 门诊标志 = n_门诊标志 And
                       (Instr(',' || v_序号 || ',', ',' || 序号 || ',') > 0 Or v_序号 Is Null)) Loop
      Begin
        Select Count(1)
        Into n_Count
        From 输液配药内容 A, 药品收发记录 B
        Where a.收发id = b.Id And b.费用id = v_费用.Id And Instr(',8,9,10,21,24,25,26,', ',' || b.单据 || ',') > 0;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count <> 0 Then
        v_Err_Msg := '存在已经进入输液配药中心的待销帐药品，无法完成销帐！';
        Raise Err_Item;
      End If;
    End Loop;
  End If;

  --------------------------------------------------------------------------------- 
  --药品相关处理:主要是对销帐审核有效.(可以是部分) 
  --必须按照“收费细目id”升序排序，防止并发锁“药品库存”表
  For v_费用 In (Select ID, 序号
               From 门诊费用记录
               Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) And 收费类别 In ('4', '5', '6', '7') And 门诊标志 = n_门诊标志 And
                     (Instr(',' || v_序号 || ',', ',' || 序号 || ',') > 0 Or v_序号 Is Null)
               Order By 收费细目id) Loop
    --根据费用ID来进行相关的处理 
    n_退费数量 := 0;
    If Instr(序号_In, ':') > 0 Then
      Select Max(a.C2) Into v_Tmp From Table(f_Str2list2(序号_In, ',', ':')) A Where a.C1 = v_费用.序号;
      If Instr(v_Tmp, ':') > 0 Then
        n_退费数量 := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
      Else
        n_退费数量 := To_Number(v_Tmp);
      End If;
    End If;
    Zl_药品收发记录_销售退费(v_费用.Id, n_退费数量, v_配药id);
  End Loop;

  --删除划价记录
  n_Count := l_划价.Count;
  Forall I In 1 .. l_划价.Count
    Delete From 门诊费用记录 Where ID = l_划价(I);

  --删除之后再统一调整序号
  If n_Count > 0 Then
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.价格父号 Is Null Then
        n_父号 := n_Count;
      End If;
    
      Update 门诊费用记录
      Set 序号 = n_Count, 价格父号 = Decode(价格父号, Null, Null, n_父号)
      Where NO = No_In And 记录性质 = 2 And 序号 = r_Serial.序号;
    
      Update 门诊费用记录 Set 从属父号 = n_Count Where NO = No_In And 记录性质 = 2 And 从属父号 = r_Serial.序号;
    
      n_Count := n_Count + 1;
    End Loop;
  End If;

  --整张单据全部冲完时，删除病人医嘱附费
  For c_医嘱 In (Select Distinct 医嘱序号
               From 门诊费用记录
               Where NO = No_In And 记录性质 = 2 And 记录状态 = 3 And 医嘱序号 Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select 序号, Sum(数量) As 剩余数量
           From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                  From 门诊费用记录
                  Where 记录性质 = 2 And 医嘱序号 + 0 = c_医嘱.医嘱序号 And NO = No_In
                  Group By 记录状态, Nvl(价格父号, 序号))
           Group By 序号
           Having Sum(数量) <> 0);
  
    If n_Count = 0 Then
      Delete From 病人医嘱附费 Where 医嘱id = c_医嘱.医嘱序号 And 记录性质 = 2 And NO = No_In;
    End If;
  End Loop;

  If v_医嘱ids Is Not Null Then
    --医嘱处理
    --场合_In    Integer:=0, --0:门诊;1-住院
    --性质_In    Integer:=1, --1-收费单;2-记帐单
    --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
    --No_In      门诊费用记录.No%Type,
    --医嘱ids_In Varchar2 := Null
    v_医嘱ids := Substr(v_医嘱ids, 2);
    Zl_医嘱发送_计费状态_Update(0, 2, 0, No_In, v_医嘱ids);
  Else
    Zl_医嘱发送_计费状态_Update(0, 2, 2, No_In, v_医嘱ids);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊记帐记录_Delete;
/

--139063:冉俊明,2019-04-01,门诊留观病人按门诊流程就诊
Create Or Replace Procedure Zl_门诊记帐记录_Insert
(
  No_In         门诊费用记录.No%Type,
  序号_In       门诊费用记录.序号%Type,
  病人id_In     门诊费用记录.病人id%Type,
  标识号_In     门诊费用记录.标识号%Type,
  姓名_In       门诊费用记录.姓名%Type,
  性别_In       门诊费用记录.性别%Type,
  年龄_In       门诊费用记录.年龄%Type,
  费别_In       门诊费用记录.费别%Type,
  加班标志_In   门诊费用记录.加班标志%Type,
  婴儿费_In     门诊费用记录.婴儿费%Type,
  病人科室id_In 门诊费用记录.病人科室id%Type,
  开单部门id_In 门诊费用记录.开单部门id%Type,
  开单人_In     门诊费用记录.开单人%Type,
  从属父号_In   门诊费用记录.从属父号%Type,
  收费细目id_In 门诊费用记录.收费细目id%Type,
  收费类别_In   门诊费用记录.收费类别%Type,
  计算单位_In   门诊费用记录.计算单位%Type,
  付数_In       门诊费用记录.付数%Type,
  数次_In       门诊费用记录.数次%Type,
  附加标志_In   门诊费用记录.附加标志%Type,
  执行部门id_In 门诊费用记录.执行部门id%Type,
  价格父号_In   门诊费用记录.价格父号%Type,
  收入项目id_In 门诊费用记录.收入项目id%Type,
  收据费目_In   门诊费用记录.收据费目%Type,
  标准单价_In   门诊费用记录.标准单价%Type,
  应收金额_In   门诊费用记录.应收金额%Type,
  实收金额_In   门诊费用记录.实收金额%Type,
  发生时间_In   门诊费用记录.发生时间%Type,
  登记时间_In   门诊费用记录.登记时间%Type,
  药品摘要_In   药品收发记录.摘要%Type,
  划价_In       Number,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type,
  记帐单id_In   门诊费用记录.记帐单id%Type := Null,
  费用摘要_In   门诊费用记录.摘要%Type := Null,
  医嘱序号_In   门诊费用记录.医嘱序号%Type := Null,
  频次_In       药品收发记录.频次%Type := Null,
  单量_In       药品收发记录.单量%Type := Null,
  用法_In       药品收发记录.用法%Type := Null, --用法[|煎法]
  期效_In       药品收发记录.扣率%Type := Null,
  计价特性_In   药品收发记录.扣率%Type := Null,
  门诊标志_In   门诊费用记录.门诊标志%Type := 1,
  中药形态_In   门诊费用记录.结论%Type := Null,
  备货材料_In   Number := 0,
  批次_In       药品收发记录.批次%Type := Null,
  主页id_In     门诊费用记录.主页id%Type := Null,
  病人病区id_In 门诊费用记录.病人病区id%Type := Null
) As
  --功能：新收一张门诊记帐单据
  --参数：
  --   药品摘要_IN:修改保存新单据时用。目前仅用于存放于药品收发记录的摘要中。
  --         原单据(记录状态=2)记录修改产生的新单据号。
  --         新单据(记录状态=1)记录所修改的原单据号。
  v_费用id 门诊费用记录.Id%Type;
  n_急诊   病人挂号记录.急诊%Type;
  n_主页id 门诊费用记录.主页id%Type;

  --临时变量
  v_用法     药品收发记录.用法%Type;
  v_煎法     药品收发记录.外观%Type;
  n_单价小数 Number;
  n_挂号id   病人挂号记录.Id%Type;

  n_Dec     Number;
  n_Count   Number;
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_发药窗口 药品收发记录.发药窗口%Type;
  n_跟踪在用 材料特性.跟踪在用%Type;

Begin
  n_跟踪在用 := 0;
  If 收费类别_In = '4' Then
    --跟踪在用的卫材才处理
    Select Nvl(跟踪在用, 0) Into n_跟踪在用 From 材料特性 Where 材料id = 收费细目id_In;
  End If;

  --金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))
  Into n_Dec, n_单价小数
  From Dual;

  n_主页id := 主页id_In;
  If Nvl(n_主页id, 0) = 0 Then
    Select Max(主页id) Into n_主页id From 病案主页 Where 病人id = 病人id_In And 病人性质 = 1 And 出院日期 Is Null;
  End If;

  If (收费类别_In In ('5', '6', '7') Or 收费类别_In = '4' And n_跟踪在用 = 1) And Nvl(划价_In, 0) = 0 Then
    --同一张单据,满足同一药房同一窗口
    Begin
      Select 发药窗口
      Into v_发药窗口
      From 门诊费用记录
      Where 收费类别 In ('5', '6', '7', '4') And NO = No_In And 记录性质 = 2 And 执行部门id = 执行部门id_In And 发药窗口 Is Not Null And
            Rownum <= 1;
    Exception
      When Others Then
        v_发药窗口 := Null;
    End;
    If v_发药窗口 Is Null Then
      --同一病人在普通号挂号有效挂号天数内且未发药的且上班的,以最近一次记账窗口为准
      n_Count := To_Number(Substr(Nvl(zl_GetSysParameter(21), '11') || '11', 1, 1));
      If n_Count = 0 Then
        n_Count := 1;
      End If;
    
      Begin
        Select 发药窗口
        Into v_发药窗口
        From (Select 登记时间, 发药窗口
               From 门诊费用记录 A
               Where 收费类别 In ('5', '6', '7', '4') And 病人id = 病人id_In And 登记时间 Between Sysdate - n_Count And Sysdate And
                     记录性质 = 2 And 执行部门id = 执行部门id_In And 发药窗口 Is Not Null And Exists
                (Select 1
                      From 未发药品记录
                      Where a.No = NO And 单据 In (9, 26) And 库房id + 0 = 执行部门id_In And 病人id + 0 = 病人id_In) And Exists
                (Select 1
                      From 发药窗口
                      Where Nvl(上班否, 0) = 1 And 名称 = a.发药窗口 And Nvl(专家, 0) = 0 And 药房id = 执行部门id_In)
               Order By 登记时间 Desc)
        Where Rownum <= 1;
      
      Exception
        When Others Then
          v_发药窗口 := Null;
      End;
      If v_发药窗口 Is Null Then
        v_发药窗口 := Zl_Get发药窗口(执行部门id_In);
      End If;
    End If;
  End If;
  --门诊费用记录
  Select 病人费用记录_Id.Nextval Into v_费用id From Dual;

  --是否是急诊挂号单
  If Nvl(医嘱序号_In, 0) <> 0 Then
    Begin
      Select Nvl(Max(急诊), 0), Max(ID)
      Into n_急诊, n_挂号id
      From 病人挂号记录
      Where NO In (Select 挂号单 From 病人医嘱记录 Where ID = Nvl(医嘱序号_In, 0)) And 病人id = 病人id_In;
    Exception
      When Others Then
        n_急诊   := Null;
        n_挂号id := Null;
    End;
  End If;

  Insert Into 门诊费用记录
    (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 数次, 加班标志,
     附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 记帐费用, 划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行状态, 操作员编号, 操作员姓名, 婴儿费, 记帐单id,
     摘要, 医嘱序号, 结论, 发药窗口, 是否急诊, 主页id, 挂号id, 病人病区id)
  Values
    (v_费用id, 2, No_In, Decode(划价_In, 1, 0, 1), 序号_In, Decode(从属父号_In, 0, Null, 从属父号_In),
     Decode(价格父号_In, 0, Null, 价格父号_In), 门诊标志_In, 病人id_In, Decode(标识号_In, 0, Null, 标识号_In), 姓名_In, 性别_In, 年龄_In,
     病人科室id_In, 费别_In, 收费类别_In, 收费细目id_In, 计算单位_In, 付数_In, 数次_In, 加班标志_In, 附加标志_In, 收入项目id_In, 收据费目_In, 标准单价_In, 应收金额_In,
     实收金额_In, 1, 操作员姓名_In, 开单部门id_In, 开单人_In, 发生时间_In, 登记时间_In, 执行部门id_In, 0, Decode(划价_In, 1, Null, 操作员编号_In),
     Decode(划价_In, 1, Null, 操作员姓名_In), 婴儿费_In, 记帐单id_In, 费用摘要_In, 医嘱序号_In, 中药形态_In, v_发药窗口, Nvl(n_急诊, 0), n_主页id, n_挂号id,
     Decode(病人病区id_In, 0, Null, 病人病区id_In));

  --相关汇总表的处理
  If Nvl(划价_In, 0) = 0 Then
    --病人余额
    If Nvl(门诊标志_In, 0) <> 4 Then
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + 实收金额_In
      Where 病人id = 病人id_In And 性质 = 1 And 类型 = Decode(门诊标志_In, 2, 2, 1);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 费用余额, 预交余额)
        Values
          (病人id_In, 1, Decode(门诊标志_In, 2, 2, 1), 实收金额_In, 0);
      End If;
    End If;
  
    --病人未结费用
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + 实收金额_In
    Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(n_主页id, 0) And Nvl(病人病区id, 0) = Nvl(病人病区id_In, 0) And
          Nvl(病人科室id, 0) = Nvl(病人科室id_In, 0) And Nvl(开单部门id, 0) = Nvl(开单部门id_In, 0) And
          Nvl(执行部门id, 0) = Nvl(执行部门id_In, 0) And 收入项目id + 0 = 收入项目id_In And 来源途径 + 0 = 门诊标志_In;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (病人id_In, n_主页id, 病人病区id_In, 病人科室id_In, 开单部门id_In, 执行部门id_In, 收入项目id_In, 门诊标志_In, 实收金额_In);
    End If;
  
  End If;

  --药品和卫生材料部分
  If 收费类别_In In ('4', '5', '6', '7') Then
    --药品用法煎法分解
    If 用法_In Is Not Null Then
      If Instr(用法_In, '|') > 0 Then
        v_用法 := Substr(用法_In, 1, Instr(用法_In, '|') - 1);
        v_煎法 := Substr(用法_In, Instr(用法_In, '|') + 1);
      Else
        v_用法 := 用法_In;
      End If;
    End If;
    Zl_药品收发记录_销售出库(v_费用id, 药品摘要_In, 频次_In, 单量_In, v_用法, v_煎法, 期效_In, 计价特性_In, n_主页id, 备货材料_In, 批次_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊记帐记录_Insert;
/

--139063:冉俊明,2019-04-01,门诊留观病人按门诊流程就诊
Create Or Replace Function Zl1_Getdef_Prepaymoney
(
  病人id_In 病人信息.病人id%Type,
  主页id_In 病案主页.主页id%Type,
  来源_In   Number := 2
  --功能：获取默认的预交款缴款额 
  --     本函数主要供病人费用查询时调用缴预交时调用,主要是读取缺省的预交金额 
  --     用户可以根据实际产生的规则,来生成这个缺省值 
  --参数： 
  --    病人ID_In：病人ID 
  --    主页ID_In:主页ID 
  --    来源_In:1-门诊，2-住院
) Return Number Is
  Err_Custom Exception;
  n_预交余额 病人余额.预交余额%Type;
  n_费用余额 病人余额.费用余额%Type;
  n_预结费用 病人余额.预交余额%Type;
  n_本次预交 病人余额.预交余额%Type;
Begin
  --目前按规则:（总费用-预交款总额-报销金额 >0） 
  Select Nvl(Sum(预交余额), 0), Nvl(Sum(费用余额), 0), Nvl(Sum(预结费用), 0)
  Into n_预交余额, n_费用余额, n_预结费用
  From (Select Nvl(预交余额, 0) 预交余额, Nvl(费用余额, 0) 费用余额, 0 As 预结费用
         From 病人余额
         Where 性质 = 1 And 病人id = 病人id_In And Nvl(类型, 2) = Nvl(来源_In, 2)
         Union All
         Select 0 As 预交余额, 0 As 费用余额, Sum(b.金额) As 预结费用
         From 病人信息 A, 保险模拟结算 B
         Where a.病人id = b.病人id And a.主页id = b.主页id And a.病人id = 病人id_In);
  n_本次预交 := n_预交余额 - n_费用余额 + n_预结费用;
  If n_本次预交 > 0 Then
    n_本次预交 := 0;
  End If;
  n_本次预交 := Abs(n_本次预交);
  Return n_本次预交;

End Zl1_Getdef_Prepaymoney;
/





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0056' Where 编号=&n_System;
Commit;
