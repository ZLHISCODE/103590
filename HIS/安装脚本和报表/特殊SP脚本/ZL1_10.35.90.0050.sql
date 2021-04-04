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


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--134869:李小东,2019-02-19,采集站病区条码打印添加权限“完成采集”
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1211,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0
Union All Select '完成采集',7,'有该权限时允许在病区条码打印中完成采集',0 From Dual) A;




-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--137272:李南春,2019-02-22,预约提前接收锁定序号
CREATE OR REPLACE Procedure Zl_挂号序号状态_Lock
(
  操作_In       Number, --1-锁定,2-清除锁定
  操作员姓名_In 挂号序号状态.操作员姓名%Type,
  号码_In       挂号序号状态.号码%Type := Null,
  日期_In       挂号序号状态.日期%Type := Null,
  序号_In       挂号序号状态.序号%Type := Null,
  出诊记录ID_In 临床出诊记录.ID%type := Null,
  备注_In       挂号序号状态.备注%Type := Null
) As

  v_姓名       挂号序号状态.操作员姓名%Type;
  v_状态       挂号序号状态.状态%Type;
  v_机器名     挂号序号状态.机器名%Type;
  v_验证机器名 挂号序号状态.机器名%Type;
  v_备注       挂号序号状态.备注%Type;
  v_工作站IP   临床出诊序号控制.工作站IP%Type;
  v_Error      Varchar2(255);
  Err_Custom Exception;
Begin
  Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  Select SYS_CONTEXT('USERENV','IP_ADDRESS') Into v_工作站IP from dual;
  If 操作_In = 1 Then
    If 备注_In Is Null Then
      v_备注 := '自助机锁号';
    Else
      v_备注 := 备注_In;
    End If;
    --锁定挂号序号状态
    If 出诊记录ID_In is Null then
      Begin
        Select 操作员姓名, 状态, 机器名
        Into v_姓名, v_状态, v_验证机器名
        From 挂号序号状态
        Where 号码 = 号码_In And 日期 = 日期_In And 序号 = 序号_In;
      Exception
        When Others Then
          Null;
      End;
      If v_姓名 Is Null Then
        Insert Into 挂号序号状态
          (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
        Values
          (号码_In, 日期_In, 序号_In, 5, 操作员姓名_In, v_备注, Sysdate, v_机器名);
      Else
        v_Error := '序号' || 序号_In || '已被操作员' || v_姓名;
        If v_状态 = 1 Then
          v_Error := v_Error || '使用';
        Elsif v_状态 = 2 Then
          v_Error := v_Error || '预约';
        Elsif v_状态 = 3 Then
          v_Error := v_Error || '预留';
        Elsif v_状态 = 4 Then
          v_Error := v_Error || '退号';
        Elsif v_状态 = 5 Then
          v_Error := v_Error || '(' || v_验证机器名 || ')锁定';
        End If;
        Raise Err_Custom;
      End If;
    Else
      Begin
        Select 操作员姓名, 挂号状态, 工作站名称
        Into v_姓名, v_状态, v_验证机器名
        From 临床出诊序号控制
        Where 记录ID = 出诊记录ID_In And 序号 = 序号_In;
      Exception
        When Others Then
          Null;
      End;
      
      If Nvl(v_状态,0) = 0 Then
        Update 临床出诊序号控制 set 挂号状态=5,锁号时间=Sysdate,操作员姓名=操作员姓名_In,工作站IP=v_工作站IP,工作站名称=v_机器名,备注=v_备注
        Where 记录ID=出诊记录ID_In  And 序号=序号_In;
      Else
        v_Error := '序号' || 序号_In || '已被操作员' || v_姓名;
        If v_状态 = 1 Then
          v_Error := v_Error || '使用';
        Elsif v_状态 = 2 Then
          v_Error := v_Error || '预约';
        Elsif v_状态 = 3 Then
          v_Error := v_Error || '预留';
        Elsif v_状态 = 4 Then
          v_Error := v_Error || '退号';
        Elsif v_状态 = 5 Then
          v_Error := v_Error || '(' || v_验证机器名 || ')锁定';
        End If;
        Raise Err_Custom;
      End If;
    End If;
  Elsif 操作_In = 2 Then
    If 出诊记录ID_In is Null then
       Delete 挂号序号状态 Where 机器名 = v_机器名 And 操作员姓名 = 操作员姓名_In And 状态 = 5;
       
       Update 临床出诊序号控制 A set A.挂号状态=0,A.锁号时间=NULL,A.操作员姓名=NULL,A.工作站IP=NULL,A.工作站名称=NULL,A.类型=NULL,A.名称=NULL,A.备注=NULL
       Where A.工作站名称 =v_机器名 And A.工作站IP=v_工作站IP And A.操作员姓名 = 操作员姓名_In And A.挂号状态 = 5 And A.锁号时间 > Trunc(Sysdate);
    Else
      Update 临床出诊序号控制 A set A.挂号状态=0,A.锁号时间=NULL,A.操作员姓名=NULL,A.工作站IP=NULL,A.工作站名称=NULL,A.类型=NULL,A.名称=NULL,A.备注=NULL
      Where A.工作站名称 =v_机器名 And A.工作站IP=v_工作站IP And A.操作员姓名 = 操作员姓名_In And A.挂号状态 = 5 And A.锁号时间 > Trunc(Sysdate)
        And Exists (Select 1 From 临床出诊记录 B Where A.记录ID=B.ID And B.是否序号控制 = 1);
      
      Update 临床出诊序号控制 A set A.挂号状态=4,A.锁号时间=NULL,A.操作员姓名=NULL,A.工作站IP=NULL,A.工作站名称=NULL,A.类型=NULL,A.名称=NULL,A.备注=NULL
      Where A.工作站名称 =v_机器名 And A.工作站IP=v_工作站IP And A.操作员姓名 = 操作员姓名_In And A.挂号状态 = 5 And A.锁号时间 > Trunc(Sysdate)
        And Exists (Select 1 From 临床出诊记录 B Where A.记录ID=B.ID And B.是否序号控制 = 0);
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_挂号序号状态_Lock;
/

--137272:李南春,2019-02-22,预约提前接收锁定序号
Create Or Replace Procedure Zl_预约挂号接收_Insert
(
  No_In            门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  领用id_In        票据使用明细.领用id%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  结算方式_In      病人预交记录.结算方式%Type, --现金的结算名称
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额
  预交支付_In      病人预交记录.冲预交%Type, --挂号时使用的预交金额
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额
  发生时间_In      门诊费用记录.发生时间%Type,
  号序_In          挂号序号状态.序号%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  生成队列_In      Number := 0,
  登记时间_In      门诊费用记录.登记时间%Type := Null,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  冲预交病人ids_In Varchar2 := Null,
  三方调用_In      Number := 0,
  更新交款余额_In  Number := 1,
  摘要_In          病人挂号记录.摘要%Type := Null,
  收费单_In        病人挂号记录.收费单%Type := Null
) As
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  --更新交款余额_In:0-在zl_人员缴款余额_Update 中更新 1-在本过程中更新
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select 病人id, No, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, Id, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select /*+cardinality(d,10)*/
                                        d.Column_Value
                                       From Table(f_Num2list(v_冲预交病人ids)) d) And Nvl(预交类别, 2) = 1 Having
     Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By No, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 结帐id, No;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  Err_Special Exception;

  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  v_队列名称 排队叫号队列.队列名称%Type;
  v_号别     门诊费用记录.计算单位%Type;
  v_号序     门诊费用记录.发药窗口%Type;
  v_排队号码 排队叫号队列.排队号码 %Type;
  v_预约方式 病人挂号记录.预约方式 %Type;

  n_病人余额      病人预交记录.金额%Type;
  n_预交金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  v_冲预交病人ids Varchar2(4000);

  n_挂号id         病人挂号记录.Id%Type;
  n_分诊台签到排队 Number;
  n_组id           财务缴款分组.Id%Type;
  n_Count          Number(18);
  n_排队           Number;
  n_当天排队       Number;
  n_当前金额       病人预交记录.金额%Type;
  n_预交id         病人预交记录.Id%Type;

  d_Date     Date;
  d_预约时间 门诊费用记录.发生时间%Type;
  d_发生时间 Date;
  d_排队时间 Date;
  n_时段     Number := 0;
  n_存在     Number := 0;
  v_排队序号 排队叫号队列.排队序号%Type;
  n_结算模式 病人信息.结算模式%Type;

  v_付款方式   病人挂号记录.医疗付款方式%Type;
  v_操作员姓名 病人挂号记录.接收人%Type;
  n_接收模式   Number := 0;
Begin
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
  n_接收模式      := Nvl(zl_GetSysParameter('预约接收模式', 1111), 0);

  --获取结算方式名称
  Begin
    Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
  Exception
    When Others Then
      v_现金 := '现金';
  End;
  Begin
    Select 名称 Into v_个人帐户 From 结算方式 Where 性质 = 3;
  Exception
    When Others Then
      v_个人帐户 := '个人帐户';
  End;
  If 登记时间_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := 登记时间_In;
  End If;

  --更新挂号序号状态
  Begin
    Select 号别, 号序, Trunc(发生时间), 发生时间, 预约方式
    Into v_号别, v_号序, d_预约时间, d_发生时间, v_预约方式
    From 病人挂号记录
    Where 记录性质 = 2 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      Select Max(接收人) Into v_操作员姓名 From 病人挂号记录 Where 记录性质 = 2 And 记录状态 In (1, 3) And NO = No_In;
      If v_操作员姓名 Is Null Then
        v_Err_Msg := '当前预约挂号单已被取消';
        Raise Err_Item;
      Else
        If v_操作员姓名 = 操作员姓名_In Then
          v_Err_Msg := '当前预约挂号单已被接收';
          Raise Err_Special;
        Else
          v_Err_Msg := '当前预约挂号单已被其它人接收';
          Raise Err_Item;
        End If;
      End If;
  End;

  --判断是否分时段
  Begin
    Select 1
    Into n_时段
    From Dual
    Where Exists (Select 1
           From 挂号安排时段 A, 挂号安排 B
           Where a.安排id = b.Id And b.号码 = v_号别 And Rownum < 2
           Union All
           Select 1
           From 挂号计划时段 C, 挂号安排计划 D 　
           Where c.计划id = d.Id And d.号码 = v_号别 And d.生效时间 > Sysdate And Rownum < 2);
  Exception
    When Others Then
      n_时段 := 0;
  End;
  --分时段的号别，只能当天接收
  If n_时段 = 1 And 三方调用_In = 0 And n_接收模式 = 0 Then
    If Trunc(发生时间_In) <> Trunc(Sysdate) Then
      v_Err_Msg := '分时段的预约挂号单只能当天接收！';
      Raise Err_Item;
    End If;
  End If;

  If n_时段 = 0 And 三方调用_In = 0 Then
    If n_接收模式 = 0 Then
      If Trunc(发生时间_In) = Trunc(Sysdate) Then
        d_发生时间 := 发生时间_In;
      Else
        d_发生时间 := Sysdate;
      End If;
    Else
      d_发生时间 := 发生时间_In;
    End If;
  Else
    If Not 发生时间_In Is Null Then
      d_发生时间 := 发生时间_In;
    End If;
  End If;
  If Not v_号序 Is Null Then
    If 号序_In Is Null Then
      Delete 挂号序号状态 Where 号码 = v_号别 And Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序;
    Else
      If Trunc(d_预约时间) <> Trunc(Sysdate) And n_接收模式 = 0 Then
      
        If n_时段 = 0 And 三方调用_In = 0 Then
          --提前接收或延迟接收
          Delete 挂号序号状态 Where 号码 = v_号别 And Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序;
          --检查是否被锁定,n_存在 表示是否被他人使用
          Select Count(1)
          Into n_存在
          From 挂号序号状态
          Where 号码 = v_号别 And Trunc(日期) = Trunc(Sysdate) And 序号 = 号序_In And (状态 <> 5 Or 状态 = 5 And 操作员姓名 <> 操作员姓名_In);
          If n_存在 = 0 Then
            Delete 挂号序号状态 Where 号码 = v_号别 And Trunc(日期) = Trunc(Sysdate) And 序号 = 号序_In;
            v_号序 := 号序_In;
          Else
            Begin
              Select 1 Into n_存在 From 挂号序号状态 Where 号码 = v_号别 And 日期 = Trunc(Sysdate) And 序号 = v_号序;
            Exception
              When Others Then
                n_存在 := 0;
            End;
          End IF;
          If n_存在 = 0 Then
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
            Values
              (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
          Else
            --号码已被使用的情况
            Begin
              v_号序 := 1;
              Insert Into 挂号序号状态
                (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
              Values
                (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
            Exception
              When Others Then
                Select Min(序号 + 1)
                Into v_号序
                From 挂号序号状态 A
                Where 号码 = v_号别 And 日期 = Trunc(Sysdate) And Not Exists
                 (Select 1 From 挂号序号状态 Where 号码 = a.号码 And 日期 = a.日期 And 序号 = a.序号 + 1);
                Insert Into 挂号序号状态
                  (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
                Values
                  (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
            End;
          End If;
        Else
          Update 挂号序号状态
          Set 状态 = 1, 登记时间 = Sysdate
          Where Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序 And 号码 = v_号别 And 状态 = 2;
          If Sql% NotFound Then
            Begin
              Insert Into 挂号序号状态
                (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
              Values
                (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
            Exception
              When Others Then
                v_Err_Msg := '序号' || v_号序 || '已被其它人使用,请重新选择一个序号.';
                Raise Err_Item;
            End;
          End If;
        
        End If;
      
      Else
        Update 挂号序号状态
        Set 序号 = 号序_In, 状态 = 1, 登记时间 = Sysdate
        Where 号码 = v_号别 And Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
            Values
              (v_号别, Trunc(d_发生时间), v_号序, 1, 操作员姓名_In, Sysdate);
          Exception
            When Others Then
              v_Err_Msg := '序号' || v_号序 || '已被其它人使用,请重新选择一个序号.';
              Raise Err_Item;
          End;
        End If;
      End If;
    End If;
  Else
    If Not 号序_In Is Null Then
      Delete 挂号序号状态
      Where 号码 = v_号别 And Trunc(日期) = Trunc(Sysdate) And 序号 = 号序_In And 状态 = 5 And 操作员姓名 = 操作员姓名_In;
      Begin
        Insert Into 挂号序号状态
          (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
        Values
          (v_号别, Trunc(Sysdate), 号序_In, 1, 操作员姓名_In, Sysdate);
      Exception
        When Others Then
          v_Err_Msg := '序号' || 号序_In || '已被其它人使用,请重新选择一个序号.';
          Raise Err_Item;
      End;
      v_号序 := 号序_In;
    Else
      v_号序 := Null;
    End If;
  End If;

  --更新门诊费用记录
  Update 门诊费用记录
  Set 记录状态 = 1, 实际票号 = Decode(Nvl(记帐费用_In, 0), 1, Null, 票据号_In), 结帐id = Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In),
      结帐金额 = Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额), 发药窗口 = 诊室_In, 病人id = 病人id_In, 标识号 = 门诊号_In, 姓名 = 姓名_In, 年龄 = 年龄_In,
      性别 = 性别_In, 付款方式 = 付款方式_In, 费别 = 费别_In, 发生时间 = d_发生时间, 登记时间 = d_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In,
      缴款组id = n_组id, 记帐费用 = Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 摘要 = Decode(收费单_In, Null, Nvl(摘要_In, 摘要), '划价:' || 收费单_In)
  Where 记录性质 = 4 And 记录状态 = 0 And NO = No_In;

  --病人挂号记录
  Update 病人挂号记录
  Set 接收人 = 操作员姓名_In, 接收时间 = d_Date, 记录性质 = 1, 病人id = 病人id_In, 门诊号 = 门诊号_In, 发生时间 = d_发生时间, 姓名 = 姓名_In, 性别 = 性别_In,
      年龄 = 年龄_In, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 险类 = Decode(Nvl(险类_In, 0), 0, Null, 险类_In), 号序 = v_号序, 诊室 = 诊室_In,
      摘要 = Nvl(摘要_In, 摘要), 收费单 = 收费单_In
  Where 记录状态 = 1 And NO = No_In And 记录性质 = 2
  Returning ID Into n_挂号id;
  If Sql%NotFound Then
    Begin
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Begin
        Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
      Exception
        When Others Then
          v_付款方式 := Null;
      End;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         摘要, 号序, 预约, 预约方式, 接收人, 接收时间, 预约时间, 险类, 医疗付款方式, 收费单)
        Select n_挂号id, No_In, 1, 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 计算单位, 加班标志, 诊室_In, Null, 执行部门id, 执行人, 0, Null,
               登记时间, 发生时间, 操作员编号, 操作员姓名, Nvl(摘要_In, 摘要), v_号序, 1, Substr(结论, 1, 10) As 预约方式, 操作员姓名_In,
               Nvl(登记时间_In, Sysdate), 发生时间, Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, 收费单_In
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '由于并发原因,单据号为【' || No_In || '】的病人' || 姓名_In || '已经被接收';
        Raise Err_Item;
    End;
  End If;

  --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
  If Nvl(生成队列_In, 0) <> 0 Then
    For v_挂号 In (Select ID, 姓名, 诊室, 执行人, 执行部门id, 发生时间, 号别, 号序 From 病人挂号记录 Where NO = No_In) Loop
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(v_挂号.执行部门id, 0)));
      If Nvl(n_分诊台签到排队, 0) = 0 Then
        Begin
          Select 1,
                 Case
                   When 排队时间 < Trunc(Sysdate) Then
                    1
                   Else
                    0
                 End
          Into n_排队, n_当天排队
          From 排队叫号队列
          Where 业务类型 = 0 And 业务id = v_挂号.Id And Rownum <= 1;
        Exception
          When Others Then
            n_排队 := 0;
        End;
        If n_排队 = 0 Then
          --产生队列
          --按”执行部门”产生队列
          n_挂号id   := v_挂号.Id;
          v_队列名称 := v_挂号.执行部门id;
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, n_挂号id, v_挂号.号别 || '|' || v_挂号.号序);
          v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        
          --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
          d_排队时间 := Zl_Get_Queuedate(n_挂号id, v_挂号.号别, v_挂号.号序, v_挂号.发生时间);
          --   队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In,
          Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, v_挂号.执行部门id, v_排队号码, Null, 姓名_In, 病人id_In, v_挂号.诊室, v_挂号.执行人, d_排队时间,
                           v_预约方式, Null, v_排队序号);
        Elsif Nvl(n_当天排队, 0) = 1 Then
          --更新队列号
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, v_挂号.Id, v_挂号.号别 || '|' || Nvl(v_挂号.号序, 0));
          v_排队序号 := Zlgetsequencenum(0, v_挂号.Id, 1);
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人, v_排队号码, v_排队序号);
        
        Else
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人);
        End If;
        --预约接收时，改变记录标志
        Update 病人挂号记录 Set 记录标志 = 1 Where ID = n_挂号id;
      End If;
    End Loop;
  End If;

  --汇总结算到病人预交记录
  If (Nvl(现金支付_In, 0) <> 0 Or (Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0)) And
     Nvl(记帐费用_In, 0) = 0 Then
    Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算序号,
       结算性质)
    Values
      (n_预交id, 4, 1, No_In, 病人id_In, Nvl(结算方式_In, v_现金), Nvl(现金支付_In, 0), d_Date, 操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费',
       n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 结帐id_In, 4);
  
    If Nvl(结算卡序号_In, 0) <> 0 And Nvl(现金支付_In, 0) <> 0 Then
      Zl_病人卡结算记录_支付(结算卡序号_In, 卡号_In, 0, 现金支付_In, n_预交id, 操作员编号_In, 操作员姓名_In, d_Date);
    End If;
  
  End If;

  --对于就诊卡通过预交金挂号
  If Nvl(预交支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    Select Nvl(Sum(Nvl(预交余额, 0) - Nvl(费用余额, 0)), 0)
    Into n_病人余额
    From 病人余额
    Where 病人id In (Select /*+cardinality(d,10)*/
                    d.Column_Value
                   From Table(f_Num2list(v_冲预交病人ids)) d) And Nvl(性质, 0) = 1 And Nvl(类型, 0) = 1;
    if n_病人余额 < 预交支付_In Then
      v_Err_Msg := '病人的当前预交余额为 ' || Ltrim(To_Char(n_病人余额, '9999999990.00')) || '，小于本次支付金额 ' ||
                   Ltrim(To_Char(预交支付_In, '9999999990.00')) || '，支付失败！';
      Raise Err_Item;
    End if;
    
    n_预交金额 := 预交支付_In;
    For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
      n_当前金额 := Case
                  When r_Deposit.金额 - n_预交金额 < 0 Then
                   r_Deposit.金额
                  Else
                   n_预交金额
                End;
      If r_Deposit.结帐id = 0 Then
        --第一次冲预交(填上结帐ID,金额为0)
        Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.原预交id;
      End If;
      --冲上次剩余额
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结算序号, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 登记时间_In,
               操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结帐id_In, 4
        From 病人预交记录
        Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
    
      --更新病人预交余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
      Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (r_Deposit.病人id, Nvl(1, 2), -1 * n_当前金额, 1);
        n_返回值 := -1 * n_当前金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
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
    IF n_预交金额 > 0 Then
      v_Err_Msg := '病人的当前预交余额小于本次支付金额 ' || Ltrim(To_Char(预交支付_In, '9999999990.00')) || '，不能继续操作！';
      Raise Err_Item;
    End IF;
  End If;

  --对于医保挂号
  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
       预交类别, 结算序号, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 4, 1, No_In, 病人id_In, v_个人帐户, 个帐支付_In, d_Date, 操作员编号_In, 操作员姓名_In, 结帐id_In, '医保挂号', n_组id,
       Null, Null, Null, Null, Null, Null, Null, 结帐id_In, 4);
  End If;

  --相关汇总表的处理
  --人员缴款余额
  If Nvl(现金支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 And Nvl(更新交款余额_In, 1) = 1 Then
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 现金支付_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(结算方式_In, v_现金)
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (操作员姓名_In, Nvl(结算方式_In, v_现金), 1, 现金支付_In);
      n_返回值 := 现金支付_In;
    
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = Nvl(结算方式_In, v_现金) And Nvl(余额, 0) = 0;
    End If;
  End If;

  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 And Nvl(更新交款余额_In, 1) = 1 Then
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 个帐支付_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
      n_返回值 := 个帐支付_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额 Where 收款员 = 操作员姓名_In And 性质 = 1 And Nvl(余额, 0) = 0;
    End If;
  End If;

  If Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
    For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                 From 门诊费用记录
                 Where 记录性质 = 4 And 记录状态 = 1 And NO = No_In And Nvl(记帐费用, 0) = 1) Loop
      --病人余额
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 费用余额, 预交余额)
        Values
          (病人id_In, 1, 1, Nvl(c_费用.实收金额, 0), 0);
      End If;
    
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And 收入项目id + 0 = c_费用.收入项目id And
            来源途径 + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, Nvl(c_费用.实收金额, 0));
      End If;
    End Loop;
  End If;
  If Nvl(病人id_In, 0) <> 0 Then
    n_结算模式 := 0;
    Update 病人信息
    Set 就诊时间 = d_发生时间, 就诊状态 = 1, 就诊诊室 = 诊室_In
    Where 病人id = 病人id_In
    Returning Nvl(结算模式, 0) Into n_结算模式;
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
      If n_结算模式 = 1 And Nvl(结算模式_In, 0) = 0 Then
        --病人已经是"先诊疗后结算的",本次是"先结算后诊疗的",则检查是否存在未结数据
        Select Count(1)
        Into n_Count
        From 病人未结费用
        Where 病人id = 病人id_In And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --存在未结算数据，必须先结算后才允许执行
          v_Err_Msg := '当前病人的就诊模式为先诊疗后结算且存在未结费用，不允许调整该病人的就诊模式,你可以先对未结费用结帐后再挂号或不调整病人的就诊模式!';
          Raise Err_Item;
        End If;
        --检查
        --未发生医嘱业务的（即当时就挂号的,需要保证同一次的就诊模式是一至的(程序已经检查，不用再处理)
      End If;
      Update 病人信息 Set 结算模式 = 结算模式_In Where 病人id = 病人id_In;
    End If;
  End If;

  --病人担保信息
  If 病人id_In Is Not Null Then
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Nvl(在院, 0) = 0 And Exists
     (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = d_Date
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, d_Date) >= d_Date;
    End If;
  End If;
  --消息推送
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 1, n_挂号id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_预约挂号接收_Insert;
/

--137272:李南春,2019-02-22,预约提前接收锁定序号
Create Or Replace Procedure Zl_预约挂号接收_出诊_Insert
(
  No_In            门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  领用id_In        票据使用明细.领用id%Type,

  结帐id_In        门诊费用记录.结帐id%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  结算方式_In      Varchar2, --现金的结算名称
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额
  预交支付_In      病人预交记录.冲预交%Type, --挂号时使用的预交金额
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额
  发生时间_In      门诊费用记录.发生时间%Type,
  号序_In          挂号序号状态.序号%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  生成队列_In      Number := 0,
  登记时间_In      门诊费用记录.登记时间%Type := Null,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  冲预交病人ids_In Varchar2 := Null,
  三方调用_In      Number := 0,
  更新交款余额_In  Number := 1, --是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况
  摘要_In          病人挂号记录.摘要%Type := Null,
  收费单_In        病人挂号记录.收费单%Type := Null
) As
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select 病人id, No, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, Id, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select /*+cardinality(d,10)*/
                                        d.Column_Value
                                       From Table(f_Num2list(v_冲预交病人ids)) d) And Nvl(预交类别, 2) = 1 Having
     Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By No, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 结帐id, No;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  Err_Special Exception;
  v_操作员姓名 病人挂号记录.接收人%Type;
  v_现金       结算方式.名称%Type;
  v_个人帐户   结算方式.名称%Type;
  v_队列名称   排队叫号队列.队列名称%Type;
  v_号别       门诊费用记录.计算单位%Type;
  v_号序       门诊费用记录.发药窗口%Type;
  v_排队号码   排队叫号队列.排队号码 %Type;
  v_预约方式   病人挂号记录.预约方式 %Type;

  n_病人余额      病人预交记录.金额%Type;
  n_预交金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  v_冲预交病人ids Varchar2(4000);

  n_挂号id         病人挂号记录.Id%Type;
  n_分诊台签到排队 Number;
  n_组id           财务缴款分组.Id%Type;
  n_Count          Number(18);
  n_排队           Number;
  n_当天排队       Number;
  n_当前金额       病人预交记录.金额%Type;
  n_预交id         病人预交记录.Id%Type;

  d_Date         Date;
  d_预约时间     门诊费用记录.发生时间%Type;
  d_发生时间     Date;
  d_排队时间     Date;
  n_时段         Number := 0;
  n_存在         Number := 0;
  v_结算内容     Varchar2(2000);
  v_当前结算     Varchar2(500);
  n_结算金额     病人预交记录.冲预交%Type;
  v_结算号码     病人预交记录.结算号码%Type;
  v_结算方式     病人预交记录.结算方式%Type;
  n_三方卡标志   Number(3);
  v_排队序号     排队叫号队列.排队序号%Type;
  n_结算模式     病人信息.结算模式%Type;
  v_付款方式     病人挂号记录.医疗付款方式%Type;
  n_接收模式     Number := 0;
  n_出诊记录id   病人挂号记录.出诊记录id%Type;
  n_新出诊记录id 病人挂号记录.出诊记录id%Type;
  n_号源id       临床出诊记录.号源id%Type;
  n_预约顺序号   临床出诊序号控制.预约顺序号%Type;
  n_旧分时段     临床出诊记录.是否分时段%Type;
  n_旧序号控制   临床出诊记录.是否序号控制%Type;
  n_旧科室id     临床出诊记录.科室id%Type;
  n_旧项目id     临床出诊记录.项目id%Type;
  n_旧医生id     临床出诊记录.医生id%Type;
  n_挂号模式     Number(3);
  d_启用时间     Date;
  v_Paratemp     Varchar2(500);
  v_Registtemp   Varchar2(500);
  n_检查         Number(3);
  n_序号控制     临床出诊记录.是否序号控制%Type;
  v_旧上班时段   临床出诊记录.上班时段%Type;
Begin
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
  v_Paratemp      := Nvl(zl_GetSysParameter('挂号排班模式'), 0);
  n_接收模式      := Nvl(zl_GetSysParameter('预约接收模式', 1111), 0);
  n_挂号模式      := To_Number(Substr(v_Paratemp, 1, 1));
  If n_挂号模式 = 1 Then
    Begin
      d_启用时间 := To_Date(Substr(v_Paratemp, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_启用时间 := Null;
    End;
  End If;

  --获取结算方式名称
  Begin
    Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
  Exception
    When Others Then
      v_现金 := '现金';
  End;
  Begin
    Select 名称 Into v_个人帐户 From 结算方式 Where 性质 = 3;
  Exception
    When Others Then
      v_个人帐户 := '个人帐户';
  End;
  If 登记时间_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := 登记时间_In;
  End If;

  --更新挂号序号状态
  Begin
    Select 号别, 号序, Trunc(发生时间), 发生时间, 预约方式, 出诊记录id
    Into v_号别, v_号序, d_预约时间, d_发生时间, v_预约方式, n_出诊记录id
    From 病人挂号记录
    Where 记录性质 = 2 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      Select Max(接收人) Into v_操作员姓名 From 病人挂号记录 Where 记录性质 = 2 And 记录状态 In (1, 3) And NO = No_In;
      If v_操作员姓名 Is Null Then
        v_Err_Msg := '当前预约挂号单已被取消';
        Raise Err_Item;
      Else
        If v_操作员姓名 = 操作员姓名_In Then
          v_Err_Msg := '当前预约挂号单已被接收';
          Raise Err_Special;
        Else
          v_Err_Msg := '当前预约挂号单已被其它人接收';
          Raise Err_Item;
        End If;
      End If;
  End;

  --判断是否分时段
  Select Nvl(是否分时段, 0), 号源id, Nvl(是否序号控制, 0)
  Into n_时段, n_号源id, n_序号控制
  From 临床出诊记录
  Where ID = n_出诊记录id;

  If n_时段 = 1 And 三方调用_In = 0 And n_接收模式 = 0 Then
    If Trunc(发生时间_In) <> Trunc(Sysdate) Then
      v_Err_Msg := '分时段的预约挂号单只能当天接收！';
      Raise Err_Item;
    End If;
  End If;

  If n_时段 = 0 And 三方调用_In = 0 Then
    If n_接收模式 = 0 Then
      If Trunc(发生时间_In) = Trunc(Sysdate) Then
        d_发生时间 := 发生时间_In;
      Else
        d_发生时间 := Sysdate;
      End If;
    Else
      d_发生时间 := 发生时间_In;
    End If;
  Else
    If Not 发生时间_In Is Null Then
      d_发生时间 := 发生时间_In;
    End If;
  End If;

  If d_启用时间 Is Not Null Then
    If d_发生时间 < d_启用时间 Then
      v_Err_Msg := '当前预约挂号单属于出诊表排班模式安排，不能在' || To_Char(d_启用时间, 'yyyy-mm-dd hh24:mi:ss') || '之前接收!';
      Raise Err_Item;
    End If;
  End If;

  If Not v_号序 Is Null Then
    If 号序_In Is Null Then
      Update 临床出诊序号控制 Set 挂号状态 = 0 Where (序号 = v_号序 Or 备注 = v_号序) And 记录id = n_出诊记录id;
    Else
      If Trunc(d_预约时间) <> Trunc(Sysdate) And n_接收模式 = 0 Then
        If n_时段 = 0 And 三方调用_In = 0 Then
          --提前接收或延迟接收
          Update 临床出诊序号控制 Set 挂号状态 = 0 Where 序号 = v_号序 And 记录id = n_出诊记录id;
        
          Select 是否分时段, 是否序号控制, 科室id, 医生id, 项目id, 上班时段
          Into n_旧分时段, n_旧序号控制, n_旧科室id, n_旧医生id, n_旧项目id, v_旧上班时段
          From 临床出诊记录
          Where ID = n_出诊记录id;
          Begin
            Select ID
            Into n_新出诊记录id
            From 临床出诊记录
            Where 号源id = n_号源id And 是否分时段 = n_旧分时段 And 是否序号控制 = n_旧序号控制 And 科室id = n_旧科室id And
                  Nvl(医生id, 0) = Nvl(n_旧医生id, 0) And 上班时段 = v_旧上班时段 And Nvl(是否发布, 0) = 1 And 出诊日期 = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '接收当天没有对应的出诊安排,无法接收!';
              Raise Err_Item;
          End;
        
          --检查是否被锁定，n_存在 表示序号可用
          Select Count(1)
          Into n_存在
          From 临床出诊序号控制
          Where 记录id = n_新出诊记录id And 序号 = 号序_In And (Nvl(挂号状态, 0) = 0 Or Nvl(挂号状态, 0) = 5 And 操作员姓名 = 操作员姓名_In) And
                Rownum < 2;
          If n_存在 = 1 Then
            Update 临床出诊序号控制 Set 挂号状态 = 0 Where 记录id = n_新出诊记录id And 序号 = 号序_In;
            v_号序 := 号序_In;
          ElsIF v_号序 <> 号序_In Then
            Begin
              Select 1
              Into n_存在
              From 临床出诊序号控制
              Where 记录id = n_新出诊记录id And 序号 = v_号序 And Nvl(挂号状态, 0) = 0;
            Exception
              When Others Then
                n_存在 := 0;
            End;
          End if;
        
          If n_存在 = 1 Then
            Update 临床出诊序号控制
            Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
            Where 记录id = n_新出诊记录id And 序号 = v_号序 And Nvl(挂号状态, 0) = 0;
          Else
            --号码已被使用的情况
            Select Min(序号) Into v_号序 From 临床出诊序号控制 Where 记录id = n_新出诊记录id And Nvl(挂号状态, 0) = 0;
            If v_号序 Is Null Then
              v_Err_Msg := '接收当天没有可用序号,无法接收!';
              Raise Err_Item;
            End If;
            Update 临床出诊序号控制
            Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
            Where 记录id = n_新出诊记录id And 序号 = v_号序 And Nvl(挂号状态, 0) = 0;
          End If;
        Else
          Select 是否分时段, 是否序号控制, 科室id, 医生id, 项目id, 上班时段
          Into n_旧分时段, n_旧序号控制, n_旧科室id, n_旧医生id, n_旧项目id, v_旧上班时段
          From 临床出诊记录
          Where ID = n_出诊记录id;
          Begin
            Select ID
            Into n_新出诊记录id
            From 临床出诊记录
            Where 号源id = n_号源id And 是否分时段 = n_旧分时段 And 是否序号控制 = n_旧序号控制 And 科室id = n_旧科室id And
                  Nvl(医生id, 0) = Nvl(n_旧医生id, 0) And 上班时段 = v_旧上班时段 And Nvl(是否发布, 0) = 1 And 出诊日期 = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '接收当天没有对应的出诊安排,无法接收!';
              Raise Err_Item;
          End;
          Update 临床出诊序号控制
          Set 挂号状态 = 0, 操作员姓名 = 操作员姓名_In
          Where (序号 = v_号序 Or 备注 = v_号序) And 记录id = n_出诊记录id And Nvl(挂号状态, 0) = 2
          Returning 预约顺序号 Into n_预约顺序号;
        
          Update 临床出诊序号控制
          Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In, 预约顺序号 = n_预约顺序号
          Where 序号 = v_号序 And 记录id = n_新出诊记录id And (Nvl(挂号状态, 0) = 0 Or Nvl(挂号状态, 0) = 5 And 操作员姓名 = 操作员姓名_In);
          If Sql% RowCount = 0 Then
            v_Err_Msg := '接收当天序号' || v_号序 || '已被其它人使用,无法接收.';
            Raise Err_Item;
          End If;
        End If;
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
        Where (序号 = v_号序 Or 备注 = v_号序) And 记录id = n_出诊记录id;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '序号' || v_号序 || '已被其它人使用,请重新选择一个序号.';
          Raise Err_Item;
        End If;
      End If;
    End If;
  Else
    If Not 号序_In Is Null Then
      If Trunc(d_预约时间) <> Trunc(Sysdate) And n_接收模式 = 0 Then
        Select 是否分时段, 是否序号控制, 科室id, 医生id, 项目id, 上班时段
        Into n_旧分时段, n_旧序号控制, n_旧科室id, n_旧医生id, n_旧项目id, v_旧上班时段
        From 临床出诊记录
        Where ID = n_出诊记录id;
        Begin
          Select ID
          Into n_新出诊记录id
          From 临床出诊记录
          Where 号源id = n_号源id And 是否分时段 = n_旧分时段 And 是否序号控制 = n_旧序号控制 And 科室id = n_旧科室id And
                Nvl(医生id, 0) = Nvl(n_旧医生id, 0) And 上班时段 = v_旧上班时段 And Nvl(是否发布, 0) = 1 And 出诊日期 = Trunc(Sysdate) And
                Rownum < 2;
        Exception
          When Others Then
            v_Err_Msg := '接收当天没有对应的出诊安排,无法接收!';
            Raise Err_Item;
        End;
        Update 临床出诊序号控制
        Set 挂号状态 = 0, 操作员姓名 = 操作员姓名_In
        Where (序号 = 号序_In Or 备注 = 号序_In) And 记录id = n_出诊记录id And Nvl(挂号状态, 0) = 2
        Returning 预约顺序号 Into n_预约顺序号;
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In, 预约顺序号 = n_预约顺序号
        Where 序号 = 号序_In And 记录id = n_新出诊记录id And (Nvl(挂号状态, 0) = 0 Or Nvl(挂号状态, 0) = 5 And 操作员姓名 = 操作员姓名_In);
        If Sql%RowCount = 0 Then
          v_Err_Msg := '接收当天序号' || 号序_In || '已被其它人使用,无法接收.';
          Raise Err_Item;
        End If;
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
        Where (序号 = 号序_In Or 备注 = 号序_In) And 记录id = n_出诊记录id;
      End If;
      v_号序 := 号序_In;
    Else
      v_号序 := Null;
    End If;
  End If;
  
  --检查挂号项目,因为zl_Custom_GetRegeventItem项目可能不一致，只检查科室
  Select Count(1)
  Into n_Count
  From 临床出诊记录 a, 门诊费用记录 b
  Where a.Id = Nvl(n_新出诊记录id, n_出诊记录id) And b.No = No_In And b.序号 = 1 And a.科室id = b.执行部门id;
  If n_Count = 0 Then
    v_Err_Msg := '挂号科室不一致，无法接收！';
    Raise Err_Item;
  End If;

  --更新门诊费用记录
  Update 门诊费用记录
  Set 记录状态 = 1, 实际票号 = Decode(Nvl(记帐费用_In, 0), 1, Null, 票据号_In), 结帐id = Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In),
      结帐金额 = Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额), 发药窗口 = 诊室_In, 病人id = 病人id_In, 标识号 = 门诊号_In, 姓名 = 姓名_In, 年龄 = 年龄_In,
      性别 = 性别_In, 付款方式 = 付款方式_In, 费别 = 费别_In, 发生时间 = d_发生时间, 登记时间 = d_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In,
      缴款组id = n_组id, 记帐费用 = Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 摘要 = Decode(收费单_In, Null, Nvl(摘要_In, 摘要), '划价:' || 收费单_In)
  Where 记录性质 = 4 And 记录状态 = 0 And NO = No_In;

  v_Registtemp := zl_GetSysParameter('挂号排班模式');
  If Substr(v_Registtemp, 1, 1) = 1 Then
    Begin
      If To_Date(Substr(v_Registtemp, 3), 'yyyy-mm-dd hh24:mi:ss') > d_发生时间 Then
        v_Err_Msg := '接收时间' || To_Char(d_发生时间, 'yyyy-mm-dd hh24:mi:ss') || '未启用出诊表排班模式,目前无法接收!';
        Raise Err_Item;
      End If;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select 1
      Into n_检查
      From 临床出诊记录
      Where ID = Nvl(n_新出诊记录id, n_出诊记录id) And d_发生时间 Between 停诊开始时间 And 停诊终止时间;
    Exception
      When Others Then
        n_检查 := 0;
    End;
    If n_检查 = 1 And Not (n_时段 = 1 And n_序号控制 = 1) Then
      v_Err_Msg := '接收时间' || To_Char(d_发生时间, 'yyyy-mm-dd hh24:mi:ss') || '的安排已经被停诊,无法接收!';
      Raise Err_Item;
    End If;
  End If;

  --病人挂号记录
  Update 病人挂号记录
  Set 接收人 = 操作员姓名_In, 接收时间 = d_Date, 记录性质 = 1, 病人id = 病人id_In, 门诊号 = 门诊号_In, 发生时间 = d_发生时间, 姓名 = 姓名_In, 性别 = 性别_In,
      年龄 = 年龄_In, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 险类 = Decode(Nvl(险类_In, 0), 0, Null, 险类_In), 号序 = v_号序, 诊室 = 诊室_In,
      出诊记录id = Nvl(n_新出诊记录id, n_出诊记录id), 摘要 = Nvl(摘要_In, 摘要), 收费单 = 收费单_In
  Where 记录状态 = 1 And NO = No_In And 记录性质 = 2
  Returning ID Into n_挂号id;
  If Sql%NotFound Then
    Begin
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Begin
        Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
      Exception
        When Others Then
          v_付款方式 := Null;
      End;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         摘要, 号序, 预约, 预约方式, 接收人, 接收时间, 预约时间, 险类, 医疗付款方式, 出诊记录id, 收费单)
        Select n_挂号id, No_In, 1, 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 计算单位, 加班标志, 诊室_In, Null, 执行部门id, 执行人, 0, Null,
               登记时间, 发生时间, 操作员编号, 操作员姓名, Nvl(摘要_In, 摘要), v_号序, 1, Substr(结论, 1, 10) As 预约方式, 操作员姓名_In,
               Nvl(登记时间_In, Sysdate), 发生时间, Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, Nvl(n_新出诊记录id, n_出诊记录id),
               收费单_In
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '由于并发原因,单据号为【' || No_In || '】的病人' || 姓名_In || '已经被接收';
        Raise Err_Item;
    End;
  End If;

  --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
  If Nvl(生成队列_In, 0) <> 0 Then
    For v_挂号 In (Select ID, 姓名, 诊室, 执行人, 执行部门id, 发生时间, 号别, 号序 From 病人挂号记录 Where NO = No_In) Loop
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(v_挂号.执行部门id, 0)));
      If Nvl(n_分诊台签到排队, 0) = 0 Then
        Begin
          Select 1,
                 Case
                   When 排队时间 < Trunc(Sysdate) Then
                    1
                   Else
                    0
                 End
          Into n_排队, n_当天排队
          From 排队叫号队列
          Where 业务类型 = 0 And 业务id = v_挂号.Id And Rownum <= 1;
        Exception
          When Others Then
            n_排队 := 0;
        End;
        If n_排队 = 0 Then
          --产生队列
          --按”执行部门”产生队列
          n_挂号id   := v_挂号.Id;
          v_队列名称 := v_挂号.执行部门id;
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, n_挂号id, v_挂号.号别 || '|' || v_挂号.号序);
          v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        
          --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
          d_排队时间 := Zl_Get_Queuedate(n_挂号id, v_挂号.号别, v_挂号.号序, v_挂号.发生时间);
          --   队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In,
          Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, v_挂号.执行部门id, v_排队号码, Null, 姓名_In, 病人id_In, v_挂号.诊室, v_挂号.执行人, d_排队时间,
                           v_预约方式, Null, v_排队序号);
        Elsif Nvl(n_当天排队, 0) = 1 Then
          --更新队列号
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, v_挂号.Id, v_挂号.号别 || '|' || Nvl(v_挂号.号序, 0));
          v_排队序号 := Zlgetsequencenum(0, v_挂号.Id, 1);
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人, v_排队号码, v_排队序号);
        
        Else
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人);
        End If;
      End If;
    End Loop;
  End If;

  --汇总结算到病人预交记录
  If Nvl(记帐费用_In, 0) = 0 Then
    If Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
         结算性质)
      Values
        (n_预交id, 4, 1, No_In, Decode(病人id_In, 0, Null, 病人id_In), v_现金, 0, 登记时间_In, 操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费',
         n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, Null, 4);
    End If;
    If Nvl(现金支付_In, 0) <> 0 Then
      v_结算内容 := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        v_结算号码 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_三方卡标志 := To_Number(v_当前结算);
      
        If n_三方卡标志 = 0 Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, No_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, Null, Null, Null, Null, Null, Null, 4, v_结算号码);
        Else
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, No_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, Null, 4, v_结算号码);
        
          If Nvl(结算卡序号_In, 0) <> 0 And Nvl(n_结算金额, 0) <> 0 Then
            Zl_病人卡结算记录_支付(结算卡序号_In, 卡号_In, 0, n_结算金额, n_预交id, 操作员编号_In, 操作员姓名_In, 登记时间_In);
          End If;
        End If;
      
        If Nvl(更新交款余额_In, 1) = 1 Then
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + n_结算金额
          Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金)
          Returning 余额 Into n_返回值;
        
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (操作员姓名_In, Nvl(v_结算方式, v_现金), 1, n_结算金额);
            n_返回值 := n_结算金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金) And 性质 = 1 And Nvl(余额, 0) = 0;
          End If;
        End If;
      
        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
      End Loop;
    End If;
  End If;

  --对于就诊卡通过预交金挂号
  If Nvl(预交支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    Select Nvl(Sum(Nvl(预交余额, 0) - Nvl(费用余额, 0)), 0)
    Into n_病人余额
    From 病人余额
    Where 病人id In (Select /*+cardinality(d,10)*/
                    d.Column_Value
                   From Table(f_Num2list(v_冲预交病人ids)) d) And Nvl(性质, 0) = 1 And Nvl(类型, 0) = 1;
    if n_病人余额 < 预交支付_In Then
      v_Err_Msg := '病人的当前预交余额为 ' || Ltrim(To_Char(n_病人余额, '9999999990.00')) || '，小于本次支付金额 ' ||
                   Ltrim(To_Char(预交支付_In, '9999999990.00')) || '，支付失败！';
      Raise Err_Item;
    End if;
    
    n_预交金额 := 预交支付_In;
    For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
      n_当前金额 := Case
                  When r_Deposit.金额 - n_预交金额 < 0 Then
                   r_Deposit.金额
                  Else
                   n_预交金额
                End;
      If r_Deposit.结帐id = 0 Then
        --第一次冲预交(填上结帐ID,金额为0)
        Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.原预交id;
      End If;
      --冲上次剩余额
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结算序号, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 登记时间_In,
               操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结帐id_In, 4
        From 病人预交记录
        Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
    
      --更新病人预交余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
      Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (r_Deposit.病人id, Nvl(1, 2), -1 * n_当前金额, 1);
        n_返回值 := -1 * n_当前金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
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
    IF n_预交金额 > 0 Then
      v_Err_Msg := '病人的当前预交余额小于本次支付金额 ' || Ltrim(To_Char(预交支付_In, '9999999990.00')) || '，不能继续操作！';
      Raise Err_Item;
    End IF;
  End If;

  --对于医保挂号
  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
       预交类别, 结算序号, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 4, 1, No_In, 病人id_In, v_个人帐户, 个帐支付_In, d_Date, 操作员编号_In, 操作员姓名_In, 结帐id_In, '医保挂号', n_组id,
       Null, Null, Null, Null, Null, Null, Null, 结帐id_In, 4);
  End If;

  --相关汇总表的处理
  --人员缴款余额
  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 And Nvl(更新交款余额_In, 1) = 1 Then
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 个帐支付_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
      n_返回值 := 个帐支付_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额 Where 收款员 = 操作员姓名_In And 性质 = 1 And Nvl(余额, 0) = 0;
    End If;
  End If;

  If Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
    For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                 From 门诊费用记录
                 Where 记录性质 = 4 And 记录状态 = 1 And NO = No_In And Nvl(记帐费用, 0) = 1) Loop
      --病人余额
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 费用余额, 预交余额)
        Values
          (病人id_In, 1, 1, Nvl(c_费用.实收金额, 0), 0);
      End If;
    
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And 收入项目id + 0 = c_费用.收入项目id And
            来源途径 + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, Nvl(c_费用.实收金额, 0));
      End If;
    End Loop;
  End If;
  If Nvl(病人id_In, 0) <> 0 Then
    n_结算模式 := 0;
    Update 病人信息
    Set 就诊时间 = d_发生时间, 就诊状态 = 1, 就诊诊室 = 诊室_In
    Where 病人id = 病人id_In
    Returning Nvl(结算模式, 0) Into n_结算模式;
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
      If n_结算模式 = 1 And Nvl(结算模式_In, 0) = 0 Then
        --病人已经是"先诊疗后结算的",本次是"先结算后诊疗的",则检查是否存在未结数据
        Select Count(1)
        Into n_Count
        From 病人未结费用
        Where 病人id = 病人id_In And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --存在未结算数据，必须先结算后才允许执行
          v_Err_Msg := '当前病人的就诊模式为先诊疗后结算且存在未结费用，不允许调整该病人的就诊模式,你可以先对未结费用结帐后再挂号或不调整病人的就诊模式!';
          Raise Err_Item;
        End If;
        --检查
        --未发生医嘱业务的（即当时就挂号的,需要保证同一次的就诊模式是一至的(程序已经检查，不用再处理)
      End If;
      Update 病人信息 Set 结算模式 = 结算模式_In Where 病人id = 病人id_In;
    End If;
  End If;

  --病人担保信息
  If 病人id_In Is Not Null Then
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Nvl(在院, 0) = 0 And Exists
     (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = d_Date
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, d_Date) >= d_Date;
    End If;
  End If;
  --消息推送
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 1, n_挂号id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_预约挂号接收_出诊_Insert;
/

--124583:李业庆,2019-02-21,部门发药,退药填写发药类型
Create Or Replace Procedure Zl_药品收发记录_批量发药
(
  Billinfo_In   In Varchar2, --格式:"id1,批次1|id2,批次2|....."
  Partid_In     In 药品收发记录.库房id%Type,
  People_In     In 药品收发记录.审核人%Type,
  Date_In       In 药品收发记录.审核日期%Type,
  发药方式_In   In 药品收发记录.发药方式%Type := 3,
  领药人_In     In 药品收发记录.领用人%Type := Null,
  汇总发药号_In In 药品收发记录.汇总发药号%Type := Null,
  Intdigit_In   In Number := 2,
  配药人_In     In 药品收发记录.配药人%Type := Null,
  核查人_In     In 药品收发记录.核查人%Type := Null
) Is
  --只读变量
  v_Infotmp     Varchar2(4000);
  v_Fields      Varchar2(4000);
  n_Billid      药品收发记录.Id%Type;
  n_批次        药品收发记录.批次%Type;
  Lng入出类别id Number(18);
  Int入出系数   Number;
  Int执行状态   Number;
  Int单据       药品收发记录.单据%Type;
  Strno         药品收发记录.No%Type;
  Lng库房id     药品收发记录.库房id%Type;
  Lng药品id     药品收发记录.药品id%Type;
  Lng费用id     药品收发记录.费用id%Type;
  Dbl差价率     Number;
  v_零售价      药品收发记录.零售价%Type;
  Int未发数     未发药品记录.未发数%Type;
  v_核查日期    药品收发记录.核查日期%Type;
  --可写变量
  Dbl实际数量 药品收发记录.实际数量%Type;
  Dbl实际金额 药品收发记录.零售金额%Type;
  Dbl成本金额 药品收发记录.成本金额%Type;
  Dbl实际差价 药品收发记录.差价%Type;
  --2002-07-31朱玉宝
  --LNGLAST批次 发药前确定的批次(已减可用数量)
  Str药名           Varchar2(200);
  Dbl可用数量       药品收发记录.填写数量%Type;
  Lnglast批次       药品收发记录.批次%Type;
  Lngcur批次        药品收发记录.批次%Type;
  Str批号           药品收发记录.批号%Type;
  Str效期           药品收发记录.效期%Type;
  n_上次供应商id    药品库存.上次供应商id%Type;
  n_上次采购价      药品库存.上次采购价%Type;
  v_上次产地        药品库存.上次产地%Type;
  d_上次生产日期    药品库存.上次生产日期%Type;
  v_批准文号        药品库存.批准文号%Type;
  n_记录状态        药品收发记录.记录状态%Type;
  n_平均成本价      药品库存.平均成本价%Type;
  n_发药方式        药品收发记录.发药方式%Type;
  v_摘要            药品收发记录.摘要%Type;
  Bln收费与发药分离 Number(1);
  v_Error           Varchar2(255);
  Err_Custom Exception;
  n_流通售价小数   Number;
  n_流通金额小数   Number;
  n_零差价管理模式 Number(1);
  n_药品零差价管理 Number(1);
  n_处方类型       未发药品记录.处方类型%Type;
Begin
  Select Sysdate Into v_核查日期 From Dual;
  If Billinfo_In Is Null Then
    v_Infotmp := Null;
  Else
    v_Infotmp := Billinfo_In || '|';
  End If;
  While v_Infotmp Is Not Null Loop
    --分解单据ID串
    v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
    n_Billid  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
    n_批次    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
    v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
  
    --获取该收发记录的单据、药品ID、库房ID,零售金额及实际数量、入出类别ID
    Begin
      Select a.单据, a.No, a.药品id, a.库房id, a.费用id, Nvl(a.零售价, 0), Nvl(a.零售金额, 0), Nvl(a.实际数量, 0) * Nvl(a.付数, 1), a.入出类别id,
             a.入出系数, Nvl(a.批次, 0), a.批号, a.效期, a.供药单位id, a.产地, a.生产日期, a.批准文号, Nvl(a.发药方式, 0), a.摘要, 记录状态
      Into Int单据, Strno, Lng药品id, Lng库房id, Lng费用id, v_零售价, Dbl实际金额, Dbl实际数量, Lng入出类别id, Int入出系数, Lnglast批次, Str批号, Str效期,
           n_上次供应商id, v_上次产地, d_上次生产日期, v_批准文号, n_发药方式, v_摘要, n_记录状态
      From 药品收发记录 A
      Where a.Id = n_Billid And a.审核日期 Is Null
      For Update Nowait;
    
      Select '[' || c.编码 || ']' || c.名称 Into Str药名 From 收费项目目录 C Where c.Id = Lng药品id;
    Exception
      When Others Then
        Int单据 := 0;
        v_Error := '已有其他用户在执行发药，不能重复操作！';
        Raise Err_Custom;
    End;
  
    --取流通业务精度位数
    --类别:1-药品 2-卫材
    --内容：2-零售价 4-金额
    --单位：药品:1-售价 5-金额单位
    Begin
      Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;
    Exception
      When Others Then
        n_流通金额小数 := 2;
    End;
  
    Begin
      Select 精度 Into n_流通售价小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 2 And 单位 = 1;
    Exception
      When Others Then
        n_流通售价小数 := 2;
    End;
  
    Select Zl_To_Number(Nvl(zl_GetSysParameter(275), '0')) Into n_零差价管理模式 From Dual;
  
    If n_发药方式 = -1 Or v_摘要 = '拒发' Then
      Int单据 := 0;
    End If;
  
    If Int单据 > 0 Then
      If Nvl(n_批次, 0) = 0 Then
        Lngcur批次 := Lnglast批次;
      Else
        Lngcur批次 := Nvl(n_批次, 0);
      End If;
    
      --检查是否已经填写库房
      Bln收费与发药分离 := 0;
      If Lng库房id Is Null Then
        Bln收费与发药分离 := 1;
      End If;
      Lng库房id := Partid_In;
    
      --取该批药品的批号
      Begin
        Select 上次批号, 效期, Nvl(可用数量, 0), 上次供应商id, 上次产地, 上次生产日期, 批准文号, 上次采购价
        Into Str批号, Str效期, Dbl可用数量, n_上次供应商id, v_上次产地, d_上次生产日期, v_批准文号, n_上次采购价
        From 药品库存
        Where 库房id + 0 = Lng库房id And 药品id = Lng药品id And 性质 = 1 And Nvl(批次, 0) = Lngcur批次;
      Exception
        When Others Then
          n_上次采购价 := 0;
          Dbl可用数量  := 0;
      End;
    
      --可用数量不足则退出
      If Lngcur批次 <> Nvl(Lnglast批次, 0) Then
        If Dbl可用数量 < Dbl实际数量 And Lngcur批次 <> 0 Then
          v_Error := Str药名 || '的可用数量不足，操作中止！';
          Raise Err_Custom;
        End If;
      End If;
    
      If n_零差价管理模式 <> 0 Then
        Select Nvl(是否零差价管理, 0) Into n_药品零差价管理 From 药品规格 Where 药品id = Lng药品id;
      End If;
    
      If n_记录状态 = 1 Then
        --原始发药记录，取最新价格
        n_平均成本价 := Zl_Fun_Getoutcost(Lng药品id, Lngcur批次, Lng库房id);
      
        If n_零差价管理模式 <> 0 And n_药品零差价管理 = 1 And (v_零售价 = n_平均成本价 Or Round(v_零售价, n_流通售价小数) = Round(n_平均成本价, n_流通售价小数)) Then
          Dbl成本金额 := Dbl实际金额;
        Else
          Dbl成本金额 := Round(n_平均成本价 * Nvl(Dbl实际数量, 0), n_流通金额小数);
        End If;
      Else
        --退药再发记录，取原始单据价格
        Select a.成本价
        Into n_平均成本价
        From 药品收发记录 A, 药品收发记录 B
        Where b.Id = n_Billid And a.单据 = b.单据 And a.No = b.No And a.库房id = b.库房id And a.药品id + 0 = b.药品id And
              a.序号 = b.序号 And Nvl(a.批次, 0) = Nvl(b.批次, 0) And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0);
      
        Dbl成本金额 := Round(n_平均成本价 * Nvl(Dbl实际数量, 0), n_流通金额小数);
      End If;
    
      Dbl实际差价 := Round(Dbl实际金额 - Dbl成本金额, n_流通金额小数);
    
      --更新药品收发记录的零售金额、成本金额及差价
      Update 药品收发记录
      Set 库房id = Lng库房id, 成本价 = n_平均成本价, 成本金额 = Dbl成本金额, 差价 = Dbl实际差价, 批次 = Lngcur批次, 批号 = Str批号, 效期 = Str效期,
          配药人 = 配药人_In, 核查人 = 核查人_In, 核查日期 = v_核查日期, 审核人 = People_In, 审核日期 = Date_In, 发药方式 = 发药方式_In, 领用人 = 领药人_In,
          汇总发药号 = 汇总发药号_In, 供药单位id = n_上次供应商id, 产地 = v_上次产地, 生产日期 = d_上次生产日期, 批准文号 = v_批准文号
      Where ID = n_Billid;
      --并发操作检查
      If Sql%RowCount = 0 Then
        v_Error := '要发药的药品记录"' || Str药名 || '"不存在，操作中止！';
        Raise Err_Custom;
      End If;
    
      --更新住院费用记录的执行状态(已执行)
      Select Decode(Sum(Nvl(付数, 1) * 实际数量), Null, 1, 0, 1, 2)
      Into Int执行状态
      From 药品收发记录
      Where 单据 = Int单据 And NO = Strno And 费用id = Lng费用id And 审核人 Is Null;
      Update 住院费用记录
      Set 执行状态 = Int执行状态, 执行人 = Decode(People_In, Null, Zl_Username, People_In), 执行时间 = Date_In, 执行部门id = Partid_In
      Where ID = Lng费用id;
    
      --更新未发药品记录(如果未发数为零则删除)
      Select Count(*)
      Into Int未发数
      From 药品收发记录
      Where 单据 = Int单据 And (库房id + 0 = Lng库房id Or 库房id Is Null) And NO = Strno And 审核人 Is Null And
            Nvl(LTrim(RTrim(摘要)), '小宝') <> '拒发';
    
      If Int未发数 = 0 Then
        Delete 未发药品记录
        Where NO = Strno And 单据 = Int单据 And (库房id + 0 = Lng库房id Or 库房id Is Null)
        Returning 处方类型 Into n_处方类型;
      
        --更新处方类型，按整张处方更新
        Update 药品收发记录 Set 注册证号 = n_处方类型 Where 单据 = Int单据 And NO = Strno And 库房id = Lng库房id;
      End If;
    
      --更新库存
      Zl_药品库存_Update(n_Billid, 2, 1);
    
      Zl_未审药品记录_Delete(n_Billid);
    
      --处理调价修正
      Zl_药品收发记录_调价修正(n_Billid);
    
      b_Message.Zlhis_Drug_005(Lng库房id, n_Billid);
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_批量发药;
/

--124583:李业庆,2019-02-21,部门发药,退药填写发药类型
Create Or Replace Procedure Zl_药品收发记录_部门退药
(
  Billid_In     In 药品收发记录.Id%Type,
  People_In     In 药品收发记录.审核人%Type,
  Date_In       In 药品收发记录.审核日期%Type,
  批号_In       In 药品库存.上次批号%Type := Null,
  效期_In       In 药品库存.效期%Type := Null,
  产地_In       In 药品库存.上次产地%Type := Null,
  退药数量_In   In 药品收发记录.实际数量%Type := Null,
  退药库房_In   In 药品收发记录.库房id%Type := Null,
  退药人_In     In 药品收发记录.领用人%Type := Null,
  Intdigit_In   In Number := 2,
  门诊_In       In Number := 2,
  汇总发药号_In In 药品收发记录.汇总发药号%Type := Null
) Is
  --只读变量
  Int记录状态   药品收发记录.记录状态%Type;
  Int执行状态   住院费用记录.执行状态%Type;
  Bln部分退药   Number;
  Lng入出类别id Number(18);
  Strno         药品收发记录.No%Type;
  Int单据       药品收发记录.单据%Type;
  Lng库房id     药品收发记录.库房id%Type;
  Lng药品id     药品收发记录.药品id%Type;
  Dbl实际数量   药品收发记录.实际数量%Type;
  Dbl实际金额   药品收发记录.零售金额%Type;
  Dbl实际成本   药品收发记录.成本金额%Type;
  Dbl实际差价   药品收发记录.差价%Type;
  Lng费用id     药品收发记录.费用id%Type;
  n_零售价      药品收发记录.零售价%Type;
  n_是否变价    Number;
  n_时价分批    Number;

  --20020731 Modified by zyb
  --处理退药时，分批核算性质改变后的处理
  Lng新批次 药品收发记录.批次%Type;
  Lng分批   药品规格.药房分批%Type;
  Lng批次   药品收发记录.批次%Type; --原批次

  Str批号        药品收发记录.批号%Type; --原批号
  Date效期       药品收发记录.效期%Type; --原效期
  n_上次供应商id 药品库存.上次供应商id%Type;
  n_上次采购价   药品库存.上次采购价%Type;
  v_上次产地     药品库存.上次产地%Type;
  v_原产地       药品库存.原产地%Type;
  d_上次生产日期 药品库存.上次生产日期%Type;
  v_批准文号     药品库存.批准文号%Type;

  n_记录性质   住院费用记录.记录性质%Type;
  v_收费类别   住院费用记录.收费类别%Type;
  n_付数       药品收发记录.付数%Type;
  n_原始数量   药品收发记录.实际数量%Type;
  v_冲销记录id 药品收发记录.Id%Type;
  Err_Custom Exception;
  v_Error    Varchar2(255);
  v_配药确认 药房配药控制.配药确认%Type;
  v_配药     药房配药控制.配药%Type;
  v_排队状态 Number(1);
  v_执行时间 药品收发记录.审核日期%Type;

Begin
  If 退药数量_In Is Not Null Then
    If 退药数量_In = 0 Then
      Return;
    End If;
  End If;

  --获取该收发记录的单据、药品ID、库房ID
  Select a.单据, a.No, a.库房id, a.药品id, a.费用id, a.入出类别id, a.记录状态, Nvl(a.批次, 0), a.批号, a.效期, a.供药单位id, a.产地, a.原产地, a.生产日期,
         a.批准文号, a.成本价, a.付数, Nvl(a.实际数量, 0) * Nvl(a.付数, 1) As 实际数量, a.零售价, Nvl(b.是否变价, 0) 是否变价
  Into Int单据, Strno, Lng库房id, Lng药品id, Lng费用id, Lng入出类别id, Int记录状态, Lng批次, Str批号, Date效期, n_上次供应商id, v_上次产地, v_原产地,
       d_上次生产日期, v_批准文号, n_上次采购价, n_付数, n_原始数量, n_零售价, n_是否变价
  From 药品收发记录 A, 收费项目目录 B
  Where a.药品id = b.Id And a.Id = Billid_In;

  Begin
    Select Nvl(配药确认, 0), Nvl(配药, 0)
    Into v_配药确认, v_配药
    From 药房配药控制
    Where 药房id = Lng库房id And Rownum = 1;
  
  Exception
    When Others Then
      v_配药确认 := 0;
      v_配药     := 0;
      Null;
  End;

  If v_配药确认 = 0 And v_配药 = 0 Then
    v_排队状态 := 2;
  Elsif v_配药确认 = 1 Then
    v_排队状态 := 0;
  Elsif v_配药 = 1 Then
    v_排队状态 := 1;
  End If;

  --获取该笔记录剩余未退数量、金额及差价
  --尽量避免金额及差价未出完的现象
  Select Sum(Nvl(实际数量, 0) * Nvl(付数, 1)), Sum(Nvl(零售金额, 0)), Sum(Nvl(成本金额, 0)), Sum(Nvl(差价, 0))
  Into Dbl实际数量, Dbl实际金额, Dbl实际成本, Dbl实际差价
  From 药品收发记录
  Where 审核人 Is Not Null And NO = Strno And 单据 = Int单据 And 序号 = (Select 序号 From 药品收发记录 Where ID = Billid_In);

  --如果允许退药数为零，表示已退药
  If Dbl实际数量 = 0 Then
    v_Error := '该单据已被其他操作员退药，请刷新后再试！';
    Raise Err_Custom;
  End If;
  If Nvl(退药数量_In, 0) > Dbl实际数量 Then
    v_Error := '该单据已被其他操作员部分退药，请刷新后再试！';
    Raise Err_Custom;
  End If;

  --获取该药品当前是否分批的信息
  Select Nvl(药房分批, 0) Into Lng分批 From 药品规格 Where 药品id = Lng药品id;
  --如果是部分退药，则重新计算零售金额及差价
  Bln部分退药 := 0;
  If Not (退药数量_In Is Null Or Nvl(退药数量_In, 0) = Dbl实际数量) Then
    Bln部分退药 := 1;
  End If;
  If Bln部分退药 = 1 Then
    Dbl实际金额 := Round(Dbl实际金额 * 退药数量_In / Dbl实际数量, Intdigit_In);
    Dbl实际成本 := Round(Dbl实际成本 * 退药数量_In / Dbl实际数量, Intdigit_In);
    Dbl实际差价 := Round(Dbl实际差价 * 退药数量_In / Dbl实际数量, Intdigit_In);
    Dbl实际数量 := 退药数量_In;
  End If;

  If n_原始数量 = 退药数量_In Then
    Dbl实际数量 := 退药数量_In / n_付数;
  Else
    n_付数 := 1;
  End If;

  --lng分批:0-不分批;1-分批;2-原分批，现不分批，按不分批处理;3-原不分批，现分批，产生新批次
  If Lng分批 = 0 And Lng批次 <> 0 Then
    --原分批，现不分批，按不分批处理
    Lng分批 := 2;
  Elsif Lng分批 <> 0 And Lng批次 = 0 Then
    --原不分批,现分批,产生新的批次，并在新产生的发药记录中使用
    Lng分批 := 3;
  Else
    If Lng批次 = 0 Then
      Lng分批 := 0;
    Else
      Lng分批 := 1;
    End If;
  End If;
  --判断是否时价分批
  If (Lng分批 = 1 Or Lng分批 = 3) And n_是否变价 = 1 Then
    n_时价分批 := 1;
  Else
    n_时价分批 := 0;
  End If;

  --记录状态的含义有所变化
  --冲销的记录状态        :iif(int记录状态=1,0,1)+1
  --被冲销的记录状态        :iif(int记录状态=1,0,1)+2
  --等待发药的记录状态    :iif(int记录状态=1,0,1)+3

  --产生冲销记录
  Select 药品收发记录_Id.Nextval Into v_冲销记录id From Dual;
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 成本价, 成本金额, 扣率, 零售价, 零售金额,
     差价, 摘要, 填制人, 填制日期, 配药人, 审核人, 审核日期, 费用id, 单量, 频次, 用法, 发药窗口, 外观, 领用人, 供药单位id, 生产日期, 批准文号, 汇总发药号, 发药方式, 注册证号, 原产地)
    Select v_冲销记录id, Int记录状态 + Decode(Int记录状态, 1, 0, 1) + 1, Int单据, Strno, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地,
           批号, 效期, n_付数, -dbl实际数量, -dbl实际数量, 成本价, -dbl实际成本, 扣率, 零售价, -dbl实际金额, -dbl实际差价, 摘要, People_In, Date_In, 配药人,
           People_In, Date_In, 费用id, 单量, 频次, 用法, 发药窗口, 退药库房_In, 退药人_In, 供药单位id, 生产日期, 批准文号, 汇总发药号_In, 发药方式, 注册证号, 原产地
    From 药品收发记录
    Where ID = Billid_In;

  --如果是部分冲销，则付数填为1，实际数量为付数与实际数量的积
  --产生正常记录以供继续发药
  Select 药品收发记录_Id.Nextval Into Lng新批次 From Dual;
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 成本价, 成本金额, 扣率, 零售价, 零售金额,
     差价, 摘要, 填制人, 填制日期, 配药人, 审核人, 审核日期, 费用id, 单量, 频次, 用法, 发药窗口, 供药单位id, 生产日期, 批准文号, 注册证号, 原产地)
    Select Lng新批次, Int记录状态 + Decode(Int记录状态, 1, 0, 1) + 3, Int单据, Strno, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id,
           Decode(Lng分批, 1, 批次, 3, Lng新批次, 0), Decode(Lng分批, 3, 产地_In, 1, 产地, 产地), Decode(Lng分批, 3, 批号_In, 批号),
           Decode(Lng分批, 3, 效期_In, 效期), n_付数, Dbl实际数量, Dbl实际数量, 成本价, Dbl实际成本, 扣率, 零售价, Dbl实际金额, Dbl实际差价, 摘要, 填制人, 填制日期,
           Null, Null, Null, 费用id, 单量, 频次, 用法, 发药窗口, 供药单位id, 生产日期, 批准文号, 注册证号, 原产地
    
    From 药品收发记录
    Where ID = Billid_In;

  Zl_未审药品记录_Insert(Lng新批次);

  --更新费用记录的执行状态(0-未执行;1-完全执行;2-部分执行)
  Select Decode(Sum(Nvl(付数, 1) * 实际数量), Null, 0, 0, 0, 2)
  Into Int执行状态
  From 药品收发记录
  Where 单据 = Int单据 And NO = Strno And 费用id = Lng费用id And 审核人 Is Not Null;

  If 门诊_In = 1 Then
    Select 记录性质, 收费类别 Into n_记录性质, v_收费类别 From 门诊费用记录 Where ID = Lng费用id;
  Else
    Select 记录性质, 收费类别 Into n_记录性质, v_收费类别 From 住院费用记录 Where ID = Lng费用id;
  End If;

  If Int执行状态 = 0 Then
    If 门诊_In = 1 Then
      Update 门诊费用记录
      Set 执行状态 = Int执行状态, 执行人 = Null, 执行时间 = Null
      Where NO = Strno And
            序号 = (Select 序号 From 门诊费用记录 Where ID = (Select 费用id From 药品收发记录 Where ID = Billid_In)) And
            Mod(记录性质, 10) = n_记录性质 And 记录状态 <> 2 And 执行部门id = Lng库房id;
    Else
      Update 住院费用记录 Set 执行状态 = Int执行状态, 执行人 = Null, 执行时间 = Null Where ID = Lng费用id;
    End If;
  Else
    If 门诊_In = 1 Then
      Update 门诊费用记录
      Set 执行状态 = Int执行状态
      Where NO = Strno And
            序号 = (Select 序号 From 门诊费用记录 Where ID = (Select 费用id From 药品收发记录 Where ID = Billid_In)) And
            Mod(记录性质, 10) = n_记录性质 And 记录状态 <> 2 And 执行部门id = Lng库房id;
    Else
      Update 住院费用记录 Set 执行状态 = Int执行状态 Where ID = Lng费用id;
    End If;
  End If;

  --插入未发药品记录
  Begin
    If 门诊_In = 1 Then
      Insert Into 未发药品记录
        (单据, NO, 病人id, 主页id, 姓名, 优先级, 对方部门id, 库房id, 发药窗口, 填制日期, 已收费, 配药人, 打印状态, 未发数, 领药号, 排队状态)
        Select a.单据, a.No, a.病人id, Null, a.姓名, Nvl(b.优先级, 0) 优先级, a.对方部门id, a.库房id, a.发药窗口, a.填制日期, a.已收费, Null, 1, 1,
               a.产品合格证, v_排队状态
        From (Select b.单据, b.No, a.病人id, a.姓名, Decode(a.记录状态, 0, 0, 1) 已收费, b.对方部门id, b.库房id, b.发药窗口, b.填制日期, c.身份,
                      b.产品合格证
               From 门诊费用记录 A, 药品收发记录 B, 病人信息 C
               Where b.Id = Billid_In And a.Id = b.费用id + 0 And a.病人id = c.病人id(+)) A, 身份 B
        Where b.名称(+) = a.身份;
    Else
      Insert Into 未发药品记录
        (单据, NO, 病人id, 主页id, 姓名, 优先级, 对方部门id, 库房id, 发药窗口, 填制日期, 已收费, 配药人, 打印状态, 未发数, 领药号, 排队状态)
        Select a.单据, a.No, a.病人id, a.主页id, a.姓名, Nvl(b.优先级, 0) 优先级, a.对方部门id, a.库房id, a.发药窗口, a.填制日期, a.已收费, Null, 1, 1,
               a.产品合格证, v_排队状态
        From (Select b.单据, b.No, a.病人id, a.主页id, a.姓名, Decode(a.记录状态, 0, 0, 1) 已收费, b.对方部门id, b.库房id, b.发药窗口, b.填制日期,
                      c.身份, b.产品合格证
               From 住院费用记录 A, 药品收发记录 B, 病人信息 C
               Where b.Id = Billid_In And a.Id = b.费用id + 0 And a.病人id = c.病人id(+)) A, 身份 B
        Where b.名称(+) = a.身份;
    End If;
  
    --修改处方类型
    Zl_Prescription_Type_Update(Strno, n_记录性质, Lng药品id, v_收费类别);
  Exception
    When Others Then
      Null;
  End;

  --修改原记录为被冲销记录
  Update 药品收发记录 Set 记录状态 = Int记录状态 + Decode(Int记录状态, 1, 0, 1) + 2 Where ID = Billid_In;

  --修改药品库存(反冲库存)
  If Lng分批 <> 3 Then
    --正常单据需要将库存表实际数量和金额、差价还回去，如果库存表没有则在库存表插入数据
    Zl_药品库存_Update(v_冲销记录id, 3, 0);
  Else
    --原不分批，现在分批，直接在库存表产生新单据
    Insert Into 药品库存
      (库房id, 药品id, 批次, 效期, 性质, 实际数量, 实际金额, 实际差价, 零售价, 上次批号, 上次产地, 上次供应商id, 上次采购价, 上次生产日期, 批准文号, 平均成本价)
    Values
      (Lng库房id, Lng药品id, Lng新批次, 效期_In, 1, Dbl实际数量 * n_付数, Dbl实际金额, Dbl实际差价, Decode(n_时价分批, 1, n_零售价, Null), 批号_In,
       产地_In, n_上次供应商id, n_上次采购价, d_上次生产日期, v_批准文号, n_上次采购价);
  End If;

  Delete 药品库存
  Where 库房id + 0 = Lng库房id And 药品id = Lng药品id And 性质 = 1 And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
        Nvl(实际差价, 0) = 0;

  --处理调价修正
  Zl_药品收发记录_调价修正(v_冲销记录id);

  Begin
    --移动支付宝项目在发药后动态调用生成推送信息的过程
    Execute Immediate 'Begin zl_服务窗消息_发送(:1,:2); End;'
      Using 7, Billid_In || ',' || 退药数量_In || ',' || 门诊_In;
  Exception
    When Others Then
      Null;
  End;

  --消息处理，剩余全部退数量传0
  If Bln部分退药 = 1 Then
    b_Message.Zlhis_Drug_006(v_冲销记录id, Lng新批次, Dbl实际数量 * n_付数, Lng费用id);
  Else
    b_Message.Zlhis_Drug_006(v_冲销记录id, Lng新批次, 0, Lng费用id);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_部门退药;
/



------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0050' Where 编号=&n_System;
Commit;
