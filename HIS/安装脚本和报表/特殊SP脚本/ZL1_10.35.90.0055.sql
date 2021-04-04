----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;

------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--138776:黄捷,2019-03-26,按照影像类别设置对应的图像排序方式
alter table 影像屏幕布局 add 图像排序 NUMBER(1);

--138725:秦龙,2019-03-25,带量采购问题
Create Table 中选药品对照(
    序号 NUMBER(4),
    中选药品id NUMBER(18),
    对照药品id NUMBER(18))
    TABLESPACE zl9BaseItem;

Alter Table 药品规格 Add 是否带量采购 number(1);

Alter Table 中选药品对照 Add Constraint 中选药品对照_UQ_中选药品id Unique (中选药品id,对照药品id) Using Index Tablespace zl9Indexhis;

Alter Table 中选药品对照 Modify 中选药品id Constraint 中选药品对照_NN_中选药品id Not Null;

Alter Table 中选药品对照 Add Constraint 中选药品对照_FK_中选药品id Foreign Key (中选药品id) References 收费项目目录(ID) On Delete Cascade;
Alter Table 中选药品对照 Add Constraint 中选药品对照_FK_对照药品id Foreign Key (对照药品id) References 收费项目目录(ID) On Delete Cascade;


------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--138725:秦龙,2019-03-25,带量采购问题
Insert into zlTables(系统,表名,表空间,分类) Values(100,'中选药品对照','ZL9BASEITEM','A2');

--129946:胡俊勇,2019-03-22,传染病报告卡一病一卡
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1277, 0, 0, 0, 0, 0, 0, 3, '传染病报告卡一病一卡', '0', '0',
         '中华人民共和国传染病报告卡，在选择病种时是否允许只选一个病种。', '0-可以选多个病种，1-只能选一种病。', Null, '适用于一张报告卡填写一种传染病的情况', Null
  From Dual;

-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--138725:秦龙,2019-03-25,带量采购问题
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1023,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select '中选药品对照','SELECT' From Dual
Union All Select 'Zl_中选药品对照_Update','EXECUTE' From Dual) A;




-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--138776:黄捷,2019-03-26,按照影像类别设置对应的图像排序方式
Create Or Replace Procedure Zl_影像屏幕布局_Update
(
  人员id_In       In 影像屏幕布局.人员id%Type,
  影像类型_In     In 影像屏幕布局.影像类型%Type,
  自动序列布局_In In 影像屏幕布局.自动序列布局%Type,
  自动图像布局_In In 影像屏幕布局.自动图像布局%Type,
  序列行数_In     In 影像屏幕布局.序列行数%Type,
  序列列数_In     In 影像屏幕布局.序列列数%Type,
  图像行数_In     In 影像屏幕布局.图像行数%Type,
  图像列数_In     In 影像屏幕布局.图像列数%Type,
  自动反白_In     In 影像屏幕布局.自动反白%Type,
  显示病人信息_In In 影像屏幕布局.显示病人信息%Type,
  选择定位线_In   In 影像屏幕布局.选择定位线%Type,
  选择序列同步_In In 影像屏幕布局.选择序列同步%Type,
  插值模式_In     In 影像屏幕布局.插值模式%Type,
  图像排序_In     In 影像屏幕布局.图像排序%Type
) Is
Begin
  Update 影像屏幕布局
  Set 自动序列布局 = 自动序列布局_In, 自动图像布局 = 自动图像布局_In, 序列行数 = 序列行数_In, 序列列数 = 序列列数_In, 图像行数 = 图像行数_In, 图像列数 = 图像列数_In,
      自动反白 = 自动反白_In, 显示病人信息 = 显示病人信息_In, 选择定位线 = 选择定位线_In, 选择序列同步 = 选择序列同步_In, 插值模式 = 插值模式_In, 图像排序 = 图像排序_In
  Where 人员id = 人员id_In And 影像类型 = 影像类型_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_影像屏幕布局_Update;
/

--138960:李南春,2019-03-26,挂号取号排除被锁号
Create Or Replace Procedure Zl_病人挂号记录_Insert
(
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  单据号_In        门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  序号_In          门诊费用记录.序号%Type,
  价格父号_In      门诊费用记录.价格父号%Type,
  从属父号_In      门诊费用记录.从属父号%Type,
  收费类别_In      门诊费用记录.收费类别%Type,
  收费细目id_In    门诊费用记录.收费细目id%Type,
  数次_In          门诊费用记录.数次%Type,
  标准单价_In      门诊费用记录.标准单价%Type,
  收入项目id_In    门诊费用记录.收入项目id%Type,
  收据费目_In      门诊费用记录.收据费目%Type,
  结算方式_In      病人预交记录.结算方式%Type, --现金的结算名称
  应收金额_In      门诊费用记录.应收金额%Type,
  实收金额_In      门诊费用记录.实收金额%Type,
  病人科室id_In    门诊费用记录.病人科室id%Type,
  开单部门id_In    门诊费用记录.开单部门id%Type,
  执行部门id_In    门诊费用记录.执行部门id%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  发生时间_In      门诊费用记录.发生时间%Type,
  登记时间_In      门诊费用记录.登记时间%Type,
  医生姓名_In      挂号安排.医生姓名%Type,
  医生id_In        挂号安排.医生id%Type,
  病历费_In        Number, --该条记录是否病历工本费
  急诊_In          Number,
  号别_In          挂号安排.号码%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  领用id_In        票据使用明细.领用id%Type,
  预交支付_In      病人预交记录.冲预交%Type, --刷卡挂号时使用的预交金额,序号为1传入.
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额,序号为1传入.
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额,,序号为1传入.
  保险大类id_In    门诊费用记录.保险大类id%Type,
  保险项目否_In    门诊费用记录.保险项目否%Type,
  统筹金额_In      门诊费用记录.统筹金额%Type,
  摘要_In          门诊费用记录.摘要%Type, --预约挂号摘要信息
  预约挂号_In      Number := 0, --预约挂号时用(记录状态=0,发生时间为预约时间),此时不需要传入结算相关参数
  收费票据_In      Number := 0, --挂号是否使用收费票据
  保险编码_In      门诊费用记录.保险编码%Type,
  复诊_In          病人挂号记录.复诊%Type := 0,
  号序_In          挂号序号状态.序号%Type := Null, --预约时填入费用记录的发药窗口字段,挂号时填入挂号记录
  社区_In          病人挂号记录.社区%Type := Null,
  预约接收_In      Number := 0,
  预约方式_In      预约方式.名称%Type := Null,
  生成队列_In      Number := 0,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  合作单位_In      病人预交记录.合作单位%Type := Null,
  操作类型_In      Number := 0,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  退号重用_In      Number := 1,
  冲预交病人ids_In Varchar2 := Null,
  修正病人费别_In  Number := 0,
  修正病人年龄_In  Number := 0,
  收费单_In        病人挂号记录.收费单%Type := Null,
  更新交款余额_In  Number := 1
) As
  ---------------------------------------------------------------------------
  --
  --参数:
  --     操作类型_in:0-正常挂号或者预约 1-操作员拥有加号权限加号
  --     修正病人费别_In:0-不修改病人费别 1-修改病人费别
  --     更新交款余额_In:0-在zl_人员缴款余额_Update 中更新 1-在本过程中更新
  ----------------------------------------------------------------------------
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select 病人id, No, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, Id, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id, Min(Decode(记录性质, 1, 收款时间, Null)) As 收款时间
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select /*+cardinality(d,10)*/
                                        d.Column_Value
                                       From Table(f_Num2list(v_冲预交病人ids)) d) And Nvl(预交类别, 2) = 1 Having
     Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By No, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 收款时间;
  --功能：新增一行病人挂号费用，并将结算汇总到病人预交记录
  --       同时汇总相关的汇总表(病人挂号汇总、费用汇总)
  --       第一行费用处理票据使用情况(领用ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_排队号码 排队叫号队列.排队号码%Type;
  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  v_队列名称 排队叫号队列.队列名称%Type;

  n_分时段       Number;
  n_时段限号     Number;
  n_时段限约     Number;
  d_时段时间     Date;
  d_最大序号时间 Date;
  n_追加号       Number := 0; --处理时段过期 追加挂号的情况
  n_已约数       病人挂号汇总.已约数%Type;
  n_预约有效时间 Number;
  n_失效数       Number;
  n_失约挂号     Number := 0;
  n_已用数量     Number;
  n_锁定         Number := 0;

  n_费用id        门诊费用记录.Id%Type;
  n_病人余额      病人预交记录.金额%Type;
  n_预交金额      病人预交记录.金额%Type;
  n_当前金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  n_预交id        病人预交记录.Id%Type;
  n_挂号id        病人挂号记录.Id%Type;
  v_冲预交病人ids Varchar2(4000);

  n_组id           财务缴款分组.Id%Type;
  n_门诊号         病人信息.门诊号%Type;
  n_序号           挂号序号状态.序号%Type;
  n_已用序号       挂号序号状态.序号%Type;
  n_序号控制       挂号安排.序号控制%Type;
  n_分诊台签到排队 Number;
  n_Count          Number;
  n_限号数         Number(18);
  d_排队时间       Date;
  d_序号时间       Date;
  v_费别           费别.名称%Type;
  n_时段序号       Number := -1;
  n_预约生成队列   Number;
  n_安排id         挂号安排.Id%Type;
  n_计划id         挂号安排计划.Id%Type := 0;
  v_星期           挂号安排限制.限制项目%Type;
  n_限约数         Number(18);
  n_已挂数         Number(4) := 0;

  n_挂出的最大序号 Number(4) := 0;
  n_结算模式       病人信息.结算模式%Type;
  v_排队序号       排队叫号队列.排队序号%Type;
  v_机器名         挂号序号状态.机器名%Type;
  v_序号操作员     挂号序号状态.操作员姓名%Type;
  v_序号机器名     挂号序号状态.机器名%Type;
  v_付款方式       病人挂号记录.医疗付款方式%Type;
  v_Temp           Varchar2(3000);
  v_时间段         时间段.时间段%Type;
  d_检查开始时间   时间段.开始时间%Type;
  d_检查结束时间   时间段.终止时间%Type;
  n_出诊记录id     临床出诊记录.Id%Type;
  n_分时点显示     Number(3);
  d_启用时间       Date;
Begin
  --获取当前机器名称
  Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);

  If 费别_In Is Null Then
    Begin
      Select 名称 Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
      Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
    Exception
      When Others Then
        v_Err_Msg := '无法确定病人费别，请检查缺省费别是否正确设置！';
        Raise Err_Item;
    End;
  Else
    v_费别 := 费别_In;
    If Nvl(修正病人费别_In, 0) = 1 Then
      Begin
        Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '没有找到对应的病人！';
          Raise Err_Item;
      End;
    End If;
  End If;

  If Nvl(修正病人年龄_In, 0) = 1 Then
    Begin
      Update 病人信息 Set 年龄 = 年龄_In Where 病人id = 病人id_In;
    Exception
      When Others Then
        v_Err_Msg := '没有找到对应的病人！';
        Raise Err_Item;
    End;
  End If;

  If 门诊号_In Is Not Null Then
    Begin
      Select Nvl(门诊号, 0) Into n_门诊号 From 病人信息 Where 病人id = 病人id_In;
    Exception
      When Others Then
        n_门诊号 := 0;
    End;
    If n_门诊号 = 0 Then
      Update 病人信息 Set 门诊号 = 门诊号_In Where 病人id = 病人id_In;
    End If;
  End If;

  Begin
    Delete From 挂号序号状态
    Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + (1 - 1/24/60/60) And 序号 = 号序_In 
    And ((状态 = 3 And 操作员姓名 = 操作员姓名_In) Or (状态 = 5 And 操作员姓名_In = 操作员姓名 And v_机器名 = 机器名));
  Exception
    When Others Then
      Null;
  End;
  
  v_Temp := zl_GetSysParameter(256);
  If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
    Null;
  Else
    Begin
      d_启用时间 := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
    Exception
      When Others Then
        d_启用时间 := Null;
    End;
    If d_启用时间 Is Not Null Then
      If 发生时间_In > d_启用时间 Then
        v_Err_Msg := '当前挂号的发生时间' || To_Char(发生时间_In, 'yyyy-mm-dd hh24:mi:ss') || '已经启用了出诊表排班模式,不能再使用计划排班模式挂号!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 0 Then
    --挂号或者预约接收
    --因为现在有按日编单据号规则,日挂号量不能超过10000次,所以要检查唯一约束。
    Select Count(*)
    Into n_Count
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 In (1, 3) And 序号 = 序号_In And NO = 单据号_In;
    If n_Count <> 0 Then
      v_Err_Msg := '挂号单据号重复,不能保存！' || Chr(13) || '如果使用了按日顺序编号,当日挂号量不能超过10000人次。';
      Raise Err_Item;
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
  End If;

  n_序号 := 号序_In;
  Select Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)
  Into v_星期
  From Dual;

  --挂号获取安排
  Begin
    Select a.Id, a.序号控制, Nvl(b.限号数, 0), Nvl(b.限约数, 0)
    Into n_安排id, n_序号控制, n_限号数, n_限约数
    From 挂号安排 A, 挂号安排限制 B
    Where a.Id = b.安排id(+) And b.限制项目(+) = v_星期 And a.号码 = 号别_In;
  
  Exception
    When Others Then
      n_安排id := -1;
  End;

  --如果是病历费或者号别为空时不检查
  If Nvl(病历费_In, 0) = 0 Or 号别_In Is Not Null Then
    If n_安排id = -1 Then
      v_Err_Msg := '不存相应的挂号安排数据,请检查';
      Raise Err_Item;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 1 Then
    --首先获取计划
    Begin
      Select ID
      Into n_计划id
      From 挂号安排计划
      Where 安排id = n_安排id And 审核时间 Is Not Null And
            Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) =
            (Select Max(a.生效时间) As 生效
             From 挂号安排计划 A
             Where a.审核时间 Is Not Null And 发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                   a.失效时间 And a.安排id = n_安排id) And
            发生时间_In Between Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And 失效时间;
    
    Exception
      When Others Then
        n_计划id := 0;
    End;
    If Nvl(n_计划id, 0) <> 0 Then
      Begin
        --获取计划的限制
        Select a.Id, a.序号控制, Nvl(b.限号数, 0) As 限号数, Nvl(b.限约数, 0) As 限约数
        Into n_计划id, n_序号控制, n_限号数, n_限约数
        From 挂号安排计划 A, 挂号计划限制 B
        Where a.号码 = 号别_In And a.Id = n_计划id And a.审核时间 Is Not Null And a.Id = b.计划id(+) And b.限制项目(+) = v_星期;
      Exception
        When Others Then
          v_Err_Msg := '不存相应的挂号安排或计划数据,请检查';
          Raise Err_Item;
      End;
    End If;
  End If;

  --获取是否分时段
  Begin
    If Nvl(n_计划id, 0) = 0 Then
      Select Count(Rownum) Into n_分时段 From 挂号安排时段 Where 星期 = v_星期 And 安排id = n_安排id And Rownum <= 1;
      Select Decode(To_Char(发生时间_In, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排
      Where ID = n_安排id;
    Else
      Select Count(Rownum) Into n_分时段 From 挂号计划时段 Where 星期 = v_星期 And 计划id = n_计划id And Rownum <= 1;
      Select Decode(To_Char(发生时间_In, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排计划
      Where ID = n_计划id;
    End If;
  Exception
    When Others Then
      v_时间段 := Null;
  End;

  If v_时间段 Is Not Null And d_启用时间 Is Not Null And 序号_In = 1 Then
    --检查是否跨模式挂号安排
    Select To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
           To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
    Into d_检查开始时间, d_检查结束时间
    From 时间段
    Where 时间段 = v_时间段 And 站点 Is Null And 号类 Is Null;
    If d_检查开始时间 > d_检查结束时间 Then
      d_检查结束时间 := d_检查结束时间 + 1;
    End If;
    If d_检查开始时间 < d_启用时间 And d_检查结束时间 > d_启用时间 Then
      --获取出诊记录id
      Begin
        Select a.Id
        Into n_出诊记录id
        From 临床出诊记录 A, 临床出诊号源 B
        Where a.号源id = b.Id And b.号码 = 号别_In And 上班时段 = v_时间段 And 发生时间_In Between 开始时间 And 终止时间;
      Exception
        When Others Then
          n_出诊记录id := Null;
      End;
    End If;
  End If;

  --分时段挂号时判断是否是过期挂号 也就是追加号的情况 目前只针对专家号分时段进行处理
  If Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 And Nvl(n_序号控制, 0) = 1 And 号序_In Is Null And Nvl(操作类型_In, 0) = 0 Then
    --发生时间_in>Sysdate 发生时间>最大的时段时间--号序_in is null
    Begin
      Select Max(To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_最大序号时间
      From 挂号安排时段
      Where 安排id = n_安排id And 星期 = v_星期 And Nvl(限制数量, 0) <> 0;
      n_追加号 := Case Sign(发生时间_In - d_最大序号时间)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_追加号 := 0;
    End;
  End If;
  d_时段时间 := 发生时间_In;

  If 序号_In = 1 And Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 Then
    --挂号时检查 是否分了时段,分了时段,把时段的限制条件给取出来
    Begin
      Select Nvl(序号, 0),
             To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
      Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
      From 挂号安排时段
      Where 安排id = n_安排id And 星期 = v_星期 And
            (序号, 安排id, 星期) In (Select Nvl(Max(序号), -1), 安排id, 星期
                               From 挂号安排时段
                               Where 安排id = n_安排id And 星期 = v_星期 And
                                     Decode(操作类型_In + n_追加号, 0, To_Char(发生时间_In, 'hh24:mi'),
                                            To_Char(Nvl(开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')) =
                                     To_Char(Nvl(开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
                               Group By 安排id, 星期);
    Exception
      When Others Then
        n_时段序号 := -1;
        n_分时段   := 0;
        d_时段时间 := 发生时间_In;
        n_时段限号 := 0;
        n_时段限约 := 0;
    End;
  End If;

  If 序号_In = 1 And Nvl(预约挂号_In, 0) = 1 And n_分时段 > 0 Then
    --预约号,取计划
    Begin
      If Nvl(n_计划id, 0) = 0 Then
        --没计划生效,取安排的数据
        Select Nvl(序号, 0),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(c.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 时段时间,
               限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 挂号安排时段 C
        Where 安排id = n_安排id And 星期 = v_星期 And
              (序号, 安排id, 星期) In
              (Select Nvl(Max(c.序号), -1), 安排id, 星期
               From 挂号安排时段 C
               Where 安排id = n_安排id And c.星期 = v_星期 And
                     Decode(操作类型_In, 1, To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(发生时间_In, 'hh24:mi')) =
                     To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By 安排id, 星期);
      Else
        --有计划生效取计划
        --没生效，代表是从挂号计划时段查询
        Select Nvl(序号, -1),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(c.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 时段时间,
               限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 挂号计划时段 C
        Where 计划id = n_计划id And 星期 = v_星期 And
              (序号, 计划id, 星期) In
              (Select Nvl(Max(c.序号), -1), 计划id, 星期
               From 挂号计划时段 C
               Where 计划id = n_计划id And c.星期 = v_星期 And
                     Decode(操作类型_In, 1, To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(发生时间_In, 'hh24:mi')) =
                     To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By 计划id, 星期);
      End If;
    Exception
      When Others Then
        n_时段序号 := -1;
        n_分时段   := 0;
        d_时段时间 := 发生时间_In;
        n_时段限号 := 0;
        n_时段限约 := 0;
    End;
  End If;

  If 序号_In = 1 Then
  
    --获取当前未使用的序号
    If Nvl(n_序号控制, 0) = 1 And n_分时段 = 0 Then
      --<序号控制 未设置时段 获取可用的最大序号,以及已经使用的数量>
      Begin
        --最大序号
        If 退号重用_In = 1 Then
          Select Max(Nvl(序号, 0)), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Nvl(预约, 0))
          Into n_已用序号, n_已用数量, n_已约数
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 <> 4;
        Else
          Select Max(Nvl(序号, 0)), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Nvl(预约, 0))
          Into n_已用序号, n_已用数量, n_已约数
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60;
        End If;
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
      End;
      If n_序号 Is Null Then
        n_序号 := Nvl(n_已用序号, 0) + 1;
      End If;
      --<序号控制 未设置时段 获取可用的最大序号 以及已经使用的数量 --end>
    
      --非加号的情况需要检查是否超过了限制
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号时检查
          --启用序号控制未分时段 达到了限制
          If n_限号数 <= n_已用数量 And n_限号数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd ') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约时检查
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
        
          If n_限约数 <= n_已约数 And n_限约数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd') || '已达到最大限约数！';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --启用序号控制,未分时段 加号情况   不处理,如果以后有限制条件以后补充
      End If;
    
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 0 Then
      --<--普通号分时段 这里只有预约一种情况-->
      If 操作类型_In = 0 Then
        --<正常预约挂号-->
        Begin
          Select Count(0) As 已挂数, Nvl(Sum(Decode(Nvl(Sign(a.日期 - d_时段时间), 0), 0, 1, 0)), 0) As 已约数
          Into n_已挂数, n_已约数
          From 挂号序号状态 A
          Where a.号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And
                状态 <> 4;
        Exception
          When Others Then
            n_已挂数 := 0;
            n_已约数 := 0;
        End;
      
        n_时段限约 := n_时段限号; --普通号分时段的情况,29中n_时段限约始终是0 这里特殊处理
        --检查限制数量
        If n_限约数 = 0 Then
          n_限约数 := n_限号数;
        End If;
        If n_时段限约 <= n_已约数 Or n_限约数 <= n_已约数 Then
          v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限约数！';
          Raise Err_Item;
        End If;
        If n_限号数 <= n_已挂数 Then
          v_Err_Msg := '号别' || 号别_In || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
          Raise Err_Item;
        End If;
      End If;
    
      --没有达到时段的限号数 号码在当前时段往后追加
    
      --获取当天挂出的最大号序
      If 退号重用_In = 1 Then
        Select Nvl(Max(序号), 0)
        Into n_挂出的最大序号
        From 挂号序号状态 A
        Where a.日期 = d_时段时间 And 号码 = 号别_In And 状态 <> 4;
      Else
        Select Nvl(Max(序号), 0)
        Into n_挂出的最大序号
        From 挂号序号状态 A
        Where a.日期 = d_时段时间 And 号码 = 号别_In;
      End If;
    
      --设置序号
      n_序号 := RPad(Nvl(n_时段序号, 0), Length(n_限号数) + Length(Nvl(n_时段序号, 0)), 0) + n_已约数 + 1;
      If n_序号 <= Nvl(n_挂出的最大序号, 0) Then
        n_序号 := Nvl(n_挂出的最大序号, 0) + 1;
      End If;
    
      --<--普通号分时段--End>
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 1 Then
      --<启用序号控制 设置时段
      --专家号分时段
      Begin
        If 退号重用_In = 1 Then
          Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(日期 - d_时段时间), 0, 1, 0))
          Into n_已用序号, n_已挂数, n_已用数量
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 <> 4;
        Else
          Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(日期 - d_时段时间), 0, 1, 0))
          Into n_已用序号, n_已挂数, n_已用数量
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60;
        End If;
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
          n_已挂数   := 0;
      End;
    
      n_失效数 := 0;
      If Nvl(预约挂号_In, 0) = 0 Then
        n_预约有效时间 := Zl_To_Number(zl_GetSysParameter('预约有效时间', 1111));
        n_失约挂号     := Zl_To_Number(zl_GetSysParameter('失约用于挂号', 1111));
        If Nvl(n_预约有效时间, 0) <> 0 And Nvl(n_失约挂号, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_预约有效时间 / 24 / 60) - 日期), 1, 1, 0))
            Into n_失效数
            From 挂号序号状态
            Where 号码 = 号别_In And 日期 Between Trunc(Sysdate) And Sysdate And Nvl(预约, 0) = 1 And 状态 = 2;
          Exception
            When Others Then
              n_失效数 := 0;
          End;
        End If;
      End If;
    
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
        
          --挂号 追加号码时不检查时段限号数
          If n_时段限号 <= n_已用数量 And Nvl(n_追加号, 0) = 0 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 - n_失效数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --挂号
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
          If n_限约数 <= n_已用数量 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_序号 Is Null Then
        --设置序号
        If Nvl(n_已用序号, 0) < Nvl(n_时段序号, 0) Then
          n_已用序号 := Nvl(n_时段序号, 0);
        End If;
        n_序号 := Nvl(n_已用序号, 0) + 1;
      End If;
    Elsif Nvl(n_分时段, 0) = 0 And Nvl(n_序号控制, 0) = 0 And Nvl(病历费_In, 0) = 0 And Nvl(号别_In, 0) > 0 Then
      ---<--普通号  -->
      Begin
        Select 已挂数, 已约数
        Into n_已用数量, n_已约数
        From 病人挂号汇总
        Where 日期 = Trunc(发生时间_In) And 号码 = 号别_In;
      Exception
        When Others Then
          n_已用数量 := 0;
          n_已约数   := 0;
      End;
      If Nvl(操作类型_In, 0) = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号
          If Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约
          If (Nvl(n_限约数, 0) > 0 And Nvl(n_已约数, 0) >= Nvl(n_限约数, 0)) Or
             (Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0) Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--普通号  -->
    End If;
  End If;

  --更新挂号序号状态
  If 序号_In = 1 And Not n_序号 Is Null Then
    If n_分时段 = 1 Then
      d_序号时间 := 发生时间_In;
    Else
      d_序号时间 := Trunc(发生时间_In);
    End If;
    --锁定序号的处理
    Begin
      Select 操作员姓名, 机器名
      Into v_序号操作员, v_序号机器名
      From 挂号序号状态
      Where 状态 = 5 And 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
      n_锁定 := 1;
    Exception
      When Others Then
        v_序号操作员 := Null;
        v_序号机器名 := Null;
        n_锁定       := 0;
    End;
    If n_锁定 = 0 Then
      Update 挂号序号状态
      Set 状态 = Decode(预约挂号_In, 1, 2, 1), 预约 = Decode(预约接收_In, 1, 1, 0), 登记时间 = Sysdate
      Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号 And 状态 = 3 And 操作员姓名 = 操作员姓名_In;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_分时段, 0) = 0 Or Nvl(预约挂号_In, 0) = 1 Or (Nvl(n_序号控制, 0) = 0 And Nvl(号序_In, 0) = 0) Then
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 预约, 登记时间)
            Values
              (号别_In, d_序号时间, n_序号, Decode(预约挂号_In, 1, 2, 1), 操作员姓名_In, Decode(预约接收_In, 1, 1, 0), Sysdate);
            If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
              Update 挂号序号状态 Set 预约 = 1 Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
            End If;
          Elsif Nvl(n_分时段, 0) > 0 Then
            --分时段后专家号 失约的预约号允许挂号
            Update 挂号序号状态
            Set 状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In, 预约 = Decode(预约接收_In, 1, 1, 0), 登记时间 = Sysdate
            Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号 And 状态 = 2;
            If Sql%NotFound Then
              Insert Into 挂号序号状态
                (号码, 日期, 序号, 状态, 操作员姓名, 预约, 登记时间)
              Values
                (号别_In, d_序号时间, n_序号, Decode(预约挂号_In, 1, 2, 1), 操作员姓名_In, Decode(预约接收_In, 1, 1, 0), Sysdate);
            End If;
            If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
              Update 挂号序号状态 Set 预约 = 1 Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
            Raise Err_Item;
        End;
      End If;
    Else
      If 操作员姓名_In <> v_序号操作员 Or v_机器名 <> v_序号机器名 Then
        v_Err_Msg := '序号' || n_序号 || '已被自助机' || v_机器名 || '锁定,请重新选择一个序号.';
        Raise Err_Item;
      Else
        Update 挂号序号状态
        Set 状态 = Decode(预约挂号_In, 1, 2, 1), 预约 = Decode(预约接收_In, 1, 1, 0), 登记时间 = Sysdate
        Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号 And 状态 = 5 And 操作员姓名 = 操作员姓名_In And 机器名 = v_机器名;
        If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
          Update 挂号序号状态 Set 预约 = 1 Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
        End If;
      End If;
    End If;
  End If;

  If n_出诊记录id Is Not Null Then
    Update 临床出诊序号控制
    Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In
    Where 记录id = n_出诊记录id And 序号 = n_序号;
    If 预约挂号_In = 1 Then
      Update 临床出诊记录 Set 已约数 = 已约数 + 1 Where ID = n_出诊记录id;
    Else
      If 预约接收_In = 1 Then
        Update 临床出诊记录
        Set 已约数 = 已约数 + 1, 已挂数 = 已挂数 + 1, 其中已接收 = 其中已接收 + 1
        Where ID = n_出诊记录id;
      Else
        Update 临床出诊记录 Set 已挂数 = 已挂数 + 1 Where ID = n_出诊记录id;
      End If;
    End If;
  End If;

  --产生病人挂号费用(可能单独是或包括病历费用)
  Select 病人费用记录_Id.Nextval Into n_费用id From Dual; --应该通过程序得到

  Insert Into 门诊费用记录
    (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id, 收费类别,
     计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人, 操作员编号, 操作员姓名,
     发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
  Values
    (n_费用id, 4, Decode(预约挂号_In, 1, 0, 1), 序号_In, Decode(价格父号_In, 0, Null, 价格父号_In), 从属父号_In, 单据号_In, 票据号_In, 1, 急诊_In,
     病历费_In, Decode(预约挂号_In, 1, To_Char(n_序号), 诊室_In), Decode(病人id_In, 0, Null, 病人id_In),
     Decode(门诊号_In, 0, Null, 门诊号_In), 付款方式_In, 姓名_In, Decode(姓名_In, Null, Null, 性别_In), Decode(姓名_In, Null, Null, 年龄_In),
     v_费别, 病人科室id_In, 收费类别_In, 号别_In, 收费细目id_In, 收入项目id_In, 收据费目_In, 1, 数次_In, 标准单价_In, 应收金额_In, 实收金额_In,
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额_In)),
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In)), Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 开单部门id_In,
     操作员姓名_In, Decode(预约挂号_In, 1, 操作员姓名_In, Null), 执行部门id_In, 医生姓名_In, 操作员编号_In, 操作员姓名_In, 发生时间_In, 登记时间_In, 保险大类id_In,
     保险项目否_In, 保险编码_In, 统筹金额_In, Decode(收费单_In, Null, 摘要_In, '划价:' || 收费单_In), 预约方式_In, Decode(预约挂号_In, 1, Null, n_组id));

  --汇总结算到病人预交记录
  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 0 Then
  
    If (Nvl(现金支付_In, 0) <> 0 Or (Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0)) And 序号_In = 1 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
         结算性质)
      Values
        (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(结算方式_In, v_现金), Nvl(现金支付_In, 0), 登记时间_In,
         操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4);
    
      If Nvl(结算卡序号_In, 0) <> 0 And Nvl(现金支付_In, 0) <> 0 Then
        Zl_病人卡结算记录_支付(结算卡序号_In, 卡号_In, 0, 现金支付_In, n_预交id, 操作员编号_In, 操作员姓名_In, 登记时间_In);
      End If;
    End If;
  
    --对于医保挂号
    If Nvl(个帐支付_In, 0) <> 0 And 序号_In = 1 Then
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), v_个人帐户, 个帐支付_In, 登记时间_In, 操作员编号_In,
         操作员姓名_In, 结帐id_In, '医保挂号', n_组id, 4);
    End If;
  
    --对于就诊卡通过预交金挂号
    If Nvl(预交支付_In, 0) <> 0 And 序号_In = 1 Then
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
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
           冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                 登记时间_In, 操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
      
        --更新病人预交余额
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2);
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
      If n_预交金额 > 0 Then
        v_Err_Msg := '预交余不够支付本次支付金额,不能继续操作！';
        Raise Err_Item;
      
      End If;
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
    End If;
  
    --相关汇总表的处理
    --人员缴款余额
    If 序号_In = 1 And Nvl(更新交款余额_In, 1) = 1 Then
      If Nvl(现金支付_In, 0) <> 0 Then
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
          Where 收款员 = 操作员姓名_In And 结算方式 = Nvl(结算方式_In, v_现金) And 性质 = 1 And Nvl(余额, 0) = 0;
        End If;
      End If;
    
      If Nvl(个帐支付_In, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + 个帐支付_In
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
          n_返回值 := 个帐支付_In;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_个人帐户 And Nvl(余额, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --病人挂号汇总(只处理一次,且单独收取病历费不处理)
  If Nvl(预约挂号_In, 0) = 0 Then
    --病人本次就诊(以发生时间为准)
    If Nvl(病人id_In, 0) <> 0 And 序号_In = 1 Then
      Update 病人信息 Set 就诊时间 = 发生时间_In, 就诊状态 = 1, 就诊诊室 = 诊室_In Where 病人id = 病人id_In;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
  
    --病人余额
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (病人id_In, 1, 1, Nvl(实收金额_In, 0), 0);
    End If;
  
    --病人未结费用
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(病人科室id_In, 0) And
          Nvl(开单部门id, 0) = Nvl(开单部门id_In, 0) And Nvl(执行部门id, 0) = Nvl(执行部门id_In, 0) And 收入项目id + 0 = 收入项目id_In And
          来源途径 + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (病人id_In, Null, Null, 病人科室id_In, 开单部门id_In, 执行部门id_In, 收入项目id_In, 1, Nvl(实收金额_In, 0));
    End If;
  End If;

  --病人挂号记录
  If 号别_In Is Not Null And 序号_In = 1 Then
    --And Nvl(预约挂号_In, 0) = 0
    Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
    Begin
      Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
    Exception
      When Others Then
        v_付款方式 := Null;
    End;
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 预约时间, 操作员编号,
       操作员姓名, 复诊, 号序, 社区, 预约, 预约方式, 摘要, 交易流水号, 交易说明, 合作单位, 接收时间, 接收人, 预约操作员, 预约操作员编号, 险类, 医疗付款方式, 收费单)
    Values
      (n_挂号id, 单据号_In, Decode(Nvl(预约挂号_In, 0), 1, 2, 1), 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 号别_In, 急诊_In, 诊室_In,
       Null, 执行部门id_In, 医生姓名_In, 0, Null, 登记时间_In, 发生时间_In, Decode(Nvl(预约挂号_In, 0), 1, 发生时间_In, Null), 操作员编号_In,
       操作员姓名_In, 复诊_In, n_序号, 社区_In, Decode(预约接收_In, 1, 1, 0), 预约方式_In, 摘要_In, 交易流水号_In, 交易说明_In, 合作单位_In,
       Decode(Nvl(预约挂号_In, 0), 0, 登记时间_In, Null), Decode(Nvl(预约挂号_In, 0), 0, 操作员姓名_In, Null),
       Decode(Nvl(预约挂号_In, 0), 1, 操作员姓名_In, Null), Decode(Nvl(预约挂号_In, 0), 1, 操作员编号_In, Null),
       Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, 收费单_In);
    If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
      Update 病人挂号记录
      Set 预约 = 1, 预约时间 = 发生时间_In, 预约操作员 = 操作员姓名_In, 预约操作员编号 = 操作员编号_In
      Where ID = n_挂号id;
    End If;
    n_预约生成队列 := 0;
    If Nvl(预约挂号_In, 0) = 1 Then
      n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    End If;
  
    --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
    If Nvl(生成队列_In, 0) <> 0 And Nvl(预约挂号_In, 0) = 0 Or n_预约生成队列 = 1 Then
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(执行部门id_In, 0)));
      If Nvl(n_分诊台签到排队, 0) = 0 Or n_预约生成队列 = 1 Then
        n_分时点显示 := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
        If Nvl(预约挂号_In, 0) = 1 And n_分时点显示 = 1 And n_分时段 = 1 Then
          n_分时点显示 := 1;
        Else
          n_分时点显示 := Null;
        End If;
      
        --产生队列
        --.按”执行部门” 的方式生成队列
        v_队列名称 := 执行部门id_In;
        v_排队号码 := Zlgetnextqueue(执行部门id_In, n_挂号id, 号别_In || '|' || n_序号);
        v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号别_In, n_序号, 发生时间_In);
        --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
        Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, 执行部门id_In, v_排队号码, Null, 姓名_In, 病人id_In, 诊室_In, 医生姓名_In, d_排队时间, 预约方式_In,
                         n_分时点显示, v_排队序号);
      
        --挂号立即排队
        If Nvl(n_分诊台签到排队, 0) = 0 Then
          Update 病人挂号记录 Set 记录标志 = 1 Where ID = n_挂号id;
        End If;
      End If;
    End If;
  End If;
  --病人担保信息
  If 病人id_In Is Not Null And 序号_In = 1 Then
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(结算模式, 0) Into n_结算模式 From 病人信息 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的病人信息,不允许挂号';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
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
  
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Nvl(在院, 0) = 0 And Exists
     (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
  
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = Sysdate
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, Sysdate) >= Sysdate;
    End If;
  End If;
  If 序号_In = 1 Then
    --消息推送
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
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_Insert;
/

--138960:李南春,2019-03-26,挂号取号排除被锁号
Create Or Replace Procedure Zl_病人挂号记录_出诊_Insert
(
  出诊记录id_In    临床出诊记录.Id%Type,
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  单据号_In        门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  序号_In          门诊费用记录.序号%Type,
  价格父号_In      门诊费用记录.价格父号%Type,
  从属父号_In      门诊费用记录.从属父号%Type,
  收费类别_In      门诊费用记录.收费类别%Type,
  收费细目id_In    门诊费用记录.收费细目id%Type,
  数次_In          门诊费用记录.数次%Type,
  标准单价_In      门诊费用记录.标准单价%Type,
  收入项目id_In    门诊费用记录.收入项目id%Type,
  收据费目_In      门诊费用记录.收据费目%Type,
  结算方式_In      Varchar2,
  应收金额_In      门诊费用记录.应收金额%Type,
  实收金额_In      门诊费用记录.实收金额%Type,
  病人科室id_In    门诊费用记录.病人科室id%Type,
  开单部门id_In    门诊费用记录.开单部门id%Type,
  执行部门id_In    门诊费用记录.执行部门id%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  发生时间_In      门诊费用记录.发生时间%Type,
  登记时间_In      门诊费用记录.登记时间%Type,
  医生姓名_In      挂号安排.医生姓名%Type,
  医生id_In        挂号安排.医生id%Type,
  病历费_In        Number, --该条记录是否病历工本费
  急诊_In          Number,
  号别_In          挂号安排.号码%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  领用id_In        票据使用明细.领用id%Type,
  预交支付_In      病人预交记录.冲预交%Type, --刷卡挂号时使用的预交金额,序号为1传入.
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额,序号为1传入.
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额,,序号为1传入.
  保险大类id_In    门诊费用记录.保险大类id%Type,
  保险项目否_In    门诊费用记录.保险项目否%Type,
  统筹金额_In      门诊费用记录.统筹金额%Type,
  摘要_In          门诊费用记录.摘要%Type, --预约挂号摘要信息
  预约挂号_In      Number := 0, --预约挂号时用(记录状态=0,发生时间为预约时间),此时不需要传入结算相关参数
  收费票据_In      Number := 0, --挂号是否使用收费票据
  保险编码_In      门诊费用记录.保险编码%Type,
  复诊_In          病人挂号记录.复诊%Type := 0,
  号序_In          挂号序号状态.序号%Type := Null, --预约时填入费用记录的发药窗口字段,挂号时填入挂号记录
  社区_In          病人挂号记录.社区%Type := Null,
  预约接收_In      Number := 0,
  预约方式_In      预约方式.名称%Type := Null,
  生成队列_In      Number := 0,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  合作单位_In      病人预交记录.合作单位%Type := Null,
  操作类型_In      Number := 0,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  退号重用_In      Number := 1,
  冲预交病人ids_In Varchar2 := Null,
  修正病人费别_In  Number := 0,
  预约顺序号_In    临床出诊序号控制.预约顺序号%Type := Null,
  修正病人年龄_In  Number := 0,
  收费单_In        病人挂号记录.收费单%Type := Null,
  更新交款余额_In  Number := 1 --是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况
) As
  ---------------------------------------------------------------------------
  --
  --参数:
  --     操作类型_in:0-正常挂号或者预约 1-操作员拥有加号权限加号
  --     修正病人费别_In:0-不修改病人费别 1-修改病人费别
  ----------------------------------------------------------------------------
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select 病人id, No, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, Id, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id, Min(Decode(记录性质, 1, 收款时间, Null)) As 收款时间
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select /*+cardinality(d,10)*/
                                        d.Column_Value
                                       From Table(f_Num2list(v_冲预交病人ids)) d) And Nvl(预交类别, 2) = 1 Having
     Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By No, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 收款时间;
  --功能：新增一行病人挂号费用，并将结算汇总到病人预交记录
  --       同时汇总相关的汇总表(病人挂号汇总、费用汇总)
  --       第一行费用处理票据使用情况(领用ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_排队号码 排队叫号队列.排队号码%Type;
  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  v_队列名称 排队叫号队列.队列名称%Type;

  n_分时段       Number;
  n_原始分时段   Number;
  n_时段限号     Number;
  n_时段限约     Number;
  d_时段时间     Date;
  d_最大序号时间 Date;
  n_追加号       Number := 0; --处理时段过期 追加挂号的情况
  n_已约数       病人挂号汇总.已约数%Type;
  n_预约有效时间 Number;
  n_失效数       Number;
  n_失约挂号     Number := 0;
  n_已用数量     Number;
  n_锁定         Number := 0;

  n_费用id        门诊费用记录.Id%Type;
  n_病人余额      病人预交记录.金额%Type;
  n_预交金额      病人预交记录.金额%Type;
  n_当前金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  n_预交id        病人预交记录.Id%Type;
  n_挂号id        病人挂号记录.Id%Type;
  v_冲预交病人ids Varchar2(4000);

  n_组id           财务缴款分组.Id%Type;
  n_门诊号         病人信息.门诊号%Type;
  n_序号           挂号序号状态.序号%Type;
  n_已用序号       挂号序号状态.序号%Type;
  n_序号控制       挂号安排.序号控制%Type;
  n_分诊台签到排队 Number;
  n_Count          Number;
  n_限号数         Number(18);
  d_排队时间       Date;
  v_结算方式记录   Varchar2(1000);
  d_序号时间       Date;
  v_费别           费别.名称%Type;
  n_时段序号       Number := -1;
  n_预约生成队列   Number;
  v_结算方式       结算方式.名称%Type;
  v_结算内容       Varchar2(1000);
  v_当前结算       Varchar2(200);
  v_结算号码       病人预交记录.结算号码%Type;
  n_结算金额       病人预交记录.冲预交%Type;
  n_三方卡标志     Number(2);
  n_预约顺序号     临床出诊序号控制.预约顺序号%Type;
  n_限约数         Number(18);
  n_已挂数         Number(4) := 0;
  n_Exists         Number;
  n_挂出的最大序号 Number(4) := 0;
  n_分时点显示     Number(3);
  n_结算模式       病人信息.结算模式%Type;
  v_排队序号       排队叫号队列.排队序号%Type;
  v_机器名         挂号序号状态.机器名%Type;
  v_序号操作员     挂号序号状态.操作员姓名%Type;
  v_序号机器名     挂号序号状态.机器名%Type;
  v_付款方式       病人挂号记录.医疗付款方式%Type;
  n_状态           临床出诊序号控制.挂号状态%Type;
Begin
  --记录锁定判断
  If Nvl(序号_In, 0) = 1 Then
    If 出诊记录id_In Is Not Null Then
      Begin
        Select 1
        Into n_Exists
        From 临床出诊记录 a, 临床出诊号源 b
        Where a.Id = 出诊记录id_In And a.号源id = b.Id And b.号码 = 号别_In And a.科室id = 执行部门id_In And Nvl(a.是否发布, 0) = 1 And
              Nvl(a.是否锁定, 0) = 0;
      Exception
        When Others Then
          v_Err_Msg := '无法确定出诊记录，请检查出诊记录是否存在或被锁定！';
          Raise Err_Item;
      End;
    End If;
  End if;

  --获取当前机器名称
  Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);

  If 费别_In Is Null Then
    Begin
      Select 名称 Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
      Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
    Exception
      When Others Then
        v_Err_Msg := '无法确定病人费别，请检查缺省费别是否正确设置！';
        Raise Err_Item;
    End;
  Else
    v_费别 := 费别_In;
    If Nvl(修正病人费别_In, 0) = 1 Then
      Begin
        Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '没有找到对应的病人！';
          Raise Err_Item;
      End;
    End If;
  End If;

  If Nvl(修正病人年龄_In, 0) = 1 Then
    Begin
      Update 病人信息 Set 年龄 = 年龄_In Where 病人id = 病人id_In;
    Exception
      When Others Then
        v_Err_Msg := '没有找到对应的病人！';
        Raise Err_Item;
    End;
  End If;

  If 门诊号_In Is Not Null Then
    Begin
      Select Nvl(门诊号, 0) Into n_门诊号 From 病人信息 Where 病人id = 病人id_In;
    Exception
      When Others Then
        n_门诊号 := 0;
    End;
    If n_门诊号 = 0 Then
      Update 病人信息 Set 门诊号 = 门诊号_In Where 病人id = 病人id_In;
    End If;
  End If;

  Begin
    Update 临床出诊序号控制
    Set 挂号状态 = 0
    Where 记录id = 出诊记录id_In And 序号 = 号序_In And 
          ((Nvl(挂号状态, 0) = 3 And 操作员姓名 = 操作员姓名_In) Or (Nvl(挂号状态, 0) = 5 And 操作员姓名_In = 操作员姓名 And v_机器名 = 工作站名称));
  Exception
    When Others Then
      Null;
  End;

  If Nvl(预约挂号_In, 0) = 0 Then
    --挂号或者预约接收
    --因为现在有按日编单据号规则,日挂号量不能超过10000次,所以要检查唯一约束。
    Select Count(*)
    Into n_Count
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 In (1, 3) And 序号 = 序号_In And NO = 单据号_In;
    If n_Count <> 0 Then
      v_Err_Msg := '挂号单据号重复,不能保存！' || Chr(13) || '如果使用了按日顺序编号,当日挂号量不能超过10000人次。';
      Raise Err_Item;
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
  End If;

  n_序号 := 号序_In;

  --获取是否分时段
  Begin
    Select Nvl(是否分时段, 0), Nvl(是否序号控制, 0), 限号数, 限约数
    Into n_分时段, n_序号控制, n_限号数, n_限约数
    From 临床出诊记录
    Where ID = 出诊记录id_In;
    n_原始分时段 := n_分时段;
  Exception
    When Others Then
      n_分时段     := 0;
      n_原始分时段 := n_分时段;
      n_序号控制   := 0;
      n_限号数     := Null;
      n_限约数     := Null;
  End;

  If n_序号 Is Null And n_分时段 = 1 And n_序号控制 = 0 Then
    Begin
      Select 序号
      Into n_序号
      From 临床出诊序号控制
      Where 记录id = 出诊记录id_In And 开始时间 = 发生时间_In And Rownum < 2;
    Exception
      When Others Then
        n_序号 := Null;
    End;
  End If;

  --分时段挂号时判断是否是过期挂号 也就是追加号的情况 目前只针对专家号分时段进行处理
  If Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 And Nvl(n_序号控制, 0) = 1 And 号序_In Is Null And Nvl(操作类型_In, 0) = 0 Then
    Begin
      Select Max(To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_最大序号时间
      From 临床出诊序号控制
      Where 记录id = 出诊记录id_In And Nvl(数量, 0) <> 0;
    
      n_追加号 := Case Sign(发生时间_In - d_最大序号时间)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_追加号 := 0;
    End;
  End If;
  d_时段时间 := 发生时间_In;

  If 序号_In = 1 And n_分时段 > 0 Then
    If Nvl(n_序号控制, 0) = 1 Then
      Begin
        Select Nvl(序号, 0),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               数量, Decode(Nvl(是否预约, 0), 0, 0, 数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And 序号 = n_序号;
      Exception
        When Others Then
          n_时段序号 := -1;
          n_分时段   := 0;
          d_时段时间 := 发生时间_In;
          n_时段限号 := 0;
          n_时段限约 := 0;
      End;
    Else
      --挂号时检查 是否分了时段,分了时段,把时段的限制条件给取出来
      Begin
        Select Nvl(序号, 0),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               数量, Decode(Nvl(是否预约, 0), 0, 0, 数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And 序号 = n_序号 And 预约顺序号 Is Null;
      Exception
        When Others Then
          n_时段序号 := -1;
          n_分时段   := 0;
          d_时段时间 := 发生时间_In;
          n_时段限号 := 0;
          n_时段限约 := 0;
      End;
    End If;
  End If;

  If 序号_In = 1 Then
    --获取当前未使用的序号
    If Nvl(预约挂号_In, 0) = 0 Then
      n_预约有效时间 := Zl_To_Number(zl_GetSysParameter('预约有效时间', 1111));
      n_失约挂号     := Zl_To_Number(zl_GetSysParameter('失约用于挂号', 1111));
    End If;
    If Nvl(n_序号控制, 0) = 1 And n_分时段 = 0 Then
      --<序号控制 未设置时段 获取可用的最大序号,以及已经使用的数量>
      Begin
        --最大序号
        Select Count(1) Into n_已用数量 From 病人挂号记录 Where 出诊记录id = 出诊记录id_In And 记录状态 = 1;
        Select Max(序号) Into n_已用序号 From 临床出诊序号控制 Where 记录id = 出诊记录id_In;
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
      End;
      Begin
        --最大序号
        Select Sum(Nvl(数量, 0))
        
        Into n_已约数
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And Nvl(挂号状态, 0) = 2;
      Exception
        When Others Then
          n_已约数 := 0;
      End;
    
      n_失效数 := 0;
      If Nvl(预约挂号_In, 0) = 0 Then
        If Nvl(n_预约有效时间, 0) <> 0 And Nvl(n_失约挂号, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_预约有效时间 / 24 / 60) - 预约时间), 1, 1, 0))
            Into n_失效数
            From 病人挂号记录
            Where 出诊记录id = 出诊记录id_In And 记录状态 = 1 And 记录性质 = 2;
          Exception
            When Others Then
              n_失效数 := 0;
          End;
        End If;
      End If;
    
      If n_原始分时段 = 0 Then
        Select Min(序号) Into n_已用序号 From 临床出诊序号控制 Where 记录id = 出诊记录id_In And Nvl(挂号状态, 0) = 0;
        If n_序号 Is Null Then
          n_序号 := Nvl(n_已用序号, 0);
        End If;
        IF nvl(n_序号,0)=0 THEN 
          Select Nvl(Max(序号), 0) + 1 Into n_序号 From 临床出诊序号控制 Where 记录id = 出诊记录id_In;
        END IF;
      Else
        Select Max(序号) Into n_已用序号 From 临床出诊序号控制 Where 记录id = 出诊记录id_In;
        If n_序号 Is Null Then
          n_序号 := Nvl(n_已用序号, 0) + 1;
        End If;
      End If;
      --<序号控制 未设置时段 获取可用的最大序号 以及已经使用的数量 --end>
    
      --非加号的情况需要检查是否超过了限制
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号时检查
          --启用序号控制未分时段 达到了限制
          If n_限号数 <= n_已用数量 And n_限号数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd ') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --预约时检查
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
        
          If n_限约数 <= n_已约数 And n_限约数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd') || '已达到最大限约数！';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --启用序号控制,未分时段 加号情况   不处理,如果以后有限制条件以后补充
      End If;
    
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 0 Then
      --<--普通号分时段 这里只有预约一种情况-->
      If 操作类型_In = 0 Then
        --<正常预约挂号-->
        Begin
          Select Count(0) As 已挂数, Nvl(Sum(Decode(Nvl(Sign(a.开始时间 - d_时段时间), 0), 0, 1, 0)), 0) As 已约数
          Into n_已挂数, n_已约数
          From 临床出诊序号控制 A
          Where 记录id = 出诊记录id_In And 挂号状态 Not In (0, 4);
        Exception
          When Others Then
            n_已挂数 := 0;
            n_已约数 := 0;
        End;
      
        n_时段限约 := n_时段限号; --普通号分时段的情况,29中n_时段限约始终是0 这里特殊处理
        --检查限制数量
        If n_限约数 = 0 Then
          n_限约数 := n_限号数;
        End If;
        If n_时段限约 <= n_已约数 Or n_限约数 <= n_已约数 Then
          v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限约数！';
          Raise Err_Item;
        End If;
        If n_限号数 <= n_已挂数 Then
          v_Err_Msg := '号别' || 号别_In || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
          Raise Err_Item;
        End If;
      End If;
    
      --没有达到时段的限号数 号码在当前时段往后追加
    
      --获取当天挂出的最大号序
      Select Nvl(Max(序号), 0)
      Into n_挂出的最大序号
      From 临床出诊序号控制 A
      Where 记录id = 出诊记录id_In And 预约顺序号 Is Null And 挂号状态 <> 0;
      If 预约顺序号_In Is Not Null Then
        n_预约顺序号 := 预约顺序号_In;
      Else
        Begin
          Select Nvl(Max(预约顺序号), 0) + 1
          Into n_预约顺序号
          From 临床出诊序号控制
          Where 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 Is Not Null;
        Exception
          When Others Then
            n_预约顺序号 := Null;
        End;
      End If;
      --设置序号
      n_序号 := RPad(Nvl(n_时段序号, 0), Length(n_限号数) + Length(Nvl(n_时段序号, 0)), 0) + n_预约顺序号;
      If n_预约顺序号 Is Null Then
        n_序号 := Nvl(n_挂出的最大序号, 0) + 1;
      End If;
    
      --<--普通号分时段--End>
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 1 Then
      --<启用序号控制 设置时段
      --专家号分时段
      Begin
        Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(开始时间 - d_时段时间), 0, 1, 0))
        Into n_已用序号, n_已挂数, n_已用数量
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And 挂号状态 Not In (0, 4);
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
          n_已挂数   := 0;
      End;
    
      n_失效数 := 0;
      If Nvl(预约挂号_In, 0) = 0 Then
        If Nvl(n_预约有效时间, 0) <> 0 And Nvl(n_失约挂号, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_预约有效时间 / 24 / 60) - 开始时间), 1, 1, 0))
            Into n_失效数
            From 临床出诊序号控制
            Where 记录id = 出诊记录id_In And 开始时间 Between Trunc(Sysdate) And Sysdate And Nvl(挂号状态, 0) = 2;
          Exception
            When Others Then
              n_失效数 := 0;
          End;
        End If;
      End If;
    
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
        
          --挂号 追加号码时不检查时段限号数
          If n_时段限号 <= n_已用数量 And Nvl(n_追加号, 0) = 0 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 - n_失效数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --挂号
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
          If n_限约数 <= n_已用数量 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_序号 Is Null Then
        --设置序号
        If Nvl(n_已用序号, 0) < Nvl(n_时段序号, 0) Then
          n_已用序号 := Nvl(n_时段序号, 0);
        End If;
        n_序号 := Nvl(n_已用序号, 0) + 1;
      End If;
    Elsif Nvl(n_分时段, 0) = 0 And Nvl(n_序号控制, 0) = 0 And Nvl(病历费_In, 0) = 0 And Nvl(号别_In, 0) > 0 Then
      ---<--普通号  -->
      Begin
        Select 已挂数, 已约数 Into n_已用数量, n_已约数 From 临床出诊记录 Where ID = 出诊记录id_In;
      Exception
        When Others Then
          n_已用数量 := 0;
          n_已约数   := 0;
      End;
      If Nvl(操作类型_In, 0) = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号
          If Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约
          If (Nvl(n_限约数, 0) > 0 And Nvl(n_已约数, 0) >= Nvl(n_限约数, 0)) Or
             (Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0) Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--普通号  -->
    End If;
  End If;

  --更新挂号序号状态
  If 序号_In = 1 And Not n_序号 Is Null Then
    If n_分时段 = 1 Then
      d_序号时间 := 发生时间_In;
    Else
      d_序号时间 := Trunc(发生时间_In);
    End If;
    --锁定序号的处理
    Begin
      If n_预约顺序号 Is Null Then
        Select 操作员姓名, 工作站名称
        Into v_序号操作员, v_序号机器名
        From 临床出诊序号控制
        Where Nvl(挂号状态, 0) = 5 And 记录id = 出诊记录id_In And 序号 = n_序号;
      Else
        Select 操作员姓名, 工作站名称
        Into v_序号操作员, v_序号机器名
        From 临床出诊序号控制
        Where Nvl(挂号状态, 0) = 5 And 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 = n_预约顺序号;
      End If;
      n_锁定 := 1;
    Exception
      When Others Then
        v_序号操作员 := Null;
        v_序号机器名 := Null;
        n_锁定       := 0;
    End;
    If n_锁定 = 0 Then
      If n_预约顺序号 Is Null Then
        Update 临床出诊序号控制
        Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
        Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) = 3 And 操作员姓名 = 操作员姓名_In;
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
        Where 记录id = 出诊记录id_In And 序号 = n_序号 And 预约顺序号 = n_预约顺序号 And Nvl(挂号状态, 0) = 3 And 操作员姓名 = 操作员姓名_In;
      End If;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_分时段, 0) > 0 Then
            If Nvl(n_序号控制, 0) = 1 Then
              --分时段后专家号 失约的预约号允许挂号
              Update 临床出诊序号控制
              Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In
              Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) In (0, 2);
              If Sql%NotFound Then
                Begin
                  Select 挂号状态 Into n_状态 From 临床出诊序号控制 Where 记录id = 出诊记录id_In And 序号 = n_序号;
                Exception
                  When Others Then
                    n_状态 := -1;
                End;
                If n_状态 = -1 Then
                  Insert Into 临床出诊序号控制
                    (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 类型, 名称, 操作员姓名, 备注)
                    Select 出诊记录id_In, n_序号, d_序号时间, d_序号时间, 1, Decode(预约挂号_In, 1, 1, 0), Decode(预约挂号_In, 1, 2, 1), Null,
                           Null, Null, 操作员姓名_In, '追加号'
                    From Dual;
                Else
                  v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
                  Raise Err_Item;
                End If;
              End If;
            Else
              If Nvl(预约接收_In, 0) = 1 Then
                Insert Into 临床出诊序号控制
                  (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 类型, 名称, 操作员姓名, 备注, 预约顺序号)
                  Select 记录id, 序号, 开始时间, 终止时间, 1, 1, Decode(预约挂号_In, 1, 2, 1), Null, Null, Null, 操作员姓名_In, n_序号, n_预约顺序号
                  From 临床出诊序号控制
                  Where 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 Is Null;
              End If;
            End If;
          Else
            If Nvl(n_序号控制, 0) = 1 Then
              Update 临床出诊序号控制
              Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In
              Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) = 0;
            
              If Sql%RowCount = 0 Then
                Begin
                  Select 挂号状态 Into n_状态 From 临床出诊序号控制 Where 记录id = 出诊记录id_In And 序号 = n_序号;
                Exception
                  When Others Then
                    n_状态 := -1;
                End;
                If n_状态 = -1 Then
                  Insert Into 临床出诊序号控制
                    (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 类型, 名称, 操作员姓名, 备注)
                    Select 出诊记录id_In, n_序号, 发生时间_In, 发生时间_In, 1, Decode(预约挂号_In, 1, 1, 0), Decode(预约挂号_In, 1, 2, 1),
                           Null, Null, Null, 操作员姓名_In, '追加号'
                    From Dual;
                Else
                  v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
                  Raise Err_Item;
                End If;
              End If;
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
            Raise Err_Item;
        End;
      End If;
    Else
      If 操作员姓名_In <> v_序号操作员 Or v_机器名 <> v_序号机器名 Then
        v_Err_Msg := '序号' || n_序号 || '已被自助机' || v_机器名 || '锁定,请重新选择一个序号.';
        Raise Err_Item;
      Else
        If n_预约顺序号 Is Null Then
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
          Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) = 5 And 操作员姓名 = 操作员姓名_In And 工作站名称 = v_机器名;
        Else
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
          Where 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 = n_预约顺序号 And Nvl(挂号状态, 0) = 5 And 操作员姓名 = 操作员姓名_In And
                工作站名称 = v_机器名;
        End If;
      End If;
    End If;
  End If;

  --产生病人挂号费用(可能单独是或包括病历费用)
  Select 病人费用记录_Id.Nextval Into n_费用id From Dual; --应该通过程序得到

  Insert Into 门诊费用记录
    (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id, 收费类别,
     计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人, 操作员编号, 操作员姓名,
     发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
  Values
    (n_费用id, 4, Decode(预约挂号_In, 1, 0, 1), 序号_In, Decode(价格父号_In, 0, Null, 价格父号_In), 从属父号_In, 单据号_In, 票据号_In, 1, 急诊_In,
     病历费_In, Decode(预约挂号_In, 1, To_Char(n_序号), 诊室_In), Decode(病人id_In, 0, Null, 病人id_In),
     Decode(门诊号_In, 0, Null, 门诊号_In), 付款方式_In, 姓名_In, Decode(姓名_In, Null, Null, 性别_In), Decode(姓名_In, Null, Null, 年龄_In),
     v_费别, 病人科室id_In, 收费类别_In, 号别_In, 收费细目id_In, 收入项目id_In, 收据费目_In, 1, 数次_In, 标准单价_In, 应收金额_In, 实收金额_In,
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额_In)),
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In)), Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 开单部门id_In,
     操作员姓名_In, Decode(预约挂号_In, 1, 操作员姓名_In, Null), 执行部门id_In, 医生姓名_In, 操作员编号_In, 操作员姓名_In, 发生时间_In, 登记时间_In, 保险大类id_In,
     保险项目否_In, 保险编码_In, 统筹金额_In, Decode(收费单_In, Null, 摘要_In, '划价:' || 收费单_In), 预约方式_In, Decode(预约挂号_In, 1, Null, n_组id));

  --汇总结算到病人预交记录
  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 0 Then
    If Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0 And 序号_In = 1 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
         结算性质)
      Values
        (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), v_现金, 0, 登记时间_In, 操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费',
         n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4);
    End If;
    If Nvl(现金支付_In, 0) <> 0 And 序号_In = 1 Then
      v_结算内容     := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
      v_结算方式记录 := '';
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
      
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
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, Null, Null, Null, Null, Null, 合作单位_In, 4, v_结算号码);
        Else
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4,
             v_结算号码);
        
          If Nvl(结算卡序号_In, 0) <> 0 And Nvl(现金支付_In, 0) <> 0 Then
            Zl_病人卡结算记录_支付(结算卡序号_In, 卡号_In, 0, Nvl(n_结算金额, 0), n_预交id, 操作员编号_In, 操作员姓名_In, 登记时间_In);
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
  
    --对于医保挂号
    If Nvl(个帐支付_In, 0) <> 0 And 序号_In = 1 Then
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), v_个人帐户, 个帐支付_In, 登记时间_In, 操作员编号_In,
         操作员姓名_In, 结帐id_In, '医保挂号', n_组id, 4);
    End If;
  
    --对于就诊卡通过预交金挂号
    If Nvl(预交支付_In, 0) <> 0 And 序号_In = 1 Then
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
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
           冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                 登记时间_In, 操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
      
        --更新病人预交余额
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2);
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
      If n_预交金额 > 0 Then
        v_Err_Msg := '预交余不够支付本次支付金额,不能继续操作！';
        Raise Err_Item;
      
      End If;
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
    End If;
  
    --相关汇总表的处理
    --人员缴款余额
    If 序号_In = 1 And Nvl(更新交款余额_In, 1) = 1 Then
      If Nvl(个帐支付_In, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + 个帐支付_In
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
          n_返回值 := 个帐支付_In;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_个人帐户 And Nvl(余额, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --病人挂号汇总(只处理一次,且单独收取病历费不处理)
  If Nvl(预约挂号_In, 0) = 0 Then
    --病人本次就诊(以发生时间为准)
    If Nvl(病人id_In, 0) <> 0 And 序号_In = 1 Then
      Update 病人信息 Set 就诊时间 = 发生时间_In, 就诊状态 = 1, 就诊诊室 = 诊室_In Where 病人id = 病人id_In;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
  
    --病人余额
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (病人id_In, 1, 1, Nvl(实收金额_In, 0), 0);
    End If;
  
    --病人未结费用
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(病人科室id_In, 0) And
          Nvl(开单部门id, 0) = Nvl(开单部门id_In, 0) And Nvl(执行部门id, 0) = Nvl(执行部门id_In, 0) And 收入项目id + 0 = 收入项目id_In And
          来源途径 + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (病人id_In, Null, Null, 病人科室id_In, 开单部门id_In, 执行部门id_In, 收入项目id_In, 1, Nvl(实收金额_In, 0));
    End If;
  End If;

  --病人挂号记录
  If 号别_In Is Not Null And 序号_In = 1 Then
    --And Nvl(预约挂号_In, 0) = 0
    Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
    Begin
      Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
    Exception
      When Others Then
        v_付款方式 := Null;
    End;
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 预约时间, 操作员编号,
       操作员姓名, 复诊, 号序, 社区, 预约, 预约方式, 摘要, 交易流水号, 交易说明, 合作单位, 接收时间, 接收人, 预约操作员, 预约操作员编号, 险类, 医疗付款方式, 出诊记录id, 收费单)
    Values
      (n_挂号id, 单据号_In, Decode(Nvl(预约挂号_In, 0), 1, 2, 1), 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 号别_In, 急诊_In, 诊室_In,
       Null, 执行部门id_In, 医生姓名_In, 0, Null, 登记时间_In, 发生时间_In, Decode(Nvl(预约挂号_In, 0), 1, 发生时间_In, Null), 操作员编号_In,
       操作员姓名_In, 复诊_In, n_序号, 社区_In, Decode(预约接收_In, 1, 1, 0), 预约方式_In, 摘要_In, 交易流水号_In, 交易说明_In, 合作单位_In,
       Decode(Nvl(预约挂号_In, 0), 0, 登记时间_In, Null), Decode(Nvl(预约挂号_In, 0), 0, 操作员姓名_In, Null),
       Decode(Nvl(预约挂号_In, 0), 1, 操作员姓名_In, Null), Decode(Nvl(预约挂号_In, 0), 1, 操作员编号_In, Null),
       Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, 出诊记录id_In, 收费单_In);
  
    If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
      Update 病人挂号记录
      Set 预约 = 1, 预约时间 = 发生时间_In, 预约操作员 = 操作员姓名_In, 预约操作员编号 = 操作员编号_In
      Where ID = n_挂号id;
    End If;
  
    n_预约生成队列 := 0;
    If Nvl(预约挂号_In, 0) = 1 Then
      n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    End If;
  
    --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
    If Nvl(生成队列_In, 0) <> 0 And Nvl(预约挂号_In, 0) = 0 Or n_预约生成队列 = 1 Then
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(执行部门id_In, 0)));
      If Nvl(n_分诊台签到排队, 0) = 0 Or n_预约生成队列 = 1 Then
        n_分时点显示 := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
        If Nvl(预约挂号_In, 0) = 1 And n_分时点显示 = 1 And n_分时段 = 1 Then
          n_分时点显示 := 1;
        Else
          n_分时点显示 := Null;
        End If;
        --产生队列
        --.按”执行部门” 的方式生成队列
        v_队列名称 := 执行部门id_In;
        v_排队号码 := Zlgetnextqueue(执行部门id_In, n_挂号id, 号别_In || '|' || n_序号);
        v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号别_In, n_序号, 发生时间_In);
        --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
        Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, 执行部门id_In, v_排队号码, Null, 姓名_In, 病人id_In, 诊室_In, 医生姓名_In, d_排队时间, 预约方式_In,
                         n_分时点显示, v_排队序号);
      
        --挂号立即排队
        If Nvl(n_分诊台签到排队, 0) = 0 Then
          Update 病人挂号记录 Set 记录标志 = 1 Where ID = n_挂号id;
        End If;
      End If;
    End If;
  End If;
  --病人担保信息
  If 病人id_In Is Not Null And 序号_In = 1 Then
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(结算模式, 0) Into n_结算模式 From 病人信息 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的病人信息,不允许挂号';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
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
  
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Exists (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
  
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = Sysdate
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, Sysdate) > Sysdate;
    End If;
  End If;
  If 序号_In = 1 Then
    --消息推送
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
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_出诊_Insert;
/

--138725:秦龙,2019-03-25,带量采购药品
Create Or Replace Procedure Zl_中选药品对照_Update(对照数据_In In Varchar2 --以"|"分隔的对照数据内容，每条记录按 序号^ 中选药品ID^ 对照药品ID
                                             ) Is
  v_Records    Varchar2(4000);
  v_Currrec    Varchar2(1000);
  v_Fields     Varchar2(1000);
  n_序号       中选药品对照.序号%Type;
  n_中选药品id 中选药品对照.中选药品id%Type;
  n_对照药品id 中选药品对照.对照药品id%Type;
Begin

  Delete From 中选药品对照;

  If 对照数据_In Is Null Then
    v_Records := Null;
  Else
    v_Records := 对照数据_In;
  End If;
  While v_Records Is Not Null Loop
    v_Currrec    := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
    v_Fields     := v_Currrec;
    n_序号       := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    n_中选药品id := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    n_对照药品id := To_Number(v_Fields);
  
    Insert Into 中选药品对照 (序号, 中选药品id, 对照药品id) Values (n_序号, n_中选药品id, n_对照药品id);
    v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_中选药品对照_Update;
/

Create Or Replace Procedure Zl_中选药品对照_Delete
(
  中选药品id_In In 中选药品对照.中选药品id%Type
) Is
Begin
  Delete From 中选药品对照 Where 中选药品id = 中选药品id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_中选药品对照_Delete;
/

--138725:秦龙,2019-03-25,增加传参“是否带量采购”
Create Or Replace Procedure Zl_成药规格_Update
(
  药品id_In         In 药品规格.药品id%Type,
  编码_In           In 收费项目目录.编码%Type,
  规格_In           In 收费项目目录.规格%Type,
  产地_In           In 收费项目目录.产地%Type := Null,
  品名_In           In 收费项目别名.名称%Type := Null,
  拼音_In           In 收费项目别名.简码%Type := Null,
  五笔_In           In 收费项目别名.简码%Type := Null,
  数字码_In         In 收费项目别名.简码%Type := Null,
  标识码_In         In 药品规格.标识码%Type := Null,
  药品来源_In       In 药品规格.药品来源%Type := Null,
  批准文号_In       In 药品规格.批准文号%Type := Null,
  注册商标_In       In 药品规格.注册商标%Type := Null,
  售价单位_In       In 收费项目目录.计算单位%Type := Null,
  剂量系数_In       In 药品规格.剂量系数%Type := Null,
  门诊单位_In       In 药品规格.门诊单位%Type := Null,
  门诊包装_In       In 药品规格.门诊包装%Type := Null,
  住院单位_In       In 药品规格.住院单位%Type := Null,
  住院包装_In       In 药品规格.住院包装%Type := Null,
  药库单位_In       In 药品规格.药库单位%Type := Null,
  药库包装_In       In 药品规格.药库包装%Type := Null,
  申领单位_In       In 药品规格.申领单位%Type := 1,
  申领阀值_In       In 药品规格.申领阀值%Type := Null,
  是否变价_In       In 收费项目目录.是否变价%Type := Null,
  指导批发价_In     In 药品规格.指导批发价%Type := Null,
  扣率_In           In 药品规格.扣率%Type := 95,
  指导零售价_In     In 药品规格.指导零售价%Type := Null,
  加成率_In         In 药品规格.加成率%Type := Null,
  管理费比例_In     In 药品规格.管理费比例%Type := Null,
  药价级别_In       In 药品规格.药价级别%Type := Null,
  费用类型_In       In 收费项目目录.费用类型%Type := Null,
  服务对象_In       In 收费项目目录.服务对象%Type := Null,
  Gmp认证_In        In 药品规格.Gmp认证%Type := 0,
  招标药品_In       In 药品规格.招标药品%Type := 0,
  屏蔽费别_In       In 收费项目目录.屏蔽费别%Type := 0,
  住院可否分零_In   In 药品规格.住院可否分零%Type := 0,
  药库分批_In       In 药品规格.药库分批%Type := Null,
  药房分批_In       In 药品规格.药房分批%Type := Null,
  最大效期_In       In 药品规格.最大效期%Type := Null,
  差价让利比_In     In 药品规格.差价让利比%Type := 0,
  成本价_In         In 药品规格.成本价%Type := 0,
  当前售价_In       In 收费价目.现价%Type := 0,
  收入id_In         In 收费价目.收入项目id%Type := Null,
  合同单位id_In     In 药品规格.合同单位id%Type := Null,
  说明_In           In 收费项目目录.说明%Type := Null,
  动态分零_In       In 药品规格.动态分零%Type := 0,
  发药类型_In       In 药品规格.发药类型%Type := Null,
  备选码_In         In 收费项目目录.备选码%Type := Null,
  增值税率_In       In 药品规格.增值税率%Type := Null,
  基本药物_In       In 药品规格.基本药物%Type := Null,
  站点_In           In 收费项目目录.站点%Type := Null,
  是否常备_In       In 药品规格.是否常备%Type := Null,
  存储温度_In       In 输液药品属性.存储温度%Type := Null,
  存储条件_In       In 输液药品属性.存储条件%Type := Null,
  配药类型_In       In 输液药品属性.配药类型%Type := Null,
  是否不予配置_In   In 输液药品属性.是否不予配置%Type := Null,
  容量_In           In 药品规格.容量%Type := Null,
  病案费目_In       In 收费项目目录.病案费目%Type := Null,
  门诊可否分零_In   In 药品规格.门诊可否分零%Type := 0,
  Ddd值_In          In 药品规格.Ddd值%Type := 0,
  高危药品_In       药品规格.高危药品%Type := Null,
  送货单位_In       In 药品规格.送货单位%Type := Null,
  送货包装_In       In 药品规格.送货包装%Type := Null,
  输液注意事项_In   In 输液药品属性.输液注意事项%Type := Null,
  是否摆药_In       In 药品规格.是否摆药%Type := Null,
  是否零差价管理_In In 药品规格.是否零差价管理%Type := Null,
  本位码_In         In 药品规格.本位码%Type := Null,
  是否易至跌倒_In   In 药品规格.是否易至跌倒%Type := Null,
  带量采购_In       In 药品规格.是否带量采购%Type := Null
) Is
  v_药名id   诊疗项目目录.Id%Type;
  v_名称     诊疗项目目录.名称%Type;
  v_是否变价 收费项目目录.是否变价%Type; --允许定价药品随时改为时价，时价药品只能在未发生的情况下修改为定价，其它情况不允许修改定价属性 
  v_发生     Number(2);
  Err_Notfind Exception;
  v_No           收费价目.No%Type;
  v_Temp         收费项目目录.病案费目%Type;
  v_病案费目     收费项目目录.病案费目%Type;
  n_指导差价率   药品规格.指导差价率%Type;
  n_药品上次售价 药品规格.上次售价%Type;
  n_零售金额     药品库存.实际金额%Type;
  n_收发id       药品收发记录.Id%Type;
  n_流通金额小数 Number;
  n_序号         Number(8);
  Classid        Number(18); --入出类别
  v_Billno       药品收发记录.No%Type; --调价单号
  n_价格id       收费价目.Id%Type;
  n_收费价目现价 收费价目.现价%Type;
  n_原价         药品价格记录.原价%Type;
  n_药品价格记录 Number(1);
  v_类别         收费项目目录.类别%Type;
  --定价->时价后更新药品价格记录的值

  Cursor c_Priceadjust Is
    Select s.药品id, s.库房id, Nvl(s.批次, 0) As 批次, s.上次供应商id As 供应商id, s.上次批号 As 批号, s.效期, s.上次产地 As 产地,
           Nvl(s.实际数量, 0) As 实际数量, s.上次扣率 As 扣率, Nvl(s.实际金额, 0) As 实际金额, Nvl(s.实际差价, 0) As 实际差价, Nvl(s.零售价, 0) As 零售价,
           s.平均成本价, s.灭菌效期, s.批准文号, s.上次生产日期 As 生产日期
    From 药品库存 S
    Where s.药品id = 药品id_In And s.性质 = 1 
    Order By s.药品id, s.批次, s.库房id;

  r_Priceadjust c_Priceadjust%RowType;

Begin
  v_病案费目 := 病案费目_In;
  --判断病案费目 
  If v_病案费目 Is Null Then
    If 收入id_In Is Not Null Then
      Begin
        Select 病案费目 Into v_Temp From 收入项目 Where ID = 收入id_In;
      Exception
        When Others Then
          v_Temp := Null;
      End;
      If v_Temp Is Not Null Then
        v_病案费目 := v_Temp;
      End If;
    End If;
  End If;
  --通用名称 
  Select ID, 名称
  Into v_药名id, v_名称
  From 诊疗项目目录
  Where ID = (Select 药名id From 药品规格 Where 药品id = 药品id_In);
  --取原始的定价属性 
  Select 是否变价 Into v_是否变价 From 收费项目目录 Where ID = 药品id_In;
  --规格信息 
  Update 收费项目目录
  Set 编码 = 编码_In, 名称 = v_名称, 规格 = 规格_In, 产地 = 产地_In, 计算单位 = 售价单位_In, 费用类型 = 费用类型_In, 服务对象 = 服务对象_In, 屏蔽费别 = 屏蔽费别_In,
      病案费目 = v_病案费目, 说明 = 说明_In, 备选码 = 备选码_In, 站点 = 站点_In
  Where ID = 药品id_In
  Returning 类别 Into v_类别;
  If Sql%RowCount = 0 Then
    Raise Err_Notfind;
  End If;
  n_指导差价率 := (1 - 1 / (1 + 加成率_In / 100)) * 100;
  Update 药品规格
  Set 标识码 = 标识码_In, 药品来源 = 药品来源_In, 批准文号 = 批准文号_In, 注册商标 = 注册商标_In, 剂量系数 = 剂量系数_In, 门诊单位 = 门诊单位_In, 门诊包装 = 门诊包装_In,
      住院单位 = 住院单位_In, 住院包装 = 住院包装_In, 药库单位 = 药库单位_In, 药库包装 = 药库包装_In, 申领单位 = 申领单位_In, 申领阀值 = 申领阀值_In, 指导批发价 = 指导批发价_In,
      扣率 = 扣率_In, 指导零售价 = 指导零售价_In, 指导差价率 = n_指导差价率, 管理费比例 = 管理费比例_In, 药价级别 = 药价级别_In, 住院可否分零 = 住院可否分零_In,
      药库分批 = 药库分批_In, 药房分批 = 药房分批_In, 最大效期 = 最大效期_In, 招标药品 = 招标药品_In, Gmp认证 = Gmp认证_In, 差价让利比 = 差价让利比_In,
      合同单位id = 合同单位id_In, 动态分零 = 动态分零_In, 发药类型 = 发药类型_In, 增值税率 = 增值税率_In, 基本药物 = 基本药物_In, 是否常备 = 是否常备_In, 容量 = 容量_In,
      门诊可否分零 = 门诊可否分零_In, Ddd值 = Ddd值_In, 高危药品 = 高危药品_In, 送货单位 = 送货单位_In, 送货包装 = 送货包装_In, 加成率 = 加成率_In, 是否摆药 = 是否摆药_In,
      是否零差价管理 = 是否零差价管理_In, 本位码 = 本位码_In, 是否易至跌倒 = 是否易至跌倒_In, 是否带量采购 = 带量采购_In
  Where 药品id = 药品id_In;

  --如果没有带量采购属性，则删除中选药品对照(可能对照表本来也没有数据，但执行该过程没有影响)
  If Nvl(带量采购_In, 0) = 0 Then
    Zl_中选药品对照_Delete(药品id_In);
  End If;

  --别名的处理 
  If 数字码_In Is Null Then
    Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 1 And 码类 = 3;
  Else
    Update 收费项目别名 Set 名称 = v_名称, 简码 = 数字码_In Where 收费细目id = 药品id_In And 性质 = 1 And 码类 = 3;
    If Sql%RowCount = 0 Then
      Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, v_名称, 1, 数字码_In, 3);
    End If;
  End If;
  If 品名_In Is Null Then
    Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 3;
  Else
    If 拼音_In Is Null Then
      Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 1;
    Else
      Update 收费项目别名 Set 名称 = 品名_In, 简码 = 拼音_In Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 1;
      If Sql%RowCount = 0 Then
        Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, 品名_In, 3, 拼音_In, 1);
      End If;
    End If;
    If 五笔_In Is Null Then
      Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 2;
    Else
      Update 收费项目别名 Set 名称 = 品名_In, 简码 = 五笔_In Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 2;
      If Sql%RowCount = 0 Then
        Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, 品名_In, 3, 五笔_In, 2);
      End If;
    End If;
  End If;

  --定价信息：如果已经有发生，则不允许直接更改这些信息 
  Select Nvl(Count(*), 0) Into v_发生 From 药品收发记录 Where 药品id = 药品id_In And Rownum < 2;
  If v_发生 = 0 Then
    Update 药品规格 Set 成本价 = 成本价_In Where 药品id = 药品id_In;
    If 收入id_In Is Not Null Then
      Update 收费价目
      Set 现价 = 当前售价_In, 收入项目id = 收入id_In, 变动原因 = 1, 调价说明 = '修改定价', 调价人 = User
      Where 收费细目id = 药品id_In And (Sysdate Between 执行日期 And 终止日期 Or Sysdate >= 执行日期 And 终止日期 Is Null) And 变动原因 = 1;
      If Sql%RowCount = 0 Then
        v_No := Nextno(9);
        Insert Into 收费价目
          (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
        Values
          (收费价目_Id.Nextval, Null, 药品id_In, 0, 当前售价_In, 收入id_In, 1, '新增定价', User, Sysdate,
           To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
      End If;
    End If;
  Else
    --发生业务数据时，不能修改价格但是可以修改收入项目 
    Update 收费价目
    Set 收入项目id = 收入id_In
    Where 收费细目id = 药品id_In And (Sysdate Between 执行日期 And 终止日期 Or Sysdate >= 执行日期 And 终止日期 Is Null) And 变动原因 = 1;
  End If;

  --时价->定价
  If v_是否变价 = 1 And 是否变价_In = 0 Then
    Update 收费项目目录 Set 是否变价 = 是否变价_In Where ID = 药品id_In;
  
    Begin
      Select 现价, ID As 价格id
      Into n_收费价目现价, n_价格id
      From 收费价目
      Where 收费细目id = 药品id_In And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And 变动原因 = 1;
    Exception
      When Others Then
        n_收费价目现价 := Null;
        n_价格id       := Null;
    End;
  
    Begin
      Select 上次售价 Into n_药品上次售价 From 药品规格 Where 药品id = 药品id_In;
    Exception
      When Others Then
        n_药品上次售价 := Null;
    End;
  
    If n_药品上次售价 Is Null Then
      n_药品上次售价 := n_收费价目现价;
    End If;
  
    Zl_收费价目_Stop(药品id_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into 收费价目
      (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
    Values
      (收费价目_Id.Nextval, n_价格id, 药品id_In, n_收费价目现价, n_药品上次售价, 收入id_In, 1, '时价转定价', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;
  
    --取入出类别ID
    Select 类别id Into Classid From 药品单据性质 Where 单据 = 13;
  
    n_序号   := 0;
    v_Billno := Null;
  
    For r_Priceadjust In c_Priceadjust Loop
      If n_药品上次售价 <> r_Priceadjust.零售价 Then
        If v_Billno Is Null Then
          Select Nextno(147) Into v_Billno From Dual;
        End If;
        n_序号 := n_序号 + 1;
        Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
        n_零售金额 := Round(n_药品上次售价 * r_Priceadjust.实际数量, n_流通金额小数) -
                  Round(r_Priceadjust.零售价 * r_Priceadjust.实际数量, n_流通金额小数);
        --产生调价影响记录
        Insert Into 药品收发记录
          (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要, 填制人,
           填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次, 供药单位id)
        Values
          (n_收发id, 1, 13, v_Billno, n_序号, Classid, r_Priceadjust.药品id, r_Priceadjust.批次, r_Priceadjust.批号,
           r_Priceadjust.效期, r_Priceadjust.产地, 1, r_Priceadjust.实际数量, 0, r_Priceadjust.零售价, 0, n_药品上次售价,
           r_Priceadjust.扣率, n_零售金额, n_零售金额, '时价转定价', Zl_Username, Sysdate, r_Priceadjust.库房id, 1, n_价格id, Zl_Username,
           Sysdate, r_Priceadjust.实际金额, r_Priceadjust.实际差价, r_Priceadjust.供应商id);
      
        Zl_药品库存_Update(n_收发id, 2, 0);
      End If;
    End Loop;
  
    --定价->时价
  Elsif v_是否变价 = 0 And 是否变价_In = 1 Then
    Update 收费项目目录 Set 是否变价 = 是否变价_In Where ID = 药品id_In;
    Begin
      Select 现价, ID As 价格id
      Into n_收费价目现价, n_价格id
      From 收费价目
      Where 收费细目id = 药品id_In And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And 变动原因 = 1;
    Exception
      When Others Then
        n_收费价目现价 := Null;
        n_价格id       := Null;
    End;
  
    Zl_收费价目_Stop(药品id_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into 收费价目
      (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
    Values
      (收费价目_Id.Nextval, n_价格id, 药品id_In, n_收费价目现价, n_收费价目现价, 收入id_In, 1, '定价转时价', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    For r_Priceadjust In c_Priceadjust Loop
      n_药品价格记录 := 0;
      Begin
        Select 1, 现价
        Into n_药品价格记录, n_原价
        From 药品价格记录
        Where 药品id = r_Priceadjust.药品id And 库房id = r_Priceadjust.库房id And Nvl(批次, 0) = r_Priceadjust.批次 And 记录状态 = 1 And
              价格类型 = 1;
      Exception
        When Others Then
          n_药品价格记录 := 0;
          n_原价         := n_收费价目现价;
      End;
    
      If n_药品价格记录 = 1 Then
        Zl_药品价格记录_Stop(1, r_Priceadjust.库房id, r_Priceadjust.药品id, r_Priceadjust.批次, Sysdate - 1 / 24 / 60 / 60, 2);
      End If;
      Zl_药品价格记录_Insert(0, 1, r_Priceadjust.库房id, r_Priceadjust.药品id, r_Priceadjust.批次, n_原价, n_收费价目现价, Sysdate, '定价转时价',
                       Zl_Username, Null, r_Priceadjust.供应商id, r_Priceadjust.批号, r_Priceadjust.效期, r_Priceadjust.产地,
                       r_Priceadjust.灭菌效期, Null, Null, Null, Null, 1);
    
      Update 药品库存
      Set 零售价 = n_收费价目现价
      Where 性质 = 1 And 库房id = r_Priceadjust.库房id And 药品id = r_Priceadjust.药品id And Nvl(批次, 0) = r_Priceadjust.批次;
    
    End Loop;
  End If;

  --药品生产商比较增加 
  If 产地_In Is Not Null Then
    Update 药品生产商 Set 名称 = 产地_In Where 名称 = 产地_In;
    If Sql%RowCount = 0 Then
      Insert Into 药品生产商
        (编码, 名称, 简码)
        Select Nvl(Max(To_Number(编码)), 0) + 1, 产地_In, zlSpellCode(产地_In, 10) From 药品生产商;
    End If;
  End If;

  --修改输液药品属性 
  Update 输液药品属性
  Set 存储温度 = 存储温度_In, 存储条件 = 存储条件_In, 配药类型 = 配药类型_In, 是否不予配置 = 是否不予配置_In, 输液注意事项 = 输液注意事项_In
  Where 药品id = 药品id_In;

  If Sql%NotFound Then
    Insert Into 输液药品属性
      (药品id, 存储温度, 存储条件, 配药类型, 是否不予配置, 输液注意事项)
    Values
      (药品id_In, 存储温度_In, 存储条件_In, 配药类型_In, 是否不予配置_In, 输液注意事项_In);
  End If;

  --药品精度调整(零差价模式时)
  Zl_药品卫材精度_零差价调整;

  b_Message.Zlhis_Dict_036(v_类别, 药品id_In);
Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]该规格不存在，可能已被其他用户删除！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_成药规格_Update;
/

--138725:秦龙,2019-03-25,增加传参“,是否带量采购”
Create Or Replace Procedure Zl_成药规格_Insert
(
  药名id_In         In 药品规格.药名id%Type,
  药品id_In         In 药品规格.药品id%Type,
  编码_In           In 收费项目目录.编码%Type,
  规格_In           In 收费项目目录.规格%Type,
  产地_In           In 收费项目目录.产地%Type := Null,
  品名_In           In 收费项目别名.名称%Type := Null,
  拼音_In           In 收费项目别名.简码%Type := Null,
  五笔_In           In 收费项目别名.简码%Type := Null,
  数字码_In         In 收费项目别名.简码%Type := Null,
  标识码_In         In 药品规格.标识码%Type := Null,
  药品来源_In       In 药品规格.药品来源%Type := Null,
  批准文号_In       In 药品规格.批准文号%Type := Null,
  注册商标_In       In 药品规格.注册商标%Type := Null,
  售价单位_In       In 收费项目目录.计算单位%Type := Null,
  剂量系数_In       In 药品规格.剂量系数%Type := Null,
  门诊单位_In       In 药品规格.门诊单位%Type := Null,
  门诊包装_In       In 药品规格.门诊包装%Type := Null,
  住院单位_In       In 药品规格.住院单位%Type := Null,
  住院包装_In       In 药品规格.住院包装%Type := Null,
  药库单位_In       In 药品规格.药库单位%Type := Null,
  药库包装_In       In 药品规格.药库包装%Type := Null,
  申领单位_In       In 药品规格.申领单位%Type := 1,
  申领阀值_In       In 药品规格.申领阀值%Type := Null,
  是否变价_In       In 收费项目目录.是否变价%Type := Null,
  指导批发价_In     In 药品规格.指导批发价%Type := Null,
  扣率_In           In 药品规格.扣率%Type := 95,
  指导零售价_In     In 药品规格.指导零售价%Type := Null,
  加成率_In         In 药品规格.加成率%Type := Null,
  管理费比例_In     In 药品规格.管理费比例%Type := Null,
  药价级别_In       In 药品规格.药价级别%Type := Null,
  费用类型_In       In 收费项目目录.费用类型%Type := Null,
  服务对象_In       In 收费项目目录.服务对象%Type := Null,
  Gmp认证_In        In 药品规格.Gmp认证%Type := 0,
  招标药品_In       In 药品规格.招标药品%Type := 0,
  屏蔽费别_In       In 收费项目目录.屏蔽费别%Type := 0,
  住院可否分零_In   In 药品规格.住院可否分零%Type := 0,
  药库分批_In       In 药品规格.药库分批%Type := Null,
  药房分批_In       In 药品规格.药房分批%Type := Null,
  最大效期_In       In 药品规格.最大效期%Type := Null,
  差价让利比_In     In 药品规格.差价让利比%Type := 0,
  成本价_In         In 药品规格.成本价%Type := 0,
  当前售价_In       In 收费价目.现价%Type := 0,
  收入id_In         In 收费价目.收入项目id%Type := Null,
  合同单位id_In     In 药品规格.合同单位id%Type := Null,
  说明_In           In 收费项目目录.说明%Type := Null,
  动态分零_In       In 药品规格.动态分零%Type := 0,
  发药类型_In       In 药品规格.发药类型%Type := Null,
  备选码_In         In 收费项目目录.备选码%Type := Null,
  增值税率_In       In 药品规格.增值税率%Type := Null,
  基本药物_In       In 药品规格.基本药物%Type := Null,
  站点_In           In 收费项目目录.站点%Type := Null,
  是否常备_In       In 药品规格.是否常备%Type := Null,
  存储温度_In       In 输液药品属性.存储温度%Type := Null,
  存储条件_In       In 输液药品属性.存储条件%Type := Null,
  配药类型_In       In 输液药品属性.配药类型%Type := Null,
  是否不予配置_In   In 输液药品属性.是否不予配置%Type := Null,
  容量_In           In 药品规格.容量%Type := Null,
  病案费目_In       In 收费项目目录.病案费目%Type := Null,
  门诊可否分零_In   In 药品规格.门诊可否分零%Type := 0,
  Ddd值_In          In 药品规格.Ddd值%Type := 0,
  高危药品_In       In 药品规格.高危药品%Type := Null,
  送货单位_In       In 药品规格.送货单位%Type := Null,
  送货包装_In       In 药品规格.送货包装%Type := Null,
  输液注意事项_In   In 输液药品属性.输液注意事项%Type := Null,
  是否摆药_In       In 药品规格.是否摆药%Type := Null,
  是否零差价管理_In In 药品规格.是否零差价管理%Type := Null,
  本位码_In         In 药品规格.本位码%Type := Null,
  是否易至跌倒_In   In 药品规格.是否易至跌倒%Type := Null,
  带量采购_In       In 药品规格.是否带量采购%Type := Null
) Is

  v_类别       诊疗项目目录.类别%Type;
  v_名称       诊疗项目目录.名称%Type;
  v_Kind       Varchar2(20);
  v_No         收费价目.No%Type;
  v_Temp       收费项目目录.病案费目%Type;
  v_病案费目   收费项目目录.病案费目%Type;
  n_指导差价率 药品规格.指导差价率%Type;

  --盘点库房的工作性质 
  Cursor c_Storageid Is
    Select Distinct 部门id From 部门性质说明 Where 工作性质 Like v_Kind Or 工作性质 = '制剂室';
  r_Storageid c_Storageid%RowType;
Begin
  v_病案费目 := 病案费目_In;
  --判断病案费目 
  If v_病案费目 Is Null Then
    If 收入id_In Is Not Null Then
      Begin
        Select 病案费目 Into v_Temp From 收入项目 Where ID = 收入id_In;
      Exception
        When Others Then
          v_Temp := Null;
      End;
      If v_Temp Is Not Null Then
        v_病案费目 := v_Temp;
      End If;
    End If;
  End If;
  --类别和名称 
  Select 类别, 名称 Into v_类别, v_名称 From 诊疗项目目录 Where ID = 药名id_In;
  n_指导差价率 := (1 - 1 / (1 + 加成率_In / 100)) * 100;
  --规格信息 
  Insert Into 收费项目目录
    (类别, ID, 编码, 名称, 规格, 产地, 计算单位, 费用类型, 服务对象, 屏蔽费别, 是否变价, 建档时间, 撤档时间, 说明, 备选码, 站点, 病案费目)
  Values
    (v_类别, 药品id_In, 编码_In, v_名称, 规格_In, 产地_In, 售价单位_In, 费用类型_In, 服务对象_In, 屏蔽费别_In, 是否变价_In, Sysdate,
     To_Date('3000-01-01', 'YYYY-MM-DD'), 说明_In, 备选码_In, 站点_In, v_病案费目);
  Insert Into 药品规格
    (药名id, 药品id, 标识码, 药品来源, 批准文号, 注册商标, 剂量系数, 门诊单位, 门诊包装, 住院单位, 住院包装, 药库单位, 药库包装, 申领单位, 申领阀值, 指导批发价, 扣率, 指导零售价, 指导差价率,
     管理费比例, 药价级别, 成本价, Gmp认证, 招标药品, 差价让利比, 住院可否分零, 药库分批, 药房分批, 最大效期, 合同单位id, 动态分零, 发药类型, 增值税率, 基本药物, 是否常备, 容量, 门诊可否分零,
     Ddd值, 高危药品, 送货单位, 送货包装, 加成率, 是否摆药, 是否零差价管理, 本位码, 是否易至跌倒, 是否带量采购)
  Values
    (药名id_In, 药品id_In, 标识码_In, 药品来源_In, 批准文号_In, 注册商标_In, 剂量系数_In, 门诊单位_In, 门诊包装_In, 住院单位_In, 住院包装_In, 药库单位_In, 药库包装_In,
     申领单位_In, 申领阀值_In, 指导批发价_In, 扣率_In, 指导零售价_In, n_指导差价率, 管理费比例_In, 药价级别_In, 成本价_In, Gmp认证_In, 招标药品_In, 差价让利比_In,
     住院可否分零_In, 药库分批_In, 药房分批_In, 最大效期_In, 合同单位id_In, 动态分零_In, 发药类型_In, 增值税率_In, 基本药物_In, 是否常备_In, 容量_In, 门诊可否分零_In,
     Ddd值_In, 高危药品_In, 送货单位_In, 送货包装_In, 加成率_In, 是否摆药_In, 是否零差价管理_In, 本位码_In, 是否易至跌倒_In, 带量采购_In);

  --朱玉宝修改：建立药品（西成药、中成药）时，缺省服务对象为门诊和住院，因此建立规格药品时，不再根据规格药品的服务对象更新药品的服务对象 
  --诊疗项目服务对象的更改 
  --select nvl(sum(distinct I.服务对象),0) into v_对象 
  --from 收费项目目录 I,药品规格 S 
  --where I.ID=S.药品ID and S.药名ID=药名ID_IN; 
  --update 诊疗项目目录 
  --set 服务对象=decode(v_对象,0,0,1,1,2,2,3) 
  --where ID=药名ID_IN; 

  --别名的处理 
  Insert Into 收费项目别名
    (收费细目id, 名称, 性质, 简码, 码类)
    Select 药品id_In, 名称, 性质, 简码, 码类 From 诊疗项目别名 Where 诊疗项目id = 药名id_In;
  If 数字码_In Is Not Null Then
    Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, v_名称, 1, 数字码_In, 3);
  End If;
  If (品名_In Is Not Null) And (拼音_In Is Not Null) Then
    Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, 品名_In, 3, 拼音_In, 1);
  End If;
  If (品名_In Is Not Null) And (五笔_In Is Not Null) Then
    Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, 品名_In, 3, 五笔_In, 2);
  End If;

  --定价信息 
  If 收入id_In Is Not Null Then
    v_No := Nextno(9);
    Insert Into 收费价目
      (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
    Values
      (收费价目_Id.Nextval, Null, 药品id_In, 0, 当前售价_In, 收入id_In, 1, '新增定价', User, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  End If;

  --药品生产商比较增加 
  If 产地_In Is Not Null Then
    Update 药品生产商 Set 名称 = 产地_In Where 名称 = 产地_In;
    If Sql%RowCount = 0 Then
      Insert Into 药品生产商
        (编码, 名称, 简码)
        Select Nvl(Max(To_Number(编码)), 0) + 1, 产地_In, zlSpellCode(产地_In, 10) From 药品生产商;
    End If;
  End If;

  --插入该规格的服务科室 
  Insert Into 收费执行科室
    (收费细目id, 病人来源, 开单科室id, 执行科室id)
    Select 药品id_In, 病人来源, 开单科室id, 执行科室id From 诊疗执行科室 Where 诊疗项目id = 药名id_In;

  --插入盘点属性 

  If v_类别 = 5 Then
    v_Kind := '西药%';
  Else
    v_Kind := '成药%';
  End If;

  For r_Storageid In c_Storageid Loop
    Insert Into 药品储备限额
      (库房id, 药品id, 上限, 下限, 盘点属性, 库房货位)
    Values
      (r_Storageid.部门id, 药品id_In, 0, 0, '1111', Null);
  End Loop;

  --插入输液药品属性 
  Insert Into 输液药品属性
    (药品id, 存储温度, 存储条件, 配药类型, 是否不予配置, 输液注意事项)
  Values
    (药品id_In, 存储温度_In, 存储条件_In, 配药类型_In, 是否不予配置_In, 输液注意事项_In);

  --药品精度调整(零差价模式时)
  Zl_药品卫材精度_零差价调整;

  b_Message.Zlhis_Dict_035(v_类别, 药品id_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_成药规格_Insert;
/





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0055' Where 编号=&n_System;
Commit;