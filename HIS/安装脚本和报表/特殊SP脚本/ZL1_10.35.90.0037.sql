----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------

--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
Alter Table 挂号安排计划 Add 上次计划ID number(18);

Create Index 挂号安排计划_IX_上次计划ID on 挂号安排计划(上次计划ID) Tablespace zl9Indexhis;

Alter table 挂号安排计划
  add constraint 挂号安排计划_FK_上次计划ID foreign key (上次计划ID)
  references 挂号安排计划 (ID);

------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
Declare
  Cursor c_长期计划 Is
          Select Id, 安排id, 生效时间, 失效时间, 审核时间
          From 挂号安排计划
          Where 失效时间 = To_Date('3000-01-01', 'yyyy-mm-dd') And 审核时间 Is Not Null
          Order By 安排id, 生效时间, Id;

  n_计划ID 挂号安排计划.ID%Type;
  n_安排ID 挂号安排计划.安排ID%Type;
Begin
  For r_长期计划 In c_长期计划 Loop
    IF Nvl(n_安排ID, 0) <> r_长期计划.安排ID Then
      n_安排ID := r_长期计划.安排ID;
      n_计划ID := 0;
    End if;
    IF Nvl(n_计划ID, 0) <> 0 Then
      Update 挂号安排计划 Set 失效时间 = r_长期计划.生效时间 Where ID = n_计划ID;
      Update 挂号安排计划 Set 上次计划ID = n_计划ID Where ID = r_长期计划.ID;
    End IF;
    n_计划ID := r_长期计划.ID;
  End Loop;
End;
/

-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--134254:王振涛,2018-11-16,添加三方报告打印
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1214, '基本', User, 'Zl_Lob_Read', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1214 And 功能 = '基本' And Upper(对象) = Upper('Zl_Lob_Read'));


-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
Create Or Replace Procedure Zl_挂号安排计划_Verify
(
  Id_In         In 挂号安排计划.Id%Type,
  立即生效_In   Number := 0,
  上次计划ID_In In 挂号安排计划.上次计划Id%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg   Varchar2(100);
  v_User_Name 人员表.姓名%Type;
  n_Valied    Number(1);
  d_生效时间  挂号安排计划.生效时间%Type;
Begin
  
  Select Nvl(Max(p.姓名),'') Into v_User_Name From 上机人员表 o, 人员表 p Where o.人员id = p.Id And 用户名 = User;
  If v_User_Name Is Null Then
    v_Err_Msg := '[ZLSOFT]当前用户未设置对应的人员信息,请与' || Chr(10) || Chr(13) ||
                 '系统管理员联系,先到用户授权管理中设置！[ZLSOFT]';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Valied From 挂号安排计划 a Where Nvl(生效时间, Sysdate) < Sysdate And a.Id = Id_In;
  If n_Valied > 0 And Nvl(立即生效_In, 0) = 0 Then
    v_Err_Msg := '[ZLSOFT]该计划安排的生效时间已经到期，不能进行审核！[ZLSOFT]';
    Raise Err_Item;
  End If;
  
  Update 挂号安排计划
  Set 审核人 = v_User_Name, 审核时间 = Sysdate, 上次计划ID = 上次计划ID_In,
      生效时间 = Case Nvl(立即生效_In, 0) When 0 Then 生效时间 Else Sysdate - 1 / 24 / 60 / 60 End
  Where Id = Id_In And 审核时间 Is Null
  Return 生效时间 Into d_生效时间;
  If Sql%Notfound Then
    v_Err_Msg := '[ZLSOFT]该计划安排已经被他人审核或删除,不能再审核![ZLSOFT]';
    Raise Err_Item;
  End If;
  IF Nvl(上次计划ID_In, 0) <> 0 Then
    Update 挂号安排计划 Set 失效时间 = d_生效时间 Where ID = 上次计划ID_In;
  End IF;
  If Nvl(立即生效_In, 0) = 1 Then
    Begin
      Zl_挂号安排_Autoupdate();
    Exception
      When Others Then
        Null;
    End;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_挂号安排计划_Verify;
/

--133584:李南春,2018-11-15,利用失效时间查询挂号安排计划
Create Or Replace Procedure Zl_挂号安排计划_Cancel(Id_In In 挂号安排计划.ID%Type) Is
  Err_Item        Exception;
  v_Err_Msg       Varchar2(100);
  v_User_Name     人员表.姓名%Type;
  n_上次计划ID    挂号安排计划.上次计划Id%Type;
Begin
  Begin
    Select P.姓名 Into v_User_Name From 上机人员表 O, 人员表 P Where O.人员id = P.ID And 用户名 = User;
  Exception
    When Others Then
      v_User_Name := Null;
  End;
  If v_User_Name Is Null Then
    v_Err_Msg := '[ZLSOFT]当前用户未设置对应的人员信息,请与' || Chr(10) || Chr(13) ||
                 '系统管理员联系,先到用户授权管理中设置！[ZLSOFT]';
    Raise Err_Item;
  End If;
  Begin
    Select 'Yes'
    Into v_Err_Msg
    From 挂号安排计划 L
    Where ID = Id_In And Nvl(实际生效, To_Date('3000-01-01', 'yyyy-mm-dd')) >= To_Date('3000-01-01', 'yyyy-mm-dd');
  Exception
    When Others Then
      v_Err_Msg := 'No';
  End;
  If v_Err_Msg = 'No' Then
    v_Err_Msg := '[ZLSOFT]该计划安排已经被生效,不能取消审核![ZLSOFT]';
    Raise Err_Item;
  End If;

  Update 挂号安排计划 Set 审核人 = Null, 审核时间 = Null Where ID = Id_In And 审核时间 Is Not Null
  Return 上次计划ID Into n_上次计划ID;
  If Sql%NotFound Then
    v_Err_Msg := '[ZLSOFT]该计划安排已经被他人取消审核或删除,不能再取消审核![ZLSOFT]';
    Raise Err_Item;
  End If;
  If Nvl(n_上次计划ID, 0) <> 0 Then
    Update 挂号安排计划 Set 失效时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Where ID = n_上次计划ID;
    Update 挂号安排计划 Set 上次计划ID = NULL Where ID = Id_In;
  End IF;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);

  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_挂号安排计划_Cancel;
/

--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
Create Or Replace Procedure Zl_挂号安排_Autoupdate Is
  Err_Item Exception;
  v_Date Date;
  -- v_Err_Msg Varchar2(100); 
  v_Unitscount Number;
Begin
  --n_更新执行人 ：是否更新病人挂号记录 和门诊费用记录中的执行人 
  --               如果计划中更改了 挂号项目 则不允许更新 病人挂号记录和门诊费用记录中的数据 
  Select Sysdate Into v_Date From Dual;
  Select Count(0) Into v_Unitscount From 合作单位安排控制 Where Rownum = 1;

  For v_生效 In (Select ID, 安排id, 号码, 生效时间, 失效时间, 周日, 周一, 周二, 周三, 周四, 周五, 周六, 分诊方式, 序号控制, 执行时间 As 上次生效时间, 项目id, 医生姓名, 医生id,
                      序号, 科室id
               From (Select a.Id, a.安排id, a.号码, a.生效时间, a.失效时间, a.周日, a.周一, a.周二, a.周三, a.周四, a.周五, a.周六, a.分诊方式, a.序号控制,
                             b.执行时间, a.项目id, a.医生姓名, a.医生id, Nvl(b.执行计划id, 0) As 执行计划id,
                             Row_Number() Over(Partition By a.安排id Order By a.生效时间 Desc) As 顺序号, b.序号, b.科室id
                      From 挂号安排计划 A, 挂号安排 B
                      Where Sysdate Between a.生效时间 + 0 And a.失效时间 And a.安排id = b.Id And
                            a.实际生效 >= To_Date('3000-01-01', 'yyyy-mm-dd') And 审核时间 Is Not Null And
                            b.停用日期 Is Null)
               Where 顺序号 = 1 And ID <> Nvl(执行计划id, 0)) Loop
    Update 挂号安排计划 Set 实际生效 = v_生效.上次生效时间 Where ID = v_生效.安排id And 失效时间 < v_生效.失效时间;
  
    Update 挂号安排
    Set 周日 = v_生效.周日, 周一 = v_生效.周一, 周二 = v_生效.周二, 周三 = v_生效.周三, 周四 = v_生效.周四, 周五 = v_生效.周五, 周六 = v_生效.周六,
        分诊方式 = v_生效.分诊方式, 序号控制 = v_生效.序号控制, 开始时间 = Sysdate, 终止时间 = v_生效.失效时间, 项目id = Nvl(v_生效.项目id, 项目id), 执行时间 = v_Date,
        执行计划id = v_生效.Id, 序号 = 9999999, 医生姓名 = v_生效.医生姓名, 医生id = v_生效.医生id
    Where ID = v_生效.安排id;
  
    --重新调整序号 
    Update 挂号安排 A
    Set 序号 = -1 * 序号
    Where 项目id = v_生效.项目id And a.科室id = v_生效.科室id And Nvl(a.医生姓名, '-') = Nvl(v_生效.医生姓名, '-') And
          Nvl(a.医生id, 0) = Nvl(v_生效.医生id, 0);
    For v_序号 In (Select a.Id, Rownum As 序号
                 From 挂号安排 A
                 Where a.项目id = v_生效.项目id And a.科室id = v_生效.科室id And Nvl(a.医生姓名, '-') = Nvl(v_生效.医生姓名, '-') And
                       Nvl(a.医生id, 0) = Nvl(v_生效.医生id, 0)
                 Order By a.Id) Loop
      Update 挂号安排 A Set 序号 = v_序号.序号 Where ID = v_序号.Id;
    End Loop;
    Delete 挂号安排诊室 Where 号表id = v_生效.安排id;
    Insert Into 挂号安排诊室
      (号表id, 门诊诊室)
      Select v_生效.安排id, 门诊诊室 From 挂号计划诊室 Where 计划id = v_生效.Id;
    Delete 挂号安排限制 Where 安排id = v_生效.安排id;
    Insert Into 挂号安排限制
      (安排id, 限制项目, 限号数, 限约数)
      Select v_生效.安排id, 限制项目, 限号数, 限约数 From 挂号计划限制 Where 计划id = v_生效.Id;
    Delete 挂号安排时段 Where 安排id = v_生效.安排id;
    Insert Into 挂号安排时段
      (安排id, 序号, 开始时间, 结束时间, 限制数量, 是否预约, 星期)
      Select v_生效.安排id, 序号, 开始时间, 结束时间, 限制数量, 是否预约, 星期
      From 挂号计划时段
      Where 计划id = v_生效.Id;
    If Nvl(v_Unitscount, 0) > 0 Then
      Delete 合作单位安排控制 Where 安排id = v_生效.安排id;
      Insert Into 合作单位安排控制
        (安排id, 合作单位, 限制项目, 序号, 数量)
        Select v_生效.安排id, 合作单位, 限制项目, 序号, 数量 From 合作单位计划控制 Where 计划id = v_生效.Id;
    End If;
  End Loop;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_挂号安排_Autoupdate;
/

--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
Create Or Replace Procedure Zl_分诊预约接收_取号
(
  No_In         门诊费用记录.No%Type,
  诊室_In       门诊费用记录.发药窗口%Type,
  病人id_In     门诊费用记录.病人id%Type,
  医生姓名_In   门诊费用记录.执行人 %Type := Null,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type,
  登记时间_In   门诊费用记录.登记时间%Type := Null,
  摘要_In       病人挂号记录.摘要%Type := Null,
  险类_In       病人挂号记录.险类%Type := Null
) As

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count    Number(18);
  n_排队     Number;
  n_当天排队 Number;
  n_生成队列 Number(3);
  d_Date     Date;

  v_排队号码 排队叫号队列.排队号码%Type;
  v_队列名称 排队叫号队列.队列名称%Type;
  v_排队序号 排队叫号队列.排队序号%Type;
  d_排队时间 排队叫号队列.排队时间%Type;

  d_预约时间   门诊费用记录.发生时间%Type;
  d_发生时间   门诊费用记录.发生时间%Type;
  n_出诊记录id 病人挂号记录.出诊记录id%Type;
  v_操作员姓名 病人挂号记录.接收人%Type;
  v_预约方式   病人挂号记录.预约方式 %Type;
  n_挂号id     病人挂号记录.Id%Type;
  n_组id       财务缴款分组.Id%Type;

  n_记录状态   病人挂号记录.记录状态%Type;
  n_号序       病人挂号记录.号序%Type;
  v_号码       病人挂号记录.号别%Type;
  n_预交id     病人预交记录.Id%Type;
  n_号源id     临床出诊号源.Id%Type;
  n_结帐id     病人结帐记录.Id%Type;
  n_安排id     挂号安排.Id%Type;
  n_计划id     挂号安排计划.Id%Type;
  n_是否分时段 临床出诊记录.是否分时段%Type := 0;

  n_接收模式 Number := 0;
  v_Paratemp Varchar2(100);
  d_启用时间 Date;
  n_挂号模式 Number(3);

Begin
  n_组id := Zl_Get组id(操作员姓名_In);

  n_生成队列 := To_Number(Nvl(Zl_Getsysparameter('排队叫号模式', 1113), 0));

  --0-预约接收立即就诊模式 1-预约接收不就诊模式 
  n_接收模式 := Nvl(Zl_Getsysparameter('预约接收模式', 1111), 0);
  v_Paratemp := Nvl(Zl_Getsysparameter('挂号排班模式'), 0);
  n_接收模式 := Nvl(Zl_Getsysparameter('预约接收模式', 1111), 0);

  If n_挂号模式 = 1 Then
    Begin
      d_启用时间 := To_Date(Substr(v_Paratemp, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_启用时间 := Null;
    End;
  End If;

  If 登记时间_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := 登记时间_In;
  End If;

  --更新挂号序号状态
  Begin
    Select 号别, 号序, Trunc(发生时间), 发生时间, 预约方式, 记录状态, 记录性质, 接收人, 出诊记录id
    Into v_号码, n_号序, d_预约时间, d_发生时间, v_预约方式, n_记录状态, n_Count, v_操作员姓名, n_出诊记录id
    From 病人挂号记录 A
    Where 记录状态 In (1, 3) And NO = No_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := -1;
  End;

  If n_Count = -1 Then
    v_Err_Msg := '预约挂号单:' || No_In || '不存在';
    Raise Err_Item;
  End If;

  If n_Count = 1 Then
    If n_记录状态 = 3 Then
      v_Err_Msg := '预约挂号单:' || No_In || '已经被退号';
      Raise Err_Item;
    End If;
    If v_操作员姓名 <> 操作员姓名_In Then
      v_Err_Msg := '预约挂号单:' || No_In || '已被接收';
      Raise Err_Item;
    Else
      v_Err_Msg := '预约挂号单:' || No_In || '已被他人接收';
      Raise Err_Item;
    End If;
  End If;

  If d_启用时间 Is Not Null Then
    If d_发生时间 < d_启用时间 Then
      v_Err_Msg := '当前预约挂号单属于出诊表排班模式安排，不能在' || To_Char(d_启用时间, 'yyyy-mm-dd hh24:mi:ss') || '之前接收!';
      Raise Err_Item;
    End If;
  End If;
  
  --表示签道了
  UPDATE 病人挂号记录 SET 记录标志=1 WHERE no=no_In;

  --判断是否分时段
  n_出诊记录id := Nvl(n_出诊记录id, 0);
  If n_出诊记录id = 0 Then
    n_Count := 0;
    Select Max(ID) Into n_安排id From 挂号安排 Where 号码 = v_号码;
  
    Select Max(ID)
    Into n_计划id
    From 挂号安排计划 A
    Where a.安排id = n_安排id And 审核时间 Is Not Null And
          a.生效时间 = (Select Max(生效时间)
                    From 挂号安排计划 A
                    Where 安排id = n_安排id And Sysdate Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                          失效时间 And 审核时间 Is Not Null);
    If Nvl(n_计划id, 0) = 0 Then
      Select Max(1) Into n_是否分时段 From 挂号安排时段 A Where a.安排id = n_安排id And Rownum < 2;
    Else
      Select Max(1) Into n_是否分时段 From 挂号计划时段 A Where a.计划id = n_计划id And Rownum < 2;
    End If;
  Else
    --判断是否分时段
    Select Nvl(是否分时段, 0), 号源id Into n_是否分时段, n_号源id From 临床出诊记录 Where ID = n_出诊记录id;
  End If;

  --分时段的号别，只能当天接收
  If n_是否分时段 = 1 Then
    If Trunc(d_发生时间) <> Trunc(Sysdate) Then
      v_Err_Msg := '分时段的预约挂号单只能取当天号！';
      Raise Err_Item;
    End If;
  End If;

  If n_是否分时段 = 0 Then
    --0-预约接收立即就诊模式 1-预约接收不就诊模式
    If n_接收模式 = 0 Then
      If Trunc(d_发生时间) <> Trunc(Sysdate) Then
        d_发生时间 := Sysdate;
      End If;
    End If;
  End If;

  If Not n_号序 Is Null Then
  
    If Trunc(d_预约时间) <> Trunc(Sysdate) And n_接收模式 = 0 Then
      If Nvl(n_出诊记录id, 0) = 0 Then
      
        --提前接收或延迟接收:先删除当天的预约时的序号
        Delete 挂号序号状态 Where 号码 = v_号码 And Trunc(日期) = Trunc(d_预约时间) And 序号 = n_号序;
      
        --锁当前时间的号
        Zl_挂号安排_传统_Lockno(2, v_号码, d_发生时间, Null, n_号序, Null, 操作员姓名_In, n_安排id, n_计划id, 0, 操作员姓名_In || '锁号', Null, Null);
        Update 挂号序号状态
        Set 状态 = 1, 登记时间 = Sysdate
        Where Trunc(日期) = Trunc(d_预约时间) And 序号 = n_号序 And 号码 = v_号码 And 状态 In (2, 5);
      Else
        --提前接收或延迟接收:先修改当天的预约时的序号
        Update 临床出诊序号控制 Set 挂号状态 = 0 Where 序号 = n_号序 And 记录id = n_出诊记录id;
        For c_旧安排 In (Select 是否分时段, 是否序号控制, 科室id, 医生id, 项目id, 上班时段
                      
                      From 临床出诊记录
                      Where ID = n_出诊记录id) Loop
        
          Begin
            n_Count := 1;
            Select ID
            Into n_出诊记录id
            From 临床出诊记录
            Where 号源id = n_号源id And 是否分时段 = c_旧安排.是否分时段 And 是否序号控制 = c_旧安排.是否序号控制 And 科室id = c_旧安排.科室id And
                  Nvl(医生id, 0) = Nvl(c_旧安排.医生id, 0) And 上班时段 = c_旧安排.上班时段 And Nvl(是否发布, 0) = 1 And
                  出诊日期 = Trunc(Sysdate) And Rownum < 2;
          Exception
            When Others Then
              n_Count := 0;
          End;
          If n_Count = 0 Then
            v_Err_Msg := '接收当天没有对应的出诊安排,无法接收!';
            Raise Err_Item;
          End If;
        
          Zl_挂号安排_临床出诊_Lockno(2, n_出诊记录id, d_发生时间, Null, n_号序, 0, 操作员姓名_In || '锁号', Null, 操作员姓名_In, Null, Null, Null,
                              v_号码);
        
          Update 临床出诊序号控制
          Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
          Where 记录id = n_出诊记录id And 序号 = n_号序 And Nvl(挂号状态, 0) In (0, 5);

          Update 病人挂号记录 Set 出诊记录id = n_出诊记录id Where NO = No_In;
        
        End Loop;
      
      End If;
    Else
      If Nvl(n_出诊记录id, 0) = 0 Then
      
        Update 挂号序号状态
        Set 序号 = n_号序, 状态 = 1, 登记时间 = Sysdate
        Where 号码 = v_号码 And Trunc(日期) = Trunc(d_预约时间) And 序号 = n_号序;
        If Sql%Rowcount = 0 Then
          Begin
            n_Count := 1;
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
            Values
              (v_号码, Trunc(d_发生时间), n_号序, 1, 操作员姓名_In, Sysdate);
          Exception
            When Others Then
              n_Count := 0;
          End;
          If n_Count = 0 Then
            v_Err_Msg := '序号' || n_号序 || '已被其它人使用,请重新选择一个序号.';
            Raise Err_Item;
          
          End If;
        End If;
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
        Where (序号 = n_号序 Or 备注 = n_号序) And 记录id = n_出诊记录id;
        If Sql%Rowcount = 0 Then
          v_Err_Msg := '序号' || n_号序 || '已被其它人使用,请重新选择一个序号.';
          Raise Err_Item;
        End If;
      End If;
    End If;
  End If;

  Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  
  --更新病人信息，含就诊信息
  Update 病人信息 Set 就诊时间 = d_Date, 就诊状态 = 2, 就诊诊室 = 诊室_In Where 病人id = 病人id_In;
  If Sql%NotFound Then
    v_Err_Msg := '传入的病人信息(病人ID)无效，请检查.';
    Raise Err_Item;
  End If;
  For c_病人 In (Select a.病人id, a.门诊号, a.费别, a.姓名, a.性别, a.年龄, a.医疗付款方式, b.编码 As 医疗付款方式编码
               From 病人信息 A, 医疗付款方式 B
               Where a.医疗付款方式 = b.名称(+) And a.病人id = 病人id_In) Loop

    --更新门诊费用记录
    Update 门诊费用记录
    Set 记录状态 = 1, 实际票号 = Null, 结帐id = n_结帐id, 结帐金额 = 0,实收金额=0, 病人id = c_病人.病人id, 标识号 = c_病人.门诊号, 姓名 = c_病人.姓名, 年龄 = c_病人.年龄,
        性别 = c_病人.性别, 付款方式 = c_病人.医疗付款方式编码, 费别 = c_病人.费别, 发生时间 = d_发生时间, 登记时间 = d_Date, 操作员编号 = 操作员编号_In,
        操作员姓名 = 操作员姓名_In, 缴款组id = n_组id, 发药窗口 = 诊室_In, 执行人 = 医生姓名_In
    Where 记录性质 = 4 And 记录状态 = 0 And NO = No_In;
    --病人挂号记录
    Update 病人挂号记录
    Set 接收人 = 操作员姓名_In, 接收时间 = d_Date, 记录性质 = 1, 病人id = c_病人.病人id, 门诊号 = c_病人.门诊号, 发生时间 = d_发生时间, 姓名 = c_病人.姓名,
        性别 = c_病人.性别, 年龄 = c_病人.年龄, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 险类 = Decode(Nvl(险类_In, 0), 0, Null, 险类_In),
        号序 = n_号序, 诊室 = 诊室_In, 执行人 = 医生姓名_In, 摘要 = Nvl(摘要_In, 摘要),记录标志=1,取号标志=1
    Where 记录状态 = 1 And NO = No_In And 记录性质 = 2;
    If Sql%NotFound Then
      v_Err_Msg := '由于并发原因,单据号为【' || No_In || '】的病人' || c_病人.姓名 || '已经被接收';
      Raise Err_Item;
    End If;
  
    --保存预交记录
    Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算序号,
       结算性质)
    Values
      (n_预交id, 4, 1, No_In, c_病人.病人id, '现金', 0, d_Date, 操作员编号_In, 操作员姓名_In, n_结帐id, '挂号收费', n_组id, Null, Null, Null,
       Null, Null, n_结帐id, 4);
  
    --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
    If Nvl(n_生成队列, 0) <> 0 Then
    
      For v_挂号 In (Select ID, 姓名, 诊室, 执行人, 执行部门id, 发生时间, 号别, 号序 From 病人挂号记录 Where NO = No_In) Loop
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
          Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, v_挂号.执行部门id, v_排队号码, Null, c_病人.姓名, 病人id_In, v_挂号.诊室, v_挂号.执行人, d_排队时间,
                           
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
      
      End Loop;
    End If;
  
  End Loop;

  --病人担保信息
  If 病人id_In Is Not Null Then
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Nvl(在院, 0) = 0 And Exists
     (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
    If Sql%Rowcount > 0 Then
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
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_分诊预约接收_取号;
/

--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
Create Or Replace Procedure Zl_挂号安排_传统_Lockno
(
  操作类型_In   Integer,
  号码_In       挂号安排.号码%Type,
  日期_In       挂号序号状态.日期%Type,
  号序_In       挂号序号状态.序号%Type,
  序号_Out      Out 挂号序号状态.序号%Type,
  机器名_In     挂号序号状态.机器名%Type := Null,
  操作员姓名_In 挂号序号状态.操作员姓名%Type := Null,
  安排id_In     挂号安排.Id%Type := Null,
  计划id_In     挂号安排计划.Id%Type := Null,
  是否预约_In   Number := 0,
  备注_In       挂号序号状态.备注%Type := Null,
  合作单位_In   Varchar2 := Null,
  时间段_In     Varchar2 := Null
  
) Is
  --功能:传统模式的锁号操作
  --操作类型_In： 0-解锁,1-加锁（根据传入的日期来加锁，没找到取一个）;2-加锁(直接取一下有效号进行加锁)
  --安排ID_In:如果不为空，则直接从安排中取数
  --计划ID_In:如果不为空，则直接从计划中取数
  --时间段_In:可以不传入，不传入时，则直接取下一个有效号 格式:HH24:mi:ss-HH24:mi:ss
  n_安排id       挂号安排.Id%Type;
  n_计划id       挂号安排计划.Id%Type;
  v_号码         挂号安排.号码%Type;
  n_状态         挂号序号状态.状态%Type;
  v_星期         挂号安排限制.限制项目%Type;
  n_号序         挂号序号状态.序号%Type;
  v_验证姓名     挂号序号状态.操作员姓名%Type;
  v_操作员姓名   挂号序号状态.操作员姓名%Type;
  v_验证机器名   挂号序号状态.机器名%Type;
  v_机器名       挂号序号状态.机器名%Type;
  n_限号数       挂号安排限制.限号数%Type;
  n_限约数       挂号安排限制.限约数%Type;
  n_合约模式     Number(3);
  n_序号控制     Number(3);
  n_分时段       Number(3);
  n_存在         Number(18);
  n_启用合作单位 Number(3);
  n_是否挂号     Number(3); --1-挂号;0-预约
  n_自锁号       Number(3);
  d_时段开始     Date;
  d_序号时间     Date;
  d_时段结束     Date;
  n_Rowid        Rowid;
  v_Temp         Varchar2(32767); --临时XML
  Err_Item Exception;

  Function Check_Nums_Valied
  (
    安排id1_In  In 挂号安排.Id%Type,
    计划id1_In  In 挂号安排计划.Id%Type,
    星期1_In    In 挂号安排限制.限制项目%Type,
    是否挂号_In Number
  ) Return Number Is
    --功能：检查是否超出了限号或限约
    --入参:是否挂号_IN-1:挂号;0-预约
    --返回:1-表示数据合法;0-表示数据不合法:超出了限号或限约数
    n_Count Number(18);
    n_Temp  Number(18);
  Begin
    If Nvl(n_计划id, 0) <> 0 Then
      Select Max(限号数), Max(限约数)
      Into n_限号数, n_限约数
      From 挂号计划限制
      Where 计划id = 计划id1_In And 限制项目 = 星期1_In;
    Else
      Select Max(限号数), Max(限约数)
      Into n_限号数, n_限约数
      From 挂号安排限制
      Where 安排id = 安排id1_In And 限制项目 = 星期1_In;
    End If;
  
    Select Count(*)
    Into n_Count
    From (Select 序号
           From 挂号序号状态
           Where 号码 = 号码_In And Trunc(日期) = Trunc(日期_In)
           Union
           Select 序号
           From 合作单位计划控制
           Where 计划id = Decode(是否挂号_In, 1, 0, 计划id1_In) And Decode(是否挂号_In, 1, 0, 0) = 0 And 限制项目 = 星期1_In And 数量 <> 0
           Union
           Select 序号
           From 合作单位安排控制
           Where 安排id = Decode(是否挂号_In, 1, 0, 安排id1_In) And Decode(是否挂号_In, 1, 0, 0) = 0 And 限制项目 = 星期1_In And 数量 <> 0);
  
    If 是否挂号_In = 1 And Nvl(n_限号数, 0) <> 0 And Nvl(n_限号数, 0) < n_Count Then
      Return 0;
    Elsif 是否挂号_In = 0 Then
      n_Temp := Nvl(n_限约数, 0);
      If n_Temp = 0 Then
        n_Temp := Nvl(n_限号数, 0);
      End If;
      If n_Temp <> 0 And n_Temp < n_Count Then
        Return 0;
      End If;
    End If;
    Return 1;
  End;

  Function Get_Next_Plannum
  (
    号码1_In       In 挂号安排.号码%Type,
    日期1_In       In Date,
    安排id1_In     In 挂号安排.Id%Type,
    计划id1_In     In 挂号安排计划.Id%Type,
    星期1_In       In 挂号安排限制.限制项目%Type,
    操作员姓名1_In 人员表.姓名%Type,
    机器名1_In     挂号序号状态.机器名%Type,
    备注1_In       In 挂号序号状态.备注%Type
  ) Return Number Is
    n_Temp_序号 Number(18);
    n_Find      Number(2);
    n_自锁号    Number(2);
    d_序号时间  Date;
    n_Rowid     Rowid;
  Begin
    If Nvl(计划id_In, 0) <> 0 Then
      Select Max(序号) + 1
      Into n_Temp_序号
      From (Select Distinct 序号
             From 挂号计划时段
             Where 计划id = 计划id1_In And 星期 = 星期1_In
             Union All
             Select Distinct 序号
             From 挂号序号状态
             Where 号码 = 号码1_In And Trunc(日期) = Trunc(日期1_In));
    Else
      Select Max(序号) + 1
      Into n_Temp_序号
      From (Select Distinct 序号
             From 挂号安排时段
             Where 安排id = 安排id1_In And 星期 = 星期1_In
             Union
             Select Distinct 序号
             From 挂号序号状态
             Where 号码 = 号码1_In And Trunc(日期) = Trunc(日期1_In));
    End If;
  
    n_Find := 0;
    While n_Find = 0 Loop
      Begin
        Select Rowid, 1,
               Case
                 When 机器名 = 机器名1_In And 操作员姓名 = 操作员姓名1_In And 状态 = 5 Then
                  1
                 Else
                  0
               End
        Into n_Rowid, n_存在, n_自锁号
        From 挂号序号状态
        Where 号码 = 号码1_In And 日期 Between Trunc(日期1_In) And Trunc(日期1_In) + 1 - 1 / 24 / 60 / 60 And 序号 = n_Temp_序号;
      Exception
        When Others Then
          n_存在   := 0;
          n_自锁号 := 0;
      End;
      If Nvl(n_存在, 0) = 1 And Nvl(n_自锁号, 0) = 1 Then
        --自己锁的号，独站起:
        Update 挂号序号状态 Set 状态 = 5 Where Rowid = n_Rowid;
        n_Find := 1;
        Return n_Temp_序号;
      End If;
    
      If Nvl(n_存在, 0) = 0 Then
        --未发现该序号被站用，插入记录
        d_序号时间 := 日期1_In;
        If 时间段_In Is Not Null Then
          Begin
            If Nvl(计划id1_In, 0) <> 0 Then
              Select To_Date(To_Char(日期1_In, 'yyyy-mm-dd') || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
              Into d_序号时间
              From 挂号计划时段
              Where 计划id = 计划id1_In And 星期 = v_星期 And 序号 = n_号序;
            Else
            
              Select To_Date(To_Char(日期1_In, 'yyyy-mm-dd') || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
              Into d_序号时间
              From 挂号安排时段
              Where 安排id = 安排id1_In And 星期 = v_星期 And 序号 = n_号序;
            End If;
          Exception
            When Others Then
              d_序号时间 := 日期1_In;
          End;
        End If;
        Insert Into 挂号序号状态
          (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
        Values
          (号码1_In, d_序号时间, n_Temp_序号, 5, 操作员姓名1_In, 备注1_In, Sysdate, 机器名1_In);
      
        n_Find := 1;
        Return n_Temp_序号;
      End If;
      n_Temp_序号 := n_Temp_序号 + 1;
    End Loop;
  End;

Begin

  v_机器名 := 机器名_In;
  If v_机器名 Is Null Then
    Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  End If;
  v_操作员姓名 := 操作员姓名_In;
  If v_操作员姓名 Is Null Then
    v_操作员姓名 := Zl_Username;
  End If;

  n_号序     := 号序_In;
  v_号码     := 号码_In;
  n_是否挂号 := Case
              When Nvl(是否预约_In, 0) = 0 Then
               1
              Else
               0
            End;

  If 操作类型_In = 0 Then
    --解锁
    Delete 挂号序号状态
    Where 机器名 = v_机器名 And 操作员姓名 = v_操作员姓名 And 状态 = 5 And 序号 = n_号序 And Trunc(日期) = Trunc(日期_In) And 号码 = 号码_In;
    If Sql%NotFound Then
      v_Temp := '没有发现需要解锁的序号';
      Raise Err_Item;
    End If;
    序号_Out := n_号序;
    Return;
  End If;

  --锁号
  If 时间段_In Is Not Null Then
    Begin
      d_时段开始 := To_Date(To_Char(日期_In, 'yyyy-mm-dd') || Substr(时间段_In, 1, Instr(时间段_In, '-') - 1),
                        'yyyy-mm-dd hh24:mi:ss');
      If Substr(时间段_In, Instr(时间段_In, '-') + 1) Is Null Then
        d_时段结束 := Null;
      Else
        d_时段结束 := To_Date(To_Char(日期_In, 'yyyy-mm-dd') || Substr(时间段_In, Instr(时间段_In, '-') + 1),
                          'yyyy-mm-dd hh24:mi:ss');
      End If;
    Exception
      When Others Then
        v_Temp := '无法解析传入的时间段格式，请检查！';
        Raise Err_Item;
    End;
  End If;

  Select Decode(To_Char(日期_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)
  Into v_星期
  From Dual;

  n_计划id := Nvl(计划id_In, 0);
  n_安排id := Nvl(安排id_In, 0);
  If Nvl(n_计划id, 0) <> 0 Then
    Select Max(a.序号控制), Max(b.号码)
    Into n_序号控制, v_号码
    From 挂号安排计划 A, 挂号安排 B
    Where a.Id = n_计划id And a.安排id = b.Id;
  End If;
  If Nvl(n_安排id, 0) <> 0 Then
    Select Max(序号控制), Max(号码) Into n_序号控制, v_号码 From 挂号安排 Where ID = n_安排id;
  End If;

  If Nvl(n_计划id, 0) = 0 And Nvl(n_安排id, 0) = 0 Then
    Begin
      Select 序号控制, ID
      Into n_序号控制, n_计划id
      From (Select 序号控制, ID
             From 挂号安排计划
             Where 号码 = v_号码 And 日期_In Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                   失效时间 And 审核时间 Is Not Null
             Order By 生效时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Select 序号控制, ID Into n_序号控制, n_安排id From 挂号安排 Where 号码 = v_号码;
    End;
  End If;

  If Nvl(n_序号控制, 0) = 0 Then
    --未启用序号，不能锁号
    Return;
  End If;

  If Nvl(n_计划id, 0) <> 0 Then
    Select Nvl(Max(1), 0) Into n_分时段 From 挂号计划时段 Where 星期 = v_星期 And 计划id = n_计划id And Rownum < 2;
  
    Select Nvl(Max(1), 0)
    Into n_启用合作单位
    From 合作单位计划控制
    Where 限制项目 = v_星期 And 计划id = n_计划id And 合作单位 = 合作单位_In And Rownum < 2;
  Else
  
    Select Nvl(Max(1), 0) Into n_分时段 From 挂号安排时段 Where 星期 = v_星期 And 安排id = n_安排id And Rownum < 2;
    Select Nvl(Max(1), 0)
    Into n_启用合作单位
    From 合作单位安排控制
    Where 限制项目 = v_星期 And 安排id = n_安排id And 合作单位 = 合作单位_In And Rownum < 2;
  End If;

  If 操作类型_In = 2 Then
    --直接取一下号来进行锁号操作
    v_Temp := Zl_Fun_挂号安排_传统_Nextsn(日期_In, n_安排id, n_计划id, v_操作员姓名, v_星期, 备注_In, v_机器名, 合作单位_In, 0, Nvl(是否预约_In, 0));
    If v_Temp Is Not Null Then
      序号_Out := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    End If;
    Return;
  End If;

  n_存在 := 0;
  If 时间段_In Is Null And 操作类型_In = 1 Then
    If Nvl(n_计划id, 0) <> 0 Then
      Begin
        Select 1, a.状态, a.操作员姓名, a.机器名, a.序号
        Into n_存在, n_状态, v_验证姓名, v_验证机器名, n_号序
        From 挂号序号状态 A, 挂号计划时段 B
        Where a.号码 = v_号码 And Trunc(a.日期) = Trunc(日期_In) And a.序号 = b.序号 And b.计划id = n_计划id And b.星期 = v_星期 And
              To_Char(b.开始时间, 'hh24:mi') = To_Char(日期_In, 'hh24:mi') And Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    Else
      Begin
        Select 1, a.状态, a.操作员姓名, a.机器名, a.序号
        Into n_存在, n_状态, v_验证姓名, v_验证机器名, n_号序
        From 挂号序号状态 A, 挂号安排时段 B
        Where a.号码 = v_号码 And Trunc(a.日期) = Trunc(日期_In) And a.序号 = b.序号 And b.安排id = n_安排id And b.星期 = v_星期 And
              To_Char(b.开始时间, 'hh24:mi') = To_Char(日期_In, 'hh24:mi') And Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    End If;
  End If;

  If n_存在 = 1 Then
    If Not (n_状态 = 5 And v_验证姓名 = v_操作员姓名 And v_机器名 = v_验证机器名) Then
      --传入时间的序号已经被使用
      v_Temp := '传入时间' || 日期_In || '的序号已被使用';
      Raise Err_Item;
    End If;
    序号_Out := n_号序;
    Return;
  End If;

  If n_分时段 = 1 And 操作类型_In = 1 Then
    If 时间段_In Is Null Then
      --精确定位序号
      Begin
        n_存在 := 1;
        If Nvl(n_计划id, 0) <> 0 Then
          Select 序号
          Into n_号序
          From 挂号计划时段
          Where 计划id = n_计划id And 星期 = v_星期 And To_Char(开始时间, 'hh24:mi') = To_Char(日期_In, 'hh24:mi') And Rownum < 2;
        Else
          Select 序号
          Into n_号序
          From 挂号安排时段
          Where 安排id = n_安排id And 星期 = v_星期 And To_Char(开始时间, 'hh24:mi') = To_Char(日期_In, 'hh24:mi') And Rownum < 2;
        End If;
      Exception
        When Others Then
          n_存在 := 0;
      End;
    
      If n_存在 = 1 Then
        --存在，则检查是否被其他人站用。
        Begin
          Select Rowid, 1,
                 Case
                   When 机器名 = v_机器名 And 操作员姓名 = v_操作员姓名 And 状态 = 5 Then
                    1
                   Else
                    0
                 End
          Into n_Rowid, n_存在, n_自锁号
          From 挂号序号状态
          Where 号码 = v_号码 And 日期 = 日期_In And 序号 = n_号序;
        Exception
          When Others Then
            n_存在   := 0;
            n_自锁号 := 0;
        End;
      
        If Nvl(n_存在, 0) = 1 And Nvl(n_自锁号, 0) = 1 Then
          --自己锁的号，独站起:
          Update 挂号序号状态 Set 状态 = 5 Where Rowid = n_Rowid;
          序号_Out := n_号序;
          Return;
        End If;
        If Nvl(n_存在, 0) = 1 And Nvl(n_自锁号, 0) = 0 Then
          v_Temp := '传入时间' || 日期_In || '的序号已被使用';
          Raise Err_Item;
        End If;
        Insert Into 挂号序号状态
          (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
        Values
          (v_号码, 日期_In, n_号序, 5, v_操作员姓名, 备注_In, Sysdate, v_机器名);
        序号_Out := n_号序;
        Return;
      End If;
      --不存在时，取下一个号,同时检查限号数是否正确
      n_号序 := Get_Next_Plannum(v_号码, 日期_In, n_安排id, n_计划id, v_星期, v_操作员姓名, v_机器名, 备注_In);
    
      If Check_Nums_Valied(n_安排id, n_计划id, v_星期, n_是否挂号) = 0 Then
        v_Temp := '传入号别' || 号码_In || '当前已无余号';
        Raise Err_Item;
      End If;
      序号_Out := n_号序;
      Return;
    Else
      If Nvl(n_计划id, 0) <> 0 Then
        --预约判断: 如果全部序号可以预约，则是否预约全部为0，所以需要单独处理这种情况
      
        Select Nvl(Max(1), 0)
        Into n_存在
        From 挂号计划时段
        Where 计划id = n_计划id And 星期 = v_星期 And 是否预约 = 1 And Rownum < 2;
      
        Select Min(a.序号), Min(b.状态)
        Into n_号序, n_状态
        From 挂号计划时段 A, 挂号序号状态 B
        Where a.序号 = b.序号(+) And (Nvl(b.状态, 0) = 0 Or Nvl(b.状态, 0) = 5) And b.号码(+) = v_号码 And
              Trunc(b.日期(+)) = Trunc(日期_In) And a.计划id = n_计划id And a.星期 = v_星期 And
              To_Char(a.开始时间, 'hh24:mi') >= To_Char(d_时段开始, 'hh24:mi') And
              To_Char(a.开始时间, 'hh24:mi') < To_Char(d_时段结束, 'hh24:mi') And Case
                When n_是否挂号 = 1 Or n_存在 = 0 Then
                 1
                Else
                 a.是否预约
              End = 1; --  Decode(n_是否挂号, 1, 1, Decode(n_存在, 1, a.是否预约, 1))) = 1;
      Else
        Select Nvl(Max(1), 0)
        Into n_存在
        From 挂号安排时段
        Where 安排id = n_安排id And 星期 = v_星期 And 是否预约 = 1 And Rownum < 2;
      
        Select Min(a.序号), Min(b.状态)
        Into n_号序, n_状态
        From 挂号安排时段 A, 挂号序号状态 B
        Where a.序号 = b.序号(+) And (Nvl(b.状态, 0) = 0 Or Nvl(b.状态, 0) = 5) And b.号码(+) = v_号码 And
              Trunc(b.日期(+)) = Trunc(日期_In) And a.安排id = n_安排id And a.星期 = v_星期 And
              To_Char(a.开始时间, 'hh24:mi') >= To_Char(d_时段开始, 'hh24:mi') And
              To_Char(a.开始时间, 'hh24:mi') < To_Char(d_时段结束, 'hh24:mi') And Case
                When n_是否挂号 = 1 Or n_存在 = 0 Then
                 1
                Else
                 a.是否预约
              End = 1; --  Decode(n_是否挂号, 1, 1, Decode(n_存在, 1, a.
      
      End If;
    
      If Nvl(n_号序, 0) = 0 Then
        If n_存在 = 1 Then
          v_Temp := '传入时间段' || 时间段_In || '的序号已被使用或未开放预约。';
          Raise Err_Item;
        End If;
      
        If d_时段结束 Is Not Null Then
          v_Temp := '传入时间段' || 时间段_In || '的序号已被使用';
          Raise Err_Item;
        End If;
        --不存在时，取下一个号,同时检查限号数是否正确
        n_号序 := Get_Next_Plannum(v_号码, 日期_In, n_安排id, n_计划id, v_星期, v_操作员姓名, v_机器名, 备注_In);
      
        If Check_Nums_Valied(n_安排id, n_计划id, v_星期, n_是否挂号) = 0 Then
          v_Temp := '传入号别' || v_号码 || '当前已无余号';
          Raise Err_Item;
        End If;
        序号_Out := n_号序;
        Return;
      End If;
      --存在序号
      If Nvl(n_状态, 0) = 0 Then
        --合法时间段，插入记录
        Insert Into 挂号序号状态
          (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
        Values
          (v_号码, 日期_In, n_号序, 5, v_操作员姓名, 备注_In, Sysdate, v_机器名);
        序号_Out := n_号序;
        Return;
      
      End If;
    
      If Nvl(n_状态, 0) = 5 Then
        Select Nvl(Max(1), 0)
        Into n_存在
        From 挂号序号状态
        Where 机器名 = v_机器名 And 操作员姓名 = v_操作员姓名 And 序号 = n_号序 And 号码 = v_号码;
      
        If n_存在 = 0 Then
          v_Temp := '传入时间段' || 时间段_In || '的序号已被使用';
          Raise Err_Item;
        End If;
        序号_Out := n_号序;
        Return;
      End If;
    
      If d_时段结束 Is Not Null Then
        v_Temp := '传入时间段' || 时间段_In || '的序号已被使用';
        Raise Err_Item;
      End If;
      --不存在时，取下一个号,同时检查限号数是否正确
      n_号序 := Get_Next_Plannum(v_号码, 日期_In, n_安排id, n_计划id, v_星期, v_操作员姓名, v_机器名, 备注_In);
      If Check_Nums_Valied(n_安排id, n_计划id, v_星期, n_是否挂号) = 0 Then
        v_Temp := '传入号别' || 号码_In || '当前已无余号';
        Raise Err_Item;
      End If;
    End If;
    序号_Out := n_号序;
    Return;
  End If;

  --不分时段,但启用了序号的
  If Nvl(n_计划id, 0) <> 0 Then
    Select Decode(n_是否挂号, 1, Max(限号数), Max(限约数))
    Into n_限号数
    From 挂号计划限制
    Where 计划id = n_计划id And 限制项目 = v_星期;
  Else
    Select Decode(n_是否挂号, 1, Max(限号数), Max(限约数))
    Into n_限号数
    From 挂号安排限制
    Where 安排id = n_安排id And 限制项目 = v_星期;
  End If;

  n_号序 := 1;
  If 合作单位_In Is Null Or n_启用合作单位 = 0 Then
    For r_序号 In (Select 序号, 状态, 操作员姓名, 机器名
                 From 挂号序号状态
                 Where 号码 = 号码_In And Trunc(日期) = Trunc(日期_In)
                 Union
                 Select 序号, Null, Null, Null
                 From 合作单位计划控制
                 Where 计划id = Decode(n_是否挂号, 1, 0, n_计划id) And Decode(n_是否挂号, 1, 1, 0) = 0 And 限制项目 = v_星期 And 数量 <> 0
                 Union
                 Select 序号, Null, Null, Null
                 From 合作单位安排控制
                 Where 安排id = Decode(n_是否挂号, 1, 0, n_安排id) And Decode(n_是否挂号, 1, 1, 0) = 0 And 限制项目 = v_星期 And 数量 <> 0
                 Order By 序号) Loop
      --存在锁号的，则退出
      Exit When r_序号.状态 = 5 And r_序号.操作员姓名 = v_操作员姓名 And r_序号.机器名 = v_机器名;
      If r_序号.序号 = n_号序 Then
        n_号序 := n_号序 + 1;
      End If;
    End Loop;
  
    If n_号序 > n_限号数 Then
      v_Temp := '传入号别' || 号码_In || '当前已无余号';
      Raise Err_Item;
    End If;
  
    Begin
      Select Rowid, 1,
             Case
               When 机器名 = v_机器名 And 操作员姓名 = v_操作员姓名 And 状态 = 5 Then
                1
               Else
                0
             End
      Into n_Rowid, n_存在, n_自锁号
      From 挂号序号状态
      Where 号码 = 号码_In And Trunc(日期) = Trunc(日期_In) And 序号 = n_号序;
    Exception
      When Others Then
        n_存在   := 0;
        n_自锁号 := 0;
    End;
    If n_存在 = 1 And n_自锁号 = 1 Then
      序号_Out := n_号序;
      Return;
    End If;
    If n_存在 = 1 And n_自锁号 = 0 Then
      --已经站用了
      v_Temp := '序号' || n_号序 || '已被使用';
      Raise Err_Item;
    End If;
    Insert Into 挂号序号状态
      (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
    Values
      (号码_In, Trunc(日期_In), n_号序, 5, v_操作员姓名, 备注_In, Sysdate, v_机器名);
    序号_Out := n_号序;
    Return;
  End If;

  --启用了合作单位控制的:合约单位处理
  If Nvl(n_计划id, 0) <> 0 Then
    Select Count(1)
    Into n_合约模式
    From 合作单位计划控制
    Where 序号 = 0 And 计划id = n_计划id And 合作单位 = 合作单位_In And 限制项目 = v_星期 And 数量 <> 0;
  Else
    Select Count(1)
    Into n_合约模式
    From 合作单位安排控制
    Where 序号 = 0 And 安排id = n_安排id And 合作单位 = 合作单位_In And 限制项目 = v_星期 And 数量 <> 0;
  End If;

  If n_合约模式 = 0 Then
    If Nvl(n_计划id, 0) <> 0 Then
      Select Nvl(Max(序号), 0)
      Into n_号序
      From (Select 序号
             From 合作单位计划控制 A
             Where 计划id = n_计划id And 合作单位 = 合作单位_In And 限制项目 = v_星期 And 数量 <> 0 And
                   (Not Exists
                    (Select 1
                     From 挂号序号状态
                     Where 号码 = 号码_In And 序号 = a.序号 And Trunc(日期) = Trunc(日期_In) And 状态 <> 0) Or Exists
                    (Select 1
                     From 挂号序号状态
                     Where 号码 = 号码_In And 序号 = a.序号 And Trunc(日期) = Trunc(日期_In) And 状态 = 5 And 操作员姓名 = v_操作员姓名 And
                           机器名 = v_机器名))
             Order By 序号)
      Where Rownum < 2;
    Else
    
      Select Nvl(Max(序号), 0)
      Into n_号序
      From (Select 序号
             From 合作单位安排控制 A
             Where 安排id = n_安排id And 合作单位 = 合作单位_In And 限制项目 = v_星期 And 数量 <> 0 And
                   (Not Exists
                    (Select 1
                     From 挂号序号状态
                     Where 号码 = 号码_In And 序号 = a.序号 And Trunc(日期) = Trunc(日期_In) And 状态 <> 0) Or Exists
                    (Select 1
                     From 挂号序号状态
                     Where 号码 = 号码_In And 序号 = a.序号 And Trunc(日期) = Trunc(日期_In) And 状态 = 5 And 操作员姓名 = v_操作员姓名 And
                           机器名 = v_机器名))
             Order By 序号)
      Where Rownum < 2;
    End If;
  
    If Nvl(n_号序, 0) = 0 Then
      v_Temp := '传入号别' || 号码_In || '当前已无余号';
      Raise Err_Item;
    End If;
    Begin
      Select Rowid, 1,
             Case
               When 机器名 = v_机器名 And 操作员姓名 = v_操作员姓名 And 状态 = 5 Then
                1
               Else
                0
             End
      Into n_Rowid, n_存在, n_自锁号
      From 挂号序号状态
      Where 号码 = 号码_In And 日期 Between Trunc(日期_In) And Trunc(日期_In) + 1 - 1 / 24 / 60 / 60 And 序号 = n_号序;
    Exception
      When Others Then
        n_存在   := 0;
        n_自锁号 := 0;
    End;
    If n_存在 = 1 And n_自锁号 = 0 Then
      v_Temp := '序号为' || n_号序 || '已被使用';
      Raise Err_Item;
    End If;
    If Nvl(n_存在, 0) = 0 Then
      Insert Into 挂号序号状态
        (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
      Values
        (号码_In, Trunc(日期_In), n_号序, 5, v_操作员姓名, 备注_In, Sysdate, v_机器名);
    End If;
    序号_Out := n_号序;
    Return;
  
  Else
    n_号序 := 1;
    Select 限号数 Into n_限号数 From 挂号计划限制 Where 计划id = n_计划id And 限制项目 = v_星期;
    For r_序号 In (Select 序号, 状态, 操作员姓名, 机器名
                 From 挂号序号状态
                 Where 号码 = 号码_In And Trunc(日期) = Trunc(日期_In)
                 Union All
                 Select 序号, Null, Null, Null
                 From 合作单位计划控制
                 Where 计划id = n_计划id And Decode(Nvl(n_计划id, 0), 0, 0, 1) = 1 And 限制项目 = v_星期 And 数量 <> 0
                 Union All
                 Select 序号, Null, Null, Null
                 From 合作单位安排控制
                 Where 安排id = n_安排id And Decode(Nvl(n_计划id, 0), 0, 1, 0) = 1 And 限制项目 = v_星期 And 数量 <> 0
                 Order By 序号) Loop
      If r_序号.状态 = 5 And r_序号.操作员姓名 = v_操作员姓名 And r_序号.机器名 = v_机器名 Then
        n_号序 := r_序号.序号;
        Exit;
      End If;
      If r_序号.序号 = n_号序 Then
        n_号序 := n_号序 + 1;
      End If;
    End Loop;
  
    If n_号序 > n_限号数 Then
      v_Temp := '传入号别' || 号码_In || '当前已无余号';
      Raise Err_Item;
    End If;
  
    Begin
      Select Rowid, 1,
             Case
               When 机器名 = v_机器名 And 操作员姓名 = v_操作员姓名 And 状态 = 5 Then
                1
               Else
                0
             End
      Into n_Rowid, n_存在, n_自锁号
      From 挂号序号状态
      Where 号码 = 号码_In And 日期 Between Trunc(日期_In) And Trunc(日期_In) + 1 - 1 / 24 / 60 / 60 And 序号 = n_号序;
    Exception
      When Others Then
        n_存在   := 0;
        n_自锁号 := 0;
    End;
    If n_存在 = 1 And n_自锁号 = 0 Then
      v_Temp := '序号为' || n_号序 || '已被使用';
      Raise Err_Item;
    End If;
  
    If Nvl(n_存在, 0) = 0 Then
      Insert Into 挂号序号状态
        (号码, 日期, 序号, 状态, 操作员姓名, 备注, 登记时间, 机器名)
      Values
        (号码_In, Trunc(日期_In), n_号序, 5, v_操作员姓名, 备注_In, Sysdate, v_机器名);
    End If;
    序号_Out := n_号序;
    Return;
  End If;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Temp || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_挂号安排_传统_Lockno;
/

--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
Create Or Replace Procedure Zl_出诊表挂号_Turn(启用日期_In In Date) Is
  ------------------------------------------------ 
  --功能：将出诊表排班模式启用时间之后的计划排班挂号记录转换为出诊表排班模式挂号记录
  --返回：转换记录条数
  ------------------------------------------------ 
  v_Error Varchar2(255);
  Err_Custom Exception;
  n_处理数量   Number(10);
  v_Para       Varchar2(500);
  n_挂号模式   Number(3);
  d_启用时间   Date;
  n_安排id     挂号安排.Id%Type;
  n_计划id     挂号安排计划.Id%Type;
  v_时间段     时间段.时间段%Type;
  n_分时段     Number(3);
  n_分时段序号 临床出诊序号控制.序号%Type;
  n_序号控制   挂号安排.序号控制%Type;
  n_出诊记录id 临床出诊记录.Id%Type;
  n_未处理数量 Number(10);
Begin
  Begin
    d_启用时间 := 启用日期_In;
  Exception
    When Others Then
      d_启用时间 := Null;
  End;

  For r_挂号 In (Select ID, NO, 号别, 执行部门id, 执行人, 记录性质, 预约, 号序, 病人id, 发生时间, 操作员姓名
               From 病人挂号记录
               Where 记录状态 = 1 And 发生时间 >= Trunc(d_启用时间) And 出诊记录id Is Null) Loop
    v_时间段 := Null;
    Begin
      Select ID, 序号控制
      Into n_计划id, n_序号控制
      From (Select a.Id, a.序号控制
             From 挂号安排计划 A, 挂号安排 B
             Where a.安排id = b.Id And a.审核时间 Is Not Null And b.号码 = r_挂号.号别 And r_挂号.发生时间 Between 生效时间 + 0 And 失效时间
             Order By 生效时间 Desc)
      Where Rownum < 2;
      Select Decode(To_Char(r_挂号.发生时间, 'D'), '1', a.周日, '2', a.周一, '3', a.周二, '4', a.周三, '5', a.周四, '6', a.周五, '7', a.周六,
                     Null)
      Into v_时间段
      From 挂号安排计划 A
      Where a.Id = n_计划id;
    Exception
      When Others Then
        n_计划id := Null;
        Select ID, 序号控制 Into n_安排id, n_序号控制 From 挂号安排 Where 号码 = r_挂号.号别;
        Select Decode(To_Char(r_挂号.发生时间, 'D'), '1', a.周日, '2', a.周一, '3', a.周二, '4', a.周三, '5', a.周四, '6', a.周五, '7',
                       a.周六, Null)
        Into v_时间段
        From 挂号安排 A
        Where a.Id = n_安排id;
    End;
    If v_时间段 Is Not Null Then
      Begin
        If Nvl(n_计划id, 0) = 0 Then
          Select 1 Into n_分时段 From 挂号安排时段 Where 安排id = n_安排id And Rownum < 2;
        Else
          Select 1 Into n_分时段 From 挂号计划时段 Where 计划id = n_计划id And Rownum < 2;
        End If;
      Exception
        When Others Then
          n_分时段 := 0;
      End;
      Begin
        Select a.Id
        Into n_出诊记录id
        From 临床出诊记录 A, 临床出诊号源 B
        Where a.号源id = b.Id And b.号码 = r_挂号.号别 And r_挂号.发生时间 Between a.开始时间 And a.终止时间 And 上班时段 = v_时间段;
      Exception
        When Others Then
          n_出诊记录id := Null;
      End;
      If n_出诊记录id Is Null Then
        v_Error := '已经挂号的记录存在无法对应的出诊记录,立即启用失败!';
        Raise Err_Custom;
      End If;
      If Nvl(n_序号控制, 0) = 0 Then
        If n_分时段 = 0 Then
          If r_挂号.记录性质 = 1 Then
            If Nvl(r_挂号.预约, 0) = 1 Then
              Update 病人挂号记录 Set 出诊记录id = n_出诊记录id Where ID = r_挂号.Id;
              Update 临床出诊记录
              Set 已挂数 = 已挂数 + 1, 已约数 = 已约数 + 1, 其中已接收 = 其中已接收 + 1
              Where ID = n_出诊记录id;
            Else
              Update 病人挂号记录 Set 出诊记录id = n_出诊记录id Where ID = r_挂号.Id;
              Update 临床出诊记录 Set 已挂数 = 已挂数 + 1 Where ID = n_出诊记录id;
            End If;
          Else
            Update 病人挂号记录 Set 出诊记录id = n_出诊记录id Where ID = r_挂号.Id;
            Update 临床出诊记录 Set 已约数 = 已约数 + 1 Where ID = n_出诊记录id;
          End If;
          n_处理数量 := n_处理数量 + 1;
        Else
          --非序号控制分时段,特殊处理
          Select 序号 Into n_分时段序号 From 临床出诊序号控制 Where 预约顺序号 Is Null And 开始时间 = r_挂号.发生时间;
          If r_挂号.记录性质 = 1 Then
            If Nvl(r_挂号.预约, 0) = 1 Then
              Update 病人挂号记录 Set 出诊记录id = n_出诊记录id Where ID = r_挂号.Id;
              Update 临床出诊记录
              Set 已挂数 = 已挂数 + 1, 已约数 = 已约数 + 1, 其中已接收 = 其中已接收 + 1
              Where ID = n_出诊记录id;
            Else
              Update 病人挂号记录 Set 出诊记录id = n_出诊记录id Where ID = r_挂号.Id;
              Update 临床出诊记录 Set 已挂数 = 已挂数 + 1 Where ID = n_出诊记录id;
            End If;
            Insert Into 临床出诊序号控制
              (记录id, 序号, 预约顺序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 操作员姓名, 备注)
              Select 记录id, n_分时段序号, r_挂号.号序, 开始时间, 终止时间, 1, 是否预约, 1, r_挂号.操作员姓名, r_挂号.号序
              From 临床出诊序号控制
              Where 记录id = n_出诊记录id And 序号 = n_分时段序号 And 预约顺序号 Is Null;
          Else
            Update 病人挂号记录 Set 出诊记录id = n_出诊记录id Where ID = r_挂号.Id;
            Update 临床出诊记录 Set 已约数 = 已约数 + 1 Where ID = n_出诊记录id;
            Insert Into 临床出诊序号控制
              (记录id, 序号, 预约顺序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 操作员姓名, 备注)
              Select 记录id, n_分时段序号, r_挂号.号序, 开始时间, 终止时间, 1, 是否预约, 2, r_挂号.操作员姓名, r_挂号.号序
              From 临床出诊序号控制
              Where 记录id = n_出诊记录id And 序号 = n_分时段序号 And 预约顺序号 Is Null;
          End If;
          n_处理数量 := n_处理数量 + 1;
        End If;
      Else
        If r_挂号.记录性质 = 1 Then
          If Nvl(r_挂号.预约, 0) = 1 Then
            Update 病人挂号记录 Set 出诊记录id = n_出诊记录id Where ID = r_挂号.Id;
            Update 临床出诊记录
            Set 已挂数 = 已挂数 + 1, 已约数 = 已约数 + 1, 其中已接收 = 其中已接收 + 1
            Where ID = n_出诊记录id;
            Update 临床出诊序号控制
            Set 挂号状态 = 1, 操作员姓名 = r_挂号.操作员姓名
            Where 记录id = n_出诊记录id And 序号 = r_挂号.号序;
            If Sql%RowCount = 0 Then
              Insert Into 临床出诊序号控制
                (记录id, 序号, 开始时间, 终止时间, 数量, 挂号状态, 操作员姓名, 备注)
              Values
                (n_出诊记录id, r_挂号.号序, r_挂号.发生时间, r_挂号.发生时间, 1, 1, r_挂号.操作员姓名, '自动转换产生序号');
            End If;
          Else
            Update 病人挂号记录 Set 出诊记录id = n_出诊记录id Where ID = r_挂号.Id;
            Update 临床出诊记录 Set 已挂数 = 已挂数 + 1 Where ID = n_出诊记录id;
            Update 临床出诊序号控制
            Set 挂号状态 = 1, 操作员姓名 = r_挂号.操作员姓名
            Where 记录id = n_出诊记录id And 序号 = r_挂号.号序;
            If Sql%RowCount = 0 Then
              Insert Into 临床出诊序号控制
                (记录id, 序号, 开始时间, 终止时间, 数量, 挂号状态, 操作员姓名, 备注)
              Values
                (n_出诊记录id, r_挂号.号序, r_挂号.发生时间, r_挂号.发生时间, 1, 1, r_挂号.操作员姓名, '自动转换产生序号');
            End If;
          End If;
        Else
          Update 病人挂号记录 Set 出诊记录id = n_出诊记录id Where ID = r_挂号.Id;
          Update 临床出诊记录 Set 已约数 = 已约数 + 1 Where ID = n_出诊记录id;
          Update 临床出诊序号控制
          Set 挂号状态 = 2, 操作员姓名 = r_挂号.操作员姓名
          Where 记录id = n_出诊记录id And 序号 = r_挂号.号序;
          If Sql%RowCount = 0 Then
            Insert Into 临床出诊序号控制
              (记录id, 序号, 开始时间, 终止时间, 数量, 挂号状态, 操作员姓名, 备注)
            Values
              (n_出诊记录id, r_挂号.号序, r_挂号.发生时间, r_挂号.发生时间, 1, 2, r_挂号.操作员姓名, '自动转换产生序号');
          End If;
        End If;
        n_处理数量 := n_处理数量 + 1;
      End If;
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_出诊表挂号_Turn;
/

--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
Create Or Replace Procedure Zl_临床出诊表_导入
(
  号码_In 挂号安排.号码%Type := Null,
  开始_In Number := 1
) As
  -------------------------------------------------------------------------
  --功能说明：导放临床出诊表,主要是根据挂号安排，挂号计划安排等表进行数据导入
  --入参：
  --    开始_In:传入号码时有效，表示第一个
  -------------------------------------------------------------------------
  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  l_限制id t_Numlist := t_Numlist();
  n_Count  Number(18);

  v_时间段           Varchar2(4000);
  n_出诊id           临床出诊表.Id%Type;
  v_全院号源归属站点 部门表.站点%Type;

  Procedure Zl_Register_Import
  (
    号码_In             挂号安排.号码%Type,
    出诊id_In           临床出诊表.Id%Type,
    全院号源归属站点_In 部门表.站点%Type
  ) As
    n_号源id   临床出诊号源.Id%Type;
    d_建档时间 临床出诊号源.建档时间%Type;
  
    n_出诊id 临床出诊表.Id%Type;
    n_安排id 临床出诊安排.Id%Type;
  
    n_限制id 临床出诊号源限制.Id%Type;
    n_诊室id 门诊诊室.Id%Type;
  
    n_是否导入     Number(2);
    n_是否临时安排 临床出诊安排.是否临时安排%Type;
  
    n_Count  Number(18);
    l_限制id t_Numlist := t_Numlist();
  Begin
    For c_号源 In (Select a.Id, a.号类, a.号码, a.科室id, a.项目id, a.医生姓名, Decode(a.医生id, 0, Null, a.医生id) As 医生id, a.序号, a.周日,
                        a.周一, a.周二, a.周三, a.周四, a.周五, a.周六, a.病案必须, a.分诊方式, a.序号控制, a.开始时间, a.终止时间, a.执行时间, a.执行计划id,
                        a.默认时段间隔, a.预约天数, Nvl(a.是否删除, 0) As 是否删除,
                        Nvl(a.停用日期, To_Date('3000-01-01', 'yyyy-mm-dd')) As 停用日期, Nvl(b.站点, 全院号源归属站点_In) As 站点
                 From 挂号安排 A, 部门表 B
                 Where a.科室id = b.Id And a.号码 = 号码_In
                      --科室，项目，医生相同的已经导入了一个号别就不再导入
                       And Not Exists (Select 1
                        From 临床出诊号源
                        Where 科室id = a.科室id And 项目id = a.项目id And Nvl(医生姓名, '-') = Nvl(a.医生姓名, '-') And
                              Nvl(医生id, 0) = Nvl(a.医生id, 0))) Loop
    
      n_是否导入 := 1;
      --对于科室，项目，医生三者都相同的多个号别，首先考虑导入有效号别中的第一个，如果没有，则导入失效号别中的第一个
      Select Count(1)
      Into n_Count
      From 挂号安排
      Where 科室id = c_号源.科室id And 项目id = c_号源.项目id And Nvl(医生姓名, '-') = Nvl(c_号源.医生姓名, '-') And
            Nvl(医生id, 0) = Nvl(c_号源.医生id, 0);
      If Nvl(n_Count, 0) = 1 Then
        --科室，项目，医生是唯一的
        n_是否导入 := 1;
      Else
        --是否存在未停用且未删除的号别
        Select Count(1)
        Into n_Count
        From 挂号安排
        Where 科室id = c_号源.科室id And 项目id = c_号源.项目id And Nvl(医生姓名, '-') = Nvl(c_号源.医生姓名, '-') And
              Nvl(医生id, 0) = Nvl(c_号源.医生id, 0) And c_号源.是否删除 = 0 And
              (停用日期 Is Null Or 停用日期 = To_Date('3000-01-01', 'yyyy-mm-dd'));
        If Nvl(n_Count, 0) = 0 Then
          --不存在未停用且未删除的号别，直接导入当前号别，即失效号别中的第一个
          n_是否导入 := 1;
        Elsif Nvl(n_Count, 0) = 1 Then
          --只存在一个未停用且未删除的号别，检查是不是当前号别
          If c_号源.是否删除 = 0 And c_号源.停用日期 = To_Date('3000-01-01', 'yyyy-mm-dd') Then
            n_是否导入 := 1;
          Else
            n_是否导入 := 0;
          End If;
        Else
          --检查当前号别是否已停用或已删除
          If Not (c_号源.是否删除 = 0 And c_号源.停用日期 = To_Date('3000-01-01', 'yyyy-mm-dd')) Then
            --已停用或已删除则不导入
            n_是否导入 := 0;
          Else
            --当前号别安排/计划是否有效
            Select Count(1)
            Into n_Count
            From 挂号安排计划
            Where 安排id = c_号源.Id And 审核时间 Is Not Null And 失效时间 > Sysdate And
                  Rownum < 2;
            If Nvl(n_Count, 0) = 0 Then
              --无计划
              Select Count(1)
              Into n_Count
              From 挂号安排 A
              Where a.Id = c_号源.Id And Nvl(a.终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate And
                    Not (a.周日 Is Null And a.周一 Is Null And a.周二 Is Null And a.周三 Is Null And a.周四 Is Null And
                     a.周五 Is Null And a.周六 Is Null);
            Else
              --只要生效时间大于当前时间或者不存在大于其生效时间小于当前时间的都是有效的，
              --因为1.生效时间和号码是唯一的，2.是以生效时间最近的来确定是有效的，
              --     即如果当前计划的生效时间小于等于当前时间且是所有生效时间小于等于当前时间的计划中生效时间最大的，则当前计划是有效的



              Select Count(1)
              Into n_Count
              From 挂号安排计划 A
              Where a.审核时间 Is Not Null And a.失效时间 > Sysdate And
                    a.安排id = c_号源.Id And
                    (Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) >= Sysdate Or
                    Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) < Sysdate And Not Exists
                     (Select 1
                      From 挂号安排计划
                      Where 安排id = a.安排id And 审核时间 Is Not Null And
                            失效时间 > Sysdate And
                            Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) < Sysdate And
                            Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) >
                            Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')))) And
                    Not (a.周日 Is Null And a.周一 Is Null And a.周二 Is Null And a.周三 Is Null And a.周四 Is Null And
                     a.周五 Is Null And a.周六 Is Null);
            End If;
          
            If Nvl(n_Count, 0) <> 0 Then
              --当前号别安排有效
              n_是否导入 := 1;
            Else
              --当前号别安排无效
              n_是否导入 := 0;
            End If;
          End If;
        End If;
      End If;
    
      If Nvl(n_是否导入, 0) = 1 Then
        Select 临床出诊号源_Id.Nextval Into n_号源id From Dual;
      
        Select Nvl(Min(开始时间), Sysdate)
        Into d_建档时间
        From (Select Min(开始时间) As 开始时间
               From 挂号安排
               Where ID = c_号源.Id
               Union All
               Select Min(生效时间) As 开始时间
               From 挂号安排计划
               Where 审核时间 Is Not Null And 失效时间 > Sysdate And 安排id = c_号源.Id);
      
        --1.处理临床出诊号源
        Insert Into 临床出诊号源
          (ID, 号类, 号码, 科室id, 项目id, 医生id, 医生姓名, 是否建病案, 预约天数, 出诊频次, 假日控制状态, 是否临床排班, 排班方式, 是否删除, 建档时间, 撤档时间)
        Values
          (n_号源id, c_号源.号类, c_号源.号码, c_号源.科室id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, c_号源.病案必须, c_号源.预约天数, c_号源.默认时段间隔, 2,
           0, 0, c_号源.是否删除, d_建档时间, c_号源.停用日期);
      
        --2.处理临床出诊停诊记录
        --一个医生一个停诊计划只导入一个，可能存在一个医生多个号别的情况，他们的停诊计划一样
        Insert Into 临床出诊停诊记录
          (ID, 记录id, 开始时间, 终止时间, 停诊原因, 申请人, 申请时间, 审批人, 审批时间, 登记人)
          Select 临床出诊停诊记录_Id.Nextval, Null, a.开始停止时间, a.结束停止时间, a.备注, b.医生姓名, a.制订日期, a.制订人, a.制订日期, a.制订人
          From 挂号安排停用状态 A, 挂号安排 B
          Where a.安排id = b.Id And b.Id = c_号源.Id And b.医生id Is Not Null And Not Exists
           (Select 1
                 From 临床出诊停诊记录
                 Where 记录id Is Null And 申请人 = b.医生姓名 And 开始时间 = a.开始停止时间 And 终止时间 = a.结束停止时间);
      
        --3.处理相关的出诊表数据
        --3.1 固定出诊表
        If c_号源.站点 Is Null Then
          n_出诊id := 出诊id_In;
        Else
          Begin
            Select ID Into n_出诊id From 临床出诊表 Where 排班方式 = 0 And Nvl(站点, '-') = c_号源.站点;
          Exception
            When Others Then
              n_出诊id := 0;
          End;
          If n_出诊id = 0 Then
            Update 临床出诊表
            Set 站点 = c_号源.站点
            Where 排班方式 = 0 And Nvl(站点, '-') = '-'
            Returning ID Into n_出诊id;
            If Sql%NotFound Then
              Select 临床出诊表_Id.Nextval Into n_出诊id From Dual;
              Insert Into 临床出诊表
                (ID, 排班方式, 出诊表名, 年份, 站点)
              Values
                (n_出诊id, 0, '固定出诊表', To_Number(To_Char(Sysdate, 'yyyy')), c_号源.站点);
            End If;
          End If;
        End If;
      
        --3.2导入临床出诊安排
        --失效的安排和计划不导入
        --只要生效时间大于当前时间或者不存在大于其生效时间小于当前时间的都是有效的，
        --因为1.生效时间和号码是唯一的，2.是以生效时间最近的来确定是有效的，
        --     即如果当前计划的生效时间小于等于当前时间且是所有生效时间小于等于当前时间的计划中生效时间最大的，则当前计划是有效的



        For c_详情 In (
                     --1.无计划的安排
                     Select a.Id As 安排id, -1 * Null As 计划id, a.科室id, a.项目id, a.医生姓名,
                             Decode(a.医生id, 0, Null, a.医生id) As 医生id, a.周日, a.周一, a.周二, a.周三, a.周四, a.周五, a.周六, a.分诊方式,
                             a.序号控制, Nvl(a.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')) As 开始时间,
                             Nvl(a.终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间
                     From 挂号安排 A
                     Where a.Id = c_号源.Id And Not Exists (Select 1
                            From 挂号安排计划
                            Where 安排id = a.Id And 审核时间 Is Not Null And
                                  失效时间 > Sysdate) And
                           Nvl(a.终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate
                     Union All
                     --有计划的安排,只导入有效的
                     Select a.安排id, a.Id As 计划id, b.科室id, a.项目id, a.医生姓名, Decode(a.医生id, 0, Null, a.医生id) As 医生id, a.周日,
                            a.周一, a.周二, a.周三, a.周四, a.周五, a.周六, a.分诊方式, a.序号控制,
                            Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) As 开始时间,
                            Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间
                     From 挂号安排计划 A, 挂号安排 B
                     Where a.安排id = b.Id And a.审核时间 Is Not Null And
                           a.失效时间 > Sysdate And b.Id = c_号源.Id And
                           (Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) >= Sysdate Or
                           Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) < Sysdate And Not Exists
                            (Select 1
                             From 挂号安排计划
                             Where 安排id = a.安排id And 审核时间 Is Not Null And
                                   失效时间 > Sysdate And
                                   Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) < Sysdate And
                                   Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) >
                                   Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd'))))) Loop
        
          Select 临床出诊安排_Id.Nextval Into n_安排id From Dual;
        
          n_诊室id := Null;
          If Nvl(c_详情.分诊方式, 0) = 1 Then
            Begin
              If Nvl(c_详情.计划id, 0) <> 0 Then
                Select a.Id
                Into n_诊室id
                From 门诊诊室 A, 挂号计划诊室 B
                Where a.名称 = b.门诊诊室 And b.计划id = c_详情.计划id And Rownum < 2;
              Else
                Select a.Id
                Into n_诊室id
                From 门诊诊室 A, 挂号安排诊室 B
                Where a.名称 = b.门诊诊室 And b.号表id = c_详情.安排id And Rownum < 2;
              End If;
            Exception
              When Others Then
                n_诊室id := Null;
            End;
          End If;
        
          --a.临床出诊安排
          Select Count(1)
          Into n_是否临时安排
          From 临床出诊安排
          Where 出诊id = n_出诊id And 号源id = n_号源id And Rownum < 2;
          Insert Into 临床出诊安排
            (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间, 是否临时安排)
          Values
            (n_安排id, n_出诊id, n_号源id, c_详情.项目id, c_详情.医生id, c_详情.医生姓名, c_详情.开始时间, c_详情.终止时间, Zl_Username, c_详情.开始时间,
             n_是否临时安排);
        
          --b.临床出诊限制
          --说明：限约数等于0表示禁止预约，限约数为空表示不限制预约
          If Nvl(c_详情.计划id, 0) <> 0 Then
            Select Count(1) Into n_Count From 挂号计划限制 Where 计划id = c_详情.计划id And Rownum < 2;
            If n_Count = 0 Then
              Insert Into 临床出诊限制
                (ID, 安排id, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 限制项目, 上班时段, 分诊方式, 诊室id)
                Select 临床出诊限制_Id.Nextval, n_安排id, Null, Null, c_详情.序号控制, Decode(b.星期, Null, 0, 1) As 是否分时段,
                       Nvl((Select 1
                            From 挂号计划限制
                            Where 计划id = c_详情.计划id And 限制项目 = a.限制项目 And 限约数 = 0 And Rownum < 2), 0), a.限制项目, a.上班时段,
                       c_详情.分诊方式, n_诊室id
                From (Select '周日' As 限制项目, c_详情.周日 As 上班时段
                       From Dual
                       Where c_详情.周日 Is Not Null
                       Union All
                       Select '周一', c_详情.周一
                       From Dual
                       Where c_详情.周一 Is Not Null
                       Union All
                       Select '周二', c_详情.周二
                       From Dual
                       Where c_详情.周二 Is Not Null
                       Union All
                       Select '周三', c_详情.周三
                       From Dual
                       Where c_详情.周三 Is Not Null
                       Union All
                       Select '周四', c_详情.周四
                       From Dual
                       Where c_详情.周四 Is Not Null
                       Union All
                       Select '周五', c_详情.周五
                       From Dual
                       Where c_详情.周五 Is Not Null
                       Union All
                       Select '周六', c_详情.周六 From Dual Where c_详情.周六 Is Not Null) A,
                     (Select Distinct 星期 From 挂号计划时段 Where 计划id = c_详情.计划id) B
                Where a.限制项目 = b.星期(+);
            Else
              Insert Into 临床出诊限制
                (ID, 安排id, 限制项目, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 分诊方式, 诊室id)
                Select 临床出诊限制_Id.Nextval, n_安排id, 限制项目,
                       Decode(限制项目, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                               c_详情.周五, '周六', c_详情.周六, Null), 限号数, 限约数, c_详情.序号控制, Decode(b.星期, Null, 0, 1) As 是否分时段,
                       Nvl((Select 1
                            From 挂号计划限制
                            Where 计划id = c_详情.计划id And 限制项目 = a.限制项目 And 限约数 = 0 And Rownum < 2), 0), c_详情.分诊方式, n_诊室id
                From 挂号计划限制 A, (Select Distinct 星期 From 挂号计划时段 Where 计划id = c_详情.计划id) B
                Where a.限制项目 = b.星期(+) And 计划id = c_详情.计划id;
            End If;
          Else
            Select Count(1) Into n_Count From 挂号安排限制 Where 安排id = c_详情.安排id And Rownum < 2;
            If n_Count = 0 Then
              Insert Into 临床出诊限制
                (ID, 安排id, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 限制项目, 上班时段, 分诊方式, 诊室id)
                Select 临床出诊限制_Id.Nextval, n_安排id, Null, Null, c_详情.序号控制, Decode(b.星期, Null, 0, 1) As 是否分时段,
                       Nvl((Select 1
                            From 挂号安排限制
                            Where 安排id = c_详情.安排id And 限制项目 = a.限制项目 And 限约数 = 0 And Rownum < 2), 0), a.限制项目, a.上班时段,
                       c_详情.分诊方式, n_诊室id
                From (Select '周日' As 限制项目, c_详情.周日 As 上班时段
                       From Dual
                       Where c_详情.周日 Is Not Null
                       Union All
                       Select '周一', c_详情.周一
                       From Dual
                       Where c_详情.周一 Is Not Null
                       Union All
                       Select '周二', c_详情.周二
                       From Dual
                       Where c_详情.周二 Is Not Null
                       Union All
                       Select '周三', c_详情.周三
                       From Dual
                       Where c_详情.周三 Is Not Null
                       Union All
                       Select '周四', c_详情.周四
                       From Dual
                       Where c_详情.周四 Is Not Null
                       Union All
                       Select '周五', c_详情.周五
                       From Dual
                       Where c_详情.周五 Is Not Null
                       Union All
                       Select '周六', c_详情.周六 From Dual Where c_详情.周六 Is Not Null) A,
                     (Select Distinct 星期 From 挂号安排时段 Where 安排id = c_详情.安排id) B
                Where a.限制项目 = b.星期(+);
            Else
              Insert Into 临床出诊限制
                (ID, 安排id, 限制项目, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 分诊方式, 诊室id)
                Select 临床出诊限制_Id.Nextval, n_安排id, 限制项目,
                       Decode(限制项目, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                               c_详情.周五, '周六', c_详情.周六, Null), 限号数, 限约数, c_详情.序号控制, Decode(b.星期, Null, 0, 1) As 是否分时段,
                       Nvl((Select 1
                            From 挂号安排限制
                            Where 安排id = c_详情.安排id And 限制项目 = a.限制项目 And 限约数 = 0 And Rownum < 2), 0), c_详情.分诊方式, n_诊室id
                From 挂号安排限制 A, (Select Distinct 星期 From 挂号安排时段 Where 安排id = c_详情.安排id) B
                Where a.限制项目 = b.星期(+) And 安排id = c_详情.安排id;
            End If;
          End If;
        
          --c.临床出诊诊室
          If Nvl(c_详情.分诊方式, 0) > 0 Then
            If Nvl(c_详情.计划id, 0) <> 0 Then
              Insert Into 临床出诊诊室
                (限制id, 诊室id)
                Select a.Id, b.诊室id
                From 临床出诊限制 A,
                     (Select Distinct a.Id As 诊室id
                       From 门诊诊室 A, 挂号计划诊室 B
                       Where a.名称 = b.门诊诊室 And b.计划id = c_详情.计划id) B
                Where a.安排id = n_安排id;
            Else
              Insert Into 临床出诊诊室
                (限制id, 诊室id)
                Select a.Id, b.诊室id
                From 临床出诊限制 A,
                     (Select Distinct a.Id As 诊室id
                       From 门诊诊室 A, 挂号安排诊室 B
                       Where a.名称 = b.门诊诊室 And b.号表id = c_详情.安排id) B
                Where a.安排id = n_安排id;
            End If;
          End If;
        
          --D.临床出诊时段
          If Nvl(c_详情.计划id, 0) <> 0 Then
            Insert Into 临床出诊时段
              (限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约)
              Select a.Id, b.序号, b.开始时间, b.结束时间, b.限制数量, b.是否预约
              From 临床出诊限制 A,
                   (Select n_安排id As 安排id, 星期,
                            Decode(星期, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                                    c_详情.周五, '周六', c_详情.周六, Null) As 上班时段, 序号, 开始时间, 结束时间, 限制数量, 是否预约



                     
                     From 挂号计划时段
                     Where 计划id = c_详情.计划id) B
              Where a.安排id = b.安排id And a.限制项目 = b.星期 And a.上班时段 = b.上班时段;
          
          Else
            Insert Into 临床出诊时段
              (限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约)
              Select a.Id, b.序号, b.开始时间, b.结束时间, b.限制数量, b.是否预约
              From 临床出诊限制 A,
                   (Select n_安排id As 安排id, 星期,
                            Decode(星期, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                                    c_详情.周五, '周六', c_详情.周六, Null) As 上班时段, 序号, 开始时间, 结束时间, 限制数量, 是否预约



                     
                     From 挂号安排时段
                     Where 安排id = c_详情.安排id) B
              Where a.安排id = b.安排id And a.限制项目 = b.星期 And a.上班时段 = b.上班时段;
          End If;
        
          --不分时段的序号控制号先生成序号
          --开始时间、终止时间填写时间段的开始时间和结束时间
          For c_限制项目 In (Select ID, 限号数, 上班时段
                         From 临床出诊限制
                         Where 安排id = n_安排id And Nvl(限号数, 0) <> 0 And Nvl(是否序号控制, 0) = 1 And Nvl(是否分时段, 0) = 0) Loop
            For I In 1 .. c_限制项目.限号数 Loop
              Insert Into 临床出诊时段
                (限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约)
                Select c_限制项目.Id, I, 开始时间, 终止时间, 1, 1
                From 时间段
                Where 站点 Is Null And 号类 Is Null And 时间段 = c_限制项目.上班时段;
            End Loop;
          End Loop;
        
          --任何一个都不允许预约时表示全部允许预约
          Update 临床出诊时段 A
          Set a.是否预约 = 1
          Where 限制id In (Select ID From 临床出诊限制 Where 安排id = n_安排id) And Not Exists
           (Select 1 From 临床出诊时段 B Where a.限制id = b.限制id And Nvl(b.是否预约, 0) = 1);
        
          --E.合作单位挂号控制
          If Nvl(c_详情.计划id, 0) <> 0 Then
            Insert Into 临床出诊挂号控制
              (限制id, 类型, 性质, 名称, 序号, 控制方式, 数量)
              Select a.Id, b.类型, b.性质, b.合作单位, b.序号, b.控制方式, b.数量
              From 临床出诊限制 A,
                   (Select 1 As 类型, 1 As 性质, 合作单位, n_安排id As 安排id, 限制项目,
                            Decode(限制项目, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                                    c_详情.周五, '周六', c_详情.周六, Null) As 上班时段, 序号,
                            Case
                               When Nvl(序号, 0) = 0 And Nvl(数量, 0) = 0 Then
                                0
                               When 序号 = 0 And Nvl(数量, 0) <> 0 Then
                                2
                               When Nvl(序号, 0) <> 0 And Nvl(数量, 0) <> 0 Then
                                3
                               Else
                                4
                             End As 控制方式, 数量
                     From 合作单位计划控制
                     Where 计划id = c_详情.计划id And
                           Decode(限制项目, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                                  c_详情.周五, '周六', c_详情.周六, Null) Is Not Null) B
              Where a.安排id = b.安排id And a.限制项目 = b.限制项目 And a.上班时段 = b.上班时段;
          Else
            Insert Into 临床出诊挂号控制
              (限制id, 类型, 性质, 名称, 序号, 控制方式, 数量)
              Select a.Id, b.类型, b.性质, b.合作单位, b.序号, b.控制方式, b.数量
              From 临床出诊限制 A,
                   (Select 1 As 类型, 1 As 性质, 合作单位, n_安排id As 安排id, 限制项目,
                            Decode(限制项目, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                                    c_详情.周五, '周六', c_详情.周六, Null) As 上班时段, 序号,
                            Case
                              When Nvl(序号, 0) = 0 And Nvl(数量, 0) = 0 Then
                               0
                              When 序号 = 0 And Nvl(数量, 0) <> 0 Then
                               2
                              When Nvl(序号, 0) <> 0 And Nvl(数量, 0) <> 0 Then
                               3
                              Else
                               4
                            End As 控制方式, 数量
                     From 合作单位安排控制
                     Where 安排id = c_详情.安排id And
                           Decode(限制项目, '周日', c_详情.周日, '周一', c_详情.周一, '周二', c_详情.周二, '周三', c_详情.周三, '周四', c_详情.周四, '周五',
                                  c_详情.周五, '周六', c_详情.周六, Null) Is Not Null) B
              Where a.安排id = b.安排id And a.限制项目 = b.限制项目 And a.上班时段 = b.上班时段;
          End If;
        End Loop;
      
        --4.停用没有有效安排的号源，并删除无效的安排
        --主要是处理所有计划已失效或者有效计划只有一个且这个计划周一到周日都没有上班时段的
        Select Count(1)
        Into n_Count
        From 临床出诊限制 A, 临床出诊安排 B, 临床出诊号源 C
        Where a.安排id = b.Id And b.号源id = c.Id And c.Id = n_号源id And a.上班时段 Is Not Null And Nvl(c.是否删除, 0) = 0 And
              Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And Rownum < 2;
        If n_Count = 0 Then
          --没有有效的安排，停用号源，删除安排
          Select a.Id Bulk Collect
          Into l_限制id
          From 临床出诊限制 A, 临床出诊安排 B
          Where a.安排id = b.Id And b.号源id = n_号源id;
        
          Forall I In 1 .. l_限制id.Count
            Delete From 临床出诊诊室 Where 限制id = l_限制id(I);
        
          Forall I In 1 .. l_限制id.Count
            Delete From 临床出诊时段 Where 限制id = l_限制id(I);
        
          Forall I In 1 .. l_限制id.Count
            Delete From 临床出诊挂号控制 Where 限制id = l_限制id(I);
        
          Forall I In 1 .. l_限制id.Count
            Delete From 临床出诊限制 Where ID = l_限制id(I);
        
          Delete From 临床出诊安排 Where 号源id = n_号源id;
        
          Update 临床出诊号源 Set 撤档时间 = Sysdate Where ID = n_号源id;
        End If;
      
        --5.拷贝一份出诊信息作为号源控制信息
        --说明：上班时段按安排的登记时间倒序取第一个
        For c_限制 In (Select ID, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 是否独占, 分诊方式, 诊室id
                     From (Select a.Id, a.上班时段, a.限号数, a.限约数, a.是否序号控制, a.是否分时段, a.预约控制, a.是否独占, a.分诊方式, a.诊室id,
                                   Row_Number() Over(Partition By a.上班时段 Order By b.登记时间 Desc) As 组号
                            From 临床出诊限制 A, 临床出诊安排 B
                            Where a.安排id = b.Id And b.号源id = n_号源id)
                     Where 组号 = 1) Loop
          --a.临床出诊号源限制
          Select 临床出诊号源限制_Id.Nextval Into n_限制id From Dual;
          Insert Into 临床出诊号源限制
            (ID, 号源id, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 是否独占, 分诊方式, 诊室id)
          Values
            (n_限制id, n_号源id, c_限制.上班时段, c_限制.限号数, c_限制.限约数, c_限制.是否序号控制, c_限制.是否分时段, c_限制.预约控制, c_限制.是否独占, c_限制.分诊方式,
             c_限制.诊室id);
          --b.临床出诊号源诊室
          Insert Into 临床出诊号源诊室
            (限制id, 诊室id)
            Select n_限制id, 诊室id From 临床出诊诊室 Where 限制id = c_限制.Id;
          --c.临床出诊号源时段
          Insert Into 临床出诊号源时段
            (限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约)
            Select n_限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约 From 临床出诊时段 Where 限制id = c_限制.Id;
          --d.临床出诊号源控制
          Insert Into 临床出诊号源控制
            (限制id, 类型, 性质, 名称, 序号, 控制方式, 数量)
            Select n_限制id, 类型, 性质, 名称, 序号, 控制方式, 数量 From 临床出诊挂号控制 Where 限制id = c_限制.Id;
        End Loop;
      End If;
    End Loop;
  End;
Begin
  If Nvl(开始_In, 0) = 1 Then
    Select Count(1) Into n_Count From 临床出诊表 A, 临床出诊安排 B Where a.Id = b.出诊id And Rownum < 2;
    If n_Count <> 0 Then
      v_Err_Msg := '当前已经存在临床出诊安排了，请先删除，否则不允许导入！';
      Raise Err_Item;
    End If;
  
    Begin
      Select f_List2str(Cast(Collect(s.时间段) As t_Strlist))
      Into v_时间段
      From (Select 时间段, Row_Number() Over(Partition By 时间段 Order By 时间段) As 组号
             From (Select Decode(b.行号, 1, a.周一, 2, a.周二, 3, a.周三, 4, a.周四, 5, a.周五, 6, a.周六, a.周日) As 时间段
                    From (Select 周一, 周二, 周三, 周四, 周五, 周六, 周日
                           From 挂号安排
                           Union All
                           Select 周一, 周二, 周三, 周四, 周五, 周六, 周日
                           From 挂号安排计划
                           Where 审核时间 Is Not Null And 失效时间 > Sysdate) A,
                         (Select Level As 行号 From Dual Connect By Level <= 7) B)
             Where 时间段 Is Not Null) S, 时间段 T
      Where s.时间段 = t.时间段(+) And t.时间段 Is Null And s.组号 = 1;
    Exception
      When Others Then
        v_时间段 := Null;
    End;
  
    If v_时间段 Is Not Null Then
      v_Err_Msg := '原挂号安排中的上班时间段【' || v_时间段 || '】不存在，请先在“基础设置>上班时间管理”中添加！';
      Raise Err_Item;
    End If;
  
    --删除现有所有号源，在调用之前已进行了提示
    Select a.Id Bulk Collect Into l_限制id From 临床出诊号源限制 A, 临床出诊号源 B Where a.号源id = b.Id;
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊号源诊室 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊号源时段 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊号源控制 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊号源限制 Where ID = l_限制id(I);
  
    Delete From 临床出诊号源;
  
    --删除所有停诊记录
    Delete From 临床出诊停诊记录;
  End If;

  --不分站点号源如果没有指定站点，则放入第一个出诊表中
  v_全院号源归属站点 := zl_GetSysParameter('未区分站点的号源的维护站点', 1114);
  Begin
    Select Min(ID) Into n_出诊id From 临床出诊表 Where 排班方式 = 0;
  Exception
    When Others Then
      n_出诊id := 0;
  End;
  If Nvl(n_出诊id, 0) = 0 Then
    Select 临床出诊表_Id.Nextval Into n_出诊id From Dual;
    Insert Into 临床出诊表
      (ID, 排班方式, 出诊表名, 年份, 站点)
    Values
      (n_出诊id, 0, '固定出诊表', To_Number(To_Char(Sysdate, 'yyyy')), Null);
  End If;

  If Not 号码_In Is Null Then
    Zl_Register_Import(号码_In, n_出诊id, v_全院号源归属站点);
    Return;
  End If;

  For c_号源 In (Select 号码 From 挂号安排 Order By ID Desc) Loop
    --删除以及停用的号源也全部导入
    Zl_Register_Import(c_号源.号码, n_出诊id, v_全院号源归属站点);
  End Loop;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊表_导入;
/

--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
Create Or Replace Procedure Zl_Third_Getregstatus
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --功能:获取某个科室当天排班医生的挂号数量
  --入参:Xml_In: 
  --  <IN> 
  --      <KSID></KSID>   --科室ID
  --  </IN> 
  --出参:Xml_Out 
  --  <OUTPUT> 
  --    <YS>
  --      <YSXM></YSXM>    --医生姓名
  --      <SYGHS></SYGHS>  --剩余挂号数
  --      <DDJZS></DDJZS>  --等待就诊数
  --      <SWGHS></SWGHS>  --上午挂号数
  --      <XWGHS></XWGHS>  --下午挂号数
  --      <QTGHS></QTGHS>  --全天挂号数
  --      <YSZJ><YSZJ>     --医生职级
  --    </YS>
  --    <YS>
  --      ...
  --    </YS>  
  --  </OUTPUT> 
  -------------------------------------------------------------------------------------------------- 

  n_科室id      挂号安排.科室id%Type;
  d_日期        Date;
  v_Para        Varchar2(5000);
  n_挂号模式    Number(3);
  d_启用时间    Date;
  v_Temp        Varchar2(32767); --临时XML 
  x_Templet     Xmltype; --模板XML 
  v_Err_Msg     Varchar2(200);
  n_已挂数      病人挂号汇总.已挂数%Type;
  n_限号数      挂号安排限制.限号数%Type;
  n_上午接诊    Number;
  n_下午接诊    Number;
  n_等待就诊    Number;
  v_出诊记录ids Varchar2(5000);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/KSID') Into n_科室id From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  d_日期 := Sysdate;

  v_Para     := zl_GetSysParameter(256);
  n_挂号模式 := Substr(v_Para, 1, 1);
  Begin
    d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_启用时间 := Null;
  End;

  If n_挂号模式 = 0 Or (n_挂号模式 = 1 And d_日期 < d_启用时间) Then
    --计划排班模式
    For r_医生 In (Select Distinct 医生id, 医生姓名, y.专业技术职务
                 From (Select a.医生id, a.医生姓名
                        From 挂号安排计划 A, 挂号安排 B
                        Where a.安排id = b.Id And b.停用日期 Is Null And a.审核时间 Is Not Null And b.科室id = n_科室id And
                              a.生效时间 = (Select Max(生效时间)
                                        From 挂号安排计划
                                        Where 安排id = b.Id And 审核时间 Is Not Null And d_日期 Between 生效时间 + 0 And 失效时间) And
                              Not Exists (Select 1
                               From 挂号安排停用状态
                               Where 安排id = b.Id And d_日期 Between 开始停止时间 And 结束停止时间) And
                              d_日期 Between Nvl(b.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                              Nvl(b.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))
                        Union All
                        Select a.医生id, a.医生姓名
                        From 挂号安排 A
                        Where a.停用日期 Is Null And a.科室id = n_科室id And
                              d_日期 Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                              Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Not Exists
                         (Select 1
                               From 挂号安排停用状态
                               Where 安排id = a.Id And d_日期 Between 开始停止时间 And 结束停止时间) And Not Exists
                         (Select 1
                               From 挂号安排计划
                               Where 安排id = a.Id And 审核时间 Is Not Null And d_日期 Between 生效时间 + 0 And 失效时间)) X, 人员表 Y
                 Where x.医生id = y.Id(+)) Loop
      If r_医生.医生姓名 Is Not Null Then
        v_Temp := '<YS>';
        v_Temp := v_Temp || '<YSXM>' || r_医生.医生姓名 || '</YSXM>';
        Select Nvl(Sum(a.已挂数), 0)
        Into n_已挂数
        From 病人挂号汇总 A,
             (Select Distinct 号码
               From (Select b.号码
                      From 挂号安排计划 A, 挂号安排 B
                      Where a.安排id = b.Id And b.停用日期 Is Null And a.审核时间 Is Not Null And b.科室id = n_科室id And
                            a.生效时间 = (Select Max(生效时间)
                                      From 挂号安排计划
                                      Where 安排id = b.Id And 审核时间 Is Not Null And d_日期 Between 生效时间 + 0 And 失效时间) And Not Exists
                       (Select 1
                             From 挂号安排停用状态
                             Where 安排id = b.Id And d_日期 Between 开始停止时间 And 结束停止时间) And
                            d_日期 Between Nvl(b.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                            Nvl(b.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))
                      Union All
                      Select a.号码
                      From 挂号安排 A
                      Where a.停用日期 Is Null And a.科室id = n_科室id And
                            d_日期 Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                            Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Not Exists
                       (Select 1
                             From 挂号安排停用状态
                             Where 安排id = a.Id And d_日期 Between 开始停止时间 And 结束停止时间) And Not Exists
                       (Select 1
                             From 挂号安排计划
                             Where 安排id = a.Id And 审核时间 Is Not Null And d_日期 Between 生效时间 + 0 And 失效时间))) B
        Where a.医生姓名 = r_医生.医生姓名 And a.科室id = n_科室id And a.号码 = b.号码 And 日期 = Trunc(d_日期);
      
        Select Nvl(Sum(限号数), 0)
        Into n_限号数
        From (Select c.限号数
               From 挂号安排计划 A, 挂号安排 B, 挂号计划限制 C
               Where a.Id = c.计划id And c.限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                       '周四', '6', '周五', '7', '周六', Null) And a.安排id = b.Id And
                     b.停用日期 Is Null And a.审核时间 Is Not Null And b.科室id = n_科室id And
                     a.生效时间 = (Select Max(生效时间)
                               From 挂号安排计划
                               Where 安排id = b.Id And 审核时间 Is Not Null And d_日期 Between 生效时间 + 0 And 失效时间) And Not Exists
                (Select 1 From 挂号安排停用状态 Where 安排id = b.Id And d_日期 Between 开始停止时间 And 结束停止时间) And
                     d_日期 Between Nvl(b.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                     Nvl(b.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))
               Union All
               Select b.限号数
               From 挂号安排 A, 挂号安排限制 B
               Where a.Id = b.安排id And b.限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                       '周四', '6', '周五', '7', '周六', Null) And a.停用日期 Is Null And
                     a.科室id = n_科室id And d_日期 Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                     Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Not Exists
                (Select 1 From 挂号安排停用状态 Where 安排id = a.Id And d_日期 Between 开始停止时间 And 结束停止时间) And Not Exists
                (Select 1
                      From 挂号安排计划
                      Where 安排id = a.Id And 审核时间 Is Not Null And d_日期 Between 生效时间 + 0 And 失效时间));
        If n_限号数 - n_已挂数 < 0 Then
          v_Temp := v_Temp || '<SYGHS>' || 0 || '</SYGHS>';
        Else
          v_Temp := v_Temp || '<SYGHS>' || To_Char(n_限号数 - n_已挂数) || '</SYGHS>';
        End If;
        
        Select Count(1)
        Into n_等待就诊
        From 病人挂号记录
        Where 执行人 = r_医生.医生姓名 And 执行部门id = n_科室id And Nvl(执行状态, 0) = 0 And 记录性质 = 1 And 记录状态 = 1 And
              发生时间 Between Trunc(d_日期) And Trunc(d_日期 + 1) - 1 / 24 / 60 / 60;
        v_Temp := v_Temp || '<DDJZS>' || n_等待就诊 || '</DDJZS>';
        
        Select Count(1)
        Into n_上午接诊
        From 病人挂号记录
        Where 执行人 = r_医生.医生姓名 And 执行部门id = n_科室id And 记录性质 = 1 And 记录状态 = 1 And
              发生时间 Between Trunc(d_日期) And Trunc(d_日期) + 0.5 - 1 / 24 / 60 / 60;
        v_Temp := v_Temp || '<SWGHS>' || n_上午接诊 || '</SWGHS>';
        
        Select Count(1)
        Into n_下午接诊
        From 病人挂号记录
        Where 执行人 = r_医生.医生姓名 And 执行部门id = n_科室id And 记录性质 = 1 And 记录状态 = 1 And
              发生时间 Between Trunc(d_日期) + 0.5 And Trunc(d_日期 + 1) - 1 / 24 / 60 / 60;
        v_Temp := v_Temp || '<XWGHS>' || n_下午接诊 || '</XWGHS>';
        
        v_Temp := v_Temp || '<QTGHS>' || To_Char(n_上午接诊 + n_下午接诊) || '</QTGHS>';
        v_Temp := v_Temp || '<YSZJ>' || r_医生.专业技术职务 || '</YSZJ>';
        v_Temp := v_Temp || '</YS>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      End If;
    End Loop;
  Else
    --出诊表排班模式
    For r_医生 In (Select Distinct a.医生id, a.医生姓名, b.专业技术职务
                 From 临床出诊记录 A, 人员表 B
                 Where a.医生id = b.Id(+) And a.科室id = n_科室id And 是否发布 = 1 And Nvl(是否锁定, 0) = 0 And
                       (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间))) Loop
      For r_出诊记录 In (Select Distinct ID
                     From 临床出诊记录
                     Where 医生姓名 = r_医生.医生姓名 And 科室id = n_科室id And 出诊日期 = Trunc(d_日期) And 是否发布 = 1 And Nvl(是否锁定, 0) = 0 And
                           (开始时间 < Nvl(停诊开始时间, 终止时间) Or 终止时间 > Nvl(停诊终止时间, 开始时间))) Loop
        v_出诊记录ids := v_出诊记录ids || ',' || r_出诊记录.Id;
      End Loop;
      If v_出诊记录ids Is Not Null Then
        v_出诊记录ids := Substr(v_出诊记录ids, 2);
      End If;
      Select Sum(限号数), Sum(已挂数)
      Into n_限号数, n_已挂数
      From 临床出诊记录 A, Table(f_Str2list(v_出诊记录ids)) B
      Where a.Id = b.Column_Value;
      v_Temp := '<YS>';
      v_Temp := v_Temp || '<YSXM>' || r_医生.医生姓名 || '</YSXM>';
      If n_限号数 - n_已挂数 < 0 Then
        v_Temp := v_Temp || '<SYGHS>' || 0 || '</SYGHS>';
      Else
        v_Temp := v_Temp || '<SYGHS>' || To_Char(n_限号数 - n_已挂数) || '</SYGHS>';
      End If;
      Select Count(1)
      Into n_等待就诊
      From 病人挂号记录 A, Table(f_Str2list(v_出诊记录ids)) B
      Where 执行人 = r_医生.医生姓名 And 执行部门id = n_科室id And Nvl(执行状态, 0) = 0 And 记录性质 = 1 And 记录状态 = 1 And
            a.出诊记录id = b.Column_Value;
      v_Temp := v_Temp || '<DDJZS>' || n_等待就诊 || '</DDJZS>';
      
      Select Count(1)
      Into n_上午接诊
      From 病人挂号记录 A, Table(f_Str2list(v_出诊记录ids)) B
      Where 执行人 = r_医生.医生姓名 And 执行部门id = n_科室id And 记录性质 = 1 And 记录状态 = 1 And
            发生时间 Between Trunc(d_日期) And Trunc(d_日期) + 0.5 - 1 / 24 / 60 / 60 And a.出诊记录id = b.Column_Value;
      v_Temp := v_Temp || '<SWGHS>' || n_上午接诊 || '</SWGHS>';
      
      Select Count(1)
      Into n_下午接诊
      From 病人挂号记录 A, Table(f_Str2list(v_出诊记录ids)) B
      Where 执行人 = r_医生.医生姓名 And 执行部门id = n_科室id And 记录性质 = 1 And 记录状态 = 1 And
            发生时间 Between Trunc(d_日期) + 0.5 And Trunc(d_日期 + 1) - 1 / 24 / 60 / 60 And a.出诊记录id = b.Column_Value;
      v_Temp := v_Temp || '<XWGHS>' || n_下午接诊 || '</XWGHS>';
      
      v_Temp := v_Temp || '<QTGHS>' || To_Char(n_上午接诊 + n_下午接诊) || '</QTGHS>';
      v_Temp := v_Temp || '<YSZJ>' || r_医生.专业技术职务 || '</YSZJ>';
      v_Temp := v_Temp || '</YS>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    End Loop;
  End If;

  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getregstatus;
/

--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
Create Or Replace Procedure Zl_Third_Getdoctor
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --功能:获取某个科室下的排班的医生，用于护士分诊时选择医生
  --入参:Xml_In: 
  --  <IN> 
  --      <KSID></KSID>   --科室ID
  --      <RQ></RQ>   --日期,默认为当天
  --  </IN> 
  --出参:Xml_Out 
  --  <OUTPUT> 
  --    <YS>
  --      <YSID></YSID>    --医生ID
  --      <YSXM></YSXM>     --医生姓名
  --    </YS>
  --    <YS>
  --      ...
  --    </YS>  
  --  </OUTPUT> 
  -------------------------------------------------------------------------------------------------- 

  n_科室id   挂号安排.科室id%Type;
  d_日期     Date;
  v_Para     Varchar2(5000);
  n_挂号模式 Number(3);
  d_启用时间 Date;
  n_所有医生 Number(3);
  v_Temp     Varchar2(32767); --临时XML 
  x_Templet  Xmltype; --模板XML 
  v_Err_Msg  Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/KSID'), To_Date(Extractvalue(Value(A), 'IN/RQ'), 'yyyy-mm-dd hh24:mi:ss')
  Into n_科室id, d_日期
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If d_日期 Is Null Then
    d_日期 := Sysdate;
  End If;

  v_Para     := zl_GetSysParameter(256);
  n_挂号模式 := Substr(v_Para, 1, 1);
  Begin
    d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_启用时间 := Null;
  End;

  If n_挂号模式 = 0 Or (n_挂号模式 = 1 And d_日期 < d_启用时间) Then
    --计划排班模式
    For r_医生 In (Select Distinct 医生id, 医生姓名
                 From (Select a.医生id, a.医生姓名
                        From 挂号安排计划 A, 挂号安排 B
                        Where a.安排id = b.Id And b.停用日期 Is Null And a.审核时间 Is Not Null And b.科室id = n_科室id And
                              a.生效时间 = (Select Max(生效时间)
                                        From 挂号安排计划
                                        Where 安排id = b.Id And 审核时间 Is Not Null And d_日期 Between 生效时间 + 0 And 失效时间) And
                              Not Exists (Select 1
                               From 挂号安排停用状态
                               Where 安排id = b.Id And d_日期 Between 开始停止时间 And 结束停止时间) And
                              d_日期 Between Nvl(b.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                              Nvl(b.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))
                        Union All
                        Select a.医生id, a.医生姓名
                        From 挂号安排 A
                        Where a.停用日期 Is Null And a.科室id = n_科室id And
                              d_日期 Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                              Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Not Exists
                         (Select 1
                               From 挂号安排停用状态
                               Where 安排id = a.Id And d_日期 Between 开始停止时间 And 结束停止时间) And Not Exists
                         (Select 1
                               From 挂号安排计划
                               Where 安排id = a.Id And 审核时间 Is Not Null And d_日期 Between 生效时间 + 0 And 失效时间))) Loop
      If Nvl(r_医生.医生id, 0) <> 0 Or Nvl(r_医生.医生姓名, '-') <> '-' Then
        v_Temp := '<YS><YSID>' || r_医生.医生id || '</YSID><YSXM>' || r_医生.医生姓名 || '</YSXM></YS>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      Else
        n_所有医生 := 1;
        Exit;
      End If;
    End Loop;
  Else
    --出诊表排班模式
    For r_医生 In (Select Distinct 医生id, 医生姓名
                 From 临床出诊记录
                 Where 科室id = n_科室id And 出诊日期 = Trunc(d_日期) And 是否发布 = 1 And Nvl(是否锁定, 0) = 0 And
                       (开始时间 < Nvl(停诊开始时间, 终止时间) Or 终止时间 > Nvl(停诊终止时间, 开始时间))) Loop
      If Nvl(r_医生.医生id, 0) <> 0 Or Nvl(r_医生.医生姓名, '-') <> '-' Then
        v_Temp := '<YS><YSID>' || r_医生.医生id || '</YSID><YSXM>' || r_医生.医生姓名 || '</YSXM></YS>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      Else
        n_所有医生 := 1;
        Exit;
      End If;
    End Loop;
  End If;

  If Nvl(n_所有医生, 0) = 1 Then
    x_Templet := Xmltype('<OUTPUT></OUTPUT>');
    For r_医生 In (Select Distinct a.Id, a.姓名
                 From 人员表 A, 部门人员 B, 人员性质说明 C
                 Where a.Id = b.人员id And b.部门id = n_科室id And a.Id = c.人员id And c.人员性质 = '医生' And
                       (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)) Loop
      v_Temp := '<YS><YSID>' || r_医生.Id || '</YSID><YSXM>' || r_医生.姓名 || '</YSXM></YS>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    End Loop;
  End If;

  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getdoctor;
/

--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
CREATE OR REPLACE Procedure Zl_Third_Docarrange
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:医生排班计划
  --入参:Xml_In:
  --<IN>
  --   <YSID>870</YSID>    //医生ID
  --   <KDID>870</KSID>    //科室ID
  --   <KSSJ>2014-10-29 </KSSJ>    //开始时间
  --   <CXTS>14</CXTS>    //查询天数
  --   <HZDW>支付宝</HZDW> //合作单位
  --   <HL>号类</HL>      //号类，可传多个，用逗号分隔，格式:普通,专家,...
  --   <ZD></ZD>        //站点
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  --   <PBLIST>       //未返回该节点表示没有数据
  --    <PB>
  --     <RQ>2014-10-29</RQ>     //日期
  --     <SYHS>5</SYHS>    //剩余号数
  --     <SBSJ>全日</SBSJ>             //上班时间
  --     <YGS>5</YGS>    //已挂号数
  --    </PB>
  --   <PBLIST>
  --   <ERROR><MSG></MSG></ERROR> //错误情况返回
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  n_限号数       挂号安排限制.限号数%Type;
  n_已挂数       挂号安排限制.限号数%Type;
  n_总已挂数     挂号安排限制.限号数%Type;
  n_限约数       挂号安排限制.限号数%Type;
  n_已约数       挂号安排限制.限号数%Type;
  n_剩余数       挂号安排限制.限号数%Type;
  v_上班时间     varchar2(300);
  n_医生id       人员表.Id%Type;
  n_科室id       部门表.Id%Type;
  n_查询天数     Number(4);
  n_合作单位数量 Number(5);
  n_合约已挂数   Number(4);
  n_合约存在     Number(3);
  n_安排存在     Number(3);
  v_号码         挂号安排.号码%Type;
  n_安排id       挂号安排计划.安排id%Type;
  n_计划id       挂号安排计划.Id%Type;
  v_合作单位     挂号合作单位.名称%Type;
  n_Daycount     Number(4);
  d_开始时间     Date;
  d_原始时间     Date;
  n_禁用         Number(3);
  v_Temp         varchar2(32767); --临时XML
  x_Templet      Xmltype; --模板XML
  v_Err_Msg      varchar2(200);
  v_号类         varchar2(200);
  n_Exists       Number(2);
  n_预约天数     Number(5);
  n_补充天数     Number(5);
  n_挂号模式     Number(3);
  n_合约模式     临床出诊挂号控制记录.控制方式%Type;
  v_启用时间     varchar2(500);
  v_普通等级     varchar2(100);
  v_Pricegrade   varchar2(500);
  v_站点         部门表.站点%Type;
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/YSID'), Extractvalue(Value(A), 'IN/KSID'), Extractvalue(Value(A), 'IN/CXTS'),
         To_Date(Extractvalue(Value(A), 'IN/KSSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/HZDW'),
         Extractvalue(Value(A), 'IN/HL'), Extractvalue(Value(A), 'IN/ZD')
  Into n_医生id, n_科室id, n_查询天数, d_开始时间, v_合作单位, v_号类, v_站点
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_查询天数 := Nvl(n_查询天数, 14);
  n_预约天数 := Nvl(zl_GetSysParameter(66), 7);
  d_原始时间 := Trunc(d_开始时间);
  d_开始时间 := Trunc(d_开始时间);
  n_Daycount := 0;

  v_Pricegrade := Zl_Get_Pricegrade(v_站点);
  v_普通等级   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);

  n_挂号模式 := To_Number(Substr(Nvl(zl_GetSysParameter('挂号排班模式'), '0'), 1, 1));
  v_启用时间 := Substr(Nvl(zl_GetSysParameter('挂号排班模式'), '0'), 3);
  If n_挂号模式 = 0 Then
    If Nvl(n_科室id, 0) = 0 Then
      While (n_Daycount < n_查询天数) Loop
        If Trunc(d_开始时间 + n_Daycount) = Trunc(Sysdate) Then
          d_开始时间 := Sysdate - n_Daycount;
        Else
          d_开始时间 := d_原始时间;
        End If;
        n_安排存在 := 0;
        v_上班时间 := Null;
        n_总已挂数 := 0;
        n_已挂数   := 0;
        n_剩余数   := 0;
        n_限号数   := 0;
        n_已约数   := 0;
        n_限约数   := 0;
        For r_排班 In (Select d_开始时间 + n_Daycount As 日期, a.排班, a.限号数, a.限约数, Nvl(Hz.已挂数, 0) As 已挂数, Nvl(Hz.已约数, 0) As 已约数,
                            a.安排id, a.计划id, a.号码, a.号类
                     
                     From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Nvl(Ap.医生id, 0) As 医生id, Ry.专业技术职务 As 职称, Ap.号码,
                                   Ap.安排id, Ap.计划id, Ap.排班, Ap.项目id, Fy.名称 As 项目名称, Ap.序号控制, Ap.限号数, Ap.限约数



                            
                            From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Ap.医生id, Ap.号码, Ap.Id As 安排id, 0 As 计划id,
                                          Ap.项目id, Nvl(Ap.序号控制, 0) As 序号控制,
                                          Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Ap.周日, '2', Ap.周一, '3', Ap.周二, '4',
                                                  Ap.周三, '5', Ap.周四, '6', Ap.周五, '7', Ap.周六, Null) As 排班,
                                          Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                   From 挂号安排 Ap, 部门表 Bm, 挂号安排限制 Xz
                                   Where Ap.科室id = Bm.Id(+) And Ap.医生id = n_医生id And
                                         Sysdate + Nvl(Ap.预约天数, n_预约天数) >= d_开始时间 + n_Daycount And Ap.停用日期 Is Null And
                                         d_开始时间 + n_Daycount Between Nvl(Ap.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Nvl(Ap.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.安排id(+) = Ap.Id And
                                         Xz.限制项目(+) =
                                         Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                    (Select Rownum
                                          From 挂号安排停用状态 Ty
                                          Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                         Not Exists
                                    (Select Rownum
                                          From 挂号安排计划 Jh
                                          Where Jh.安排id = Ap.Id And Jh.审核时间 Is Not Null And
                                                d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Jh.失效时间)
                                   Union All
                                   Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Jh.医生姓名, Jh.医生id, Ap.号码, Ap.Id As 安排id,
                                          Jh.Id As 计划id, Jh.项目id, Nvl(Jh.序号控制, 0) As 序号控制,
                                          Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Jh.周日, '2', Jh.周一, '3', Jh.周二, '4',
                                                  Jh.周三, '5', Jh.周四, '6', Jh.周五, '7', Jh.周六, Null) As 排班,
                                          Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                   From 挂号安排计划 Jh, 挂号安排 Ap, 部门表 Bm, 挂号计划限制 Xz
                                   Where Jh.安排id = Ap.Id And Ap.科室id = Bm.Id(+) And
                                         Sysdate + Nvl(Ap.预约天数, n_预约天数) >= d_开始时间 + n_Daycount And Ap.停用日期 Is Null And
                                         Jh.医生id = n_医生id And
                                         d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Jh.失效时间 And Xz.计划id(+) = Jh.Id And
                                         Xz.限制项目(+) =
                                         Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                    (Select Rownum
                                          From 挂号安排停用状态 Ty
                                          Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                         (Jh.生效时间, Jh.安排id) =
                                         (Select Max(Sxjh.生效时间) As 生效时间, 安排id
                                          From 挂号安排计划 Sxjh
                                          Where Sxjh.审核时间 Is Not Null And d_开始时间 + n_Daycount Between
                                                Nvl(Sxjh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Nvl(Sxjh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.安排id = Jh.安排id
                                          Group By Sxjh.安排id)) Ap, 部门表 Bm, 人员表 Ry, 收费项目目录 Fy
                            Where Ap.科室id = Bm.Id(+) And Ap.医生id = Ry.Id(+) And Ap.项目id = Fy.Id And 排班 Is Not Null) A,
                          病人挂号汇总 Hz, 收费价目 B
                     Where a.号码 = Hz.号码(+) And Hz.日期(+) = Trunc(d_开始时间 + n_Daycount) And a.项目id = b.收费细目id And
                           b.终止日期 > d_开始时间 + n_Daycount And b.执行日期 <= d_开始时间 + n_Daycount And
                           (b.价格等级 = v_普通等级 Or
                           (b.价格等级 Is Null And Not Exists
                            (Select 1
                              From 收费价目
                              Where b.收费细目id = 收费细目id And 价格等级 = v_普通等级 And Sysdate Between 执行日期 And
                                    Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))) Loop
          If v_号类 Is Null Or Instr(',' || v_号类 || ',', ',' || r_排班.号类 || ',') > 0 Then
            v_上班时间 := v_上班时间 || '+' || r_排班.排班;
            n_总已挂数 := n_总已挂数 + r_排班.已挂数;
            n_已挂数   := r_排班.已挂数;
            n_限号数   := r_排班.限号数;
            n_已约数   := r_排班.已约数;
            n_限约数   := r_排班.限约数;
            n_安排id   := Nvl(r_排班.安排id, 0);
            n_计划id   := Nvl(r_排班.计划id, 0);
            v_号码     := r_排班.号码;
            n_安排存在 := 1;
            If v_上班时间 Is Not Null Then
              If v_合作单位 Is Not Null Then
                If n_计划id <> 0 Then
                  Begin
                    Select 1
                    Into n_合约存在
                    From 合作单位计划控制
                    Where 计划id = n_计划id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_合约存在 := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_合约存在
                    From 合作单位安排控制
                    Where 安排id = n_安排id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_合约存在 := 0;
                  End;
                End If;
              End If;
            
              If v_合作单位 Is Null Or Nvl(n_合约存在, 0) = 0 Then
                If n_计划id <> 0 Then
                  Begin
                    Select Sum(数量)
                    Into n_合作单位数量
                    From 合作单位计划控制
                    Where 计划id = n_计划id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null);
                  Exception
                    When Others Then
                      n_合作单位数量 := 0;
                  End;
                Else
                  Begin
                    Select Sum(数量)
                    Into n_合作单位数量
                    From 合作单位安排控制
                    Where 安排id = n_安排id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null);
                  Exception
                    When Others Then
                      n_合作单位数量 := 0;
                  End;
                End If;
                Begin
                  Select Count(1)
                  Into n_合约已挂数
                  From 病人挂号记录
                  Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 Is Not Null And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                        Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                Exception
                  When Others Then
                    n_合约已挂数 := 0;
                End;
                If n_合作单位数量 = 0 Then
                  n_合作单位数量 := Null;
                End If;
                If Trunc(d_开始时间 + n_Daycount) > Trunc(Sysdate) Then
                  n_剩余数 := Nvl(n_剩余数, 0) + n_限约数 - n_已约数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                Else
                  n_剩余数 := Nvl(n_剩余数, 0) + n_限号数 - n_已挂数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                End If;
              Else
                --合约单位
                If n_计划id <> 0 Then
                  Begin
                    Select 1
                    Into n_禁用
                    From 合作单位计划控制
                    Where 计划id = n_计划id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_禁用 := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_禁用
                    From 合作单位安排控制
                    Where 安排id = n_安排id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_禁用 := 0;
                  End;
                End If;
                If Nvl(n_禁用, 0) = 0 Then
                  Begin
                    Select Count(1)
                    Into n_合约已挂数
                    From 病人挂号记录
                    Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                          Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_合约已挂数 := 0;
                  End;
                  If n_计划id <> 0 Then
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位计划控制
                      Where 计划id = n_计划id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  Else
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位安排控制
                      Where 安排id = n_安排id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  End If;
                  If n_合作单位数量 = 0 Then
                    n_合作单位数量 := Null;
                  End If;
                  n_剩余数 := Nvl(n_剩余数, 0) + Nvl(n_合作单位数量, n_合约已挂数) - Nvl(n_合约已挂数, 0);
                
                End If;
              End If;
            End If;
            n_合作单位数量 := 0;
            n_合约存在     := 0;
            n_禁用         := 0;
          End If;
        End Loop;
        v_上班时间 := Substr(v_上班时间, 2);
        If n_安排存在 = 1 Then
          v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                    '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 || '</YGS>' ||
                    '</PB>';
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    Else
      While (n_Daycount < n_查询天数) Loop
        If Trunc(d_开始时间 + n_Daycount) = Trunc(Sysdate) Then
          d_开始时间 := Sysdate - n_Daycount;
        Else
          d_开始时间 := d_原始时间;
        End If;
        v_上班时间 := Null;
        n_总已挂数 := 0;
        n_已挂数   := 0;
        n_剩余数   := 0;
        n_限号数   := 0;
        n_已约数   := 0;
        n_限约数   := 0;
        n_安排存在 := 0;
        For r_排班 In (Select d_开始时间 + n_Daycount As 日期, a.排班, a.限号数, a.限约数, Nvl(Hz.已挂数, 0) As 已挂数, Nvl(Hz.已约数, 0) As 已约数,
                            a.安排id, a.计划id, a.号码, a.号类
                     
                     From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Nvl(Ap.医生id, 0) As 医生id, Ry.专业技术职务 As 职称, Ap.号码,
                                   Ap.安排id, Ap.计划id, Ap.排班, Ap.项目id, Fy.名称 As 项目名称, Ap.序号控制, Ap.限号数, Ap.限约数



                            
                            From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Ap.医生id, Ap.号码, Ap.Id As 安排id, 0 As 计划id,
                                          Ap.项目id, Nvl(Ap.序号控制, 0) As 序号控制,
                                          Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Ap.周日, '2', Ap.周一, '3', Ap.周二, '4',
                                                  Ap.周三, '5', Ap.周四, '6', Ap.周五, '7', Ap.周六, Null) As 排班,
                                          Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                   From 挂号安排 Ap, 部门表 Bm, 挂号安排限制 Xz
                                   Where Ap.科室id = Bm.Id(+) And Sysdate + Nvl(Ap.预约天数, n_预约天数) >= d_开始时间 + n_Daycount And
                                         Ap.医生id = n_医生id And Ap.科室id = n_科室id And Ap.停用日期 Is Null And
                                         d_开始时间 + n_Daycount Between Nvl(Ap.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Nvl(Ap.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.安排id(+) = Ap.Id And
                                         Xz.限制项目(+) =
                                         Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                    (Select Rownum
                                          From 挂号安排停用状态 Ty
                                          Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                         Not Exists
                                    (Select Rownum
                                          From 挂号安排计划 Jh
                                          Where Jh.安排id = Ap.Id And Jh.审核时间 Is Not Null And
                                                d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Jh.失效时间)
                                   Union All
                                   Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Jh.医生姓名, Jh.医生id, Ap.号码, Ap.Id As 安排id,
                                          Jh.Id As 计划id, Jh.项目id, Nvl(Jh.序号控制, 0) As 序号控制,
                                          Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Jh.周日, '2', Jh.周一, '3', Jh.周二, '4',
                                                  Jh.周三, '5', Jh.周四, '6', Jh.周五, '7', Jh.周六, Null) As 排班,
                                          Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                   From 挂号安排计划 Jh, 挂号安排 Ap, 部门表 Bm, 挂号计划限制 Xz
                                   Where Jh.安排id = Ap.Id And Sysdate + Nvl(Ap.预约天数, n_预约天数) >= d_开始时间 + n_Daycount And
                                         Ap.科室id = Bm.Id(+) And Ap.停用日期 Is Null And Jh.医生id = n_医生id And Ap.科室id = n_科室id And
                                         d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Jh.失效时间 And Xz.计划id(+) = Jh.Id And
                                         Xz.限制项目(+) =
                                         Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                    (Select Rownum
                                          From 挂号安排停用状态 Ty
                                          Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                         (Jh.生效时间, Jh.安排id) =
                                         (Select Max(Sxjh.生效时间) As 生效时间, 安排id
                                          From 挂号安排计划 Sxjh
                                          Where Sxjh.审核时间 Is Not Null And d_开始时间 + n_Daycount Between
                                                Nvl(Sxjh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Nvl(Sxjh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.安排id = Jh.安排id
                                          Group By Sxjh.安排id)) Ap, 部门表 Bm, 人员表 Ry, 收费项目目录 Fy
                            Where Ap.科室id = Bm.Id(+) And Ap.医生id = Ry.Id(+) And Ap.项目id = Fy.Id And 排班 Is Not Null) A,
                          病人挂号汇总 Hz, 收费价目 B
                     Where a.号码 = Hz.号码(+) And Hz.日期(+) = Trunc(d_开始时间 + n_Daycount) And a.项目id = b.收费细目id And
                           b.终止日期 > d_开始时间 + n_Daycount And b.执行日期 <= d_开始时间 + n_Daycount And
                           (b.价格等级 = v_普通等级 Or
                           (b.价格等级 Is Null And Not Exists
                            (Select 1
                              From 收费价目
                              Where b.收费细目id = 收费细目id And 价格等级 = v_普通等级 And Sysdate Between 执行日期 And
                                    Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                     
                     ) Loop
          If v_号类 Is Null Or Instr(',' || v_号类 || ',', ',' || r_排班.号类 || ',') > 0 Then
            v_上班时间 := v_上班时间 || '+' || r_排班.排班;
            n_总已挂数 := n_总已挂数 + r_排班.已挂数;
            n_已挂数   := r_排班.已挂数;
            n_限号数   := r_排班.限号数;
            n_已约数   := r_排班.已约数;
            n_限约数   := r_排班.限约数;
            n_安排id   := Nvl(r_排班.安排id, 0);
            n_计划id   := Nvl(r_排班.计划id, 0);
            v_号码     := r_排班.号码;
            n_安排存在 := 1;
          
            If v_上班时间 Is Not Null Then
              If v_合作单位 Is Not Null Then
                If n_计划id <> 0 Then
                  Begin
                    Select 1
                    Into n_合约存在
                    From 合作单位计划控制
                    Where 计划id = n_计划id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_合约存在 := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_合约存在
                    From 合作单位安排控制
                    Where 安排id = n_安排id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_合约存在 := 0;
                  End;
                End If;
              End If;
            
              If v_合作单位 Is Null Or Nvl(n_合约存在, 0) = 0 Then
                If n_计划id <> 0 Then
                  Begin
                    Select Sum(数量)
                    Into n_合作单位数量
                    From 合作单位计划控制
                    Where 计划id = n_计划id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null);
                  Exception
                    When Others Then
                      n_合作单位数量 := 0;
                  End;
                Else
                  Begin
                    Select Sum(数量)
                    Into n_合作单位数量
                    From 合作单位安排控制
                    Where 安排id = n_安排id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null);
                  Exception
                    When Others Then
                      n_合作单位数量 := 0;
                  End;
                End If;
                Begin
                  Select Count(1)
                  Into n_合约已挂数
                  From 病人挂号记录
                  Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 Is Not Null And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                        Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                Exception
                  When Others Then
                    n_合约已挂数 := 0;
                End;
                If n_合作单位数量 = 0 Then
                  n_合作单位数量 := Null;
                End If;
                If Trunc(d_开始时间 + n_Daycount) > Trunc(Sysdate) Then
                  n_剩余数 := Nvl(n_剩余数, 0) + n_限约数 - n_已约数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                Else
                  n_剩余数 := Nvl(n_剩余数, 0) + n_限号数 - n_已挂数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                End If;
              Else
                --合约单位
                If n_计划id <> 0 Then
                  Begin
                    Select 1
                    Into n_禁用
                    From 合作单位计划控制
                    Where 计划id = n_计划id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_禁用 := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_禁用
                    From 合作单位安排控制
                    Where 安排id = n_安排id And
                          限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                        '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_禁用 := 0;
                  End;
                End If;
                If Nvl(n_禁用, 0) = 0 Then
                  Begin
                    Select Count(1)
                    Into n_合约已挂数
                    From 病人挂号记录
                    Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                          Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_合约已挂数 := 0;
                  End;
                  If n_计划id <> 0 Then
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位计划控制
                      Where 计划id = n_计划id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  Else
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位安排控制
                      Where 安排id = n_安排id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  End If;
                  If n_合作单位数量 = 0 Then
                    n_合作单位数量 := Null;
                  End If;
                  n_剩余数 := Nvl(n_剩余数, 0) + Nvl(n_合作单位数量, n_合约已挂数) - Nvl(n_合约已挂数, 0);
                
                End If;
              End If;
            End If;
            n_合作单位数量 := 0;
            n_合约存在     := 0;
            n_禁用         := 0;
          End If;
        End Loop;
        v_上班时间 := Substr(v_上班时间, 2);
        If n_安排存在 = 1 Then
          v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                    '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 || '</YGS>' ||
                    '</PB>';
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    End If;
    v_Temp := '<PBLIST>' || v_Temp || '</PBLIST>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Else
    --出诊表排班模式
    n_补充天数 := Zl_Fun_Getappointmentdays;
    If Nvl(n_科室id, 0) = 0 Then
      --通过医生查找
      While (n_Daycount < n_查询天数) Loop
        If Trunc(d_开始时间 + n_Daycount) < To_Date(Substr(v_启用时间, 1, Instr(v_启用时间, ' ') - 1), 'yyyy-mm-dd') Then
          n_安排存在 := 0;
          v_上班时间 := Null;
          n_总已挂数 := 0;
          n_已挂数   := 0;
          n_剩余数   := 0;
          n_限号数   := 0;
          n_已约数   := 0;
          n_限约数   := 0;
          For r_排班 In (Select d_开始时间 + n_Daycount As 日期, a.排班, a.限号数, a.限约数, Nvl(Hz.已挂数, 0) As 已挂数, Nvl(Hz.已约数, 0) As 已约数,
                              a.安排id, a.计划id, a.号码, a.号类
                       From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Nvl(Ap.医生id, 0) As 医生id, Ry.专业技术职务 As 职称,
                                     Ap.号码, Ap.安排id, Ap.计划id, Ap.排班, Ap.项目id, Fy.名称 As 项目名称, Ap.序号控制, Ap.限号数, Ap.限约数
                              
                              From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Ap.医生id, Ap.号码, Ap.Id As 安排id, 0 As 计划id,
                                            Ap.项目id, Nvl(Ap.序号控制, 0) As 序号控制,
                                            Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Ap.周日, '2', Ap.周一, '3', Ap.周二, '4',
                                                    Ap.周三, '5', Ap.周四, '6', Ap.周五, '7', Ap.周六, Null) As 排班,
                                            Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                     From 挂号安排 Ap, 部门表 Bm, 挂号安排限制 Xz
                                     Where Ap.科室id = Bm.Id(+) And Sysdate + Nvl(Ap.预约天数, n_预约天数) >= d_开始时间 + n_Daycount And
                                           Ap.医生id = n_医生id And Ap.停用日期 Is Null And
                                           d_开始时间 + n_Daycount Between Nvl(Ap.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Nvl(Ap.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.安排id(+) = Ap.Id And
                                           Xz.限制项目(+) =
                                           Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                                  '周三', '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                      (Select Rownum
                                            From 挂号安排停用状态 Ty
                                            Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                           Not Exists (Select Rownum
                                            From 挂号安排计划 Jh
                                            Where Jh.安排id = Ap.Id And Jh.审核时间 Is Not Null And
                                                  d_开始时间 + n_Daycount Between
                                                  Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Jh.失效时间)
                                     Union All
                                     Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Jh.医生姓名, Jh.医生id, Ap.号码, Ap.Id As 安排id,
                                            Jh.Id As 计划id, Jh.项目id, Nvl(Jh.序号控制, 0) As 序号控制,
                                            Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Jh.周日, '2', Jh.周一, '3', Jh.周二, '4',
                                                    Jh.周三, '5', Jh.周四, '6', Jh.周五, '7', Jh.周六, Null) As 排班,
                                            Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                     From 挂号安排计划 Jh, 挂号安排 Ap, 部门表 Bm, 挂号计划限制 Xz
                                     Where Jh.安排id = Ap.Id And Sysdate + Nvl(Ap.预约天数, n_预约天数) >= d_开始时间 + n_Daycount And
                                           Ap.科室id = Bm.Id(+) And Ap.停用日期 Is Null And Jh.医生id = n_医生id And
                                           d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Jh.失效时间 And Xz.计划id(+) = Jh.Id And
                                           Xz.限制项目(+) =
                                           Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                                  '周三', '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                      (Select Rownum
                                            From 挂号安排停用状态 Ty
                                            Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                           (Jh.生效时间, Jh.安排id) =
                                           (Select Max(Sxjh.生效时间) As 生效时间, 安排id
                                            From 挂号安排计划 Sxjh
                                            Where Sxjh.审核时间 Is Not Null And
                                                  d_开始时间 + n_Daycount Between
                                                  Nvl(Sxjh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Nvl(Sxjh.失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.安排id = Jh.安排id
                                            Group By Sxjh.安排id)) Ap, 部门表 Bm, 人员表 Ry, 收费项目目录 Fy
                              Where Ap.科室id = Bm.Id(+) And Ap.医生id = Ry.Id(+) And Ap.项目id = Fy.Id And 排班 Is Not Null) A,
                            病人挂号汇总 Hz, 收费价目 B
                       Where a.号码 = Hz.号码(+) And Hz.日期(+) = Trunc(d_开始时间 + n_Daycount) And a.项目id = b.收费细目id And
                             b.终止日期 > d_开始时间 + n_Daycount And b.执行日期 <= d_开始时间 + n_Daycount And
                             (b.价格等级 = v_普通等级 Or
                             (b.价格等级 Is Null And Not Exists
                              (Select 1
                                From 收费价目
                                Where b.收费细目id = 收费细目id And 价格等级 = v_普通等级 And Sysdate Between 执行日期 And
                                      Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))) Loop
            If v_号类 Is Null Or Instr(',' || v_号类 || ',', ',' || r_排班.号类 || ',') > 0 Then
              v_上班时间 := v_上班时间 || '+' || r_排班.排班;
              n_总已挂数 := n_总已挂数 + r_排班.已挂数;
              n_已挂数   := r_排班.已挂数;
              n_限号数   := r_排班.限号数;
              n_已约数   := r_排班.已约数;
              n_限约数   := r_排班.限约数;
              n_安排id   := Nvl(r_排班.安排id, 0);
              n_计划id   := Nvl(r_排班.计划id, 0);
              v_号码     := r_排班.号码;
              n_安排存在 := 1;
              If v_上班时间 Is Not Null Then
                If v_合作单位 Is Not Null Then
                  If n_计划id <> 0 Then
                    Begin
                      Select 1
                      Into n_合约存在
                      From 合作单位计划控制
                      Where 计划id = n_计划id And 合作单位 = v_合作单位 And Rownum < 2;
                    Exception
                      When Others Then
                        n_合约存在 := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_合约存在
                      From 合作单位安排控制
                      Where 安排id = n_安排id And 合作单位 = v_合作单位 And Rownum < 2;
                    Exception
                      When Others Then
                        n_合约存在 := 0;
                    End;
                  End If;
                End If;
              
                If v_合作单位 Is Null Or Nvl(n_合约存在, 0) = 0 Then
                  If n_计划id <> 0 Then
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位计划控制
                      Where 计划id = n_计划id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null);
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  Else
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位安排控制
                      Where 安排id = n_安排id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null);
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  End If;
                  Begin
                    Select Count(1)
                    Into n_合约已挂数
                    From 病人挂号记录
                    Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 Is Not Null And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                          Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_合约已挂数 := 0;
                  End;
                  If n_合作单位数量 = 0 Then
                    n_合作单位数量 := Null;
                  End If;
                  If Trunc(d_开始时间 + n_Daycount) > Trunc(Sysdate) Then
                    n_剩余数 := Nvl(n_剩余数, 0) + n_限约数 - n_已约数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                  Else
                    n_剩余数 := Nvl(n_剩余数, 0) + n_限号数 - n_已挂数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                  End If;
                Else
                  --合约单位
                  If n_计划id <> 0 Then
                    Begin
                      Select 1
                      Into n_禁用
                      From 合作单位计划控制
                      Where 计划id = n_计划id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_禁用 := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_禁用
                      From 合作单位安排控制
                      Where 安排id = n_安排id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_禁用 := 0;
                    End;
                  End If;
                  If Nvl(n_禁用, 0) = 0 Then
                    Begin
                      Select Count(1)
                      Into n_合约已挂数
                      From 病人挂号记录
                      Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                            Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                    Exception
                      When Others Then
                        n_合约已挂数 := 0;
                    End;
                    If n_计划id <> 0 Then
                      Begin
                        Select Sum(数量)
                        Into n_合作单位数量
                        From 合作单位计划控制
                        Where 计划id = n_计划id And
                              限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                            '周三', '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                      Exception
                        When Others Then
                          n_合作单位数量 := 0;
                      End;
                    Else
                      Begin
                        Select Sum(数量)
                        Into n_合作单位数量
                        From 合作单位安排控制
                        Where 安排id = n_安排id And
                              限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                            '周三', '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                      Exception
                        When Others Then
                          n_合作单位数量 := 0;
                      End;
                    End If;
                    If n_合作单位数量 = 0 Then
                      n_合作单位数量 := Null;
                    End If;
                    n_剩余数 := Nvl(n_剩余数, 0) + Nvl(n_合作单位数量, n_合约已挂数) - Nvl(n_合约已挂数, 0);
                  
                  End If;
                End If;
              End If;
              n_合作单位数量 := 0;
              n_合约存在     := 0;
              n_禁用         := 0;
            End If;
          End Loop;
          v_上班时间 := Substr(v_上班时间, 2);
          If n_安排存在 = 1 Then
            v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                      '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 ||
                      '</YGS>' || '</PB>';
          End If;
        Else
          n_安排存在 := 0;
          v_上班时间 := Null;
          n_总已挂数 := 0;
          n_已挂数   := 0;
          n_剩余数   := 0;
          n_限号数   := 0;
          n_已约数   := 0;
          n_限约数   := 0;
          If v_合作单位 Is Null Then
            --非合作单位
            If Trunc(d_开始时间 + n_Daycount) = Trunc(Sysdate) Then
              --当天挂号
              For r_出诊 In (Select a.已挂数, a.限号数, a.上班时段
                           From 临床出诊记录 A, 临床出诊号源 B
                           Where a.出诊日期 = Trunc(d_开始时间 + n_Daycount) And a.号源id = b.Id And
                                 Sysdate + Nvl(b.预约天数, n_预约天数) + n_补充天数 >= d_开始时间 + n_Daycount And a.医生id = n_医生id And
                                 (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                n_已挂数   := n_已挂数 + Nvl(r_出诊.已挂数, 0);
                n_限号数   := n_限号数 + r_出诊.限号数;
                v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                n_安排存在 := 1;
              End Loop;
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_限号数 - n_已挂数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --预约挂号
              For r_出诊 In (Select a.已约数, a.限号数, a.限约数, a.上班时段
                           From 临床出诊记录 A, 临床出诊号源 B
                           Where a.出诊日期 = Trunc(d_开始时间 + n_Daycount) And a.号源id = b.Id And
                                 Sysdate + Nvl(b.预约天数, n_预约天数) + n_补充天数 >= d_开始时间 + n_Daycount And a.医生id = n_医生id And
                                 (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                n_已挂数   := n_已挂数 + Nvl(r_出诊.已约数, 0);
                n_限号数   := n_限号数 + Nvl(r_出诊.限约数, r_出诊.限号数);
                v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                n_安排存在 := 1;
              End Loop;
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_限号数 - n_已挂数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          Else
            --合作单位
            If Trunc(d_开始时间 + n_Daycount) = Trunc(Sysdate) Then
              For r_出诊 In (Select a.Id, a.已挂数, a.限号数, a.限约数, a.上班时段, a.是否序号控制
                           From 临床出诊记录 A, 临床出诊号源 B
                           Where a.出诊日期 = Trunc(d_开始时间 + n_Daycount) And a.号源id = b.Id And
                                 Sysdate + Nvl(b.预约天数, n_预约天数) + n_补充天数 >= d_开始时间 + n_Daycount And a.医生id = n_医生id And
                                 (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                Begin
                  Select 控制方式
                  Into n_合约模式
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_合约模式 := 4;
                End;
                If n_合约模式 = 1 Or n_合约模式 = 2 Then
                  Select 数量
                  Into n_限号数
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1;
                  If n_合约模式 = 1 Then
                    n_限号数 := n_限号数 * Nvl(r_出诊.限约数, r_出诊.限号数) / 100;
                  End If;
                  Select Count(1)
                  Into n_已挂数
                  From 病人挂号记录
                  Where 出诊记录id = r_出诊.Id And 记录状态 = 1 And 合作单位 = v_合作单位;
                  If r_出诊.限号数 - r_出诊.已挂数 < n_限号数 - n_已挂数 Then
                    n_剩余数 := n_剩余数 + r_出诊.限号数 - r_出诊.已挂数;
                  Else
                    n_剩余数 := n_剩余数 + n_限号数 - n_已挂数;
                  End If;
                  n_总已挂数 := n_总已挂数 + n_已挂数;
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
                If n_合约模式 = 3 Then
                  If r_出诊.是否序号控制 = 0 Then
                    n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已挂数, 0);
                    n_剩余数   := n_剩余数 + r_出诊.限号数 - Nvl(r_出诊.已挂数, 0);
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                    n_安排存在 := 1;
                  Else
                    For r_合作 In (Select 序号, 数量
                                 From 临床出诊挂号控制记录
                                 Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From 临床出诊序号控制
                        Where 记录id = r_出诊.Id And 序号 = r_合作.序号 And Nvl(挂号状态, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_剩余数   := n_剩余数 + 1;
                        n_安排存在 := 1;
                      End If;
                    End Loop;
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                    n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已挂数, 0);
                  End If;
                End If;
                If n_合约模式 = 4 Then
                  n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已挂数, 0);
                  n_剩余数   := n_剩余数 + r_出诊.限号数 - Nvl(r_出诊.已挂数, 0);
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
              End Loop;
            
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --非当天
              For r_出诊 In (Select a.Id, a.已约数, a.已挂数, a.限号数, a.限约数, a.上班时段, a.是否序号控制
                           From 临床出诊记录 A, 临床出诊号源 B
                           Where a.出诊日期 = Trunc(d_开始时间 + n_Daycount) And a.号源id = b.Id And
                                 Sysdate + Nvl(b.预约天数, n_预约天数) + n_补充天数 >= d_开始时间 + n_Daycount And a.医生id = n_医生id And
                                 (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                Begin
                  Select 控制方式
                  Into n_合约模式
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_合约模式 := 4;
                End;
                If n_合约模式 = 1 Or n_合约模式 = 2 Then
                  Select 数量
                  Into n_限号数
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1;
                  If n_合约模式 = 1 Then
                    n_限号数 := n_限号数 * Nvl(r_出诊.限约数, r_出诊.限号数) / 100;
                  End If;
                  Select Count(1)
                  Into n_已挂数
                  From 病人挂号记录
                  Where 出诊记录id = r_出诊.Id And 记录状态 = 1 And 合作单位 = v_合作单位;
                  If Nvl(r_出诊.限约数, r_出诊.限号数) - r_出诊.已约数 < n_限号数 - n_已挂数 Then
                    n_剩余数 := n_剩余数 + Nvl(r_出诊.限约数, r_出诊.限号数) - r_出诊.已约数;
                  Else
                    n_剩余数 := n_剩余数 + n_限号数 - n_已挂数;
                  End If;
                  n_总已挂数 := n_总已挂数 + n_已挂数;
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
                If n_合约模式 = 3 Then
                  If r_出诊.是否序号控制 = 0 Then
                    --分时段非序号控制
                    For r_合作 In (Select 序号, 数量
                                 From 临床出诊挂号控制记录
                                 Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1) Loop
                      Select Count(1)
                      Into n_Exists
                      From 临床出诊序号控制
                      Where 预约顺序号 Is Not Null And 记录id = r_出诊.Id And 序号 = r_合作.序号 And Nvl(挂号状态, 0) <> 0;
                      If r_合作.数量 - n_Exists > 0 Then
                        n_剩余数   := n_剩余数 + r_合作.数量 - n_Exists;
                        n_总已挂数 := n_总已挂数 + n_Exists;
                        n_安排存在 := 1;
                      Else
                        n_总已挂数 := n_总已挂数 + n_Exists;
                        n_安排存在 := 1;
                      End If;
                    End Loop;
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  Else
                    For r_合作 In (Select 序号
                                 From 临床出诊挂号控制记录
                                 Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From 临床出诊序号控制
                        Where 记录id = r_出诊.Id And 序号 = r_合作.序号 And Nvl(挂号状态, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_剩余数   := n_剩余数 + 1;
                        n_安排存在 := 1;
                      Else
                        n_总已挂数 := n_总已挂数 + 1;
                        n_安排存在 := 1;
                      End If;
                    End Loop;
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  End If;
                End If;
                If n_合约模式 = 4 Then
                  n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已约数, 0);
                  n_剩余数   := n_剩余数 + r_出诊.限约数 - Nvl(r_出诊.已约数, 0);
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
              End Loop;
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          End If;
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    Else
      --通过科室+医生查找
      While (n_Daycount < n_查询天数) Loop
        If Trunc(d_开始时间 + n_Daycount) < To_Date(Substr(v_启用时间, 1, Instr(v_启用时间, ' ') - 1), 'yyyy-mm-dd') Then
          v_上班时间 := Null;
          n_总已挂数 := 0;
          n_已挂数   := 0;
          n_剩余数   := 0;
          n_限号数   := 0;
          n_已约数   := 0;
          n_限约数   := 0;
          n_安排存在 := 0;
          For r_排班 In (Select d_开始时间 + n_Daycount As 日期, a.排班, a.限号数, a.限约数, Nvl(Hz.已挂数, 0) As 已挂数, Nvl(Hz.已约数, 0) As 已约数,
                              a.安排id, a.计划id, a.号码, a.号类
                       
                       From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Nvl(Ap.医生id, 0) As 医生id, Ry.专业技术职务 As 职称,
                                     Ap.号码, Ap.安排id, Ap.计划id, Ap.排班, Ap.项目id, Fy.名称 As 项目名称, Ap.序号控制, Ap.限号数, Ap.限约数
                              
                              From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Ap.医生id, Ap.号码, Ap.Id As 安排id, 0 As 计划id,
                                            Ap.项目id, Nvl(Ap.序号控制, 0) As 序号控制,
                                            Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Ap.周日, '2', Ap.周一, '3', Ap.周二, '4',
                                                    Ap.周三, '5', Ap.周四, '6', Ap.周五, '7', Ap.周六, Null) As 排班,
                                            Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                     From 挂号安排 Ap, 部门表 Bm, 挂号安排限制 Xz
                                     Where Ap.科室id = Bm.Id(+) And Sysdate + Nvl(Ap.预约天数, n_预约天数) >= d_开始时间 + n_Daycount And
                                           Ap.医生id = n_医生id And Ap.科室id = n_科室id And Ap.停用日期 Is Null And
                                           d_开始时间 + n_Daycount Between Nvl(Ap.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Nvl(Ap.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.安排id(+) = Ap.Id And
                                           Xz.限制项目(+) =
                                           Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                                  '周三', '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                      (Select Rownum
                                            From 挂号安排停用状态 Ty
                                            Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                           Not Exists (Select Rownum
                                            From 挂号安排计划 Jh
                                            Where Jh.安排id = Ap.Id And Jh.审核时间 Is Not Null And
                                                  d_开始时间 + n_Daycount Between
                                                  Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Jh.失效时间)
                                     Union All
                                     Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Jh.医生姓名, Jh.医生id, Ap.号码, Ap.Id As 安排id,
                                            Jh.Id As 计划id, Jh.项目id, Nvl(Jh.序号控制, 0) As 序号控制,
                                            Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', Jh.周日, '2', Jh.周一, '3', Jh.周二, '4',
                                                    Jh.周三, '5', Jh.周四, '6', Jh.周五, '7', Jh.周六, Null) As 排班,
                                            Nvl(Xz.限约数, Xz.限号数) As 限约数, Nvl(Xz.限号数, 0) As 限号数
                                     From 挂号安排计划 Jh, 挂号安排 Ap, 部门表 Bm, 挂号计划限制 Xz
                                     Where Jh.安排id = Ap.Id And Sysdate + Nvl(Ap.预约天数, n_预约天数) >= d_开始时间 + n_Daycount And
                                           Ap.科室id = Bm.Id(+) And Ap.停用日期 Is Null And Jh.医生id = n_医生id And Ap.科室id = n_科室id And
                                           d_开始时间 + n_Daycount Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Jh.失效时间 And Xz.计划id(+) = Jh.Id And
                                           Xz.限制项目(+) =
                                           Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                                  '周三', '5', '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                      (Select Rownum
                                            From 挂号安排停用状态 Ty
                                            Where Ty.安排id = Ap.Id And d_开始时间 + n_Daycount Between Ty.开始停止时间 And Ty.结束停止时间) And
                                           (Jh.生效时间, Jh.安排id) =
                                           (Select Max(Sxjh.生效时间) As 生效时间, 安排id
                                            From 挂号安排计划 Sxjh
                                            Where Sxjh.审核时间 Is Not Null And
                                                  d_开始时间 + n_Daycount Between
                                                  Nvl(Sxjh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Sxjh.失效时间 And Sxjh.安排id = Jh.安排id
                                            Group By Sxjh.安排id)) Ap, 部门表 Bm, 人员表 Ry, 收费项目目录 Fy
                              Where Ap.科室id = Bm.Id(+) And Ap.医生id = Ry.Id(+) And Ap.项目id = Fy.Id And 排班 Is Not Null) A,
                            病人挂号汇总 Hz, 收费价目 B
                       Where a.号码 = Hz.号码(+) And Hz.日期(+) = Trunc(d_开始时间 + n_Daycount) And a.项目id = b.收费细目id And
                             b.终止日期 > d_开始时间 + n_Daycount And b.执行日期 <= d_开始时间 + n_Daycount And
                             (b.价格等级 = v_普通等级 Or
                             (b.价格等级 Is Null And Not Exists
                              (Select 1
                                From 收费价目
                                Where b.收费细目id = 收费细目id And 价格等级 = v_普通等级 And Sysdate Between 执行日期 And
                                      Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))) Loop
            If v_号类 Is Null Or Instr(',' || v_号类 || ',', ',' || r_排班.号类 || ',') > 0 Then
              v_上班时间 := v_上班时间 || '+' || r_排班.排班;
              n_总已挂数 := n_总已挂数 + r_排班.已挂数;
              n_已挂数   := r_排班.已挂数;
              n_限号数   := r_排班.限号数;
              n_已约数   := r_排班.已约数;
              n_限约数   := r_排班.限约数;
              n_安排id   := Nvl(r_排班.安排id, 0);
              n_计划id   := Nvl(r_排班.计划id, 0);
              v_号码     := r_排班.号码;
              n_安排存在 := 1;
            
              If v_上班时间 Is Not Null Then
                If v_合作单位 Is Not Null Then
                  If n_计划id <> 0 Then
                    Begin
                      Select 1
                      Into n_合约存在
                      From 合作单位计划控制
                      Where 计划id = n_计划id And 合作单位 = v_合作单位 And Rownum < 2;
                    Exception
                      When Others Then
                        n_合约存在 := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_合约存在
                      From 合作单位安排控制
                      Where 安排id = n_安排id And 合作单位 = v_合作单位 And Rownum < 2;
                    Exception
                      When Others Then
                        n_合约存在 := 0;
                    End;
                  End If;
                End If;
              
                If v_合作单位 Is Null Or Nvl(n_合约存在, 0) = 0 Then
                  If n_计划id <> 0 Then
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位计划控制
                      Where 计划id = n_计划id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null);
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  Else
                    Begin
                      Select Sum(数量)
                      Into n_合作单位数量
                      From 合作单位安排控制
                      Where 安排id = n_安排id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null);
                    Exception
                      When Others Then
                        n_合作单位数量 := 0;
                    End;
                  End If;
                  Begin
                    Select Count(1)
                    Into n_合约已挂数
                    From 病人挂号记录
                    Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 Is Not Null And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                          Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_合约已挂数 := 0;
                  End;
                  If n_合作单位数量 = 0 Then
                    n_合作单位数量 := Null;
                  End If;
                  If Trunc(d_开始时间 + n_Daycount) > Trunc(Sysdate) Then
                    n_剩余数 := Nvl(n_剩余数, 0) + n_限约数 - n_已约数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                  Else
                    n_剩余数 := Nvl(n_剩余数, 0) + n_限号数 - n_已挂数 - Nvl(n_合作单位数量, n_合约已挂数) + Nvl(n_合约已挂数, 0);
                  End If;
                Else
                  --合约单位
                  If n_计划id <> 0 Then
                    Begin
                      Select 1
                      Into n_禁用
                      From 合作单位计划控制
                      Where 计划id = n_计划id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_禁用 := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_禁用
                      From 合作单位安排控制
                      Where 安排id = n_安排id And
                            限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                          '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位 And 数量 = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_禁用 := 0;
                    End;
                  End If;
                  If Nvl(n_禁用, 0) = 0 Then
                    Begin
                      Select Count(1)
                      Into n_合约已挂数
                      From 病人挂号记录
                      Where 号别 = v_号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_开始时间 + n_Daycount) And
                            Trunc(d_开始时间 + n_Daycount + 1) - 1 / 60 / 60 / 24;
                    Exception
                      When Others Then
                        n_合约已挂数 := 0;
                    End;
                    If n_计划id <> 0 Then
                      Begin
                        Select Sum(数量)
                        Into n_合作单位数量
                        From 合作单位计划控制
                        Where 计划id = n_计划id And
                              限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                            '周三', '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                      Exception
                        When Others Then
                          n_合作单位数量 := 0;
                      End;
                    Else
                      Begin
                        Select Sum(数量)
                        Into n_合作单位数量
                        From 合作单位安排控制
                        Where 安排id = n_安排id And
                              限制项目 = Decode(To_Char(d_开始时间 + n_Daycount, 'D'), '1', '周日', '2', '周一', '3', '周二', '4',
                                            '周三', '5', '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
                      Exception
                        When Others Then
                          n_合作单位数量 := 0;
                      End;
                    End If;
                    If n_合作单位数量 = 0 Then
                      n_合作单位数量 := Null;
                    End If;
                    n_剩余数 := Nvl(n_剩余数, 0) + Nvl(n_合作单位数量, n_合约已挂数) - Nvl(n_合约已挂数, 0);
                  
                  End If;
                End If;
              End If;
              n_合作单位数量 := 0;
              n_合约存在     := 0;
              n_禁用         := 0;
            End If;
          End Loop;
          v_上班时间 := Substr(v_上班时间, 2);
          If n_安排存在 = 1 Then
            v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                      '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 ||
                      '</YGS>' || '</PB>';
          End If;
        Else
          n_安排存在 := 0;
          v_上班时间 := Null;
          n_总已挂数 := 0;
          n_已挂数   := 0;
          n_剩余数   := 0;
          n_限号数   := 0;
          n_已约数   := 0;
          n_限约数   := 0;
          If v_合作单位 Is Null Then
            --非合作单位
            If Trunc(d_开始时间 + n_Daycount) = Trunc(Sysdate) Then
              --当天挂号
              For r_出诊 In (Select a.已挂数, a.限号数, a.上班时段
                           From 临床出诊记录 A, 临床出诊号源 B
                           Where a.出诊日期 = Trunc(d_开始时间 + n_Daycount) And a.号源id = b.Id And
                                 Sysdate + Nvl(b.预约天数, n_预约天数) + n_补充天数 >= d_开始时间 + n_Daycount And a.医生id = n_医生id And
                                 a.科室id = n_科室id And (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                n_已挂数   := n_已挂数 + Nvl(r_出诊.已挂数, 0);
                n_限号数   := n_限号数 + r_出诊.限号数;
                v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                n_安排存在 := 1;
              End Loop;
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_限号数 - n_已挂数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --预约挂号
              For r_出诊 In (Select a.已约数, a.限号数, a.限约数, a.上班时段
                           From 临床出诊记录 A, 临床出诊号源 B
                           Where a.出诊日期 = Trunc(d_开始时间 + n_Daycount) And a.号源id = b.Id And
                                 Sysdate + Nvl(b.预约天数, n_预约天数) + n_补充天数 >= d_开始时间 + n_Daycount And a.医生id = n_医生id And
                                 a.科室id = n_科室id And (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                n_已挂数   := n_已挂数 + Nvl(r_出诊.已约数, 0);
                n_限号数   := n_限号数 + Nvl(r_出诊.限约数, r_出诊.限号数);
                v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                n_安排存在 := 1;
              End Loop;
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_限号数 - n_已挂数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          Else
            --合作单位
            If Trunc(d_开始时间 + n_Daycount) = Trunc(Sysdate) Then
              For r_出诊 In (Select a.Id, a.已挂数, a.限号数, a.限约数, a.上班时段, a.是否序号控制
                           From 临床出诊记录 A, 临床出诊号源 B
                           Where a.出诊日期 = Trunc(d_开始时间 + n_Daycount) And a.号源id = b.Id And
                                 Sysdate + Nvl(b.预约天数, n_预约天数) + n_补充天数 >= d_开始时间 + n_Daycount And a.医生id = n_医生id And
                                 a.科室id = n_科室id And (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                Begin
                  Select 控制方式
                  Into n_合约模式
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_合约模式 := 4;
                End;
                If n_合约模式 = 1 Or n_合约模式 = 2 Then
                  Select 数量
                  Into n_限号数
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1;
                  If n_合约模式 = 1 Then
                    n_限号数 := n_限号数 * Nvl(r_出诊.限约数, r_出诊.限号数) / 100;
                  End If;
                  Select Count(1)
                  Into n_已挂数
                  From 病人挂号记录
                  Where 出诊记录id = r_出诊.Id And 记录状态 = 1 And 合作单位 = v_合作单位;
                  If r_出诊.限号数 - r_出诊.已挂数 < n_限号数 - n_已挂数 Then
                    n_剩余数 := n_剩余数 + r_出诊.限号数 - r_出诊.已挂数;
                  Else
                    n_剩余数 := n_剩余数 + n_限号数 - n_已挂数;
                  End If;
                  n_总已挂数 := n_总已挂数 + n_已挂数;
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
                If n_合约模式 = 3 Then
                  If r_出诊.是否序号控制 = 0 Then
                    n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已挂数, 0);
                    n_剩余数   := n_剩余数 + r_出诊.限号数 - Nvl(r_出诊.已挂数, 0);
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                    n_安排存在 := 1;
                  Else
                    For r_合作 In (Select 序号, 数量
                                 From 临床出诊挂号控制记录
                                 Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From 临床出诊序号控制
                        Where 记录id = r_出诊.Id And 序号 = r_合作.序号 And Nvl(挂号状态, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_剩余数   := n_剩余数 + 1;
                        n_安排存在 := 1;
                      End If;
                    End Loop;
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                    n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已挂数, 0);
                  End If;
                End If;
                If n_合约模式 = 4 Then
                  n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已挂数, 0);
                  n_剩余数   := n_剩余数 + r_出诊.限号数 - Nvl(r_出诊.已挂数, 0);
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
              End Loop;
            
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --非当天
              For r_出诊 In (Select a.Id, a.已约数, a.已挂数, a.限号数, a.限约数, a.上班时段, a.是否序号控制
                           From 临床出诊记录 A, 临床出诊号源 B
                           Where a.出诊日期 = Trunc(d_开始时间 + n_Daycount) And a.号源id = b.Id And
                                 Sysdate + Nvl(b.预约天数, n_预约天数) + n_补充天数 >= d_开始时间 + n_Daycount And a.医生id = n_医生id And
                                 a.科室id = n_科室id And (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.停诊终止时间, a.开始时间)) And
                                 Nvl(a.是否锁定, 0) = 0 And Exists
                            (Select 1
                                  From 临床出诊安排 M, 临床出诊表 N
                                  Where m.Id = a.安排id And m.出诊id = n.Id And n.发布时间 Is Not Null)) Loop
                Begin
                  Select 控制方式
                  Into n_合约模式
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_合约模式 := 4;
                End;
                If n_合约模式 = 1 Or n_合约模式 = 2 Then
                  Select 数量
                  Into n_限号数
                  From 临床出诊挂号控制记录
                  Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1;
                  If n_合约模式 = 1 Then
                    n_限号数 := n_限号数 * Nvl(r_出诊.限约数, r_出诊.限号数) / 100;
                  End If;
                  Select Count(1)
                  Into n_已挂数
                  From 病人挂号记录
                  Where 出诊记录id = r_出诊.Id And 记录状态 = 1 And 合作单位 = v_合作单位;
                  If Nvl(r_出诊.限约数, r_出诊.限号数) - r_出诊.已约数 < n_限号数 - n_已挂数 Then
                    n_剩余数 := n_剩余数 + Nvl(r_出诊.限约数, r_出诊.限号数) - r_出诊.已约数;
                  Else
                    n_剩余数 := n_剩余数 + n_限号数 - n_已挂数;
                  End If;
                  n_总已挂数 := n_总已挂数 + n_已挂数;
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
                If n_合约模式 = 3 Then
                  If r_出诊.是否序号控制 = 0 Then
                    --分时段非序号控制
                    For r_合作 In (Select 序号, 数量
                                 From 临床出诊挂号控制记录
                                 Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1) Loop
                      Select Count(1)
                      Into n_Exists
                      From 临床出诊序号控制
                      Where 预约顺序号 Is Not Null And 记录id = r_出诊.Id And 序号 = r_合作.序号 And Nvl(挂号状态, 0) <> 0;
                      If r_合作.数量 - n_Exists > 0 Then
                        n_剩余数   := n_剩余数 + r_合作.数量 - n_Exists;
                        n_总已挂数 := n_总已挂数 + n_Exists;
                        n_安排存在 := 1;
                      Else
                        n_总已挂数 := n_总已挂数 + n_Exists;
                        n_安排存在 := 1;
                      End If;
                    End Loop;
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  Else
                    For r_合作 In (Select 序号
                                 From 临床出诊挂号控制记录
                                 Where 记录id = r_出诊.Id And 类型 = 1 And 名称 = v_合作单位 And 性质 = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From 临床出诊序号控制
                        Where 记录id = r_出诊.Id And 序号 = r_合作.序号 And Nvl(挂号状态, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_剩余数   := n_剩余数 + 1;
                        n_安排存在 := 1;
                      Else
                        n_总已挂数 := n_总已挂数 + 1;
                        n_安排存在 := 1;
                      End If;
                    End Loop;
                    v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  End If;
                End If;
                If n_合约模式 = 4 Then
                  n_总已挂数 := n_总已挂数 + Nvl(r_出诊.已约数, 0);
                  n_剩余数   := n_剩余数 + r_出诊.限约数 - Nvl(r_出诊.已约数, 0);
                  v_上班时间 := v_上班时间 || '+' || r_出诊.上班时段;
                  n_安排存在 := 1;
                End If;
              End Loop;
              If v_上班时间 Is Not Null Then
                v_上班时间 := Substr(v_上班时间, 2);
              End If;
              If n_安排存在 = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_开始时间 + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_剩余数 || '</SYHS>' || '<SBSJ>' || v_上班时间 || '</SBSJ>' || '<YGS>' || n_总已挂数 ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          End If;
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    End If;
    v_Temp := '<PBLIST>' || v_Temp || '</PBLIST>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Docarrange;
/

--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
Create Or Replace Procedure Zl_Third_Getnolist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:获取号源列表(简易模式)
  --入参:Xml_In:
  --<IN>
  --  <RQ>日期</RQ>
  --  <KSID>科室ID</KSID>
  --  <YSID>医生ID</YSID>
  --  <YSXM>医生姓名</YSXM>
  --  <HZDW>支付宝</HZDW>    //合作单位，传入了的时候，只取合作单位的号;为空时，只取非合作单位的号
  --  <FKFS>付款方式</FKFS>  
  --  <SJJG>60</SJJG>     //时间间隔,不传则返回序号时段
  --  <ZD></ZD>           //站点
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  --  <GROUP>
  --    <RQ>日期</RQ>
  --    <HBLIST>
  --     <HB>
  --        <CZJLID>1</CZJLID>     //出诊记录ID
  --        <APID></APID>          //挂号安排ID
  --        <JHID></JHID>          //挂号计划ID
  --        <HM>235</HM>       //号码
  --        <YSID>549</YSID>      //医生ID
  --        <YS>张锐</YS>       //医生姓名
  --        <KSID>123</KSID>   //科室ID
  --        <KSMC>内科</KSMC>   //科室名称
  --        <ZC>主治医师</ZC> //职称
  --        <XMID>10086<XMID> //挂号项目的ID
  --        <XMMC>挂号费</XMMC> //挂号项目的名称
  --        <YGHS>0</YGHS>      //已挂号数
  --        <SYHS>99</SYHS>   //剩余号数
  --        <PRICE>15</PRICE>      //价格
  --        <HL>普通</HL>       //挂号类型
  --        <HCXH>1</HCXH>    //是否存在缓冲序号时间段，1-存在 0或者空-不存在
  --        <FSD>0</FSD>      //是否分时段
  --        <FWMC>白天</FWMC>     //号别时段
  --        <HBTIME>(08:00-17:59)</HBTIME> //可挂时间
  --     <SPANLIST>
  --            <SPAN>
  --                  <SJD></SJD>          //时间段,格式:hh24:mi-hh24:mi
  --                  <GHZS></GHZS>      //时段挂号总数
  --                  <SL></SL>      //剩余数量
  --            </SPAN>
  --            ……
  --          </SPANLIST>
  --      </HB>
  --      <HB>
  --      ……
  --      </HB>
  --    </HBLIST>
  --  </GROUP>
  --</OUTPUT>
  --
  --“序号时段”<SPANLIST>和“剩余号数”<SYHS>节点说明：
  --  1.出诊表排班模式：
  --    1.1启用分时段，同时启用序号控制
  --      1.1.1传入合作单位
  --          1）当日：
  --                  序号时段：按比例或按总量时，号源剩余的可挂号时段；按序号时，分配给该合作单位剩余的可预约时段
  --                  剩余号数：分配给该合作单位剩余的可预约数量
  --          2）当日以后：
  --                  序号时段：按比例或按总量时，号源剩余的可预约时段；按序号时，分配给该合作单位剩余的可预约时段
  --                  剩余号数：分配给该合作单位剩余的可预约数量
  --      1.1.2不传入合作单位
  --          1）当日：
  --                  序号时段：号源剩余的可挂号时段
  --                  剩余号数：号源剩余的可挂号数量
  --          2）当日以后：
  --                  序号时段：号源剩余的可预约时段
  --                  剩余号数：号源剩余的可预约数量
  --------------------------------------------------------------------------------------------------
  x_Templet Xmltype; --模板XML

  d_日期     临床出诊记录.出诊日期%Type;
  n_科室id   挂号安排.科室id%Type;
  n_医生id   挂号安排.医生id%Type;
  v_医生姓名 挂号安排.医生姓名%Type;
  v_合作单位 挂号合作单位.名称%Type;
  n_时间间隔 挂号安排.默认时段间隔%Type;

  v_挂号模式 Varchar2(500);
  n_挂号模式 Number(3);
  d_启用时间 Date;
  n_预约天数 挂号安排.预约天数%Type;
  n_补充天数 挂号安排.预约天数%Type;

  v_剩余数量 挂号安排时段.限制数量%Type;
  n_禁用     Number(3);
  v_Temp     Varchar2(32767); --临时XML
  c_Xmlmain  Clob; --临时XML 
  v_Xmlmain  Clob; --临时XML 

  d_时段开始 挂号安排时段.开始时间%Type;
  d_时段结束 挂号安排时段.结束时间%Type;
  n_时段总数 挂号安排时段.限制数量%Type;
  n_时段剩余 挂号安排时段.限制数量%Type;
  n_时段已挂 挂号安排时段.限制数量%Type;

  v_星期         挂号安排限制.限制项目%Type;
  v_时间段       Varchar2(100);
  n_分时段       Number(3);
  n_单个剩余     Number(5);
  n_已挂数       Number(5);
  n_合约已挂数   Number(5);
  n_合计金额     收费价目.现价%Type;
  n_合约总数量   Number(5);
  n_合约剩余数量 Number(5);
  n_最大可用数量 Number(5);
  n_合约模式     Number(3); --合约模式:1-合约单位限数量模式 0-合约单位指定序号模式
  n_非合约       Number(3);
  n_是否预留     Number(3);

  d_开始时间 临床出诊记录.开始时间%Type;
  d_终止时间 临床出诊记录.终止时间%Type;
  d_加号时间 临床出诊记录.开始时间%Type;

  n_缓冲序号 Number(3);
  n_时段数量 Number(5);
  n_预留数量 Number(5);
  n_特殊预约 Number(3);
  v_Timetemp Varchar2(100);
  n_Exists   Number(5);

  v_普通等级   Varchar2(100);
  v_Pricegrade Varchar2(500);
  v_站点       部门表.站点%Type;
  v_付款方式   医疗付款方式.名称%Type;
  v_方式       Varchar2(20);

  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  --获取序号时段XML
  Function Gettimexml
  (
    时段开始_In 临床出诊序号控制.开始时间%Type,
    时段结束_In 临床出诊序号控制.终止时间%Type,
    时段总数_In 临床出诊序号控制.数量%Type,
    时段剩余_In 临床出诊序号控制.数量%Type
  ) Return Varchar2 Is
    v_Temp Varchar2(4000);
  Begin
    v_Temp := '';
    v_Temp := v_Temp || '<SPAN>';
    v_Temp := v_Temp || '<SJD>' || To_Char(时段开始_In, 'hh24:mi:ss') || '-' || To_Char(时段结束_In, 'hh24:mi:ss') || '</SJD>';
    v_Temp := v_Temp || '<GHZS>' || 时段总数_In || '</GHZS>';
    v_Temp := v_Temp || '<SL>' || 时段剩余_In || '</SL>';
    v_Temp := v_Temp || '</SPAN>';
    Return v_Temp;
  End;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select To_Date(Extractvalue(Value(A), 'IN/RQ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/KSID'),
         Extractvalue(Value(A), 'IN/YSID'), Extractvalue(Value(A), 'IN/YSXM'), Extractvalue(Value(A), 'IN/HZDW'),
         Extractvalue(Value(A), 'IN/SJJG'), Extractvalue(Value(A), 'IN/ZD'), Extractvalue(Value(A), 'IN/FKFS')
  Into d_日期, n_科室id, n_医生id, v_医生姓名, v_合作单位, n_时间间隔, v_站点, v_方式
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  --日期节点为空的情况
  d_日期 := Nvl(d_日期, Trunc(Sysdate));

  If v_方式 Is Null Then
    Select Max(名称) Into v_付款方式 From 医疗付款方式 Where 缺省标志 = 1;
  Else
    Select Nvl(Max(名称), v_方式) Into v_付款方式 From 医疗付款方式 Where 编码 = v_方式;
  End If;
  v_Pricegrade := Zl_Get_Pricegrade(v_站点, Null, Null, v_付款方式);
  v_普通等级   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);

  v_挂号模式 := zl_GetSysParameter('挂号排班模式') || '|||';
  n_挂号模式 := To_Number(Substr(v_挂号模式, 1, Instr(v_挂号模式, '|') - 1));
  If n_挂号模式 = 1 Then
    v_挂号模式 := Substr(v_挂号模式, Instr(v_挂号模式, '|') + 1);
    v_Temp     := Substr(v_挂号模式, 1, Instr(v_挂号模式, '|') - 1);
    d_启用时间 := To_Date(Nvl(v_Temp, '1900-01-01'), 'yyyy-mm-dd hh24:mi:ss');
    If d_日期 < d_启用时间 Then
      n_挂号模式 := 0;
    End If;
  End If;

  n_预约天数 := Nvl(zl_GetSysParameter(66), 7);
  c_Xmlmain  := '';

  --===========================================================================================
  --计划排班模式 
  --===========================================================================================
  If n_挂号模式 = 0 Then
    Select Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)
    Into v_星期
    From Dual;
    n_合约剩余数量 := 0;
  
    For r_No In (Select a.科室id, a.号类, a.科室名称, a.医生姓名, a.医生id, a.职称, a.号码, a.安排id, a.计划id, a.排班, a.项目id, a.项目名称, a.序号控制,
                        a.限号数, a.限约数, a.预约天数, Nvl(Hz.已挂数, 0) As 已挂数, Nvl(Hz.已约数, 0) As 已约数, Nvl(Hz.其中已接收, 0) As 已接收,
                        Sum(b.现价) As 价格
                 From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Nvl(Ap.医生id, 0) As 医生id, Ry.专业技术职务 As 职称, Ap.号码,
                               Ap.安排id, Ap.计划id, Ap.排班, Ap.项目id, Fy.名称 As 项目名称, Ap.序号控制, Ap.限号数, Ap.限约数, Ap.预约天数
                        From (Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Ap.医生姓名, Ap.医生id, Ap.号码, Ap.Id As 安排id, 0 As 计划id,
                                      Ap.项目id, Nvl(Ap.序号控制, 0) As 序号控制,
                                      Decode(To_Char(d_日期, 'D'), '1', Ap.周日, '2', Ap.周一, '3', Ap.周二, '4', Ap.周三, '5', Ap.周四,
                                              '6', Ap.周五, '7', Ap.周六, Null) As 排班, Xz.限约数, Xz.限号数,
                                      Nvl(Ap.预约天数, n_预约天数) As 预约天数
                               From 挂号安排 Ap, 部门表 Bm, 挂号安排限制 Xz
                               Where Ap.科室id = Bm.Id(+) And Decode(Nvl(n_科室id, 0), 0, 0, Ap.科室id) = Nvl(n_科室id, 0) And
                                     Decode(Nvl(n_医生id, 0), 0, 0, Ap.医生id) = Nvl(n_医生id, 0) And
                                     Decode(Nvl(v_医生姓名, '-'), '-', '-', Ap.医生姓名) = Nvl(v_医生姓名, '-') And Ap.停用日期 Is Null And
                                     d_日期 Between Nvl(Ap.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                     Nvl(Ap.终止时间, To_Date('3000 - 01 - 01', 'YYYY-MM-DD')) And Xz.安排id(+) = Ap.Id And
                                     Xz.限制项目(+) = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                         '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                (Select Rownum
                                      From 挂号安排停用状态 Ty
                                      Where Ty.安排id = Ap.Id And d_日期 Between Ty.开始停止时间 And Ty.结束停止时间) And Not Exists
                                (Select Rownum
                                      From 挂号安排计划 Jh
                                      Where Jh.安排id = Ap.Id And Jh.审核时间 Is Not Null And
                                            d_日期 Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                            Jh.失效时间)
                               Union All
                               Select Ap.科室id, Ap.号类, Bm.名称 As 科室名称, Jh.医生姓名, Jh.医生id, Ap.号码, Ap.Id As 安排id, Jh.Id As 计划id,
                                      Jh.项目id, Nvl(Jh.序号控制, 0) As 序号控制,
                                      Decode(To_Char(d_日期, 'D'), '1', Jh.周日, '2', Jh.周一, '3', Jh.周二, '4', Jh.周三, '5', Jh.周四,
                                              '6', Jh.周五, '7', Jh.周六, Null) As 排班, Xz.限约数, Xz.限号数,
                                      Nvl(Ap.预约天数, n_预约天数) As 预约天数
                               From 挂号安排计划 Jh, 挂号安排 Ap, 部门表 Bm, 挂号计划限制 Xz
                               Where Jh.安排id = Ap.Id And Ap.科室id = Bm.Id(+) And Ap.停用日期 Is Null And
                                     Decode(Nvl(n_科室id, 0), 0, 0, Ap.科室id) = Nvl(n_科室id, 0) And
                                     Decode(Nvl(n_医生id, 0), 0, 0, Jh.医生id) = Nvl(n_医生id, 0) And
                                     Decode(Nvl(v_医生姓名, '-'), '-', '-', Jh.医生姓名) = Nvl(v_医生姓名, '-') And
                                     d_日期 Between Nvl(Jh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                     Jh.失效时间 And Xz.计划id(+) = Jh.Id And
                                     Xz.限制项目(+) = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                         '周四', '6', '周五', '7', '周六', Null) And Not Exists
                                (Select Rownum
                                      From 挂号安排停用状态 Ty
                                      Where Ty.安排id = Ap.Id And d_日期 Between Ty.开始停止时间 And Ty.结束停止时间) And
                                     (Jh.生效时间, Jh.安排id) =
                                     (Select Max(Sxjh.生效时间) As 生效时间, 安排id
                                      From 挂号安排计划 Sxjh
                                      Where Sxjh.审核时间 Is Not Null And
                                            d_日期 Between Nvl(Sxjh.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                            Sxjh.失效时间 And Sxjh.安排id = Jh.安排id
                                      Group By Sxjh.安排id)) Ap, 部门表 Bm, 人员表 Ry, 收费项目目录 Fy
                        Where Ap.科室id = Bm.Id(+) And Ap.医生id = Ry.Id(+) And Ap.项目id = Fy.Id And 排班 Is Not Null) A,
                      病人挂号汇总 Hz, 收费价目 B
                 Where a.号码 = Hz.号码(+) And Hz.日期(+) = Trunc(d_日期) And a.项目id = b.收费细目id And
                       Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-Mm-DD')) > d_日期 And b.执行日期 <= d_日期 And
                       (b.价格等级 = v_普通等级 Or (b.价格等级 Is Null And Not Exists
                        (Select 1
                                             From 收费价目
                                             Where 收费细目id = b.收费细目id And 价格等级 = v_普通等级 And d_日期 Between 执行日期 And
                                                   Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                 Group By a.科室id, a.号类, a.科室名称, a.医生姓名, a.医生id, a.职称, a.号码, a.安排id, a.计划id, a.排班, a.项目id, a.项目名称, a.序号控制,
                          a.限号数, a.限约数, a.预约天数, Nvl(Hz.已挂数, 0), Nvl(Hz.已约数, 0), Nvl(Hz.其中已接收, 0)) Loop
      Zl_挂号序号状态_Delete(1, r_No.号码);
      If Sysdate + Nvl(r_No.预约天数, n_预约天数) >= d_日期 Then
        If r_No.计划id <> 0 Then
          Select Sign(Count(Rownum))
          Into n_分时段
          From 挂号安排计划 Jh, 挂号计划时段 Sd
          Where Jh.Id = Sd.计划id And Jh.Id = r_No.计划id And
                Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                               '周六', Null) And Rownum < 2;
        Else
          Select Sign(Count(Rownum))
          Into n_分时段
          From 挂号安排 Ap, 挂号安排时段 Sd
          Where Ap.Id = Sd.安排id And Ap.Id = r_No.安排id And
                Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                               '周六', Null) And Rownum < 2;
        End If;
        --分时段不序号控制当天号为普通号
        If Trunc(Sysdate) = Trunc(d_日期) And n_分时段 = 1 And r_No.序号控制 = 0 Then
          n_分时段 := 0;
        End If;
        If n_分时段 = 0 Then
          v_Temp := '';
          If v_合作单位 Is Not Null And r_No.序号控制 = 1 Then
            If r_No.计划id <> 0 Then
              Select Nvl(Sum(数量), 0)
              Into n_合约总数量
              From 合作单位计划控制
              Where 计划id = r_No.计划id And 合作单位 = v_合作单位 And
                    限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null);
              Select Count(1)
              Into n_合约模式
              From 合作单位计划控制
              Where 计划id = r_No.计划id And 合作单位 = v_合作单位 And
                    限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null) And 序号 = 0;
            Else
              Select Nvl(Sum(数量), 0)
              Into n_合约总数量
              From 合作单位安排控制
              Where 安排id = r_No.安排id And 合作单位 = v_合作单位 And
                    限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null);
              Select Count(1)
              Into n_合约模式
              From 合作单位安排控制
              Where 安排id = r_No.安排id And 合作单位 = v_合作单位 And
                    限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null) And 序号 = 0;
            End If;
            If n_合约模式 = 0 Then
              If r_No.计划id <> 0 Then
                Select Count(1)
                Into n_合约已挂数
                From 病人挂号记录 A
                Where 号别 = r_No.号码 And 记录状态 = 1 And 发生时间 Between Trunc(d_日期) And Trunc(d_日期 + 1) - 1 / 60 / 60 / 24 And
                      Exists
                 (Select 1
                       From 合作单位计划控制
                       Where 计划id = r_No.计划id And 合作单位 = v_合作单位 And
                             限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6',
                                           '周五', '7', '周六', Null) And 序号 = a.号序 And 数量 <> 0);
              Else
                Select Count(1)
                Into n_合约已挂数
                From 病人挂号记录 A
                Where 号别 = r_No.号码 And 记录状态 = 1 And 发生时间 Between Trunc(d_日期) And Trunc(d_日期 + 1) - 1 / 60 / 60 / 24 And
                      Exists
                 (Select 1
                       From 合作单位安排控制
                       Where 安排id = r_No.安排id And 合作单位 = v_合作单位 And
                             限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6',
                                           '周五', '7', '周六', Null) And 序号 = a.号序 And 数量 <> 0);
              End If;
            Else
              Select Count(1)
              Into n_合约已挂数
              From 病人挂号记录
              Where 号别 = r_No.号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_日期) And
                    Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
            End If;
            If n_合约总数量 = 0 Then
              n_合约剩余数量 := 0;
            Else
              n_合约剩余数量 := n_合约总数量 - n_合约已挂数;
              If n_合约剩余数量 > (Nvl(r_No.限号数, 0) - r_No.已挂数) Then
                n_合约剩余数量 := Nvl(r_No.限号数, 0) - r_No.已挂数;
              End If;
            End If;
          End If;
        Else
          v_Temp := '<SPANLIST>';
          If r_No.计划id <> 0 Then
            Select To_Date(To_Char(d_日期, 'yyyy-mm-dd') || To_Char(Max(结束时间), 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
            Into d_加号时间
            From 挂号计划时段
            Where 计划id = r_No.计划id And 星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                   '周四', '6', '周五', '7', '周六', Null);
            If r_No.序号控制 = 1 Then
              If Trunc(d_日期) = Trunc(Sysdate) Then
                n_特殊预约 := 0;
              Else
                Select Nvl(Max(Jh.是否预约), 0)
                Into n_特殊预约
                From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码,
                              To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                              To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                       From 挂号安排计划 Jh, 挂号计划时段 Sd
                       Where Jh.Id = Sd.计划id And Jh.Id = r_No.计划id And
                             Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6',
                                            '周五', '7', '周六', Null)) Jh, 挂号序号状态 Zt
                Where Zt.日期(+) = Jh.开始时间 And Zt.号码(+) = Jh.号码 And Zt.序号(+) = Jh.序号 And
                      Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) <> 1;
              End If;
            
              d_时段开始 := Null;
              d_时段结束 := Null;
              n_时段总数 := 0;
              n_时段剩余 := 0;
              For r_Time In (Select Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约, 0 As 已约数,
                                    Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数,
                                    Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) As 失效时段
                             From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                    From 挂号安排计划 Jh, 挂号计划时段 Sd
                                    Where Jh.Id = Sd.计划id And Jh.Id = r_No.计划id And Not Exists
                                     (Select 1
                                           From 挂号安排停用状态
                                           Where 安排id = Jh.安排id And
                                                 To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' ||
                                                         To_Char(Sd.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Between
                                                 开始停止时间 And 结束停止时间) And
                                          Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                         '5', '周四', '6', '周五', '7', '周六', Null)) Jh, 挂号序号状态 Zt
                             Where Zt.日期(+) = Jh.开始时间 And Zt.号码(+) = Jh.号码 And Zt.序号(+) = Jh.序号
                             Order By 序号) Loop
                If Nvl(n_时间间隔, 0) <> 0 Then
                  If d_时段开始 Is Null Then
                    d_时段开始 := r_Time.开始时间;
                    d_时段结束 := r_Time.开始时间 + n_时间间隔 / 24 / 60;
                    n_时段总数 := n_时段总数 + r_Time.限制数量;
                    If d_加号时间 < d_时段结束 Then
                      d_时段结束 := d_加号时间;
                    End If;
                  Else
                    If r_Time.开始时间 >= d_时段结束 Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_时段开始, 'hh24:mi:ss') || '-' ||
                                    To_Char(d_时段结束, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_时段总数 || '</GHZS>' ||
                                    '<SL>' || n_时段剩余 || '</SL>' || '</SPAN>';
                      n_时段总数 := r_Time.限制数量;
                      n_时段剩余 := 0;
                      d_时段开始 := r_Time.开始时间;
                      d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                      If d_加号时间 < d_时段结束 Then
                        d_时段结束 := d_加号时间;
                      End If;
                    Else
                      n_时段总数 := n_时段总数 + r_Time.限制数量;
                    End If;
                  End If;
                End If;
                If v_合作单位 Is Not Null Then
                  Select Nvl(Max(1), 0)
                  Into n_合约模式
                  From 合作单位计划控制
                  Where 限制项目 = r_Time.星期 And 计划id = r_No.计划id And 序号 = 0 And 合作单位 = v_合作单位;
                Else
                  n_合约模式 := 0;
                End If;
                If r_Time.剩余数 = 0 Then
                  n_单个剩余 := 0;
                Else
                  n_单个剩余 := r_Time.限制数量;
                End If;
                If v_合作单位 Is Null Or n_合约模式 = 1 Then
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From 合作单位计划控制
                  Where 限制项目 = r_Time.星期 And 计划id = r_No.计划id And 序号 = r_Time.序号 And Rownum < 2;
                  If n_Exists = 0 And r_Time.失效时段 <> 1 Then
                    If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                      Null;
                    Else
                      Select Nvl(Max(1), 0)
                      Into n_是否预留
                      From 挂号序号状态
                      Where 状态 In (3, 4) And 号码 = r_No.号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                      If n_是否预留 = 0 Then
                        If Nvl(n_时间间隔, 0) = 0 Then
                          v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                    '</SPAN>';
                        Else
                          n_时段剩余 := n_时段剩余 + n_单个剩余;
                        End If;
                        n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                      End If;
                    End If;
                  End If;
                Else
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From 合作单位计划控制
                  Where 限制项目 = r_Time.星期 And 计划id = r_No.计划id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
                  Begin
                    Select 0
                    Into n_非合约
                    From 合作单位计划控制
                    Where 计划id = r_No.计划id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_非合约 := 1;
                  End;
                  If (n_Exists = 1 Or n_非合约 = 1) And r_Time.失效时段 <> 1 Then
                    If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                      Null;
                    Else
                      Select Nvl(Max(1), 0)
                      Into n_是否预留
                      From 挂号序号状态
                      Where 状态 In (3, 4) And 号码 = r_No.号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                      If n_是否预留 = 0 Then
                        If Nvl(n_时间间隔, 0) = 0 Then
                          v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                    '</SPAN>';
                        Else
                          n_时段剩余 := n_时段剩余 + n_单个剩余;
                        End If;
                        n_合约剩余数量 := n_合约剩余数量 + 1;
                      End If;
                    End If;
                  End If;
                End If;
              End Loop;
              If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
                v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_时段开始, 'hh24:mi:ss') || '-' ||
                              To_Char(d_时段结束, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_时段总数 || '</GHZS>' || '<SL>' ||
                              n_时段剩余 || '</SL>' || '</SPAN>';
                n_时段总数 := 0;
                n_时段剩余 := 0;
                d_时段开始 := Null;
                d_时段结束 := Null;
              End If;
            Else
              n_最大可用数量 := Nvl(r_No.限约数, Nvl(r_No.限号数, 0)) - Nvl(r_No.已约数, 0);
              n_时段总数     := 0;
              n_时段剩余     := 0;
              d_时段开始     := Null;
              d_时段结束     := Null;
              For r_Time In (Select Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约,
                                    Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 已约数,
                                    Jh.限制数量 - Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 剩余数,
                                    Decode(Sign(Sysdate - Jh.开始时间), -1, 0, 1) As 失效时段
                             From (Select Sd.计划id, Sd.序号, Sd.星期, Jh.号码,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                    From 挂号安排计划 Jh, 挂号计划时段 Sd
                                    Where Jh.Id = Sd.计划id And Jh.Id = r_No.计划id And Not Exists
                                     (Select 1
                                           From 挂号安排停用状态
                                           Where 安排id = Jh.安排id And
                                                 To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' ||
                                                         To_Char(Sd.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Between
                                                 开始停止时间 And 结束停止时间) And
                                          Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                         '5', '周四', '6', '周五', '7', '周六', Null)) Jh, 挂号序号状态 Zt
                             Where Jh.号码 = Zt.号码(+) And Jh.开始时间 = Zt.日期(+)
                             Group By Jh.号码, Jh.序号, Jh.星期, Jh.开始时间, Jh.结束时间, Jh.限制数量, Jh.是否预约
                             Order By Jh.序号) Loop
                If Nvl(n_时间间隔, 0) <> 0 Then
                  If d_时段开始 Is Null Then
                    d_时段开始 := r_Time.开始时间;
                    d_时段结束 := r_Time.开始时间 + n_时间间隔 / 24 / 60;
                    n_时段总数 := n_时段总数 + r_Time.限制数量;
                    If d_加号时间 < d_时段结束 Then
                      d_时段结束 := d_加号时间;
                    End If;
                  Else
                    If r_Time.开始时间 >= d_时段结束 Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_时段开始, 'hh24:mi:ss') || '-' ||
                                    To_Char(d_时段结束, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_时段总数 || '</GHZS>' ||
                                    '<SL>' || n_时段剩余 || '</SL>' || '</SPAN>';
                      n_时段总数 := r_Time.限制数量;
                      n_时段剩余 := 0;
                      d_时段开始 := r_Time.开始时间;
                      d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                      If d_加号时间 < d_时段结束 Then
                        d_时段结束 := d_加号时间;
                      End If;
                    Else
                      n_时段总数 := n_时段总数 + r_Time.限制数量;
                    End If;
                  End If;
                End If;
                If v_合作单位 Is Not Null Then
                  Select Nvl(Max(1), 0)
                  Into n_合约模式
                  From 合作单位计划控制
                  Where 限制项目 = r_Time.星期 And 计划id = r_No.计划id And 序号 = 0 And 合作单位 = v_合作单位;
                Else
                  n_合约模式 := 0;
                End If;
                n_单个剩余 := r_Time.剩余数;
                If v_合作单位 Is Null Or n_合约模式 = 1 Then
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From 合作单位计划控制
                  Where 限制项目 = r_Time.星期 And 计划id = r_No.计划id And 序号 = r_Time.序号 And Rownum < 2;
                  If n_Exists = 0 And r_Time.失效时段 <> 1 Then
                    If n_最大可用数量 < n_单个剩余 Then
                      If Nvl(n_时间间隔, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_时段剩余 := n_时段剩余 + n_最大可用数量;
                      End If;
                      n_时段数量 := Nvl(n_时段数量, 0) + n_最大可用数量;
                    Else
                      If Nvl(n_时间间隔, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_时段剩余 := n_时段剩余 + n_单个剩余;
                      End If;
                      n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                    End If;
                  End If;
                Else
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From 合作单位计划控制
                  Where 限制项目 = r_Time.星期 And 计划id = r_No.计划id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
                  Begin
                    Select 0
                    Into n_非合约
                    From 合作单位计划控制
                    Where 计划id = r_No.计划id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_非合约 := 1;
                  End;
                  If (n_Exists = 1 Or n_非合约 = 1) And r_Time.失效时段 <> 1 Then
                    If n_最大可用数量 < n_单个剩余 Then
                      If Nvl(n_时间间隔, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_时段剩余 := n_时段剩余 + n_最大可用数量;
                      End If;
                      n_合约剩余数量 := n_合约剩余数量 + Nvl(n_最大可用数量, 0);
                    Else
                      If Nvl(n_时间间隔, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_时段剩余 := n_时段剩余 + n_单个剩余;
                      End If;
                      n_合约剩余数量 := n_合约剩余数量 + Nvl(n_单个剩余, 0);
                    End If;
                  End If;
                End If;
              End Loop;
              If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
                v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_时段开始, 'hh24:mi:ss') || '-' ||
                              To_Char(d_时段结束, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_时段总数 || '</GHZS>' || '<SL>' ||
                              n_时段剩余 || '</SL>' || '</SPAN>';
                n_时段总数 := 0;
                n_时段剩余 := 0;
                d_时段开始 := Null;
                d_时段结束 := Null;
              End If;
            End If;
          Else
            Select To_Date(To_Char(d_日期, 'yyyy-mm-dd') || To_Char(Max(结束时间), 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
            Into d_加号时间
            From 挂号安排时段
            Where 安排id = r_No.安排id And 星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                   '周四', '6', '周五', '7', '周六', Null);
            If r_No.序号控制 = 1 Then
              If Trunc(d_日期) = Trunc(Sysdate) Then
                n_特殊预约 := 0;
              Else
                Select Nvl(Max(Ap.是否预约), 0)
                Into n_特殊预约
                From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码,
                              To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                              To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                       From 挂号安排 Ap, 挂号安排时段 Sd
                       Where Ap.Id = Sd.安排id And Ap.Id = r_No.安排id And
                             Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6',
                                            '周五', '7', '周六', Null)) Ap, 挂号序号状态 Zt
                Where Zt.日期(+) = Ap.开始时间 And Zt.号码(+) = Ap.号码 And Zt.序号(+) = Ap.序号 And
                      Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) <> 1;
              End If;
              n_时段总数 := 0;
              n_时段剩余 := 0;
              d_时段开始 := Null;
              d_时段结束 := Null;
              For r_Time In (Select Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约, 0 As 已约数,
                                    Decode(Nvl(Zt.序号, 0), 0, 1, 0) As 剩余数,
                                    Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) As 失效时段
                             
                             From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                    From 挂号安排 Ap, 挂号安排时段 Sd
                                    Where Ap.Id = Sd.安排id And Ap.Id = r_No.安排id And Not Exists
                                     (Select 1
                                           From 挂号安排停用状态
                                           Where 安排id = Ap.Id And
                                                 To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' ||
                                                         To_Char(Sd.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Between
                                                 开始停止时间 And 结束停止时间) And
                                          Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                         '5', '周四', '6', '周五', '7', '周六', Null)) Ap, 挂号序号状态 Zt
                             Where Zt.日期(+) = Ap.开始时间 And Zt.号码(+) = Ap.号码 And Zt.序号(+) = Ap.序号
                             Order By 序号) Loop
                If Nvl(n_时间间隔, 0) <> 0 Then
                  If d_时段开始 Is Null Then
                    d_时段开始 := r_Time.开始时间;
                    n_时段总数 := n_时段总数 + r_Time.限制数量;
                    d_时段结束 := r_Time.开始时间 + n_时间间隔 / 24 / 60;
                    If d_加号时间 < d_时段结束 Then
                      d_时段结束 := d_加号时间;
                    End If;
                  Else
                    If r_Time.开始时间 >= d_时段结束 Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_时段开始, 'hh24:mi:ss') || '-' ||
                                    To_Char(d_时段结束, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_时段总数 || '</GHZS>' ||
                                    '<SL>' || n_时段剩余 || '</SL>' || '</SPAN>';
                      n_时段总数 := r_Time.限制数量;
                      n_时段剩余 := 0;
                      d_时段开始 := r_Time.开始时间;
                      d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                      If d_加号时间 < d_时段结束 Then
                        d_时段结束 := d_加号时间;
                      End If;
                    Else
                      n_时段总数 := n_时段总数 + r_Time.限制数量;
                    End If;
                  End If;
                End If;
                If v_合作单位 Is Not Null Then
                  Begin
                    Select 1
                    Into n_合约模式
                    From 合作单位安排控制
                    Where 限制项目 = r_Time.星期 And 安排id = r_No.安排id And 序号 = 0 And 合作单位 = v_合作单位;
                  Exception
                    When Others Then
                      n_合约模式 := 0;
                  End;
                Else
                  n_合约模式 := 0;
                End If;
                If r_Time.剩余数 = 0 Then
                  n_单个剩余 := 0;
                Else
                  n_单个剩余 := r_Time.限制数量;
                End If;
                If v_合作单位 Is Null Or n_合约模式 = 1 Then
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From 合作单位安排控制
                  Where 限制项目 = r_Time.星期 And 安排id = r_No.安排id And 序号 = r_Time.序号 And Rownum < 2;
                  If n_Exists = 0 And r_Time.失效时段 <> 1 Then
                    If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                      Null;
                    Else
                      Select Nvl(Max(1), 0)
                      Into n_是否预留
                      From 挂号序号状态
                      Where 状态 In (3, 4) And 号码 = r_No.号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                      If n_是否预留 = 0 Then
                        If Nvl(n_时间间隔, 0) = 0 Then
                          v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                    '</SPAN>';
                        Else
                          n_时段剩余 := n_时段剩余 + n_单个剩余;
                        End If;
                        n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                      End If;
                    End If;
                  End If;
                Else
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From 合作单位安排控制
                  Where 限制项目 = r_Time.星期 And 安排id = r_No.安排id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
                  Begin
                    Select 0
                    Into n_非合约
                    From 合作单位安排控制
                    Where 安排id = r_No.安排id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_非合约 := 1;
                  End;
                  If (n_Exists = 1 Or n_非合约 = 1) And r_Time.失效时段 <> 1 Then
                    If n_特殊预约 = 1 And r_Time.是否预约 = 0 Then
                      Null;
                    Else
                      Select Nvl(Max(1), 0)
                      Into n_是否预留
                      From 挂号序号状态
                      Where 状态 In (3, 4) And 号码 = r_No.号码 And 序号 = r_Time.序号 And Trunc(日期) = Trunc(d_日期);
                      If n_是否预留 = 0 Then
                        If Nvl(n_时间间隔, 0) = 0 Then
                          v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                    '</SPAN>';
                        Else
                          n_时段剩余 := n_时段剩余 + n_单个剩余;
                        End If;
                        n_合约剩余数量 := n_合约剩余数量 + n_单个剩余;
                      End If;
                    End If;
                  End If;
                End If;
              End Loop;
              If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
                v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_时段开始, 'hh24:mi:ss') || '-' ||
                              To_Char(d_时段结束, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_时段总数 || '</GHZS>' || '<SL>' ||
                              n_时段剩余 || '</SL>' || '</SPAN>';
                n_时段总数 := 0;
                n_时段剩余 := 0;
                d_时段开始 := Null;
                d_时段结束 := Null;
              End If;
            Else
              n_最大可用数量 := Nvl(r_No.限约数, Nvl(r_No.限号数, 0)) - Nvl(r_No.已约数, 0);
              n_时段总数     := 0;
              n_时段剩余     := 0;
              d_时段开始     := Null;
              d_时段结束     := Null;
              For r_Time In (Select Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约,
                                    Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 已约数,
                                    Ap.限制数量 - Sum(Decode(Nvl(Zt.序号, 0), 0, 0, 1)) As 剩余数,
                                    Decode(Sign(Sysdate - Ap.开始时间), -1, 0, 1) As 失效时段
                             From (Select Sd.安排id, Sd.序号, Sd.星期, Ap.号码,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.开始时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                           To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' || To_Char(Sd.结束时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As 结束时间, Sd.限制数量, Sd.是否预约
                                    From 挂号安排 Ap, 挂号安排时段 Sd
                                    Where Ap.Id = Sd.安排id And Ap.Id = r_No.安排id And Not Exists
                                     (Select 1
                                           From 挂号安排停用状态
                                           Where 安排id = Ap.Id And
                                                 To_Date(To_Char(d_日期, 'yyyy-mm-dd') || ' ' ||
                                                         To_Char(Sd.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Between
                                                 开始停止时间 And 结束停止时间) And
                                          Sd.星期 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三',
                                                         '5', '周四', '6', '周五', '7', '周六', Null)) Ap, 挂号序号状态 Zt
                             Where Ap.号码 = Zt.号码(+) And Ap.开始时间 = Zt.日期(+)
                             Group By Ap.号码, Ap.序号, Ap.星期, Ap.开始时间, Ap.结束时间, Ap.限制数量, Ap.是否预约
                             Order By Ap.序号) Loop
                If Nvl(n_时间间隔, 0) <> 0 Then
                  If d_时段开始 Is Null Then
                    d_时段开始 := r_Time.开始时间;
                    d_时段结束 := r_Time.开始时间 + n_时间间隔 / 24 / 60;
                    n_时段总数 := n_时段总数 + r_Time.限制数量;
                    If d_加号时间 < d_时段结束 Then
                      d_时段结束 := d_加号时间;
                    End If;
                  Else
                    If r_Time.开始时间 >= d_时段结束 Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_时段开始, 'hh24:mi:ss') || '-' ||
                                    To_Char(d_时段结束, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_时段总数 || '</GHZS>' ||
                                    '<SL>' || n_时段剩余 || '</SL>' || '</SPAN>';
                      n_时段总数 := r_Time.限制数量;
                      n_时段剩余 := 0;
                      d_时段开始 := r_Time.开始时间;
                      d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                      If d_加号时间 < d_时段结束 Then
                        d_时段结束 := d_加号时间;
                      End If;
                    Else
                      n_时段总数 := n_时段总数 + r_Time.限制数量;
                    End If;
                  End If;
                End If;
                If v_合作单位 Is Not Null Then
                  Select Nvl(Max(1), 0)
                  Into n_合约模式
                  From 合作单位安排控制
                  Where 限制项目 = r_Time.星期 And 安排id = r_No.安排id And 序号 = 0 And 合作单位 = v_合作单位;
                Else
                  n_合约模式 := 0;
                End If;
                n_单个剩余 := r_Time.剩余数;
                If v_合作单位 Is Null Or n_合约模式 = 1 Then
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From 合作单位安排控制
                  Where 限制项目 = r_Time.星期 And 安排id = r_No.安排id And 序号 = r_Time.序号 And Rownum < 2;
                  If n_Exists = 0 And r_Time.失效时段 <> 1 Then
                    If n_最大可用数量 < n_单个剩余 Then
                      If Nvl(n_时间间隔, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_时段剩余 := n_时段剩余 + n_最大可用数量;
                      End If;
                      n_时段数量 := Nvl(n_时段数量, 0) + n_最大可用数量;
                    Else
                      If Nvl(n_时间间隔, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_时段剩余 := n_时段剩余 + n_单个剩余;
                      End If;
                      n_时段数量 := Nvl(n_时段数量, 0) + n_单个剩余;
                    End If;
                  End If;
                Else
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From 合作单位安排控制
                  Where 限制项目 = r_Time.星期 And 安排id = r_No.安排id And 序号 = r_Time.序号 And 合作单位 = v_合作单位 And Rownum < 2;
                  Begin
                    Select 0
                    Into n_非合约
                    From 合作单位安排控制
                    Where 安排id = r_No.安排id And 合作单位 = v_合作单位 And Rownum < 2;
                  Exception
                    When Others Then
                      n_非合约 := 1;
                  End;
                  If (n_Exists = 1 Or n_非合约 = 1) And r_Time.失效时段 <> 1 Then
                    If n_最大可用数量 < n_单个剩余 Then
                      If Nvl(n_时间间隔, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_最大可用数量 || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_时段剩余 := n_时段剩余 + n_最大可用数量;
                      End If;
                      n_合约剩余数量 := n_合约剩余数量 + Nvl(n_最大可用数量, 0);
                    Else
                      If Nvl(n_时间间隔, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.开始时间, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.结束时间, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_单个剩余 || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_时段剩余 := n_时段剩余 + n_单个剩余;
                      End If;
                      n_合约剩余数量 := n_合约剩余数量 + Nvl(n_单个剩余, 0);
                    End If;
                  End If;
                End If;
              End Loop;
              If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
                v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_时段开始, 'hh24:mi:ss') || '-' ||
                              To_Char(d_时段结束, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_时段总数 || '</GHZS>' || '<SL>' ||
                              n_时段剩余 || '</SL>' || '</SPAN>';
                n_时段总数 := 0;
                n_时段剩余 := 0;
                d_时段开始 := Null;
                d_时段结束 := Null;
              End If;
            End If;
          End If;
        End If;
        If v_合作单位 Is Not Null Then
          If Nvl(r_No.计划id, 0) <> 0 Then
            Begin
              Select 0
              Into n_非合约
              From 合作单位计划控制
              Where 计划id = r_No.计划id And 合作单位 = v_合作单位 And Rownum < 2;
            Exception
              When Others Then
                n_非合约 := 1;
            End;
          Else
            Begin
              Select 0
              Into n_非合约
              From 合作单位安排控制
              Where 安排id = r_No.安排id And 合作单位 = v_合作单位 And Rownum < 2;
            Exception
              When Others Then
                n_非合约 := 1;
            End;
          End If;
        End If;
        If v_合作单位 Is Null Or n_非合约 = 1 Then
          If r_No.限号数 = 0 Then
            v_剩余数量 := '';
          Else
            If Nvl(r_No.计划id, 0) <> 0 Then
              Select Sum(数量)
              Into n_合约总数量
              From 合作单位计划控制
              Where 计划id = r_No.计划id And
                    限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null);
            Else
              Select Sum(数量)
              Into n_合约总数量
              From 合作单位安排控制
              Where 安排id = r_No.安排id And
                    限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null);
            End If;
            Select Count(1)
            Into n_合约已挂数
            From 病人挂号记录
            Where 号别 = r_No.号码 And 记录状态 = 1 And 合作单位 Is Not Null And 发生时间 Between Trunc(d_日期) And
                  Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
            Select Count(1)
            Into n_预留数量
            From 挂号序号状态
            Where 状态 = 3 And 号码 = r_No.号码 And Trunc(日期) = Trunc(d_日期);
            If Trunc(d_日期) = Trunc(Sysdate) Then
              If Nvl(n_合约总数量, 0) = 0 Then
                v_剩余数量 := r_No.限号数 - r_No.已挂数 - r_No.已约数 + r_No.已接收 - n_预留数量;
              Else
                v_剩余数量 := r_No.限号数 - r_No.已挂数 - r_No.已约数 + r_No.已接收 - n_合约总数量 + n_合约已挂数 - n_预留数量;
              End If;
              n_已挂数 := r_No.已挂数;
              If Nvl(n_时段数量, 0) < v_剩余数量 And n_分时段 <> 0 Then
                n_缓冲序号 := 1;
                v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_加号时间, 'hh24:mi:ss') || '-' || '</SJD>' ||
                              '<SL>' || To_Number(v_剩余数量 - Nvl(n_时段数量, 0) - Nvl(n_合约剩余数量, 0)) || '</SL>' || '</SPAN>';
              Else
                n_缓冲序号 := 0;
              End If;
            Else
              If Nvl(n_合约总数量, 0) = 0 Then
                v_剩余数量 := r_No.限约数 - r_No.已约数 - n_预留数量;
                If v_剩余数量 Is Null Then
                  v_剩余数量 := r_No.限号数 - r_No.已挂数 - r_No.已约数 + r_No.已接收 - n_预留数量;
                End If;
              Else
                v_剩余数量 := r_No.限约数 - r_No.已约数 - n_合约总数量 + n_合约已挂数 - n_预留数量;
                If v_剩余数量 Is Null Then
                  v_剩余数量 := r_No.限号数 - r_No.已挂数 - r_No.已约数 + r_No.已接收 - n_合约总数量 + n_合约已挂数 - n_预留数量;
                End If;
              End If;
              n_已挂数 := r_No.已挂数;
            End If;
          End If;
        Else
          If Nvl(r_No.计划id, 0) <> 0 Then
            If v_合作单位 Is Not Null Then
              Select Nvl(Max(1), 0)
              Into n_合约模式
              From 合作单位计划控制
              Where 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null) And 计划id = r_No.计划id And 序号 = 0 And 合作单位 = v_合作单位;
            Else
              n_合约模式 := 0;
            End If;
            Select Sum(数量)
            Into n_合约总数量
            From 合作单位计划控制
            Where 计划id = r_No.计划id And 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                     '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
          Else
            If v_合作单位 Is Not Null Then
              Select Nvl(Max(1), 0)
              Into n_合约模式
              From 合作单位安排控制
              Where 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null) And 安排id = r_No.安排id And 序号 = 0 And 合作单位 = v_合作单位;
            Else
              n_合约模式 := 0;
            End If;
            Select Sum(数量)
            Into n_合约总数量
            From 合作单位安排控制
            Where 安排id = r_No.安排id And 限制项目 = Decode(To_Char(d_日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                                     '周四', '6', '周五', '7', '周六', Null) And 合作单位 = v_合作单位;
          End If;
          If n_合约模式 = 0 Then
            v_剩余数量   := n_合约剩余数量;
            n_已挂数     := r_No.已挂数;
            n_合约已挂数 := Nvl(n_合约总数量, 0) - n_合约剩余数量;
          Else
            n_已挂数 := r_No.已挂数;
            Select Count(1)
            Into n_合约已挂数
            From 病人挂号记录
            Where 号别 = r_No.号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_日期) And
                  Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
            If Nvl(n_合约总数量, 0) = 0 Then
              v_剩余数量 := '0';
            Else
              v_剩余数量 := n_合约总数量 - n_合约已挂数;
            End If;
          End If;
        End If;
        Select To_Char(开始时间, 'hh24:mi') Into v_Timetemp From 时间段 Where 时间段 = r_No.排班;
        v_时间段 := v_Timetemp || '-';
        Select To_Char(终止时间, 'hh24:mi') Into v_Timetemp From 时间段 Where 时间段 = r_No.排班;
        v_时间段 := v_时间段 || v_Timetemp;
        If v_Temp Is Not Null Then
          v_Temp := v_Temp || '</SPANLIST>';
        End If;
        If v_合作单位 Is Not Null Then
          If Nvl(r_No.计划id, 0) <> 0 Then
            Select Nvl(Max(1), 0)
            Into n_禁用
            From 合作单位计划控制
            Where 计划id = r_No.计划id And 合作单位 = v_合作单位 And 数量 = 0 And Rownum < 2;
          Else
            Select Nvl(Max(1), 0)
            Into n_禁用
            From 合作单位安排控制
            Where 安排id = r_No.安排id And 合作单位 = v_合作单位 And 数量 = 0 And Rownum < 2;
          End If;
        End If;
        --限约数=0的预约禁止
        If Trunc(d_日期) <> Trunc(Sysdate) Then
          If r_No.限约数 = 0 Then
            n_禁用 := 1;
          End If;
        End If;
        If Nvl(n_禁用, 0) = 0 Then
          --从项金额计算
          n_合计金额 := r_No.价格;
          For r_Subfee In (Select 现价, 从项数次
                           From 收费从属项目 A, 收费价目 B
                           Where a.主项id = r_No.项目id And a.从项id = b.收费细目id And d_日期 Between b.执行日期 And
                                 Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                                 (b.价格等级 = v_普通等级 Or
                                 (b.价格等级 Is Null And Not Exists
                                  (Select 1
                                    From 收费价目
                                    Where 收费细目id = b.收费细目id And 价格等级 = v_普通等级 And d_日期 Between 执行日期 And
                                          Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))) Loop
            n_合计金额 := n_合计金额 + r_Subfee.现价 * r_Subfee.从项数次;
          End Loop;
          If Trunc(Sysdate) = Trunc(d_日期) Then
            Select Nvl(Max(1), 0)
            Into n_Exists
            From (Select 时间段
                   From 时间段
                   Where ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') < '3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS')) Or
                         ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') <
                         Decode(Sign(开始时间 - 终止时间), 1, '3000-01-11 ' || To_Char(终止时间, 'HH24:MI:SS'),
                                 '3000-01-10 ' || To_Char(终止时间, 'HH24:MI:SS'))))
            Where 时间段 = r_No.排班;
          Else
            n_Exists := 1;
          End If;
          If n_Exists = 1 Then
            If v_剩余数量 > 0 Then
              c_Xmlmain := '<HB>' || '<APID>' || r_No.安排id || '</APID>' || '<JHID>' || r_No.计划id || '</JHID>' || '<HM>' ||
                           r_No.号码 || '</HM>' || '<YSID>' || r_No.医生id || '</YSID>' || '<YS>' || r_No.医生姓名 || '</YS>' ||
                           '<KSID>' || r_No.科室id || '</KSID>' || '<KSMC>' || r_No.科室名称 || '</KSMC>' || '<ZC>' ||
                           r_No.职称 || '</ZC>' || '<XMID>' || r_No.项目id || '</XMID>' || '<XMMC>' || r_No.项目名称 ||
                           '</XMMC>' || '<YGHS>' || n_已挂数 || '</YGHS>' || '<SYHS>' || v_剩余数量 || '</SYHS>' || '<PRICE>' ||
                           n_合计金额 || '</PRICE>' || '<HCXH>' || n_缓冲序号 || '</HCXH>' || '<HL>' || r_No.号类 || '</HL>' ||
                           '<FSD>' || n_分时段 || '</FSD>' || '<HBTIME>' || v_时间段 || '</HBTIME>' || '<FWMC>' || r_No.排班 ||
                           '</FWMC>' || v_Temp || '</HB>';
            Else
              c_Xmlmain := '<HB>' || '<APID>' || r_No.安排id || '</APID>' || '<JHID>' || r_No.计划id || '</JHID>' || '<HM>' ||
                           r_No.号码 || '</HM>' || '<YSID>' || r_No.医生id || '</YSID>' || '<YS>' || r_No.医生姓名 || '</YS>' ||
                           '<KSID>' || r_No.科室id || '</KSID>' || '<KSMC>' || r_No.科室名称 || '</KSMC>' || '<ZC>' ||
                           r_No.职称 || '</ZC>' || '<XMID>' || r_No.项目id || '</XMID>' || '<XMMC>' || r_No.项目名称 ||
                           '</XMMC>' || '<YGHS>' || n_已挂数 || '</YGHS>' || '<SYHS>' || v_剩余数量 || '</SYHS>' || '<PRICE>' ||
                           n_合计金额 || '</PRICE>' || '<HL>' || r_No.号类 || '</HL>' || '<FSD>' || n_分时段 || '</FSD>' ||
                           '<HBTIME>' || v_时间段 || '</HBTIME>' || '<FWMC>' || r_No.排班 || '</FWMC>' || '</HB>';
            End If;
            v_Xmlmain := v_Xmlmain || c_Xmlmain;
          End If;
        End If;
      End If;
      n_合约剩余数量 := 0;
      n_合约总数量   := 0;
      n_时段数量     := 0;
      n_禁用         := 0;
      n_非合约       := 0;
    End Loop;
  
    v_Xmlmain := '<GROUP>' || '<RQ>' || To_Char(d_日期, 'yyyy-mm-dd') || '</RQ>' || '<HBLIST>' || v_Xmlmain ||
                 '</HBLIST>' || '</GROUP>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Xmlmain)) Into x_Templet From Dual;
    Xml_Out := x_Templet;
    Return;
  End If;

  --===========================================================================================
  --出诊表排班模式 
  --===========================================================================================
  n_合约剩余数量 := 0;
  n_补充天数     := Zl_Fun_Getappointmentdays;
  --注意出诊记录停用了，但可能启用了部分时段
  --临床出诊序号控制 中，开始时间与终止时间相等的是加号的序号
  --记录性质：1-正常出诊记录,2-替诊出诊记录
  For r_No In (Select a.记录性质, a.记录id, a.号源id, b.号类, b.号码, a.科室id, c.名称 As 科室名称, a.项目id, e.名称 As 项目名称, a.医生id, a.医生姓名,
                      d.专业技术职务 As 职称, a.排班, a.开始时间, a.终止时间, a.序号控制, a.分时段, a.预约控制, a.限号数, a.限约数, a.已挂数, a.已约数, a.已接收,
                      a.替诊开始时间, a.替诊终止时间, a.停诊开始时间, a.停诊终止时间, Nvl(b.预约天数, n_预约天数) + n_补充天数 As 预约天数
               From (Select 1 As 记录性质, a.Id As 记录id, a.号源id, a.科室id, a.项目id, a.医生id, a.医生姓名, a.上班时段 As 排班, a.开始时间, a.终止时间,
                             Nvl(a.是否序号控制, 0) As 序号控制, Nvl(a.是否分时段, 0) As 分时段, a.预约控制, a.限号数, Nvl(a.限约数, a.限号数) As 限约数,
                             Nvl(a.已挂数, 0) As 已挂数, Nvl(a.已约数, 0) As 已约数, Nvl(a.其中已接收, 0) As 已接收, a.替诊开始时间, a.替诊终止时间,
                             a.停诊开始时间, a.停诊终止时间
                      From 临床出诊记录 A
                      Where Nvl(a.是否发布, 0) = 1 And Nvl(a.是否锁定, 0) = 0 And
                            (a.开始时间 < Nvl(a.替诊开始时间, a.终止时间) Or a.终止时间 > Nvl(a.替诊终止时间, a.开始时间)) And a.开始时间 > Trunc(d_启用时间) And
                            a.终止时间 > Sysdate And
                            (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) And Nvl(a.停诊开始时间, a.终止时间) > Sysdate Or
                             a.终止时间 > Nvl(a.停诊终止时间, a.开始时间) And a.终止时间 > Sysdate Or Exists
                             (Select 1
                              From 临床出诊序号控制
                              Where 记录id = a.Id And Nvl(是否停诊, 0) = 0 And Nvl(a.是否序号控制, 0) = 1 And Nvl(a.是否分时段, 0) = 1 And
                                    开始时间 <> 终止时间 And 开始时间 >= Sysdate)) And
                            Decode(Nvl(n_医生id, 0), 0, 0, a.医生id) = Nvl(n_医生id, 0) And
                            Decode(Nvl(v_医生姓名, '-'), '-', '-', a.医生姓名) = Nvl(v_医生姓名, '-') And
                            Decode(Nvl(n_科室id, 0), 0, 0, a.科室id) = Nvl(n_科室id, 0) And a.出诊日期 = Trunc(d_日期)
                      Union All
                      Select 2 As 记录性质, a.Id As 记录id, a.号源id, a.科室id, a.项目id, a.替诊医生id As 医生id, a.替诊医生姓名 As 医生姓名,
                             a.上班时段 As 排班, a.开始时间, a.终止时间, Nvl(a.是否序号控制, 0) As 序号控制, Nvl(a.是否分时段, 0) As 分时段, a.预约控制, a.限号数,
                             Nvl(a.限约数, a.限号数) As 限约数, Nvl(a.已挂数, 0) As 已挂数, Nvl(a.已约数, 0) As 已约数, Nvl(a.其中已接收, 0) As 已接收,
                             a.替诊开始时间, a.替诊终止时间, a.停诊开始时间, a.停诊终止时间
                      From 临床出诊记录 A
                      Where Nvl(a.是否发布, 0) = 1 And Nvl(a.是否锁定, 0) = 0 And a.开始时间 > Trunc(d_启用时间) And a.终止时间 > Sysdate And
                            (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) And Nvl(a.停诊开始时间, a.终止时间) > Sysdate Or
                             a.终止时间 > Nvl(a.停诊终止时间, a.开始时间) And a.终止时间 > Sysdate Or Exists
                             (Select 1
                              From 临床出诊序号控制
                              Where 记录id = a.Id And Nvl(是否停诊, 0) = 0 And Nvl(a.是否序号控制, 0) = 1 And Nvl(a.是否分时段, 0) = 1 And
                                    开始时间 <> 终止时间 And 开始时间 >= Sysdate)) And
                            Decode(Nvl(n_医生id, 0), 0, 0, a.替诊医生id) = Nvl(n_医生id, 0) And
                            Decode(Nvl(v_医生姓名, '-'), '-', '-', a.替诊医生姓名) = Nvl(v_医生姓名, '-') And
                            Decode(Nvl(n_科室id, 0), 0, 0, a.科室id) = Nvl(n_科室id, 0) And a.替诊医生姓名 Is Not Null And
                            a.出诊日期 = Trunc(d_日期)) A, 临床出诊号源 B, 部门表 C, 人员表 D, 收费项目目录 E
               Where a.号源id = b.Id And a.科室id = c.Id And a.项目id = e.Id And a.医生id = d.Id(+)) Loop
  
    Zl_挂号序号状态_出诊_Delete(r_No.记录id);
    v_Temp := '';
    n_禁用 := 0;
    If Sysdate + Nvl(r_No.预约天数, n_预约天数) + n_补充天数 >= d_日期 Then
      If Trunc(d_日期) = Trunc(Sysdate) Then
        --当日
        If v_合作单位 Is Null Then
          --未传入合作单位
          n_已挂数   := r_No.已挂数;
          v_剩余数量 := r_No.限号数 - Nvl(r_No.已挂数, 0) - (Nvl(r_No.已约数, 0) - Nvl(r_No.已接收, 0));
          If r_No.分时段 = 1 And r_No.序号控制 = 1 Then
            v_Temp     := '<SPANLIST>';
            n_Exists   := 0;
            n_时段总数 := 0;
            n_时段剩余 := 0;
            d_时段开始 := Null;
            d_时段结束 := Null;
            Select Max(终止时间) Into d_加号时间 From 临床出诊序号控制 Where 记录id = r_No.记录id;
            For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态
                           From 临床出诊序号控制
                           Where 记录id = r_No.记录id And (r_No.记录性质 = 1 And 开始时间 Not Between Nvl(r_No.替诊开始时间, Sysdate) And
                                 Nvl(r_No.替诊终止时间, Sysdate - 1) Or
                                 r_No.记录性质 = 2 And 开始时间 Between Nvl(r_No.替诊开始时间, Sysdate) And
                                 Nvl(r_No.替诊终止时间, Sysdate - 1)) And Nvl(是否停诊, 0) = 0) Loop
              If r_Time.开始时间 > Sysdate Then
                If Nvl(n_时间间隔, 0) = 0 Then
                  If Nvl(r_Time.挂号状态, 0) = 0 Then
                    n_时段剩余 := 1;
                    n_Exists   := n_Exists + 1;
                  Else
                    n_时段剩余 := 0;
                  End If;
                  v_Temp := v_Temp || Gettimexml(r_Time.开始时间, r_Time.终止时间, 1, n_时段剩余);
                Else
                  If d_时段开始 Is Null Then
                    n_时段总数 := 1;
                    n_时段剩余 := 0;
                    d_时段开始 := r_Time.开始时间;
                    d_时段结束 := r_Time.开始时间 + n_时间间隔 / 24 / 60;
                    If d_加号时间 < d_时段结束 Then
                      d_时段结束 := d_加号时间;
                    End If;
                  Else
                    If r_Time.开始时间 >= d_时段结束 Then
                      v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                    
                      n_时段总数 := 1;
                      n_时段剩余 := 0;
                      d_时段开始 := r_Time.开始时间;
                      d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                      If d_加号时间 < d_时段结束 Then
                        d_时段结束 := d_加号时间;
                      End If;
                    Else
                      n_时段总数 := n_时段总数 + 1;
                    End If;
                  End If;
                
                  If Nvl(r_Time.挂号状态, 0) = 0 Then
                    n_时段剩余 := n_时段剩余 + 1;
                    n_Exists   := n_Exists + 1;
                  End If;
                End If;
              End If;
            End Loop;
          
            If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
              v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
            End If;
          
            If r_No.记录性质 = 1 And n_Exists < To_Number(v_剩余数量) Then
              v_Temp := v_Temp || Gettimexml(d_加号时间, '', v_剩余数量, To_Number(v_剩余数量) - n_Exists);
            End If;
            v_Temp := v_Temp || '</SPANLIST>';
          End If;
        Else
          --传入合作单位
          n_已挂数 := r_No.已挂数;
          Begin
            Select 控制方式
            Into n_合约模式
            From 临床出诊挂号控制记录
            Where 记录id = r_No.记录id And 类型 = 1 And 性质 = 1 And 名称 = v_合作单位 And Rownum < 2;
          Exception
            When Others Then
              n_合约模式 := 4;
          End;
        
          If n_合约模式 = 0 Then
            n_禁用 := 1;
          Elsif n_合约模式 = 1 Or n_合约模式 = 2 Then
            Select 数量
            Into n_合约总数量
            From 临床出诊挂号控制记录
            Where 记录id = r_No.记录id And 类型 = 1 And 性质 = 1 And 名称 = v_合作单位;
            If n_合约模式 = 1 Then
              n_合约总数量 := Floor(r_No.限约数 * n_合约总数量 / 100);
            End If;
          
            Select Count(1)
            Into n_合约已挂数
            From 病人挂号记录
            Where 号别 = r_No.号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_日期) And
                  Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
            n_合约剩余数量 := n_合约总数量 - n_合约已挂数;
          
            If r_No.限号数 - Nvl(r_No.已挂数, 0) < n_合约剩余数量 Then
              v_剩余数量 := r_No.限号数 - Nvl(r_No.已挂数, 0);
            Else
              v_剩余数量 := n_合约剩余数量;
            End If;
          
            If r_No.分时段 = 1 And r_No.序号控制 = 1 Then
              v_Temp     := '<SPANLIST>';
              n_Exists   := 0;
              n_时段总数 := 0;
              n_时段剩余 := 0;
              d_时段开始 := Null;
              d_时段结束 := Null;
              Select Max(终止时间) Into d_加号时间 From 临床出诊序号控制 Where 记录id = r_No.记录id;
              For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态
                             From 临床出诊序号控制
                             Where 记录id = r_No.记录id And (r_No.记录性质 = 1 And 开始时间 Not Between Nvl(r_No.替诊开始时间, Sysdate) And
                                   Nvl(r_No.替诊终止时间, Sysdate - 1) Or
                                   r_No.记录性质 = 2 And 开始时间 Between Nvl(r_No.替诊开始时间, Sysdate) And
                                   Nvl(r_No.替诊终止时间, Sysdate - 1)) And Nvl(是否停诊, 0) = 0) Loop
              
                If r_Time.开始时间 > Sysdate Then
                  If Nvl(n_时间间隔, 0) = 0 Then
                    If Nvl(r_Time.挂号状态, 0) = 0 Then
                      n_时段剩余 := 1;
                      n_Exists   := n_Exists + 1;
                    Else
                      n_时段剩余 := 0;
                    End If;
                    v_Temp := v_Temp || Gettimexml(r_Time.开始时间, r_Time.终止时间, 1, n_时段剩余);
                  Else
                    If d_时段开始 Is Null Then
                      n_时段总数 := 1;
                      n_时段剩余 := 0;
                      d_时段开始 := r_Time.开始时间;
                      d_时段结束 := r_Time.开始时间 + n_时间间隔 / 24 / 60;
                      If d_加号时间 < d_时段结束 Then
                        d_时段结束 := d_加号时间;
                      End If;
                    Else
                      If r_Time.开始时间 >= d_时段结束 Then
                        v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                      
                        n_时段总数 := 1;
                        n_时段剩余 := 0;
                        d_时段开始 := r_Time.开始时间;
                        d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                        If d_加号时间 < d_时段结束 Then
                          d_时段结束 := d_加号时间;
                        End If;
                      Else
                        n_时段总数 := n_时段总数 + 1;
                      End If;
                    End If;
                  
                    If Nvl(r_Time.挂号状态, 0) = 0 Then
                      n_时段剩余 := n_时段剩余 + 1;
                      n_Exists   := n_Exists + 1;
                    End If;
                  End If;
                End If;
              End Loop;
            
              If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
                v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
              End If;
            
              If r_No.记录性质 = 1 And n_Exists < To_Number(v_剩余数量) Then
                v_Temp := v_Temp || Gettimexml(d_加号时间, '', v_剩余数量, To_Number(v_剩余数量) - n_Exists);
              End If;
              v_Temp := v_Temp || '</SPANLIST>';
            End If;
          Elsif n_合约模式 = 3 Then
            If r_No.序号控制 = 0 Then
              n_已挂数   := r_No.已挂数;
              v_剩余数量 := r_No.限号数 - Nvl(r_No.已挂数, 0) - (Nvl(r_No.已约数, 0) - Nvl(r_No.已接收, 0));
            Else
              v_Temp     := '<SPANLIST>';
              n_已挂数   := 0;
              v_剩余数量 := 0;
              n_时段总数 := 0;
              n_时段剩余 := 0;
              d_时段开始 := Null;
              d_时段结束 := Null;
              For r_合作 In (Select 序号
                           From 临床出诊挂号控制记录
                           Where 记录id = r_No.记录id And 类型 = 1 And 性质 = 1 And 名称 = v_合作单位) Loop
              
                Begin
                  Select 1, 开始时间, 终止时间
                  Into n_Exists, d_开始时间, d_终止时间
                  From 临床出诊序号控制
                  Where 记录id = r_No.记录id And 序号 = r_合作.序号 And Nvl(挂号状态, 0) = 0 And
                        (r_No.记录性质 = 1 And 开始时间 Not Between Nvl(r_No.替诊开始时间, Sysdate) And Nvl(r_No.替诊终止时间, Sysdate - 1) Or
                        r_No.记录性质 = 2 And 开始时间 Between Nvl(r_No.替诊开始时间, Sysdate) And Nvl(r_No.替诊终止时间, Sysdate - 1)) And
                        Nvl(是否停诊, 0) = 0;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
              
                If n_Exists = 1 Then
                  v_剩余数量 := v_剩余数量 + 1;
                Else
                  n_已挂数 := n_已挂数 + 1;
                End If;
              
                If d_开始时间 > Sysdate Then
                  If Nvl(n_时间间隔, 0) = 0 Then
                    If n_Exists = 1 Then
                      n_时段剩余 := 1;
                    Else
                      n_时段剩余 := 0;
                    End If;
                    v_Temp := v_Temp || Gettimexml(d_开始时间, d_终止时间, 1, n_时段剩余);
                  Else
                    If d_时段开始 Is Null Then
                      n_时段总数 := 1;
                      n_时段剩余 := 0;
                      d_时段开始 := d_开始时间;
                      d_时段结束 := d_开始时间 + n_时间间隔 / 24 / 60;
                    Else
                      If d_开始时间 >= d_时段结束 Then
                        v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                      
                        n_时段总数 := 1;
                        n_时段剩余 := 0;
                        d_时段开始 := d_开始时间;
                        d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                      Else
                        n_时段总数 := n_时段总数 + 1;
                      End If;
                    End If;
                  
                    If n_Exists = 1 Then
                      n_时段剩余 := n_时段剩余 + 1;
                    End If;
                  End If;
                End If;
              End Loop;
            
              If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
                v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
              End If;
              v_Temp := v_Temp || '</SPANLIST>';
            End If;
          Elsif n_合约模式 = 4 Then
            n_已挂数   := r_No.已挂数;
            v_剩余数量 := r_No.限号数 - Nvl(r_No.已挂数, 0) - (Nvl(r_No.已约数, 0) - Nvl(r_No.已接收, 0));
            If r_No.分时段 = 1 And r_No.序号控制 = 1 Then
              v_Temp     := '<SPANLIST>';
              n_Exists   := 0;
              n_时段总数 := 0;
              n_时段剩余 := 0;
              d_时段开始 := Null;
              d_时段结束 := Null;
              Select Max(终止时间) Into d_加号时间 From 临床出诊序号控制 Where 记录id = r_No.记录id;
              For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态
                             From 临床出诊序号控制
                             Where 记录id = r_No.记录id And (r_No.记录性质 = 1 And 开始时间 Not Between Nvl(r_No.替诊开始时间, Sysdate) And
                                   Nvl(r_No.替诊终止时间, Sysdate - 1) Or
                                   r_No.记录性质 = 2 And 开始时间 Between Nvl(r_No.替诊开始时间, Sysdate) And
                                   Nvl(r_No.替诊终止时间, Sysdate - 1)) And Nvl(是否停诊, 0) = 0) Loop
              
                If r_Time.开始时间 > Sysdate Then
                  If Nvl(n_时间间隔, 0) = 0 Then
                    If Nvl(r_Time.挂号状态, 0) = 0 Then
                      n_时段剩余 := 1;
                      n_Exists   := n_Exists + 1;
                    Else
                      n_时段剩余 := 0;
                    End If;
                    v_Temp := v_Temp || Gettimexml(r_Time.开始时间, r_Time.终止时间, 1, n_时段剩余);
                  Else
                    If d_时段开始 Is Null Then
                      n_时段总数 := 1;
                      n_时段剩余 := 0;
                      d_时段开始 := r_Time.开始时间;
                      d_时段结束 := r_Time.开始时间 + n_时间间隔 / 24 / 60;
                      If d_加号时间 < d_时段结束 Then
                        d_时段结束 := d_加号时间;
                      End If;
                    Else
                      If r_Time.开始时间 >= d_时段结束 Then
                        v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                      
                        n_时段总数 := 1;
                        n_时段剩余 := 0;
                        d_时段开始 := r_Time.开始时间;
                        d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                        If d_加号时间 < d_时段结束 Then
                          d_时段结束 := d_加号时间;
                        End If;
                      Else
                        n_时段总数 := n_时段总数 + 1;
                      End If;
                    End If;
                  
                    If Nvl(r_Time.挂号状态, 0) = 0 Then
                      n_时段剩余 := n_时段剩余 + 1;
                      n_Exists   := n_Exists + 1;
                    End If;
                  End If;
                End If;
              End Loop;
            
              If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
                v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
              End If;
            
              If r_No.记录性质 = 1 And n_Exists < To_Number(v_剩余数量) Then
                v_Temp := v_Temp || Gettimexml(d_加号时间, '', v_剩余数量, To_Number(v_剩余数量) - n_Exists);
              End If;
              v_Temp := v_Temp || '</SPANLIST>';
            End If;
          End If;
        End If;
      Else
        --预约挂号
        If r_No.预约控制 = 1 Then
          n_禁用 := 1;
        Else
          --不限制预约
          If v_合作单位 Is Null Then
            If r_No.分时段 = 0 Then
              n_已挂数   := r_No.已约数;
              v_剩余数量 := Nvl(r_No.限约数, r_No.限号数) - Nvl(r_No.已约数, 0);
            Else
              --分时段
              v_Temp     := '<SPANLIST>';
              n_已挂数   := 0;
              v_剩余数量 := 0;
              n_时段总数 := 0;
              n_时段剩余 := 0;
              d_时段开始 := Null;
              d_时段结束 := Null;
              Select Max(终止时间) Into d_加号时间 From 临床出诊序号控制 Where 记录id = r_No.记录id;
              If r_No.序号控制 = 0 Then
                --非序号控制分时段预约
                For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态, 数量
                               From 临床出诊序号控制
                               Where 记录id = r_No.记录id And 预约顺序号 Is Null And 是否预约 = 1 And
                                     (r_No.记录性质 = 1 And 开始时间 Not Between Nvl(r_No.替诊开始时间, Sysdate) And
                                     Nvl(r_No.替诊终止时间, Sysdate - 1) Or
                                     r_No.记录性质 = 2 And 开始时间 Between Nvl(r_No.停诊开始时间, Sysdate) And
                                     Nvl(r_No.停诊终止时间, Sysdate - 1))) Loop
                
                  Select Count(1)
                  Into n_时段已挂
                  From 临床出诊序号控制
                  Where 记录id = r_No.记录id And 序号 = r_Time.序号 And 预约顺序号 Is Not Null And Nvl(挂号状态, 0) <> 0;
                
                  n_已挂数   := n_已挂数 + n_时段已挂;
                  v_剩余数量 := v_剩余数量 + (r_Time.数量 - n_时段已挂);
                
                  If r_Time.开始时间 > Sysdate Then
                    If Nvl(n_时间间隔, 0) = 0 Then
                      n_时段剩余 := r_Time.数量 - n_时段已挂;
                      v_Temp     := v_Temp || Gettimexml(r_Time.开始时间, r_Time.终止时间, r_Time.数量, n_时段剩余);
                    Else
                      If d_时段开始 Is Null Then
                        n_时段总数 := r_Time.数量;
                        n_时段剩余 := 0;
                        d_时段开始 := r_Time.开始时间;
                        d_时段结束 := r_Time.开始时间 + n_时间间隔 / 24 / 60;
                        If d_加号时间 < d_时段结束 Then
                          d_时段结束 := d_加号时间;
                        End If;
                      Else
                        If r_Time.开始时间 >= d_时段结束 Then
                          v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                        
                          n_时段总数 := r_Time.数量;
                          n_时段剩余 := 0;
                          d_时段开始 := r_Time.开始时间;
                          d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                          If d_加号时间 < d_时段结束 Then
                            d_时段结束 := d_加号时间;
                          End If;
                        Else
                          n_时段总数 := n_时段总数 + r_Time.数量;
                        End If;
                      End If;
                    
                      n_时段剩余 := n_时段剩余 + r_Time.数量 - n_时段已挂;
                    End If;
                  End If;
                End Loop;
              
                If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
                  v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                End If;
              Else
                --序号控制分时段预约
                For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态
                               From 临床出诊序号控制
                               Where 记录id = r_No.记录id And 是否预约 = 1 And
                                     (r_No.记录性质 = 1 And 开始时间 Not Between Nvl(r_No.替诊开始时间, Sysdate) And
                                     Nvl(r_No.替诊终止时间, Sysdate - 1) Or
                                     r_No.记录性质 = 2 And 开始时间 Between Nvl(r_No.替诊开始时间, Sysdate) And
                                     Nvl(r_No.替诊终止时间, Sysdate - 1)) And Nvl(是否停诊, 0) = 0) Loop
                
                  If Nvl(r_Time.挂号状态, 0) = 0 Then
                    v_剩余数量 := v_剩余数量 + 1;
                  Else
                    n_已挂数 := n_已挂数 + 1;
                  End If;
                
                  If r_Time.开始时间 > Sysdate Then
                    If Nvl(n_时间间隔, 0) = 0 Then
                      If Nvl(r_Time.挂号状态, 0) = 0 Then
                        n_时段剩余 := 1;
                      Else
                        n_时段剩余 := 0;
                      End If;
                      v_Temp := v_Temp || Gettimexml(r_Time.开始时间, r_Time.终止时间, 1, n_时段剩余);
                    Else
                      If d_时段开始 Is Null Then
                        n_时段总数 := 1;
                        n_时段剩余 := 0;
                        d_时段开始 := r_Time.开始时间;
                        d_时段结束 := r_Time.开始时间 + n_时间间隔 / 24 / 60;
                        If d_加号时间 < d_时段结束 Then
                          d_时段结束 := d_加号时间;
                        End If;
                      Else
                        If r_Time.开始时间 >= d_时段结束 Then
                          v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                        
                          n_时段总数 := 1;
                          n_时段剩余 := 0;
                          d_时段开始 := r_Time.开始时间;
                          d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                          If d_加号时间 < d_时段结束 Then
                            d_时段结束 := d_加号时间;
                          End If;
                        Else
                          n_时段总数 := n_时段总数 + 1;
                        End If;
                      End If;
                    
                      If Nvl(r_Time.挂号状态, 0) = 0 Then
                        n_时段剩余 := n_时段剩余 + 1;
                      End If;
                    End If;
                  End If;
                End Loop;
              
                If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
                  v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                End If;
              End If;
              v_Temp := v_Temp || '</SPANLIST>';
            End If;
          Else
            --合作单位预约挂号
            If r_No.预约控制 = 2 Then
              n_禁用 := 1;
            Else
              Begin
                Select 控制方式
                Into n_合约模式
                From 临床出诊挂号控制记录
                Where 记录id = r_No.记录id And 类型 = 1 And 性质 = 1 And 名称 = v_合作单位 And Rownum < 2;
              Exception
                When Others Then
                  n_合约模式 := 4;
              End;
            
              If n_合约模式 = 0 Then
                n_禁用 := 1;
              Elsif n_合约模式 = 1 Or n_合约模式 = 2 Then
                Select 数量
                Into n_合约总数量
                From 临床出诊挂号控制记录
                Where 记录id = r_No.记录id And 类型 = 1 And 性质 = 1 And 名称 = v_合作单位;
                If n_合约模式 = 1 Then
                  n_合约总数量 := Floor(r_No.限约数 * n_合约总数量 / 100);
                End If;
              
                Select Count(1)
                Into n_合约已挂数
                From 病人挂号记录
                Where 号别 = r_No.号码 And 记录状态 = 1 And 合作单位 = v_合作单位 And 发生时间 Between Trunc(d_日期) And
                      Trunc(d_日期 + 1) - 1 / 60 / 60 / 24;
              
                n_合约剩余数量 := n_合约总数量 - n_合约已挂数;
              
                If Nvl(r_No.限约数, r_No.限号数) - Nvl(r_No.已约数, 0) < n_合约剩余数量 Then
                  v_剩余数量 := Nvl(r_No.限约数, r_No.限号数) - Nvl(r_No.已约数, 0);
                Else
                  v_剩余数量 := n_合约剩余数量;
                End If;
                n_已挂数 := r_No.已约数;
              
                If r_No.分时段 = 1 Then
                  v_Temp     := '<SPANLIST>';
                  n_Exists   := 0;
                  n_时段总数 := 0;
                  n_时段剩余 := 0;
                  d_时段开始 := Null;
                  d_时段结束 := Null;
                  Select Max(终止时间) Into d_加号时间 From 临床出诊序号控制 Where 记录id = r_No.记录id;
                  If r_No.序号控制 = 1 Then
                    --分时段,序号控制
                    For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态
                                   From 临床出诊序号控制
                                   Where 记录id = r_No.记录id And 是否预约 = 1 And
                                         (r_No.记录性质 = 1 And 开始时间 Not Between Nvl(r_No.替诊开始时间, Sysdate) And
                                         Nvl(r_No.替诊终止时间, Sysdate - 1) Or
                                         r_No.记录性质 = 2 And 开始时间 Between Nvl(r_No.替诊开始时间, Sysdate) And
                                         Nvl(r_No.替诊终止时间, Sysdate - 1)) And Nvl(是否停诊, 0) = 0) Loop
                    
                      If r_Time.开始时间 > Sysdate Then
                        If Nvl(n_时间间隔, 0) = 0 Then
                          If Nvl(r_Time.挂号状态, 0) = 0 Then
                            n_时段剩余 := 1;
                            n_Exists   := n_Exists + 1;
                          Else
                            n_时段剩余 := 0;
                          End If;
                          v_Temp := v_Temp || Gettimexml(r_Time.开始时间, r_Time.终止时间, 1, n_时段剩余);
                        Else
                          If d_时段开始 Is Null Then
                            n_时段总数 := 1;
                            n_时段剩余 := 0;
                            d_时段开始 := r_Time.开始时间;
                            d_时段结束 := r_Time.开始时间 + n_时间间隔 / 24 / 60;
                            If d_加号时间 < d_时段结束 Then
                              d_时段结束 := d_加号时间;
                            End If;
                          Else
                            If r_Time.开始时间 >= d_时段结束 Then
                              v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                            
                              n_时段总数 := 1;
                              n_时段剩余 := 0;
                              d_时段开始 := r_Time.开始时间;
                              d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                              If d_加号时间 < d_时段结束 Then
                                d_时段结束 := d_加号时间;
                              End If;
                            Else
                              n_时段总数 := n_时段总数 + 1;
                            End If;
                          End If;
                        
                          If Nvl(r_Time.挂号状态, 0) = 0 Then
                            n_时段剩余 := 1;
                            n_Exists   := n_Exists + 1;
                          End If;
                        End If;
                      End If;
                    End Loop;
                  
                    If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
                      v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                    End If;
                  
                    If n_Exists < To_Number(v_剩余数量) Then
                      v_剩余数量 := n_Exists;
                    End If;
                  Else
                    --分时段,非序号控制
                    For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态, 数量
                                   From 临床出诊序号控制
                                   Where 记录id = r_No.记录id And 预约顺序号 Is Null And 是否预约 = 1 And
                                         (r_No.记录性质 = 1 And 开始时间 Not Between Nvl(r_No.替诊开始时间, Sysdate) And
                                         Nvl(r_No.替诊终止时间, Sysdate - 1) Or
                                         r_No.记录性质 = 2 And 开始时间 Between Nvl(r_No.停诊开始时间, Sysdate) And
                                         Nvl(r_No.停诊终止时间, Sysdate - 1))) Loop
                    
                      Select Count(1)
                      Into n_时段已挂
                      From 临床出诊序号控制
                      Where 记录id = r_No.记录id And 序号 = r_Time.序号 And 预约顺序号 Is Not Null And Nvl(挂号状态, 0) <> 0;
                    
                      If r_Time.开始时间 > Sysdate Then
                        If Nvl(n_时间间隔, 0) = 0 Then
                          n_时段剩余 := r_Time.数量 - n_时段已挂;
                          n_Exists   := n_Exists + n_时段剩余;
                          v_Temp     := v_Temp || Gettimexml(r_Time.开始时间, r_Time.终止时间, r_Time.数量, n_时段剩余);
                        Else
                          If d_时段开始 Is Null Then
                            n_时段总数 := r_Time.数量;
                            n_时段剩余 := 0;
                            d_时段开始 := r_Time.开始时间;
                            d_时段结束 := r_Time.开始时间 + n_时间间隔 / 24 / 60;
                            If d_加号时间 < d_时段结束 Then
                              d_时段结束 := d_加号时间;
                            End If;
                          Else
                            If r_Time.开始时间 >= d_时段结束 Then
                              v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                            
                              n_时段总数 := r_Time.数量;
                              n_时段剩余 := 0;
                              d_时段开始 := r_Time.开始时间;
                              d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                              If d_加号时间 < d_时段结束 Then
                                d_时段结束 := d_加号时间;
                              End If;
                            Else
                              n_时段总数 := n_时段总数 + r_Time.数量;
                            End If;
                          End If;
                        
                          n_时段剩余 := n_时段剩余 + (r_Time.数量 - n_时段已挂);
                          n_Exists   := n_Exists + (r_Time.数量 - n_时段已挂);
                        End If;
                      End If;
                    End Loop;
                  
                    If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
                      v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                    End If;
                  
                    If n_Exists < To_Number(v_剩余数量) Then
                      v_剩余数量 := n_Exists;
                    End If;
                  End If;
                  v_Temp := v_Temp || '</SPANLIST>';
                End If;
              Elsif n_合约模式 = 3 Then
                If r_No.分时段 = 0 Then
                  If r_No.序号控制 = 0 Then
                    n_禁用 := 1;
                  Else
                    --序号控制不分时段
                    n_已挂数   := 0;
                    v_剩余数量 := 0;
                    For r_合作 In (Select 序号
                                 From 临床出诊挂号控制记录
                                 Where 记录id = r_No.记录id And 类型 = 1 And 性质 = 1 And 名称 = v_合作单位) Loop
                    
                      Select Count(1)
                      Into n_Exists
                      From 临床出诊序号控制
                      Where 记录id = r_No.记录id And 序号 = r_合作.序号 And 是否预约 = 1 And Nvl(挂号状态, 0) = 0 And
                            (r_No.记录性质 = 1 And 开始时间 Not Between Nvl(r_No.替诊开始时间, Sysdate) And
                            Nvl(r_No.替诊终止时间, Sysdate - 1) Or
                            r_No.记录性质 = 2 And 开始时间 Between Nvl(r_No.替诊开始时间, Sysdate) And Nvl(r_No.替诊终止时间, Sysdate - 1)) And
                            Nvl(是否停诊, 0) = 0;
                    
                      If n_Exists = 1 Then
                        v_剩余数量 := v_剩余数量 + 1;
                      Else
                        n_已挂数 := n_已挂数 + 1;
                      End If;
                    End Loop;
                  
                    If Nvl(r_No.限约数, r_No.限号数) - Nvl(r_No.已约数, 0) < v_剩余数量 Then
                      v_剩余数量 := Nvl(r_No.限约数, r_No.限号数) - Nvl(r_No.已约数, 0);
                    End If;
                  End If;
                Else
                  v_Temp     := '<SPANLIST>';
                  n_已挂数   := 0;
                  v_剩余数量 := 0;
                  n_时段总数 := 0;
                  n_时段剩余 := 0;
                  d_时段开始 := Null;
                  d_时段结束 := Null;
                  Select Max(终止时间) Into d_加号时间 From 临床出诊序号控制 Where 记录id = r_No.记录id;
                  If r_No.序号控制 = 0 Then
                    --分时段,非序号控制
                    For r_合作 In (Select 序号, 数量
                                 From 临床出诊挂号控制记录
                                 Where 记录id = r_No.记录id And 类型 = 1 And 性质 = 1 And 名称 = v_合作单位) Loop
                    
                      Select Count(1), Max(开始时间), Max(终止时间)
                      Into n_时段已挂, d_开始时间, d_终止时间
                      From 临床出诊序号控制
                      Where 记录id = r_No.记录id And 序号 = r_合作.序号 And 预约顺序号 Is Not Null And Nvl(挂号状态, 0) <> 0 And
                            (r_No.记录性质 = 1 And 开始时间 Not Between Nvl(r_No.替诊开始时间, Sysdate) And
                            Nvl(r_No.替诊终止时间, Sysdate - 1) Or
                            r_No.记录性质 = 2 And 开始时间 Between Nvl(r_No.替诊开始时间, Sysdate) And Nvl(r_No.替诊终止时间, Sysdate - 1));
                    
                      n_已挂数   := n_已挂数 + n_时段已挂;
                      v_剩余数量 := v_剩余数量 + r_合作.数量 - n_时段已挂;
                    
                      If d_开始时间 > Sysdate Then
                        If Nvl(n_时间间隔, 0) = 0 Then
                          n_时段剩余 := r_合作.数量 - n_时段已挂;
                          v_Temp     := v_Temp || Gettimexml(d_开始时间, d_终止时间, r_合作.数量, n_时段剩余);
                        Else
                          If d_时段开始 Is Null Then
                            n_时段总数 := r_合作.数量;
                            n_时段剩余 := 0;
                            d_时段开始 := d_开始时间;
                            d_时段结束 := d_开始时间 + n_时间间隔 / 24 / 60;
                            If d_加号时间 < d_时段结束 Then
                              d_时段结束 := d_加号时间;
                            End If;
                          Else
                            If d_开始时间 >= d_时段结束 Then
                              v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                            
                              n_时段总数 := r_合作.数量;
                              n_时段剩余 := 0;
                              d_时段开始 := d_开始时间;
                              d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                              If d_加号时间 < d_时段结束 Then
                                d_时段结束 := d_加号时间;
                              End If;
                            Else
                              n_时段总数 := n_时段总数 + r_合作.数量;
                            End If;
                          End If;
                        
                          n_时段剩余 := n_时段剩余 + r_合作.数量 - n_时段已挂;
                        End If;
                      End If;
                    End Loop;
                  
                    If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
                      v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                    End If;
                  Else
                    --分时段,序号控制
                    For r_合作 In (Select 序号
                                 From 临床出诊挂号控制记录
                                 Where 记录id = r_No.记录id And 类型 = 1 And 性质 = 1 And 名称 = v_合作单位) Loop
                    
                      Select Max(1), Max(开始时间), Max(终止时间)
                      Into n_Exists, d_开始时间, d_终止时间
                      From 临床出诊序号控制
                      Where 记录id = r_No.记录id And 序号 = r_合作.序号 And Nvl(挂号状态, 0) = 0 And
                            (r_No.记录性质 = 1 And 开始时间 Not Between Nvl(r_No.替诊开始时间, Sysdate) And
                            Nvl(r_No.替诊终止时间, Sysdate - 1) Or
                            r_No.记录性质 = 2 And 开始时间 Between Nvl(r_No.替诊开始时间, Sysdate) And Nvl(r_No.替诊终止时间, Sysdate - 1)) And
                            Nvl(是否停诊, 0) = 0;
                    
                      If n_Exists = 1 Then
                        v_剩余数量 := v_剩余数量 + 1;
                      Else
                        n_已挂数 := n_已挂数 + 1;
                      End If;
                    
                      If d_开始时间 > Sysdate Then
                        If Nvl(n_时间间隔, 0) = 0 Then
                          If n_Exists = 1 Then
                            n_时段剩余 := 1;
                          Else
                            n_时段剩余 := 0;
                          End If;
                          v_Temp := v_Temp || Gettimexml(d_开始时间, d_终止时间, 1, n_时段剩余);
                        Else
                          If d_时段开始 Is Null Then
                            n_时段总数 := 1;
                            n_时段剩余 := 0;
                            d_时段开始 := d_开始时间;
                            d_时段结束 := d_开始时间 + n_时间间隔 / 24 / 60;
                            If d_加号时间 < d_时段结束 Then
                              d_时段结束 := d_加号时间;
                            End If;
                          Else
                            If d_开始时间 >= d_时段结束 Then
                              v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                            
                              n_时段总数 := 1;
                              n_时段剩余 := 0;
                              d_时段开始 := d_开始时间;
                              d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                              If d_加号时间 < d_时段结束 Then
                                d_时段结束 := d_加号时间;
                              End If;
                            Else
                              n_时段总数 := n_时段总数 + 1;
                            End If;
                          End If;
                        
                          If n_Exists = 1 Then
                            n_时段剩余 := n_时段剩余 + 1;
                          End If;
                        End If;
                      End If;
                    End Loop;
                  
                    If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
                      v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                    End If;
                  End If;
                  v_Temp := v_Temp || '</SPANLIST>';
                End If;
              Elsif n_合约模式 = 4 Then
                If r_No.分时段 = 0 Then
                  n_已挂数   := r_No.已约数;
                  v_剩余数量 := Nvl(r_No.限约数, r_No.限号数) - Nvl(r_No.已约数, 0);
                Else
                  --分时段
                  v_Temp     := '<SPANLIST>';
                  n_已挂数   := 0;
                  v_剩余数量 := 0;
                  n_时段总数 := 0;
                  n_时段剩余 := 0;
                  d_时段开始 := Null;
                  d_时段结束 := Null;
                  Select Max(终止时间) Into d_加号时间 From 临床出诊序号控制 Where 记录id = r_No.记录id;
                  If r_No.序号控制 = 0 Then
                    --非序号控制分时段预约
                    For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态, 数量
                                   From 临床出诊序号控制
                                   Where 记录id = r_No.记录id And 预约顺序号 Is Null And 是否预约 = 1 And
                                         (r_No.记录性质 = 1 And 开始时间 Not Between Nvl(r_No.替诊开始时间, Sysdate) And
                                         Nvl(r_No.替诊终止时间, Sysdate - 1) Or
                                         r_No.记录性质 = 2 And 开始时间 Between Nvl(r_No.停诊开始时间, Sysdate) And
                                         Nvl(r_No.停诊终止时间, Sysdate - 1))) Loop
                    
                      Select Count(1)
                      Into n_时段已挂
                      From 临床出诊序号控制
                      Where 记录id = r_No.记录id And 序号 = r_Time.序号 And 预约顺序号 Is Not Null And Nvl(挂号状态, 0) <> 0;
                    
                      n_已挂数   := n_已挂数 + n_时段已挂;
                      v_剩余数量 := v_剩余数量 + n_时段剩余;
                    
                      If r_Time.开始时间 > Sysdate Then
                        If Nvl(n_时间间隔, 0) = 0 Then
                          n_时段剩余 := r_Time.数量 - n_时段已挂;
                          v_Temp     := v_Temp || Gettimexml(r_Time.开始时间, r_Time.终止时间, r_Time.数量, n_时段剩余);
                        Else
                          If d_时段开始 Is Null Then
                            n_时段总数 := r_Time.数量;
                            n_时段剩余 := 0;
                            d_时段开始 := r_Time.开始时间;
                            d_时段结束 := r_Time.开始时间 + n_时间间隔 / 24 / 60;
                            If d_加号时间 < d_时段结束 Then
                              d_时段结束 := d_加号时间;
                            End If;
                          Else
                            If r_Time.开始时间 >= d_时段结束 Then
                              v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                            
                              n_时段总数 := r_Time.数量;
                              n_时段剩余 := 0;
                              d_时段开始 := r_Time.开始时间;
                              d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                              If d_加号时间 < d_时段结束 Then
                                d_时段结束 := d_加号时间;
                              End If;
                            Else
                              n_时段总数 := n_时段总数 + r_Time.数量;
                            End If;
                          End If;
                        
                          n_时段剩余 := n_时段剩余 + (r_Time.数量 - n_时段已挂);
                        End If;
                      End If;
                    End Loop;
                  
                    If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
                      v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                    End If;
                  Else
                    For r_Time In (Select 序号, 开始时间, 终止时间, 挂号状态
                                   From 临床出诊序号控制
                                   Where 记录id = r_No.记录id And 是否预约 = 1 And
                                         (r_No.记录性质 = 1 And 开始时间 Not Between Nvl(r_No.替诊开始时间, Sysdate) And
                                         Nvl(r_No.替诊终止时间, Sysdate - 1) Or
                                         r_No.记录性质 = 2 And 开始时间 Between Nvl(r_No.替诊开始时间, Sysdate) And
                                         Nvl(r_No.替诊终止时间, Sysdate - 1)) And Nvl(是否停诊, 0) = 0) Loop
                    
                      If Nvl(r_Time.挂号状态, 0) = 0 Then
                        v_剩余数量 := v_剩余数量 + 1;
                      Else
                        n_已挂数 := n_已挂数 + 1;
                      End If;
                    
                      If r_Time.开始时间 > Sysdate Then
                        If Nvl(n_时间间隔, 0) = 0 Then
                          If Nvl(r_Time.挂号状态, 0) = 0 Then
                            n_时段剩余 := 1;
                          Else
                            n_时段剩余 := 0;
                          End If;
                          v_Temp := v_Temp || Gettimexml(r_Time.开始时间, r_Time.终止时间, 1, n_时段剩余);
                        Else
                          If d_时段开始 Is Null Then
                            n_时段总数 := 1;
                            n_时段剩余 := 0;
                            d_时段开始 := r_Time.开始时间;
                            d_时段结束 := r_Time.开始时间 + n_时间间隔 / 24 / 60;
                            If d_加号时间 < d_时段结束 Then
                              d_时段结束 := d_加号时间;
                            End If;
                          Else
                            If r_Time.开始时间 >= d_时段结束 Then
                              v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                            
                              n_时段总数 := 1;
                              n_时段剩余 := 0;
                              d_时段开始 := r_Time.开始时间;
                              d_时段结束 := d_时段开始 + n_时间间隔 / 24 / 60;
                              If d_加号时间 < d_时段结束 Then
                                d_时段结束 := d_加号时间;
                              End If;
                            Else
                              n_时段总数 := n_时段总数 + 1;
                            End If;
                          End If;
                        
                          If Nvl(r_Time.挂号状态, 0) = 0 Then
                            n_时段剩余 := n_时段剩余 + 1;
                          End If;
                        End If;
                      End If;
                    End Loop;
                  
                    If Nvl(n_时间间隔, 0) <> 0 And n_时段总数 <> 0 Then
                      v_Temp := v_Temp || Gettimexml(d_时段开始, d_时段结束, n_时段总数, n_时段剩余);
                    End If;
                  End If;
                  v_Temp := v_Temp || '</SPANLIST>';
                End If;
              End If;
            End If;
          End If;
        End If;
      End If;
    
      If Not (r_No.分时段 = 1 And r_No.序号控制 = 1) Then
        If d_日期 Between Nvl(r_No.替诊开始时间, Sysdate) And Nvl(r_No.替诊终止时间, Sysdate - 1) Then
          n_禁用 := 1;
        End If;
        If d_日期 Between Nvl(r_No.停诊开始时间, Sysdate) And Nvl(r_No.停诊终止时间, Sysdate - 1) Then
          n_禁用 := 1;
        End If;
      End If;
    
      If Nvl(n_禁用, 0) = 0 Then
        n_合计金额 := 0;
        For r_Fee In (Select b.现价, a.从项数次
                      From 收费从属项目 A, 收费价目 B
                      Where a.从项id = b.收费细目id And a.主项id = r_No.项目id And d_日期 Between b.执行日期 And
                            Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                            (b.价格等级 = v_普通等级 Or
                            (b.价格等级 Is Null And Not Exists
                             (Select 1
                               From 收费价目
                               Where 收费细目id = b.收费细目id And 价格等级 = v_普通等级 And d_日期 Between 执行日期 And
                                     Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                      Union All
                      Select b.现价, 1 As 从项数次
                      From 收费项目目录 A, 收费价目 B
                      Where a.Id = b.收费细目id And a.Id = r_No.项目id And d_日期 Between b.执行日期 And
                            Nvl(b.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                            (b.价格等级 = v_普通等级 Or
                            (b.价格等级 Is Null And Not Exists
                             (Select 1
                               From 收费价目
                               Where 收费细目id = b.收费细目id And 价格等级 = v_普通等级 And d_日期 Between 执行日期 And
                                     Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))) Loop
          n_合计金额 := n_合计金额 + r_Fee.现价 * r_Fee.从项数次;
        End Loop;
      
        v_时间段  := To_Char(r_No.开始时间, 'HH24:MI') || '-' || To_Char(r_No.终止时间, 'HH24:MI');
        c_Xmlmain := '<HB>';
        c_Xmlmain := c_Xmlmain || '<CZJLID>' || r_No.记录id || '</CZJLID>';
        c_Xmlmain := c_Xmlmain || '<HM>' || r_No.号码 || '</HM>';
        c_Xmlmain := c_Xmlmain || '<YSID>' || r_No.医生id || '</YSID>';
        c_Xmlmain := c_Xmlmain || '<YS>' || r_No.医生姓名 || '</YS>';
        c_Xmlmain := c_Xmlmain || '<KSID>' || r_No.科室id || '</KSID>';
        c_Xmlmain := c_Xmlmain || '<KSMC>' || r_No.科室名称 || '</KSMC>';
        c_Xmlmain := c_Xmlmain || '<ZC>' || r_No.职称 || '</ZC>';
        c_Xmlmain := c_Xmlmain || '<XMID>' || r_No.项目id || '</XMID>';
        c_Xmlmain := c_Xmlmain || '<XMMC>' || r_No.项目名称 || '</XMMC>';
        c_Xmlmain := c_Xmlmain || '<PRICE>' || n_合计金额 || '</PRICE>';
        c_Xmlmain := c_Xmlmain || '<HL>' || r_No.号类 || '</HL>';
        c_Xmlmain := c_Xmlmain || '<FSD>' || r_No.分时段 || '</FSD>';
        c_Xmlmain := c_Xmlmain || '<HBTIME>' || v_时间段 || '</HBTIME>';
        c_Xmlmain := c_Xmlmain || '<FWMC>' || r_No.排班 || '</FWMC>';
        If Trunc(Sysdate) = Trunc(d_日期) Or r_No.已约数 < r_No.限约数 Then
          c_Xmlmain := c_Xmlmain || '<YGHS>' || n_已挂数 || '</YGHS>';
          c_Xmlmain := c_Xmlmain || '<SYHS>' || v_剩余数量 || '</SYHS>';
          c_Xmlmain := c_Xmlmain || v_Temp;
        Else
          c_Xmlmain := c_Xmlmain || '<YGHS>' || r_No.已约数 || '</YGHS>';
          c_Xmlmain := c_Xmlmain || '<SYHS>' || 0 || '</SYHS>';
        End If;
        c_Xmlmain := c_Xmlmain || '</HB>';
        v_Xmlmain := v_Xmlmain || c_Xmlmain;
      End If;
    End If;
  End Loop;
  v_Xmlmain := '<GROUP>' || '<RQ>' || To_Char(d_日期, 'yyyy-mm-dd') || '</RQ>' || '<HBLIST>' || v_Xmlmain || '</HBLIST>' ||
               '</GROUP>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Xmlmain)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getnolist;
/

--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
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
                         Where 安排id = p.Id And (d_发生时间_In Between 生效时间 + 0 And 失效时间) And 审核时间 Is Not Null) And Not Exists
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
                               a.失效时间 And b.号码 = 号码_In
                         Group By 安排id) E
                  Where p.安排id = c.Id And p.Id = b.计划id(+) And p.生效时间 = e.生效 And p.安排id = e.安排id And
                        Nvl(p.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) = Nvl(e.生效, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                        b.限制项目(+) = Decode(To_Char(d_发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四',
                                           '6', '周五', '7', '周六', Null) And (d_发生时间_In Between p.生效时间 + 0 And p.失效时间) And
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
              失效时间 And Rownum < 2
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
                                  d.失效时间)
                     Union All
                     Select c.序号, 数量
                     From 挂号安排计划 A, 挂号安排 D, 合作单位计划控制 C,
                          (Select Max(a.生效时间) As 生效, 安排id
                            From 挂号安排计划 A, 挂号安排 B
                            Where a.安排id = b.Id And a.审核时间 Is Not Null And
                                  发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                  a.失效时间 And b.号码 = 号码_In
                            Group By 安排id) E
                     Where a.安排id = d.Id And a.审核时间 Is Not Null And d.号码 = 号码_In And a.安排id = e.安排id And
                           Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) =
                           Nvl(e.生效, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                           Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',
                                  '7', '周六', Null) = c.限制项目(+) And a.Id = c.计划id And c.合作单位 = 合作单位_In And c.序号 = n_序号 And
                           发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                           a.失效时间) Loop
      
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

--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
Create Or Replace Procedure Zl_三方机构挂号_Delete
(
  单据号_In     门诊费用记录.No%Type,
  交易流水号_In 病人预交记录.交易流水号%Type,
  交易说明_In   病人预交记录.交易说明%Type,
  退号时间_In   门诊费用记录.登记时间%Type := Null,
  预交id_In     病人预交记录.Id%Type := Null
) As
  v_Error Varchar2(255);
  Err_Custom Exception;

  --该游标用于判断是否单独收病历费,及挂号汇总表处理
  Cursor c_Registinfo
  (
    v_状态     病人挂号记录.记录状态%Type,
    v_性质     病人挂号记录.记录性质%Type,
    v_无效单据 Number := 0
  ) Is
  
    Select a.发生时间, a.登记时间, b.收费细目id As 项目id, a.执行部门id As 科室id, a.执行人 As 医生姓名, c.Id As 医生id, a.号别 As 号码
    From 病人挂号记录 A, 门诊费用记录 B, 人员表 C
    Where a.记录性质 = Decode(v_无效单据, 0, v_性质, a.记录性质) And b.记录性质 = 4 And a.记录状态 = v_状态 And a.No = 单据号_In And a.No = b.No And
          a.执行人 = c.姓名(+) And b.序号 = 1 And Rownum = 1;

  r_Registrow c_Registinfo%RowType;

  --该光标用于处理人员缴款余额中退的不同结算方式的金额
  Cursor c_Opermoney Is
    Select Distinct b.结算方式, b.冲预交
    From 门诊费用记录 A, 病人预交记录 B
    Where a.结帐id = b.结帐id And a.No = 单据号_In And a.记录性质 = 4 And a.记录状态 = 3 And b.记录性质 = 4 And b.记录状态 = 3 And
          Nvl(b.冲预交, 0) <> 0;

  n_执行状态       病人挂号记录.执行状态%Type;
  n_打印id         票据打印内容.Id%Type;
  n_结帐id         门诊费用记录.结帐id%Type;
  n_原结帐id       病人预交记录.结帐id%Type;
  n_病人id         病人信息.病人id%Type;
  n_返回值         病人余额.预交余额%Type;
  n_分诊台签到排队 Number;
  n_预约挂号       Number;
  n_无效单据       Number; --无效单据没有产生费用单据
  n_挂号生成队列   Number;
  n_Count          Number;
  n_组id           财务缴款分组.Id%Type;
  d_退号时间       Date;
  v_操作员编号     人员表.编号%Type;
  v_操作员姓名     人员表.姓名%Type;
  v_合作单位       合作单位挂号汇总.合作单位%Type;
  n_预约状态       病人挂号记录.预约%Type;
  v_Temp           Varchar2(100);
  d_登记时间       病人挂号记录.登记时间%Type;
  v_号别           病人挂号记录.号别%Type;
  n_号序           病人挂号记录.号序%Type;
  d_预约时间       病人挂号记录.预约时间%Type;
  n_合作单位限制   Number(18);
  n_预约生成队列   Number;
  n_记录性质       Number;
  n_状态           Number;
  n_退号重用       Number(3);
  n_挂号排班模式   Number;
  n_记帐           门诊费用记录.记帐费用%Type;
  n_挂号id         病人挂号记录.Id%Type;
  n_已结帐         Number;
  n_返回额         病人余额.费用余额%Type;
  n_预交支付       Number(3);
  n_正常支付       Number(3);
  n_安排id         挂号安排.Id%Type;
  n_计划id         挂号安排计划.Id%Type;
  v_时间段         时间段.时间段%Type;
  d_检查开始时间   Date;
  d_检查结束时间   Date;
  d_启用时间       Date;
  n_出诊记录id     Number(18);
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

  Procedure Zl_三方机构挂号_出诊_Delete
  (
    单据号_In     门诊费用记录.No%Type,
    交易流水号_In 病人预交记录.交易流水号%Type,
    交易说明_In   病人预交记录.交易说明%Type,
    退号时间_In   门诊费用记录.登记时间%Type := Null,
    预交id_In     病人预交记录.Id%Type := Null
  ) As
    v_Error Varchar2(255);
    Err_Custom Exception;
  
    --该游标用于判断是否单独收病历费,及挂号汇总表处理
    Cursor c_Registinfo
    (
      v_状态     病人挂号记录.记录状态%Type,
      v_性质     病人挂号记录.记录性质%Type,
      v_无效单据 Number := 0
    ) Is
      Select a.发生时间, a.登记时间, b.项目id, b.科室id, b.医生姓名, b.医生id, b.Id As 记录id, a.号别 As 号码
      From 病人挂号记录 A, 临床出诊记录 B
      Where a.记录性质 = Decode(v_无效单据, 0, v_性质, a.记录性质) And a.记录状态 = v_状态 And a.No = 单据号_In And a.出诊记录id = b.Id And
            Rownum < 2;
  
    r_Registrow c_Registinfo%RowType;
  
    --该光标用于处理人员缴款余额中退的不同结算方式的金额
    Cursor c_Opermoney Is
      Select Distinct b.结算方式, b.冲预交
      From 门诊费用记录 A, 病人预交记录 B
      Where a.结帐id = b.结帐id And a.No = 单据号_In And a.记录性质 = 4 And a.记录状态 = 3 And b.记录性质 = 4 And b.记录状态 = 3 And
            Nvl(b.冲预交, 0) <> 0;
  
    n_执行状态       病人挂号记录.执行状态%Type;
    n_打印id         票据打印内容.Id%Type;
    n_结帐id         门诊费用记录.结帐id%Type;
    n_原结帐id       病人预交记录.结帐id%Type;
    n_病人id         病人信息.病人id%Type;
    n_返回值         病人余额.预交余额%Type;
    n_分诊台签到排队 Number;
    n_预约挂号       Number;
    n_无效单据       Number; --无效单据没有产生费用单据
    n_挂号生成队列   Number;
    n_Count          Number;
    n_组id           财务缴款分组.Id%Type;
    d_退号时间       Date;
    v_操作员编号     人员表.编号%Type;
    v_操作员姓名     人员表.姓名%Type;
    v_合作单位       合作单位挂号汇总.合作单位%Type;
    n_预约状态       病人挂号记录.预约%Type;
    v_Temp           Varchar2(100);
    d_登记时间       病人挂号记录.登记时间%Type;
    v_号别           病人挂号记录.号别%Type;
    n_号序           病人挂号记录.号序%Type;
    d_预约时间       病人挂号记录.预约时间%Type;
    n_合作单位限制   Number(18);
    n_预约生成队列   Number;
    n_记录性质       Number;
    n_状态           Number;
    n_退号重用       Number(3);
    n_挂号id         病人挂号记录.Id%Type;
    n_记帐           门诊费用记录.记帐费用%Type;
    n_记录id         临床出诊记录.Id%Type;
  
    n_已结帐   Number;
    n_返回额   病人余额.费用余额%Type;
    n_预交支付 Number(3);
    n_正常支付 Number(3);
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
    --部门ID,部门名称;人员ID,人员编号,人员姓名
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Error := '当前操作人员未设置对应的人员关系,不能继续。';
      Raise Err_Custom;
    End If;
    v_操作员编号 := Zl_操作员(1, v_Temp);
    v_操作员姓名 := Zl_操作员(2, v_Temp);
  
    n_组id := Zl_Get组id(v_操作员姓名);
  
    d_退号时间 := 退号时间_In;
    If d_退号时间 Is Null Then
      d_退号时间 := Sysdate;
    End If;
  
    --首先判断要退号/取消预约的记录是否存在
    Begin
      Select Decode(记录性质, 2, 1, 0), 记录性质, 登记时间, 号别, 号序, Nvl(预约时间, 发生时间), 合作单位, Nvl(预约, 0), Decode(记录状态, 0, 1, 0), 出诊记录id
      Into n_预约挂号, n_记录性质, d_登记时间, v_号别, n_号序, d_预约时间, v_合作单位, n_预约状态, n_无效单据, n_记录id
      From 病人挂号记录
      Where NO = 单据号_In And 记录状态 In (0, 1) And Rownum < 2;
    Exception
      When Others Then
        n_预约挂号 := -1;
    End;
  
    If n_预约挂号 = -1 Then
      v_Error := '单据可能已经被退号或单据输入错误!';
      Raise Err_Custom;
    End If;
  
    Begin
      Select Nvl(记帐费用, 0), Decode(Sign(Nvl(结帐id, 0)), 0, 0, 1)
      Into n_记帐, n_已结帐
      From 门诊费用记录
      Where 记录性质 = 4 And NO = 单据号_In And 记录状态 In (1, 3) And Rownum < 2;
    Exception
      When Others Then
        n_记帐   := 0;
        n_已结帐 := 0;
    End;
  
    --预约检查是否添加合作单位控制
    --如果设置了合作单位控制 则
    Select Count(0) Into n_合作单位限制 From 临床出诊挂号控制记录 Where 类型 = 1 And 性质 = 1 And Rownum < 2;
    --更新挂号序号状态
    n_退号重用 := Zl_To_Number(zl_GetSysParameter('已退序号允许挂号', 1111));
    If n_退号重用 = 0 Then
      Update 临床出诊序号控制 Set 挂号状态 = 4 Where 记录id = n_记录id And (序号 = n_号序 Or 备注 = To_Char(n_号序));
    Else
      Update 临床出诊序号控制
      Set 挂号状态 = 0, 类型 = Null, 名称 = Null, 操作员姓名 = Null, 工作站名称 = Null
      Where 记录id = n_记录id And (序号 = n_号序 Or 备注 = To_Char(n_号序));
    End If;
    If Nvl(n_预约挂号, 0) = 1 Or Nvl(n_无效单据, 0) = 1 Then
      If Nvl(n_无效单据, 0) = 0 Then
        --N天内不能取消预约号
        n_Count := Zl_To_Number(zl_GetSysParameter('N天内不能取消预约号', 1111));
        If n_Count <> 0 Then
          If Trunc(Sysdate - n_Count) < d_登记时间 Then
            v_Error := '不能退掉预约在' || To_Char(Trunc(Sysdate - n_Count), 'yyyy-mm-dd') || '以前的预约单!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      n_状态 := Case n_无效单据
                When 1 Then
                 0
                Else
                 1
              End;
      --减少已约数
      Open c_Registinfo(n_状态, 2, n_无效单据);
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已挂数 = Nvl(已挂数, 0) - Decode(n_预约状态, 0, 1, 0)
      Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
            Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
            (号码 = r_Registrow.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数, 已挂数)
        Values
          (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
           Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -n_预约状态, Decode(n_预约状态, 0, 1, 0));
      End If;
    
      Update 临床出诊记录
      Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已挂数 = Nvl(已挂数, 0) - Decode(n_预约状态, 0, 1, 0)
      Where ID = n_记录id;
      Close c_Registinfo;
    
      If Nvl(n_无效单据, 0) = 0 Then
        --删除门诊费用记录
        Delete From 门诊费用记录 Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 0;
        --如果预约生成队列时需要清除队列
        n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
        If Nvl(n_挂号生成队列, 0) = 1 Then
          n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
          If Nvl(n_预约生成队列, 0) = 1 Then
            --要删除队列
            For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
              Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
            End Loop;
          End If;
        End If;
      End If;
    Else
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    
      --更新挂号序号状态
    
      --病人就诊状态
      Select 病人id
      Into n_病人id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 序号 = 1;
    
      If n_病人id Is Not Null Then
        Update 病人信息 Set 就诊状态 = 0, 就诊诊室 = Null Where 病人id = n_病人id;
        --删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理
      End If;
    
      --门诊费用记录
      Insert Into 门诊费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位,
         付数, 数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间,
         登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
        Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
               收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
               v_操作员编号, v_操作员姓名, 发生时间, d_退号时间, n_结帐id,
               Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额, 摘要,
               Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
    
      --原始记录
      If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
        Update 门诊费用记录
        Set 记录状态 = 3, 结帐id = n_结帐id, 结帐金额 = 实收金额
        Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
      Else
        Update 门诊费用记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
      End If;
    
      n_原结帐id := 0;
      If n_记帐 = 0 Then
        --获取结帐ID
        Select Nvl(结帐id, 0)
        Into n_原结帐id
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Rownum < 2;
      End If;
    
      If n_记帐 = 1 Then
        --记帐
        For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                     From 门诊费用记录
                     Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Nvl(记帐费用, 0) = 1) Loop
          --病人余额
          Update 病人余额
          Set 费用余额 = Nvl(费用余额, 0) - Nvl(c_费用.实收金额, 0)
          Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1
          Returning 费用余额 Into n_返回额;
        
          If Sql%RowCount = 0 Then
            Insert Into 病人余额
              (病人id, 性质, 类型, 费用余额, 预交余额)
            Values
              (n_病人id, 1, 1, -1 * Nvl(c_费用.实收金额, 0), 0);
            n_返回额 := Nvl(c_费用.实收金额, 0);
          End If;
          If Nvl(n_返回额, 0) = 0 Then
            Delete 病人余额
            Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;
          --病人未结费用
          Update 病人未结费用
          Set 金额 = Nvl(金额, 0) - Nvl(c_费用.实收金额, 0)
          Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
                Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And
                收入项目id + 0 = c_费用.收入项目id And 来源途径 + 0 = 1;
        
          If Sql%RowCount = 0 Then
            Insert Into 病人未结费用
              (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
            Values
              (n_病人id, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, -1 * Nvl(c_费用.实收金额, 0));
          End If;
        End Loop;
        Delete 病人未结费用
        Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(金额, 0) = 0 And 来源途径 + 0 = 1;
      End If;
    
      If n_记帐 = 0 Then
        Begin
          Select 1
          Into n_预交支付
          From 病人预交记录
          Where Mod(记录性质, 10) = 1 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum < 2;
        Exception
          When Others Then
            n_预交支付 := 0;
        End;
        Begin
          Select 1
          Into n_正常支付
          From 病人预交记录
          Where Mod(记录性质, 10) = 4 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum < 2;
        Exception
          When Others Then
            n_正常支付 := 0;
        End;
        If n_预交支付 = 1 And n_正常支付 = 1 Then
          v_Error := '不能处理多种结算方式,请检查传入的退号单据是否正确!';
          Raise Err_Custom;
        End If;
        If n_预交支付 = 1 Then
          --原样退回预交
          If Nvl(预交id_In, 0) = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                     d_退号时间, v_操作员姓名, v_操作员编号, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
          Else
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 预交id_In, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_退号时间,
                     v_操作员姓名, v_操作员编号, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
          End If;
          --处理病人预交余额
          For c_预交 In (Select 病人id, 预交类别, -1 * Sum(Nvl(冲预交, 0)) As 冲预交
                       From 病人预交记录
                       Where 记录性质 In (1, 11) And 结帐id = n_结帐id
                       Group By 病人id, 预交类别) Loop
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + Nvl(c_预交.冲预交, 0)
            Where 病人id = c_预交.病人id And 类型 = Nvl(c_预交.预交类别, 2) And 性质 = 1
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额
                (病人id, 预交余额, 性质, 类型)
              Values
                (c_预交.病人id, Nvl(c_预交.冲预交, 0), 1, Nvl(c_预交.预交类别, 2));
              n_返回值 := Nvl(c_预交.冲预交, 0);
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = c_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
          End Loop;
        Else
          If Nvl(预交id_In, 0) = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易流水号, 交易说明,
               合作单位, 结算序号, 卡类别id, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_退号时间, v_操作员编号, v_操作员姓名, -冲预交,
                     n_结帐id, n_组id, 交易流水号_In, 交易说明_In, 合作单位, n_结帐id, 卡类别id, 结算性质
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
          Else
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易流水号, 交易说明,
               合作单位, 结算序号, 卡类别id, 结算性质)
              Select 预交id_In, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_退号时间, v_操作员编号, v_操作员姓名, -冲预交, n_结帐id,
                     n_组id, 交易流水号_In, 交易说明_In, 合作单位, n_结帐id, 卡类别id, 结算性质
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
          End If;
          Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
        End If;
        --退卡收回票据(可能上次挂号使用票据,不能收回)
        --从最后一次的打印内容中取
        Select Max(ID)
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 4 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
        If n_打印id Is Not Null Then
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_退号时间, v_操作员姓名, 票据金额
            From 票据使用明细
            Where 打印id = n_打印id And 性质 = 1;
        End If;
      End If;
    
      --相关汇总表的处理
    
      --病人挂号汇总
      Open c_Registinfo(1, 1);
      Fetch c_Registinfo
        Into r_Registrow;
    
      If c_Registinfo%RowCount = 0 Then
        --只收病历费时无号别,不处理
        Close c_Registinfo;
      Else
      
        --需要确定是否预约挂号
        --1.如果是退预约挂号产生的挂号记录,则需要减已挂数和其中已接数
        --2.如果是正常挂号,则只减已挂数
        Begin
          Select Decode(预约, Null, 0, 0, 0, 1), 执行状态
          Into n_预约挂号, n_执行状态
          From 病人挂号记录
          Where NO = 单据号_In And 记录状态 = 1 And Rownum = 1;
        Exception
          When Others Then
            n_预约挂号 := 0;
        End;
        --0-等待接诊,1-完成就诊,2-正在就诊,-1标记为不就诊
        If n_执行状态 > 0 Then
          If n_执行状态 = 1 Then
            v_Error := '该病人已经完成就诊,不能再退号!';
          Else
            v_Error := '该病人正在就诊, 不能退号!';
          End If;
          Raise Err_Custom;
        End If;
      
        Update 病人挂号汇总
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约状态, 已约数 = Nvl(已约数, 0) - n_预约状态
        Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
              Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
              (号码 = r_Registrow.号码 Or 号码 Is Null);
      
        If Sql%RowCount = 0 Then
          Insert Into 病人挂号汇总
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收)
          Values
            (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
             Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1, -1 * n_预约挂号);
        End If;
      
        Update 临床出诊记录
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约状态, 已约数 = Nvl(已约数, 0) - n_预约状态
        Where ID = n_记录id;
        Close c_Registinfo;
      End If;
    
      --人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款)
      If n_记帐 = 0 Then
        For r_Opermoney In c_Opermoney Loop
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + (-1 * r_Opermoney.冲预交)
          Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1
          Returning 余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (v_操作员姓名, r_Opermoney.结算方式, 1, -1 * r_Opermoney.冲预交);
            n_返回值 := r_Opermoney.冲预交;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
          End If;
        End Loop;
      End If;
    
      n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
      If n_挂号生成队列 <> 0 Then
        --要删除队列
        For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
          n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, Nvl(v_挂号.执行部门id, 0)));
          If Nvl(n_分诊台签到排队, 0) = 0 Then
            Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
          End If;
        End Loop;
      End If;
    
      --医保产生的就诊登记记录
      Delete From 就诊登记记录
      Where (病人id, 主页id, 就诊时间) In (Select 病人id, 主页id, 发生时间 From 病人挂号记录 Where NO = 单据号_In);
    End If;
  
    If Nvl(n_无效单据, 0) = 0 Then
      Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录状态 = 1;
      If Sql%NotFound Then
        v_Error := '未找到挂号单据,请检查!';
        Raise Err_Custom;
      End If;
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         复诊, 号序, 预约, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 出诊记录id)
        Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_退号时间, 发生时间,
               v_操作员编号, v_操作员姓名, 复诊, 号序, 预约, 接收人, 接收时间, 交易流水号_In, 交易说明_In, 合作单位, 出诊记录id
        From 病人挂号记录
        Where NO = 单据号_In And 记录状态 = 3;
    End If;
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 2, 单据号_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_挂号id, 单据号_In);
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Error || '+[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

Begin
  n_挂号排班模式 := To_Number(Substr(Nvl(zl_GetSysParameter('挂号排班模式'), 0), 1, 1));
  If n_挂号排班模式 = 1 Then
    --出诊表排班模式
    Zl_三方机构挂号_出诊_Delete(单据号_In, 交易流水号_In, 交易说明_In, 退号时间_In, 预交id_In);
  Else
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
    End If;
    --部门ID,部门名称;人员ID,人员编号,人员姓名
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Error := '当前操作人员未设置对应的人员关系,不能继续。';
      Raise Err_Custom;
    End If;
    v_操作员编号 := Zl_操作员(1, v_Temp);
    v_操作员姓名 := Zl_操作员(2, v_Temp);
  
    n_组id := Zl_Get组id(v_操作员姓名);
  
    d_退号时间 := 退号时间_In;
    If d_退号时间 Is Null Then
      d_退号时间 := Sysdate;
    End If;
  
    --首先判断要退号/取消预约的记录是否存在
    Begin
      Select Decode(记录性质, 2, 1, 0), 记录性质, 登记时间, 号别, 号序, Nvl(预约时间, 发生时间), 合作单位, Nvl(预约, 0), Decode(记录状态, 0, 1, 0)
      Into n_预约挂号, n_记录性质, d_登记时间, v_号别, n_号序, d_预约时间, v_合作单位, n_预约状态, n_无效单据
      From 病人挂号记录
      Where NO = 单据号_In And 记录状态 In (0, 1) And Rownum <= 1;
    Exception
      When Others Then
        n_预约挂号 := -1;
    End;
  
    If n_预约挂号 = -1 Then
      v_Error := '单据可能已经被退号或单据输入错误!';
      Raise Err_Custom;
    End If;
  
    Begin
      Select Nvl(记帐费用, 0), Decode(Sign(Nvl(结帐id, 0)), 0, 0, 1)
      Into n_记帐, n_已结帐
      From 门诊费用记录
      Where 记录性质 = 4 And NO = 单据号_In And 记录状态 In (1, 3) And Rownum < 2;
    Exception
      When Others Then
        n_记帐   := 0;
        n_已结帐 := 0;
    End;
  
    Begin
      Select a.Id Into n_安排id From 挂号安排 A Where a.号码 = v_号别;
    Exception
      When Others Then
        n_安排id := -1;
    End;
  
    Begin
      Select ID
      Into n_计划id
      From 挂号安排计划
      Where 安排id = n_安排id And 审核时间 Is Not Null And
            Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) =
            (Select Max(a.生效时间) As 生效
             From 挂号安排计划 A
             Where a.审核时间 Is Not Null And d_预约时间 Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                   a.失效时间 And a.安排id = n_安排id) And
            d_预约时间 Between Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And 失效时间;
    Exception
      When Others Then
        n_计划id := 0;
    End;
    If Nvl(n_计划id, 0) = 0 Then
      Select Decode(To_Char(d_预约时间, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排
      Where ID = n_安排id;
    Else
      Select Decode(To_Char(d_预约时间, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排计划
      Where ID = n_计划id;
    End If;
  
    If v_时间段 Is Not Null And d_启用时间 Is Not Null Then
      --检查是否跨模式挂号安排
      Select To_Date(To_Char(d_预约时间, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             To_Date(To_Char(d_预约时间, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
      Into d_检查开始时间, d_检查结束时间
      From 时间段
      Where 时间段 = v_时间段 And 站点 Is Null And 号类 Is Null;
      If d_检查开始时间 > d_检查结束时间 Then
        d_检查结束时间 := d_检查结束时间 + 1;
      End If;
      If d_检查结束时间 > d_启用时间 Then
        --获取出诊记录id
        Select Max(a.Id)
        Into n_出诊记录id
        From 临床出诊记录 A, 临床出诊号源 B
        Where a.号源id = b.Id And b.号码 = v_号别 And 上班时段 = v_时间段 And d_预约时间 Between 开始时间 And 终止时间;
      End If;
    End If;
  
    --预约检查是否添加合作单位控制
    --如果设置了合作单位控制 则
    Select Count(0) Into n_合作单位限制 From 合作单位安排控制 Where Rownum = 1;
    --更新挂号序号状态
    n_退号重用 := Zl_To_Number(zl_GetSysParameter('已退序号允许挂号', 1111));
    If n_退号重用 = 0 Then
      Update 挂号序号状态
      Set 状态 = 4
      Where 号码 = v_号别 And 序号 = n_号序 And 日期 Between Trunc(d_预约时间) And Trunc(d_预约时间 + 1) - 1 / 24 / 60 / 60;
    Else
      Delete 挂号序号状态
      Where 号码 = v_号别 And 序号 = n_号序 And 日期 Between Trunc(d_预约时间) And Trunc(d_预约时间 + 1) - 1 / 24 / 60 / 60;
    End If;
    If Nvl(n_预约挂号, 0) = 1 Or Nvl(n_无效单据, 0) = 1 Then
      If Nvl(n_无效单据, 0) = 0 Then
        --N天内不能取消预约号
        n_Count := Zl_To_Number(zl_GetSysParameter('N天内不能取消预约号', 1111));
        If n_Count <> 0 Then
          If Trunc(Sysdate - n_Count) < d_登记时间 Then
            v_Error := '不能退掉预约在' || To_Char(Trunc(Sysdate - n_Count), 'yyyy-mm-dd') || '以前的预约单!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      n_状态 := Case n_无效单据
                When 1 Then
                 0
                Else
                 1
              End;
      --减少已约数
      Open c_Registinfo(n_状态, 2, n_无效单据);
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已挂数 = Nvl(已挂数, 0) - Decode(n_预约状态, 0, 1, 0)
      Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
            Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
            (号码 = r_Registrow.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数, 已挂数)
        Values
          (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
           Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -n_预约状态, Decode(n_预约状态, 0, 1, 0));
      End If;
    
      If n_出诊记录id Is Not Null Then
        Update 临床出诊记录
        Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已挂数 = Nvl(已挂数, 0) - Decode(n_预约状态, 0, 1, 0)
        Where ID = n_出诊记录id And Nvl(已约数, 0) > 0;
        Update 临床出诊序号控制 Set 挂号状态 = Null, 操作员姓名 = Null Where 记录id = n_出诊记录id And 序号 = n_号序;
      End If;
      If Nvl(n_合作单位限制, 0) <> 0 And v_合作单位 Is Not Null And Nvl(n_预约状态, 0) <> 0 Then
        Update 合作单位挂号汇总
        Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已接数 = Nvl(已接数, 0) - Decode(n_预约状态, 0, 1, 0)
        Where 日期 = Trunc(r_Registrow.发生时间) And (号码 = r_Registrow.号码 Or 号码 Is Null) And 合作单位 = v_合作单位 And
              序号 = Nvl(n_号序, 0);
        If Sql%RowCount = 0 Then
          Insert Into 合作单位挂号汇总
            (日期, 号码, 已约数, 合作单位, 序号, 已接数)
          Values
            (Trunc(r_Registrow.发生时间), r_Registrow.号码, -n_预约状态, v_合作单位, Nvl(n_号序, 0), -decode(n_预约状态, 0, 1, 0));
        End If;
      End If;
      Close c_Registinfo;
    
      If Nvl(n_无效单据, 0) = 0 Then
        --删除门诊费用记录
        Delete From 门诊费用记录 Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 0;
        --如果预约生成队列时需要清除队列
        n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
        If Nvl(n_挂号生成队列, 0) = 1 Then
          n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
          If Nvl(n_预约生成队列, 0) = 1 Then
            --要删除队列
            For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
              Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
            End Loop;
          End If;
        End If;
      End If;
    Else
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    
      --更新挂号序号状态
    
      --病人就诊状态
      Select 病人id
      Into n_病人id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 序号 = 1;
    
      If n_病人id Is Not Null Then
        Update 病人信息 Set 就诊状态 = 0, 就诊诊室 = Null Where 病人id = n_病人id;
        --删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理
      End If;
    
      --门诊费用记录
      Insert Into 门诊费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位,
         付数, 数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间,
         登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
        Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
               收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
               v_操作员编号, v_操作员姓名, 发生时间, d_退号时间, n_结帐id,
               Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额, 摘要,
               Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
    
      --原始记录
      If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
        Update 门诊费用记录
        Set 记录状态 = 3, 结帐id = n_结帐id, 结帐金额 = 实收金额
        Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
      Else
        Update 门诊费用记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
      End If;
    
      n_原结帐id := 0;
      If n_记帐 = 0 Then
        --获取结帐ID
        Select Nvl(结帐id, 0)
        Into n_原结帐id
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Rownum < 2;
      End If;
    
      If n_记帐 = 1 Then
        --记帐
        For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                     From 门诊费用记录
                     Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Nvl(记帐费用, 0) = 1) Loop
          --病人余额
          Update 病人余额
          Set 费用余额 = Nvl(费用余额, 0) - Nvl(c_费用.实收金额, 0)
          Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1
          Returning 费用余额 Into n_返回额;
        
          If Sql%RowCount = 0 Then
            Insert Into 病人余额
              (病人id, 性质, 类型, 费用余额, 预交余额)
            Values
              (n_病人id, 1, 1, -1 * Nvl(c_费用.实收金额, 0), 0);
            n_返回额 := Nvl(c_费用.实收金额, 0);
          End If;
          If Nvl(n_返回额, 0) = 0 Then
            Delete 病人余额
            Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;
          --病人未结费用
          Update 病人未结费用
          Set 金额 = Nvl(金额, 0) - Nvl(c_费用.实收金额, 0)
          Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
                Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And
                收入项目id + 0 = c_费用.收入项目id And 来源途径 + 0 = 1;
        
          If Sql%RowCount = 0 Then
            Insert Into 病人未结费用
              (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
            Values
              (n_病人id, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, -1 * Nvl(c_费用.实收金额, 0));
          End If;
        End Loop;
        Delete 病人未结费用
        Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(金额, 0) = 0 And 来源途径 + 0 = 1;
      End If;
    
      If n_记帐 = 0 Then
        Begin
          Select 1
          Into n_预交支付
          From 病人预交记录
          Where Mod(记录性质, 10) = 1 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum < 2;
        Exception
          When Others Then
            n_预交支付 := 0;
        End;
      
        Begin
          Select 1
          Into n_正常支付
          From 病人预交记录
          Where Mod(记录性质, 10) = 4 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum < 2;
        Exception
          When Others Then
            n_正常支付 := 0;
        End;
      
        If n_预交支付 = 1 And n_正常支付 = 1 Then
          v_Error := '不能处理多种结算方式,请检查传入的退号单据是否正确!';
          Raise Err_Custom;
        End If;
      
        If n_预交支付 = 1 Then
          --原样退回预交
          If Nvl(预交id_In, 0) = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                     d_退号时间, v_操作员姓名, v_操作员编号, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
          Else
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 预交id_In, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_退号时间,
                     v_操作员姓名, v_操作员编号, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
          End If;
          --处理病人预交余额
          For c_预交 In (Select 病人id, 预交类别, -1 * Sum(Nvl(冲预交, 0)) As 冲预交
                       From 病人预交记录
                       Where 记录性质 In (1, 11) And 结帐id = n_结帐id
                       Group By 病人id, 预交类别) Loop
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + Nvl(c_预交.冲预交, 0)
            Where 病人id = c_预交.病人id And 类型 = Nvl(c_预交.预交类别, 2) And 性质 = 1
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额
                (病人id, 预交余额, 性质, 类型)
              Values
                (c_预交.病人id, Nvl(c_预交.冲预交, 0), 1, Nvl(c_预交.预交类别, 2));
              n_返回值 := Nvl(c_预交.冲预交, 0);
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = c_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
          End Loop;
        Else
          If Nvl(预交id_In, 0) = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易流水号, 交易说明,
               合作单位, 结算序号, 卡类别id, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_退号时间, v_操作员编号, v_操作员姓名, -冲预交,
                     n_结帐id, n_组id, 交易流水号_In, 交易说明_In, 合作单位, n_结帐id, 卡类别id, 结算性质
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
          Else
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易流水号, 交易说明,
               合作单位, 结算序号, 卡类别id, 结算性质)
              Select 预交id_In, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_退号时间, v_操作员编号, v_操作员姓名, -冲预交, n_结帐id,
                     n_组id, 交易流水号_In, 交易说明_In, 合作单位, n_结帐id, 卡类别id, 结算性质
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
          End If;
        
          Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
        End If;
        --退卡收回票据(可能上次挂号使用票据,不能收回)
        --从最后一次的打印内容中取
        Select Max(ID)
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 4 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
        If n_打印id Is Not Null Then
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_退号时间, v_操作员姓名, 票据金额
            From 票据使用明细
            Where 打印id = n_打印id And 性质 = 1;
        End If;
      End If;
    
      --相关汇总表的处理
    
      --病人挂号汇总
      Open c_Registinfo(1, 1);
      Fetch c_Registinfo
        Into r_Registrow;
    
      If c_Registinfo%RowCount = 0 Then
        --只收病历费时无号别,不处理
        Close c_Registinfo;
      Else
      
        --需要确定是否预约挂号
        --1.如果是退预约挂号产生的挂号记录,则需要减已挂数和其中已接数
        --2.如果是正常挂号,则只减已挂数
        Begin
          Select Decode(预约, Null, 0, 0, 0, 1), 执行状态
          Into n_预约挂号, n_执行状态
          From 病人挂号记录
          Where NO = 单据号_In And 记录状态 = 1 And Rownum = 1;
        Exception
          When Others Then
            n_预约挂号 := 0;
        End;
        --0-等待接诊,1-完成就诊,2-正在就诊,-1标记为不就诊
        If n_执行状态 > 0 Then
          If n_执行状态 = 1 Then
            v_Error := '该病人已经完成就诊,不能再退号!';
          Else
            v_Error := '该病人正在就诊, 不能退号!';
          End If;
          Raise Err_Custom;
        End If;
      
        Update 病人挂号汇总
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约状态, 已约数 = Nvl(已约数, 0) - n_预约状态
        Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
              Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
              (号码 = r_Registrow.号码 Or 号码 Is Null);
      
        If Sql%RowCount = 0 Then
          Insert Into 病人挂号汇总
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收)
          Values
            (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
             Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1, -1 * n_预约挂号);
        End If;
        If n_出诊记录id Is Not Null Then
          Update 临床出诊记录
          Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约状态, 已约数 = Nvl(已约数, 0) - n_预约状态
          Where ID = n_出诊记录id And Nvl(已约数, 0) > 0;
          Update 临床出诊序号控制 Set 挂号状态 = Null, 操作员姓名 = Null Where 记录id = n_出诊记录id And 序号 = n_号序;
        End If;
        If Nvl(n_合作单位限制, 0) <> 0 And v_合作单位 Is Not Null And Nvl(n_预约状态, 0) <> 0 Then
          Update 合作单位挂号汇总
          Set 已接数 = Nvl(已接数, 0) - 1, 已约数 = Nvl(已约数, 0) - n_预约挂号
          Where 日期 = Trunc(r_Registrow.发生时间) And (号码 = r_Registrow.号码 Or 号码 Is Null) And 合作单位 = v_合作单位 And
                序号 = Nvl(n_号序, 0);
          If Sql%RowCount = 0 Then
            Insert Into 合作单位挂号汇总
              (日期, 号码, 已约数, 合作单位, 已接数, 序号)
            Values
              (Trunc(r_Registrow.发生时间), r_Registrow.号码, -1, v_合作单位, -1 * n_预约挂号, Nvl(n_号序, 0));
          End If;
        End If;
        Close c_Registinfo;
      End If;
    
      --人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款)
      If n_记帐 = 0 Then
        For r_Opermoney In c_Opermoney Loop
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + (-1 * r_Opermoney.冲预交)
          Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1
          Returning 余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (v_操作员姓名, r_Opermoney.结算方式, 1, -1 * r_Opermoney.冲预交);
            n_返回值 := r_Opermoney.冲预交;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
          End If;
        End Loop;
      End If;
    
      n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
      If n_挂号生成队列 <> 0 Then
        --要删除队列
        For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
          n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(v_挂号.执行部门id, 0)));
          If Nvl(n_分诊台签到排队, 0) = 0 Then
            Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
          End If;
        End Loop;
      End If;
    
      --医保产生的就诊登记记录
      Delete From 就诊登记记录
      Where (病人id, 主页id, 就诊时间) In (Select 病人id, 主页id, 发生时间 From 病人挂号记录 Where NO = 单据号_In);
    End If;
  
    If Nvl(n_无效单据, 0) = 0 Then
      Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录状态 = 1;
      If Sql%NotFound Then
        v_Error := '未找到挂号单据,请检查!';
        Raise Err_Custom;
      End If;
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         复诊, 号序, 预约, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位)
        Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_退号时间, 发生时间,
               v_操作员编号, v_操作员姓名, 复诊, 号序, 预约, 接收人, 接收时间, 交易流水号_In, 交易说明_In, 合作单位
        From 病人挂号记录
        Where NO = 单据号_In And 记录状态 = 3;
    End If;
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 2, 单据号_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_挂号id, 单据号_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Error || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_三方机构挂号_Delete;
/

--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
Create Or Replace Procedure Zl_Third_Regist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --功能:HIS挂号 
  --入参:Xml_In: 
  --<IN> 
  --   <CZFS>3</CZFS>    //操作方式 
  --   <CZJLID>1</CZJLID>    //出诊记录ID 
  --   <HM>号码</HM>    //号码 
  --   <HX>号序</HX>     //号序 
  --   <JKFS>0</JKFS>  //缴款方式,0-挂号或预约缴款;1-预约不缴款 
  --   <YYSJ>2014-10-21 </YYSJ>    //预约日期 YYYY-MM-DD,分时段非序号控制需要传入时间 
  --   <JE>金额</JE>     //金额 
  --   <JSLIST> 
  --     <JS>            //结算信息，挂号非医保结算目前仅支持一个，结构与收费一致 
  --       <JSKLB>结算卡类别</JSKLB>    //结算卡类别 
  --       <JSKH>支付宝帐号</JSKH>           //结算卡号(支付宝帐号) 
  --       <JYSM>交易说明</JYSM>            //说明，固定传支付宝 
  --       <JYLSH>流水号</JYLSH>           //流水号，传订单号 
  --       <JSFS>结算方式</JSFS>            //结算方式:现金、支票，如果是三方卡,可以传空 
  --       <JSJE>结算金额</JSJE>            //结算金额 
  --       <ZY>摘要</ZY>                  //摘要 
  --       <SFCYJ></SFCYJ>              //是否冲预交，挂号目前不传 
  --       <SFXFK></SFXFK>              //是否消费卡,挂号目前不传 
  --       <EXPENDLIST>                 //扩展信息 
  --         <EXPEND> 
  --           <JYMC>交易名称</JYMC>        //交易名称 
  --           <JYLR>交易内容<JYLR>         //交易内容 
  --         </EXPEND> 
  --         <EXPEND> 
  --           ... 
  --         </EXPEND> 
  --       </EXPENDLIST> 
  --     </JS> 
  --   </JSLIST> 
  --   <HZDW>合作单位</HZDW>        //合作单位名称 
  --   <YYFS>支付宝<YYFS>    //预约方式,如自助机，支付宝 
  --   <BRID>病人ID</BRID>     //病人ID 
  --   <SFZH>身份证号</SFZH>     //身份证号 
  --   <XM>姓名</XM>            //姓名 
  --   <BRLX></BRLX>             //医保病人类型 
  --   <FB>普通</FB>               //病人费别，可以不传 
  --   <JQM>机器名</JQM>            //机器名 
  --</IN> 

  --出参:Xml_Out 
  --<OUTPUT> 
  -- <GHDH>挂号单号</GHDH>          //挂号单号 
  -- <CZSJ>操作时间</CZSJ>          //HIS的登记时间 
  -- <JZID>结帐ID</JZID>          //本次结帐ID 
  -- <ERROR><MSG>错误信息</MSG></ERROR>  //出错时返回 
  --</ OUTPUT> 
  -------------------------------------------------------------------------------------------------- 
  v_号码     挂号安排.号码%Type;
  d_发生时间 Date;
  d_原始时间 Date;
  d_登记时间 Date;

  n_应收金额   门诊费用记录.应收金额%Type;
  v_流水号     病人预交记录.交易流水号%Type;
  v_说明       门诊费用记录.摘要%Type;
  n_病人id     病人信息.病人id%Type;
  v_身份证号   病人信息.身份证号%Type;
  v_预约方式   预约方式.名称%Type;
  v_卡类别名称 医疗卡类别.名称%Type;
  v_结算卡号   病人预交记录.卡号%Type;
  n_门诊号     门诊费用记录.标识号%Type;
  v_姓名       门诊费用记录.姓名%Type;
  v_性别       门诊费用记录.性别%Type;
  v_年龄       门诊费用记录.年龄%Type;
  v_付款方式   门诊费用记录.付款方式%Type;
  v_费别       门诊费用记录.费别%Type;
  v_No         病人挂号记录.No%Type;
  v_结算方式   医疗卡类别.结算方式%Type;
  n_收费细目id 门诊费用记录.收费细目id%Type;
  n_病人科室id 门诊费用记录.病人科室id%Type;
  n_开单部门id 门诊费用记录.开单部门id%Type;
  v_操作员编号 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  v_医生姓名   挂号安排.医生姓名%Type;
  n_医生id     挂号安排.医生id%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  n_卡类别id   医疗卡类别.Id%Type;
  v_排班       挂号安排.周日%Type;
  n_安排id     挂号安排.Id%Type;
  n_计划id     挂号安排计划.Id%Type;
  n_预交id     病人预交记录.Id%Type;
  n_序号控制   挂号安排.序号控制%Type;
  n_号序       挂号序号状态.序号%Type;
  v_星期       挂号安排限制.限制项目%Type;
  v_病人类型   病人信息.病人类型%Type;
  n_存在       Number(3);
  v_现金       结算方式.名称%Type;
  n_分时段     Number(3);
  v_结算内容   Varchar2(3000);
  v_合作单位   病人挂号记录.合作单位%Type;
  v_机器名     挂号序号状态.机器名%Type;
  n_缴款方式   Number(3);
  n_挂号模式   Number(3);
  n_Exists     Number(3);
  v_保险结算   Varchar2(1000);
  n_记录id     临床出诊记录.Id%Type;
  v_Temp       Varchar2(32767); --临时XML 
  x_Templet    Xmltype; --模板XML 
  v_Err_Msg    Varchar2(200);
  d_启用时间   Date;
  n_Count      Number(3);
  v_卡类别     三方交易记录.类别%Type;
  n_冲预交     病人预交记录.冲预交%Type;
  v_Para       Varchar2(2000);
  Err_Item Exception;
  Err_Special Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/HM'), Extractvalue(Value(A), 'IN/HX'),
         To_Date(Extractvalue(Value(A), 'IN/YYSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/JE'),
         Extractvalue(Value(A), 'IN/YYFS'), Extractvalue(Value(A), 'IN/HZDW'), Extractvalue(Value(A), 'IN/BRID'),
         Extractvalue(Value(A), 'IN/BRLX'), Extractvalue(Value(A), 'IN/FB'), Extractvalue(Value(A), 'IN/JQM'),
         Extractvalue(Value(A), 'IN/JKFS'), Extractvalue(Value(A), 'IN/CZJLID'), Extractvalue(Value(A), 'IN/SFZH'),
         Extractvalue(Value(A), 'IN/XM')
  Into v_号码, n_号序, d_原始时间, n_应收金额, v_预约方式, v_合作单位, n_病人id, v_病人类型, v_费别, v_机器名, n_缴款方式, n_记录id, v_身份证号, v_姓名
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
  End If;
  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '无法确定病人信息,请检查!';
    Raise Err_Item;
  End If;

  v_Para     := zl_GetSysParameter(256);
  n_挂号模式 := Substr(v_Para, 1, 1);
  Begin
    d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_启用时间 := Null;
  End;

  If Sysdate - 10 > Nvl(d_启用时间, Sysdate - 30) Then
    If n_挂号模式 = 1 And Nvl(d_原始时间, Sysdate) > Nvl(d_启用时间, Sysdate - 30) And n_记录id Is Null Then
      v_Err_Msg := '系统当前处于出诊表排班模式，传入的参数无法确定挂号安排，请重试！';
      Raise Err_Item;
    End If;
  Else
    If n_挂号模式 = 1 And Nvl(d_原始时间, Sysdate) > Nvl(d_启用时间, Sysdate - 30) And n_记录id Is Null Then
      Begin
        Select a.Id
        Into n_记录id
        From 临床出诊记录 A, 临床出诊号源 B
        Where a.号源id = b.Id And b.号码 = v_号码 And Nvl(d_原始时间, Sysdate) Between a.开始时间 And a.终止时间;
      Exception
        When Others Then
          v_Err_Msg := '系统当前处于出诊表排班模式，传入的参数无法确定挂号安排，请重试！';
          Raise Err_Item;
      End;
    End If;
  End If;

  d_登记时间 := Sysdate;
  d_发生时间 := Trunc(d_原始时间);

  For c_交易记录 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    --冲预交不需要三方交易锁 
    If Nvl(c_交易记录.是否冲预交, 0) = 0 Then
      If c_交易记录.结算卡类别 Is Null Then
        v_卡类别 := c_交易记录.结算方式;
      Else
        Select Decode(Translate(Nvl(c_交易记录.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Count
        From Dual;
      
        If Nvl(n_Count, 0) = 1 Then
          Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(c_交易记录.结算卡类别);
        Else
          Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = c_交易记录.结算卡类别;
        End If;
      End If;
    
      If v_卡类别 Is Null Then
        v_Err_Msg := '不支持的结算方式,请检查！';
        Raise Err_Item;
      End If;
    
      If Zl_Fun_三方交易记录_Locked(v_卡类别, c_交易记录.交易流水号, c_交易记录.结算卡号, c_交易记录.摘要, 4) = 0 Then
        v_Err_Msg := '交易流水号为:' || c_交易记录.交易流水号 || '的交易正在进行中，不允许再次提交此交易!';
        Raise Err_Special;
      End If;
    End If;
  End Loop;

  If v_病人类型 Is Not Null Then
    Begin
      Select 1 Into n_存在 From 病人类型 Where 名称 = v_病人类型;
    Exception
      When Others Then
        v_Err_Msg := '没有发现为(' || v_病人类型 || ')的病人类型';
        Raise Err_Item;
    End;
    Update 病人信息 Set 病人类型 = Nvl(病人类型, v_病人类型) Where 病人id = n_病人id;
  End If;

  Select a.门诊号, a.姓名, a.性别, a.年龄, Nvl(b.编码, c.编码)
  Into n_门诊号, v_姓名, v_性别, v_年龄, v_付款方式
  From 病人信息 A, 医疗付款方式 B, (Select 编码 From 医疗付款方式 Where 缺省标志 = '1' And Rownum < 2) C
  Where a.病人id = n_病人id And a.医疗付款方式 = b.名称(+);

  v_Temp := Zl_Identity(1);
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_操作员编号 From Dual;
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_操作员姓名 From Dual;
  v_Temp := Zl_Identity(2);
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into n_开单部门id From Dual;

  v_No := Nextno(12);
  Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  Select 病人预交记录_Id.Nextval Into n_预交id From Dual;

  If n_记录id Is Null Then
    For r_结算 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      If Nvl(r_结算.是否冲预交, 0) = 0 Then
        If r_结算.结算方式 Is Null Then
          Begin
            Select b.结算方式, b.Id
            Into v_结算方式, n_卡类别id
            From 医疗卡类别 B
            Where b.名称 = r_结算.结算卡类别 And Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '没有发现该结算卡的相关信息';
              Raise Err_Item;
          End;
        Else
          Select Nvl(Max(1), 0) Into n_Exists From 结算方式 Where 名称 = r_结算.结算方式 And 性质 In (3, 4);
          If n_Exists = 1 Then
            v_保险结算 := v_保险结算 || '||' || r_结算.结算方式 || '|' || r_结算.结算金额;
          Else
            If v_结算方式 Is Null Then
              v_结算方式 := r_结算.结算方式;
            Else
              v_Err_Msg := '目前计划排班挂号不支持非医保外的多种结算方式,请检查!';
              Raise Err_Item;
            End If;
          End If;
        End If;
      
        If r_结算.结算卡类别 Is Not Null Then
          v_卡类别名称 := r_结算.结算卡类别;
          v_结算卡号   := r_结算.结算卡号;
          v_流水号     := r_结算.交易流水号;
          v_说明       := r_结算.交易说明;
        
          If n_卡类别id Is Null Then
            Begin
              Select b.Id Into n_卡类别id From 医疗卡类别 B Where b.名称 = r_结算.结算卡类别 And Rownum < 2;
            Exception
              When Others Then
                v_Err_Msg := '没有发现该结算卡的相关信息';
                Raise Err_Item;
            End;
          End If;
        
          Select Decode(Translate(Nvl(r_结算.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Count
          From Dual;
        
          If Nvl(n_Count, 0) = 1 Then
            Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(r_结算.结算卡类别);
          Else
            Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = r_结算.结算卡类别;
          End If;
        Else
          v_卡类别 := r_结算.结算方式;
        End If;
      
        Update 三方交易记录
        Set 业务结算id = n_结帐id
        Where 流水号 = r_结算.交易流水号 And 类别 = v_卡类别 And 业务类型 = 4;
      Else
        n_冲预交 := r_结算.结算金额;
      End If;
    End Loop;
  
    If v_保险结算 Is Not Null Then
      v_保险结算 := Substr(v_保险结算, 3);
    End If;
  
    Select Decode(To_Char(d_原始时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六',
                   Null)
    Into v_星期
    From Dual;
  
    Begin
      Select ID
      Into n_计划id
      From (Select ID
             From 挂号安排计划
             Where 号码 = v_号码 And d_原始时间 Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                   失效时间 And 审核时间 Is Not Null
             Order By 生效时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Select ID Into n_安排id From 挂号安排 Where 号码 = v_号码;
    End;
  
    If Nvl(n_计划id, 0) <> 0 Then
      --从计划读取信息 
      Select a.项目id, b.科室id, a.医生姓名, a.医生id,
             Decode(To_Char(d_发生时间, 'D'), '1', a.周日, '2', a.周一, '3', a.周二, '4', a.周三, '5', a.周四, '6', a.周五, '7', a.周六,
                     Null), Nvl(a.序号控制, 0)
      Into n_收费细目id, n_病人科室id, v_医生姓名, n_医生id, v_排班, n_序号控制
      From 挂号安排计划 A, 挂号安排 B
      Where a.Id = n_计划id And b.Id = a.安排id;
      Select Count(Rownum) Into n_分时段 From 挂号计划时段 Where 星期 = v_星期 And 计划id = n_计划id And Rownum < 2;
    
      --合作单位检查 
      If v_合作单位 Is Not Null Then
        Begin
          Select 1 Into n_存在 From 合作单位计划控制 Where 计划id = n_计划id And 数量 = 0 And 合作单位 = v_合作单位;
        Exception
          When Others Then
            n_存在 := 0;
        End;
      End If;
    
      If n_存在 = 1 Then
        v_Err_Msg := '传入的合作单位在此号码上被禁用！';
        Raise Err_Item;
      End If;
    
      If n_分时段 = 1 And n_序号控制 = 0 Then
        d_发生时间 := d_原始时间;
        Select 序号
        Into n_号序
        From 挂号计划时段
        Where 计划id = n_计划id And 星期 = v_星期 And To_Char(开始时间, 'hh24:mi:ss') = To_Char(d_发生时间, 'hh24:mi:ss');
      Else
        Begin
          Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
          Into d_发生时间
          From 挂号计划时段
          Where 计划id = n_计划id And 星期 = v_星期 And 序号 = Nvl(n_号序, 0);
        Exception
          When Others Then
            If n_分时段 = 1 And n_序号控制 = 1 Then
              Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(Max(结束时间), 'hh24:mi:ss'),
                              'YYYY-MM-DD hh24:mi:ss')
              Into d_发生时间
              From 挂号计划时段
              Where 计划id = n_计划id And 星期 = v_星期;
            Else
              Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
              Into d_发生时间
              From 时间段
              Where 时间段 = v_排班;
            End If;
            If d_发生时间 < d_登记时间 Then
              d_发生时间 := d_登记时间;
            End If;
        End;
      End If;
    Else
      --从安排读取信息 
      Select b.项目id, b.科室id, b.医生姓名, b.医生id,
             Decode(To_Char(d_发生时间, 'D'), '1', b.周日, '2', b.周一, '3', b.周二, '4', b.周三, '5', b.周四, '6', b.周五, '7', b.周六,
                     Null), Nvl(b.序号控制, 0)
      Into n_收费细目id, n_病人科室id, v_医生姓名, n_医生id, v_排班, n_序号控制
      From 挂号安排 B
      Where b.Id = n_安排id;
      Select Count(Rownum) Into n_分时段 From 挂号安排时段 Where 星期 = v_星期 And 安排id = n_安排id And Rownum < 2;
    
      --合作单位检查 
      If v_合作单位 Is Not Null Then
        Begin
          Select 1 Into n_存在 From 合作单位安排控制 Where 安排id = n_安排id And 数量 = 0 And 合作单位 = v_合作单位;
        Exception
          When Others Then
            n_存在 := 0;
        End;
      End If;
    
      If n_存在 = 1 Then
        v_Err_Msg := '传入的合作单位在此号码上被禁用！';
        Raise Err_Item;
      End If;
    
      If n_分时段 = 1 And n_序号控制 = 0 Then
        d_发生时间 := d_原始时间;
        Select 序号
        Into n_号序
        From 挂号安排时段
        Where 安排id = n_安排id And 星期 = v_星期 And To_Char(开始时间, 'hh24:mi:ss') = To_Char(d_发生时间, 'hh24:mi:ss');
      Else
        Begin
          Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
          Into d_发生时间
          From 挂号安排时段
          Where 安排id = n_安排id And 星期 = v_星期 And 序号 = Nvl(n_号序, 0);
        Exception
          When Others Then
            If n_分时段 = 1 And n_序号控制 = 1 Then
              Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(Max(结束时间), 'hh24:mi:ss'),
                              'YYYY-MM-DD hh24:mi:ss')
              Into d_发生时间
              From 挂号安排时段
              Where 安排id = n_安排id And 星期 = v_星期;
            Else
              Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
              Into d_发生时间
              From 时间段
              Where 时间段 = v_排班;
            End If;
            If d_发生时间 < d_登记时间 Then
              d_发生时间 := d_登记时间;
            End If;
        End;
      End If;
    End If;
  
    If Trunc(d_发生时间) <> Trunc(Sysdate) Then
      If Nvl(n_缴款方式, 0) = 0 Then
        Zl_三方机构挂号_Insert(3, n_病人id, v_号码, n_号序, v_No, Null, v_结算方式, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1);
      Else
        Zl_三方机构挂号_Insert(2, n_病人id, v_号码, n_号序, v_No, Null, v_结算方式, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1);
      End If;
    Else
      If Nvl(n_缴款方式, 0) = 0 Then
        Zl_三方机构挂号_Insert(1, n_病人id, v_号码, n_号序, v_No, Null, v_结算方式, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1);
      Else
        Zl_三方机构挂号_Insert(2, n_病人id, v_号码, n_号序, v_No, Null, v_结算方式, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1);
      End If;
    End If;
  
    For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                          Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
      Zl_三方结算交易_Insert(n_卡类别id, 0, v_结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
    End Loop;
  
    v_Temp := '<GHDH>' || v_No || '</GHDH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_结帐id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Else
    --出诊表排班模式 
    For r_结算 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      If Nvl(r_结算.是否冲预交, 0) = 0 Then
        If r_结算.结算方式 Is Null Then
          Begin
            Select b.结算方式, b.Id
            Into v_结算方式, n_卡类别id
            From 医疗卡类别 B
            Where b.名称 = r_结算.结算卡类别 And Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '没有发现该结算卡的相关信息';
              Raise Err_Item;
          End;
          v_结算内容 := v_结算内容 || '|' || v_结算方式 || ',' || r_结算.结算金额 || ',,';
        Else
          v_结算内容 := v_结算内容 || '|' || r_结算.结算方式 || ',' || r_结算.结算金额 || ',,';
        End If;
      
        If r_结算.结算卡类别 Is Not Null Then
          v_结算内容   := v_结算内容 || '1';
          v_卡类别名称 := r_结算.结算卡类别;
          v_结算卡号   := r_结算.结算卡号;
          v_流水号     := r_结算.交易流水号;
          v_说明       := r_结算.交易说明;
          If n_卡类别id Is Null Then
            Begin
              Select b.Id Into n_卡类别id From 医疗卡类别 B Where b.名称 = r_结算.结算卡类别 And Rownum < 2;
            Exception
              When Others Then
                v_Err_Msg := '没有发现该结算卡的相关信息';
                Raise Err_Item;
            End;
          End If;
        
          Select Decode(Translate(Nvl(r_结算.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Count
          From Dual;
        
          If Nvl(n_Count, 0) = 1 Then
            Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(r_结算.结算卡类别);
          Else
            Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = r_结算.结算卡类别;
          End If;
        Else
          v_结算内容 := v_结算内容 || '0';
          v_卡类别   := r_结算.结算方式;
        End If;
      
        Update 三方交易记录
        Set 业务结算id = n_结帐id
        Where 流水号 = r_结算.交易流水号 And 类别 = v_卡类别 And 业务类型 = 4;
      Else
        n_冲预交 := r_结算.结算金额;
      End If;
    End Loop;
  
    If v_结算内容 Is Not Null Then
      v_结算内容 := Substr(v_结算内容, 2);
    Else
      Begin
        Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
      Exception
        When Others Then
          v_现金 := '现金';
      End;
      v_结算内容 := v_现金 || ',' || 0 || ',,0';
    End If;
  
    Select 项目id, 科室id, 医生姓名, 医生id, 是否序号控制, 是否分时段
    Into n_收费细目id, n_病人科室id, v_医生姓名, n_医生id, n_序号控制, n_分时段
    From 临床出诊记录
    Where ID = n_记录id;
  
    Begin
      Select 开始时间 Into d_发生时间 From 临床出诊序号控制 Where 记录id = n_记录id And 序号 = n_号序;
    Exception
      When Others Then
        d_发生时间 := d_原始时间;
    End;
  
    If Trunc(d_发生时间) <> Trunc(Sysdate) Then
      If Nvl(n_缴款方式, 0) = 0 Then
        Zl_三方机构挂号_Insert(3, n_病人id, v_号码, n_号序, v_No, Null, v_结算内容, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, v_保险结算, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1, 0, n_记录id);
      Else
        Zl_三方机构挂号_Insert(2, n_病人id, v_号码, n_号序, v_No, Null, v_结算内容, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, v_保险结算, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1, 0, n_记录id);
      End If;
    Else
      If Nvl(n_缴款方式, 0) = 0 Then
        Zl_三方机构挂号_Insert(1, n_病人id, v_号码, n_号序, v_No, Null, v_结算内容, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, v_保险结算, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1, 0, n_记录id);
      Else
        Zl_三方机构挂号_Insert(2, n_病人id, v_号码, n_号序, v_No, Null, v_结算内容, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, v_保险结算, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1, 0, n_记录id);
      End If;
    End If;
  
    For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                          Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
      Zl_三方结算交易_Insert(n_卡类别id, 0, v_结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
    End Loop;
  
    v_Temp := '<GHDH>' || v_No || '</GHDH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_结帐id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Regist;
/

--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
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
    Select 病人id, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id, Min(Decode(记录性质, 1, 收款时间, Null)) As 收款时间
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1
     Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By NO, 病人id
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
    Where 号码 = 号别_In And 日期 = 发生时间_In And 序号 = 号序_In And 状态 = 3 And 操作员姓名 = 操作员姓名_In;
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
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 Not In (4, 5);
        Else
          Select Max(Nvl(序号, 0)), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Nvl(预约, 0))
          Into n_已用序号, n_已用数量, n_已约数
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 <> 5;
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
                状态 Not In (4, 5);
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
        Where a.日期 = d_时段时间 And 号码 = 号别_In And 状态 Not In (4, 5);
      Else
        Select Nvl(Max(序号), 0)
        Into n_挂出的最大序号
        From 挂号序号状态 A
        Where a.日期 = d_时段时间 And 号码 = 号别_In And 状态 <> 5;
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
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 Not In (4, 5);
        Else
          Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(日期 - d_时段时间), 0, 1, 0))
          Into n_已用序号, n_已挂数, n_已用数量
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 <> 5;
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

--133584:李南春,2018-11-12,利用失效时间查询挂号安排计划
Create Or Replace Procedure Zl_病人挂号记录_Delete
(
  单据号_In       门诊费用记录.No%Type,
  操作员编号_In   门诊费用记录.操作员编号%Type,
  操作员姓名_In   门诊费用记录.操作员姓名%Type,
  摘要_In         门诊费用记录.摘要%Type := Null, --预约取消时 填写 存放预约取消原因
  删除门诊号_In   Number := 0,
  非原样退结算_In Varchar2 := Null,
  退费类型_In     In Number := 0, --0-全退 1-退挂号费 2-退病历费 3-退附加费 4-退挂号与病历 5-退挂号与附加
  退指定结算_In   病人预交记录.结算方式%Type := Null,
  退号重用_In     Number := 1,
  收回票据号_In   Varchar2 := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null
) As
  --退费类型_In,在一下几种情况下不准进行部分退费
  --    2.三方接口,暂时不支持
  -- 挂号费病历费分开退,规则
  --    普通结算方式:原结算方式退部分费用
  --    预交款:预交款,退部分
  --    预交款与普通结算方式混合:退款按照普通结算方式部分退
  --    消费卡:原样将费用部分退入消费卡
  --非原样退结算_In:指不能退还给原样结算方式(如医保的个人账户,三方账户的退现等),多个用逗分离
  --退指定结算_IN:指非原样退结算部分,应该退给哪种结算方式,为空时缺省退给现金,否则退给指定的结算方式

  --该游标用于判断是否单独收病历费,及挂号汇总表处理
  Cursor c_Registinfo(v_状态 门诊费用记录.记录状态%Type) Is
    Select a.发生时间, a.登记时间, c.接收时间, a.收费细目id As 项目id, c.执行部门id As 科室id, c.执行人 As 医生姓名, d.Id As 医生id, c.号别 As 号码
    From 门诊费用记录 A, 挂号安排 B, 病人挂号记录 C, 人员表 D
    Where a.记录性质 = 4 And a.序号 = 1 And a.记录状态 = v_状态 And c.No = a.No And c.执行人 = d.姓名(+) And a.No = 单据号_In And
          Nvl(a.计算单位, '号别') = c.号别 And Rownum < 2;
  r_Registrow c_Registinfo%RowType;

  --该游标用于判断记录是否存在,及费用汇总表处理
  Cursor c_Moneyinfo(v_状态 门诊费用记录.记录状态%Type) Is
    Select 病人科室id, 开单部门id, 执行部门id, 收入项目id, Nvl(Sum(应收金额), 0) As 应收, Nvl(Sum(实收金额), 0) As 实收, Nvl(Sum(结帐金额), 0) As 结帐
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = v_状态 And NO = 单据号_In
    Group By 病人科室id, 开单部门id, 执行部门id, 收入项目id;
  r_Moneyrow c_Moneyinfo%RowType;

  --该光标用于处理人员缴款余额中退的不同结算方式的金额
  Cursor c_Opermoney(n_Id 病人预交记录.结帐id%Type) Is
    Select Distinct b.结算方式, -1 * Nvl(b.冲预交, 0) As 冲预交
    From 病人预交记录 B
    Where b.结帐id = n_Id And b.记录性质 = 4 And b.记录状态 = 2 And Nvl(b.冲预交, 0) <> 0;

  v_Err_Msg Varchar(255);
  Err_Item Exception;

  n_结帐id 病人预交记录.结帐id%Type;
  n_销帐id 门诊费用记录.结帐id%Type;

  v_退指定结算方式 病人预交记录.结算方式%Type;
  n_退款金额       病人预交记录.冲预交%Type;
  n_打印id         票据打印内容.Id%Type;
  n_病人id         病人信息.病人id%Type;
  n_退费金额       病人预交记录.冲预交%Type;
  n_预交金额       病人预交记录.冲预交%Type; --原记录 预交缴款金额
  n_返回值         病人余额.预交余额%Type;
  n_挂号id         病人挂号记录.Id%Type;
  n_组id           财务缴款分组.Id%Type;

  n_二次退费       Number; --记录是否是此单据的第二次退费
  n_分诊台签到排队 Number;
  n_预约生成队列   Number;
  n_预约挂号       Number;
  n_挂号生成队列   Number;
  d_Date           Date;
  n_记帐           门诊费用记录.记帐费用%Type;
  n_病人id1        病人信息.病人id%Type;
  n_返回额         门诊费用记录.实收金额%Type;
  n_已结帐         Number;
  n_安排id         挂号安排.Id%Type;
  n_计划id         挂号安排计划.Id%Type;
  d_启用时间       Date;
  d_发生时间       病人挂号记录.发生时间%Type;
  n_出诊记录id     临床出诊记录.Id%Type;
  v_号码           挂号安排.号码%Type;
  n_序号           病人挂号记录.号序%Type;
  v_时间段         时间段.时间段%Type;
  d_检查开始时间   Date;
  d_检查结束时间   Date;
  v_Temp           Varchar2(500);
  v_附加ids        Varchar2(500);
  n_就诊病人id     病人信息.病人id%Type;
  d_就诊时间       就诊登记记录.就诊时间%Type;
Begin
  n_组id           := Zl_Get组id(操作员姓名_In);
  v_退指定结算方式 := 退指定结算_In;

  --首先判断要退号/取消预约的记录是否存在
  Open c_Moneyinfo(1);
  Fetch c_Moneyinfo
    Into r_Moneyrow;
  If c_Moneyinfo%RowCount = 0 Then
    Close c_Moneyinfo;
    Open c_Moneyinfo(0);
    Fetch c_Moneyinfo
      Into r_Moneyrow;
    If c_Moneyinfo%RowCount = 0 Then
      v_Err_Msg := '要处理的单据不存在。';
      Raise Err_Item;
    End If;
    n_预约挂号 := 1;
  End If;
  Close c_Moneyinfo;

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
  End If;

  Select 号别, 号序, 发生时间 Into v_号码, n_序号, d_发生时间 From 病人挂号记录 Where NO = 单据号_In And Rownum < 2;

  Begin
    Select a.Id Into n_安排id From 挂号安排 A Where a.号码 = v_号码;
  Exception
    When Others Then
      n_安排id := -1;
  End;

  Begin
    Select ID
    Into n_计划id
    From 挂号安排计划
    Where 安排id = n_安排id And 审核时间 Is Not Null And
          Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) =
          (Select Max(a.生效时间) As 生效
           From 挂号安排计划 A
           Where a.审核时间 Is Not Null And d_发生时间 Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                 a.失效时间 And a.安排id = n_安排id) And
          d_发生时间 Between Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And 失效时间;
  Exception
    When Others Then
      n_计划id := 0;
  End;

  Begin
    If Nvl(n_计划id, 0) = 0 Then
      Select Decode(To_Char(d_发生时间, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排
      Where ID = n_安排id;
    Else
      Select Decode(To_Char(d_发生时间, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排计划
      Where ID = n_计划id;
    End If;
  Exception
    When Others Then
      v_时间段 := Null;
  End;

  If v_时间段 Is Not Null And d_启用时间 Is Not Null Then
    --检查是否跨模式挂号安排
    Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
           To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
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
        Where a.号源id = b.Id And b.号码 = v_号码 And 上班时段 = v_时间段 And d_发生时间 Between 开始时间 And 终止时间;
      Exception
        When Others Then
          n_出诊记录id := Null;
      End;
    End If;
  End If;

  Begin
    Select Zl_Fun_Regcustomname Into v_Temp From Dual;
    If v_Temp Is Not Null Then
      v_附加ids := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    End If;
  Exception
    When Others Then
      v_附加ids := Null;
  End;

  --1.预约处理
  If Nvl(n_预约挂号, 0) = 1 Then
    --减少已约数
    Open c_Registinfo(0);
    Fetch c_Registinfo
      Into r_Registrow;
  
    Update 病人挂号汇总
    Set 已约数 = Nvl(已约数, 0) - 1
    Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
          (号码 = r_Registrow.号码 Or 号码 Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into 病人挂号汇总
        (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数)
      Values
        (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
         Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1);
    End If;
    Close c_Registinfo;
  
    --更新挂号序号状态
    Delete 挂号序号状态
    Where 状态 = 2 And
          (号码, 序号, 日期) = (Select 计算单位, 发药窗口, Trunc(发生时间)
                          From 门诊费用记录
                          Where 记录性质 = 4 And 记录状态 = 0 And 序号 = 1 And Rownum = 1 And NO = 单据号_In) Or
          (号码, 序号, 日期) = (Select 计算单位, 发药窗口, 发生时间
                          From 门诊费用记录
                          Where 记录性质 = 4 And 记录状态 = 0 And 序号 = 1 And Rownum = 1 And NO = 单据号_In);
  
    --添加病人挂号记录的 冲销记录
    Select 病人挂号记录_Id.Nextval, Sysdate Into n_挂号id, d_Date From Dual;
  
    Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录状态 = 1 And 记录性质 = 2;
    If Sql%NotFound Then
      v_Err_Msg := '预约单【' || 单据号_In || '】不存在或由于并发原因已经被取消预约';
      Raise Err_Item;
    End If;
  
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
       复诊, 号序, 社区, 预约, 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式)
      Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_Date, 发生时间,
             操作员编号_In, 操作员姓名_In, 复诊, 号序, 社区, 预约, Nvl(摘要_In, 摘要) As 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 医疗付款方式
      From 病人挂号记录
      Where NO = 单据号_In;
  
    If n_出诊记录id Is Not Null Then
      Update 临床出诊记录 Set 已约数 = 已约数 - 1 Where ID = n_出诊记录id And Nvl(已约数, 0) > 0;
      Update 临床出诊序号控制 Set 挂号状态 = Null, 操作员姓名 = Null Where 记录id = n_出诊记录id And 序号 = n_序号;
    End If;
  
    --Update 病人挂号记录 set 摘要=nvl(摘要_IN,摘要) where NO=单据号_IN;
    --删除门诊费用记录
    Delete From 门诊费用记录 Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 0;
    --如果预约生成队列时需要清除队列
  
    n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    If Nvl(n_预约生成队列, 0) = 1 Then
      --要删除队列
      For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
        Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
      End Loop;
    End If;
    Return;
  End If;

  Select Nvl(记帐费用, 0), 病人id, Decode(Sign(Nvl(结帐id, 0)), 0, 0, 1)
  Into n_记帐, n_病人id, n_已结帐
  From 门诊费用记录
  Where 记录性质 = 4 And NO = 单据号_In And 记录状态 In (1, 3) And Rownum < 2;

  --2.挂号处理
  n_已结帐 := Nvl(n_已结帐, 0);

  If n_已结帐 = 1 And n_记帐 = 1 Then
    Select Sysdate, Null Into d_Date, n_销帐id From Dual;
  Else
    Select Sysdate, 病人结帐记录_Id.Nextval Into d_Date, n_销帐id From Dual;
  End If;

  ----0-全退 1-退挂号费 2-退病历费
  If Nvl(退费类型_In, 0) Not In (2, 3) Then
    --不是光退病历费时处理
    --更新挂号序号状态
    If 退号重用_In = 1 Then
      Delete 挂号序号状态
      Where 状态 = 1 And
            (号码, 序号, 日期) = (Select 号别, 号序, Trunc(发生时间) From 病人挂号记录 Where NO = 单据号_In And Rownum = 1) Or
            (号码, 序号, 日期) = (Select 号别, 号序, 发生时间 From 病人挂号记录 Where NO = 单据号_In And Rownum = 1);
    Else
      Update 挂号序号状态
      Set 状态 = 4
      Where 状态 = 1 And
            (号码, 序号, 日期) = (Select 号别, 号序, Trunc(发生时间) From 病人挂号记录 Where NO = 单据号_In And Rownum = 1) Or
            (号码, 序号, 日期) = (Select 号别, 号序, 发生时间 From 病人挂号记录 Where NO = 单据号_In And Rownum = 1);
    End If;
  
    --病人就诊状态
    If n_病人id Is Not Null Then
      Update 病人信息 Set 就诊状态 = 0, 就诊诊室 = Null Where 病人id = n_病人id;
    
      --删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理
      If 删除门诊号_In = 1 Then
        Delete 门诊病案记录 Where 病人id = n_病人id;
        Update 病人信息 Set 门诊号 = Null Where 病人id = n_病人id;
        --费用记录包括挂号及病案、就诊卡费用,以及病人交费后退费或销帐的费用,挂号记录在最后处理
        Update 门诊费用记录 Set 标识号 = Null Where 门诊标志 = 1 And 病人id = n_病人id;
      End If;
    End If;
  
    --如果挂时收了就诊卡费,退费时清除就诊卡号,在非光退病历费时
    n_病人id1 := Null;
    Begin
      Select 病人id
      Into n_病人id1
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 附加标志 = 2 And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If n_病人id1 Is Not Null And Nvl(退费类型_In, 0) Not In (2, 3) Then
      Update 病人信息
      Set 就诊卡号 = Null, 卡验证码 = Null, Ic卡号 = Decode(Ic卡号, 就诊卡号, Null, Ic卡号)
      Where 病人id = n_病人id1;
    End If;
  
  End If;

  --检查前面是否已经部分退过费用
  Begin
    Select 1 Into n_二次退费 From 门诊费用记录 Where 记录性质 = 4 And NO = 单据号_In And 记录状态 = 3 And Rownum < 2;
  Exception
    When Others Then
      n_二次退费 := 0;
  End;

  If Nvl(退费类型_In, 0) = 0 Or Nvl(退费类型_In, 0) = 2 Then
    --全退,退病历费
    --门诊费用记录，冲销记录
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
       数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
       结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
             收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
             操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_销帐id,
             Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额,
             Nvl(摘要_In, 摘要) As 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  
    --原始记录
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    End If;
  Elsif Nvl(退费类型_In, 0) = 1 Or Nvl(退费类型_In, 0) = 4 Then
    --退挂号费,退挂号与病历费
    --门诊费用记录，冲销记录
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
       数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
       结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
             收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
             操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_销帐id,
             Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额,
             Nvl(摘要_In, 摘要) As 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') = 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  
    --原始记录
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') = 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') = 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    End If;
  Elsif Nvl(退费类型_In, 0) = 3 Then
    --退附加费
    --门诊费用记录，冲销记录
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
       数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
       结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
             收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
             操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_销帐id,
             Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额,
             Nvl(摘要_In, 摘要) As 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') > 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
  
    --原始记录
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') > 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Instr(',' || v_附加ids || ',', ',' || 收费细目id || ',') > 0 And
            Nvl(附加标志, 0) = Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0));
    End If;
  Elsif Nvl(退费类型_In, 0) = 5 Then
    --退挂号与附加费
    --门诊费用记录，冲销记录
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位, 付数,
       数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间, 登记时间,
       结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
             收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
             操作员编号_In, 操作员姓名_In, 发生时间, d_Date, n_销帐id,
             Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额,
             Nvl(摘要_In, 摘要) As 摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Nvl(附加标志, 0) <> 1;
  
    --原始记录
    If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
      Update 门诊费用记录
      Set 记录状态 = 3, 结帐id = n_销帐id, 结帐金额 = 实收金额
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Nvl(附加标志, 0) <> 1;
    Else
      Update 门诊费用记录
      Set 记录状态 = 3
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And Nvl(附加标志, 0) <> 1;
    End If;
  End If;

  n_结帐id := 0;
  If n_记帐 = 0 Then
    --获取结帐ID
    Select Nvl(结帐id, 0)
    Into n_结帐id
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Rownum < 2;
  End If;

  If n_记帐 = 1 Then
    --记帐
    For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                 From 门诊费用记录
                 Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And
                       Nvl(附加标志, 0) =
                       Decode(Nvl(退费类型_In, 0), 2, 1, 1, Decode(Nvl(附加标志, 0), 1, -1, Nvl(附加标志, 0)), Nvl(附加标志, 0)) And
                       Nvl(记帐费用, 0) = 1) Loop
      --病人余额
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) - Nvl(c_费用.实收金额, 0)
      Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1
      Returning 费用余额 Into n_返回额;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 费用余额, 预交余额)
        Values
          (n_病人id, 1, 1, -1 * Nvl(c_费用.实收金额, 0), 0);
        n_返回额 := Nvl(c_费用.实收金额, 0);
      End If;
      If Nvl(n_返回额, 0) = 0 Then
        Delete 病人余额
        Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) - Nvl(c_费用.实收金额, 0)
      Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And 收入项目id + 0 = c_费用.收入项目id And
            来源途径 + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (n_病人id, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, -1 * Nvl(c_费用.实收金额, 0));
      End If;
    End Loop;
    Delete 病人未结费用
    Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(金额, 0) = 0 And 来源途径 + 0 = 1;
  End If;

  If n_记帐 = 0 Then
    --1.退费
    --病人挂号结算:现金和个人帐户部份
    If 非原样退结算_In Is Not Null Then
      --退款金额获取
      If Nvl(退费类型_In, 0) <> 0 Or Nvl(n_二次退费, 0) <> 0 Then
        --如果是单独退病历费,或者只退挂号费,先获取退费金额
        Begin
          --获取本次退款金额
          Select -1 * Sum(Nvl(实收金额, 0)) As 收款金额
          Into n_退款金额
          From 门诊费用记录
          Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id;
        
        Exception
          When Others Then
            v_Err_Msg := '单据【' || 单据号_In || '】的' || Case Nvl(退费类型_In, 0)
                           When 1 Then
                            '挂号费用'
                           When 2 Then
                            '病历费'
                         End || '可能由于并发原因已经进行了退费或者单据不存在!';
            Raise Err_Item;
        End;
        Begin
          Select 冲预交
          Into n_退费金额
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
        Exception
          When Others Then
            n_退费金额 := 0;
        End;
      
        --a.允许的结算方式
      
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
           结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In, -n_退款金额,
                 n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
      
        If n_退费金额 = 0 Then
          --b.不允许的退现金
          If n_退款金额 <> 0 Then
            If v_退指定结算方式 Is Null Then
              --退给现金
              Begin
                Select 名称 Into v_退指定结算方式 From 结算方式 Where 性质 = 1;
              Exception
                When Others Then
                  v_退指定结算方式 := '现金';
              End;
            End If;
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * n_退款金额), 交易说明 = Nvl(交易说明_In, 交易说明)
            Where 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id And 结算方式 = v_退指定结算方式;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
                 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
                Select 病人预交记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.病人id, a.主页id, a.科室id, a.摘要, v_退指定结算方式, d_Date,
                       操作员编号_In, 操作员姓名_In, -1 * n_退款金额, n_销帐id, n_组id, 预交类别, Decode(交易说明_In, Null, 卡类别id, Null), 结算卡序号,
                       Decode(交易说明_In, Null, 卡号, Null), Decode(交易说明_In, Null, 交易流水号, Null), Nvl(交易说明_In, 交易说明), 合作单位, 4
                From 病人预交记录 A
                Where a.记录性质 = 4 And a.记录状态 = 1 And a.结帐id = n_结帐id And Rownum < 2;
            End If;
          End If;
        End If;
      Else
        --a.允许的结算方式原样退
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
           结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In, -冲预交,
                 n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
      
        --b.不允许的退现金
        Begin
          Select Sum(冲预交)
          Into n_退费金额
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') > 0;
        Exception
          When Others Then
            n_退费金额 := 0;
        End;
        If n_退费金额 <> 0 Then
          If v_退指定结算方式 Is Null Then
            --退给现金
            Begin
              Select 结算方式
              Into v_退指定结算方式
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id And Instr(',' || 非原样退结算_In || ',', ',' || 结算方式 || ',') = 0;
            
            Exception
              When Others Then
                Begin
                  Select 名称 Into v_退指定结算方式 From 结算方式 Where 性质 = 1;
                Exception
                  When Others Then
                    v_退指定结算方式 := '现金';
                End;
            End;
          End If;
          Update 病人预交记录
          Set 冲预交 = 冲预交 + (-1 * n_退费金额), 交易说明 = Nvl(交易说明_In, 交易说明)
          Where 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id And 结算方式 = v_退指定结算方式;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
               结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.病人id, a.主页id, a.科室id, a.摘要, v_退指定结算方式, d_Date,
                     操作员编号_In, 操作员姓名_In, -1 * n_退费金额, n_销帐id, n_组id, 预交类别, Decode(交易说明_In, Null, 卡类别id, Null), 结算卡序号,
                     Decode(交易说明_In, Null, 卡号, Null), Decode(交易说明_In, Null, 交易流水号, Null), Nvl(交易说明_In, 交易说明), 合作单位, 4
              From 病人预交记录 A
              Where a.记录性质 = 4 And a.记录状态 = 1 And a.结帐id = n_结帐id And Rownum < 2;
          End If;
        End If;
      End If;
    Else
      --退款金额获取
      If Nvl(退费类型_In, 0) <> 0 Or Nvl(n_二次退费, 0) <> 0 Then
        --如果是单独退病历费,或者只退挂号费,先获取退费金额
        Begin
          --获取本次退款金额
          Select -1 * Sum(Nvl(实收金额, 0)) As 收款金额
          Into n_退款金额
          From 门诊费用记录
          Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 2 And 结帐id = n_销帐id;
        Exception
          When Others Then
            v_Err_Msg := '单据【' || 单据号_In || '】的' || Case Nvl(退费类型_In, 0)
                           When 1 Then
                            '挂号费用'
                           When 2 Then
                            '病历费'
                         End || '可能由于并发原因已经进行了退费或者单据不存在!';
            Raise Err_Item;
        End;
      End If;
      If Nvl(n_二次退费, 0) = 0 And Nvl(退费类型_In, 0) = 0 Then
        --首次全退
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
           结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In, -1 * 冲预交,
                 n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_结帐id;
      Else
        --二次退费,或者本次单退一部分
        --二次退费时,记录状态=3 ,首次部分退,记录状态为1
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
           结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                 -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And 摘要 = '医保挂号' And 冲预交 = n_退款金额 And
                Rownum < 2;
        If Sql%RowCount = 0 Then
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And 冲预交 = n_退款金额 And Rownum < 2;
        End If;
        If Sql%RowCount = 0 Then
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id,
             结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
            Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_Date, 操作员编号_In, 操作员姓名_In,
                   -1 * n_退款金额, n_销帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
            From 病人预交记录
            Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id And Rownum < 2;
          If Sql%RowCount = 0 Then
            --部分退费,并且全部使用预交款缴费时才存在此种情况
            n_预交金额 := n_退款金额;
          End If;
        End If;
      
      End If;
    End If;
    --首次退费时,记录状态便调整为了3
    Update 病人预交记录
    Set 记录状态 = 3
    Where 记录性质 = 4 And 记录状态 = Decode(Nvl(n_二次退费, 0), 0, 1, 3) And 结帐id = n_结帐id;
  
    --冲预交 1-全退 2-部分退,部分退时当全部使用预交进行缴款
    If Nvl(退费类型_In, 0) = 0 Or (Nvl(退费类型_In, 0) <> 0 And n_预交金额 <> 0) Then
      --病人挂号结算:冲预交款部份
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_Date,
               操作员姓名_In, 操作员编号_In, -1 * Decode(Nvl(退费类型_In, 0) + Nvl(n_二次退费, 0), 0, 冲预交, n_预交金额), n_销帐id, n_组id, 预交类别,
               卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
        From 病人预交记录
        Where 记录性质 In (1, 11) And 结帐id = n_结帐id And Nvl(冲预交, 0) <> 0 And
              Rownum = Decode(Nvl(退费类型_In, 0) + Nvl(n_二次退费, 0), 0, Rownum, 1);
    End If;
  
    --处理病人预交余额
    For c_预交 In (Select 病人id, 预交类别, -1 * Sum(Nvl(冲预交, 0)) As 冲预交
                 From 病人预交记录
                 Where 记录性质 In (1, 11) And 结帐id = n_销帐id
                 Group By 病人id, 预交类别) Loop
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) + Nvl(c_预交.冲预交, 0)
      Where 病人id = c_预交.病人id And 类型 = Nvl(c_预交.预交类别, 2) And 性质 = 1
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 预交余额, 性质, 类型)
        Values
          (c_预交.病人id, Nvl(c_预交.冲预交, 0), 1, Nvl(c_预交.预交类别, 2));
        n_返回值 := Nvl(c_预交.冲预交, 0);
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = c_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
      End If;
    End Loop;
  
    If 收回票据号_In Is Not Null Then
      --光退挂号费,不回收票据
      --退卡收回票据(可能上次挂号使用票据,不能收回)
      Begin
        --从最后一次打印的内容中取
        Select ID
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 4 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          n_打印id := Null;
      End;
    
      --先收回原票据
      If n_打印id Is Not Null Then
        Begin
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In, 票据金额
            From 票据使用明细
            Where 打印id = n_打印id And 性质 = 1 And Instr(',' || 收回票据号_In || ',', ',' || 号码 || ',') > 0;
        Exception
          When Others Then
            Delete From 票据使用明细 Where 打印id = n_打印id And 性质 = 2 And 原因 = 2;
            Insert Into 票据使用明细
              (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
              Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_Date, 操作员姓名_In, 票据金额
              From 票据使用明细
              Where 打印id = n_打印id And 性质 = 1 And Instr(',' || 收回票据号_In || ',', ',' || 号码 || ',') > 0;
        End;
      End If;
    End If;
  End If;

  --单独退病历费用,不处理汇总记录
  --相关汇总表的处理

  --病人挂号汇总
  If Nvl(退费类型_In, 0) Not In (2, 3) Then
    Open c_Registinfo(3);
    Fetch c_Registinfo
      Into r_Registrow;
  
    If c_Registinfo%RowCount = 0 Then
      --只收病历费时无号别,不处理
      Close c_Registinfo;
    Else
    
      --需要确定是否预约挂号
      --1.如果是退预约挂号产生的挂号记录,则需要减已挂数和其中已接数
      --2.如果是正常挂号,则只减已挂数
    
      Begin
        Select Decode(预约, Null, 0, 0, 0, 1) Into n_预约挂号 From 病人挂号记录 Where NO = 单据号_In And Rownum = 1;
      Exception
        When Others Then
          n_预约挂号 := 0;
      End;
    
      Update 病人挂号汇总
      Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约挂号, 已约数 = Nvl(已约数, 0) - n_预约挂号
      Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
            (号码 = r_Registrow.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收, 已约数)
        Values
          (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
           Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1, -1 * n_预约挂号, -1 * n_预约挂号);
      End If;
    
      If n_出诊记录id Is Not Null Then
        Update 临床出诊记录
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约挂号, 已约数 = Nvl(已约数, 0) - n_预约挂号
        Where ID = n_出诊记录id And Nvl(已约数, 0) > 0;
        Update 临床出诊序号控制 Set 挂号状态 = Null, 操作员姓名 = Null Where 记录id = n_出诊记录id And 序号 = n_序号;
      End If;
    
      Close c_Registinfo;
    End If;
  End If;

  If n_记帐 = 0 Then
    --人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款)
    For r_Opermoney In c_Opermoney(n_销帐id) Loop
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) + (-1 * r_Opermoney.冲预交)
      Where 收款员 = 操作员姓名_In And 结算方式 = r_Opermoney.结算方式 And 性质 = 1
      Returning 余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (操作员姓名_In, r_Opermoney.结算方式, 1, -1 * r_Opermoney.冲预交);
        n_返回值 := r_Opermoney.冲预交;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 收款员 = 操作员姓名_In And 结算方式 = r_Opermoney.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
      End If;
    End Loop;
  End If;
  If Nvl(退费类型_In, 0) Not In (2, 3) Then
    n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
    If n_挂号生成队列 <> 0 Then
      --要删除队列
      For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
        n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113, 1, Nvl(v_挂号.执行部门id, 0)));
        If Nvl(n_分诊台签到排队, 0) = 0 Then
          Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
        End If;
      End Loop;
    End If;
  
    --医保产生的就诊登记记录
    Begin
      Select 病人id, 发生时间 Into n_就诊病人id, d_就诊时间 From 病人挂号记录 Where NO = 单据号_In;
      Delete From 就诊登记记录 Where 病人id = n_就诊病人id And 就诊时间 = d_就诊时间 And 主页id Is Null;
    Exception
      When Others Then
        Null;
    End;
  
    --病人挂号记录
    Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
  
    Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录性质 = 1 And 记录状态 = 1;
    If Sql%NotFound Then
      v_Err_Msg := '挂号单【' || 单据号_In || '】不存在或由于并发原因已经被退号';
      Raise Err_Item;
    End If;
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
       复诊, 号序, 社区, 预约, 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式)
      Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_Date, 发生时间,
             操作员编号_In, 操作员姓名_In, 复诊, 号序, 社区, 预约, Nvl(摘要_In, 摘要) As 摘要, 预约方式, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 险类, 医疗付款方式
      From 病人挂号记录
      Where NO = 单据号_In;
  End If;
  --消息推送
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 2, 单据号_In;
  Exception
    When Others Then
      Null;
  End;
  b_Message.Zlhis_Regist_003(n_挂号id, 单据号_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_Delete;
/





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0037' Where 编号=&n_System;
Commit;