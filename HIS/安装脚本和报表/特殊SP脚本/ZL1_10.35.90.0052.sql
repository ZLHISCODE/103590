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



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--132271:李南春,2019-03-04,加号时自动取序号
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
    是否挂号_In Number,
    合作单位Check_In   Varchar2 := Null
  ) Return Number Is
    --功能：检查是否超出了限号或限约
    --入参:是否挂号_IN-1:挂号;0-预约
    --返回:1-表示数据合法;0-表示数据不合法:超出了限号或限约数
    n_Count Number(18);
    n_Temp  Number(18);
  Begin
    If Nvl(计划id1_In, 0) <> 0 Then
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
           Where 号码 = 号码_In And 日期 between Trunc(日期_In) And Trunc(日期_In) + (1-1/24/60/60) And Nvl(状态, 0) <> 4
           Union
           Select 序号
           From 合作单位计划控制
           Where 计划id = Decode(是否挂号_In, 1, 0, 计划id1_In) And 合作单位 <> Nvl(合作单位Check_In,'-') And 序号 <> 0 And 限制项目 = 星期1_In And 数量 <> 0
           Union
           Select 序号
           From 合作单位安排控制
           Where 安排id = Decode(是否挂号_In, 1, 0, 安排id1_In) And 合作单位 <> Nvl(合作单位Check_In,'-') And 序号 <> 0 And 限制项目 = 星期1_In And 数量 <> 0);
  
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
    If Nvl(计划id1_In, 0) <> 0 Then
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
    
      If Check_Nums_Valied(n_安排id, n_计划id, v_星期, n_是否挂号, 合作单位_In) = 0 Then
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
      
        If Check_Nums_Valied(n_安排id, n_计划id, v_星期, n_是否挂号, 合作单位_In) = 0 Then
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
      If Check_Nums_Valied(n_安排id, n_计划id, v_星期, n_是否挂号, 合作单位_In) = 0 Then
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





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0052' Where 编号=&n_System;
Commit;
