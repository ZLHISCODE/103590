----10.35.70---》10.35.80
--00000:刘硕,2017-09-25,唯一键名命不符合规范调整
alter table Zltools.zlProgFuncs rename constraint zlProgFuncs_PK to zlProgFuncs_UQ_功能;
alter index Zltools.zlProgFuncs_PK rename to zlProgFuncs_UQ_功能;
Alter Table Zltools.zlProgFuncs Modify 序号  constraint zlProgFuncs_NN_序号   not  null;
alter table Zltools.zlProgPrivs rename constraint zlProgPrivs_PK to zlProgPrivs_UQ_权限;
alter index Zltools.zlProgPrivs_PK rename to zlProgPrivs_UQ_权限;
Alter Table Zltools.zlProgPrivs Modify 序号  constraint zlProgPrivs_NN_序号   not  null;
alter table Zltools.zlPrograms rename constraint zlPrograms_PK to zlPrograms_UQ_序号;
alter index Zltools.zlPrograms_PK rename to zlPrograms_UQ_序号;
Alter Table Zltools.zlPrograms Modify 序号  constraint zlPrograms_NN_序号   not  null;

--111526:高腾,2017-9-27,调试日志
Alter Table Zltools.Zllogconfig Modify(名称 Varchar2(50));
Alter Table Zltools.Zllogconfig Add(系统 Number(5));
Alter Table Zltools.Zllogconfig Drop Primary Key;
Alter Table Zltools.Zllogconfig Add Constraint Zllogconfig_Pk Primary Key(名称) Using Index;
Alter Table Zltools.Zllogconfig Add Constraint Zllogconfig_Uq_编号 Unique(系统, 编号) Using Index;
Alter Table Zltools.Zllogconfig Add Constraint Zllogconfig_Fk_系统 Foreign Key(系统) References Zlsystems(编号) On Delete Cascade;

--114503:高腾,2017-9-25,院区界面关联修改
Update Zltools.Zlnodelist Set 名称 = Decode(名称,'站点0','总院','站点1','一分院','站点2','二分院','站点3','三分院','站点4','四分院','站点5','五分院','站点6','六分院','站点7','七分院','站点8','八分院','站点9','九分院',名称);

--116691:杨周一,2017-11-29,远程控制帐号密码保存
Alter Table Zltools.zlClients Add(管理员用户 Varchar2(20));
Alter Table Zltools.zlClients Add(管理员密码 Varchar2(20));

--113406:高腾,2017-9-5,删除弃用工作站
Alter Table Zltools.Zlclients Add 最近登陆时间 Date;
Update Zltools.Zlsvrtools Set 标题 = '客户端运行控制' Where 编号 = '0308';

--113406:高腾,2017-9-5,删除弃用工作站
--116691:杨周一,2017-11-29,远程控制帐号密码保存
Create Or Replace Procedure Zltools.Zl_Zlclients_Set
(
  n_Mode_In       Number,
  n_Rowid_In      Varchar2 := Null,
  v_工作站_In     Zlclients.工作站%Type := Null,
  v_Ip_In         Zlclients.Ip%Type := Null,
  v_Cpu_In        Zlclients.Cpu%Type := Null,
  v_内存_In       Zlclients.内存%Type := Null,
  v_硬盘_In       Zlclients.硬盘%Type := Null,
  v_操作系统_In   Zlclients.操作系统%Type := Null,
  v_部门_In       Zlclients.部门%Type := Null,
  v_用途_In       Zlclients.用途%Type := Null,
  v_说明_In       Zlclients.说明%Type := Null,
  n_升级服务器_In Zlclients.升级服务器%Type := Null,
  n_升级标志_In   Zlclients.升级标志%Type := 0,
  n_连接数_In     Zlclients.连接数%Type := 0,
  v_站点_In       Zlclients.站点%Type := Null,
  n_Apply_In      Number := 0,
  v_Ipbegin_In    Varchar2 := Null,
  v_Ipend_In      Varchar2 := Null,
  n_启用视频源    Zlclients.启用视频源%Type := Null,
  v_管理员用户_In Zlclients.管理员用户%Type := Null,
  v_管理员密码_In Zlclients.管理员密码%Type := Null
  --功能：新增客户端或站点 或者更新客户端属性
  --应用：1、管理工具：新增或修改站点 （修改时以IP与客户端做判断条件，不需传入N_Rowid_In）
  --      2：应用系统：登录时根据当前登录的客户短来判断是否
  --                   新增站点或修改站点参数（更新时N_Rowid_In需传入）
  --站点设置:0-新增站点，1-更新站点
  --N_Apply_In,站点参数应用范围，0-本站点，1，本部门，2，所有站点，3，固定IP段
  --V_Ipbegin_In,V_Ipend_In:在固定IP断应用时传入,两者在一个IP断上，即前面部分相同
) Is
  n_Pos         Number(3);
  n_Ipbegin_Num Number;
  n_Ipend_Num   Number;
  n_Ip_Num      Number;
  n_Count       Number;

  v_Err Varchar2(500);
  Err_Custom Exception;

  Function Get_Ipnum(v_Ip_Input Varchar2) Return Number Is
    v_Ip_Num  Varchar2(20);
    n_Pos_Tmp Number;
    v_Ip_Tmp  Varchar2(20);
  Begin
    n_Pos_Tmp := Length(v_Ip_Input);
    n_Pos_Tmp := n_Pos_Tmp - Length(Replace(v_Ip_Input, '.', ''));
    If n_Pos_Tmp <> 3 Then
      Return Null;
    Else
      v_Ip_Tmp := v_Ip_Input;
      Loop
        n_Pos_Tmp := Instr(v_Ip_Tmp, '.');
        Exit When(Nvl(n_Pos_Tmp, 0) = 0);
        --将每一断数字转化为3位数
        v_Ip_Num := v_Ip_Num || Trim(To_Char(Substr(v_Ip_Tmp, 1, n_Pos_Tmp - 1), '099'));
        v_Ip_Tmp := Substr(v_Ip_Tmp, n_Pos_Tmp + 1);
      End Loop;
      v_Ip_Num := v_Ip_Num || Trim(To_Char(v_Ip_Tmp, '099'));
      n_Ip_Num := To_Number(Trim(v_Ip_Num));
      Return n_Ip_Num;
    End If;
  End;
Begin
  If n_Mode_In = 0 Then
  
    Select Count(1) Into n_Count From zlClients Where 工作站 = v_工作站_In;
    If n_Count = 0 Then
      Insert Into zlClients
        (Ip, 工作站, Cpu, 内存, 硬盘, 操作系统, 部门, 用途, 说明, 升级服务器, 升级标志, 连接数, 站点, 启用视频源, 最近登陆时间, 管理员用户, 管理员密码)
      Values
        (v_Ip_In, v_工作站_In, v_Cpu_In, v_内存_In, v_硬盘_In, v_操作系统_In, v_部门_In, v_用途_In, v_说明_In, n_升级服务器_In, n_升级标志_In,
         n_连接数_In, v_站点_In, n_启用视频源, Sysdate, v_管理员用户_In, v_管理员密码_In);
    Else
      v_Err := '已经设置了相同IP地址或工作站,不能再设!';
      Raise Err_Custom;
    End If;
  Else
    If n_Rowid_In Is Null Then
      Update zlClients
      Set Cpu = v_Cpu_In, 内存 = v_内存_In, 硬盘 = v_硬盘_In, 操作系统 = v_操作系统_In, 部门 = v_部门_In, 用途 = v_用途_In, 说明 = v_说明_In,
          连接数 = n_连接数_In, 站点 = v_站点_In, 启用视频源 = n_启用视频源, 升级服务器 = n_升级服务器_In, 升级标志 = n_升级标志_In, 最近登陆时间 = Sysdate,
          管理员用户 = Decode(v_管理员用户_In, '空空', 管理员用户, Nvl(v_管理员用户_In, 管理员用户)),
          管理员密码 = Decode(v_管理员密码_In, '空空', 管理员密码, Nvl(v_管理员密码_In, 管理员密码))
      Where 工作站 = v_工作站_In And Ip = v_Ip_In;
    Else
      Update zlClients
      Set 工作站 = v_工作站_In, Ip = v_Ip_In, Cpu = Decode(v_Cpu_In, Null, Cpu, v_Cpu_In),
          内存 = Decode(v_内存_In, Null, 内存, v_内存_In), 硬盘 = Decode(v_硬盘_In, Null, 硬盘, v_硬盘_In),
          操作系统 = Decode(v_操作系统_In, Null, 操作系统, v_操作系统_In), 部门 = v_部门_In, 站点 = v_站点_In, 启用视频源 = n_启用视频源, 最近登陆时间 = Sysdate,
          管理员用户 = Decode(v_管理员用户_In, '空空', 管理员用户, Nvl(v_管理员用户_In, 管理员用户)),
          管理员密码 = Decode(v_管理员密码_In, '空空', 管理员密码, Nvl(v_管理员密码_In, 管理员密码))
      Where Rowid = n_Rowid_In;
    End If;
  End If;
  --本部门
  If n_Apply_In = 1 Then
    Update zlClients
    Set 连接数 = n_连接数_In, 站点 = v_站点_In
    Where Nvl(部门, 'NONE') = Nvl(v_部门_In, 'NONE') And Ip <> v_Ip_In;
  Elsif n_Apply_In = 2 Then
    Update zlClients Set 连接数 = n_连接数_In, 站点 = v_站点_In Where Ip <> v_Ip_In;
  Elsif n_Apply_In = 3 Then
    n_Pos := Length(v_Ipbegin_In);
    n_Pos := n_Pos - Length(Replace(v_Ipbegin_In, '.', ''));
    If n_Pos <> 3 Then
      v_Err := '起始IP格式有误！';
      Raise Err_Custom;
    End If;
    n_Pos := Length(v_Ipend_In);
    n_Pos := n_Pos - Length(Replace(v_Ipend_In, '.', ''));
    If n_Pos <> 3 Then
      v_Err := '结束IP格式有误！';
      Raise Err_Custom;
    End If;
  
    n_Ipbegin_Num := Get_Ipnum(v_Ipbegin_In);
    n_Ipend_Num   := Get_Ipnum(v_Ipend_In);
    For r_Ip In (Select 工作站, Ip From zlClients) Loop
      n_Ip_Num := Get_Ipnum(r_Ip.Ip);
      If n_Ip_Num >= n_Ipbegin_Num And n_Ip_Num <= n_Ipend_Num Then
        Update zlClients Set 连接数 = n_连接数_In, 站点 = v_站点_In Where 工作站 = r_Ip.工作站 And Ip = r_Ip.Ip;
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Set;
/

--113406:高腾,2017-9-5,删除弃用工作站
Create Or Replace Procedure Zltools.Zl_Zlclients_Deletebatch Is
  d_登陆时间 Zlclients.最近登陆时间%Type;
Begin
  Select Min(最近登陆时间) Into d_登陆时间 From Zlclients;
  Delete Zlclients Where Add_Months(Nvl(最近登陆时间, d_登陆时间), 3) < Sysdate;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlclients_Deletebatch;
/

--113406:高腾,2017-9-26,清除三个月未登录客户端问题关联修改
--116691:杨周一,2017-11-29,远程控制帐号密码保存
Create Or Replace Package Zltools.b_Runmana Is

  Type t_Refcur Is Ref Cursor;

  Procedure Get_Parameters
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Number := 0
  );

  Procedure Get_Parameter
  (
    Cursor_Out Out t_Refcur,
    参数id_In  In Zlparameters.Id%Type
  );

  Procedure Get_Parachangedlog
  (
    Cursor_Out Out t_Refcur,
    参数id_In  In Zlparachangedlog.参数id%Type
  );

  Procedure Get_Job_Number
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Number
  );

  Procedure Get_Depict
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Zldatamove.系统%Type,
    组号_In    In Zldatamove.组号%Type
  );

  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur);

  Procedure Get_Client
  (
    Cur_Out   Out t_Refcur,
    工作站_In In Zlclients.工作站%Type := Null
  );

  Procedure Get_Client_Station(Cur_Out Out t_Refcur);

  Procedure Get_Project_No(Cur_Out Out t_Refcur);

  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur);

  Procedure Get_Resile
  (
    Cur_Out   Out t_Refcur,
    方案号_In In Zlclientparaset.方案号%Type,
    类型_In   In Number := 0
  );

  Procedure Get_Zldatamove
  (
    Cur_Out Out t_Refcur,
    系统_In In Zldatamove.系统%Type
  );

  Procedure Get_Log
  (
    Cur_Out     Out t_Refcur,
    日志类型_In In Varchar2,
    Where_In    In Varchar2
  );

  Procedure Get_Log_Count
  (
    Cur_Out     Out t_Refcur,
    日志类型_In In Varchar2
  );

  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur);

  Procedure Get_Not_Regist(Cur_Out Out t_Refcur);

  Procedure Get_Zloption
  (
    Cur_Out   Out t_Refcur,
    参数号_In In Zloptions.参数号%Type
  );

End b_Runmana;
/
--113406:高腾,2017-9-26,清除三个月未登录客户端问题关联修改
Create Or Replace Package Body Zltools.b_Runmana Is

  --功能：取参数信息
  --frmParameters
  Procedure Get_Parameters
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Number := 0
  ) Is
  Begin
    If Nvl(系统_In, 0) = 0 Then
      Open Cursor_Out For
        Select a.Id, Nvl(a.系统, 0) 系统, Nvl(a.模块, 0) 模块, Nvl(a.私有, 0) 私有, a.参数号, a.参数名, a.参数值, a.缺省值, Nvl(a.性质, 0) 性质,
               a.影响控制说明, a.参数值含义, a.关联说明, a.适用说明, a.警告说明, Nvl(a.本机, 0) 本机, Nvl(a.授权, 0) 授权, Nvl(a.固定, 0) 固定,
               Nvl(a.部门, 0) 部门, b.标题 As 模块名称, zlSpellCode(b.标题) As 模块简码
        From zlParameters A, zlPrograms B
        Where Nvl(a.系统, 0) = 0 And Nvl(a.系统, 0) = b.系统(+) And Nvl(a.模块, 0) = b.序号(+);
    Else
      Open Cursor_Out For
        Select a.Id, Nvl(a.系统, 0) 系统, Nvl(a.模块, 0) 模块, Nvl(a.私有, 0) 私有, a.参数号, a.参数名, a.参数值, a.缺省值, Nvl(a.性质, 0) 性质,
               a.影响控制说明, a.参数值含义, a.关联说明, a.适用说明, a.警告说明, Nvl(a.本机, 0) 本机, Nvl(a.授权, 0) 授权, Nvl(a.固定, 0) 固定,
               Nvl(a.部门, 0) 部门, b.标题 As 模块名称, zlSpellCode(b.标题) As 模块简码
        From zlParameters A, zlPrograms B,
             --处理权限部分，只有授权的才能显示
             (Select Distinct f.序号
               From zlProgFuncs F, zlRegFunc R
               Where Trunc(f.系统 / 100) = r.系统(+) And f.序号 = r.序号(+) And f.功能 = r.功能(+) And
                     (r.功能 Is Not Null Or r.功能 Is Null And (f.序号 Between 10000 And 19999)) And f.系统 = 系统_In And
                     1 = (Select 1 From Zlregaudit A Where a.项目 = '授权证章')
               Union All
               Select 0 As 序号 From Dual) M
        Where a.系统 = Nvl(系统_In, 0) And Nvl(a.系统, 0) = b.系统(+) And Nvl(a.模块, 0) = b.序号(+) And Nvl(a.模块, 0) = m.序号;
    End If;
  End Get_Parameters;

  --功能：根据指定的参数ID取参数信息
  --调用列表：frmParameters;frmParaChangeSet
  Procedure Get_Parameter
  (
    Cursor_Out Out t_Refcur,
    参数id_In  In Zlparameters.Id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select a.Id, Nvl(a.系统, 0) 系统, Nvl(a.模块, 0) 模块, Nvl(a.私有, 0) 私有, a.参数号, a.参数名, a.参数值, a.缺省值, Nvl(a.性质, 0) 性质,
             a.影响控制说明, a.参数值含义, a.关联说明, a.适用说明, a.警告说明, Nvl(a.本机, 0) 本机, Nvl(a.授权, 0) 授权, Nvl(a.固定, 0) 固定,
             Nvl(a.部门, 0) 部门, b.标题 As 模块名称, zlSpellCode(b.标题) As 模块简码
      From zlParameters A, zlPrograms B
      Where a.Id = Nvl(参数id_In, 0) And Nvl(a.系统, 0) = b.系统(+) And Nvl(a.模块, 0) = b.序号(+);
  End Get_Parameter;
  --功能：取参数修改信息
  --调用列表：frmParameters
  Procedure Get_Parachangedlog
  (
    Cursor_Out Out t_Refcur,
    参数id_In  In Zlparachangedlog.参数id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 参数id, 序号, 变动说明, 变动内容, 变动人, 变动时间, 变动原因
      From Zlparachangedlog
      Where 参数id = Nvl(参数id_In, 0);
  
  End;
  --功能：取ZlAutoJob序列号
  Procedure Get_Job_Number
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Number
  ) Is
  Begin
    Open Cursor_Out For
      Select 序号 + 1 As 序号
      From zlAutoJobs
      Where Nvl(系统, 0) = 系统_In And 类型 = 3 And
            序号 + 1 Not In (Select 序号 From zlAutoJobs Where Nvl(系统, 0) = 系统_In And 类型 = 3);
  End Get_Job_Number;

  --功能：取ZlDataMove描述
  Procedure Get_Depict
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Zldatamove.系统%Type,
    组号_In    In Zldatamove.组号%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 转出描述 From zlDataMove Where Nvl(系统, 0) = 系统_In And 组号 = 组号_In;
  End Get_Depict;

  --功能：取zlClients的MAX IP
  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Max(Ip) As Ip From zlClients;
  End Get_Client_Maxip;

  --功能：取zlClients的记录
  Procedure Get_Client
  (
    Cur_Out   Out t_Refcur,
    工作站_In In Zlclients.工作站%Type := Null
  ) Is
    v_Sql Varchar2(1000);
  Begin
    If Nvl(工作站_In, '空') = '空' Then
      v_Sql := 'Select a.Ip, a.工作站, a.Cpu, a.内存, a.硬盘, a.操作系统, a.部门, a.用途, a.说明, a.升级标志, a.禁止使用,
                             a.连接数, Decode(b.Terminal, Null, 0, 1) As 状态, a.收集标志,a.升级服务器,a.站点,c.名称 院区,a.启用视频源,a.最近登陆时间,a.管理员用户,a.管理员密码
                From Zlclients a, (Select Distinct Terminal From GV$session) b, zlnodelist c
                Where a.工作站 = b.Terminal(+) and a.站点 = c.编号(+)
                Order By a.站点, a.Ip';
      Open Cur_Out For v_Sql;
    Else
      Open Cur_Out For
        Select Ip, 工作站, Cpu, 内存, 硬盘, 操作系统, 部门, 用途, 说明, 升级标志, 收集标志, 禁止使用, 连接数, 升级服务器, 站点, 启用视频源, 管理员用户, 管理员密码
        From zlClients
        Where 工作站 = 工作站_In;
    End If;
  End Get_Client;

  --功能：取zlClients的站点
  Procedure Get_Client_Station(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Distinct Upper(工作站) || '[' || Ip || ']' As 站点, Upper(工作站) 工作站 From zlClients;
  End Get_Client_Station;

  --功能：取方案号
  Procedure Get_Project_No(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select 方案号 From Zlclientparaset Where Rownum = 1;
  End Get_Project_No;

  --功能：取方案
  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select 方案号, 方案号 || '-' || 方案名称 As 方案名称, 方案描述, 工作站, 用户名 From Zlclientscheme;
  End Get_Client_Scheme;

  --功能：取恢复信息
  Procedure Get_Resile
  (
    Cur_Out   Out t_Refcur,
    方案号_In In Zlclientparaset.方案号%Type,
    类型_In   In Number := 0
  ) Is
  Begin
    If 类型_In = 0 Then
      Open Cur_Out For
        Select Distinct a.工作站 || Decode(m.工作站, Null, ' ', '[' || m.Ip || ']') As 工作站, a.用户名, a.恢复标志,
                        '[' || b.方案号 || ']' || b.方案名称 As 方案名称
        From Zlclientparaset A, Zlclientscheme B, zlClients M
        Where a.方案号 = b.方案号 And a.工作站 = m.工作站(+) And a.方案号 = 方案号_In;
    End If;
  
    If 类型_In = 1 Then
      Open Cur_Out For
        Select Distinct Upper(工作站) 工作站, Min(恢复标志) 恢复标志
        From Zlclientparaset A
        Where a.方案号 = 方案号_In
        Group By 工作站;
    End If;
  
    If 类型_In = 2 Then
      Open Cur_Out For
        Select Distinct Upper(用户名) 用户名, Max(工作站) 工作站, Min(Decode(恢复标志, 2, 0, 恢复标志)) 恢复标志
        From Zlclientparaset A
        Where a.方案号 = 方案号_In
        Group By 用户名
        Order By 用户名;
    End If;
  
  End Get_Resile;

  --功能：取zldataMove数据
  Procedure Get_Zldatamove
  (
    Cur_Out Out t_Refcur,
    系统_In In Zldatamove.系统%Type
  ) Is
  Begin
    Open Cur_Out For
      Select 组号, 组名, 说明, 日期字段, 转出描述, 上次日期 From zlDataMove Where 系统 = 系统_In Order By 组号;
  End Get_Zldatamove;

  --功能：取日志数据
  Procedure Get_Log
  (
    Cur_Out     Out t_Refcur,
    日志类型_In In Varchar2,
    Where_In    In Varchar2
  ) Is
    v_Sql Varchar2(1000);
  Begin
    If 日志类型_In = '错误日志' Then
      v_Sql := 'Select 会话号,工作站,用户名,错误序号,错误信息,To_char(时间,''yyyy-MM-dd hh24:mi:ss'') 时间
                     ,Decode(类型,1,''存储过程错误'',2,''数据联结层错误'',3,''应用程序层错误'',''客户端升级错误'') 错误类型
                        From ZlErrorLog Where ' || Where_In;
      Open Cur_Out For v_Sql;
    End If;
    If 日志类型_In = '运行日志' Then
      v_Sql := 'Select 会话号,工作站,用户名,部件名,工作内容,To_char(进入时间,''yyyy-MM-dd hh24:mi:ss'') 进入时间
                                 ,To_char(退出时间,''yyyy-MM-dd hh24:mi:ss'') 退出时间,Decode(退出原因,1,''正常退出'',''异常退出'') 退出原因
                                    From ZlDiaryLog Where ' || Where_In;
      Open Cur_Out For v_Sql;
    End If;
  End Get_Log;

  --功能：取日志记录数
  Procedure Get_Log_Count
  (
    Cur_Out     Out t_Refcur,
    日志类型_In In Varchar2
  ) Is
  Begin
    If 日志类型_In = '错误日志' Then
      Open Cur_Out For
        Select Count(*) 数量
        From zlErrorLog
        Union All
        Select Nvl(To_Number(参数值), 0) From zlOptions Where 参数号 = 4;
    End If;
    If 日志类型_In = '运行日志' Then
      Open Cur_Out For
        Select Count(*) 数量
        From zlDiaryLog
        Union All
        Select Nvl(To_Number(参数值), 0) From zlOptions Where 参数号 = 2;
    
    End If;
  End Get_Log_Count;

  --功能：取zlfilesupgradeg数据
  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select 序号, 文件名, 版本号, 修改日期, 文件说明 As 说明,
             Decode(文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '') As 类型, 安装路径 As 安装路径,
             Md5 As Md5, 加入日期
      From zlFilesUpgrade
      Order By 序号;
  End Get_Zlfilesupgrade;

  --功能：取非注册项目
  Procedure Get_Not_Regist(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select 项目, 内容
      From zlRegInfo
      Where 项目 Not In ('发行码', '版本号', '服务器目录', '访问用户', '访问密码', '收集目录', '收集类型', '注册码', '授权证章', '授权工具', '授权邮戳');
  End Get_Not_Regist;

  --功能：取参数值
  Procedure Get_Zloption
  (
    Cur_Out   Out t_Refcur,
    参数号_In In Zloptions.参数号%Type
  ) Is
  Begin
    Open Cur_Out For
      Select Nvl(参数值, 缺省值) Option_Value From zlOptions Where 参数号 = 参数号_In;
  End Get_Zloption;

End b_Runmana;
/

--109444:杨周一,2017-09-28,消息状态结构修正
drop table zltools.ZLPROCEDURENOTE;
drop INDEX zltools.zlMsgState_IX_消息ID;
Alter Table Zltools.ZLMSGSTATE Add Constraint ZLMSGSTATE_PK PRIMARY KEY(消息ID,类型,用户,身份) Using Index;


--109444:杨周一,2017-09-28,zlFilesExpired数据修正
Delete From Zltools.zlFilesExpired Where (文件名, 安装路径) In (Select 文件名, 安装路径 From Zltools.zlFilesExpired Group By 文件名, 安装路径 Having Count(1) > 1) And
Rowid Not In (Select Min(Rowid) From Zltools.zlFilesExpired Group By 文件名, 安装路径);

Alter Table zltools.zlFilesExpired Add Constraint zlFilesExpired_PK PRIMARY KEY(文件名,安装路径) Using Index;



Begin
--客户端升级日志结构修正：zlClientUpdateLog
  If Zl_Checkobject(1, 'ZLCLIENTUPDATELOG_bak') = 0 Then
    Execute Immediate 'Alter Table zltools.Zlclientupdatelog Rename To Zlclientupdatelog_Bak';
    Execute Immediate 'Create Table zltools.Zlclientupdatelog As Select * From zltools.Zlclientupdatelog_Bak Where 1 = 0';
  End If;
End;
/
Alter Table zltools.Zlclientupdatelog Add (顺序号 Number(5));
Alter Table zltools.Zlclientupdatelog Add Constraint Zlclientupdatelog_PK PRIMARY KEY (处理日期,工作站,顺序号) Using Index;


Drop Index zltools.ZLDIARYLOG_IX_会话任务;
Drop Index zltools.ZLDIARYLOG_IX_进入时间;
Begin
--运行日志结构修正:zlDiaryLog
  If Zl_Checkobject(1, 'Zldiarylog_bak') = 0 Then
    Execute Immediate 'Alter table zltools.Zldiarylog rename to Zldiarylog_bak';
    Execute Immediate 'Create table zltools.zldiarylog as select * from zltools.zldiarylog_bak where 1=0';
  End If;
End;
/
ALTER TABLE zltools.Zldiarylog ADD CONSTRAINT Zldiarylog_PK PRIMARY KEY (进入时间,会话号,窗体名) USING INDEX;

Drop Index zltools.ZLERRORLOG_IX_时间;
Begin
  If Zl_Checkobject(1, 'zlErrorLog_bak') = 0 Then
    Execute Immediate 'Alter table zltools.zlErrorLog rename to zlErrorLog_bak';
    Execute Immediate 'Create table zltools.zlErrorLog as select * from zltools.zlErrorLog_bak where 1=0';
  End If;
End;
/
Alter table zltools.zlErrorLog Add (顺序号 Number(5));
Alter Table zltools.zlErrorLog Add Constraint zlErrorLog_PK PRIMARY KEY(时间,会话号,错误序号,顺序号) Using Index;

Update zltools.zlOptions Set 参数值 = decode(sign(参数值-1000),1,1000,参数值), 缺省值 = 1000, 参数名 = '日志保存最大天数', 参数说明 = '日志最多能保存的天数，超过时系统将其自动删除。' Where 参数号 = 2;
Update zltools.zlOptions Set 参数值 = decode(sign(参数值-1000),1,1000,参数值), 缺省值 = 1000, 参数名 = '错误保存最大天数', 参数说明 = '错误最多能保存的天数，超过时系统将其自动删除。' Where 参数号 = 4;

Create Or Replace Procedure Zltools.Zl_Autologprocess As
  --功能：
  --   对多余的运行日志和错误日志进行清除
  v_Limit Number;
Begin
  --删除多余的运行日志
  Select Nvl(Max(To_Number(参数值)), 0) Into v_Limit From zlOptions Where 参数号 = 2;
  Delete From zlDiaryLog Where 进入时间 < Sysdate - v_Limit;

  --删除多余的错误日志

  Select Nvl(Max(To_Number(参数值)), 0) Into v_Limit From zlOptions Where 参数号 = 4;
  Delete From zlErrorLog Where 时间 < Sysdate - v_Limit;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Autologprocess;
/

Create Or Replace Procedure Zltools.Zl_Zlerrorlog_Insert
(
  工作站_In    Zlerrorlog.工作站%Type,
  类型_In      Zlerrorlog.类型%Type,
  错误序号_In  Zlerrorlog.错误序号%Type,
  错误信息_In  Zlerrorlog.错误信息%Type,
  Sessionid_In Number := Null
) Is
  n_Audsid Number;
  --功能：
  --   错误日志插入
Begin
  If Sessionid_In Is Null Then
    Select Userenv('SessionID') Into n_Audsid From Dual;
  Else
    n_Audsid := Sessionid_In;
  End If;
  Insert Into zlErrorLog
    (会话号, 用户名, 工作站, 时间, 类型, 错误序号, 错误信息, 顺序号)
    Select n_Audsid, User, 工作站_In, Sysdate, 类型_In, 错误序号_In, 错误信息_In, Count(1) + 1
    From zlErrorLog
    Where 会话号 = n_Audsid And 时间 = Sysdate And 错误序号 = 错误序号_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlerrorlog_Insert;
/

--   运行日志插入
Create Or Replace Procedure Zltools.Zl_Zldiarylog_Insert
(
  工作站_In    Zldiarylog.工作站%Type,
  部件名_In    Zldiarylog.部件名%Type,
  窗体名_In    Zldiarylog.窗体名%Type,
  工作内容_In  Zldiarylog.工作内容%Type,
  Sessionid_In Number := Null
) Is
  n_Audsid Number;
  --功能：
  --   运行日志插入
Begin
  If Sessionid_In Is Null Then
    Select Userenv('SessionID') Into n_Audsid From dual;
  Else
    n_Audsid := Sessionid_In;
  End If;
  Insert Into zlDiaryLog
    (会话号, 用户名, 工作站, 部件名, 窗体名, 工作内容, 进入时间)
    Select n_Audsid, User, 工作站_In, 部件名_In, 窗体名_In, 工作内容_In, Sysdate From Dual;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zldiarylog_Insert;
/

--   运行日志更新
Create Or Replace Procedure Zltools.Zl_Zldiarylog_Update
(
  工作站_In    Zldiarylog.工作站%Type,
  部件名_In    Zldiarylog.部件名%Type,
  窗体名_In    Zldiarylog.窗体名%Type,
  退出原因_In  Zldiarylog.退出原因%Type,
  Sessionid_In Number := Null
) Is
  n_会话号 Zldiarylog.会话号%Type;
  --功能：
  --   运行日志更新
Begin
  If Sessionid_In Is Null Then
    Select Userenv('SessionID') Into n_会话号 From Dual;
  Else
    n_会话号 := Sessionid_In;
  End If;
  Update zlDiaryLog
  Set 退出原因 = 退出原因_In, 退出时间 = Sysdate
  Where 退出原因 Is Null And 用户名 = User And 工作站 = 工作站_In And 会话号 = n_会话号 And 部件名 = 部件名_In And 窗体名 = 窗体名_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zldiarylog_Update;
/

--   客户端升级日志插入
Create Or Replace Procedure Zltools.Zl_Zlclientupdatelog_Insert
(
  内容_In   Zlclientupdatelog.内容%Type,
  工作站_In Zlclientupdatelog.工作站%Type
) Is
  v_工作站 Zlclientupdatelog.工作站%Type;
  --功能：
  --   客户端升级日志插入
Begin

  If 工作站_In Is Null Then
    Select Userenv('Terminal') Into v_工作站 From Dual;
  Else
    v_工作站 := 工作站_In;
  End If;
  Insert Into Zlclientupdatelog
    (工作站, 处理日期, 内容, 顺序号)
    Select v_工作站, Sysdate, 内容_In, Count(1) + 1
    From Zlclientupdatelog
    Where 工作站 = v_工作站 And 处理日期 = Sysdate;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclientupdatelog_Insert;
/

  --   自动升级插入
 Create Or Replace Procedure Zltools.Zl_Zlfilesexpired_Insert
 (
   文件名_In   Zlfilesexpired.文件名%Type,
   安装路径_In Zlfilesexpired.安装路径%Type,
   系统编号_In Zlfilesexpired.系统编号%Type,
   系统版本_In Zlfilesexpired.系统版本%Type,
   说明_In     Zlfilesexpired.说明%Type
 ) Is
  --功能：
  --   自动升级插入
 Begin
   Insert Into Zlfilesexpired
     (文件名, 安装路径, 系统编号, 系统版本, 说明)
     Select 文件名_In, 安装路径_In, 系统编号_In, 系统版本_In, 说明_In From Dual;
 Exception
   When Others Then
     zl_ErrorCenter(SQLCode, SQLErrM);
 End Zl_Zlfilesexpired_Insert;
/

--杨周一,创建包头
Create Or Replace Package zltools.b_Comfunc Is
--主要用于公共部件的过程
  Type t_Refcur Is Ref Cursor;
--功能：保存错误日志
--调用列表：clsComLib.SaveErrLog
  Procedure Save_Error_Log
  (
    类型_In     In zlErrorLog.类型%Type,
    错误序号_In In zlErrorLog.错误序号%Type,
    错误信息_In In zlErrorLog.错误信息%Type
  );
--功能：取可用功能
--调用列表：clsComLib.ShowAbout
  Procedure Get_Usable_Function
  (
    Cursor_Out Out t_Refcur,
    部件_In    In zlPrograms.部件%Type
  );
--功能：取大写金额
--调用列表：clsCommFun.UppeMoney
  Procedure Get_Uppmoney
  (
    Cursor_Out Out t_Refcur,
    金额_In    In Number
  );
--功能：根据指定的日期、组号、系统判断指定日期的数据是否已转出到后备数据表中
--调用列表：clsDatabase.DateMoved
  Procedure Get_Datamoved
  (
    Cursor_Out  Out t_Refcur,
    组号_In     In zlDataMove.组号%Type,
    系统_In     In zlDataMove.系统%Type,
    上次日期_In In zlDataMove.上次日期%Type
  );
--功能：取系统所有者
--调用列表：clsDatabase.GetOwner
  Procedure Get_Owner
  (
    Cursor_Out Out t_Refcur,
    编号_In    In zlSystems.编号%Type
  );
--功能：取简码
--调用列表：clsCommFun.SpellCode
  Procedure Get_Spell_Code
  (
    Cursor_Out Out t_Refcur,
    字符串_In  In Varchar2,
    方式_In    In Number := 0
  );
--功能：保存运行日志
--调用列表：clsComLib.RestoreWinState
  Procedure Save_Diary_Log
  (
    部件名_In   In zlDiaryLog.部件名%Type,
    窗体名_In   In zlDiaryLog.窗体名%Type,
    工作内容_In In zlDiaryLog.工作内容%Type
  );
--功能：更改运行日志
--调用列表：clsComLib.SaveWinState
  Procedure Update_Diary_Log
  (
    部件名_In In zlDiaryLog.部件名%Type,
    窗体名_In In zlDiaryLog.窗体名%Type
  );
--功能：取固定发布报表和用户发布报表
--调用列表：clsDatabase.ShowReportMenu
  Procedure Get_Report_Menu
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlPrograms.系统%Type,
    序号_In    In zlPrograms.序号%Type,
    功能_In    In zlReports.功能%Type,
    编号_In    In zlReports.编号%Type
  );
--功能：取用户提醒信息
--调用列表：zlApptools.frmAlert
  Procedure Get_Zlnoticerec
  (
    Cursor_Out Out t_Refcur,
    用户名_In  In zlNoticeRec.用户名%Type
  );
--功能：取邮件正文
--调用列表：zlApptools.frmMessageEdit.LoadMessage
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type,
    类型_In    In zlMsgState.类型%Type,
    用户_In    In zlMsgState.用户%Type
  );
--功能：取邮件内容
--调用列表：zlApptools.frmMessageManager.FillText，zlApptools.frmMessageRelate.FillText
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type
  );
--功能：取邮递地址
--调用列表：zlApptools.frmMessageEdit.LoadMessage
  Procedure Get_Zlmsgstate
  (
    Cursor_Out Out t_Refcur,
    消息id_In  In zlMsgState.消息id%Type
  );
--功能：删除消息
--调用列表：zlApptools.frmMessageManager.mnuEditDelete_Click
  Procedure Delete_Zlmsgstate
  (
    删除_In   In zlMsgState.删除%Type,
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  );
--功能：删除过期消息
--调用列表：zlApptools.frmMessageManager.DeleteMessage
  Procedure Delete_Zlmessage;
--功能：取邮件列表
--调用列表：zlApptools.frmMessageManager.FillList，zlApptools.frmMessageRelate.FillList
  Procedure Get_Mail_List
  (
    Cursor_Out  Out t_Refcur,
    消息类型_In In Varchar2,
    用户_In     In zlMsgState.用户%Type,
    显示已读_In In Number,
    会话id_In   In zlMessages.会话id%Type
  );
--功能：还原删除的消息
--调用列表：zlApptools.frmMessageManager.mnuEditRestore_Click
  Procedure Restore_Zlmsgstate
  (
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  );
--功能：保存主表消息
--调用列表：zlApptools.frmMessageEdit.SaveMessage
  Procedure Save_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type,
    会话id_In  In zlMessages.会话id%Type,
    发件人_In  In zlMessages.发件人%Type,
    收件人_In  In zlMessages.收件人%Type,
    主题_In    In zlMessages.主题%Type,
    内容_In    In zlMessages.内容%Type,
    背景色_In  In zlMessages.背景色%Type
  );
--功能：插入zlMsgstate
--调用列表：zlApptools.frmMessageEdit.SaveMessage
  Procedure Insert_Zlmsgstate
  (
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type,
    身份_In   In zlMsgState.身份%Type,
    删除_In   In zlMsgState.删除%Type,
    状态_In   In zlMsgState.状态%Type
  );
--功能：为原件加上答复或转发标志
--调用列表：zlApptools.frmMessageEdit.SaveMessage
  Procedure Update_Zlmsgstate_State
  (
    模式_In   In Number,
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  );
--功能：为原件加上答复或转发标志
--调用列表：zlApptools.frmMessageEdit.LoadMessage
  Procedure Update_Zlmsgstate_Idtntify
  (
    身份_In   In zlMsgState.身份%Type,
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  );

End b_Comfunc;
/

Create Or Replace Package Body Zltools.b_Comfunc Is
  --功能：保存错误日志
  Procedure Save_Error_Log
  (
    类型_In     In Zlerrorlog.类型%Type,
    错误序号_In In Zlerrorlog.错误序号%Type,
    错误信息_In In Zlerrorlog.错误信息%Type
  ) Is
    n_会话号 Number;
    v_工作站 Zlerrorlog.工作站%Type;
  Begin
    Select Userenv('SessionID'), Userenv('Terminal') Into n_会话号, v_工作站 From Dual;
    Insert Into zlErrorLog
      (会话号, 用户名, 工作站, 时间, 类型, 错误序号, 错误信息, 顺序号)
      Select n_会话号, User, v_工作站, Sysdate, 类型_In, 错误序号_In, 错误信息_In, Count(1) + 1
      From zlErrorLog
      Where 会话号 = n_会话号 And 时间 = Sysdate And 错误序号 = 错误序号_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Save_Error_Log;

  --功能：取可用功能
  Procedure Get_Usable_Function
  (
    Cursor_Out Out t_Refcur,
    部件_In    In Zlprograms.部件%Type
  ) Is
  Begin
    If Nvl(部件_In, '空空') = '空空' Then
      Open Cursor_Out For
        Select Distinct a.序号, a.标题, a.说明
        From zlPrograms A, zlProgFuncs B, zlRegFunc C
        Where a.系统 = b.系统 And a.序号 = b.序号 And Trunc(b.系统 / 100) = c.系统(+) And b.序号 = c.序号(+) And b.功能 = c.功能(+) And
              (c.功能 Is Not Null Or c.功能 Is Null And (a.序号 Between 10000 And 19999))
        Order By a.序号;
    Else
      Open Cursor_Out For
        Select Distinct a.序号, a.标题, a.说明
        From zlPrograms A, zlProgFuncs B, zlRegFunc C
        Where a.系统 = b.系统 And a.序号 = b.序号 And Upper(a.部件) = Upper(部件_In) And Trunc(b.系统 / 100) = c.系统(+) And
              b.序号 = c.序号(+) And b.功能 = c.功能(+) And
              (c.功能 Is Not Null Or c.功能 Is Null And (a.序号 Between 10000 And 19999))
        Order By a.序号;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Usable_Function;

  --功能：取大写金额
  Procedure Get_Uppmoney
  (
    Cursor_Out Out t_Refcur,
    金额_In    In Number
  ) Is
  Begin
    Open Cursor_Out For
      Select zlUppMoney(Nvl(金额_In, 0)) As Num From Dual;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Uppmoney;

  --功能：根据指定的日期、组号、系统判断指定日期的数据是否已转出到后备数据表中
  Procedure Get_Datamoved
  (
    Cursor_Out  Out t_Refcur,
    组号_In     In Zldatamove.组号%Type,
    系统_In     In Zldatamove.系统%Type,
    上次日期_In In Zldatamove.上次日期%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 系统, 组号
      From zlDataMove
      Where 组号 = 组号_In And 系统 = 系统_In And 上次日期 > 上次日期_In And 上次日期 Is Not Null;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Datamoved;

  --功能：取系统所有者
  Procedure Get_Owner
  (
    Cursor_Out Out t_Refcur,
    编号_In    In Zlsystems.编号%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 所有者 From zlSystems Where 编号 = 编号_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Owner;

  --功能：取简码
  Procedure Get_Spell_Code
  (
    Cursor_Out Out t_Refcur,
    字符串_In  In Varchar2,
    方式_In    In Number := 0
  ) Is
  Begin
    If Nvl(方式_In, 0) = 0 Then
      Open Cursor_Out For
        Select zlSpellCode(字符串_In) As 简码 From Dual;
    Else
      Open Cursor_Out For
        Select zlWbCode(字符串_In) As 简码 From Dual;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Spell_Code;

  --功能：保存运行日志
  Procedure Save_Diary_Log
  (
    部件名_In   In Zldiarylog.部件名%Type,
    窗体名_In   In Zldiarylog.窗体名%Type,
    工作内容_In In Zldiarylog.工作内容%Type
  ) Is
  Begin
    Insert Into zlDiaryLog
      (会话号, 用户名, 工作站, 部件名, 窗体名, 工作内容, 进入时间)
      Select Userenv('SessionID'), User, RTrim(LTrim(Replace(Userenv('Terminal'), Chr(0), ''))), 部件名_In, 窗体名_In, 工作内容_In, Sysdate
      From dual;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Save_Diary_Log;

  --功能：更改运行日志
  --调用列表：clsComLib.SaveWinState
  Procedure Update_Diary_Log
  (
    部件名_In In Zldiarylog.部件名%Type,
    窗体名_In In Zldiarylog.窗体名%Type
  ) Is
    Cursor c_Session Is
      Select Userenv('SessionID') As 会话号, User As 用户名, RTrim(LTrim(Replace(Userenv('Terminal'), Chr(0), ''))) As 工作站
      From dual;
  Begin
    For r_Tmp In c_Session Loop
      Update zlDiaryLog
      Set 退出原因 = 1, 退出时间 = Sysdate
      Where 退出原因 Is Null And 用户名 = r_Tmp.用户名 And 工作站 = r_Tmp.工作站 And 会话号 = r_Tmp.会话号 And 部件名 = 部件名_In And 窗体名 = 窗体名_In;
    End Loop;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Update_Diary_Log;

  --功能：取固定发布报表和用户发布报表
  Procedure Get_Report_Menu
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Zlprograms.系统%Type,
    序号_In    In Zlprograms.序号%Type,
    功能_In    In Zlreports.功能%Type,
    编号_In    In Zlreports.编号%Type
  ) Is
  Begin
    If Nvl(编号_In, '空空') <> '空空' Then
      Open Cursor_Out For
        Select 标志, 系统, 编号, 名称
        From (Select 1 As 标志, a.系统, a.编号, a.名称
               From zlReports A, zlPrograms B
               Where a.系统 = b.系统 And a.程序id = b.序号 And Not Upper(a.编号) Like '%BILL%' And
                     Upper(b.部件) <> Upper('zl9Report') And b.系统 = 系统_In And b.序号 = 序号_In And
                     Instr(功能_In, ';' || a.功能 || ';') > 0
               Union All
               Select Decode(a.系统, Null, 2, 1) As 标志, a.系统, a.编号, a.名称
               From zlReports A, zlRPTPuts B, zlPrograms C
               Where a.Id = b.报表id And b.系统 = c.系统 And b.程序id = c.序号 And (Not Upper(a.编号) Like '%BILL%' Or a.系统 Is Null) And
                     Instr(功能_In, ';' || b.功能 || ';') > 0 And c.系统 = 系统_In And c.序号 = 序号_In)
        Where Instr(编号_In, ',' || 编号 || ',') = 0
        Order By 标志, 编号;
    Else
      Open Cursor_Out For
        Select 标志, 系统, 编号, 名称
        From (Select 1 As 标志, a.系统, a.编号, a.名称
               From zlReports A, zlPrograms B
               Where a.系统 = b.系统 And a.程序id = b.序号 And Not Upper(a.编号) Like '%BILL%' And
                     Upper(b.部件) <> Upper('zl9Report') And b.系统 = 系统_In And b.序号 = 序号_In And
                     Instr(功能_In, ';' || a.功能 || ';') > 0
               Union All
               Select Decode(a.系统, Null, 2, 1) As 标志, a.系统, a.编号, a.名称
               From zlReports A, zlRPTPuts B, zlPrograms C
               Where a.Id = b.报表id And b.系统 = c.系统 And b.程序id = c.序号 And (Not Upper(a.编号) Like '%BILL%' Or a.系统 Is Null) And
                     Instr(功能_In, ';' || b.功能 || ';') > 0 And c.系统 = 系统_In And c.序号 = 序号_In)
        Order By 标志, 编号;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Report_Menu;

  --功能：取用户提醒信息
  Procedure Get_Zlnoticerec
  (
    Cursor_Out Out t_Refcur,
    用户名_In  In Zlnoticerec.用户名%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select a.序号, a.系统, c.程序id As 模块, c.系统 As 报表系统, b.提醒内容 As 结果内容, c.名称 As 提醒报表, a.提醒声音, b.检查时间, b.已读标志
      From zlNotices A, zlNoticeRec B, (Select * From zlReports Where 发布时间 Is Not Null) C
      Where b.用户名 = 用户名_In And b.提醒标志 > 0 And c.编号(+) = a.提醒报表 And a.序号 = b.提醒序号 And b.提醒内容 Is Not Null;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlnoticerec;

  --功能：取邮件正文
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In Zlmessages.Id%Type,
    类型_In    In Zlmsgstate.类型%Type,
    用户_In    In Zlmsgstate.用户%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select a.*, b.删除, b.状态
      From zlMessages A, zlMsgState B
      Where a.Id = b.消息id And b.消息id = Id_In And b.类型 = 类型_In And b.用户 = 用户_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlmessage;

  --功能：取邮件内容
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In Zlmessages.Id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 内容, 背景色 From zlMessages Where ID = Id_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlmessage;

  --功能：取邮递地址
  Procedure Get_Zlmsgstate
  (
    Cursor_Out Out t_Refcur,
    消息id_In  In Zlmsgstate.消息id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 类型, 用户, 身份 From zlMsgState Where 消息id = 消息id_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlmsgstate;

  --功能：删除消息
  Procedure Delete_Zlmsgstate
  (
    删除_In   In Zlmsgstate.删除%Type,
    消息id_In In Zlmsgstate.消息id%Type,
    类型_In   In Zlmsgstate.类型%Type,
    用户_In   In Zlmsgstate.用户%Type
  ) Is
    n_总数 Number(10);
    n_数量 Number(10);
  Begin
    If Nvl(删除_In, 0) = 1 Then
      Update zlMsgState Set 删除 = 1 Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
    Else
      If 类型_In = 0 Then
        --对于草稿，把收件人的也一并删除
        Update zlMsgState Set 删除 = 2 Where 消息id = 消息id_In And 用户 = 用户_In;
      Else
        Update zlMsgState Set 删除 = 2 Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
      End If;
      -- 删除指定ID的消息  mnuEditDelete_Click 调用
      Select Count(*) As 总数, Sum(Decode(删除, 2, 1, 0)) As 数量
      Into n_总数, n_数量
      From zlMsgState
      Where 消息id = 消息id_In;

      If n_总数 = n_数量 Then
        Delete From zlMessages Where ID = 消息id_In;
      End If;
    End If;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Delete_Zlmsgstate;

  --功能：删除过期消息
  Procedure Delete_Zlmessage Is
    n_Days Number;
  Begin
    Select Nvl(参数值, 缺省值) Into n_Days From zlOptions Where 参数号 = 5;
    If Nvl(n_Days, 0) > 0 Then
      Delete From zlMessages Where 时间 < Sysdate - n_Days;
      Commit;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Delete_Zlmessage;

  --功能：取邮件列表
  Procedure Get_Mail_List
  (
    Cursor_Out  Out t_Refcur,
    消息类型_In In Varchar2,
    用户_In     In Zlmsgstate.用户%Type,
    显示已读_In In Number,
    会话id_In   In Zlmessages.会话id%Type
  ) Is
    v_Sql  Varchar2(1000);
    v_已读 Varchar2(100);
    v_类型 Varchar2(100);
  Begin

    If Nvl(显示已读_In, 0) = 1 Then
      v_已读 := ' and substr(S.状态,1,1)=''0''';
    Else
      v_已读 := '';
    End If;

    If Instr(';草稿;收件箱;已发送消息;已删除消息;相关消息;', ';' || 消息类型_In || ';') <= 0 Then
      v_类型 := '草稿';
    Else
      v_类型 := 消息类型_In;
    End If;

    If v_类型 = '草稿' Then
      v_Sql := 'Select M.ID, M.会话id, M.发件人, M.收件人, M.主题, To_Char(M.时间, ''YYYY-MM-DD HH24:MI:SS'') As 时间, S.类型, S.状态
              From zlMessages M, zlMsgState S
              Where M.ID = S.消息id  and S.删除=0 and S.用户= ''' || 用户_In || ''' And S.类型=0 ' || v_已读;
    End If;

    If v_类型 = '收件箱' Then
      v_Sql := 'Select M.ID, M.会话id, M.发件人, M.收件人, M.主题, To_Char(M.时间, ''YYYY-MM-DD HH24:MI:SS'') As 时间, S.类型, S.状态
              From zlMessages M, zlMsgState S
              Where M.ID = S.消息id  and S.删除=0 and S.用户= ''' || 用户_In || ''' And S.类型=2 ' || v_已读;
    End If;

    If v_类型 = '已发送消息' Then
      v_Sql := 'Select M.ID, M.会话id, M.发件人, M.收件人, M.主题, To_Char(M.时间, ''YYYY-MM-DD HH24:MI:SS'') As 时间, S.类型, S.状态
              From zlMessages M, zlMsgState S
              Where M.ID = S.消息id  and S.删除=0 and S.用户= ''' || 用户_In || ''' And S.类型=1 ' || v_已读;
    End If;

    If v_类型 = '已删除消息' Then
      v_Sql := 'Select M.ID, M.会话id, M.发件人, M.收件人, M.主题, To_Char(M.时间, ''YYYY-MM-DD HH24:MI:SS'') As 时间, S.类型, S.状态
              From zlMessages M, zlMsgState S
              Where M.ID = S.消息id  and S.用户= ''' || 用户_In || ''' And S.删除=1 ' || v_已读;
    End If;

    If v_类型 = '相关消息' Then
      v_Sql := 'select M.ID,M.会话ID,M.发件人,M.收件人,M.主题,to_char(M.时间,''YYYY-MM-DD HH24:MI:SS'') as 时间,S.类型,S.状态
         from zlMessages M,zlMsgState S where M.ID=S.消息ID and S.删除<>2 and S.用户= ''' || 用户_In ||
               '''  and M.会话ID=' || 会话id_In;
    End If;

    If Nvl(v_Sql, '空空') <> '空空' Then
      Open Cursor_Out For v_Sql;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Mail_List;

  --功能：还原删除的消息
  Procedure Restore_Zlmsgstate
  (
    消息id_In In Zlmsgstate.消息id%Type,
    类型_In   In Zlmsgstate.类型%Type,
    用户_In   In Zlmsgstate.用户%Type
  ) Is
  Begin
    Update zlMsgState Set 删除 = 0 Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Restore_Zlmsgstate;

  --功能：保存消息
  --调用列表：zlApptools.frmMessageEdit.SaveMessage
  Procedure Save_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In Zlmessages.Id%Type,
    会话id_In  In Zlmessages.会话id%Type,
    发件人_In  In Zlmessages.发件人%Type,
    收件人_In  In Zlmessages.收件人%Type,
    主题_In    In Zlmessages.主题%Type,
    内容_In    In Zlmessages.内容%Type,
    背景色_In  In Zlmessages.背景色%Type
  ) Is
    n_Id     Zlmessages.Id%Type;
    n_会话id Zlmessages.会话id%Type;
  Begin
    If Nvl(Id_In, 0) = 0 Then
      Select Zlmessages_Id.Nextval Into n_Id From Dual;
      n_Id := Nvl(n_Id, 0);
      If Nvl(会话id_In, 0) = 0 Then
        n_会话id := n_Id;
      Else
        n_会话id := 会话id_In;
      End If;
      Insert Into zlMessages
        (ID, 会话id, 发件人, 时间, 收件人, 主题, 内容, 背景色)
      Values
        (n_Id, n_会话id, 发件人_In, Sysdate, 收件人_In, 主题_In, 内容_In, 背景色_In);
      Open Cursor_Out For
        Select n_Id As ID, n_会话id As 会话id From Dual;
    Else
      Update zlMessages
      Set 发件人 = 发件人_In, 时间 = Sysdate, 收件人 = 收件人_In, 主题 = 主题_In, 内容 = 内容_In, 背景色 = 背景色_In
      Where ID = Id_In;
      Open Cursor_Out For
        Select Id_In As ID, 会话id_In As 会话id From Dual;
    End If;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Save_Zlmessage;

  --功能：插入zlMsgstate
  --调用列表：zlApptools.frmMessageEdit.SaveMessage
  Procedure Insert_Zlmsgstate
  (
    消息id_In In Zlmsgstate.消息id%Type,
    类型_In   In Zlmsgstate.类型%Type,
    用户_In   In Zlmsgstate.用户%Type,
    身份_In   In Zlmsgstate.身份%Type,
    删除_In   In Zlmsgstate.删除%Type,
    状态_In   In Zlmsgstate.状态%Type
  ) Is
  Begin

    If 类型_In < 2 Then
      Delete From zlMsgState Where 消息id = 消息id_In;
    End If;
    Insert Into zlMsgState
      (消息id, 类型, 用户, 身份, 删除, 状态)
    Values
      (消息id_In, 类型_In, 用户_In, 身份_In, 删除_In, 状态_In);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Insert_Zlmsgstate;

  --功能：为原件加上答复或转发标志
  Procedure Update_Zlmsgstate_State
  (
    模式_In   In Number,
    消息id_In In Zlmsgstate.消息id%Type,
    类型_In   In Zlmsgstate.类型%Type,
    用户_In   In Zlmsgstate.用户%Type
  ) Is
  Begin
    If Nvl(模式_In, 0) = 1 Or Nvl(模式_In, 0) = 2 Then
      Update zlMsgState
      Set 状态 = Substr(状态, 1, 1) || '1' || Substr(状态, 3, 2)
      Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
      Commit;
    End If;
    If Nvl(模式_In, 0) = 3 Then
      Update zlMsgState
      Set 状态 = Substr(状态, 1, 1) || '1' || Substr(状态, 4, 1)
      Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
      Commit;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Update_Zlmsgstate_State;

  --功能：更新状态和身份
  Procedure Update_Zlmsgstate_Idtntify
  (
    身份_In   In Zlmsgstate.身份%Type,
    消息id_In In Zlmsgstate.消息id%Type,
    类型_In   In Zlmsgstate.类型%Type,
    用户_In   In Zlmsgstate.用户%Type
  ) Is
  Begin
    Update zlMsgState
    Set 状态 = '1' || Substr(状态, 2), 身份 = 身份_In
    Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Update_Zlmsgstate_Idtntify;

End b_Comfunc;
/
--113538:余智勇,2017-10-30,列属性表增加对齐字段
Alter Table Zltools.Zlrptcolproterty Add 对齐 Number(1);

--115010:高腾,2017-11-7,单系统登录问题创建角色表
Create Table zltools.Zlroles(
       名称   Varchar2(50),
       系统   Number(5));
ALTER TABLE zltools.Zlroles ADD CONSTRAINT ZLRoles_PK PRIMARY KEY (名称) USING INDEX;
Insert Into zlTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLROLES','ZLTOOLSTBS','A2');
--115010:高腾,2017-11-7,单系统登录问题增加角色授权表约束
Alter Table zltools.zlRoleGroups Modify 组名 Constraint zlRoleGroups_NN_组名 Not Null;
--115010:高腾,2017-11-7,单系统登录问题初始化角色数据
Truncate Table Zltools.Zlroles;
Insert Into zltools.Zlroles(名称) Select Role From Sys.Dba_Roles Where Upper(Role) Like 'ZL_%';
Delete zltools.Zlrolegroups Where 组名 = '未分组';
--115010:高腾,2017-11-7,单系统登录问题
CREATE OR REPLACE Procedure zltools.Zl_Zlroles_Edit
(
  操作_In In Number, --1-增加，2-修改，3－删除
  名称_In In Zlroles.名称%Type,
  系统_In In Zlroles.系统%Type := Null
) Is
Begin
  If 操作_In = 1 Then
    Insert Into Zlroles Values (名称_In, 系统_In);
  Elsif 操作_In = 2 Then
    Update Zlroles Set 系统 = 系统_In Where 名称 = 名称_In;
  Else
    Delete Zlroles Where 名称 = 名称_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlroles_Edit;
/
--115010:高腾,2017-11-7,单系统登录问题
Create Or Replace Procedure Zltools.Zl_Checkrolesdiff As
  --功能：检查Dba_Roles表中的数据和Zlroles中的是否一致，若不一致，则将Dba_Roles中
  --     多出来的数据插入到zlroles中
Begin
  --若Dba_Roles中存在这个角色，但Zlroles里面不存在，则将该角色添加到Zlroles中去
  Insert Into Zlroles
    (名称)
    Select a.Role
    From Dba_Roles a
    Where Not Exists (Select 1 From Zlroles b Where a.Role = b.名称) And a.Role Like 'ZL_%';
  --若Dba_Roles中不存在该角色，但Zlroles中存在，则将该角色从Zlroles中删除
  Delete Zlroles a Where Not Exists (Select 1 From Dba_Roles b Where a.名称 = b.Role);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Checkrolesdiff;
/

--116688:杨周一,2017-11-08,管理工具登录控制
Insert Into zlTools.zlSvrTools(编号,上级,标题,快键,说明,次序) Values('0607','06','用户与IP限制','I',Null,7);
Insert Into zlTools.zlSvrTools(编号,上级,标题,快键,说明,次序) Values('0608','06','应用程序授权','A',Null,8);
Insert Into zlTools.zlSvrTools(编号,上级,标题,快键,说明,次序) Values('0609','06','用户登录日志','L',Null,9);


--116688:杨周一,2017-11-08,用户登录限制
Create Table zlTools.zlLoginLimit(
    ID  Number(18),
    用户名 Varchar2(50),
    开始IP Varchar2(20),
    结束IP Varchar2(20),
    开始时间 Date,
    结束时间 Date,
    状态  Number(1),
    说明  Varchar2(200));
CREATE Sequence zlTools.zlLoginLimit_ID start with 1;
ALTER TABLE zlTools.zlLoginLimit ADD CONSTRAINT zlLoginLimit_PK PRIMARY KEY (ID) USING INDEX;
Alter Table zlTools.zlLoginLimit Add Constraint zlLoginLimit_Uq_用户名 Unique(用户名,开始IP,结束IP,开始时间,结束时间) Using Index;
Insert into zlTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLLOGINLIMIT','ZLTOOLSTBS','C1');

--116688:杨周一,2017-11-08,限制规则新增/修改
Create Or Replace Procedure Zltools.Zl_Zlloginlimit_Edit
(
  操作类型_In In Number,
  Id_In       In Zlloginlimit.Id%Type,
  开始时间_In In Zlloginlimit.开始时间%Type,
  结束时间_In In Zlloginlimit.结束时间%Type,
  开始ip_In   In Zlloginlimit.开始ip%Type,
  结束ip_In   In Zlloginlimit.结束ip%Type,
  状态_In     In Zlloginlimit.状态%Type,
  说明_In     In Zlloginlimit.说明%Type,
  用户名_In   In Zlloginlimit.用户名%Type
) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  -------------------------------------------------------------------------------------
  --功能:登录限制规则添加/修改
  --参数:
  --操作类型_In=1:添加 ,不传入ID_In , 操作类型_In-其他值=修改,需传入ID_In
  --用户名_In-一条规则可以对多个用户生效,每个用户名之间用","隔开
  -------------------------------------------------------------------------------------
Begin
  If 操作类型_In = 1 Or Id_In Is Null Then
    If 用户名_In Is Null Then
      Insert Into Zlloginlimit
        (ID, 开始时间, 结束时间, 开始ip, 结束ip, 状态, 说明, 用户名)
      Values
        (Zlloginlimit_Id.Nextval, 开始时间_In, 结束时间_In, 开始ip_In, 结束ip_In, 状态_In, 说明_In, 用户名_In);
    Else
      --操作类型为1,或者Id_为空,添加数据
      Insert Into Zlloginlimit
        (ID, 开始时间, 结束时间, 开始ip, 结束ip, 状态, 说明, 用户名)
        Select Zlloginlimit_Id.Nextval, 开始时间_In, 结束时间_In, 开始ip_In, 结束ip_In, 状态_In, 说明_In, Column_Value
        From Table(f_Str2list(用户名_In)) A;
    End If;
  Else
    --修改数据
    Update Zlloginlimit
    Set 开始时间 = 开始时间_In, 结束时间 = 结束时间_In, 开始ip = 开始ip_In, 结束ip = 结束ip_In, 状态 = 状态_In, 说明 = 说明_In, 用户名 = 用户名_In
    Where ID = Id_In;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlloginlimit_Edit;
/

--116688:杨周一,2017-11-08,登录限制规则删除
Create Or Replace Procedure Zltools.Zl_Zlloginlimit_Delete(Ids_In In Varchar2) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  -------------------------------------------------------------------------------------
  --功能:登录限制规则删除
  --参数:Ids_In-传入字符串类型,便于批量删除,每个ID之间用","作为间隔
  -------------------------------------------------------------------------------------
Begin
  Delete Zlloginlimit Where ID In (Select Column_Value From Table(f_Str2list(Ids_In)) A);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlloginlimit_Delete;
/

--116688:杨周一,2017-11-08,应用程序授权

Create Table zlTools.zlAppPermission(
    应用程序 Varchar2(100),
    用户名 Varchar2(50),
    开始IP Varchar2(20),
    结束IP Varchar2(20),
    状态 Number(1),
    说明 Varchar2(200));
ALTER TABLE zlTools.zlAppPermission ADD CONSTRAINT zlAppPermission_PK PRIMARY KEY (应用程序,用户名) USING INDEX;
Insert into zlTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLAPPPERMISSION','ZLTOOLSTBS','C1');

--116688:杨周一,2017-11-08,应用程序授权新增/修改
Create Or Replace Procedure Zltools.Zl_Zlapppermission_Edit
(
  操作类型_In    In Number,
  应用程序_In    In Zlapppermission.应用程序%Type,
  用户名_In      In Varchar,
  开始ip_In      In Zlapppermission.开始ip%Type,
  结束ip_In      In Zlapppermission.结束ip%Type,
  状态_In        In Zlapppermission.状态%Type,
  说明_In        In Zlapppermission.说明%Type,
  应用程序new_In In Zlapppermission.应用程序%Type,
  用户名new_In   In Varchar
) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  -------------------------------------------------------------------------------------
  --功能:登录限制规则添加/修改
  --参数:
  --操作类型_In=1:添加 ,不传入ID_In , 操作类型_In-其他值=修改,需传入ID_In
  --用户名_In-一条规则可以对多个用户生效,每个用户名之间用","隔开
  -------------------------------------------------------------------------------------
Begin
  If 操作类型_In = 1 Then
    Insert Into Zlapppermission
      (应用程序, 用户名, 开始ip, 结束ip, 状态, 说明)
      Select 应用程序_In, Column_Value, 开始ip_In, 结束ip_In, 状态_In, 说明_In From Table(f_Str2list(用户名_In)) A;
  Else
    Update Zlapppermission
    Set 开始ip = 开始ip_In, 结束ip = 结束ip_In, 状态 = 状态_In, 说明 = 说明_In, 应用程序 = 应用程序new_In, 用户名 = 用户名new_In
    Where 应用程序 = 应用程序_In And 用户名 = 用户名_In;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlapppermission_Edit;
/

--116688:杨周一,2017-11-08,应用程序授权删除
Create Or Replace Procedure Zltools.Zl_Zlapppermission_Delete(Ids_In In Varchar2) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  -------------------------------------------------------------------------------------
  --功能:登录限制规则删除
  --参数:Ids_In-传入字符串类型,便于批量删除,格式为 应用程序1:用户名1,应用程序2:用户名2
  -- 如 SQLPLUS:ZLHIS,plsql DEV:ZLHIS
  -------------------------------------------------------------------------------------
Begin
  Delete From Zlapppermission Where (应用程序, 用户名) In (Select C1, C2 From Table(f_Str2list2(Ids_In)) A);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlapppermission_Delete;
/

--116688:杨周一,2017-11-08,用户登录触发器
Create Or Replace Trigger zlTools.zl_Trigger_LoginLimit
        After logon On Database
Declare
  --IP_1:ip的前三位, ip_2:IP地址的后一位
  v_Ip     Varchar2(25);
  v_Ip_1   Varchar2(20);
  v_Ip_2   Varchar2(5);
  v_User   Varchar2(40);
  v_Date   Date;
  v_Module Varchar2(100);
  n_Count  Number(5);
Begin

  --如果没有限制规则,就不执行操作,防止多余耗时
  Select Count(1)
  Into n_Count
  From (Select 1
         From Zlapppermission
         Where Rownum = 1
         Union All
         Select 1 From Zlloginlimit Where Rownum = 1);

  If n_Count <> 0 Then
    Select Sys_Context('userenv', 'ip_address'), User, Sysdate, Module
    Into v_Ip, v_User, v_Date, v_Module
    From V$session
    Where Audsid = Userenv('sessionid') And Rownum = 1;
    v_Ip_1 := Substr(v_Ip, 1, Instr(v_Ip, '.', 1, 3) - 1);
    v_Ip_2 := Substr(v_Ip, Instr(v_Ip, '.', 1, 3) + 1);

    --检查登录规则
    Select Count(1)
    Into n_Count
    From Zlloginlimit
    Where 状态 = 1 And
          ((用户名 = User And Substr(开始ip, 1, Instr(开始ip, '.', 1, 3) - 1) = v_Ip_1 And
          v_Ip_2 Between Substr(开始ip, Instr(开始ip, '.', 1, 3) + 1) And Substr(结束ip, Instr(结束ip, '.', 1, 3) + 1) And
          v_Date Between 开始时间 And 结束时间) Or (用户名 = User And 开始ip Is Null And 开始时间 Is Null) Or
          (Substr(开始ip, 1, Instr(开始ip, '.', 1, 3) - 1) = v_Ip_1 And
          v_Ip_2 Between Substr(开始ip, Instr(开始ip, '.', 1, 3) + 1) And Substr(结束ip, Instr(结束ip, '.', 1, 3) + 1) And
          用户名 Is Null And 开始时间 Is Null) Or (v_Date Between 开始时间 And 结束时间 And 用户名 Is Null And 开始ip Is Null) Or
          (用户名 = User And Substr(开始ip, 1, Instr(开始ip, '.', 1, 3) - 1) = v_Ip_1 And
          v_Ip_2 Between Substr(开始ip, Instr(开始ip, '.', 1, 3) + 1) And Substr(结束ip, Instr(结束ip, '.', 1, 3) + 1) And
          开始时间 Is Null) Or (用户名 = User And v_Date Between 开始时间 And 结束时间 And 开始ip Is Null) Or
          (Substr(开始ip, 1, Instr(开始ip, '.', 1, 3) - 1) = v_Ip_1 And
          v_Ip_2 Between Substr(开始ip, Instr(开始ip, '.', 1, 3) + 1) And Substr(结束ip, Instr(结束ip, '.', 1, 3) + 1) And
          v_Date Between 开始时间 And 结束时间 And 用户名 Is Null));

    If n_Count > 0 Then
      Raise_Application_Error(-20001, '当前用户被禁止登录数据库，请联系管理员。');
    End If;

    --检查应用授权
    Select Count(1) Into n_Count From Zlapppermission Where 应用程序 = v_Module And 状态 = 1;

    If n_Count > 0 Then
      Select Count(1)
      Into n_Count
      From Zlapppermission
      Where 状态 = 1 And
            ((应用程序 = v_Module And 用户名 = v_User And 开始ip Is Null) Or
            (应用程序 = v_Module And 用户名 = v_User And Substr(开始ip, 1, Instr(开始ip, '.', 1, 3) - 1) = v_Ip_1 And
            v_Ip_2 Between Substr(开始ip, Instr(开始ip, '.', 1, 3) + 1) And Substr(结束ip, Instr(结束ip, '.', 1, 3) + 1)));
      If n_Count = 0 Then
        Raise_Application_Error(-20002, '当前用户不能使用该应用登录数据库，请联系管理员。');
      End If;
    End If;
  End If;
End Zl_Trigger_Loginlimit;
/
--116691:杨周一,2017-11-15,远程控制参数
Insert Into zlParameters(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
    Select zlParameters_ID.Nextval,-Null,-null,-Null,1,-Null,-Null,A.* From (
    Select 部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 From zlParameters Where 1 = 0 Union All
    Select 0,0,31,'允许远程控制',Null,'1001','进入导航台后,是否允许被管理员远程控制','远程控制',Null,'用于允许管理员通过远程控制来进行问题查证或操作演示',Null From dual Union All
    Select 部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 From ZLPARAMETERS Where 1 = 0) A;

--113395:高腾,2017-11-15,管理工具按功能授权
Create Table Zltools.ZLSVRFuncs(
       序号 varchar2(6),
       功能 varchar2(30),
       排列 number(3),
       说明 varchar2(250),
       缺省 number(1));
Alter Table Zltools.ZLSVRFuncs Add Constraint ZLSVRFuncs_UQ_功能 Unique(序号, 功能) Using Index;
Alter Table Zltools.ZLSVRFuncs Modify 序号 Constraint ZLSVRFuncs_NN_序号 Not Null;
Alter Table Zltools.ZLSVRFuncs Add Constraint ZLSVRFuncs_FK_序号 Foreign Key(序号) References Zlsvrtools(编号) On Delete Cascade;
Alter Table Zltools.ZLMgrGrant Modify(功能 Varchar2(4000));
Insert Into zlTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLSVRFUNCS','ZLTOOLSTBS','A2');
Insert Into Zltools.Zlsvrfuncs(序号, 功能, 排列, 说明, 缺省)
  Select 编号, '基本', '0', Null, 1 From Zlsvrtools Where 上级 Is Not Null;
Insert Into Zltools.Zlsvrfuncs(序号, 功能, 排列, 说明, 缺省) Values ('0307', '文件服务器配置', 1, '用于对客户端升级所需的服务器进行配置，包括服务器的启停，服务器类型等', 1);
Insert Into Zltools.Zlsvrfuncs(序号, 功能, 排列, 说明, 缺省) Values ('0307', '文件升级管理', 2, '用于对客户端升级所用的升级文件进行管理，主要为升级文件的增删改', 1);
Insert Into Zltools.Zlsvrfuncs(序号, 功能, 排列, 说明, 缺省) Values ('0307', '客户端升级配置', 3, '用于配置客户端升级信息，包括是否启用升级，定时升级等', 1);

--000000:杨周一,2017-11-30,公共函数添加对序列的检查
Create Or Replace Function Zltools.Zl_Checkobject
(
  n_Type        In Number, --1=表,2=字段,3=约束,4=索引 ,5=序列
  v_Object_Name In Varchar2,
  v_Table_Name  In Varchar2 := Null --仅当n_Type=2时才需要传入 
) Return Number Authid Current_User As
  --功能：以执行者的身份检查指定表的指定对象是否存在 
  --返回值：>0表示存在，0表示不存在 
  n_Count Number(5);
Begin
  If n_Type = 1 Then
    If v_Table_Name Is Null Then
      Select Count(1) Into n_Count From User_Tables Where Table_Name = Upper(v_Object_Name);
    Else
      Select Count(1) Into n_Count From User_Tables Where Table_Name = Upper(v_Table_Name);
    End If;
  Elsif n_Type = 2 Then
    Select Count(1)
    Into n_Count
    From User_Tab_Columns
    Where Table_Name = Upper(v_Table_Name) And Column_Name = Upper(v_Object_Name);
  
  Elsif n_Type = 3 Then
    Select Count(1) Into n_Count From User_Constraints Where Constraint_Name = Upper(v_Object_Name);
  
  Elsif n_Type = 4 Then
    Select Count(1) Into n_Count From User_Indexes Where Index_Name = Upper(v_Object_Name);
  
  Elsif n_Type = 5 Then
    Select Count(1) Into n_Count From User_Sequences Where Sequence_Name = Upper(v_Object_Name);
  End If;

  Return n_Count;
End Zl_Checkobject;
/