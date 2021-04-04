--10.35.0

--00000:张永康,2015-03-13,参数重整(集中处理，未登记BUG)
--84990:刘硕,2015-06-12,参数重整删除参数上传下载等功能
Delete zlPrograms where 系统 Is Null And 序号=15;
--00000:刘硕,2015-08-20,模块关联授权误加字段
alter table zltools.zlprograms drop column 性质;
--00000:刘硕,2015-05-04,参数重整(集中处理，未登记BUG)
alter table Zltools.zlParameters add 部门 NUMBER(1);
alter table Zltools.zlParameters add 性质 NUMBER(1);
Alter Table Zltools.zlParameters Add Constraint zlParameters_CK_性质 Check (性质 IN(0,1));
Alter Table Zltools.zlParameters Add Constraint zlParameters_CK_部门 Check (部门 IN(0,1));

alter table Zltools.zlParameters Add 影响控制说明 varchar2(2000);
alter table Zltools.zlParameters add 参数值含义 varchar2(2000);
alter table Zltools.zlParameters add 关联说明 varchar2(2000);
alter table Zltools.zlParameters add 适用说明 varchar2(2000);
alter table Zltools.zlParameters add 警告说明 varchar2(2000);
alter table Zltools.Zlparachangedlog modify 变动内容 varchar2(4000);

--89346:刘硕,2015-10-20,自动锁屏
Insert Into Zlparameters
  (Id, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, -null, -null, 1, -null, -null, -null, 0, 0, 25, '自动锁屏', '5', '5', '间隔指定分钟数自动锁定系统',
         '0或NUll，不进行自动锁定，>0：自动锁定间隔的分钟数', Null, Null, Null
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '自动锁屏');

--84990:刘硕,2015-06-13,参数重整
Update zlParameters Set  影响控制说明 = '记录自动消息提醒信息停留时间(秒)' , 参数值含义 = '单位：秒'  Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '自动消息停留时间';
Update zlParameters Set  影响控制说明 = '设置在邮件管理器中是否显示已读邮件' Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '显示已读邮件';
Update zlParameters Set  影响控制说明 = '记录最近使用的产品模块，用于在导航台历史菜单中显示' Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '最近使用模块';
Update zlParameters Set  影响控制说明 = '设置是否记忆当前用户的界面特性以便下次进入时保持以前的设置,包括窗口的位置、宽高，表格的列宽、顺序等。' Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '使用个性化风格';
Update zlParameters Set  影响控制说明 = '设置是否接收邮件消息通知' Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '接收邮件消息';
Update zlParameters Set  性质 = 1 , 影响控制说明 = '记录Brower风格导航台字体大小，由小到大分别为：0-9号,1-11号,2-12号' Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = 'zlBrwFontSize';
Update zlParameters Set  性质 = 1 , 影响控制说明 = '设置MDI风格导航台的字体颜色'  Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = 'zlMdiFontColor';
Update zlParameters Set  性质 = 1 , 影响控制说明 = '设置MDI风格导航台的背景图片文件路径' Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = 'zlMdiBackPic';
Update zlParameters Set  性质 = 1 , 影响控制说明 = '设置MDI风格导航台菜单排列方式' , 参数值含义 = '0-纵向排列，1-横向排列' Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = 'zlMdiMenuArray';
Update zlParameters Set  性质 = 1 , 影响控制说明 = '设置Windows风格导航台的字体颜色'  Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = 'zlWinFontColor';
Update zlParameters Set  性质 = 1 , 影响控制说明 = '设置Windows风格导航台的背景图片文件路径'  Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = 'zlWinBackPic';
Update zlParameters Set  性质 = 1 , 影响控制说明 = '记录使用哪种类型的导航台：zlBrw，zlWin，zlMdi' , 参数值含义 = '导航台名称'  Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '导航台';
Update zlParameters Set  影响控制说明 = '设置是否允许界面区域提供自动隐藏功能'  Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '界面区域隐藏';
Update zlParameters Set  影响控制说明 = '门诊和住院医生站，以及药房处方发药和部门发药等界面，药品名称显示（主界面单据明细、单据输入界面、直接进入的药品选择器时的药品名称显示）' , 参数值含义 = '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名' , 适用说明 = '临床科室人员习惯看药品通用名，药房人员习惯看药品商品名' Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '药品名称显示';
Update zlParameters Set  影响控制说明 = '门诊和住院医生站，门诊收费和住院记帐等费用相关界面，输入药品时以哪种方式显示（通过输入简码方式进入选择器时药品名称的显示）' , 参数值含义 = '0-按输入匹配显示，1-固定显示通用名和商品名'  Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '输入药品显示';
Update zlParameters Set  影响控制说明 = '允许在窗口界面的工具栏切换简码匹配方式，不允许时工具栏不显示切换按钮。' Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '简码匹配方式切换';
Update zlParameters Set  影响控制说明 = '允许在网络断网或者多重网络切换后自动重新连接数据库' , 参数值含义 = '0-不检测，1-检测' Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '网络断网自动重连';
Update zlParameters Set  影响控制说明 = '设置显示在导航台工具栏上的常用功能模块' , 参数值含义 = '模块1所属系统,模块2所属系统|模块1编号,模块2编号|模块1图标,模块2图标|模块1名称,模块2名称' , 本机 = 0 , 授权 = 0 , 固定 = 0 Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '常用功能模块';
Update zlParameters Set  影响控制说明 = '设置各种业务操作中查找输入时的匹配方向' , 参数值含义 = '0-双向匹配，1-从左匹配' Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '输入匹配';
Update zlParameters Set  影响控制说明 = '设置各种业务操作中,中文类的文本输入框，自动开启的输入法名称' , 参数值含义 = '输入法名称'  Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '输入法';
Update zlParameters Set  影响控制说明 = '设置各种业务操作中,查找输入时的简码匹配方式' , 参数值含义 = '0-拼音，1-五笔'  Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '简码方式';
Update zlParameters Set  影响控制说明 = '设置是否退出程序时自动关闭 Windows'  Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '关闭Windows';
Update zlParameters Set  影响控制说明 = '设置自动检查邮件消息的时间间隔(秒)' , 参数值含义 = '单位：秒' Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '邮件消息检查周期';
Update zlParameters Set  影响控制说明 = '设置登录时是否检查新的邮件消息' Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '登录检查邮件消息';

Create Table Zltools.zlDeptParas(
    参数ID NUMBER(18),
    部门ID NUMBER(18),
    参数值 VARCHAR2(2000))
    PCTFREE 5
    Cache Storage(Buffer_Pool Keep);
Alter Table Zltools.zlDeptParas Add Constraint zlDeptParas_UQ_参数ID Unique(参数ID,部门ID) Using Index PCTFREE 5;
Alter Table Zltools.zlDeptParas Add Constraint zlDeptParas_FK_参数ID Foreign Key (参数ID) References zlParameters(ID) On Delete Cascade;
--79998:刘硕,2015-08-18,密码复杂性控制
Insert Into zlOptions(参数号,参数名,参数值,缺省值,参数说明) Values(20, '是否控制密码长度', '','', '是否启用密码长度控制');
Insert Into zlOptions(参数号,参数名,参数值,缺省值,参数说明) Values(21, '密码长度下限', '','3', '设置密码的最小长度');
Insert Into zlOptions(参数号,参数名,参数值,缺省值,参数说明) Values(22, '密码长度上限', '','12', '设置密码的最大长度');
Insert Into zlOptions(参数号,参数名,参数值,缺省值,参数说明) Values(23, '是否控制密码复杂度', '','', '是否控制密码必须包含至少一个字母、数字、与特殊字符。部分特殊字符不能当作密码输入。');

--00000:张永康,2015-03-20,参数重整(集中处理，未登记BUG)
Create Or Replace Procedure Zltools.Zl_Parameters_Change_Value
(
  参数id_In     Zlparachangedlog.参数id%Type,
  变动内容_In   Zlparachangedlog.变动内容%Type, --原值-->新值
  变动原因_In   Zlparachangedlog.变动原因%Type,
  操作员姓名_In Zlparachangedlog.变动人%Type,
  变动时间_In   Zlparachangedlog.变动时间%Type
) Is
  n_Max序号 Zlparachangedlog.序号%Type;
Begin
  Select Nvl(Max(序号), 1)+1 Into n_Max序号 From Zlparachangedlog Where 参数id = 参数id_In;

  Insert Into Zlparachangedlog
    (参数id, 序号, 变动说明, 变动内容, 变动人, 变动时间, 变动原因)
  Values
    (参数id_In, n_Max序号, '值变动', 变动内容_In, 操作员姓名_In, 变动时间_In, 变动原因_In);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Parameters_Change_Value;
/

CREATE OR REPLACE Procedure Zltools.Zl_Parameters_Update_Batch
(
  系统编号_In   Zlsystems.编号%Type,
  参数列表_In   Varchar2, --模块号1^参数号1^参数值1#模块号2^参数号2^参数值2......
  操作员姓名_In Zlparachangedlog.变动人%Type
) Is
  t_模块   t_Numlist;
  t_参数号 t_Numlist;
  t_参数值 t_Strlist;
Begin
  Select To_Number(C1), To_Number(Substr(C2, 1, Instr(C2, '^') - 1)), Substr(C2, Instr(C2, '^') + 1) Bulk Collect
  Into t_模块, t_参数号, t_参数值
  From Table(f_Str2list2(参数列表_In, '#', '^'));

  For Rs In (Select /*+ rule*/
              a.Id, a.参数值 || '-->' || Substr(C2, Instr(C2, '^') + 1) As 变动内容, Sysdate As 变动时间
             From zlParameters A, Table(f_Str2list2(参数列表_In, '#', '^')) B
             Where a.系统 = 系统编号_In And Nvl(a.模块, 0) = To_Number(b.C1) And
                   a.参数号 = To_Number(Substr(b.C2, 1, Instr(b.C2, '^') - 1)) And a.警告说明 Is Null) Loop
    --有警告说明的关键参数，单独在界面提供变动登记
    Zl_Parameters_Change_Value(Rs.Id, Rs.变动内容, '', 操作员姓名_In, Rs.变动时间);
  End Loop;

  Forall I In 1 .. t_参数号.Count
    Update zlParameters
    Set 参数值 = t_参数值(I)
    Where 系统 = 系统编号_In And Nvl(模块, 0) = t_模块(I) And 参数号 = t_参数号(I);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Parameters_Update_Batch;
/
--83832:梁唐彬,2015-04-07,参数整理
CREATE OR REPLACE Procedure Zltools.Zl_DeptParameters_Delete
(
  参数_In   Zlparameters.参数名%Type,
  系统_In   Zlparameters.系统%Type,
  模块_In   Zlparameters.模块%Type
  --功能：删除部门类型的对应部门的所有参数
  --参数：
  --     参数_In：必须传入非Null值，以字符形式传入的参数号或参数名,注意参数名不能为数字。
  --     权限_IN：对于要求用权限控制的参数，当前用户是否有权限设置
) Is
  v_参数id Zlparameters.Id%Type;
  v_私有   Zlparameters.私有%Type;
  v_本机   Zlparameters.本机%Type;
  v_授权   Zlparameters.授权%Type;
  v_机器名 Zluserparas.机器名%Type;
  v_部门   Zlparameters.部门%Type;
Begin
  --确定参数信息
  Begin
    If Zl_To_Number(参数_In) <> 0 Then
      --以参数号为准处理
      Select ID, 私有, 本机, 授权, Sys_Context('USERENV', 'TERMINAL'), 部门
      Into v_参数id, v_私有, v_本机, v_授权, v_机器名, v_部门
      From zlParameters
      Where Nvl(系统, 0) = Nvl(系统_In, 0) And Nvl(模块, 0) = Nvl(模块_In, 0) And 参数号 = Zl_To_Number(参数_In);
    Else
      --以参数名为准处理
      Select ID, 私有, 本机, 授权, Sys_Context('USERENV', 'TERMINAL'), 部门
      Into v_参数id, v_私有, v_本机, v_授权, v_机器名, v_部门
      From zlParameters
      Where Nvl(系统, 0) = Nvl(系统_In, 0) And Nvl(模块, 0) = Nvl(模块_In, 0) And 参数名 = 参数_In;
    End If;
  Exception
    When Others Then
      Return;
  End;

  If Nvl(v_部门, 0) = 0 Then
    Return; --部门级模块参数
  End If;

  --更新参数值
  If v_参数id Is Not Null Then
     Delete From zldeptparas Where 参数id = v_参数id;
  End If;
End Zl_DeptParameters_Delete;
/

--83832:梁唐彬,2015-04-07,参数整理
Create Or Replace Procedure Zltools.Zl_Parameters_Update
(
  参数_In   Zlparameters.参数名%Type,
  参数值_In Zlparameters.参数值%Type,
  系统_In   Zlparameters.系统%Type,
  模块_In   Zlparameters.模块%Type,
  权限_In   Number := 1,
  部门id_In zldeptparas.部门id%Type := 0
  --功能：设置系统参数值，如果是用户私有参数，则用户名以当前的为准
  --参数：
  --     参数_In：必须传入非Null值，以字符形式传入的参数号或参数名,注意参数名不能为数字。
  --     权限_IN：对于要求用权限控制的参数，当前用户是否有权限设置
) Is
  v_参数id Zlparameters.Id%Type;
  v_私有   Zlparameters.私有%Type;
  v_本机   Zlparameters.本机%Type;
  v_授权   Zlparameters.授权%Type;
  v_机器名 Zluserparas.机器名%Type;
  v_部门   Zlparameters.部门%Type;
Begin
  --确定参数信息
  Begin
    If Zl_To_Number(参数_In) <> 0 Then
      --以参数号为准处理
      Select ID, 私有, 本机, 授权, Sys_Context('USERENV', 'TERMINAL'), 部门
      Into v_参数id, v_私有, v_本机, v_授权, v_机器名, v_部门
      From zlParameters
      Where Nvl(系统, 0) = Nvl(系统_In, 0) And Nvl(模块, 0) = Nvl(模块_In, 0) And 参数号 = Zl_To_Number(参数_In);
    Else
      --以参数名为准处理
      Select ID, 私有, 本机, 授权, Sys_Context('USERENV', 'TERMINAL'), 部门
      Into v_参数id, v_私有, v_本机, v_授权, v_机器名, v_部门
      From zlParameters
      Where Nvl(系统, 0) = Nvl(系统_In, 0) And Nvl(模块, 0) = Nvl(模块_In, 0) And 参数名 = 参数_In;
    End If;
  Exception
    When Others Then
      Return;
  End;

  --检查权限
  If Nvl(权限_In, 0) = 0 Then
    If Nvl(v_部门, 0) <> 0 Then
      Return; --部门级模块参数
    Elsif Nvl(系统_In, 0) <> 0 And Nvl(模块_In, 0) = 0 And Nvl(v_私有, 0) = 0 And Nvl(v_本机, 0) = 0 Then
      Return; --公共全局参数,固定需要权限
    Elsif Nvl(系统_In, 0) <> 0 And Nvl(模块_In, 0) <> 0 And Nvl(v_私有, 0) = 0 And Nvl(v_本机, 0) = 0 Then
      Return; --公共模块参数,固定需要权限
    Elsif Nvl(系统_In, 0) <> 0 And Nvl(模块_In, 0) <> 0 And Nvl(v_私有, 0) = 0 And Nvl(v_本机, 0) = 1 And Nvl(v_授权, 0) = 1 Then
      Return; --要授权控制的本机公共模块
    End If;
  End If;

  --更新参数值
  If v_参数id Is Not Null Then
    If Nvl(v_部门, 0) <> 0 Then
      Update zldeptparas Set 参数值 = 参数值_In Where 参数id = v_参数id And 部门ID= 部门id_In;
      If Sql%RowCount = 0 Then
        Insert Into zldeptparas
          (参数id, 部门ID, 参数值)
        Values
          (v_参数id,部门id_In , 参数值_In);
      End If;
    elsIf Nvl(v_私有, 0) = 0 And Nvl(v_本机, 0) = 0 Then
      Update zlParameters Set 参数值 = 参数值_In Where ID = v_参数id;
    Else
      Update zlUserParas
      Set 参数值 = 参数值_In
      Where 参数id = v_参数id And Nvl(用户名, 'NullUser') = Decode(v_私有, 1, User, 'NullUser') And
            Nvl(机器名, 'NullMachine') = Decode(v_本机, 1, v_机器名, 'NullMachine');
      If Sql%RowCount = 0 Then
        Insert Into zlUserParas
          (参数id, 用户名, 机器名, 参数值)
        Values
          (v_参数id, Decode(v_私有, 1, User, Null), Decode(v_本机, 1, v_机器名, Null), 参数值_In);
      End If;
    End If;
  End If;
End Zl_Parameters_Update;
/
--84990:刘硕,2015-09-22,参数整理
Create Or Replace Procedure Zltools.Zlparameters_Delall_Details
(
  参数列表_In Varchar2,
  n_部门      Number := 0
  --n_部门:1，部门类型的参数。0-非部门类型的参数
  --参数列表_In 系统1^模块1^参数名1#系统2......, 
) Is
  t_参数id t_Numlist;
Begin
  Select a.Id Bulk Collect
  Into t_参数id
  From Zlparameters a,
       (Select Zl_To_Number(C1) 系统, Zl_To_Number(Substr(C2, 1, Instr(C2, '^') - 1)) 模块,
                Substr(C2, Instr(C2, '^') + 1) 参数名
         From Table(f_Str2list2(参数列表_In, '#', '^'))) b
  Where Nvl(a.系统, 0) = Nvl(b.系统, 0) And Nvl(a.模块, 0) = Nvl(b.模块, 0) And a.参数名 = b.参数名;
  If n_部门 = 0 Then
    Forall i In 1 .. t_参数id.Count
      Delete Zluserparas Where 参数id = t_参数id(i);
  Else
    Forall i In 1 .. t_参数id.Count
      Delete Zldeptparas Where 参数id = t_参数id(i);
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zlparameters_Delall_Details;
/
--84990:刘硕,2015-09-22,参数整理
CREATE OR REPLACE Procedure ZLTOOLS.Zlparameters_Add_Details
(
  参数id_In Zlparameters.Id%Type,
  用户名_In Varchar2,
  机器名_In Varchar2,
  参数值_In Varchar2
  --用户名_In 以逗号分割，用户1,用户2,
  --机器名_In 以逗号分割，机器1,机器2,
) Is
  n_部门 Number(1);
Begin
  Select Nvl(部门, 0) Into n_部门 From Zlparameters Where Id = 参数id_In;
  If n_部门 = 0 Then
    Insert Into Zluserparas
      (参数id, 用户名, 机器名, 参数值)
      Select 参数id, 用户名, 机器名, 参数值
      From (Select 参数id_In 参数id, a.用户名, b.机器名, 参数值_In 参数值
             From (Select Distinct Column_Value 用户名 From Table(f_Str2list(Nvl(用户名_In, ',')))) a,
                  (Select Distinct Column_Value 机器名 From Table(f_Str2list(Nvl(机器名_In, ',')))) b) c
      Where Not Exists
       (Select 1
             From Zluserparas
             Where 参数id = c.参数id And Nvl(用户名, '空空') = Nvl(c.用户名, '空空') And Nvl(机器名, '空空') = Nvl(c.机器名, '空空'));
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zlparameters_Add_Details;
/
--84990:刘硕,2015-05-28,参数整理
Create Or Replace Procedure Zltools.Zlparameters_Update_Details
(
  参数id_In   Zlparameters.Id%Type,
  参数列表_In Varchar2
  --参数列表_In 用户名1^机器名1^参数值1#用户名2^机器名2^参数值2......,
  --           部门类型参数：部门ID1,,参数值1#部门ID2,,参数值2
) Is
  n_部门   Number(1);
  t_部门id t_Numlist;
  t_用户名 t_Strlist;
  t_机器名 t_Strlist;
  t_参数值 t_Strlist;
Begin
  Select Nvl(部门, 0) Into n_部门 From zlParameters Where ID = 参数id_In;
  If n_部门 = 0 Then
    Select C1, Substr(C2, 1, Instr(C2, '^') - 1), Substr(C2, Instr(C2, '^') + 1) Bulk Collect
    Into t_用户名, t_机器名, t_参数值
    From Table(f_Str2list2(参数列表_In, '#', '^'));
  
    Forall I In 1 .. t_参数值.Count
      Update zlUserParas
      Set 参数值 = t_参数值(I)
      Where 参数id = 参数id_In And Nvl(用户名, '空空') = Nvl(t_用户名(I), '空空') And Nvl(机器名, '空空') = Nvl(t_机器名(I), '空空');
  Else
    Select To_Number(C1), Substr(C2, 1, Instr(C2, '^') - 1), Substr(C2, Instr(C2, '^') + 1) Bulk Collect
    Into t_部门id, t_机器名, t_参数值
    From Table(f_Str2list2(参数列表_In, '#', '^'));
  
    Forall I In 1 .. t_参数值.Count
      Update Zldeptparas Set 参数值 = t_参数值(I) Where 参数id = 参数id_In And 部门id = t_部门id(I);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlparameters_Update_Details;
/
--84990:刘硕,2015-09-28,参数整理
Create Or Replace Procedure Zltools.Zlparameters_Del_Details
(
  参数id_In   Zlparameters.Id%Type,
  参数列表_In Varchar2
  --参数列表_In 用户名1^机器名1#用户名2^机器名2......,
  --           部门类型参数：部门ID1#部门ID2
) Is
  n_部门   Number(1);
  t_部门id t_Numlist;
  t_用户名 t_Strlist;
  t_机器名 t_Strlist;
Begin
  Select Nvl(部门, 0) Into n_部门 From Zlparameters Where Id = 参数id_In;
  If n_部门 = 0 Then
    Select C1, C2 Bulk Collect Into t_用户名, t_机器名 From Table(f_Str2list2(参数列表_In, '#', '^'));
  
    Forall i In 1 .. t_用户名.Count
      Delete Zluserparas
      Where 参数id = 参数id_In And Nvl(用户名, '空空') = Nvl(t_用户名(i), '空空') And Nvl(机器名, '空空') = Nvl(t_机器名(i), '空空');
  Else
    Select To_Number(Column_Value) Bulk Collect Into t_部门id From Table(f_Str2list(参数列表_In, '#'));
  
    Forall i In 1 .. t_部门id.Count
      Delete Zldeptparas Where 参数id = 参数id_In And 部门id = t_部门id(i);
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zlparameters_Del_Details;
/
--84990:刘硕,2015-05-28,参数整理
Create Or Replace Procedure Zltools.Zlparameters_Imp_Details
(
  系统_In     Zlparameters.系统%Type,
  模块_In     Zlparameters.模块%Type,
  参数_In     Zlparameters.参数名%Type,
  参数列表_In Varchar2
  --参数列表_In 用户名1^机器名1^参数值1#用户名2^机器名2^参数值2......,
  --           部门类型参数：部门ID1,,参数值1#部门ID2,,参数值2
  --参数列表为空时删除所有详细参数
) Is
  n_参数id Zlparameters.Id%Type;
  n_部门   Number(1);
  n_私有   Number(1);
  n_本机   Number(1);
  t_部门id t_Numlist;
  t_用户名 t_Strlist;
  t_机器名 t_Strlist;
  t_参数值 t_Strlist;
Begin
  --获取参数ID与部门性质
  If Zl_To_Number(参数_In) <> 0 Then
    Select Nvl(部门, 0), Nvl(私有, 0), Nvl(本机, 0), ID
    Into n_部门, n_私有, n_本机, n_参数id
    From zlParameters
    Where 参数号 = Zl_To_Number(参数_In) And Nvl(模块, 0) = Nvl(模块_In, 0) And Nvl(系统, 0) = Nvl(系统_In, 0);
  Else
    Select Nvl(部门, 0), Nvl(私有, 0), Nvl(本机, 0), ID
    Into n_部门, n_私有, n_本机, n_参数id
    From zlParameters
    Where 参数名 = 参数_In And Nvl(模块, 0) = Nvl(模块_In, 0) And Nvl(系统, 0) = Nvl(系统_In, 0);
  End If;
  If n_参数id Is Not Null Then
    If n_部门 = 0 Then
      If 参数列表_In Is Null Then
        Delete zlUserParas Where 参数id = n_参数id;
        --私有或本机参数，才插入
      Elsif n_私有 = 1 Or n_本机 = 1 Then
        Select C1, Substr(C2, 1, Instr(C2, '^') - 1), Substr(C2, Instr(C2, '^') + 1) Bulk Collect
        Into t_用户名, t_机器名, t_参数值
        From Table(f_Str2list2(参数列表_In, '#', '^'));
        --采用全部删除，再插入
        Forall I In 1 .. t_参数值.Count
          Insert Into zlUserParas
            (参数id, 用户名, 机器名, 参数值)
          Values
            (n_参数id, t_用户名(I), t_机器名(I), t_参数值(I));
      End If;
    Else
      If 参数列表_In Is Null Then
        Delete Zldeptparas Where 参数id = n_参数id;
      Else
        Select To_Number(C1), Substr(C2, 1, Instr(C2, '^') - 1), Substr(C2, Instr(C2, '^') + 1) Bulk Collect
        Into t_部门id, t_机器名, t_参数值
        From Table(f_Str2list2(参数列表_In, '#', '^'));
        --采用全部删除，再插入
        Forall I In 1 .. t_参数值.Count
          Insert Into Zldeptparas (参数id, 部门id, 参数值) Values (n_参数id, t_部门id(I), t_参数值(I));
      End If;
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlparameters_Imp_Details;
/
--84990:刘硕,2015-06-18,参数整理
--84598:刘硕,2015-06-17,zl_GetSysParameter性能问题处理
Create Or Replace Function Zltools.zl_GetSysParameter
(
  参数_In   Zlparameters.参数名%Type,
  模块_In   Zlparameters.模块%Type := Null,
  系统_In   Zlparameters.系统%Type := 1,
  部门id_In Zldeptparas.部门id%Type := 0
  --功能：获取当前系统中指定参数的参数值 
  ----本函数主要供其他过程调用，因zlParameters是公共表，程序中使用公共部件的公共函数 
  ----调用时注意,如果参数值为空或没有该参数,则返回空 
  --参数： 
  ----参数_In：必须传入非Null值，以字符形式传入的参数号或参数名,注意参数名不能为数字。 
  ----系统_IN：非标准版系统需要传入系统号，注意是示扩展的系统号，如1，而不是100。无系统传入Null 
) Return Varchar2 As
  v_系统 Zlparameters.系统%Type;
  v_私有 Zlparameters.私有%Type;
  v_本机 Zlparameters.本机%Type;
  v_部门 Zlparameters.部门%Type;

  v_参数id Zluserparas.参数id%Type;
  v_机器名 Zluserparas.机器名%Type;
  v_参数值 Zlparameters.参数值%Type;
Begin
  --确定系统,可能没有系统(如私有全局) 
  If 系统_In Is Not Null Then
    Select Min(编号) Into v_系统 From zlSystems Where Trunc(编号 / 100) = 系统_In;
  End If;

  --读取参数信息 
  Begin
    If Zl_To_Number(参数_In) <> 0 Then
      Select ID, Nvl(参数值, 缺省值), 私有, 本机, 部门
      Into v_参数id, v_参数值, v_私有, v_本机, v_部门
      From zlParameters
      Where 参数号 = Zl_To_Number(参数_In) And Nvl(模块, 0) = Nvl(模块_In, 0) And Nvl(系统, 0) = Nvl(v_系统, 0);
    Else
      Select ID, Nvl(参数值, 缺省值), 私有, 本机, 部门
      Into v_参数id, v_参数值, v_私有, v_本机, v_部门
      From zlParameters
      Where 参数名 = 参数_In And Nvl(模块, 0) = Nvl(模块_In, 0) And Nvl(系统, 0) = Nvl(v_系统, 0);
    End If;
  
  Exception
    When Others Then
      Return Null;
  End;
  If Nvl(v_部门, 0) = 0 Then
    --读取非部门参数值 
    If Nvl(v_私有, 0) = 1 Or Nvl(v_本机, 0) = 1 Then
      If Nvl(v_本机, 0) = 1 Then
        Select Sys_Context('USERENV', 'TERMINAL') Into v_机器名 From Dual;
      End If;
      Begin
        Select Nvl(参数值, v_参数值)
        Into v_参数值
        From zlUserParas
        Where 参数id = v_参数id And (用户名 = User Or Nvl(v_私有, 0) = 0) And (机器名 = Nvl(v_机器名, '空空') Or Nvl(v_本机, 0) = 0);
      Exception
        When Others Then
          Return v_参数值;
      End;
    End If;
  Else
    Begin
      Select Nvl(参数值, v_参数值) Into v_参数值 From Zldeptparas Where 参数id = v_参数id And 部门id = 部门id_In;
    Exception
      When Others Then
        Return v_参数值;
    End;
  End If;
  Return v_参数值;
End zl_GetSysParameter;
/
--84990:刘硕,2015-05-20,参数整理
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


--84990:刘硕,2015-05-28,参数整理
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
               Select 0 As 序号
               From Dual) M
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
                             a.连接数, Decode(b.Terminal, Null, 0, 1) As 状态, a.收集标志,a.升级服务器,a.站点,a.启用视频源
                From Zlclients a, (Select Distinct Terminal From V$session) b
                Where Upper(a.工作站) = Upper(b.Terminal(+))
                Order By a.Ip';
      Open Cur_Out For v_Sql;
    Else
      Open Cur_Out For
        Select Ip, 工作站, Cpu, 内存, 硬盘, 操作系统, 部门, 用途, 说明, 升级标志, 收集标志, 禁止使用, 连接数, 升级服务器, 站点, 启用视频源
        From zlClients
        Where Upper(工作站) = 工作站_In;
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
        Select Nvl(To_Number(参数值), 0)
        From zlOptions
        Where 参数号 = 4;
    End If;
    If 日志类型_In = '运行日志' Then
      Open Cur_Out For
        Select Count(*) 数量
        From zlDiaryLog
        Union All
        Select Nvl(To_Number(参数值), 0)
        From zlOptions
        Where 参数号 = 2;
    
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
--00000:刘硕,2016-08-18,解决补充版本升级上来导致的包头与包体不匹配
CREATE OR REPLACE Package b_Public Is
--公共过程
  Type t_Refcur Is Ref Cursor;
--功能：取系统日期
--调用列表：mdlMain.CurrentDate，clsDatabase.CurrentDate
  Procedure Get_Current_Date(Cursor_Out Out t_Refcur);
--功能：删除错误日志或运行日志
--调用列表：mdlMain.DeleteAllLog
  Procedure Delete_All_Log(Runtimelog_In In Number := 0);
--功能：删除当前运行日志
--调用列表：mdlMain.DeleteCurLog
  Procedure Delete_Diarylog
  (
    会话号_In   Number,
    用户名_In   Varchar2,
    工作站_In   Varchar2,
    部件名_In   Varchar2,
    工作内容_In Varchar2,
    进入时间_In Date
  );
--功能：删除当前错误日志
--调用列表：mdlMain.DeleteCurLog
  Procedure Delete_Errorlog
  (
    会话号_In   Number,
    用户名_In   Varchar2,
    工作站_In   Varchar2,
    类型_In     Number,
    错误序号_In Number,
    时间_In     Date
  );
--功能：取注册码
--调用列表：mdlMain.Get注册码
  Procedure Get_Regcode(Cursor_Out Out t_Refcur);
--功能：取版本号
--调用列表：mdlMain.UpgradeManager
  Procedure Get_Ver(Cursor_Out Out t_Refcur);
--功能：更新版本号
--调用列表：mdlMain.UpgradeManager
  Procedure Update_Ver(Verstring_In In Varchar2);
--功能：取得系统所有者名称
--调用列表：
--frmStatus.cmbsystem_Click、mdlMain.GetOwnerName、mdlMain.cmbSystem_Click
--frmAutoJobs.cmbSystem_Click、frmDataMove.cmbSystem_Click 、frmNoticeTools.cboSystem_Click
--frmProgPriv.ProgPriv、frmAppScript.cmbSystem_Click
  Procedure Get_Owner_Name
  (
    Cursor_Out Out t_Refcur,
    编号_In    In zlSystems.编号%Type := 0
  );

--功能：取注册表中信息
--调用列表：
--frmAbout.GetUnitInfo、frmAutoJobs.From_load、frmClientsUpgrade.InitInfor
--frmFilesSet.ShowEdit、frmRegist.From_load、frmAppScript.From_Load
--frmFilesSendToServer.InitInfo
  Procedure Get_Reginfo
  (
    Cursor_Out Out t_Refcur,
    项目_In    In zlRegInfo.项目%Type := Null
  );
--功能：取zlGetSvrToolsg数据
--调用列表：frmMDIMain.MDIForm_Load
  Procedure Get_Zlsvrtools(Cursor_Out Out t_Refcur);
--功能：取已安装系统清单
--调用列表：
--frmAppCheck.Form_Load、frmClearData.Form_Load、frmDataMove.Form_Load
--frmImp.FillSystem、frmLoadIn.FillSystem、frmLoadOut.FillSystem
--frmMDIMain.mnuFileRemove_Click、frmNoticeTools.Form_Activate、frmRoleGrant.FillSystem
--frmAppUpgrade.Form_Load、frmAppScript.Form_Load、frmExp.FillSystem
--frmInputTools.from_activate、fromRole.FillSystem、frmAutoJobs.From_load
--frmAppstart.sysCreated
  Procedure Get_Zlsystems
  (
    Cursor_Out Out t_Refcur,
    所有者_In  In zlSystems.所有者%Type := Null
  );

End b_Public;
/
--84990:刘硕,2015-05-28,参数整理
Create Or Replace Package Body Zltools.b_Public Is
  --功能：取系统日期
  Procedure Get_Current_Date(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select Sysdate As 日期 From Dual;
  End Get_Current_Date;

  --功能：删除错误日志或运行日志
  Procedure Delete_All_Log(Runtimelog_In In Number := 0) Is
    n_Count Number;
    n_Loop  Number;
  Begin
    If Runtimelog_In = 1 Then
      Select Count(进入时间) Into n_Count From zlDiaryLog;
      If n_Count > 1000 Then
        For n_Loop In 1 .. Ceil(n_Count - 1000) Loop
          Delete zlDiaryLog Where Rownum < 10001;
          Commit;
        End Loop;
      Else
        If n_Count > 0 Then
          Delete zlDiaryLog;
          Commit;
        End If;
      End If;
    Else
      Select Count(时间) Into n_Count From zlErrorLog;
      If n_Count > 1000 Then
        For n_Loop In 1 .. Ceil(n_Count - 1000) Loop
          Delete zlErrorLog Where Rownum < 10001;
          Commit;
        End Loop;
      Else
        If n_Count > 0 Then
          Delete zlErrorLog;
          Commit;
        End If;
      End If;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Delete_All_Log;

  --功能：删除当前运行日志
  Procedure Delete_Diarylog
  (
    会话号_In   Number,
    用户名_In   Varchar2,
    工作站_In   Varchar2,
    部件名_In   Varchar2,
    工作内容_In Varchar2,
    进入时间_In Date
  ) Is
  Begin
    Delete zlDiaryLog
    Where 会话号 = 会话号_In And 用户名 = 用户名_In And 工作站 = 工作站_In And 部件名 = 部件名_In And 工作内容 = 工作内容_In And 进入时间 = 进入时间_In;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Delete_Diarylog;

  --功能：删除当前错误日志
  Procedure Delete_Errorlog
  (
    会话号_In   Number,
    用户名_In   Varchar2,
    工作站_In   Varchar2,
    类型_In     Number,
    错误序号_In Number,
    时间_In     Date
  ) Is
  Begin
    Delete zlErrorLog
    Where 会话号 = 会话号_In And 用户名 = 用户名_In And 工作站 = 工作站_In And 类型 = 类型_In And 错误序号 = 错误序号_In And 时间 = 时间_In;
    Commit;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Delete_Errorlog;

  --功能：取注册码
  Procedure Get_Regcode(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select 内容 From zlRegInfo Where 项目 = '注册码' Or 项目 = '授权证章' Order By 行号;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Regcode;

  --功能：取版本号
  Procedure Get_Ver(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select 内容 From zlRegInfo Where 项目 = '版本号';
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Ver;

  --功能：更新版本号
  Procedure Update_Ver(Verstring_In In Varchar2) Is
  Begin
    Update zlRegInfo Set 内容 = Verstring_In Where 项目 = '版本号';
    If Sql%NotFound Then
      Insert Into zlRegInfo (项目, 行号, 内容) Values ('版本号', 1, Verstring_In);
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Update_Ver;

  --功能：取得系统所有者名称
  Procedure Get_Owner_Name
  (
    Cursor_Out Out t_Refcur,
    编号_In    In Zlsystems.编号%Type := 0
  ) Is
  Begin
    Open Cursor_Out For
      Select Upper(所有者) As 所有者 From zlSystems Where 编号 = 编号_In;
  End Get_Owner_Name;

  --功能：取注册表中信息
  Procedure Get_Reginfo
  (
    Cursor_Out Out t_Refcur,
    项目_In    In Zlreginfo.项目%Type := Null
  ) Is
  Begin
    If Trim(Nvl(项目_In, '空')) = '空' Then
      Open Cursor_Out For
        Select * From zlRegInfo;
    Else
      Open Cursor_Out For
        Select 内容 From zlRegInfo Where 项目 = 项目_In Order By 行号;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Reginfo;

  --功能：取zlGetSvrToolsg数据
  Procedure Get_Zlsvrtools(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select * From zlSvrTools Start With 上级 Is Null Connect By Prior 编号 = 上级 Order By Level, 编号;
  End Get_Zlsvrtools;

  --功能：取已安装系统清单
  Procedure Get_Zlsystems
  (
    Cursor_Out Out t_Refcur,
    所有者_In  In Zlsystems.所有者%Type := Null
  ) Is
  Begin
    If Nvl(所有者_In, '空') = '空' Then
      Open Cursor_Out For
        Select 编号, 名称, 共享号, Upper(所有者) 所有者, 安装日期, 正常安装, 版本号 From zlSystems Order By 编号;
    Else
      Open Cursor_Out For
        Select 编号, 名称, 共享号, Upper(所有者) 所有者, 安装日期, 正常安装, 版本号
        From zlSystems
        Where Upper(所有者) = Upper(所有者_In)
        Order By 编号;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlsystems;

End b_Public;
/