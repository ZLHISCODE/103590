----10.35.50---》10.35.60
--111205:刘硕,2017-07-06,部件清单重新整理
alter table ZLTOOLS.zlFilesUpgrade modify 业务部件 varchar2(500);
--110050:刘硕,2017-6-19,优化会话认证
Alter Table ZLTOOLS.zlRegFunc  Drop Constraint zlRegFunc_UQ cascade drop index;
Alter Table ZLTOOLS.zlRegFunc  Add Constraint zlRegFunc_PK PRIMARY KEY(系统, 序号, 功能) USING INDEX PCTFREE 5;
--97672:高腾,2017-4-11,老的消息程序(zlapptools),修改"邮件消息检查周期"
Update Zltools.Zlparameters Set 缺省值 = '60' Where Nvl(系统, 0) = 0 And Nvl(模块, 0) = 0 And 参数名 = '邮件消息检查周期';

--00000:张永康,2017-07-04,增加参数限制返回字符的最大长度，避免超过字段上限而出错
Create Or Replace Function Zltools.f_List2str
(
  p_Strlist   In t_Strlist,
  p_Delimiter In Varchar2 Default ',',
  p_Distinct  In Number Default 1,
  p_Maxlength In Number Default 4000
) Return Varchar2 Is
  l_String Long;
  l_Add    Number;
  --功能：将一个列表集合转换为一个缺省以逗号分隔的字符串。
  --例：
  --Select 科室, f_List2str(Cast(Collect(人员 Order By 编号) As t_Strlist)) 人员列表
  --From (Select a.名称 As 科室, c.姓名 As 人员,c.编号
  --      From 部门表 A, 部门人员 B, 人员表 C
  --      Where a.Id = b.部门id And b.人员id = c.Id
  --      Order By 科室, 人员)
  --Group By 科室

  --此函数不支持with方式构造的临时内存表，这将会报错：ORA-00932: 数据类型不一致: 应为 -, 但却获得 -。
  --例如：With Test As (Select '内科' As 科室,'张三' As 人员 From Dual Union All......)
  --     Select 科室,f_List2str(cast(COLLECT(人员) as t_Strlist)) tt From Test Group By 科室
Begin
  If p_Strlist.Count > 0 Then
    For I In p_Strlist.First .. p_Strlist.Last Loop
      l_Add := 0;
      If p_Distinct = 1 Then
        If Instr(',' || l_String || ',', ',' || p_Strlist(I) || ',') = 0 Then
          l_Add := 1;
        End If;
      Else
        l_Add := 1;
      End If;
      If l_Add = 1 Then
        If I != p_Strlist.First Then
          l_String := l_String || p_Delimiter;
        End If;
        l_String := l_String || p_Strlist(I);
        If Lengthb(l_String) > p_Maxlength Then
          l_String := Substr(l_String, 1, p_Maxlength);
          Return l_String;
        End If;
      End If;
    End Loop;
  End If;
  Return l_String;
End f_List2str;
/

--97672:高腾,2017-4-11,老的消息程序(zlapptools),将SQL语句修改为过程
CREATE OR REPLACE Procedure zltools.Zl_Zlmsgstate_Edit
(
  操作_In     Number, --0-新增,1-修改,2-删除
  消息id_In   Zlmsgstate.消息id%Type,
  类型_In     Zlmsgstate.类型%Type := Null,
  用户_In     Zlmsgstate.用户%Type := Null,
  身份_In     Zlmsgstate.身份%Type := Null,
  删除_In     Zlmsgstate.删除%Type := Null,
  状态_In     Zlmsgstate.状态%Type := Null,
  保存天数_In Number := Null
) Is
  n_总数 Number;
  n_数量 Number;
Begin
  If 操作_In = 0 Then
    Insert Into Zlmsgstate
      (消息id, 类型, 用户, 身份, 删除, 状态)
    Values
      (消息id_In, 类型_In, 用户_In, 身份_In, 删除_In, 状态_In);
  Elsif 操作_In = 1 Then
    If 状态_In Is Not Null Then
      If 身份_In Is Null Then
        Update Zlmsgstate Set 状态 = 状态_In Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
      Else
        Update Zlmsgstate
        Set 状态 = 状态_In, 身份 = 身份_In
        Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
      End If;
    End If;
  
    If 删除_In Is Not Null Then
      If 类型_In Is Not Null Then
        Update Zlmsgstate Set 删除 = 删除_In Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
      Else
        Update Zlmsgstate Set 删除 = 删除_In Where 消息id = 消息id_In And 用户 = 用户_In;
      End If;
      Select Count(*), Sum(Decode(删除, 2, 1, 0)) Into n_总数, n_数量 From Zlmsgstate Where 消息id = 消息id_In;
      If n_总数 = n_数量 Then
        Delete From Zlmessages Where Id = 消息id_In;
      End If;
    End If;
  Elsif 操作_In = 2 Then
    Delete From Zlmessages Where 时间 < Sysdate - 保存天数_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlmsgstate_Edit;
/

--97672:高腾,2017-4-11,老的消息程序(zlapptools),将SQL语句修改为过程
CREATE OR REPLACE Procedure zltools.Zl_Zlmsgstate_Addaddressee
(
  消息id_In   Zlmsgstate.消息id%Type,
  类型_In     Zlmsgstate.类型%Type,
  状态_In     Zlmsgstate.状态%Type,
  用户身份_In Varchar2 --格式：用户名1,身份1#用户名2,身份2#用户名3,身份3
) Is
Begin
  Insert Into Zlmsgstate
    (消息id, 类型, 用户, 身份, 删除, 状态)
    Select 消息id_In, 类型_In, C1, C2, 0, 状态_In From Table(f_Str2list2(用户身份_In, '#', ','));
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlmsgstate_Addaddressee;
/

--97672:高腾,2017-4-11,老的消息程序(zlapptools),将SQL语句修改为过程
CREATE OR REPLACE Procedure zltools.Zl_Zlmessages_New
(
  消息id_In   Zlmessages.Id%Type,
  会话id_In   Zlmessages.会话id%Type,
  收件人_In   Zlmessages.收件人%Type,
  主题_In     Zlmessages.主题%Type,
  内容_In     Zlmessages.内容%Type,
  背景色_In   Zlmessages.背景色%Type,
  类型_In     Zlmsgstate.类型%Type,
  用户_In     Zlmsgstate.用户%Type,
  身份_In     Zlmsgstate.身份%Type,
  状态_In     Zlmsgstate.状态%Type,
  操作类型_In Number, --操作类型。1-答复；2-全部答复；3-转发；0-新建邮件
  修改id_In   Zlmsgstate.消息id%Type,
  修改类型_In Zlmsgstate.类型%Type
) Is
  n_Count Number;
Begin
  --插入或修改消息记录
  Select Count(1) Into n_Count From Zlmessages Where Id = 消息id_In;
  If n_Count = 0 Then
    Insert Into Zlmessages
      (Id, 会话id, 发件人, 时间, 收件人, 主题, 内容, 背景色)
    Values
      (消息id_In, 会话id_In, 身份_In, Sysdate, 收件人_In, 主题_In, 内容_In, 背景色_In);
  Else
    Update Zlmessages
    Set 时间 = Sysdate, 收件人 = 收件人_In, 主题 = 主题_In, 内容 = 内容_In, 背景色 = 背景色_In
    Where Id = 消息id_In;
  End If;

  --删除所有记录
  Delete Zlmsgstate Where 消息id = 消息id_In;
  --增加发件人记录
  Insert Into Zlmsgstate
    (消息id, 类型, 用户, 身份, 删除, 状态)
  Values
    (消息id_In, 类型_In, 用户_In, 身份_In, 0, 状态_In);
  --为原件加上答复或转发标志
  If 操作类型_In = 1 Or 操作类型_In = 2 Then
    Update Zlmsgstate
    Set 状态 = Substr(状态, 1, 1) || '1' || Substr(状态, 3, 2)
    Where 消息id = 修改id_In And 类型 = 修改类型_In And 用户 = 用户_In;
  Elsif 操作类型_In = 3 Then
    Update Zlmsgstate
    Set 状态 = Substr(状态, 1, 1) || '11' || Substr(状态, 4, 1)
    Where 消息id = 修改id_In And 类型 = 修改类型_In And 用户 = 用户_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlmessages_New;
/

--97672:高腾,2017-4-17,zlSvrNotice自动提醒服务,将SQL语句修改为过程
Create Or Replace Procedure Zltools.Zl_Zlnoticerec_Edit
(
  操作_In     Number, --0:增加，1:修改，2:删除
  提醒序号_In Zlnoticerec.提醒序号%type,
  用户名_In   Zlnoticerec.用户名%type,
  检查时间_In Zlnoticerec.检查时间%type,
  检查结果_In Zlnoticerec.检查结果%type,
  提醒标志_In Zlnoticerec.提醒标志%type,
  已读标志_In Zlnoticerec.已读标志%type,
  提醒内容_In Zlnoticerec.提醒内容%type
) Is
Begin
  If 操作_In = 0 Then
    Insert Into Zlnoticerec
      (提醒序号, 用户名, 检查时间, 检查结果, 提醒内容, 提醒标志, 已读标志)
    Values
      (提醒序号_In, 用户名_In, 检查时间_In, 检查结果_In, 提醒内容_In, 提醒标志_In, 已读标志_In);
  Elsif 操作_In = 1 Then
    If 已读标志_In Is Null Then
      Update Zlnoticerec
      Set 提醒时间 = Sysdate, 提醒标志 = 提醒标志_In
      Where 提醒序号 = 提醒序号_In And 用户名 = 用户名_In;
    Else
      Update Zlnoticerec Set 已读标志 = 已读标志_In Where 提醒序号 = 提醒序号_In And 用户名 = 用户名_In;
    End If;
  Else
    Delete From Zlnoticerec Where 提醒序号 = 提醒序号_In And 用户名 = 用户名_In;
  End If;
End Zl_Zlnoticerec_Edit;
/

--105511:余智勇,2017-06-21,自定义报表增加分类
Create Table Zltools.Zlrptclasses(
  ID Number(18), 
  上级id Number(18), 
  名称 Varchar2(30), 
  说明 Varchar2(100)
) PCTFREE 5;

--105511:余智勇,2017-06-21,自定义报表增加分类
alter table Zltools.zlReports add 分类ID Number(18);
alter table Zltools.zlRPTGroups add 分类ID Number(18);
Create Sequence Zltools.Zlrptclasses_Id Start With 1;
Alter Table Zltools.Zlrptclasses Add Constraint Zlrptclasses_Pk Primary Key(ID) Using Index;
Alter Table Zltools.Zlrptclasses Add Constraint Zlrptclasses_Uq_名称 Unique(名称) Using Index;
Alter Table Zltools.Zlrptclasses Add Constraint Zlrptclasses_Fk_上级id Foreign Key(上级id) References Zlrptclasses(ID) On Delete Cascade;
Alter Table Zltools.zlReports Add Constraint Zlreports_Uq_分类id Unique(分类id, ID) Using Index;
alter table Zltools.zlReports add constraint ZLREPORTS_FK_分类ID foreign key (分类ID) references zlRPTClasses (ID);
alter table Zltools.zlRPTGroups add constraint ZLRPTGROUPS_UQ_分类ID unique (分类ID, ID) using index;
alter table Zltools.zlRPTGroups add constraint ZLRPTGROUPS_FK_分类ID foreign key (分类ID) references zlRPTClasses (ID);
Create Index Zltools.Zlrptclasses_Ix_上级id On Zlrptclasses(上级id);

--00000:张永康,2017-07-06,数据表分类
Create Table ZLTools.zlTables(
    系统    Number(5),
    表名    Varchar2(30),
    表空间  Varchar2(30),
    分类    Varchar2(3)
);
Alter Table ZLTools.zlTables Add Constraint zlTables_PK Primary Key (表名,系统) USING INDEX PCTFREE 5;

--A1:静态基础数据
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLBAKTABLEINDEX','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLBAKTABLES','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLBASECODE','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLBIGTABLES','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLCOMPONENT','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLDATAMOVE','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLFILES','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLFILESEXPIRED','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLFILESUPGRADE','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLKILLPROCESS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLMENUS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLMODULERELAS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLOPTIONS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLPARAMETERS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLPINYIN','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLPROGFUNCS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLPROGPRIVS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLPROGRAMS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLPROGRELAS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLREGFUNC','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLREGINFO','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLSVRTOOLS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLSYSTEMS','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLTools.zlTables','ZLTOOLSTBS','A1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLUSUALFUNC','ZLTOOLSTBS','A1');

--A2:动态基础数据
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLAUTOJOBS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLBAKSPACES','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLCLIENTPARALIST','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLCLIENTPARASET','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLCLIENTSCHEME','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLCONNECTIONS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLDEPTPARAS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLFUNCPARS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLFUNCTIONS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLINSUREBASE','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLINSURECOMPONENTS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLINSUREFUNCS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLINSUREMODULS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLINSUREOPERATION','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLINSUREPRIVS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLMGRGRANT','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLNODELIST','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLNOTICES','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLNOTICEUSR','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLPERIODS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLREPORTS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLROLEGRANT','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLROLEGROUPS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRPTCLASSES','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRPTCOLPROTERTY','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRPTCONDS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRPTDATAS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRPTFMTS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRPTGRAPHS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRPTGROUPS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRPTITEMS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRPTPARS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRPTPUTS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRPTRELATION','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRPTSQLS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRPTSUBS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLSYSFILES','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLUNITINFOIMAGE','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLUNITINFOITEM','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLUPGRADESERVER','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLUSERPARAS','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLUSERROLES','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLXLSDIRECTORY','ZLTOOLSTBS','A2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLXLSVERIFY','ZLTOOLSTBS','A2');

--A3:知识数据

--B1:业务活动数据
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLMESSAGES','ZLTOOLSTBS','B1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLMSGSTATE','ZLTOOLSTBS','B1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLPROCEDURE','ZLTOOLSTBS','B1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLPROCEDURETEXT','ZLTOOLSTBS','B1');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLUPGRADE','ZLTOOLSTBS','B1');

--B2:临时数据
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLREGAUDIT','','B2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLREGFILE','','B2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLTRIGGERS','ZLTOOLSTBS','B2');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLUPGRADECONFIG','ZLTOOLSTBS','B2');

--B3:日志数据
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLCLIENTUPDATELOG','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLDATAMOVELOG','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLDIARYLOG','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLERRORLOG','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLNOTICEREC','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLPARACHANGEDLOG','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRPTRUNHISTORY','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRPTSQLSHISTORY','ZLTOOLSTBS','B3');
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLUPGRADELOG','ZLTOOLSTBS','B3');

--C1:状态数据
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLCLIENTS','ZLTOOLSTBS','C1');

--C2:汇总数据

--C3:余额数据