UPDATE 临床出诊号源 set  是否序号控制 =NULL,是否分时段=NULL,预约控制=NULL ,分诊方式=NULL ,诊室ID=NULL  ;
ALTER TABLE 临床出诊号源 DROP column 是否序号控制;
ALTER TABLE 临床出诊号源 DROP column 是否分时段;
ALTER TABLE 临床出诊号源 DROP column 预约控制;
ALTER TABLE 临床出诊号源 DROP column 分诊方式;
ALTER TABLE 临床出诊号源 DROP column 诊室ID;

DROP TABLE 临床出诊号源诊室;

Create Sequence 临床出诊号源限制_ID start with 1;
Create Table 临床出诊号源限制(
   ID number(18) not null,
   号源ID number(18),
   上班时段 varchar2(10),
   限号数 number(10),
   限约数 number(10),
   是否序号控制 number(2) default 0,
   是否分时段  NUMBER(2),
   预约控制 number(2),
   是否独占 number(2) default 0,   
   分诊方式 number(3),
   诊室ID number(18))
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊号源限制  Add Constraint 临床出诊号源限制_PK  Primary Key (ID) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊号源限制 Add Constraint 临床出诊号源限制_FK_号源ID Foreign Key (号源ID) References 临床出诊号源( ID) ;
Alter Table 临床出诊号源限制  Add Constraint 临床出诊号源限制_UQ_号源ID  Unique (号源ID,上班时段) Using Index Tablespace zl9Indexhis; 
Alter Table 临床出诊号源限制 Add Constraint 临床出诊号源限制_FK_诊室ID Foreign Key (诊室ID) References 门诊诊室( ID) ;
create Index 临床出诊号源限制_IX_诊室ID on 临床出诊号源限制(诊室ID);



Create Table 临床出诊号源诊室(
   限制ID number(18),
   诊室ID number(18))
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊号源诊室  Add Constraint 临床出诊号源诊室_PK  Primary Key (限制ID,诊室ID) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊号源诊室 Add Constraint 临床出诊号源诊室_FK_限制ID Foreign Key (限制ID) References 临床出诊号源限制( ID) ;
Alter Table 临床出诊号源诊室 Add Constraint 临床出诊号源诊室_FK_诊室ID Foreign Key (诊室ID) References 门诊诊室( ID) ;
create Index 临床出诊号源诊室_IX_诊室ID on 临床出诊号源诊室(诊室ID);

Create Table 临床出诊号源时段(
   限制ID number(18),
   序号 number(18),
   开始时间 Date,
   终止时间 Date,
   限制数量 number(10),
   是否预约 number(2))
TABLESPACE zl9BaseItem;

Alter Table 临床出诊号源时段  Add Constraint 临床出诊号源时段_PK  Primary Key (限制ID,序号) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊号源时段 Add Constraint 临床出诊号源时段_FK_限制ID Foreign Key (限制ID) References 临床出诊号源限制( ID) ;



Create Table 临床出诊号源控制(
   限制ID number(18),
   类型 number(2),
   性质 number(2),
   名称 varchar2(50),
   序号 number(18),
   控制方式 number(2),
   数量 number(16,5))
TABLESPACE zl9BaseItem ;

Alter Table 临床出诊号源控制  Add Constraint 临床出诊号源控制_PK  Primary Key (限制ID,类型,性质,名称,序号) Using Index Tablespace zl9Indexhis;
Alter Table 临床出诊号源控制 Add Constraint 临床出诊号源控制_FK_限制ID Foreign Key (限制ID) References 临床出诊号源限制(ID);



Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1114,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select '临床出诊号源限制_ID','SELECT' From Dual Union All
Select '临床出诊号源限制','SELECT' From Dual Union All
Select '临床出诊号源时段','SELECT' From Dual Union All
Select '临床出诊号源控制','SELECT' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;




Insert Into zlParameters(ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
Select Zlparameters_Id.Nextval, &n_System, 1114, A.* From (
  Select 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明 From zlParameters Where 1 = 0 Union All 
  Select 1, -null, -null, 1, -null, -null, 7, '显示缺省控制信息', '', '1', '在号源管理中控制在选择号源时，是否在下方显示控制的相关信息，比如：缺省的序号信息、诊室信息、三方预约控制信息等，。', '0-不显示，1-显示。', Null, Null, Null From Dual Union All
  Select 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明 From zlParameters Where 1 = 0) A;





Create Or Replace Procedure Zl_临床出诊号源_Delete(Id_In 临床出诊号源.Id%Type) As
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_Count  Number;
  l_限制id t_Numlist := t_Numlist();
Begin
  Select Count(1) Into n_Count From 临床出诊安排 Where 号源id = Id_In;

  If n_Count = 0 Then
  
    Select ID Bulk Collect Into l_限制id From 临床出诊号源限制 Where 号源id = Id_In;
  
    Forall I In 1 .. l_限制id.Count
      Delete 临床出诊号源时段 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete 临床出诊号源控制 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete 临床出诊号源诊室 Where 限制id = l_限制id(I);
  
    Delete 临床出诊号源限制 Where 号源id = Id_In;
    --假删除
  
    Delete From 临床出诊号源 Where ID = Id_In;
    If Sql%NotFound Then
      v_Err_Msg := '当前号源可能已被他人删除，不能再删除!';
      Raise Err_Item;
    End If;
    Return;
  End If;
  Update 临床出诊号源 Set 是否删除 = 1, 撤档时间 = Sysdate Where ID = Id_In And Nvl(是否删除, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '当前号源可能已被他人删除，不能再删除!';
    Raise Err_Item;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊号源_Delete;
/



Create Or Replace Procedure Zl_临床出诊号源_Modify
(
  操作类型_In     Number,
  Id_In           临床出诊号源.Id%Type,
  号类_In         临床出诊号源.号类%Type := Null,
  号码_In         临床出诊号源.号码%Type := Null,
  科室id_In       临床出诊号源.科室id%Type := 0,
  项目id_In       临床出诊号源.项目id%Type := 0,
  医生id_In       临床出诊号源.医生id%Type := Null,
  医生姓名_In     临床出诊号源.医生姓名%Type := Null,
  是否建病案_In   临床出诊号源.是否建病案%Type := 0,
  预约天数_In     临床出诊号源.预约天数%Type := 0,
  出诊频次_In     临床出诊号源.出诊频次%Type := 0,
  假日控制状态_In 临床出诊号源.假日控制状态%Type := 0,
  是否假日换休_In 临床出诊号源.是否假日换休%Type := 0,
  是否临床排班_In 临床出诊号源.是否临床排班%Type := 0,
  排班方式_In     临床出诊号源.排班方式%Type := 0
) As
  --操作类型_In 0-新增，1-修改，2-删除
  --分诊诊室_In 诊室ID，格式：诊室ID1;诊室ID2;诊室ID13;...
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_号源id 临床出诊号源.Id%Type;
  n_Count  Number;
Begin

  If 操作类型_In = 0 Then
    --增加号源
    n_号源id := Id_In;
  
    If Nvl(n_号源id, 0) = 0 Then
      Select 临床出诊号源_Id.Nextval Into n_号源id From Dual;
    End If;
    Insert Into 临床出诊号源
      (ID, 号类, 号码, 科室id, 项目id, 医生id, 医生姓名, 是否建病案, 预约天数, 出诊频次, 假日控制状态, 是否假日换休, 是否临床排班, 排班方式, 是否删除, 建档时间, 撤档时间)
    Values
      (n_号源id, 号类_In, 号码_In, 科室id_In, 项目id_In, 医生id_In, 医生姓名_In, 是否建病案_In, 预约天数_In, 出诊频次_In, 假日控制状态_In, 是否假日换休_In,
       是否临床排班_In, 排班方式_In, 0, Sysdate, To_Date('3000-01-01', 'yyyy-mm-dd'));
  
    Return;
  End If;

  --修改号源
  Update 临床出诊号源
  Set 号类 = 号类_In, 号码 = 号码_In, 科室id = 科室id_In, 项目id = 项目id_In, 医生id = 医生id_In, 医生姓名 = 医生姓名_In, 是否建病案 = 是否建病案_In,
      预约天数 = 预约天数_In, 出诊频次 = 出诊频次_In, 假日控制状态 = 假日控制状态_In, 是否假日换休 = 是否假日换休_In, 是否临床排班 = 是否临床排班_In, 排班方式 = 排班方式_In
  Where ID = Id_In And Nvl(是否删除, 0) = 0 And Nvl(撤档时间, Sysdate) >= Sysdate;
  If Sql%NotFound Then
    v_Err_Msg := '当前号源可能已被他人删除或停用，不能对该号源信息进行修改!';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊号源_Modify;
/


Create Or Replace Procedure Zl_临床出诊号源限制_Modify
(
  Id_In           临床出诊号源限制.Id%Type,
  号源id_In       临床出诊号源限制.号源id%Type,
  上班时段_In     临床出诊号源限制.上班时段%Type,
  限号数_In       临床出诊号源限制.限号数%Type,
  限约数_In       临床出诊号源限制.限约数%Type,
  是否序号控制_In 临床出诊号源限制.是否序号控制%Type,
  是否分时段_In   临床出诊号源限制.是否分时段%Type,
  预约控制_In     临床出诊号源限制.预约控制%Type,
  是否独占_In     临床出诊号源限制.是否独占%Type,
  分诊方式_In     临床出诊号源限制.分诊方式%Type,
  诊室id_In       临床出诊号源限制.诊室id%Type,
  号源诊室_In     Varchar2 := Null,
  号源时段_In     Varchar2 := Null,
  号源控制_In     Varchar2 := Null,
  删除号源限制_In Integer := 0
  
) As
  --号源时段_IN:序号,开始时间(HH:MM:SS),终止时(HH:MM:SS)间,数量,是否预约|...
  --号源诊室_IN:诊室id1,诊室id2,....
  --号源控制_IN:类型,性质,名称,控制方式,序号,数量|
  --删除号源限制_in:1-插入数据前，先删除号源限制,0-不删除数据，直接插入

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  l_限制id   t_Numlist := t_Numlist();
  n_Count    Number;
  v_开始时间 Varchar2(20);
  v_终止时间 Varchar2(20);

  n_序号     临床出诊号源时段.序号%Type;
  d_开始时间 临床出诊号源时段.开始时间%Type;
  d_终止时间 临床出诊号源时段.终止时间%Type;
  n_数量     临床出诊号源时段.限制数量%Type;
  n_是否预约 临床出诊号源时段.是否预约%Type;
  n_类型     临床出诊号源控制.类型%Type;
  n_性质     临床出诊号源控制.性质%Type;
  v_名称     临床出诊号源控制.名称%Type;
  n_控制方式 临床出诊号源控制.控制方式%Type;

Begin
  If Nvl(删除号源限制_In, 0) = 1 Then
    Select ID Bulk Collect Into l_限制id From 临床出诊号源限制 Where 号源id = 号源id_In;
    Forall I In 1 .. l_限制id.Count
      Delete 临床出诊号源时段 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete 临床出诊号源控制 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete 临床出诊号源诊室 Where 限制id = l_限制id(I);
  
    Delete 临床出诊号源限制 Where 号源id = 号源id_In;
    Delete From 临床出诊号源限制 Where 号源id = 号源id_In;
  
  End If;

  Select Count(1) Into n_Count From 临床出诊号源限制 Where ID = Id_In;
  If n_Count = 0 Then
    Insert Into 临床出诊号源限制
      (ID, 号源id, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 是否独占, 分诊方式, 诊室id)
    Values
      (Id_In, 号源id_In, 上班时段_In, 限号数_In, 限约数_In, 是否序号控制_In, 是否分时段_In, 预约控制_In, 是否独占_In, 分诊方式_In, 诊室id_In);
  
  End If;

  If 号源时段_In Is Not Null Then
    --插入号源缺省时间段
    For c_时间段集 In (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(号源时段_In, '|'))) Loop
      n_序号     := Null;
      v_开始时间 := Null;
      v_终止时间 := Null;
      n_数量     := Null;
      n_是否预约 := Null;
      For c_时间段 In (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(c_时间段集.值)) Order By 序号) Loop
        If c_时间段.序号 = 1 Then
          n_序号 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 2 Then
          v_开始时间 := c_时间段.值;
        End If;
      
        If c_时间段.序号 = 3 Then
          v_终止时间 := c_时间段.值;
        End If;
      
        If c_时间段.序号 = 4 Then
          n_数量 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 5 Then
          n_是否预约 := To_Number(c_时间段.值);
        End If;
      
      End Loop;
      d_开始时间 := To_Date('3000-01-01 ' || Nvl(v_开始时间, ''), 'yyyy-mm-dd hh24:mi:ss');
      d_终止时间 := To_Date('3000-01-01 ' || Nvl(v_终止时间, ''), 'yyyy-mm-dd hh24:mi:ss');
    
      If d_开始时间 >= d_开始时间 Then
        d_终止时间 := d_终止时间 + 1;
      End If;
    
      If Nvl(n_序号, 0) <> 0 Then
        Insert Into 临床出诊号源时段
          (限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约)
        Values
          (Id_In, n_序号, d_开始时间, d_终止时间, n_数量, n_是否预约);
      End If;
    End Loop;
  
  End If;

  --插入号源的缺省控制
  --号源控制_IN:类型,性质,名称,控制方式,序号,数量|
  If 号源控制_In Is Not Null Then
    For c_时间段集 In (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(号源控制_In, '|'))) Loop
      n_类型     := Null;
      n_性质     := Null;
      v_名称     := Null;
      n_序号     := Null;
      n_控制方式 := Null;
      n_数量     := Null;
    
      --类型,性质,名称,控制方式,序号,数量|
      For c_时间段 In (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(c_时间段集.值)) Order By 序号) Loop
        If c_时间段.序号 = 1 Then
          n_类型 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 2 Then
          n_性质 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 3 Then
          v_名称 := c_时间段.值;
        End If;
      
        If c_时间段.序号 = 4 Then
          n_控制方式 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 5 Then
          n_序号 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 6 Then
          n_数量 := To_Number(c_时间段.值);
        End If;
      
      End Loop;
    
      If v_名称 Is Not Null Then
        Insert Into 临床出诊号源控制
          (限制id, 类型, 性质, 名称, 序号, 控制方式, 数量)
        Values
          (Id_In, n_类型, n_性质, v_名称, n_序号, n_控制方式, n_数量);
      
      End If;
    End Loop;
  End If;
  --插入号源诊室
  If 号源诊室_In Is Not Null Then
    Insert Into 临床出诊号源诊室
      (限制id, 诊室id)
      Select Id_In As 限制id, Column_Value As 科室id From Table(f_Num2list(号源诊室_In));
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊号源限制_Modify;
/
