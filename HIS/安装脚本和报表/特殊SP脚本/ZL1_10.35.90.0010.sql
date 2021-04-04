----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------

--123419:王煜,2018-05-10,新增手工销自动记帐费用附加标志含义
create or replace view 出院病人自动记帐 as
Select p.类型, p.病人id, p.主页id, Nvl(a.姓名, i.姓名) As 姓名, Nvl(a.性别, i.性别) As 性别, Nvl(a.年龄, i.年龄) As 年龄, Nvl(a.住院号, i.住院号) As 住院号,
       a.费别, p.科室id, p.病区id, p.床号, p.附加床位, p.收费细目id, p.收入项目id, 1 As 标志, p.现价 As 标准单价, p.开始日期, p.终止日期,
       p.终止日期 - p.开始日期 As 天数, p.数量, p.经治医师, p.责任护士, p.操作员编号, p.操作员姓名,p.医疗小组id
From 病人信息 I, 病案主页 A,
     (Select 2 As 类型, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量,
               b.医疗小组id
       From 自动计价项目 A,
            (Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, 床位等级id, 1 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名, 上次计算时间,
                     a.医疗小组id
              From 病人变动记录 A
              Where 开始原因 <> 10
              Union All
              Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, i.从项id As 床位等级id, i.从项数次 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名,
                     上次计算时间, b.医疗小组id
              From 病人变动记录 B, 收费从属项目 I
              Where b.床位等级id = i.主项id And b.开始原因 <> 10 And i.固有从属 > 0) B, 收费价目 P
       Where a.病区id = b.病区id And Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.计算标志 = 1 And b.床位等级id = p.收费细目id And Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2))) And p.价格等级 Is Null
       Union All
       Select 1 As 类型, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量,
              b.医疗小组id
       From 自动计价项目 A,
            (Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, 护理等级id, 1 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名, 上次计算时间, 医疗小组id
              From 病人变动记录
              Where 开始原因 <> 10
              Union All
              Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, i.从项id As 护理等级id, i.从项数次 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名,
                     上次计算时间, b.医疗小组id
              From 病人变动记录 B, 收费从属项目 I
              Where b.护理等级id = i.主项id And b.开始原因 <> 10 And i.固有从属 > 0) B, 收费价目 P, 收费项目目录 C
       Where a.病区id = b.病区id And b.附加床位 <> 1 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And a.计算标志 = 2 And
             b.护理等级id = p.收费细目id And b.护理等级id = c.Id And Nvl(c.计算方式, 0) <> 1 And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2))) And p.价格等级 Is Null
       Union All
       Select 3 As 类型, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, a.数量,
              b.医疗小组id
       From (Select 病区id, 计算标志, 收费细目id, 1 As 数量, 启用日期
              From 自动计价项目
              Union All
              Select 病区id, 计算标志, 从项id, i.从项数次 As 数量, 启用日期
              From 自动计价项目 A, 收费从属项目 I
              Where a.收费细目id = i.主项id And i.固有从属 > 0) A, 病人变动记录 B, 收费价目 P
       Where a.病区id = b.病区id And b.附加床位 <> 1 And b.开始原因 <> 10 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.收费细目id = p.收费细目id And (a.计算标志 = 6 And b.床位等级id Is Not Null Or a.计算标志 =7) And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2))) And p.价格等级 Is Null) P
Where i.病人id = p.病人id And a.病人id = p.病人id And a.主页id = p.主页id;

create or replace view 在院病人自动记帐 as
Select p.类型,p.病人id, p.主页id, Nvl(a.姓名, i.姓名) As 姓名, Nvl(a.性别, i.性别) As 性别, Nvl(a.年龄, i.年龄) As 年龄, Nvl(a.住院号, i.住院号) As 住院号,
       a.费别, p.科室id, p.病区id, p.床号, p.附加床位, p.收费细目id, p.收入项目id, 1 As 标志, p.现价 As 标准单价, p.开始日期, p.终止日期,
       p.终止日期 - p.开始日期 As 天数, p.数量, p.经治医师, p.责任护士, p.操作员编号, p.操作员姓名,p.医疗小组id
From 病人信息 I, 病案主页 A,
     (Select 2 As 类型, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量,
              b.医疗小组id
       From 自动计价项目 A,
            (Select a.病人id, a.主页id, a.开始时间, a.附加床位, a.病区id, a.科室id, a.床号, a.床位等级id, 1 As 数量, a.责任护士, a.经治医师, a.终止时间,
                     a.操作员编号, a.操作员姓名, a.上次计算时间, a.医疗小组id
              From 病人变动记录 A, 病人信息 B
              Where a.开始原因 <> 10 And a.病人id = b.病人id And a.主页id = b.主页id And b.在院 = 1
              Union All
              Select b.病人id, b.主页id, 开始时间, 附加床位, b.病区id, b.科室id, 床号, i.从项id As 床位等级id, i.从项数次 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号,
                     操作员姓名, 上次计算时间, b.医疗小组id
              From 病人变动记录 B, 收费从属项目 I, 病人信息 C
              Where b.病人id = c.病人id And b.主页id = c.主页id And c.在院 = 1 And b.床位等级id = i.主项id And b.开始原因 <> 10 And i.固有从属 > 0) B,
            收费价目 P
       Where a.病区id = b.病区id And Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.计算标志 = 1 And b.床位等级id = p.收费细目id And Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2))) And p.价格等级 Is Null
       Union All
       Select 1 As 类型, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量,
              b.医疗小组id
       From 自动计价项目 A,
            (Select a.病人id, a.主页id, 开始时间, 附加床位, a.病区id, a.科室id, 床号, 护理等级id, 1 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名, 上次计算时间,
                     a.医疗小组id
              From 病人变动记录 A, 病人信息 B
              Where 开始原因 <> 10 And a.病人id = b.病人id And a.主页id = b.主页id And b.在院 = 1
              Union All
              Select b.病人id, b.主页id, 开始时间, 附加床位, b.病区id, b.科室id, 床号, i.从项id As 护理等级id, i.从项数次 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号,
                     操作员姓名, 上次计算时间, b.医疗小组id
              From 病人变动记录 B, 收费从属项目 I, 病人信息 C
              Where b.护理等级id = i.主项id And b.病人id = c.病人id And b.主页id = c.主页id And c.在院 = 1 And b.开始原因 <> 10 And i.固有从属 > 0) B,
            收费价目 P, 收费项目目录 C
       Where a.病区id = b.病区id And b.附加床位 <> 1 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And a.计算标志 = 2 And
             b.护理等级id = p.收费细目id And b.护理等级id = c.Id And Nvl(c.计算方式, 0) <> 1 And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2))) And p.价格等级 Is Null
       Union All
       Select 3 As 类型, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, a.数量,
               b.医疗小组id
       From (Select 病区id, 计算标志, 收费细目id, 1 As 数量, 启用日期
              From 自动计价项目
              Union All
              Select 病区id, 计算标志, 从项id, i.从项数次 As 数量, 启用日期
              From 自动计价项目 A, 收费从属项目 I
              Where a.收费细目id = i.主项id And i.固有从属 > 0) A, 病人变动记录 B, 收费价目 P, 病人信息 C
       Where a.病区id = b.病区id And b.病人id = c.病人id And b.主页id = c.主页id And c.在院 = 1 And b.附加床位 <> 1 And b.开始原因 <> 10 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.收费细目id = p.收费细目id And (a.计算标志 = 6 And b.床位等级id Is Not Null Or a.计算标志=7) And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2))) And p.价格等级 Is Null) P
Where i.病人id = p.病人id And a.病人id = p.病人id And a.主页id = p.主页id;

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
--123419:王煜,2018-05-10,新增手工销自动记帐费用附加标志含义
CREATE OR REPLACE Procedure Zl1_Autocptone
( 
  病人id_In   In Number, 
  主页id_In   In Number, 
  期间_In     In Varchar2, 
  在院记帐_In In Number := 0, 
  强制记帐_In In Number := 0 
) As 
 
  ------------------------------------------------------------------------- 
  --功能说明：完成指定病人指定期间自动计价项目表设置自动计算的项目进行记帐处理 
  --          1、系统首先根据系统参数"修正上期自动计费"，修改以往该病人自动记帐记录标志; 
  --          2、综合病人的床位变化、入出转情况、调价情况等多项因素，结合期间跨度、病人费 
  --             别等完成费用的正确计算： 
  --             如果发现已经计算，则修改标志为正常;如果未计算，则插入新的自动记帐记录; 
  --             作废以前的错误计算的记录; 
  --             统计本次变动(新增和作废)，填写余额表和汇总表; 
  --入口参数： 
  --       病人ID_IN  number    病人身份ID 
  --       主页ID_IN  number    病案主页ID，两个参数共同确定需要计算的病人 
  --       期间_IN  varchar2     需要计算的最小期间 
  --       在院记帐_IN number   为1时,仅计算在院病人的费用 
  --       强制记帐_IN number   为1时,不受病案主页.禁止自动记帐属性控制 
  --调用关系：zl1_AutoCptPati/zl1_AutoCptWard/zl1_AutoCptAll 调用本过程 
  n_Count  Number(5);
   
  Cursor v_Autocur 
  ( 
    期间_In Varchar2, 
    Insure  病案主页.险类%Type 
  ) Is 
    Select l.类型,l.病人id, l.主页id, l.姓名, l.性别, l.年龄, l.住院号, l.费别, l.科室id, l.病区id, l.床号, l.附加床位, l.收费细目id, l.收入项目id, l.标志, l.标准单价, 
           Greatest(l.开始日期, Trunc(p.开始日期)) As 开始日期, l.终止日期, l.天数, l.数量, l.经治医师, l.责任护士, l.操作员编号, l.操作员姓名, i.险类, i.大类id, 
           k.算法, k.统筹比额, l.医疗小组id 
    From (Select * From 出院病人自动记帐 Where 病人id = 病人id_In And 主页id = 主页id_In) L, 
         (Select Min(开始日期) As 开始日期 From 期间表 Where 期间 >= 期间_In) P, 保险支付项目 I, 保险支付大类 K 
    Where Trunc(l.终止日期) >= Trunc(p.开始日期) And l.收费细目id = i.收费细目id(+) And i.险类(+) = Insure And i.大类id = k.Id(+) 
    Order By l.开始日期; 
 
  Cursor v_Autocurzy 
  ( 
    期间_In Varchar2, 
    Insure  病案主页.险类%Type 
  ) Is 
    Select l.类型, l.病人id, l.主页id, l.姓名, l.性别, l.年龄, l.住院号, l.费别, l.科室id, l.病区id, l.床号, l.附加床位, l.收费细目id, l.收入项目id, l.标志, l.标准单价, 
           Greatest(l.开始日期, Trunc(p.开始日期)) As 开始日期, l.终止日期, l.天数, l.数量, l.经治医师, l.责任护士, l.操作员编号, l.操作员姓名, i.险类, i.大类id, 
           k.算法, k.统筹比额,l.医疗小组id 
    From (Select * From 在院病人自动记帐 Where 病人id = 病人id_In And 主页id = 主页id_In) L, 
         (Select Min(开始日期) As 开始日期 From 期间表 Where 期间 >= 期间_In) P, 保险支付项目 I, 保险支付大类 K 
    Where Trunc(l.终止日期) >= Trunc(p.开始日期) And l.收费细目id = i.收费细目id(+) And i.险类(+) = Insure And i.大类id = k.Id(+) 
    Order By l.开始日期; 
 
  n_Insure       病案主页.险类%Type; 
  v_Billno       Varchar2(8); --费用表实际的自动记帐号码 
  n_Datecount    Integer; --日期计数器 
  d_Datefrom     Date; --开始计算日期 
  d_Dateto       Date; --终止计算日期 
  d_Datelast     Date; 
  n_Billcount    Number(5) := 0; --单据序号计数器 
  n_Exsetax      Number(16, 2) := 0; --费用收取比率 
  n_Exsetax_Temp Number(16, 2) := 0; --费用收取比率 
  n_Summoney     Number(16, 2) := 0; --金额 
 
  Cursor v_Sumcur 
  ( 
    Billno    Varchar2, 
    Datestart Date 
  ) Is 
    Select 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Sum(Decode(附加标志, 0, 1, -1) * 应收金额) As 应收金额, 
           Sum(Decode(附加标志, 0, 1, -1) * 实收金额) As 实收金额 
    From 住院费用记录 
    Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 
          (NO = Billno Or 附加标志 = 5 And 发生时间 >= Datestart) 
    Group By 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id; 
 
  n_Dec            Number; --金额小数位数 
  d_登记时间       Date; --登记时间 
  d_发生时间       Date; --发生时间 
  n_Dates          Number(3, 1); --当前记录的天数，全天为1 
  n_Do             Number(1); 
  n_返回值         病人余额.预交余额%Type; 
  n_Delete         Number; 
  n_医疗小组id     住院费用记录.医疗小组id%Type; 
  n_护理计算标准   Number(2); --护理费计算标准 
  n_收费细目id     Number(18); 
  n_Temp           Number(18); 
  l_护理id         t_Numlist := t_Numlist(); 
  l_护理等级       t_Numlist := t_Numlist(); 
  n_护理项目       Number(2); --1:是护理项目;0-非不护理 
  n_价格           收费价目.现价%Type; 
  n_护理已处理     Number(2); --1-护理费已经处理,;0-未处理 
  n_收入项目id     Number(18); 
  n_从属项目       Number(2); 
  n_审核标志       病案主页.审核标志%Type; 
  n_住院状态       病案主页.状态%Type; 
  n_病人审核方式   Number(2); 
  n_未入科禁止记账 Number(2); 
  n_禁止自动记帐   Number(2); 
 
  n_病人病区id 住院费用记录.病人病区id%Type; 
  n_开单部门id 住院费用记录.开单部门id%Type; 
 
  --已经计算了的护理类型 
  Type t_护理_Rec Is Record( 
    收费细目id 收费项目目录.Id%Type, 
    日期       Date); 
  Type t_护理 Is Table Of t_护理_Rec; 
  c_护理 t_护理 := t_护理(); 
Begin 
  Begin 
    Select 险类, Nvl(审核标志, 0), Nvl(状态, 0), Nvl(是否禁止自动记帐, 0) 
    Into n_Insure, n_审核标志, n_住院状态, n_禁止自动记帐 
    From 病案主页 
    Where 病人id = 病人id_In And 主页id = 主页id_In; 
  Exception 
    When Others Then 
      Return; 
  End; 
 
  If 强制记帐_In = 0 And n_禁止自动记帐 = 1 Then 
    Return; 
  End If; 
 
  n_病人审核方式   := Nvl(zl_GetSysParameter(185), 0); 
  n_未入科禁止记账 := Nvl(zl_GetSysParameter(215), 0); 
  If n_病人审核方式 = 1 And Nvl(n_审核标志, 0) >= 1 Then 
    Return; 
  End If; 
  If n_未入科禁止记账 = 1 And n_住院状态 = 1 Then 
    Return; 
  End If; 
 
  v_Billno := Nextno(17); 
 
  --金额小数位数 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(160), '0')) 
  Into n_Dec, n_护理计算标准 
  From Dual; 
 
  --每天5点以前，将记录时间登记为昨天，否则登记为当时 
  Select Decode(Sign(To_Number(To_Char(Sysdate, 'HH24')) - 5), -1, Trunc(Sysdate) - 1 / 24 / 60, Sysdate) 
  Into d_登记时间 
  From Dual; 
 
  --锁定该病人的记录,以免重复计算 
  Update 病案主页 Set 状态 = 状态 Where 病人id = 病人id_In And 主页id = 主页id_In; 
 
  ----------------------------------------------------------------- 
  d_Datefrom := Sysdate + 1000; 
  d_Dateto   := Sysdate - 1000; 
  n_Do       := 0; 
  -------------------------------------------------------------------- 
  If n_护理计算标准 = 1 Then 
    --同天以最高价位的护理费为准,先将其护理等级记住, 
    For v_护理 In (Select Distinct 护理等级id 
                 From (Select 护理等级id 
                        From 病人变动记录 
                        Where 开始原因 <> 10 And 病人id = 病人id_In And 主页id = 主页id_In 
                        Union All 
                        Select i.从项id As 护理等级id 
                        From 病人变动记录 B, 收费从属项目 I 
                        Where b.护理等级id = i.主项id And 病人id = 病人id_In And 主页id = 主页id_In And b.开始原因 <> 10 And i.固有从属 > 0)) Loop 
      If Nvl(v_护理.护理等级id, 0) <> 0 Then 
        l_护理id.Extend; 
        l_护理id(l_护理id.Count) := v_护理.护理等级id; 
      End If; 
    End Loop; 
  End If; 
  ----------------------------------------------------------------- 
  --循环检查计算情况，并增加正确和新计算的记录 
  ----------------------------------------------------------------- 
  If 在院记帐_In = 1 Then 
    For v_Currrow In v_Autocurzy(期间_In, n_Insure) Loop 
      If v_Currrow.医疗小组id Is Null Then 
        n_医疗小组id := Zl_医疗小组_Get(v_Currrow.科室id, v_Currrow.操作员姓名, v_Currrow.病人id, v_Currrow.主页id, d_发生时间); 
      Else 
        n_医疗小组id := v_Currrow.医疗小组id; 
      End If; 
 
      If d_Datefrom > v_Currrow.开始日期 Then 
        d_Datefrom := v_Currrow.开始日期; 
        n_Do       := 1; 
        --将本次开始计算时间以后的已计算记录标志修改 
        Update 住院费用记录 
        Set 附加标志 = 5 
        Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 附加标志 <> 8 And Nvl(医嘱序号, 0) = 0 And 
              发生时间 >= v_Currrow.开始日期; 
      End If; 
 
      If d_Dateto < v_Currrow.终止日期 Then 
        d_Dateto := v_Currrow.终止日期; 
      End If; 
      n_收费细目id := v_Currrow.收费细目id; 
      n_护理项目   := 0; 
      --护理费计算标准:0-按最后一次护理计算;1-按价格最高的护理等级计算。 
      If n_护理计算标准 = 1 Then 
        --先确定是否护理项目,如果是,则需要重新进行计算 
        Select Count(*) Into n_护理项目 From Table(l_护理id) Where Column_Value = n_收费细目id; 
      End If; 
 
      --提取当前收入项目的收费比率 
      Begin 
        Select 实收比率 
        Into n_Exsetax 
        From (Select 实收比率 
               From 费别明细 
               Where 费别 = v_Currrow.费别 And 收费细目id = v_Currrow.收费细目id And 
                     (Abs(v_Currrow.标准单价 * v_Currrow.数量) Between 应收段首值 And 应收段尾值) 
               Union All 
               Select 实收比率 
               From 费别明细 
               Where 费别 = v_Currrow.费别 And 收入项目id = v_Currrow.收入项目id And 
                     (Abs(v_Currrow.标准单价 * v_Currrow.数量) Between 应收段首值 And 应收段尾值) And Not Exists 
                (Select 1 From 费别明细 Where 费别 = v_Currrow.费别 And 收费细目id = v_Currrow.收费细目id)); 
      Exception 
        When Others Then 
          n_Exsetax := 100.00; 
      End; 
 
      n_Exsetax := Nvl(n_Exsetax, 100); 
      For n_Datecount In 0 .. (Trunc(v_Currrow.终止日期 + 0.5) - Trunc(v_Currrow.开始日期)) - 1 Loop 
        d_发生时间 := Greatest(v_Currrow.开始日期, Trunc(v_Currrow.开始日期 + n_Datecount)); 
        n_Dates    := Least(Trunc(v_Currrow.开始日期 + n_Datecount + 1), v_Currrow.终止日期) - 
                      Greatest(v_Currrow.开始日期, Trunc(v_Currrow.开始日期 + n_Datecount)); 
        --判断是否手工销帐
        Select Count(1)
        Into n_Count
        From 住院费用记录
        Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 2 And Nvl(附加标志, 0) = 1 And
              收费类别 = Decode(v_Currrow.类型, 1, 'H', 2, 'J', 收费类别) And 发生时间 = d_发生时间 And
              收费细目id = Decode(v_Currrow.类型, 3, v_Currrow.收费细目id, 收费细目id); 

        n_护理已处理 := 0; 
        If n_护理项目 = 1 Then 
          --1.先检查当天是否存在护理变动,只有存在多个护理变动的,才会去处理(以主项目为准) 
          n_从属项目 := 1; 
          If l_护理等级.Count > 0 Then 
            l_护理等级.Delete; 
          End If; 
          For v_护理 In (Select Distinct 护理等级id 
                       From 病人变动记录 
                       Where 开始原因 <> 10 And 病人id = 病人id_In And 主页id = 主页id_In And 
                             (Trunc(开始时间) = Trunc(d_发生时间) Or Trunc(Nvl(终止时间, Sysdate)) = Trunc(d_发生时间))) Loop 
            If Nvl(v_护理.护理等级id, 0) <> 0 Then 
              l_护理等级.Extend; 
              l_护理等级(l_护理等级.Count) := v_护理.护理等级id; 
              If Nvl(v_护理.护理等级id, 0) = Nvl(v_Currrow.收费细目id, 0) Then 
                n_从属项目 := 0; 
              End If; 
            End If; 
          End Loop; 
          If l_护理等级.Count > 1 Then 
            --2. 存在两个以上变动,则取价位最高的 
            n_Temp       := v_Currrow.收费细目id; 
            n_价格       := Nvl(v_Currrow.标准单价, 0); 
            n_收入项目id := v_Currrow.收入项目id; 
            --本身是从属项目时,由于主项目计算时,已经计算了的,所以就不再计算 
            If Nvl(n_从属项目, 0) = 1 Then 
              n_护理已处理 := 1; 
            End If; 
            --因为可能存在多个收入项目,但收费细目相同的情况,因此,必须先检查该项目是否已经参与计算过的 
            For I In 1 .. c_护理.Count Loop 
              If c_护理(I).收费细目id = v_Currrow.收费细目id And c_护理(I).日期 = Trunc(d_发生时间) Then 
                n_护理已处理 := 1; 
                Exit; 
              End If; 
            End Loop; 
            If Nvl(n_护理已处理, 0) = 0 Then 
              c_护理.Extend; 
              c_护理(c_护理.Count).收费细目id := v_Currrow.收费细目id; 
              c_护理(c_护理.Count).日期 := Trunc(d_发生时间); 
            End If; 
            If Nvl(n_从属项目, 0) = 0 And Nvl(n_护理已处理, 0) = 0 Then 
              --3.处理最高价位 
              For v_价位 In (Select /*+ rule */ 
                            a.Column_Value As 收费细目id, p.现价, p.收入项目id 
                           From Table(l_护理等级) A, 收费价目 P, 收费项目目录 C 
                           Where a.Column_Value = p.收费细目id And a.Column_Value = c.Id And d_发生时间 Between p.执行日期 And 
                                 Nvl(p.终止日期, Sysdate) And Nvl(c.计算方式, 0) <> 1 And p.价格等级 Is Null) Loop 
                If Nvl(v_价位.现价, 0) > n_价格 Then 
                  n_价格       := Nvl(v_价位.现价, 0); 
                  n_Temp       := v_价位.收费细目id; 
                  n_收入项目id := v_价位.收入项目id; 
                End If; 
              End Loop; 
 
              If n_Temp <> v_Currrow.收费细目id And Nvl(n_护理已处理, 0) = 0 Then 
 
                n_开单部门id := v_Currrow.科室id; 
                n_病人病区id := v_Currrow.病区id; 
 
                For c_变动记录 In (Select 病区id, 科室id 
                               From 病人变动记录 
                               Where 开始原因 <> 10 And 病人id = 病人id_In And 主页id = 主页id_In And 护理等级id + 0 = n_Temp And 
                                     (Trunc(开始时间) = Trunc(d_发生时间) Or Trunc(Nvl(终止时间, Sysdate)) = Trunc(d_发生时间)) 
                               Order By 开始时间 Desc) Loop 
                  n_开单部门id := c_变动记录.科室id; 
                  n_病人病区id := c_变动记录.病区id; 
                  Exit; 
                End Loop;               
                      
                --4. 不等的话,需要重新处理相关费用 
                For v_费用 In (Select n_Temp As 收费细目id, v_Currrow.数量 As 数量, n_价格 As 单价, n_收入项目id As 收入项目id 
                             From Dual 
                             Union All 
                             Select 从项id As 收费细目id, a.从项数次 As 数量, p.现价 As 单价, p.收入项目id 
                             From 收费从属项目 A, 收费价目 P, 收费项目目录 C 
                             Where a.从项id = p.收费细目id And a.从项id = c.Id And Nvl(c.计算方式, 0) <> 1 And a.主项id = n_Temp And 
                                   d_发生时间 Between p.执行日期 And Nvl(p.终止日期, Sysdate) And p.价格等级 Is Null) Loop 
                  --确定比例 
                  Begin 
                    Select 实收比率 
                    Into n_Exsetax_Temp 
                    From (Select 实收比率 
                           From 费别明细 
                           Where 费别 = v_Currrow.费别 And 收费细目id = v_费用.收费细目id And 
                                 (Abs(v_费用.单价 * v_费用.数量) Between 应收段首值 And 应收段尾值) 
                           Union All 
                           Select 实收比率 
                           From 费别明细 
                           Where 费别 = v_Currrow.费别 And 收入项目id = v_费用.收入项目id And 
                                 (Abs(v_费用.单价 * v_费用.数量) Between 应收段首值 And 应收段尾值) And Not Exists 
                            (Select 1 From 费别明细 Where 费别 = v_Currrow.费别 And 收费细目id = v_费用.收费细目id)); 
                  Exception 
                    When Others Then 
                      n_Exsetax_Temp := 100.00; 
                  End; 
                  n_Exsetax_Temp := Nvl(n_Exsetax_Temp, 100); 
                  --如果已经计算，原记录计算完全正确，则直接修改将标志改正 
                  Update 住院费用记录 
                  Set 附加标志 = 0 
                  Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And Nvl(加班标志, 0) = v_Currrow.附加床位 And 
                        病人科室id = Nvl(n_开单部门id, 0) And 病人病区id = Nvl(n_病人病区id, 0) And Nvl(床号, 0) = Nvl(v_Currrow.床号, 0) And 
                        收费细目id = v_费用.收费细目id And 收入项目id = v_费用.收入项目id And 发生时间 = d_发生时间 And 数次 = v_费用.数量 * n_Dates And 
                        标准单价 = v_费用.单价 And 应收金额 = Round(v_费用.单价 * v_费用.数量 * n_Dates, n_Dec) And 
                        实收金额 = Round(v_费用.单价 * v_费用.数量 * n_Dates * n_Exsetax / 100, n_Dec); 
 
                  If Sql%RowCount = 0 And n_Count=0 Then 
                    --如果未计算或计算错误，则增加正确的计算记录 
                    Insert Into 住院费用记录 
                      (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 
                       姓名, 性别, 年龄, 标识号, 床号, 费别, 记帐费用, 收费细目id, 收入项目id, 附加标志, 标准单价, 付数, 数次, 应收金额, 实收金额, 收费类别, 计算单位, 加班标志, 
                       收据费目, 开单人, 划价人, 操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 医疗小组id) 
                      Select 病人费用记录_Id.Nextval, 3, v_Billno, 1, Rownum + n_Billcount, Null, Null, 0, Null, 
                             Decode(v_Currrow.主页id, Null, 1, 2), v_Currrow.病人id, v_Currrow.主页id, n_病人病区id, n_开单部门id, 
                             n_开单部门id, n_病人病区id, v_Currrow.姓名, v_Currrow.性别, v_Currrow.年龄, v_Currrow.住院号, v_Currrow.床号, 
                             v_Currrow.费别, 1, v_费用.收费细目id, v_费用.收入项目id, 0, v_费用.单价, 1, v_费用.数量 * n_Dates, 
                             Round(v_费用.单价 * v_费用.数量 * n_Dates, n_Dec), 
                             Round(v_费用.单价 * v_费用.数量 * n_Dates * n_Exsetax / 100, n_Dec), i.类别, i.计算单位, v_Currrow.附加床位, 
                             j.收据费目, v_Currrow.经治医师, v_Currrow.责任护士, v_Currrow.操作员编号, v_Currrow.操作员姓名, d_发生时间, d_登记时间, 
                             Decode(v_Currrow.险类, Null, 0, 1), v_Currrow.大类id, 
                             Decode(v_Currrow.算法, 1, 
                                     Round(v_费用.单价 * v_费用.数量 * n_Dates * n_Exsetax / 100 * v_Currrow.统筹比额 / 100, n_Dec), 2, 
                                     v_Currrow.统筹比额, 0), n_医疗小组id 
                      From (Select 类别, 计算单位 
                             From 收费细目 
                             Where ID = v_费用.收费细目id And (撤档时间 Is Null Or 撤档时间 > d_发生时间)) I, 
                           (Select 收据费目 
                             From 收入项目 
                             Where ID = v_费用.收入项目id And (撤档时间 Is Null Or 撤档时间 > d_发生时间)) J; 
                    n_Billcount := n_Billcount + Sql%RowCount; 
                  End If; 
                  n_护理已处理 := 1; 
                End Loop; 
              End If; 
            End If; 
          End If; 
        End If; 
 
        If Nvl(n_护理已处理, 0) = 0 Then 
          --如果已经计算，原记录计算完全正确，则直接修改将标志改正 
          Update 住院费用记录 
          Set 附加标志 = 0 
          Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And Nvl(加班标志, 0) = v_Currrow.附加床位 And 
                病人科室id = v_Currrow.科室id And 病人病区id = Nvl(v_Currrow.病区id, 0) And Nvl(床号, 0) = Nvl(v_Currrow.床号, 0) And 
                收费细目id = v_Currrow.收费细目id And 收入项目id = v_Currrow.收入项目id And 发生时间 = d_发生时间 And 
                数次 = v_Currrow.数量 * n_Dates And 标准单价 = v_Currrow.标准单价 And 
                应收金额 = Round(v_Currrow.标准单价 * v_Currrow.数量 * n_Dates, n_Dec) And 
                实收金额 = Round(v_Currrow.标准单价 * v_Currrow.数量 * n_Dates * n_Exsetax / 100, n_Dec); 
 
          If Sql%RowCount = 0 And n_Count=0 Then 
            --如果未计算或计算错误，则增加正确的计算记录\ 
            Insert Into 住院费用记录 
              (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 姓名, 性别, 
               年龄, 标识号, 床号, 费别, 记帐费用, 收费细目id, 收入项目id, 附加标志, 标准单价, 付数, 数次, 应收金额, 实收金额, 收费类别, 计算单位, 加班标志, 收据费目, 开单人, 划价人, 
               操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 医疗小组id) 
              Select 病人费用记录_Id.Nextval, 3, v_Billno, 1, Rownum + n_Billcount, Null, Null, 0, Null, 
                     Decode(v_Currrow.主页id, Null, 1, 2), v_Currrow.病人id, v_Currrow.主页id, v_Currrow.病区id, v_Currrow.科室id, 
                     v_Currrow.科室id, v_Currrow.病区id, v_Currrow.姓名, v_Currrow.性别, v_Currrow.年龄, v_Currrow.住院号, 
                     v_Currrow.床号, v_Currrow.费别, 1, v_Currrow.收费细目id, v_Currrow.收入项目id, 0, v_Currrow.标准单价, 1, 
                     v_Currrow.数量 * n_Dates, Round(v_Currrow.标准单价 * v_Currrow.数量 * n_Dates, n_Dec), 
                     Round(v_Currrow.标准单价 * v_Currrow.数量 * n_Dates * n_Exsetax / 100, n_Dec), i.类别, i.计算单位, 
                     v_Currrow.附加床位, j.收据费目, v_Currrow.经治医师, v_Currrow.责任护士, v_Currrow.操作员编号, v_Currrow.操作员姓名, d_发生时间, 
                     d_登记时间, Decode(v_Currrow.险类, Null, 0, 1), v_Currrow.大类id, 
                     Decode(v_Currrow.算法, 1, 
                             Round(v_Currrow.标准单价 * v_Currrow.数量 * n_Dates * n_Exsetax / 100 * v_Currrow.统筹比额 / 100, n_Dec), 
                             2, v_Currrow.统筹比额, 0), n_医疗小组id 
              From (Select 类别, 计算单位 
                     From 收费细目 
                     Where ID = v_Currrow.收费细目id And (撤档时间 Is Null Or 撤档时间 > d_发生时间)) I, 
                   (Select 收据费目 
                     From 收入项目 
                     Where ID = v_Currrow.收入项目id And (撤档时间 Is Null Or 撤档时间 > d_发生时间)) J; 
 
            n_Billcount := n_Billcount + Sql%RowCount; 
          End If; 
        End If; 
      End Loop; 
    End Loop; 
  Else 
    For v_Currrow In v_Autocur(期间_In, n_Insure) Loop 
 
      If v_Currrow.医疗小组id Is Null Then 
        n_医疗小组id := Zl_医疗小组_Get(v_Currrow.科室id, v_Currrow.操作员姓名, v_Currrow.病人id, v_Currrow.主页id, d_发生时间); 
      Else 
        n_医疗小组id := v_Currrow.医疗小组id; 
      End If; 
 
      If d_Datefrom > v_Currrow.开始日期 Then 
        d_Datefrom := v_Currrow.开始日期; 
        n_Do       := 1; 
        --将本次开始计算时间以后的已计算记录标志修改 
        Update 住院费用记录 
        Set 附加标志 = 5 
        Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 附加标志 <> 8 And Nvl(医嘱序号, 0) = 0 And 
              发生时间 >= v_Currrow.开始日期; 
      End If; 
 
      If d_Dateto < v_Currrow.终止日期 Then 
        d_Dateto := v_Currrow.终止日期; 
      End If; 
      n_收费细目id := v_Currrow.收费细目id; 
      n_护理项目   := 0; 
      --护理费计算标准:0-按最后一次护理计算;1-按价格最高的护理等级计算。 
      If n_护理计算标准 = 1 Then 
        --先确定是否护理项目,如果是,则需要重新进行计算 
        Select Count(*) Into n_护理项目 From Table(l_护理id) Where Column_Value = n_收费细目id; 
      End If; 
 
      --提取当前收入项目的收费比率 
      Begin 
        Select 实收比率 
        Into n_Exsetax 
        From (Select 实收比率 
               From 费别明细 
               Where 费别 = v_Currrow.费别 And 收费细目id = v_Currrow.收费细目id And 
                     (Abs(v_Currrow.标准单价 * v_Currrow.数量) Between 应收段首值 And 应收段尾值) 
               Union All 
               Select 实收比率 
               From 费别明细 
               Where 费别 = v_Currrow.费别 And 收入项目id = v_Currrow.收入项目id And 
                     (Abs(v_Currrow.标准单价 * v_Currrow.数量) Between 应收段首值 And 应收段尾值) And Not Exists 
                (Select 1 From 费别明细 Where 费别 = v_Currrow.费别 And 收费细目id = v_Currrow.收费细目id)); 
      Exception 
        When Others Then 
          n_Exsetax := 100.00; 
      End; 
 
      n_Exsetax := Nvl(n_Exsetax, 100); 
      For n_Datecount In 0 .. (Trunc(v_Currrow.终止日期 + 0.5) - Trunc(v_Currrow.开始日期)) - 1 Loop 
        d_发生时间 := Greatest(v_Currrow.开始日期, Trunc(v_Currrow.开始日期 + n_Datecount)); 
        n_Dates    := Least(Trunc(v_Currrow.开始日期 + n_Datecount + 1), v_Currrow.终止日期) - 
                      Greatest(v_Currrow.开始日期, Trunc(v_Currrow.开始日期 + n_Datecount)); 
      
      --判断是否手工销帐
      Select Count(1)
      Into n_Count
      From 住院费用记录
      Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 2 And Nvl(附加标志, 0) = 1 And
            收费类别 = Decode(v_Currrow.类型, 1, 'H', 2, 'J', 收费类别) And 发生时间 = d_发生时间 And
            收费细目id = Decode(v_Currrow.类型, 3, v_Currrow.收费细目id, 收费细目id); 

        n_护理已处理 := 0; 
        If n_护理项目 = 1 Then 
          --1.先检查当天是否存在护理变动,只有存在多个护理变动的,才会去处理(以主项目为准) 
          n_从属项目 := 1; 
          If l_护理等级.Count > 0 Then 
            l_护理等级.Delete; 
          End If; 
          For v_护理 In (Select Distinct 护理等级id 
                       From 病人变动记录 
                       Where 开始原因 <> 10 And 病人id = 病人id_In And 主页id = 主页id_In And 
                             (Trunc(开始时间) = Trunc(d_发生时间) Or Trunc(Nvl(终止时间, Sysdate)) = Trunc(d_发生时间))) Loop 
            If Nvl(v_护理.护理等级id, 0) <> 0 Then 
              l_护理等级.Extend; 
              l_护理等级(l_护理等级.Count) := v_护理.护理等级id; 
              If Nvl(v_护理.护理等级id, 0) = Nvl(v_Currrow.收费细目id, 0) Then 
                n_从属项目 := 0; 
              End If; 
            End If; 
          End Loop; 
          If l_护理等级.Count > 1 Then 
            --2. 存在两个以上变动,则取价位最高的 
            n_Temp       := v_Currrow.收费细目id; 
            n_价格       := Nvl(v_Currrow.标准单价, 0); 
            n_收入项目id := v_Currrow.收入项目id; 
            --本身是从属项目时,由于主项目计算时,已经计算了的,所以就不再计算 
            If Nvl(n_从属项目, 0) = 1 Then 
              n_护理已处理 := 1; 
            End If; 
            --因为可能存在多个收入项目,但收费细目相同的情况,因此,必须先检查该项目是否已经参与计算过的 
            For I In 1 .. c_护理.Count Loop 
              If c_护理(I).收费细目id = v_Currrow.收费细目id And c_护理(I).日期 = Trunc(d_发生时间) Then 
                n_护理已处理 := 1; 
                Exit; 
              End If; 
            End Loop; 
            If Nvl(n_护理已处理, 0) = 0 Then 
              c_护理.Extend; 
              c_护理(c_护理.Count).收费细目id := v_Currrow.收费细目id; 
              c_护理(c_护理.Count).日期 := Trunc(d_发生时间); 
            End If; 
            If Nvl(n_从属项目, 0) = 0 And Nvl(n_护理已处理, 0) = 0 Then 
              --3.处理最高价位 
              For v_价位 In (Select /*+ rule */ 
                            a.Column_Value As 收费细目id, p.现价, p.收入项目id 
                           From Table(l_护理等级) A, 收费价目 P, 收费项目目录 C 
                           Where a.Column_Value = p.收费细目id And a.Column_Value = c.Id And d_发生时间 Between p.执行日期 And 
                                 Nvl(p.终止日期, Sysdate) And Nvl(c.计算方式, 0) <> 1 And p.价格等级 Is Null) Loop 
                If Nvl(v_价位.现价, 0) > n_价格 Then 
                  n_价格       := Nvl(v_价位.现价, 0); 
                  n_Temp       := v_价位.收费细目id; 
                  n_收入项目id := v_价位.收入项目id; 
                End If; 
              End Loop; 
 
              If n_Temp <> v_Currrow.收费细目id And Nvl(n_护理已处理, 0) = 0 Then 
 
                n_开单部门id := v_Currrow.科室id; 
                n_病人病区id := v_Currrow.病区id; 
 
                For c_变动记录 In (Select 病区id, 科室id 
                               From 病人变动记录 
                               Where 开始原因 <> 10 And 病人id = 病人id_In And 主页id = 主页id_In And 护理等级id + 0 = n_Temp And 
                                     (Trunc(开始时间) = Trunc(d_发生时间) Or Trunc(Nvl(终止时间, Sysdate)) = Trunc(d_发生时间)) 
                               Order By 开始时间 Desc) Loop 
                  n_开单部门id := c_变动记录.科室id; 
                  n_病人病区id := c_变动记录.病区id; 
                  Exit; 
                End Loop; 
                      
                --4. 不等的话,需要重新处理相关费用 
                For v_费用 In (Select n_Temp As 收费细目id, v_Currrow.数量 As 数量, n_价格 As 单价, n_收入项目id As 收入项目id 
                             From Dual 
                             Union All 
                             Select 从项id As 收费细目id, a.从项数次 As 数量, p.现价 As 单价, p.收入项目id 
                             From 收费从属项目 A, 收费价目 P, 收费项目目录 C 
                             Where a.从项id = p.收费细目id And a.从项id = c.Id And Nvl(c.计算方式, 0) <> 1 And a.主项id = n_Temp And 
                                   d_发生时间 Between p.执行日期 And Nvl(p.终止日期, Sysdate) And p.价格等级 Is Null) Loop 
                  --确定比例 
                  Begin 
                    Select 实收比率 
                    Into n_Exsetax_Temp 
                    From (Select 实收比率 
                           From 费别明细 
                           Where 费别 = v_Currrow.费别 And 收费细目id = v_费用.收费细目id And 
                                 (Abs(v_费用.单价 * v_费用.数量) Between 应收段首值 And 应收段尾值) 
                           Union All 
                           Select 实收比率 
                           From 费别明细 
                           Where 费别 = v_Currrow.费别 And 收入项目id = v_费用.收入项目id And 
                                 (Abs(v_费用.单价 * v_费用.数量) Between 应收段首值 And 应收段尾值) And Not Exists 
                            (Select 1 From 费别明细 Where 费别 = v_Currrow.费别 And 收费细目id = v_费用.收费细目id)); 
                  Exception 
                    When Others Then 
                      n_Exsetax_Temp := 100.00; 
                  End; 
                  n_Exsetax_Temp := Nvl(n_Exsetax_Temp, 100); 
                  --如果已经计算，原记录计算完全正确，则直接修改将标志改正 
                  Update 住院费用记录 
                  Set 附加标志 = 0 
                  Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And Nvl(加班标志, 0) = v_Currrow.附加床位 And 
                        病人科室id = Nvl(n_开单部门id, 0) And 病人病区id = Nvl(n_病人病区id, 0) And Nvl(床号, 0) = Nvl(v_Currrow.床号, 0) And 
                        收费细目id = v_费用.收费细目id And 收入项目id = v_费用.收入项目id And 发生时间 = d_发生时间 And 数次 = v_费用.数量 * n_Dates And 
                        标准单价 = v_费用.单价 And 应收金额 = Round(v_费用.单价 * v_费用.数量 * n_Dates, n_Dec) And 
                        实收金额 = Round(v_费用.单价 * v_费用.数量 * n_Dates * n_Exsetax / 100, n_Dec); 
 
                  If Sql%RowCount = 0 And n_Count=0Then 
                    --如果未计算或计算错误，则增加正确的计算记录 
                    Insert Into 住院费用记录 
                      (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 
                       姓名, 性别, 年龄, 标识号, 床号, 费别, 记帐费用, 收费细目id, 收入项目id, 附加标志, 标准单价, 付数, 数次, 应收金额, 实收金额, 收费类别, 计算单位, 加班标志, 
                       收据费目, 开单人, 划价人, 操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 医疗小组id) 
                      Select 病人费用记录_Id.Nextval, 3, v_Billno, 1, Rownum + n_Billcount, Null, Null, 0, Null, 
                             Decode(v_Currrow.主页id, Null, 1, 2), v_Currrow.病人id, v_Currrow.主页id, n_病人病区id, n_开单部门id, 
                             n_开单部门id, n_病人病区id, v_Currrow.姓名, v_Currrow.性别, v_Currrow.年龄, v_Currrow.住院号, v_Currrow.床号, 
                             v_Currrow.费别, 1, v_费用.收费细目id, v_费用.收入项目id, 0, v_费用.单价, 1, v_费用.数量 * n_Dates, 
                             Round(v_费用.单价 * v_费用.数量 * n_Dates, n_Dec), 
                             Round(v_费用.单价 * v_费用.数量 * n_Dates * n_Exsetax / 100, n_Dec), i.类别, i.计算单位, v_Currrow.附加床位, 
                             j.收据费目, v_Currrow.经治医师, v_Currrow.责任护士, v_Currrow.操作员编号, v_Currrow.操作员姓名, d_发生时间, d_登记时间, 
                             Decode(v_Currrow.险类, Null, 0, 1), v_Currrow.大类id, 
                             Decode(v_Currrow.算法, 1, 
                                     Round(v_费用.单价 * v_费用.数量 * n_Dates * n_Exsetax / 100 * v_Currrow.统筹比额 / 100, n_Dec), 2, 
                                     v_Currrow.统筹比额, 0), n_医疗小组id 
                      From (Select 类别, 计算单位 
                             From 收费细目 
                             Where ID = v_费用.收费细目id And (撤档时间 Is Null Or 撤档时间 > d_发生时间)) I, 
                           (Select 收据费目 
                             From 收入项目 
                             Where ID = v_费用.收入项目id And (撤档时间 Is Null Or 撤档时间 > d_发生时间)) J; 
                    n_Billcount := n_Billcount + Sql%RowCount; 
                  End If; 
                  n_护理已处理 := 1; 
                End Loop; 
              End If; 
            End If; 
          End If; 
        End If; 
 
        If Nvl(n_护理已处理, 0) = 0 Then 
          --如果已经计算，原记录计算完全正确，则直接修改将标志改正 
          Update 住院费用记录 
          Set 附加标志 = 0 
          Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And Nvl(加班标志, 0) = v_Currrow.附加床位 And 
                病人科室id = v_Currrow.科室id And 病人病区id = Nvl(v_Currrow.病区id, 0) And Nvl(床号, 0) = Nvl(v_Currrow.床号, 0) And 
                收费细目id = v_Currrow.收费细目id And 收入项目id = v_Currrow.收入项目id And 发生时间 = d_发生时间 And 
                数次 = v_Currrow.数量 * n_Dates And 标准单价 = v_Currrow.标准单价 And 
                应收金额 = Round(v_Currrow.标准单价 * v_Currrow.数量 * n_Dates, n_Dec) And 
                实收金额 = Round(v_Currrow.标准单价 * v_Currrow.数量 * n_Dates * n_Exsetax / 100, n_Dec); 
 
          If Sql%RowCount = 0 And n_Count=0 Then 
            --如果未计算或计算错误，则增加正确的计算记录\ 
            Insert Into 住院费用记录 
              (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 姓名, 性别, 
               年龄, 标识号, 床号, 费别, 记帐费用, 收费细目id, 收入项目id, 附加标志, 标准单价, 付数, 数次, 应收金额, 实收金额, 收费类别, 计算单位, 加班标志, 收据费目, 开单人, 划价人, 
               操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 医疗小组id) 
              Select 病人费用记录_Id.Nextval, 3, v_Billno, 1, Rownum + n_Billcount, Null, Null, 0, Null, 
                     Decode(v_Currrow.主页id, Null, 1, 2), v_Currrow.病人id, v_Currrow.主页id, v_Currrow.病区id, v_Currrow.科室id, 
                     v_Currrow.科室id, v_Currrow.病区id, v_Currrow.姓名, v_Currrow.性别, v_Currrow.年龄, v_Currrow.住院号, 
                     v_Currrow.床号, v_Currrow.费别, 1, v_Currrow.收费细目id, v_Currrow.收入项目id, 0, v_Currrow.标准单价, 1, 
                     v_Currrow.数量 * n_Dates, Round(v_Currrow.标准单价 * v_Currrow.数量 * n_Dates, n_Dec), 
                     Round(v_Currrow.标准单价 * v_Currrow.数量 * n_Dates * n_Exsetax / 100, n_Dec), i.类别, i.计算单位, 
                     v_Currrow.附加床位, j.收据费目, v_Currrow.经治医师, v_Currrow.责任护士, v_Currrow.操作员编号, v_Currrow.操作员姓名, d_发生时间, 
                     d_登记时间, Decode(v_Currrow.险类, Null, 0, 1), v_Currrow.大类id, 
                     Decode(v_Currrow.算法, 1, 
                             Round(v_Currrow.标准单价 * v_Currrow.数量 * n_Dates * n_Exsetax / 100 * v_Currrow.统筹比额 / 100, n_Dec), 
                             2, v_Currrow.统筹比额, 0), n_医疗小组id 
              From (Select 类别, 计算单位 
                     From 收费细目 
                     Where ID = v_Currrow.收费细目id And (撤档时间 Is Null Or 撤档时间 > d_发生时间)) I, 
                   (Select 收据费目 
                     From 收入项目 
                     Where ID = v_Currrow.收入项目id And (撤档时间 Is Null Or 撤档时间 > d_发生时间)) J; 
 
            n_Billcount := n_Billcount + Sql%RowCount; 
          End If; 
        End If; 
      End Loop; 
    End Loop; 
  End If; 
  If n_Do = 0 Then 
    --撤销出院后,如果修改出院时间为入院当天则不产生新费用,但以前的费用要冲销 
    Begin 
      Select Nvl(Trunc(b.上次计算时间), Trunc(b.终止时间)) 
      Into d_Datelast 
      From 病人变动记录 A, 病人变动记录 B 
      Where a.病人id = 病人id_In And a.主页id = 主页id_In And a.终止原因 = 1 And a.病人id = b.病人id And a.主页id = b.主页id And b.开始原因 = 1 And 
            Trunc(b.开始时间) = Trunc(a.终止时间) And a.附加床位 = 0 And b.附加床位 = 0; 
    Exception 
      When Others Then 
        Null; 
    End; 
    If d_Datelast Is Not Null Then 
      d_Datefrom := d_Datelast; 
      d_Dateto   := Sysdate; 
      Update 住院费用记录 
      Set 附加标志 = 5 
      Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 附加标志 <> 8 And Nvl(医嘱序号, 0) = 0 And 
            发生时间 >= d_Datefrom; 
    End If; 
  End If; 
 
  ----------------------------------------------------------------- 
  --作废以前计算的错误记录 
  ----------------------------------------------------------------- 
  Insert Into 住院费用记录 
    (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 姓名, 性别, 年龄, 标识号, 
     床号, 费别, 记帐费用, 收费细目id, 收入项目id, 附加标志, 标准单价, 付数, 数次, 应收金额, 实收金额, 收费类别, 计算单位, 加班标志, 收据费目, 开单人, 划价人, 操作员编号, 操作员姓名, 发生时间, 
     登记时间, 保险项目否, 保险大类id, 统筹金额, 医疗小组id) 
    Select 病人费用记录_Id.Nextval, 记录性质, NO, 2, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 
           姓名, 性别, 年龄, 标识号, 床号, 费别, 记帐费用, 收费细目id, 收入项目id, 0, 标准单价, 付数, -数次, -应收金额, -实收金额, 收费类别, 计算单位, 加班标志, 收据费目, 开单人, 
           划价人, 操作员编号, 操作员姓名, 发生时间, d_登记时间, 保险项目否, 保险大类id, -统筹金额, 医疗小组id 
    From 住院费用记录 
    Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 附加标志 = 5 And 发生时间 >= d_Datefrom; 
 
  ----------------------------------------------------------------- 
  --填写病人余额 
  ----------------------------------------------------------------- 
  Select Sum(Decode(附加标志, 0, 1, -1) * 实收金额) As 实收金额 
  Into n_Summoney 
  From 住院费用记录 
  Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 
        (NO = v_Billno Or 附加标志 = 5 And 发生时间 >= d_Datefrom); 
 
  Update 病人余额 
  Set 费用余额 = Nvl(费用余额, 0) + Nvl(n_Summoney, 0) 
  Where 病人id = 病人id_In And 性质 = 1 And 类型 = 2 
  Returning 费用余额 Into n_返回值; 
 
  If Sql%RowCount = 0 Then 
    Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (病人id_In, 1, 2, n_Summoney, 0); 
    n_返回值 := n_Summoney; 
  End If; 
 
  If Nvl(n_返回值, 0) = 0 Then 
    Delete From 病人余额 Where 性质 = 1 And 病人id = 病人id_In And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0; 
  End If; 
 
  ----------------------------------------------------------------- 
  --填写病人汇总费用 
  ----------------------------------------------------------------- 
  n_Delete := 0; 
  For v_Currrow In v_Sumcur(v_Billno, d_Datefrom) Loop 
    Update 病人未结费用 
    Set 金额 = Nvl(金额, 0) + Nvl(v_Currrow.实收金额, 0) 
    Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And Nvl(病人病区id, 0) = Nvl(v_Currrow.病人病区id, 0) And 
          Nvl(病人科室id, 0) = Nvl(v_Currrow.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(v_Currrow.开单部门id, 0) And 
          Nvl(执行部门id, 0) = Nvl(v_Currrow.执行部门id, 0) And Nvl(收入项目id, 0) = Nvl(v_Currrow.收入项目id, 0) And 来源途径 + 0 = 2 
    Returning 金额 Into n_返回值; 
 
    If Sql%RowCount = 0 Then 
      Insert Into 病人未结费用 
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额) 
      Values 
        (病人id_In, 主页id_In, v_Currrow.病人病区id, v_Currrow.病人科室id, v_Currrow.开单部门id, v_Currrow.执行部门id, v_Currrow.收入项目id, 2, 
         v_Currrow.实收金额); 
      n_返回值 := v_Currrow.实收金额; 
    End If; 
    If Nvl(n_返回值, 0) = 0 Then 
      n_Delete := 1; 
    End If; 
  End Loop; 
 
  If Nvl(n_Delete, 0) = 1 Then 
    Delete From 病人未结费用 Where 病人id = 病人id_In And 金额 = 0; 
  End If; 
 
  ----------------------------------------------------------------- 
  --将所有修改的附加标志还原为正常标志 
  ----------------------------------------------------------------- 
  Update 住院费用记录 
  Set 附加标志 = 0, 记录状态 = 3 
  Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 附加标志 = 5 And 发生时间 >= d_Datefrom; 
 
  ----------------------------------------------------------------- 
  --修改计算时间标志 
  ----------------------------------------------------------------- 
  Update 病人变动记录 
  Set 上次计算时间 = Least(d_Dateto, Nvl(终止时间, Greatest(开始时间, Sysdate))) 
  Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(终止时间, Sysdate) > d_Datefrom; 
  Commit; --单个病人提交 
Exception 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl1_Autocptone;
/

--123419:王煜,2018-05-10,新增手工销自动记帐费用附加标志含义
CREATE OR REPLACE Procedure Zl1_Autocalc_Pati_Charge_Nm
( 
  病人id_In       In 病案主页.病人id%Type, 
  主页id_In       In 病案主页.主页id%Type, 
  期间_In         In 期间表.期间%Type, 
  强制记帐_In     In Number := 0, 
  启用价格等级_In In Number := -1 
) As 
  ------------------------------------------------------------------------- 
  --功能说明：完成指定病人指定期间的自动记帐(主要是针对内蒙片区的自动记帐项目的计算) 
  --          1、系统首先根据系统参数"修正上期自动计费"，修改以往该病人自动记帐记录标志; 
  --          2、综合病人的床位变化、入出转情况、调价情况等多项因素，结合期间跨度、病人费 
  --             别等完成费用的正确计算： 
  --             如果发现已经计算，则修改标志为正常;如果未计算，则插入新的自动记帐记录; 
  --             作废以前的错误计算的记录; 
  --             统计本次变动(新增和作废)，填写余额表和汇总表; 
  --入口参数： 
  --       病人ID_IN  number    病人身份ID 
  --       主页ID_IN  number    病案主页ID，两个参数共同确定需要计算的病人 
  --       期间_IN  varchar2     需要计算的最小期间 
  --       强制记帐_IN number   为1时,不受病案主页.禁止自动记帐属性控制 
  --       启用价格等级_In number ：-1表示未判断价格等级,内部会自动去检查;0-不启用价格等级;1-启用了价格等级的 
  --调用关系：zl1_AutoCptPati/zl1_AutoCptWard/zl1_AutoCptAll 调用本过程 
  --自动记帐规则说明: 
  --   1. 床位:  计入不计出, 存在中途调整的(转科，转病区，等级变动等),12点以前，按转入科室为准;12点以后以转出科室为准 
  --   2.护理及其他费用:  入院当天按一天计算,出院当天中午12点之前算半天，12点之后算一天 
  ---------------------------------------------------------------------------------------------------------------------------------- 
  v_价格等级         收费价格等级.名称%Type; 
  v_付款方式价格等级 收费价格等级.名称%Type; 
 
  v_Temp      Varchar2(500); 
  v_Billno    Varchar2(8); --费用表实际的自动记帐号码 
  n_Billcount Number(5) := 0; --单据序号计数器 
 
  n_Exsetax  Number(16, 2) := 0; --费用收取比率 
  n_Summoney Number(16, 2) := 0; --金额 
 
  n_Dec    Number; --金额小数位数 
  n_Dates  Number(4, 1); --当前记录的天数，全天为1 
  n_Delete Number; 
  n_Exists Number; 
  n_Count  Number(5);
  n_返回值 病人余额.预交余额%Type; 
 
  v_收据费目   收入项目.收据费目%Type; 
  v_计算单位   收费项目目录.计算单位%Type; 
  n_住院状态   病案主页.状态%Type; 
  n_标准价格   收费价目.现价%Type; 
  n_收入项目id 收入项目.Id%Type; 
  v_类别       收费项目目录.类别%Type; 
  n_算法       保险支付大类.算法%Type; 
  n_统筹比额   保险支付大类.统筹比额%Type; 
  n_检查类型   病人自动计算.性质%Type; 
 
  n_病人审核方式   Number(2); 
  n_未入科禁止记账 Number(2); 
  n_护理价格优先   Number(2); 
  n_是否用价格等级 Number(2); 
  n_是否计算费用   Number(2); 
  n_Finded         Number(2); 
  n_类型           Number(2); --1-护理;2- 床位;3-其他 
  n_Find           Number(2); 
  n_Last           Number(2); 
  n_前科室id       病人自动计算.科室id%Type; 
  n_前病区id       病人自动计算.病区id%Type; 
  n_前收费细目id   病人自动计算.护理等级id%Type; 
  v_前床号         病人自动计算.床号%Type; 
  n_前床位等级id   病人自动计算.床位等级id%Type; 
  v_重算站点       部门表.站点%Type; 
 
  d_Start_Date Date; 
  d_登记时间   Date; --登记时间 
  d_发生时间   Date; --发生时间 
  d_Temp       Date; 
 
  d_床位时间_Max Date; 
  d_护理时间_Max Date; 
  d_其他时间_Max Date; 
 
  l_Mulit_细目id t_Numlist := t_Numlist(); 
 
  Type t_价格_Rec Is Ref Cursor; 
  c_价格_Rec t_价格_Rec; 
 
  Type t_病人变动_Rec Is Record( 
    ID         病人变动记录.Id%Type, 
    开始时间   病人变动记录.开始时间%Type, 
    终止时间   病人变动记录.开始时间%Type, 
    科室id     病人变动记录.科室id%Type, 
    病区id     病人变动记录.病区id%Type, 
    经治医师   病人变动记录.经治医师%Type, 
    责任护士   病人变动记录.责任护士%Type, 
    医疗小组id 病人变动记录.医疗小组id%Type); 
 
  Type c_病人变动_Rec Is Table Of t_病人变动_Rec; 
  r_病人变动 c_病人变动_Rec := c_病人变动_Rec(); 
  r_变动_Cur c_病人变动_Rec := c_病人变动_Rec(); 
 
  Cursor c_Sumcur_Rec 
  ( 
    Billno    Varchar2, 
    Datestart Date 
  ) Is 
    Select 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Sum(Decode(附加标志, 0, 1, -1) * 应收金额) As 应收金额, 
           Sum(Decode(附加标志, 0, 1, -1) * 实收金额) As 实收金额 
    From 住院费用记录 
    Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 
          (NO = Billno Or 附加标志 = 5 And 发生时间 >= Datestart) 
    Group By 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id; 
 
  Cursor c_Pati Is 
    Select a.病人id, a.主页id, Nvl(a.姓名, i.姓名) As 姓名, Nvl(a.性别, i.性别) As 性别, Nvl(a.年龄, i.年龄) As 年龄, Nvl(a.住院号, i.住院号) As 住院号, 
           a.费别, Nvl(a.险类, 0) As 险类, Nvl(a.审核标志, 0) As 审核标志, Nvl(a.状态, 0) As 住院状态, Nvl(a.是否禁止自动记帐, 0) As 是否禁止自动记帐, 
           a.医疗付款方式 As 付款方式, a.入院日期, a.出院日期 
    From 病案主页 A, 病人信息 I 
    Where a.病人id = i.病人id And a.病人id = 病人id_In And a.主页id = 主页id_In; 
 
  r_Pati c_Pati%RowType; 
  Cursor c_Pati_Change 
  ( 
    病人id_In In 病案主页.病人id%Type, 
    主页id_In In 病案主页.主页id%Type, 
    险类_In   In 病案主页.险类%Type, 
    期间_In   In Varchar2 
  ) Is 
    Select a.类型, a.Id, a.病人id, a.主页id, a.科室id, a.病区id, a.床号, a.附加床位, a.收费细目id, a.操作员编号, a.操作员姓名, a.开始时间, a.终止时间, a.启用日期, 
           a.数量, Greatest(a.开始日期, Trunc(p.开始日期)) As 开始日期, a.终止日期, a.天数, Nvl(Q1.站点, Q2.站点) As 站点, m.计算单位, m.类别, i.险类, 
           i.大类id, k.算法, k.统筹比额, Nvl(m.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 项目撤档时间, a.计算标志 
    From (Select 2 As 类型, b.Id, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, b.床位等级id As 收费细目id, b.操作员编号, b.操作员姓名, 
                  b.开始时间, b.终止时间, a.启用日期, b.数量, Trunc(b.开始时间) As 开始日期, Trunc(Nvl(b.终止时间, Sysdate)) As 终止日期, 
                  Trunc(Nvl(b.终止时间, Sysdate)) - Trunc(b.开始时间) As 天数, 0 As 计算标志 
           From 自动计价项目 A, 
                (Select a.Id, a.病人id, a.主页id, a.开始时间, a.附加床位, a.病区id, a.科室id, a.床号, a.床位等级id, 1 As 数量, a.终止时间, a.操作员编号, 
                         a.操作员姓名, a.上次计算时间 
                  From 病人自动计算 A 
                  Where a.性质 = 2 And a.病人id = 病人id_In And a.主页id = 主页id_In And 
                        Nvl(a.上次计算时间, a.开始时间) <= Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) 
                  Union All 
                  Select b.Id, b.病人id, b.主页id, 开始时间, 附加床位, b.病区id, b.科室id, 床号, i.从项id As 床位等级id, i.从项数次 As 数量, b.终止时间, 
                         b.操作员编号, b.操作员姓名, b.上次计算时间 
                  From 病人自动计算 B, 收费从属项目 I 
                  Where b.性质 = 2 And b.病人id = 病人id_In And b.主页id = 主页id_In And b.床位等级id = i.主项id And i.固有从属 > 0 And 
                        Nvl(b.上次计算时间, b.开始时间) <= Nvl(b.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))) B 
           Where a.病区id = b.病区id And a.计算标志 = 1 And Trunc(Nvl(b.终止时间, Sysdate)) >= a.启用日期 
           Union All 
           Select 1 As 类型, b.Id, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, b.护理等级id As 收费细目id, b.操作员编号, b.操作员姓名, 
                  b.开始时间, b.终止时间, a.启用日期, b.数量, Trunc(b.开始时间) As 开始日期, 
                  Decode(Trunc(b.终止时间), Trunc(b.开始时间), Trunc(Nvl(b.终止时间, Sysdate)), 
                          Zl_Date_Half(Nvl(b.终止时间, Trunc(Sysdate)), 1)) As 终止日期, 
                  Decode(Trunc(b.终止时间), Trunc(b.开始时间), Trunc(Nvl(b.终止时间, Sysdate)), 
                          Zl_Date_Half(Nvl(b.终止时间, Trunc(Sysdate)), 1)) - Trunc(b.开始时间) As 天数, 0 As 计算标志 
           From 自动计价项目 A, 
                (Select a.Id, a.病人id, a.主页id, a.开始时间, a.附加床位, a.病区id, a.科室id, a.床号, a.护理等级id, 1 As 数量, a.终止时间, a.操作员编号, 
                         a.操作员姓名, a.上次计算时间 
                  From 病人自动计算 A 
                  Where a.性质 = 1 And a.病人id = 病人id_In And a.主页id = 主页id_In And 
                        Nvl(a.上次计算时间, a.开始时间) <= Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) 
                  Union All 
                  Select b.Id, b.病人id, b.主页id, 开始时间, 附加床位, b.病区id, b.科室id, 床号, i.从项id As 床位等级id, i.从项数次 As 数量, b.终止时间, 
                         b.操作员编号, b.操作员姓名, b.上次计算时间 
                  From 病人自动计算 B, 收费从属项目 I 
                  Where b.性质 = 1 And b.病人id = 病人id_In And b.主页id = 主页id_In And b.护理等级id = i.主项id And i.固有从属 > 0 And 
                        Nvl(b.上次计算时间, b.开始时间) <= Nvl(b.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))) B 
           Where a.病区id = b.病区id And a.计算标志 = 2 And 
                 Decode(Trunc(b.终止时间), Trunc(b.开始时间), Trunc(Nvl(b.终止时间, Sysdate)), 
                        Zl_Date_Half(Nvl(b.终止时间, Trunc(Sysdate)), 1)) >= a.启用日期 
           Union All 
           Select 3 As 类型, b.Id, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, a.收费细目id, b.操作员编号, b.操作员姓名, b.开始时间, b.终止时间, 
                  a.启用日期, a.数量, Trunc(b.开始时间) As 开始日期, Zl_Date_Half(Nvl(b.终止时间, Trunc(Sysdate)), 1) As 终止日期, 
                  Trunc(Nvl(b.终止时间, Sysdate)) - Trunc(b.开始时间) As 天数, a.计算标志 
           From (Select 病区id, 计算标志, 收费细目id, 1 As 数量, 启用日期 
                  From 自动计价项目 
                  Union All 
                  Select 病区id, 计算标志, 从项id, i.从项数次 As 数量, 启用日期 
                  From 自动计价项目 A, 收费从属项目 I 
                  Where a.收费细目id = i.主项id And i.固有从属 > 0) A, 病人自动计算 B 
           Where a.病区id = b.病区id And b.病人id = 病人id_In And Zl_Date_Half(Nvl(b.终止时间, Trunc(Sysdate)), 1) >= a.启用日期 And 
                 b.主页id = 主页id_In And b.性质 = 3 And Nvl(b.附加床位, 0) = 0 And 
                 (a.计算标志 = 6 And b.床位等级id Is Not Null Or a.计算标志=7) And 
                 Nvl(b.上次计算时间, b.开始时间) <= Nvl(b.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))) A, 
         (Select Min(开始日期) As 开始日期 From 期间表 Where 期间 >= 期间_In) P, 保险支付项目 I, 保险支付大类 K, 收费项目目录 M, 部门表 Q1, 部门表 Q2 
    Where Trunc(Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD'))) >= Trunc(p.开始日期) And a.收费细目id = i.收费细目id(+) And 
          i.险类(+) = Nvl(险类_In, 0) And i.大类id = k.Id(+) And a.收费细目id = m.Id(+) And a.病区id = Q1.Id(+) And 
          a.科室id = Q2.Id(+) 
    Order By 类型, 附加床位, 开始时间; 
 
  r_Pati_Change     c_Pati_Change%RowType; 
  r_Pati_Change_Pre c_Pati_Change%RowType; 
 
  Function Get_Discount_Rate 
  ( 
    费别_In       病人信息.费别%Type, 
    收费细目id_In 费别明细.收费细目id%Type, 
    收入项目id_In 费别明细.收入项目id%Type, 
    金额_In       费别明细.应收段首值%Type 
  ) Return Number As 
    n_Discount_Rate Number(16, 5); 
  Begin 
    Begin 
      Select 实收比率 
      Into n_Discount_Rate 
      From (Select 实收比率 
             From 费别明细 
             Where 费别 = Nvl(费别_In, '-') And 收费细目id = Nvl(收费细目id_In, 0) And (金额_In Between 应收段首值 And 应收段尾值) 
             Union All 
             Select 实收比率 
             From 费别明细 
             Where 费别 = Nvl(费别_In, '-') And 收入项目id = Nvl(收入项目id_In, 0) And (金额_In Between 应收段首值 And 应收段尾值) And Not Exists 
              (Select 1 From 费别明细 Where 费别 = Nvl(费别_In, '-') And 收费细目id = Nvl(收费细目id_In, 0))); 
    Exception 
      When Others Then 
        n_Discount_Rate := 100.00; 
    End; 
    n_Discount_Rate := Nvl(n_Discount_Rate, 100); 
    Return n_Discount_Rate; 
  End Get_Discount_Rate; 
 
Begin 
 
  --获取病人信息 
  Begin 
    Open c_Pati; 
    Fetch c_Pati 
      Into r_Pati; 
  Exception 
    When Others Then 
      Return; 
  End; 
 
  If Nvl(强制记帐_In, 0) = 0 And Nvl(r_Pati.是否禁止自动记帐, 0) = 1 Then 
    Return; 
  End If; 
 
  n_病人审核方式   := Nvl(zl_GetSysParameter(185), 0); 
  n_未入科禁止记账 := Nvl(zl_GetSysParameter(215), 0); 
 
  If n_病人审核方式 = 1 And Nvl(r_Pati.审核标志, 0) >= 1 Then 
    Return; 
  End If; 
 
  If n_未入科禁止记账 = 1 And n_住院状态 = 1 Then 
    Return; 
  End If; 
 
  -------------------------------------------------------------------------------- 
  --1.初始化相关的参数 
  n_是否用价格等级 := 启用价格等级_In; 
  If n_是否用价格等级 < 0 Then 
    Select Nvl(Max(1), 0) Into n_是否用价格等级 From 收费价格等级应用 Where Rownum < 2; 
  End If; 
  --每天5点以前，将记录时间登记为昨天，否则登记为当时 
  Select Decode(Sign(To_Number(To_Char(Sysdate, 'HH24')) - 5), -1, Trunc(Sysdate) - 1 / 24 / 60, Sysdate) 
  Into d_登记时间 
  From Dual; 
 
  v_付款方式价格等级 := Null; 
  If Nvl(n_是否用价格等级, 0) = 1 Then 
    Select Max(价格等级) 
    Into v_付款方式价格等级 
    From 收费价格等级应用 A, 收费价格等级 B 
    Where a.价格等级 = b.名称 And a.性质 = 1 And a.医疗付款方式 = Nvl(r_Pati.付款方式, '-') And Nvl(b.是否适用普通项目, 0) = 1 And 
          Nvl(b.撤档时间, Sysdate + 1) > Sysdate; 
  End If; 
 
  --金额小数位数 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(160), '0')) 
  Into n_Dec, n_护理价格优先 
  From Dual; 
 
  n_护理价格优先 := 1; 
  --每天5点以前，将记录时间登记为昨天，否则登记为当时 
  Select Decode(Sign(To_Number(To_Char(Sysdate, 'HH24')) - 5), -1, Trunc(Sysdate) - 1 / 24 / 60, Sysdate) 
  Into d_登记时间 
  From Dual; 
  -------------------------------------------------------------------------------- 
 
  --锁定该病人的记录,以免重复计算 
  Update 病案主页 Set 状态 = 状态 Where 病人id = 病人id_In And 主页id = 主页id_In; 
 
  -------------------------------------------------------------------------------- 
  --2. 先将变动信息给记录集,以便提取经治医师和责任护士 
  d_护理时间_Max := Null; 
  d_其他时间_Max := Null; 
  d_床位时间_Max := Null; 
  For c_变动 In (Select ID, 开始时间, Nvl(终止时间, Sysdate + 1) As 终止时间, 科室id, 病区id, 护理等级id, 床位等级id, 经治医师, 责任护士, 医疗小组id 
               From 病人变动记录 A 
               Where 病人id = 病人id_In And 主页id = 主页id_In And 科室id Is Not Null And 
                     Nvl(终止时间, Sysdate) >= (Select Nvl(Min(上次计算时间), Sysdate - 1000) 
                                            From 病人自动计算 
                                            Where 病人id = 病人id_In And 主页id = 主页id_In) 
               Order By 病区id, 科室id, 开始时间 Desc) Loop 
 
    If c_变动.护理等级id Is Not Null And Nvl(d_护理时间_Max, c_变动.终止时间 - 1) <= c_变动.终止时间 Then 
      d_护理时间_Max := c_变动.终止时间; 
    End If; 
 
    If c_变动.床位等级id Is Not Null And Nvl(d_床位时间_Max, c_变动.终止时间 - 1) <= c_变动.终止时间 Then 
      d_床位时间_Max := c_变动.终止时间; 
    End If; 
 
    If c_变动.科室id Is Not Null And Nvl(d_其他时间_Max, c_变动.终止时间 - 1) <= c_变动.终止时间 Then 
      d_其他时间_Max := c_变动.终止时间; 
    End If; 
    r_病人变动.Extend; 
    r_病人变动(r_病人变动.Count).Id := c_变动.Id; 
    r_病人变动(r_病人变动.Count).开始时间 := c_变动.开始时间; 
    r_病人变动(r_病人变动.Count).终止时间 := c_变动.终止时间; 
    r_病人变动(r_病人变动.Count).科室id := c_变动.科室id; 
    r_病人变动(r_病人变动.Count).病区id := c_变动.病区id; 
    r_病人变动(r_病人变动.Count).经治医师 := c_变动.经治医师; 
    r_病人变动(r_病人变动.Count).责任护士 := c_变动.责任护士; 
    r_病人变动(r_病人变动.Count).医疗小组id := c_变动.医疗小组id; 
  End Loop; 
 
  --超过12:00,以12:00为准 
  d_护理时间_Max := Zl_Date_Half(d_护理时间_Max, 1); 
  d_床位时间_Max := Zl_Date_Half(d_床位时间_Max, 1); 
  d_其他时间_Max := Zl_Date_Half(d_其他时间_Max, 1); 
 
  ----------------------------------------------------------------- 
  --循环检查计算情况，并增加正确和新计算的记录 
  ----------------------------------------------------------------- 
 
  d_Start_Date := Sysdate + 1000; 
  d_Temp       := Sysdate - 1000;  

  --1.计算床位费 
  For c_自动记帐 In c_Pati_Change(病人id_In, 主页id_In, r_Pati.险类, 期间_In) Loop 
 
    If v_付款方式价格等级 Is Null Then 
      If Nvl(n_是否用价格等级, 0) = 1 And Nvl(r_Pati_Change_Pre.站点, '-') <> Nvl(c_自动记帐.站点, '-') Then 
        v_Temp     := Nvl(Zl_Get_Pricegrade(c_自动记帐.站点, 病人id_In, 主页id_In, r_Pati.付款方式), '|||') || '||||'; 
        v_价格等级 := Substr(v_Temp, 1, Instr(v_Temp, '|') - 1); 
      End If; 
    Else 
      v_价格等级 := v_付款方式价格等级; 
    End If; 
 
    r_Pati_Change := c_自动记帐; 
    If d_Start_Date > r_Pati_Change.开始日期 Then 
      d_Start_Date := r_Pati_Change.开始日期; 
    End If; 
 
    If Nvl(r_Pati_Change.类型, 0) <> Nvl(n_检查类型, 0) Then 
      n_检查类型 := Nvl(r_Pati_Change.类型, 0); 
      Update 住院费用记录 
      Set 附加标志 = 5 
      Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 附加标志 <> 8 And Nvl(医嘱序号, 0) = 0 And 
            发生时间 >= r_Pati_Change.开始日期 And 附加标志 <> 5 And (发药窗口 Is Null Or 发药窗口 = Nvl(n_检查类型, 0)); 
    End If; 
 
    --生成每天费用 
    For I In 0 .. r_Pati_Change.天数 Loop 
      v_重算站点    := Null; 
      r_Pati_Change := c_自动记帐; 
      d_发生时间    := Greatest(c_自动记帐.开始日期, Trunc(c_自动记帐.开始日期 + I)); 
      n_Dates       := Least(Trunc(c_自动记帐.开始日期 + I + 1), c_自动记帐.终止日期) - Greatest(c_自动记帐.开始日期, Trunc(c_自动记帐.开始日期 + I)); 
 
      If r_Pati_Change.类型 <> 2 Then 
        ----护理及其他费用:  入院当天按一天计算,出院当天中午12点之前算半天，12点之后算一天 
        If (r_Pati_Change.类型 = 1 And Trunc(d_护理时间_Max) = Trunc(d_发生时间) And d_护理时间_Max = r_Pati_Change.终止日期) Or 
           (r_Pati_Change.类型 = 3 And Trunc(d_其他时间_Max) = Trunc(d_发生时间) And d_其他时间_Max = r_Pati_Change.终止日期 And 
           r_Pati_Change.计算标志 = 7) Or (r_Pati_Change.类型 = 3 And Trunc(d_床位时间_Max) = Trunc(d_发生时间) And 
           d_床位时间_Max = r_Pati_Change.终止日期 And r_Pati_Change.计算标志 = 6) Then 
          If To_Char(r_Pati_Change.终止日期, 'hh24') >= 12 Then 
            n_Dates := 1; 
          Else 
            n_Dates    := 0.5; 
            d_发生时间 := Trunc(d_发生时间) + 0.5; 
          End If; 
        Else 
          n_Dates    := Least(Trunc(c_自动记帐.开始日期 + I + 1), Trunc(c_自动记帐.终止日期)) - 
                        Greatest(c_自动记帐.开始日期, Trunc(c_自动记帐.开始日期 + I)); 
          d_发生时间 := Trunc(d_发生时间); 
        End If; 
      End If; 
 
      If n_护理价格优先 = 1 And c_自动记帐.类型 = 1 Then 
        If d_发生时间 <> d_Temp Or Nvl(n_类型, 0) <> Nvl(c_自动记帐.类型, 0) Then 
          l_Mulit_细目id.Delete; 
          d_Temp := d_发生时间; 
          n_类型 := Nvl(c_自动记帐.类型, 0); 
        End If; 
 
        n_Finded := 0; 
        For J In 1 .. l_Mulit_细目id.Count Loop 
          If l_Mulit_细目id(J) = c_自动记帐.收费细目id Then 
            n_Finded := 1; 
            Exit; 
          End If; 
        End Loop; 
        If n_Finded = 0 Then 
          l_Mulit_细目id.Extend; 
          l_Mulit_细目id(l_Mulit_细目id.Count) := c_自动记帐.收费细目id; 
        End If; 
      End If; 
 
      n_Last         := 0; 
      n_是否计算费用 := 1; 
      If d_发生时间 > r_Pati_Change.项目撤档时间 Or n_Dates <= 0 Or (d_发生时间 > r_Pati_Change.终止日期) Then 
        Select Nvl(Max(1), 0) 
        Into n_Exists 
        From 病人自动计算 A 
        Where (a.终止原因 = 1 Or a.终止原因 = 10) And a.Id = r_Pati_Change.Id And r_Pati_Change.类型 <> 2; 
        If n_Exists = 0 Or n_Dates <= 0 Then 
          n_是否计算费用 := 0; 
        Else 
          n_Last := 0.5; 
        End If; 
      End If; 
 
      Select Nvl(Max(1), 0) 
      Into n_Exists 
      From 病人自动计算 A 
      Where a.终止原因 = 1 And a.Id = r_Pati_Change.Id And Exists 
       (Select 1 
             From 病人自动计算 
             Where 病人id = a.病人id And 主页id = a.主页id And 开始原因 = 2 And Trunc(开始时间) = Trunc(a.终止时间)); 
 
      If n_Exists = 1 Then 
        n_是否计算费用 := 1; 
        n_Dates        := 1; 
      End If; 
 
      If n_是否计算费用 = 1 Then 
 
        If d_发生时间 = Trunc(r_Pati_Change.开始时间) And Nvl(r_Pati_Change.附加床位, 0) = 0 Then 
          --12点以前，按转入出科室为准;12点以后以转出为准 
          If To_Char(r_Pati_Change.开始时间, 'hh24') >= 12 Then 
            --多条变动记录的处理 
            Begin 
              n_Find := 1; 
              Select 科室id, 病区id, Decode(r_Pati_Change.类型, 1, 护理等级id, 2, 床位等级id, r_Pati_Change.收费细目id), 床号, 床位等级id 
              Into n_前科室id, n_前病区id, n_前收费细目id, v_前床号, n_前床位等级id 
              From 病人自动计算 
              Where 病人id = 病人id_In And 主页id = 主页id_In And 性质 = r_Pati_Change.类型 And 
                    开始时间 = (Select Max(开始时间) 
                            From 病人自动计算 
                            Where 病人id = 病人id_In And 主页id = 主页id_In And 性质 = r_Pati_Change.类型 And 
                                  开始时间 <= To_Date(To_Char(r_Pati_Change.开始时间, 'yyyy-mm-dd') || ' 12:00:00', 
                                                  'yyyy-mm-dd hh24:mi:ss')); 
            Exception 
              When Others Then 
                n_Find := 0; 
            End; 
            If n_Find = 1 And n_前收费细目id Is Not Null And n_前病区id Is Not Null And n_前科室id Is Not Null And 
               Not (r_Pati_Change.类型 = 3 And r_Pati_Change.计算标志 = 6 And n_前床位等级id Is Null) Then 
              r_Pati_Change.科室id     := n_前科室id; 
              r_Pati_Change.病区id     := n_前病区id; 
              r_Pati_Change.收费细目id := n_前收费细目id; 
              r_Pati_Change.床号       := v_前床号; 
 
              Select Nvl(a.站点, b.站点) 
              Into v_重算站点 
              From 部门表 A, 部门表 B 
              Where a.Id = n_前病区id And b.Id = n_前科室id; 
            End If; 
          End If; 
        End If; 
 
        If v_重算站点 Is Not Null Then 
          If v_付款方式价格等级 Is Null Then 
            If Nvl(n_是否用价格等级, 0) = 1 Then 
              v_Temp     := Nvl(Zl_Get_Pricegrade(v_重算站点, 病人id_In, 主页id_In, r_Pati.付款方式), '|||') || '||||'; 
              v_价格等级 := Substr(v_Temp, 1, Instr(v_Temp, '|') - 1); 
            End If; 
          Else 
            v_价格等级 := v_付款方式价格等级; 
          End If; 
        Else 
          If v_付款方式价格等级 Is Null Then 
            If Nvl(n_是否用价格等级, 0) = 1 Then 
              v_Temp     := Nvl(Zl_Get_Pricegrade(r_Pati_Change.站点, 病人id_In, 主页id_In, r_Pati.付款方式), '|||') || '||||'; 
              v_价格等级 := Substr(v_Temp, 1, Instr(v_Temp, '|') - 1); 
            End If; 
          Else 
            v_价格等级 := v_付款方式价格等级; 
          End If; 
        End If; 
       --判断是否手动销帐
       Select Count(1)
       Into n_Count
       From 住院费用记录
       Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 2 And Nvl(附加标志, 0) = 1 And
             收费类别 = Decode(r_Pati_Change.类型, 1, 'H', 2, 'J', 收费类别) And 发生时间 = d_发生时间 And
             收费细目id = Decode(r_Pati_Change.类型, 3, r_Pati_Change.收费细目id, 收费细目id);
       
        If v_价格等级 Is Null Then 
          Open c_价格_Rec For 
            Select b.现价 As 标准单价, b.收入项目id, c.收据费目, m.计算单位, m.类别 
            From 收费价目 B, 收入项目 C, 收费项目目录 M 
            Where b.收费细目id = m.Id And b.收费细目id = r_Pati_Change.收费细目id And b.收入项目id = c.Id And 
                  (c.撤档时间 Is Null Or c.撤档时间 > d_发生时间) And Trunc(d_发生时间) Between Trunc(b.执行日期) And 
                  Nvl(b.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')) And b.价格等级 Is Null; 
        Else 
          Open c_价格_Rec For 
            Select b.现价 As 标准单价, b.收入项目id, c.收据费目, m.计算单位, m.类别 
            From 收费价目 B, 收入项目 C, 收费项目目录 M 
            Where b.收费细目id = m.Id And b.收费细目id = r_Pati_Change.收费细目id And b.收入项目id = c.Id And 
                  (c.撤档时间 Is Null Or c.撤档时间 > d_发生时间) And Trunc(d_发生时间) Between Trunc(b.执行日期) And 
                  Nvl(b.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')) And 
                  (b.价格等级 = v_价格等级 Or 
                  (b.价格等级 Is Null And Not Exists 
                   (Select 1 
                     From 收费价目 
                     Where 收费细目id = r_Pati_Change.收费细目id And 价格等级 = Nvl(v_价格等级, '-') And 
                           Trunc(d_发生时间) Between Trunc(执行日期) And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))))); 
        End If; 
 
        Loop 
          Fetch c_价格_Rec 
            Into n_标准价格, n_收入项目id, v_收据费目, v_计算单位, v_类别; 
          Exit When c_价格_Rec%NotFound; 
          --For c_价格 In c_价格_Rec(n_收费细目id, d_发生时间, v_价格等级) Loop 
          --提取当前收入项目的收费比率 
          n_Exsetax := Get_Discount_Rate(r_Pati.费别, r_Pati_Change.收费细目id, n_收入项目id, Abs(n_标准价格 * r_Pati_Change.数量)); 
 
          --如果已经计算，原记录计算完全正确，则直接修改将标志改正 
          Update 住院费用记录 
          Set 附加标志 = 0, 发药窗口 = r_Pati_Change.类型 
          Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 
                Nvl(加班标志, 0) = Nvl(r_Pati_Change.附加床位, 0) And 病人科室id = r_Pati_Change.科室id And 
                病人病区id = Nvl(r_Pati_Change.病区id, 0) And Nvl(床号, 0) = Nvl(r_Pati_Change.床号, 0) And 
                收费细目id = r_Pati_Change.收费细目id And 收入项目id = n_收入项目id And 发生时间 = d_发生时间 And 
                数次 = r_Pati_Change.数量 * n_Dates And 标准单价 = n_标准价格 And 
                应收金额 = Round(n_标准价格 * r_Pati_Change.数量 * n_Dates, n_Dec) And 
                实收金额 = Round(n_标准价格 * r_Pati_Change.数量 * n_Dates * n_Exsetax / 100, n_Dec); 
                          
          If Sql%RowCount = 0 And n_Count =0 Then 
            --如果未计算或计算错误，则增加正确的计算记录 
            r_变动_Cur.Delete; 
            r_变动_Cur.Extend; 
            For Q In 1 .. r_病人变动.Count Loop 
              If r_病人变动(Q).病区id = r_Pati_Change.病区id And r_病人变动(Q).科室id = r_Pati_Change.科室id And 
                  d_发生时间 - n_Last Between Trunc(r_病人变动(Q).开始时间) And r_病人变动(Q).终止时间 Then 
                r_变动_Cur(r_变动_Cur.Count).Id := r_病人变动(Q).Id; 
                r_变动_Cur(r_变动_Cur.Count).开始时间 := r_病人变动(Q).开始时间; 
                r_变动_Cur(r_变动_Cur.Count).终止时间 := r_病人变动(Q).终止时间; 
                r_变动_Cur(r_变动_Cur.Count).科室id := r_病人变动(Q).科室id; 
                r_变动_Cur(r_变动_Cur.Count).病区id := r_病人变动(Q).病区id; 
                r_变动_Cur(r_变动_Cur.Count).经治医师 := r_病人变动(Q).经治医师; 
                r_变动_Cur(r_变动_Cur.Count).责任护士 := r_病人变动(Q).责任护士; 
                r_变动_Cur(r_变动_Cur.Count).医疗小组id := r_病人变动(Q).医疗小组id; 
                Exit; 
              End If; 
            End Loop; 
 
            If v_Billno Is Null Then 
              v_Billno := Nextno(17); 
            End If; 
            Insert Into 住院费用记录 
              (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 姓名, 性别, 
               年龄, 标识号, 床号, 费别, 记帐费用, 收费细目id, 收入项目id, 附加标志, 标准单价, 付数, 数次, 应收金额, 实收金额, 收费类别, 计算单位, 加班标志, 收据费目, 开单人, 划价人, 
               操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 医疗小组id, 发药窗口) 
              Select 病人费用记录_Id.Nextval, 3, v_Billno, 1, Rownum + n_Billcount, Null, Null, 0, Null, 
                     Decode(r_Pati_Change.主页id, Null, 1, 2), r_Pati_Change.病人id, r_Pati_Change.主页id, r_Pati_Change.病区id, 
                     r_Pati_Change.科室id, r_Pati_Change.科室id, r_Pati_Change.病区id, r_Pati.姓名, r_Pati.性别, r_Pati.年龄, 
                     r_Pati.住院号, r_Pati_Change.床号, r_Pati.费别, 1, r_Pati_Change.收费细目id, n_收入项目id, 0, n_标准价格, 1, 
                     r_Pati_Change.数量 * n_Dates, Round(n_标准价格 * r_Pati_Change.数量 * n_Dates, n_Dec), 
                     Round(n_标准价格 * r_Pati_Change.数量 * n_Dates * n_Exsetax / 100, n_Dec), v_类别, v_计算单位, 
                     r_Pati_Change.附加床位, v_收据费目,r_变动_Cur(1).经治医师,r_变动_Cur(1).责任护士, r_Pati_Change.操作员编号, 
                     r_Pati_Change.操作员姓名, d_发生时间, d_登记时间, Decode(r_Pati_Change.险类, Null, 0, 1), r_Pati_Change.大类id, 
                     Decode(Nvl(n_算法, 0), 1, 
                             Round(n_标准价格 * r_Pati_Change.数量 * n_Dates * n_Exsetax / 100 * n_统筹比额 / 100, n_Dec), 2, n_统筹比额, 
                             0),r_变动_Cur(1).医疗小组id, r_Pati_Change.类型 
              From Dual; 
            n_Billcount := n_Billcount + Sql%RowCount; 
          End If; 
        End Loop; 
        Close c_价格_Rec; 
      End If; 
      r_Pati_Change_Pre := c_自动记帐; 
    End Loop; 
  End Loop; 
 
  ----------------------------------------------------------------- 
  --作废以前计算的错误记录 
  ----------------------------------------------------------------- 
  Insert Into 住院费用记录 
    (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 姓名, 性别, 年龄, 标识号, 
     床号, 费别, 记帐费用, 收费细目id, 收入项目id, 附加标志, 标准单价, 付数, 数次, 应收金额, 实收金额, 收费类别, 计算单位, 加班标志, 收据费目, 开单人, 划价人, 操作员编号, 操作员姓名, 发生时间, 
     登记时间, 保险项目否, 保险大类id, 统筹金额, 医疗小组id, 发药窗口) 
    Select 病人费用记录_Id.Nextval, 记录性质, NO, 2, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 
           姓名, 性别, 年龄, 标识号, 床号, 费别, 记帐费用, 收费细目id, 收入项目id, 0, 标准单价, 付数, -数次, -应收金额, -实收金额, 收费类别, 计算单位, 加班标志, 收据费目, 开单人, 
           划价人, 操作员编号, 操作员姓名, 发生时间, d_登记时间, 保险项目否, 保险大类id, -统筹金额, 医疗小组id, 发药窗口 
    From 住院费用记录 
    Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 附加标志 = 5 And 发生时间 >= d_Start_Date; 
 
  ----------------------------------------------------------------- 
  --填写病人余额 
  ----------------------------------------------------------------- 
  Select Sum(Decode(附加标志, 0, 1, -1) * 实收金额) As 实收金额 
  Into n_Summoney 
  From 住院费用记录 
  Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 
        (NO = v_Billno Or 附加标志 = 5 And 发生时间 >= d_Start_Date); 
 
  Update 病人余额 
  Set 费用余额 = Nvl(费用余额, 0) + Nvl(n_Summoney, 0) 
  Where 病人id = 病人id_In And 性质 = 1 And 类型 = 2 
  Returning 费用余额 Into n_返回值; 
 
  If Sql%RowCount = 0 Then 
    Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (病人id_In, 1, 2, n_Summoney, 0); 
    n_返回值 := n_Summoney; 
  End If; 
 
  If Nvl(n_返回值, 0) = 0 Then 
    Delete From 病人余额 Where 性质 = 1 And 病人id = 病人id_In And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0; 
  End If; 
 
  ----------------------------------------------------------------- 
  --填写病人汇总费用 
  ----------------------------------------------------------------- 
  n_Delete := 0; 
  For v_Currrow In c_Sumcur_Rec(v_Billno, d_Start_Date) Loop 
    Update 病人未结费用 
    Set 金额 = Nvl(金额, 0) + Nvl(v_Currrow.实收金额, 0) 
    Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And Nvl(病人病区id, 0) = Nvl(v_Currrow.病人病区id, 0) And 
          Nvl(病人科室id, 0) = Nvl(v_Currrow.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(v_Currrow.开单部门id, 0) And 
          Nvl(执行部门id, 0) = Nvl(v_Currrow.执行部门id, 0) And Nvl(收入项目id, 0) = Nvl(v_Currrow.收入项目id, 0) And 来源途径 + 0 = 2 
    Returning 金额 Into n_返回值; 
 
    If Sql%RowCount = 0 Then 
      Insert Into 病人未结费用 
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额) 
      Values 
        (病人id_In, 主页id_In, v_Currrow.病人病区id, v_Currrow.病人科室id, v_Currrow.开单部门id, v_Currrow.执行部门id, v_Currrow.收入项目id, 2, 
         v_Currrow.实收金额); 
      n_返回值 := v_Currrow.实收金额; 
    End If; 
    If Nvl(n_返回值, 0) = 0 Then 
      n_Delete := 1; 
    End If; 
  End Loop; 
 
  If Nvl(n_Delete, 0) = 1 Then 
    Delete From 病人未结费用 Where 病人id = 病人id_In And 金额 = 0; 
  End If; 
  ----------------------------------------------------------------- 
  --将所有修改的附加标志还原为正常标志 
  ----------------------------------------------------------------- 
  Update 住院费用记录 
  Set 附加标志 = 0, 记录状态 = 3 
  Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 附加标志 = 5 And 发生时间 >= d_Start_Date; 
  ----------------------------------------------------------------- 
  --修改计算时间标志 
  ----------------------------------------------------------------- 
  Update 病人自动计算 
  Set 上次计算时间 = Greatest(Sysdate, Nvl(终止时间, Sysdate)) 
  Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(终止时间, Sysdate) > d_Start_Date; 
Exception 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl1_Autocalc_Pati_Charge_Nm;
/

--123419:王煜,2018-05-10,新增手工销自动记帐费用附加标志含义
CREATE OR REPLACE Procedure Zl1_Autocalc_Pati_Charge
( 
  病人id_In       In Number, 
  主页id_In       In Number, 
  期间_In         In Varchar2, 
  强制记帐_In     In Number := 0, 
  启用价格等级_In In Number := -1 
) As 
  ------------------------------------------------------------------------- 
  --功能说明：完成指定病人指定期间的自动记帐 
  --          1、系统首先根据系统参数"修正上期自动计费"，修改以往该病人自动记帐记录标志; 
  --          2、综合病人的床位变化、入出转情况、调价情况等多项因素，结合期间跨度、病人费 
  --             别等完成费用的正确计算： 
  --             如果发现已经计算，则修改标志为正常;如果未计算，则插入新的自动记帐记录; 
  --             作废以前的错误计算的记录; 
  --             统计本次变动(新增和作废)，填写余额表和汇总表; 
  --入口参数： 
  --       病人ID_IN  number    病人身份ID 
  --       主页ID_IN  number    病案主页ID，两个参数共同确定需要计算的病人 
  --       期间_IN  varchar2     需要计算的最小期间 
  --       强制记帐_IN number   为1时,不受病案主页.禁止自动记帐属性控制 
  --       启用价格等级_In number ：-1表示未判断价格等级,内部会自动去检查;0-不启用价格等级;1-启用了价格等级的 
  --调用关系：zl1_AutoCptPati/zl1_AutoCptWard/zl1_AutoCptAll 调用本过程 
  ------------------------------------------------------------------------- 
  v_价格等级         收费价格等级.名称%Type; 
  v_站点_Pre         收费价格等级.名称%Type; 
  v_付款方式价格等级 收费价格等级.名称%Type; 
 
  v_Temp      Varchar2(500); 
  v_Billno    Varchar2(8); --费用表实际的自动记帐号码 
  n_Billcount Number(5) := 0; --单据序号计数器 
 
  n_Exsetax  Number(16, 2) := 0; --费用收取比率 
  n_Summoney Number(16, 2) := 0; --金额 
 
  n_Dec        Number; --金额小数位数 
  n_Dates      Number(6, 1); --当前记录的天数，全天为1 
  n_Delete     Number; 
  n_返回值     病人余额.预交余额%Type; 
  n_收费细目id 收费项目目录.Id%Type; 
 
  n_住院状态       病案主页.状态%Type; 
  n_病人审核方式   Number(2); 
  n_未入科禁止记账 Number(2); 
  n_Count          Number(5);
  n_床位半天模式 Number(2); 
  n_护理半天模式 Number(2); 
  n_其他半天模式 Number(2); 
 
  n_床位价格优先   Number(2); 
  n_护理价格优先   Number(2); 
  n_其他项价格优先 Number(2); 
  n_是否用价格等级 Number(2); --0-未启用;1-启用 
  n_是否计算费用   Number(2); 
  n_类型           Number(2); --1-护理;2- 床位;3-其他 
  n_Finded         Number(2); 
  v_半天模式       Varchar2(50); 
  n_开单部门id     病人变动记录.科室id%Type; 
  n_病人病区id     病人变动记录.病区id%Type; 
  n_检查类型       Number(3); 
 
  d_Start_Date Date; 
  d_登记时间   Date; --登记时间 
  d_发生时间   Date; --发生时间 
  d_Temp       Date; 
 
  l_Mulit_细目id t_Numlist := t_Numlist(); 
 
  Type t_病人变动_Rec Is Record( 
    ID         病人变动记录.Id%Type, 
    开始时间   病人变动记录.开始时间%Type, 
    终止时间   病人变动记录.开始时间%Type, 
    科室id     病人变动记录.科室id%Type, 
    病区id     病人变动记录.病区id%Type, 
    经治医师   病人变动记录.经治医师%Type, 
    责任护士   病人变动记录.责任护士%Type, 
    医疗小组id 病人变动记录.医疗小组id%Type); 
 
  Type c_病人变动_Rec Is Table Of t_病人变动_Rec; 
  r_病人变动 c_病人变动_Rec := c_病人变动_Rec(); 
  r_变动_Cur c_病人变动_Rec := c_病人变动_Rec(); 
 
  Cursor c_Sumcur_Rec 
  ( 
    Billno    Varchar2, 
    Datestart Date 
  ) Is 
    Select 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, Sum(Decode(附加标志, 0, 1, -1) * 应收金额) As 应收金额, 
           Sum(Decode(附加标志, 0, 1, -1) * 实收金额) As 实收金额 
    From 住院费用记录 
    Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 
          (NO = Billno Or 附加标志 = 5 And 发生时间 >= Datestart) 
    Group By 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id; 
 
  n_收入项目id 收入项目.Id%Type; 
  v_收据费目   收入项目.收据费目%Type; 
  v_计算单位   收费项目目录.计算单位%Type; 
  v_类别       收费项目目录.类别%Type; 
  n_标准价格   收费价目.现价%Type; 
 
  n_算法     保险支付大类.算法%Type; 
  n_统筹比额 保险支付大类.统筹比额%Type; 
 
  Type t_价格_Rec Is Ref Cursor; 
  c_价格_Rec t_价格_Rec; 
 
  Cursor c_Pati Is 
    Select a.病人id, a.主页id, Nvl(a.姓名, i.姓名) As 姓名, Nvl(a.性别, i.性别) As 性别, Nvl(a.年龄, i.年龄) As 年龄, Nvl(a.住院号, i.住院号) As 住院号, 
           a.费别, Nvl(a.险类, 0) As 险类, Nvl(a.审核标志, 0) As 审核标志, Nvl(a.状态, 0) As 住院状态, Nvl(a.是否禁止自动记帐, 0) As 是否禁止自动记帐, 
           a.医疗付款方式 As 付款方式 
    From 病案主页 A, 病人信息 I 
    Where a.病人id = i.病人id And a.病人id = 病人id_In And a.主页id = 主页id_In; 
 
  r_Pati c_Pati%RowType; 
 
  Function Get_Discount_Rate 
  ( 
    费别_In       病人信息.费别%Type, 
    收费细目id_In 费别明细.收费细目id%Type, 
    收入项目id_In 费别明细.收入项目id%Type, 
    金额_In       费别明细.应收段首值%Type 
  ) Return Number As 
    n_Discount_Rate Number(16, 5); 
  Begin 
    Begin 
      Select 实收比率 
      Into n_Discount_Rate 
      From (Select 实收比率 
             From 费别明细 
             Where 费别 = Nvl(费别_In, '-') And 收费细目id = Nvl(收费细目id_In, 0) And (金额_In Between 应收段首值 And 应收段尾值) 
             Union All 
             Select 实收比率 
             From 费别明细 
             Where 费别 = Nvl(费别_In, '-') And 收入项目id = Nvl(收入项目id_In, 0) And (金额_In Between 应收段首值 And 应收段尾值) And Not Exists 
              (Select 1 From 费别明细 Where 费别 = Nvl(费别_In, '-') And 收费细目id = Nvl(收费细目id_In, 0))); 
    Exception 
      When Others Then 
        n_Discount_Rate := 100.00; 
    End; 
    n_Discount_Rate := Nvl(n_Discount_Rate, 100); 
    Return n_Discount_Rate; 
  End Get_Discount_Rate; 
 
Begin 
 
  --获取病人信息 
  Begin 
    Open c_Pati; 
    Fetch c_Pati 
      Into r_Pati; 
  Exception 
    When Others Then 
      Return; 
  End; 
  If Nvl(强制记帐_In, 0) = 0 And Nvl(r_Pati.是否禁止自动记帐, 0) = 1 Then 
    Return; 
  End If; 
 
  n_病人审核方式   := Nvl(zl_GetSysParameter(185), 0); 
  n_未入科禁止记账 := Nvl(zl_GetSysParameter(215), 0); 
 
  If n_病人审核方式 = 1 And Nvl(r_Pati.审核标志, 0) >= 1 Then 
    Return; 
  End If; 
 
  If n_未入科禁止记账 = 1 And n_住院状态 = 1 Then 
    Return; 
  End If; 
 
  -------------------------------------------------------------------------------- 
  --1.初始化相关的参数 
  v_半天模式 := Nvl(zl_GetSysParameter(100), '0'); 
  If Length(v_半天模式) = 3 Then 
    n_床位半天模式 := To_Number(Substr(v_半天模式, 1, 1)); 
    n_护理半天模式 := To_Number(Substr(v_半天模式, 2, 1)); 
    n_其他半天模式 := To_Number(Substr(v_半天模式, 3, 1)); 
  Else 
    n_床位半天模式 := To_Number(v_半天模式); 
    n_护理半天模式 := To_Number(v_半天模式); 
    n_其他半天模式 := To_Number(v_半天模式); 
  End If; 
 
  n_床位价格优先   := 0; 
  n_其他项价格优先 := 0; 
 
  --金额小数位数 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(160), '0')) 
  Into n_Dec, n_护理价格优先 
  From Dual; 
 
  n_是否用价格等级 := 启用价格等级_In; 
  If n_是否用价格等级 < 0 Then 
    Select Nvl(Max(1), 0) Into n_是否用价格等级 From 收费价格等级应用 Where Rownum < 2; 
  End If; 
  --每天5点以前，将记录时间登记为昨天，否则登记为当时 
  Select Decode(Sign(To_Number(To_Char(Sysdate, 'HH24')) - 5), -1, Trunc(Sysdate) - 1 / 24 / 60, Sysdate) 
  Into d_登记时间 
  From Dual; 
 
  v_付款方式价格等级 := Null; 
  If Nvl(n_是否用价格等级, 0) = 1 Then 
    Select Max(价格等级) 
    Into v_付款方式价格等级 
    From 收费价格等级应用 A, 收费价格等级 B 
    Where a.价格等级 = b.名称 And a.性质 = 1 And a.医疗付款方式 = Nvl(r_Pati.付款方式, '-') And Nvl(b.是否适用普通项目, 0) = 1 And 
          Nvl(b.撤档时间, Sysdate + 1) > Sysdate; 
  End If; 
  -------------------------------------------------------------------------------- 
 
  --锁定该病人的记录,以免重复计算 
  Update 病案主页 Set 状态 = 状态 Where 病人id = 病人id_In And 主页id = 主页id_In; 
 
  -------------------------------------------------------------------------------- 
  --2. 先将变动信息给记录集,以便提取经治医师和责任护士 
  For c_变动 In (Select ID, 开始时间, Nvl(终止时间, Trunc(Sysdate) + 1) As 终止时间, 科室id, 病区id, 经治医师, 责任护士, 医疗小组id 
               From 病人变动记录 A 
               Where 病人id = 病人id_In And 主页id = 主页id_In And 科室id Is Not Null And 
                     Nvl(终止时间, Sysdate) >= (Select Nvl(Min(上次计算时间), Sysdate - 1000) 
                                            From 病人自动计算 
                                            Where 病人id = 病人id_In And 主页id = 主页id_In) 
               Order By 病区id, 科室id, 开始时间 Desc) Loop 
    r_病人变动.Extend; 
    r_病人变动(r_病人变动.Count).Id := c_变动.Id; 
    r_病人变动(r_病人变动.Count).开始时间 := c_变动.开始时间; 
    r_病人变动(r_病人变动.Count).终止时间 := c_变动.终止时间; 
    r_病人变动(r_病人变动.Count).科室id := c_变动.科室id; 
    r_病人变动(r_病人变动.Count).病区id := c_变动.病区id; 
    r_病人变动(r_病人变动.Count).经治医师 := c_变动.经治医师; 
    r_病人变动(r_病人变动.Count).责任护士 := c_变动.责任护士; 
    r_病人变动(r_病人变动.Count).医疗小组id := c_变动.医疗小组id; 
  End Loop; 
 
  ----------------------------------------------------------------- 
  --循环检查计算情况，并增加正确和新计算的记录 
  ----------------------------------------------------------------- 
  d_Start_Date := Sysdate + 1000; 
  d_Temp       := Sysdate - 1000; 
 
  --1.计算床位费 
  For c_自动记帐 In (Select a.类型, a.病人id, a.主页id, a.科室id, a.病区id, a.床号, a.附加床位, a.收费细目id, a.操作员编号, a.操作员姓名, a.开始时间, 
                        Nvl(a.终止时间, Sysdate) As 终止时间, a.启用日期, a.数量, Greatest(a.开始日期, Trunc(p.开始日期)) As 开始日期, 
                        Nvl(a.终止日期, Trunc(Sysdate)) As 终止日期, 
                        Nvl(a.终止日期, Trunc(Sysdate)) - Greatest(a.开始日期, Trunc(p.开始日期)) As 天数, Nvl(Q1.站点, Q2.站点) As 站点, 
                        m.计算单位, m.类别, i.险类, i.大类id, k.算法, k.统筹比额, 
                        Nvl(m.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 项目撤档时间, a.计算标志, a.Id 
                 From (Select b.Id, 2 As 类型, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, b.床位等级id As 收费细目id, b.操作员编号, 
                               b.操作员姓名, b.开始时间, b.终止时间, a.启用日期, b.数量, Zl_Date_Half(b.开始时间, n_床位半天模式) As 开始日期, 
                               Zl_Date_Half(b.终止时间, n_床位半天模式) As 终止日期, 0 As 计算标志 
                        From 自动计价项目 A, 
                             (Select a.Id, a.病人id, a.主页id, a.开始时间, a.附加床位, a.病区id, a.科室id, a.床号, a.床位等级id, 1 As 数量, a.终止时间, 
                                      a.操作员编号, a.操作员姓名, a.上次计算时间 
                               From 病人自动计算 A 
                               Where a.性质 = 2 And a.病人id = 病人id_In And a.主页id = 主页id_In And 
                                     Nvl(a.上次计算时间, a.开始时间) <= Nvl(a.终止时间, Sysdate) 
                               Union All 
                               Select b.Id, b.病人id, b.主页id, 开始时间, 附加床位, b.病区id, b.科室id, 床号, i.从项id As 床位等级id, i.从项数次 As 数量, 
                                      b.终止时间, b.操作员编号, b.操作员姓名, b.上次计算时间 
                               From 病人自动计算 B, 收费从属项目 I 
                               Where b.性质 = 2 And b.病人id = 病人id_In And b.主页id = 主页id_In And b.床位等级id = i.主项id And i.固有从属 > 0 And 
                                     Nvl(b.上次计算时间, b.开始时间) <= Nvl(b.终止时间, Sysdate)) B 
                        Where a.病区id = b.病区id And a.计算标志 = 1 
                        Union All 
                        Select b.Id, 1 As 类型, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, b.护理等级id As 收费细目id, b.操作员编号, 
                               b.操作员姓名, b.开始时间, b.终止时间, a.启用日期, b.数量, Zl_Date_Half(b.开始时间, n_护理半天模式) As 开始日期, 
                               Zl_Date_Half(b.终止时间, n_护理半天模式) As 终止日期, 0 As 计算标志 
                        From 自动计价项目 A, 
                             (Select a.Id, a.病人id, a.主页id, a.开始时间, a.附加床位, a.病区id, a.科室id, a.床号, a.护理等级id, 1 As 数量, a.终止时间, 
                                      a.操作员编号, a.操作员姓名, a.上次计算时间 
                               From 病人自动计算 A 
                               Where a.性质 = 1 And a.病人id = 病人id_In And a.主页id = 主页id_In And 
                                     Nvl(a.上次计算时间, a.开始时间) <= Nvl(a.终止时间, Sysdate) 
                               Union All 
                               Select b.Id, b.病人id, b.主页id, 开始时间, 附加床位, b.病区id, b.科室id, 床号, i.从项id As 床位等级id, i.从项数次 As 数量, 
                                      b.终止时间, b.操作员编号, b.操作员姓名, b.上次计算时间 
                               From 病人自动计算 B, 收费从属项目 I 
                               Where b.性质 = 1 And b.病人id = 病人id_In And b.主页id = 主页id_In And b.护理等级id = i.主项id And i.固有从属 > 0 And 
                                     Nvl(b.上次计算时间, b.开始时间) <= Nvl(b.终止时间, Sysdate)) B 
                        Where a.病区id = b.病区id And a.计算标志 = 2 
                        Union All 
                        Select b.Id, 3 As 类型, b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, a.收费细目id, b.操作员编号, b.操作员姓名, 
                               b.开始时间, b.终止时间, a.启用日期, a.数量, Zl_Date_Half(b.开始时间, n_其他半天模式) As 开始日期, 
                               Zl_Date_Half(b.终止时间, n_其他半天模式) As 终止日期, a.计算标志 
                        From (Select 病区id, 计算标志, 收费细目id, 1 As 数量, 启用日期 
                               From 自动计价项目 
                               Union All 
                               Select 病区id, 计算标志, 从项id, i.从项数次 As 数量, 启用日期 
                               From 自动计价项目 A, 收费从属项目 I 
                               Where a.收费细目id = i.主项id And i.固有从属 > 0) A, 病人自动计算 B 
                        Where a.病区id = b.病区id And b.病人id = 病人id_In And b.主页id = 主页id_In And b.性质 = 3 And 
                              Nvl(b.附加床位, 0) = 0 And (a.计算标志 = 6 And b.床位等级id Is Not Null Or a.计算标志 = 7) And 
                              Nvl(b.上次计算时间, b.开始时间) <= Nvl(b.终止时间, Sysdate)) A, 
                      (Select Min(开始日期) As 开始日期 From 期间表 Where 期间 >= 期间_In) P, 保险支付项目 I, 保险支付大类 K, 收费项目目录 M, 部门表 Q1, 
                      部门表 Q2 
                 Where Trunc(Nvl(a.终止时间, Greatest(a.开始时间, Sysdate))) >= Trunc(p.开始日期) And a.收费细目id = i.收费细目id(+) And 
                       i.险类(+) = Nvl(r_Pati.险类, 0) And i.大类id = k.Id(+) And a.收费细目id = m.Id And a.病区id = Q1.Id(+) And 
                       a.科室id = Q2.Id(+) 
                 Order By 类型, 开始时间) Loop 
    --产生数据 
    If v_付款方式价格等级 Is Null Then 
      If Nvl(n_是否用价格等级, 0) = 1 And Nvl(v_站点_Pre, '-') <> Nvl(c_自动记帐.站点, '-') Then 
        v_Temp     := Nvl(Zl_Get_Pricegrade(c_自动记帐.站点, 病人id_In, 主页id_In, r_Pati.付款方式), '|||') || '||||'; 
        v_价格等级 := Substr(v_Temp, 1, Instr(v_Temp, '|') - 1); 
      End If; 
    Else 
      v_价格等级 := v_付款方式价格等级; 
    End If; 
 
    v_站点_Pre := c_自动记帐.站点; 
 
    If d_Start_Date > c_自动记帐.开始日期 Then 
      d_Start_Date := c_自动记帐.开始日期; 
    End If; 
 
    If Nvl(c_自动记帐.类型, 0) <> Nvl(n_检查类型, 0) Then 
      n_检查类型 := Nvl(c_自动记帐.类型, 0); 
      Update 住院费用记录 
      Set 附加标志 = 5 
      Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 附加标志 <> 8 And Nvl(医嘱序号, 0) = 0 And 
            发生时间 >= c_自动记帐.开始日期 And 附加标志 <> 5 And (发药窗口 Is Null Or 发药窗口 = Nvl(n_检查类型, 0)); 
    End If; 
 
    --生成每天费用 
    For I In 0 .. c_自动记帐.天数 Loop 
      d_发生时间 := Greatest(c_自动记帐.开始日期, Trunc(c_自动记帐.开始日期 + I)); 
      n_Dates    := Least(Trunc(c_自动记帐.开始日期 + I + 1), c_自动记帐.终止日期) - Greatest(c_自动记帐.开始日期, Trunc(c_自动记帐.开始日期 + I)); 
      If n_Dates < 0 Then 
        n_Dates := 0; 
      End If; 
      If (n_护理价格优先 = 1 And c_自动记帐.类型 = 1) Or (n_床位价格优先 = 1 And c_自动记帐.类型 = 2) Or (n_其他项价格优先 = 1 And c_自动记帐.类型 = 3) Then 
 
        If d_发生时间 <> d_Temp Or Nvl(n_类型, 0) <> Nvl(c_自动记帐.类型, 0) Then 
          l_Mulit_细目id.Delete; 
          d_Temp := d_发生时间; 
          n_类型 := Nvl(c_自动记帐.类型, 0); 
        End If; 
 
        n_Finded := 0; 
        For J In 1 .. l_Mulit_细目id.Count Loop 
          If l_Mulit_细目id(J) = c_自动记帐.收费细目id Then 
            n_Finded := 1; 
            Exit; 
          End If; 
        End Loop; 
        If n_Finded = 0 Then 
          l_Mulit_细目id.Extend; 
          l_Mulit_细目id(l_Mulit_细目id.Count) := c_自动记帐.收费细目id; 
        End If; 
      End If; 
 
      n_是否计算费用 := 1; 
      If d_发生时间 > c_自动记帐.项目撤档时间 Or n_Dates = 0 Then 
        n_是否计算费用 := 0; 
      End If; 
 
      Select Decode(Nvl(Max(1), 0), 0, n_是否计算费用, 0) 
      Into n_是否计算费用 
      From 病人自动计算 A 
      Where a.终止原因 = 1 And a.Id = c_自动记帐.Id And Exists 
       (Select 1 
             From 病人自动计算 
             Where 病人id = a.病人id And 主页id = a.主页id And 开始原因 = 2 And Trunc(开始时间) = Trunc(a.终止时间)); 
 
      If n_是否计算费用 = 1 Then 
        --需要检查是否在指定日期被停用了的 
        n_收费细目id := c_自动记帐.收费细目id; 
        n_算法       := c_自动记帐.算法; 
        n_统筹比额   := c_自动记帐.统筹比额; 
 
        If (n_护理价格优先 = 1 And c_自动记帐.类型 = 1) Or (n_床位价格优先 = 1 And c_自动记帐.类型 = 2) Or 
           (n_其他项价格优先 = 1 And c_自动记帐.类型 = 3) And l_Mulit_细目id.Count > 1 Then 
          --取最高价格的收费项目 
          If v_价格等级 Is Null Then 
            Open c_价格_Rec For 
              Select b.收费细目id, Sum(b.现价) As 标准单价 
              From 收费价目 B, 收入项目 C 
              Where b.收费细目id In (Select Column_Value From Table(l_Mulit_细目id)) And b.收入项目id = c.Id And 
                    (c.撤档时间 Is Null Or c.撤档时间 > d_发生时间) And Trunc(d_发生时间) Between Trunc(b.执行日期) And 
                    Nvl(b.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')) And b.价格等级 Is Null 
 
              Group By 收费细目id 
              Order By 标准单价 Desc; 
          Else 
            Open c_价格_Rec For 
              Select b.收费细目id, Sum(b.现价) As 标准单价 
              From 收费价目 B, 收入项目 C 
              Where b.收费细目id In (Select Column_Value From Table(l_Mulit_细目id)) And b.收入项目id = c.Id And 
                    (c.撤档时间 Is Null Or c.撤档时间 > d_发生时间) And Trunc(d_发生时间) Between Trunc(b.执行日期) And 
                    Nvl(b.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')) And 
                    (b.价格等级 = Nvl(v_价格等级, '-') Or 
                    (b.价格等级 Is Null And Not Exists 
                     (Select 1 
                       From 收费价目 
                       Where 收费细目id = b.收费细目id And 价格等级 = Nvl(v_价格等级, '-') And Trunc(d_发生时间) Between Trunc(执行日期) And 
                             Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))))) 
              Group By 收费细目id 
              Order By 标准单价 Desc; 
          End If; 
 
          Begin 
            Fetch c_价格_Rec 
              Into n_收费细目id, n_标准价格; 
          Exception 
            When Others Then 
              n_收费细目id := c_自动记帐.收费细目id; 
          End; 
          Close c_价格_Rec; 
        End If; 
        n_开单部门id := c_自动记帐.科室id; 
        n_病人病区id := c_自动记帐.病区id; 
        If c_自动记帐.收费细目id <> n_收费细目id Then 
          --最高价格的收费细目不对，可能统筹比额不一样 
          Select Max(k.算法), Max(k.统筹比额) 
          Into n_算法, n_统筹比额 
          From 保险支付项目 I, 保险支付大类 K 
          Where i.收费细目id = n_收费细目id And i.险类(+) = Nvl(r_Pati.险类, 0) And i.大类id = k.Id(+); 
          If n_护理价格优先 = 1 And c_自动记帐.类型 = 1 Then 
            For c_变动记录 In (Select 病区id, 科室id 
                           From 病人变动记录 
                           Where 开始原因 <> 10 And 病人id = 病人id_In And 主页id = 主页id_In And 护理等级id + 0 = n_收费细目id And 
                                 (Trunc(开始时间) = Trunc(d_发生时间) Or Trunc(Nvl(终止时间, Sysdate)) = Trunc(d_发生时间)) 
                           Order By 开始时间 Desc) Loop 
              n_开单部门id := c_变动记录.科室id; 
              n_病人病区id := c_变动记录.病区id; 
              Exit; 
            End Loop; 
          End If; 
        End If; 
      End If; 
      --判断是否手工销帐
      Select Count(1)
      Into n_Count
      From 住院费用记录
      Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 2 And Nvl(附加标志, 0) = 1 And
           收费类别 = Decode(c_自动记帐.类型, 1, 'H', 2, 'J', 收费类别) And 发生时间 = d_发生时间 And
           收费细目id = Decode(c_自动记帐.类型, 3, c_自动记帐.收费细目id, 收费细目id); 

      If n_是否计算费用 = 1 Then 
        If v_价格等级 Is Null Then 
          Open c_价格_Rec For 
            Select b.现价 As 标准单价, b.收入项目id, c.收据费目, m.计算单位, m.类别 
            From 收费价目 B, 收入项目 C, 收费项目目录 M 
            Where b.收费细目id = m.Id And b.收费细目id = n_收费细目id And b.收入项目id = c.Id And (c.撤档时间 Is Null Or c.撤档时间 > d_发生时间) And 
                  Trunc(d_发生时间) Between Trunc(b.执行日期) And Nvl(b.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')) And 
                  b.价格等级 Is Null; 
        Else 
          Open c_价格_Rec For 
            Select b.现价 As 标准单价, b.收入项目id, c.收据费目, m.计算单位, m.类别 
            From 收费价目 B, 收入项目 C, 收费项目目录 M 
            Where b.收费细目id = m.Id And b.收费细目id = n_收费细目id And b.收入项目id = c.Id And (c.撤档时间 Is Null Or c.撤档时间 > d_发生时间) And 
                  Trunc(d_发生时间) Between Trunc(b.执行日期) And Nvl(b.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')) And 
                  (b.价格等级 = v_价格等级 Or 
                  (b.价格等级 Is Null And Not Exists 
                   (Select 1 
                     From 收费价目 
                     Where 收费细目id = n_收费细目id And 价格等级 = Nvl(v_价格等级, '-') And Trunc(d_发生时间) Between Trunc(执行日期) And 
                           Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))))); 
        End If; 
 
        Loop 
          Fetch c_价格_Rec 
            Into n_标准价格, n_收入项目id, v_收据费目, v_计算单位, v_类别; 
          Exit When c_价格_Rec%NotFound; 
          --For c_价格 In c_价格_Rec(n_收费细目id, d_发生时间, v_价格等级) Loop 
          --提取当前收入项目的收费比率 
          n_Exsetax := Get_Discount_Rate(r_Pati.费别, n_收费细目id, n_收入项目id, Abs(n_标准价格 * c_自动记帐.数量)); 
 
          --如果已经计算，原记录计算完全正确，则直接修改将标志改正 
          Update 住院费用记录 
          Set 附加标志 = 0 
          Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And Nvl(加班标志, 0) = Nvl(c_自动记帐.附加床位, 0) And 
                病人科室id = n_开单部门id And 病人病区id = Nvl(n_病人病区id, 0) And Nvl(床号, 0) = Nvl(c_自动记帐.床号, 0) And 
                收费细目id = n_收费细目id And 收入项目id = n_收入项目id And 发生时间 = d_发生时间 And 数次 = c_自动记帐.数量 * n_Dates And 
                标准单价 = n_标准价格 And 应收金额 = Round(n_标准价格 * c_自动记帐.数量 * n_Dates, n_Dec) And 
                实收金额 = Round(n_标准价格 * c_自动记帐.数量 * n_Dates * n_Exsetax / 100, n_Dec); 
 
          If Sql%RowCount = 0 And n_Count=0 Then 
            --如果未计算或计算错误，则增加正确的计算记录 
            r_变动_Cur.Delete; 
            r_变动_Cur.Extend; 
            For Q In 1 .. r_病人变动.Count Loop 
              If r_病人变动(Q) 
               .病区id = c_自动记帐.病区id And r_病人变动(Q).科室id = c_自动记帐.科室id And d_发生时间 Between Trunc(r_病人变动(Q).开始时间) And r_病人变动(Q).终止时间 Then 
                r_变动_Cur(r_变动_Cur.Count).Id := r_病人变动(Q).Id; 
                r_变动_Cur(r_变动_Cur.Count).开始时间 := r_病人变动(Q).开始时间; 
                r_变动_Cur(r_变动_Cur.Count).终止时间 := r_病人变动(Q).终止时间; 
                r_变动_Cur(r_变动_Cur.Count).科室id := r_病人变动(Q).科室id; 
                r_变动_Cur(r_变动_Cur.Count).病区id := r_病人变动(Q).病区id; 
                r_变动_Cur(r_变动_Cur.Count).经治医师 := r_病人变动(Q).经治医师; 
                r_变动_Cur(r_变动_Cur.Count).责任护士 := r_病人变动(Q).责任护士; 
                r_变动_Cur(r_变动_Cur.Count).医疗小组id := r_病人变动(Q).医疗小组id; 
                Exit; 
              End If; 
            End Loop; 
 
            If v_Billno Is Null Then 
              v_Billno := Nextno(17); 
            End If; 
            Insert Into 住院费用记录 
              (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 姓名, 性别, 
               年龄, 标识号, 床号, 费别, 记帐费用, 收费细目id, 收入项目id, 附加标志, 标准单价, 付数, 数次, 应收金额, 实收金额, 收费类别, 计算单位, 加班标志, 收据费目, 开单人, 划价人, 
               操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 医疗小组id, 发药窗口) 
              Select 病人费用记录_Id.Nextval, 3, v_Billno, 1, Rownum + n_Billcount, Null, Null, 0, Null, 
                     Decode(c_自动记帐.主页id, Null, 1, 2), c_自动记帐.病人id, c_自动记帐.主页id, c_自动记帐.病区id, n_开单部门id, n_开单部门id, 
                     n_病人病区id, r_Pati.姓名, r_Pati.性别, r_Pati.年龄, r_Pati.住院号, c_自动记帐.床号, r_Pati.费别, 1, n_收费细目id, n_收入项目id, 
                     0, n_标准价格, 1, c_自动记帐.数量 * n_Dates, Round(n_标准价格 * c_自动记帐.数量 * n_Dates, n_Dec), 
                     Round(n_标准价格 * c_自动记帐.数量 * n_Dates * n_Exsetax / 100, n_Dec), v_类别, v_计算单位, c_自动记帐.附加床位, v_收据费目, 
                     r_变动_Cur(1).经治医师,r_变动_Cur(1).责任护士, c_自动记帐.操作员编号, c_自动记帐.操作员姓名, d_发生时间, d_登记时间, 
                     Decode(c_自动记帐.险类, Null, 0, 1), c_自动记帐.大类id, 
                     Decode(Nvl(n_算法, 0), 1, Round(n_标准价格 * c_自动记帐.数量 * n_Dates * n_Exsetax / 100 * n_统筹比额 / 100, n_Dec), 
                             2, n_统筹比额, 0),r_变动_Cur(1).医疗小组id, n_检查类型 
              From Dual; 
            n_Billcount := n_Billcount + Sql%RowCount; 
          End If; 
        End Loop; 
        Close c_价格_Rec; 
      End If; 
    End Loop; 
  End Loop; 
 
  ----------------------------------------------------------------- 
  --作废以前计算的错误记录 
  ----------------------------------------------------------------- 
  Insert Into 住院费用记录 
    (ID, 记录性质, NO, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 姓名, 性别, 年龄, 标识号, 
     床号, 费别, 记帐费用, 收费细目id, 收入项目id, 附加标志, 标准单价, 付数, 数次, 应收金额, 实收金额, 收费类别, 计算单位, 加班标志, 收据费目, 开单人, 划价人, 操作员编号, 操作员姓名, 发生时间, 
     登记时间, 保险项目否, 保险大类id, 统筹金额, 医疗小组id, 发药窗口) 
    Select 病人费用记录_Id.Nextval, 记录性质, NO, 2, 序号, 从属父号, 价格父号, 多病人单, 医嘱序号, 门诊标志, 病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 
           姓名, 性别, 年龄, 标识号, 床号, 费别, 记帐费用, 收费细目id, 收入项目id, 0, 标准单价, 付数, -数次, -应收金额, -实收金额, 收费类别, 计算单位, 加班标志, 收据费目, 开单人, 
           划价人, 操作员编号, 操作员姓名, 发生时间, d_登记时间, 保险项目否, 保险大类id, -统筹金额, 医疗小组id, 发药窗口 
    From 住院费用记录 
    Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 附加标志 = 5 And 发生时间 >= d_Start_Date; 
 
  ----------------------------------------------------------------- 
  --填写病人余额 
  ----------------------------------------------------------------- 
  Select Sum(Decode(附加标志, 0, 1, -1) * 实收金额) As 实收金额 
  Into n_Summoney 
  From 住院费用记录 
  Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 
        (NO = v_Billno Or 附加标志 = 5 And 发生时间 >= d_Start_Date); 
 
  Update 病人余额 
  Set 费用余额 = Nvl(费用余额, 0) + Nvl(n_Summoney, 0) 
  Where 病人id = 病人id_In And 性质 = 1 And 类型 = 2 
  Returning 费用余额 Into n_返回值; 
 
  If Sql%RowCount = 0 Then 
    Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (病人id_In, 1, 2, n_Summoney, 0); 
    n_返回值 := n_Summoney; 
  End If; 
 
  If Nvl(n_返回值, 0) = 0 Then 
    Delete From 病人余额 Where 性质 = 1 And 病人id = 病人id_In And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0; 
  End If; 
 
  ----------------------------------------------------------------- 
  --填写病人汇总费用 
  ----------------------------------------------------------------- 
  n_Delete := 0; 
  For v_Currrow In c_Sumcur_Rec(v_Billno, d_Start_Date) Loop 
    Update 病人未结费用 
    Set 金额 = Nvl(金额, 0) + Nvl(v_Currrow.实收金额, 0) 
    Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And Nvl(病人病区id, 0) = Nvl(v_Currrow.病人病区id, 0) And 
          Nvl(病人科室id, 0) = Nvl(v_Currrow.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(v_Currrow.开单部门id, 0) And 
          Nvl(执行部门id, 0) = Nvl(v_Currrow.执行部门id, 0) And Nvl(收入项目id, 0) = Nvl(v_Currrow.收入项目id, 0) And 来源途径 + 0 = 2 
    Returning 金额 Into n_返回值; 
 
    If Sql%RowCount = 0 Then 
      Insert Into 病人未结费用 
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额) 
      Values 
        (病人id_In, 主页id_In, v_Currrow.病人病区id, v_Currrow.病人科室id, v_Currrow.开单部门id, v_Currrow.执行部门id, v_Currrow.收入项目id, 2, 
         v_Currrow.实收金额); 
      n_返回值 := v_Currrow.实收金额; 
    End If; 
    If Nvl(n_返回值, 0) = 0 Then 
      n_Delete := 1; 
    End If; 
  End Loop; 
 
  If Nvl(n_Delete, 0) = 1 Then 
    Delete From 病人未结费用 Where 病人id = 病人id_In And 金额 = 0; 
  End If; 
  ----------------------------------------------------------------- 
  --将所有修改的附加标志还原为正常标志 
  ----------------------------------------------------------------- 
  Update 住院费用记录 
  Set 附加标志 = 0, 记录状态 = 3 
  Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 3 And 记录状态 = 1 And 附加标志 = 5 And 发生时间 >= d_Start_Date; 
  ----------------------------------------------------------------- 
  --修改计算时间标志 
  ----------------------------------------------------------------- 
  Update 病人自动计算 
  Set 上次计算时间 = Greatest(Sysdate, Nvl(终止时间, Sysdate)) 
  Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(终止时间, Sysdate) > d_Start_Date; 
Exception 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl1_Autocalc_Pati_Charge;
/

--123419:王煜,2018-05-10,新增手工销自动记帐费用附加标志含义
Create Or Replace Procedure Zl_住院记帐记录_Delete
(
  No_In           住院费用记录.No%Type,
  序号_In         Varchar2,
  操作员编号_In   住院费用记录.操作员编号%Type,
  操作员姓名_In   住院费用记录.操作员姓名%Type,
  记录性质_In     住院费用记录.记录性质%Type := 2,
  操作状态_In     Number := 0,
  输液配药检查_In Number := 1,
  登记时间_In     住院费用记录.登记时间%Type := Sysdate
) As
  --功能：冲销一张住院记帐单据中指定序号行
  --序号：格式如"1,3,5,7,8",或"1:2:33456,3:2,5:2,7:2,8:2",冒号前面的数字表示行号,中间的数字表示退的数量,后面的数字表示配药记录的ID,目前仅在销帐审核时才传入
  --      为空表示冲销所有可冲销行
  --记录性质:    2-人工记帐单,3-自动记帐单
  --输液配药检查:    0-医嘱调用，不检查药品是否进入输液配药中心；1-非医嘱调用，检查药品是否进入配药中心
  --该光标用于销帐指定费用行
  --操作状态_In:0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用
  --该游标为要退费单据的所有原始记录
  Cursor c_Bill Is
    Select a.Id, a.价格父号, a.序号, a.执行状态, a.记录性质, a.收费类别, a.医嘱序号, a.收费细目id, a.病人id, a.主页id, a.收入项目id, a.开单部门id, a.病人科室id,
           a.执行部门id, a.病人病区id, a.付数, a.数次, m.跟踪在用
    From 住院费用记录 A, 材料特性 M
    Where a.No = No_In And a.记录性质 = 记录性质_In And a.记录状态 In (0, 1, 3) And a.门诊标志 = 2 And a.收费细目id + 0 = m.材料id(+)
    Order By 收费细目id, 序号;

  --该游标用于处理药品库存可用数量
  --不要管费用的执行状态,因为先于此步处理
  Cursor c_Stock(v_序号_In Varchar2) Is
    Select ID, 单据, NO, 库房id, 药品id, 批次, 发药方式, 付数, 实际数量, 灭菌效期, 效期, 产地, 批号, 填制日期, 费用id, 商品条码, 内部条码
    From 药品收发记录
    Where NO = No_In And 单据 In (9, 10, 25, 26) And Mod(记录状态, 3) = 1 And 审核人 Is Null And
          费用id In (Select ID
                   From 住院费用记录
                   Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 收费类别 In ('4', '5', '6', '7') And
                         门诊标志 = 2 And (Instr(',' || v_序号_In || ',', ',' || 序号 || ',') > 0 Or v_序号_In Is Null))
    Order By 药品id, 填制日期 Desc;

  r_Stock c_Stock%RowType;

  --该游标用于处理费用记录序号
  Cursor c_Serial Is
    Select 序号, 价格父号
    From 住院费用记录
    Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3)
    Order By 序号;

  v_医嘱id 病人医嘱记录.Id%Type;
  n_划价   Number;
  v_父号   住院费用记录.价格父号%Type;
  v_序号   Varchar2(2000);
  v_Tmp    Varchar2(4000);

  v_医嘱ids Varchar2(4000);
  l_划价    t_Numlist := t_Numlist();
  n_付数    Number;
  n_返回值  Number;
  --部分退费计算变量
  v_剩余数量 Number;
  v_剩余应收 Number;
  v_剩余实收 Number;
  v_剩余统筹 Number;

  v_准退数量 Number;
  v_退费次数 Number;
  v_应收金额 Number;
  v_实收金额 Number;
  v_统筹金额 Number;
  n_部分销帐 Number;
  v_Dec      Number;
  n_Count    Number;
  v_Curdate  Date;
  Err_Item Exception;
  v_Err_Msg        Varchar2(255);
  n_病人id         病案主页.病人id%Type;
  n_主页id         病案主页.主页id%Type;
  n_审核标志       病案主页.审核标志%Type;
  n_住院状态       病案主页.状态%Type;
  n_病人审核方式   Number(2);
  n_未入科禁止记账 Number(2);
  v_配药id         Varchar2(4000);

  n_未执行数量 药品收发记录.实际数量%Type;
  n_已执行数量 药品收发记录.实际数量%Type;
Begin
  --销帐审核时,非药品会传入行号的销帐数量
  If Not 序号_In Is Null Then
    If Instr(序号_In, ':') > 0 Then
      v_Tmp := 序号_In || ',';
      While Not v_Tmp Is Null Loop
        v_序号 := v_序号 || ',' || Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
        If Instr(Substr(v_Tmp, Instr(v_Tmp, ':') + 1, Instr(v_Tmp, ',') - Instr(v_Tmp, ':') - 1), ':') > 0 Then
          v_配药id := v_配药id || ',' ||
                    Substr(v_Tmp, Instr(v_Tmp, ':', 1, 2) + 1, Instr(v_Tmp, ',') - Instr(v_Tmp, ':', 1, 2) - 1);
        End If;
        v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
      End Loop;
      v_序号 := Substr(v_序号, 2);
      If v_配药id Is Not Null Then
        v_配药id := Substr(v_配药id, 2);
      End If;
    Else
      v_序号 := 序号_In;
    End If;
  End If;

  --是否已经全部完全执行(只是整张单据的检查)
  Select Nvl(Count(*), 0), Nvl(Max(病人id), 0), Nvl(Max(主页id), 0)
  Into n_Count, n_病人id, n_主页id
  From 住院费用记录
  Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And Nvl(执行状态, 0) <> 1 And 门诊标志 = 2;
  If n_Count = 0 Then
    v_Err_Msg := '该单据中的项目已经全部完全执行！';
    Raise Err_Item;
  End If;

  n_病人审核方式   := Nvl(zl_GetSysParameter(185), 0);
  n_未入科禁止记账 := Nvl(zl_GetSysParameter(215), 0);
  If n_病人审核方式 = 1 Or n_未入科禁止记账 = 1 Then
  
    Begin
      Select 审核标志, 状态 Into n_审核标志, n_住院状态 From 病案主页 Where 病人id = n_病人id And 主页id = n_主页id;
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

  --未完全执行的项目是否有剩余数量(只是整张单据的检查)
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select 序号, Sum(数量) As 剩余数量
         From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                From 住院费用记录
                Where NO = No_In And 记录性质 = 记录性质_In And 门诊标志 = 2 And
                      Nvl(价格父号, 序号) In
                      (Select Nvl(价格父号, 序号)
                       From 住院费用记录
                       Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And Nvl(执行状态, 0) <> 1)
                Group By 记录状态, Nvl(价格父号, 序号))
         Group By 序号
         Having Sum(数量) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '该单据中未完全执行部分项目剩余数量为零,没有可以销帐的费用！';
    Raise Err_Item;
  End If;

  --医嘱费用：检查正在执行的医嘱(注意已执行的情况在下面检查,因为不传 序号_IN 这种情况费用界面已限制)
  If Nvl(操作状态_In, 0) = 0 Then
    --走销帐申请流程的，不检查医保执行状态
    Select Nvl(Count(*), 0)
    Into n_Count
    From 病人医嘱发送
    Where 执行状态 = 3 And (NO, 记录性质, 医嘱id) In
          (Select NO, 记录性质, 医嘱序号
                        From 住院费用记录
                        Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 医嘱序号 Is Not Null And
                              (Instr(',' || v_序号 || ',', ',' || 序号 || ',') > 0 Or v_序号 Is Null));
    If n_Count > 0 Then
      v_Err_Msg := '要销帐的费用中存在对应的医嘱正在执行的情况，不能销帐！';
      Raise Err_Item;
    End If;
  End If;

  ---------------------------------------------------------------------------------
  --先打开药品对应数据集,以确保当前条件下有数据,为了处理并发判断
  --不能在游标条件中取消"审核人 is Null"条件，因为多次退药可能部份又已发
  Open c_Stock(v_序号);

  --公用变量
  Select 登记时间_In Into v_Curdate From Dual;

  --金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;

  For c_编目病案 In (Select a.姓名
                 From 病人信息 A, 病案主页 B
                 Where a.病人id = b.病人id And b.编目日期 Is Not Null And
                       (b.病人id, b.主页id) In
                       (Select Distinct 病人id, 主页id
                        From 住院费用记录
                        Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 门诊标志 = 2)) Loop
    v_Err_Msg := '病人『' || c_编目病案.姓名 || '』 已经被病案编目,不能被销帐！';
    Raise Err_Item;
  End Loop;
  v_医嘱ids := Null;
  --循环处理每行费用(收入项目行)
  For r_Bill In c_Bill Loop
    --检查已经存在病案编目的,则不能进行销帐处理
    If Instr(',' || v_序号 || ',', ',' || Nvl(r_Bill.价格父号, r_Bill.序号) || ',') > 0 Or v_序号 Is Null Then
      Select Decode(记录状态, 0, 1, 0) Into n_划价 From 住院费用记录 Where ID = r_Bill.Id;
      If Nvl(r_Bill.执行状态, 0) <> 1 Then
        --求剩余数量,剩余应收,剩余实收
        Select Sum(Nvl(付数, 1) * 数次), Sum(应收金额), Sum(实收金额), Sum(统筹金额)
        Into v_剩余数量, v_剩余应收, v_剩余实收, v_剩余统筹
        From 住院费用记录
        Where NO = No_In And 记录性质 = 记录性质_In And 序号 = r_Bill.序号;
        n_部分销帐 := 0;
        If v_剩余数量 = 0 Then
          If v_序号 Is Not Null Then
            v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经全部销帐！';
            Raise Err_Item;
          End If;
          --情况：未限定行号,原始单据中的该笔已经全部销帐(执行状态=0的一种可能)
        Else
        
          If Instr(序号_In, ':') > 0 Then
            v_Tmp := ',' || 序号_In;
            v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',' || r_Bill.序号 || ':') + Length(',' || r_Bill.序号 || ':'));
            v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp || ',', ',') - 1);
            If Instr(v_Tmp, ':') > 0 Then
              v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
            End If;
            v_准退数量 := v_Tmp;
            n_部分销帐 := 1;
          End If;
        
          --准销数量(非药品项目为剩余数量,原始数量)
          If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Then
            If Instr(序号_In, ':') = 0 Or 序号_In Is Null Then
              v_准退数量 := v_剩余数量;
            End If;
          Else
            --医嘱超期收回时,卫材可能没有发放,但申请销帐的是部分数量,所以要以申请的为准
            If Instr(序号_In, ':') = 0 Or 序号_In Is Null Then
              Select Nvl(Sum(Nvl(付数, 1) * 实际数量), 0), Count(*)
              Into v_准退数量, n_Count
              From 药品收发记录
              Where NO = No_In And 单据 In (9, 10, 25, 26) And Mod(记录状态, 3) = 1 And 审核人 Is Null And 费用id = r_Bill.Id;
            End If;
          
            --有剩余数量无准退数量的有两种情况：
            --1.不跟踪在用的卫材无对应的收发记录,这时使用剩余数量
            --2.并发操作,此时已发药或发料
            If v_准退数量 = 0 Then
              If r_Bill.收费类别 = '4' Then
                If n_Count > 0 Or Nvl(r_Bill.跟踪在用, 0) = 1 Then
                  v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发料,须退料后再退费！';
                  Raise Err_Item;
                Else
                  v_准退数量 := v_剩余数量;
                End If;
              Else
                v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发药,须退药后再退费！';
                Raise Err_Item;
              End If;
            End If;
          End If;
        
          --处理住院费用记录
          If Nvl(n_划价, 0) = 0 Then
            --划价时,直接更改数量,所以不须查划冲销次数
            --该笔项目第几次销帐
            Select Nvl(Max(Abs(执行状态)), 0) + 1
            Into v_退费次数
            From 住院费用记录
            Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 = 2 And 序号 = r_Bill.序号 And 门诊标志 = 2;
          End If;
        
          --金额=剩余金额*(准退数/剩余数)
          v_应收金额 := Round(v_剩余应收 * (v_准退数量 / v_剩余数量), v_Dec);
          v_实收金额 := Round(v_剩余实收 * (v_准退数量 / v_剩余数量), v_Dec);
          v_统筹金额 := Round(v_剩余统筹 * (v_准退数量 / v_剩余数量), v_Dec);
          If Nvl(n_划价, 0) = 1 Then
            If Nvl(n_部分销帐, 0) = 0 Then
              l_划价.Extend;
              l_划价(l_划价.Count) := r_Bill.Id;
              n_返回值 := 0;
            Else
              --更新数量
              --划价的,先将相关的数据处理在内部表集中
              n_付数 := 0;
              If r_Bill.付数 > 1 Then
                --如果是中药,超期回收肯定是回收的付数,而不是次数.因此,需要检查准退数量是否可以整 除
                If Trunc(v_准退数量 / r_Bill.数次) <> (v_准退数量 / r_Bill.数次) Then
                  v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用为中药,请按付数进行退费！';
                  Raise Err_Item;
                End If;
                n_付数 := Trunc(v_准退数量 / r_Bill.数次);
                If Nvl(r_Bill.付数, 0) - n_付数 < 0 Then
                  v_准退数量 := r_Bill.数次;
                Else
                  v_准退数量 := 0;
                End If;
              End If;
              Update 住院费用记录
              Set 付数 = 付数 - n_付数, 数次 = 数次 - v_准退数量, 应收金额 = Nvl(应收金额, 0) - v_应收金额, 实收金额 = Nvl(实收金额, 0) - v_实收金额,
                  登记时间 = v_Curdate, 统筹金额 = Nvl(统筹金额, 0) - v_统筹金额
              Where ID = r_Bill.Id
              Returning Nvl(数次, 0) * Nvl(付数, 0) Into n_返回值;
            End If;
            If Nvl(n_返回值, 0) <= 0 Then
              l_划价.Extend;
              l_划价(l_划价.Count) := r_Bill.Id;
            End If;
            If r_Bill.医嘱序号 Is Not Null Then
              If Instr(',' || Nvl(v_医嘱ids, '') || ',', ',' || r_Bill.医嘱序号 || ',') = 0 Then
                v_医嘱ids := Nvl(v_医嘱ids, '') || ',' || r_Bill.医嘱序号;
              End If;
              --记录病人医嘱附费对应的医嘱ID(不是主费用)
              If v_医嘱id Is Null Then
                v_医嘱id := r_Bill.医嘱序号;
              End If;
            End If;
          
          End If;
        
          If Nvl(n_划价, 0) = 0 Then
            --划价时,直接更改数量,所以不须查划冲销次数
            --插入退费记录
            Insert Into 住院费用记录
              (ID, NO, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 主页id, 病人id, 医嘱序号, 门诊标志, 多病人单, 婴儿费, 姓名, 性别, 年龄, 标识号, 床号, 费别, 病人病区id,
               病人科室id, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人,
               执行部门id, 划价人, 执行人, 执行状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 保险编码, 记帐单id, 摘要, 费用类型, 是否急诊,
               结论, 医疗小组id)
              Select 病人费用记录_Id.Nextval, NO, 记录性质, 2, 序号, 从属父号, 价格父号, 主页id, 病人id, 医嘱序号, 门诊标志, 多病人单, 婴儿费, 姓名, 性别, 年龄, 标识号,
                     床号, 费别, 病人病区id, 病人科室id, 收费类别, 收费细目id, 计算单位, Decode(Sign(v_准退数量 - Nvl(付数, 1) * 数次), 0, 付数, 1), 发药窗口,
                     Decode(Sign(v_准退数量 - Nvl(付数, 1) * 数次), 0, -1 * 数次, -1 * v_准退数量), 加班标志, Decode(记录性质,3,1,附加标志) 附加标志, 收入项目id, 收据费目, 记帐费用,
                     标准单价, -1 * v_应收金额, -1 * v_实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, -1 * v_退费次数, 执行时间, 操作员编号_In,
                     操作员姓名_In, 发生时间, v_Curdate, 保险项目否, 保险大类id, -1 * v_统筹金额, 保险编码, 记帐单id, 摘要, 费用类型, 是否急诊, 结论, 医疗小组id
              From 住院费用记录
              Where ID = r_Bill.Id;
          
            --记录病人医嘱附费对应的医嘱ID(不是主费用)
            If v_医嘱id Is Null And r_Bill.医嘱序号 Is Not Null Then
              v_医嘱id := r_Bill.医嘱序号;
            End If;
          
            Update 病人审批项目
            Set 已用数量 = Nvl(已用数量, 0) - v_准退数量
            Where 病人id = r_Bill.病人id And 主页id = r_Bill.主页id And 项目id = r_Bill.收费细目id And Nvl(使用限量, 0) <> 0;
          
            --病人余额
            Update 病人余额
            Set 费用余额 = Nvl(费用余额, 0) - v_实收金额
            Where 病人id = r_Bill.病人id And 类型 = 2 And 性质 = 1;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额
                (病人id, 类型, 性质, 费用余额, 预交余额)
              Values
                (r_Bill.病人id, 2, 1, -1 * v_实收金额, 0);
            End If;
          
            --病人未结费用
            Update 病人未结费用
            Set 金额 = Nvl(金额, 0) - v_实收金额
            Where 病人id = r_Bill.病人id And Nvl(主页id, 0) = Nvl(r_Bill.主页id, 0) And Nvl(病人病区id, 0) = Nvl(r_Bill.病人病区id, 0) And
                  Nvl(病人科室id, 0) = Nvl(r_Bill.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Bill.开单部门id, 0) And
                  Nvl(执行部门id, 0) = Nvl(r_Bill.执行部门id, 0) And 收入项目id + 0 = r_Bill.收入项目id And 来源途径 + 0 = 2;
            If Sql%RowCount = 0 Then
              Insert Into 病人未结费用
                (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
              Values
                (r_Bill.病人id, r_Bill.主页id, r_Bill.病人病区id, r_Bill.病人科室id, r_Bill.开单部门id, r_Bill.执行部门id, r_Bill.收入项目id, 2,
                 -1 * v_实收金额);
            End If;
          
            --标记原费用记录
            --执行状态:全部退完(准退数=剩余数)标记为0,否则保持原状态
            If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Then
              --一般情况非药品和卫材的项目,不存在部分销帐的情况,只有销帐申请和销帐审核时,才会出现部分销帐,所以
              --执行状态只有两种:0.未执行;1已执行;
              --由于在销帐审核过程中将已执行强制改为了2部分执行,因此需要在此处改为1已执行.未执行的不变.
              Update 住院费用记录
              Set 记录状态 = 3,附加标志=Decode(记录性质,3,1,附加标志), 执行状态 = Decode(Sign(v_准退数量 - v_剩余数量), 0, 0, Decode(执行状态, 2, 1, 执行状态))
              Where ID = r_Bill.Id;
            Else
              Select Nvl(Sum(Decode(审核人, Null, 1, 0) * Nvl(付数, 1) * 实际数量), 0),
                     Nvl(Sum(Decode(审核人, Null, 0, 1) * Nvl(付数, 1) * 实际数量), 0)
              Into n_未执行数量, n_已执行数量
              From 药品收发记录
              Where NO = No_In And 单据 In (9, 10, 25, 26) And 费用id = r_Bill.Id;
            
              Update 住院费用记录
              Set 记录状态 = 3,
                  执行状态 = Decode(Sign(v_准退数量 - v_剩余数量), 0, 0,
                                 Decode(Sign(n_未执行数量 - v_准退数量), 1, Decode(n_已执行数量, 0, 0, 2), 1))
              Where ID = r_Bill.Id;
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

  If Nvl(操作状态_In, 0) = 2 Then
    --转病区费用时:
    --1.药品及跟踪在用的卫材不会调用该过程
    --2.划价记账单也不会调用该过程
    --3.不需要更改医嘱信息
    For r_Bill In c_Bill Loop
      If Nvl(r_Bill.执行状态, 0) <> 1 Then
        b_Message.Zlhis_Charge_008(r_Bill.收费类别, r_Bill.Id);
      End If;
    End Loop;
    Return;
  End If;

  --不存在配药ID,检查该药品是否在输液配药中心
  If v_配药id Is Null And 输液配药检查_In = 1 Then
    For v_费用 In (Select ID
                 From 住院费用记录
                 Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 收费类别 In ('4', '5', '6', '7') And 门诊标志 = 2 And
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
  For v_费用 In (Select ID, 序号, 收费类别
               From 住院费用记录
               Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 收费类别 In ('4', '5', '6', '7') And 门诊标志 = 2 And
                     (Instr(',' || v_序号 || ',', ',' || 序号 || ',') > 0 Or v_序号 Is Null)
               Order By 收费细目id) Loop
    --根据费用ID来进行相关的处理
    v_准退数量 := 0;
    If Instr(序号_In, ':') > 0 Then
      v_Tmp := ',' || 序号_In;
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',' || v_费用.序号 || ':') + Length(',' || v_费用.序号 || ':'));
      v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp || ',', ',') - 1);
      If Instr(v_Tmp, ':') > 0 Then
        v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
      End If;
      v_准退数量 := v_Tmp;
    End If;
  
    Zl_药品收发记录_销售退费(v_费用.Id, v_准退数量, v_配药id, 1);
  End Loop;

  ---------------------------------------------------------------------------------
  --如果是划价,直接删除费用记录(药品处理后)
  n_Count := l_划价.Count;
  --删除划价记录
  Forall I In 1 .. l_划价.Count
    Delete From 住院费用记录 Where ID = l_划价(I);

  --删除之后再统一调整序号
  If n_Count > 0 Then
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.价格父号 Is Null Then
        v_父号 := n_Count;
      End If;
    
      Update 住院费用记录
      Set 序号 = n_Count, 价格父号 = Decode(价格父号, Null, Null, v_父号)
      Where NO = No_In And 记录性质 = 记录性质_In And 序号 = r_Serial.序号;
    
      Update 住院费用记录
      Set 从属父号 = n_Count
      Where NO = No_In And 记录性质 = 记录性质_In And 从属父号 = r_Serial.序号;
    
      n_Count := n_Count + 1;
    End Loop;
  
  End If;
  --整张单据全部冲完时，删除病人医嘱附费
  For c_医嘱 In (Select Distinct 医嘱序号
               From 住院费用记录
               Where NO = No_In And 记录性质 = 2 And 记录状态 = 3 And 医嘱序号 Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select 序号, Sum(数量) As 剩余数量
           From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                  From 住院费用记录
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
    Zl_医嘱发送_计费状态_Update(1, 2, 0, No_In, v_医嘱ids);
  Else
    Zl_医嘱发送_计费状态_Update(1, 2, 2, No_In);
  End If;
  For r_Bill In c_Bill Loop
    --卫材药品类别的消息放到Zl_药品收发记录_销售退费中发送
    If Nvl(r_Bill.执行状态, 0) <> 1 And Instr(',4,5,6,7,', ',' || r_Bill.收费类别 || ',') = 0 Then
      b_Message.Zlhis_Charge_008(r_Bill.收费类别, r_Bill.Id);
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_住院记帐记录_Delete;
/

--124924:冉俊明,2018-05-08,消费卡按每一张卡片设置限制类别
Create Or Replace Procedure Zl_消费卡类别目录_Update
(
  编号_In             消费卡类别目录.编号%Type,
  名称_In             消费卡类别目录.名称%Type,
  结算方式_In         消费卡类别目录.结算方式%Type,
  前缀文本_In         消费卡类别目录.前缀文本%Type,
  卡号长度_In         消费卡类别目录.卡号长度%Type,
  是否密文_In         消费卡类别目录.是否密文%Type,
  是否退现_In         消费卡类别目录.是否退现%Type,
  是否全退_In         消费卡类别目录.是否退现%Type,
  启用_In             消费卡类别目录.启用%Type,
  密码长度_In         消费卡类别目录.密码长度%Type,
  密码长度限制_In     消费卡类别目录.密码长度限制%Type,
  密码规则_In         消费卡类别目录.密码规则%Type,
  操作方式_In         Integer,
  读卡性质_In         消费卡类别目录.读卡性质%Type,
  键盘控制方式_In     消费卡类别目录.键盘控制方式%Type,
  限制类别_In         消费卡类别目录.限制类别%Type,
  是否严格控制_In     消费卡类别目录.是否严格控制%Type,
  是否特定病人_In     消费卡类别目录.是否特定病人%Type,
  是否允许换卡_In     消费卡类别目录.是否允许换卡%Type,
  是否允许补卡_In     消费卡类别目录.是否允许补卡%Type,
  是否允许余额退款_In 消费卡类别目录.是否允许余额退款%Type,
  应用场合_In         消费卡类别目录.应用场合%Type
) Is
  --操作方式_In 0-新增,else-修改 
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  v_卡名称 Varchar2(200);
Begin

  If 结算方式_In Is Null Then
    v_Err_Msg := '结算方式不能为空！';
    Raise Err_Item;
  End If;

  Begin
    Select 名称
    Into v_卡名称
    From (Select 名称
           From 医疗卡类别
           Where 结算方式 = 结算方式_In
           Union All
           Select 名称 From 消费卡类别目录 Where 编号 <> 编号_In And 结算方式 = 结算方式_In)
    Where Rownum < 2;
  Exception
    When Others Then
      v_卡名称 := Null;
  End;
  If v_卡名称 Is Not Null Then
    v_Err_Msg := '结算方式『' || 结算方式_In || '』已被' || v_卡名称 || '使用，重复使用会造成财务扎帐紊乱，请重新选定一种结算方式！';
    Raise Err_Item;
  End If;

  If 操作方式_In = 0 Then
    Insert Into 消费卡类别目录
      (编号, 名称, 结算方式, 启用, 自制卡, 前缀文本, 卡号长度, 是否密文, 是否退现, 是否全退, 密码长度, 密码长度限制, 密码规则, 读卡性质, 键盘控制方式, 限制类别, 是否严格控制, 是否特定病人,
       是否允许换卡, 是否允许补卡, 是否允许余额退款, 应用场合)
    Values
      (编号_In, 名称_In, 结算方式_In, 启用_In, 1, 前缀文本_In, 卡号长度_In, 是否密文_In, 是否退现_In, 是否全退_In, 密码长度_In, 密码长度限制_In, 密码规则_In,
       读卡性质_In, 键盘控制方式_In, 限制类别_In, 是否严格控制_In, 是否特定病人_In, 是否允许换卡_In, 是否允许补卡_In, 是否允许余额退款_In, 应用场合_In);
  Else
    Update 消费卡类别目录
    Set 名称 = 名称_In, 结算方式 = 结算方式_In, 启用 = 启用_In, 前缀文本 = 前缀文本_In, 卡号长度 = 卡号长度_In, 是否密文 = 是否密文_In, 是否退现 = 是否退现_In,
        是否全退 = 是否全退_In, 密码长度 = 密码长度_In, 密码长度限制 = 密码长度限制_In, 密码规则 = 密码规则_In, 读卡性质 = 读卡性质_In, 键盘控制方式 = 键盘控制方式_In,
        限制类别 = 限制类别_In, 是否严格控制 = 是否严格控制_In, 是否特定病人 = 是否特定病人_In, 是否允许换卡 = 是否允许换卡_In, 是否允许补卡 = 是否允许补卡_In,
        是否允许余额退款 = 是否允许余额退款_In, 应用场合 = 应用场合_In
    Where 编号 = 编号_In;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_消费卡类别目录_Update;
/

--124924:冉俊明,2018-05-08,消费卡按每一张卡片设置限制类别
Create Or Replace Procedure Zl_消费卡信息_Update
(
  Id_In         消费卡信息.Id%Type,
  卡类型_In     消费卡信息.卡类型%Type,
  可否充值_In   消费卡信息.可否充值%Type,
  有效期_In     消费卡信息.有效期%Type,
  发卡原因_In   消费卡信息.发卡原因%Type,
  领卡人_In     消费卡信息.领卡人%Type,
  病人id_In     消费卡信息.病人id%Type,
  领卡部门id_In 消费卡信息.领卡部门id%Type,
  备注_In       消费卡信息.备注%Type,
  限制类别_In   消费卡信息.限制类别%Type
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  v_卡号     消费卡信息.卡号%Type;
  v_卡名称   消费卡类别目录.名称%Type;
  n_可否充值 消费卡信息.可否充值%Type;
  d_回收时间 Date;
  d_停用时间 Date;
  n_序号     消费卡信息.序号%Type;
  n_最大序号 消费卡信息.序号%Type;
  n_Count    Number(2);
Begin
  Begin
    Select b.名称, a.卡号, a.可否充值, a.回收时间, a.停用日期, a.序号,
           (Select Max(序号) From 消费卡信息 B Where a.卡号 = b.卡号 And a.接口编号 = b.接口编号)
    Into v_卡名称, v_卡号, n_可否充值, d_回收时间, d_停用时间, n_序号, n_最大序号
    From 消费卡信息 A, 消费卡类别目录 B
    Where a.接口编号 = b.编号 And a.Id = Id_In;
  Exception
    When Others Then
      v_Err_Msg := '未找到卡信息，不能修改！';
      Raise Err_Item;
  End;

  If Nvl(n_序号, 0) < Nvl(n_最大序号, 0) Then
    v_Err_Msg := '不能修改历史发卡信息(卡号为“' || v_卡号 || '”)！';
    Raise Err_Item;
  End If;

  If Nvl(d_回收时间, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
    v_Err_Msg := '卡号为“' || v_卡号 || '”的' || v_卡名称 || '已经回收，不能修改！';
    Raise Err_Item;
  End If;
  If Nvl(d_停用时间, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
    v_Err_Msg := '卡号为“' || v_卡号 || '”的' || v_卡名称 || '已经停止使用，不能再修改！';
    Raise Err_Item;
  End If;

  If Nvl(可否充值_In, 0) = 0 And Nvl(n_可否充值, 0) = 1 Then
    --需要检查是否发生了充值记录 
    Select Count(1)
    Into n_Count
    From 病人卡结算记录　where 消费卡id = Id_In And 记录性质 = 2 And 记录状态 = 1 And Rownum < 2;
    If n_Count <> 0 Then
      v_Err_Msg := '卡号为“' || v_卡号 || '”的' || v_卡名称 || '原来是充值卡且发生了充值记录，不能更改为非充值卡！';
      Raise Err_Item;
    End If;
  End If;

  If Nvl(有效期_In, To_Date('3000-01-01', 'yyyy-mm-dd')) < Sysdate Then
    v_Err_Msg := '卡号为“' || v_卡号 || '”的' || v_卡名称 || '的效期不能小于当前系统时间！';
    Raise Err_Item;
  End If;

  Update 消费卡信息
  Set 卡类型 = 卡类型_In, 可否充值 = 可否充值_In, 有效期 = Decode(有效期_In, Null, To_Date('3000-01-01', 'yyyy-mm-dd'), 有效期_In),
      发卡原因 = 发卡原因_In, 领卡人 = 领卡人_In, 领卡部门id = 领卡部门id_In, 备注 = 备注_In, 病人id = 病人id_In, 限制类别 = 限制类别_In
  Where ID = Id_In;

  --调整卡面值有效期,退卡后取消退卡会有多条面值记录 
  --不调整升级以前的数据 And 交易序号 > 0 
  Update 帐户缴款余额
  Set 有效期 = Decode(有效期_In, Null, To_Date('3000-01-01', 'yyyy-mm-dd'), 有效期_In)
  Where 交易序号 In (Select 交易序号 From 病人卡结算记录 A Where a.消费卡id = Id_In And a.记录性质 = 1) And 消费卡id = Id_In And 交易序号 > 0;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_消费卡信息_Update;
/


--125261:胡俊勇,2018-05-08,转科医嘱校对发送处理自动停长嘱
CREATE OR REPLACE Procedure Zl_病人医嘱发送_Insert
(
  医嘱id_In     In 病人医嘱发送.医嘱id%Type,
  发送号_In     In 病人医嘱发送.发送号%Type,
  记录性质_In   In 病人医嘱发送.记录性质%Type,
  No_In         In 病人医嘱发送.No%Type,
  记录序号_In   In 病人医嘱发送.记录序号%Type,
  发送数次_In   In 病人医嘱发送.发送数次%Type,
  首次时间_In   In 病人医嘱发送.首次时间%Type,
  末次时间_In   In 病人医嘱发送.末次时间%Type,
  发送时间_In   In 病人医嘱发送.发送时间%Type,
  执行状态_In   In 病人医嘱发送.执行状态%Type,
  执行部门id_In In 病人医嘱发送.执行部门id%Type,
  计费状态_In   In 病人医嘱发送.计费状态%Type,
  First_In      In Number := 0,
  样本条码_In   In 病人医嘱发送.样本条码%Type := Null,
  操作员编号_In In 人员表.编号%Type := Null,
  操作员姓名_In In 人员表.姓名%Type := Null,
  领药号_In     In 未发药品记录.领药号%Type := Null,
  门诊记帐_In   In 病人医嘱发送.门诊记帐%Type := Null,
  分解时间_In   In Varchar2 := Null,
  原液皮试_In   In Varchar2 := Null
  --功能：填写病人医嘱发送记录
  --参数：
  --      医嘱id_In=要发送的每个医嘱ID
  --      First_IN=表示是否一组医嘱的第一医嘱行,以便处理医嘱相关内容(如成药,配方的第一行,因为给药途径,配方煎法,用法可能为叮嘱不发送)
  --      发送数次_IN,首次时间_IN,末次时间_IN:对"持续性"长嘱,不填写发送数次,可填写首末次时间(用于回退)。
  --      门诊记帐_In,住院临嘱发送到门诊记帐时才填写为1（因为记录性质是2，用于区分住院记帐），其余情况均填空。
  --      源液皮试_In 原液皮试医嘱ID，需求号7107/bug115972用于关联药品医嘱行和皮试医嘱行。关联字段为 病人医嘱发送.标本发送批号 存入药品行的医嘱ID值
  --      格式：1医嘱ID,2医嘱ID 前面一个为皮试医嘱的医嘱ID，第二个为药品行医嘱的医嘱ID
) Is
  --包含病人及医嘱(一组医嘱中第一行)相关信息的游标
  Cursor c_Advice Is
    Select Nvl(a.相关id, a.Id) As 组id, a.序号, a.病人id, a.主页id, a.婴儿, a.姓名, a.病人科室id, c.操作类型, a.诊疗类别, a.医嘱期效, a.医嘱状态, a.医嘱内容,
           a.开嘱医生, a.开嘱时间, a.开始执行时间, a.上次执行时间, a.执行终止时间, a.执行时间方案, a.频率次数, a.频率间隔, a.间隔单位, a.开嘱科室id, a.标本部位, a.执行科室id,
           a.相关id, a.诊疗项目id, a.挂号单
    From 病人医嘱记录 A, 诊疗项目目录 C
    Where a.诊疗项目id = c.Id And a.Id = 医嘱id_In;
  r_Advice c_Advice%RowType;

  --包含病人(婴儿)的所有未停长嘱(含配方长嘱),婴儿传入-1表示都处理
  Cursor c_Needstop
  (
    v_病人id   病人医嘱记录.病人id%Type,
    v_主页id   病人医嘱记录.主页id%Type,
    v_婴儿     病人医嘱记录.婴儿%Type,
    v_Stoptime Date
  ) Is
    Select a.Id, a.诊疗类别, b.操作类型, b.执行频率
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.诊疗项目id = b.Id(+) And a.病人id = v_病人id And a.主页id = v_主页id And (v_婴儿 = -1 Or Nvl(a.婴儿, 0) = Nvl(v_婴儿, 0)) And
          Nvl(a.医嘱期效, 0) = 0 And a.医嘱状态 Not In (1, 2, 4, 8, 9) And a.开始执行时间 <= v_Stoptime
    Order By a.序号;
  --包含病人(婴儿)的已停但未确认的长嘱,终止执行时间在指定时间之后,婴儿传入-1表示都处理
  Cursor c_Havestop
  (
    v_病人id   病人医嘱记录.病人id%Type,
    v_主页id   病人医嘱记录.主页id%Type,
    v_婴儿     病人医嘱记录.婴儿%Type,
    v_Stoptime Date
  ) Is
    Select ID
    From 病人医嘱记录
    Where 病人id = v_病人id And 主页id = v_主页id And (v_婴儿 = -1 Or Nvl(婴儿, 0) = Nvl(v_婴儿, 0)) And Nvl(医嘱期效, 0) = 0 And
          医嘱状态 = 8 And 执行终止时间 > v_Stoptime And 开始执行时间 < v_Stoptime
    Order By 序号;

  --其它临时变量
  v_婴儿       病人医嘱记录.婴儿%Type;
  v_持续性     Number(1); --是否持续性长嘱
  v_Autostop   Number(1);
  v_Date       Date;
  v_Temp       Varchar2(255);
  v_人员编号   人员表.编号%Type;
  v_人员姓名   人员表.姓名%Type;
  v_停止时间   病人医嘱记录.开嘱时间%Type;
  n_执行状态   病人医嘱发送.执行状态%Type;
  d_开始时间   病人医嘱记录.开始执行时间%Type;
  v_Count      Number;
  n_皮试标号   病人医嘱发送.医嘱id%Type;
  n_皮试医嘱id 病人医嘱发送.医嘱id%Type;

  v_Stopadviceids 病人医嘱记录.医嘱内容%Type;
  n_Adviceid      病人医嘱记录.病人id%Type;
  n_标记          Number(18);
  v_Error         Varchar2(255);
  Err_Custom Exception;
Begin
  --当前操作人员
  If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null Then
    v_人员编号 := 操作员编号_In;
    v_人员姓名 := 操作员姓名_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;
  --如果首次时间为空则填入开始执行时间
  If 首次时间_In Is Null Or 分解时间_In Is Null Or 末次时间_In Is Null Then
    Select 开始执行时间 Into d_开始时间 From 病人医嘱记录 Where ID = 医嘱id_In;
  End If;

  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;
  Close c_Advice;

  --是一组医嘱的第一行时处理医嘱内容
  If Nvl(First_In, 0) = 1 Then
    --并发操作检查
    ---------------------------------------------------------------------------------------
    If Nvl(r_Advice.医嘱状态, 0) = 4 Then
      --检查要发送的医嘱是否被作废
      v_Error := '"' || r_Advice.姓名 || '"的医嘱"' || r_Advice.医嘱内容 || '"已经被其他人作废。' || Chr(13) || Chr(10) ||
                 '该病人的医嘱发送失败。请重新读取发送清单再试。';
      Raise Err_Custom;
    End If;
  
    If Nvl(r_Advice.医嘱期效, 0) = 0 Then
      --长嘱：含成药长嘱,配方长嘱,非药"可选频率"长嘱,非药"持续性"长嘱
    
      --检查长嘱是否已被发送
      If r_Advice.上次执行时间 Is Not Null Then
        If r_Advice.上次执行时间 >= 首次时间_In Then
          v_Error := '"' || r_Advice.姓名 || '"的医嘱"' || r_Advice.医嘱内容 || '"已经被其他人发送。' || Chr(13) || Chr(10) ||
                     '该病人的医嘱发送失败。请重新读取发送清单再试。';
          Raise Err_Custom;
        End If;
      End If;
    
      --检查长嘱发送前是否已被自动停止(如术后)
      If r_Advice.执行终止时间 Is Not Null Then
        If 首次时间_In > r_Advice.执行终止时间 Then
          v_Error := '"' || r_Advice.姓名 || '"的医嘱"' || r_Advice.医嘱内容 || '"已经被停止。' || Chr(13) || Chr(10) ||
                     '该病人的医嘱发送失败。请重新读取发送清单再试。';
          Raise Err_Custom;
        End If;
      End If;
    Elsif Nvl(r_Advice.医嘱状态, 0) In (8, 9) Then
      --临嘱：含配方临嘱
    
      --检查是否已被发送(或因其它原因自动停止)
      v_Error := '"' || r_Advice.姓名 || '"的医嘱"' || r_Advice.医嘱内容 || '"已经被其他人发送。' || Chr(13) || Chr(10) ||
                 '该病人的医嘱发送失败。请重新读取发送清单再试。';
      Raise Err_Custom;
    End If;
  
    --发送后的医嘱处理
    ---------------------------------------------------------------------------------------
    If Nvl(r_Advice.医嘱期效, 0) = 0 Then
      --长期医嘱:更新上次执行时间
      Update 病人医嘱记录 Set 上次执行时间 = 末次时间_In Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
    
      --判断是否持续性长嘱
      v_持续性 := 0;
      If r_Advice.执行时间方案 Is Null And (Nvl(r_Advice.频率次数, 0) = 0 Or Nvl(r_Advice.频率间隔, 0) = 0 Or r_Advice.间隔单位 Is Null) Then
        v_持续性 := 1;
      End If;
    
      --预定了终止时间且未停止的自动停止
      If r_Advice.执行终止时间 Is Not Null And Nvl(r_Advice.医嘱状态, 0) Not In (8, 9) Then
        v_Autostop := 0;
        If v_持续性 = 1 Then
          --非药"持续性"长嘱
          If Trunc(末次时间_In) = Trunc(r_Advice.执行终止时间 - 1) Then
            v_Autostop := 1; --终止这天不执行
          End If;
        Elsif Zl_Advicenexttime(医嘱id_In) > r_Advice.执行终止时间 Then
          --成药长嘱或非药"可选频率"长嘱
          v_Autostop := 1; --如果是等于,还可以执行一次
        End If;
      
        If v_Autostop = 1 Then
          Update 病人医嘱记录
          Set 医嘱状态 = 8, 停嘱时间 = 末次时间_In, 停嘱医生 = r_Advice.开嘱医生
          Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
          v_Temp := zl_GetSysParameter(271);
          If v_Temp = '1' Then
            v_Temp := '自动停止：预定停止时间。';
          Else
            v_Temp := Null;
          End If;
          Insert Into 病人医嘱状态
            (医嘱id, 操作类型, 操作人员, 操作时间, 操作说明)
            Select ID, 8, r_Advice.开嘱医生, 发送时间_In, v_Temp
            From 病人医嘱记录
            Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Advice.组id;
        End If;
      End If;
    Else
      --临嘱停止。
      --住院医生发送时自动校对、停止：校对是以Sysdate取的,为避免重复,停止时间也取Sysdate
      Select Sysdate Into v_Date From Dual;
      Update 病人医嘱记录
      Set 医嘱状态 = 8, 执行终止时间 = 末次时间_In,
          --为一次性临嘱时没有
          上次执行时间 = 末次时间_In,
          --为一次性临嘱时没有
          停嘱时间 = v_Date,
          --发送时间_IN,
          停嘱医生 = r_Advice.开嘱医生
      Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
    
      Insert Into 病人医嘱状态
        (医嘱id, 操作类型, 操作人员, 操作时间)
        Select ID, 8, v_人员姓名, v_Date --发送时间_IN
        From 病人医嘱记录
        Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
    End If;
  
    --特殊医嘱的处理
    ---------------------------------------------------------------------------------------
    If r_Advice.诊疗类别 = 'Z' And Nvl(r_Advice.操作类型, '0') <> '0' Then
      --(1-留观;2-住院;)3-转科;4-术后(不发送);5-出院;6-转院,7-会诊,11-死亡
    
      --几种特殊医嘱要自动停止病人该医嘱之前(按时间算)所有未停的长嘱
      If r_Advice.操作类型 In ('3', '5', '6', '11') Then
        If Nvl(r_Advice.婴儿, 0) = 0 Then
          v_婴儿 := -1;
        Else
          v_婴儿 := Nvl(r_Advice.婴儿, 0);
        End If;
        For r_Needstop In c_Needstop(r_Advice.病人id, r_Advice.主页id, v_婴儿, r_Advice.开始执行时间) Loop
          Select Decode(Sign(开始执行时间 - r_Advice.开始执行时间), 1, 开始执行时间, r_Advice.开始执行时间)
          Into v_停止时间
          From 病人医嘱记录
          Where ID = r_Needstop.Id;
          Update 病人医嘱记录
          Set 医嘱状态 = 8, 执行终止时间 = v_停止时间, 停嘱时间 = 发送时间_In, 停嘱医生 = r_Advice.开嘱医生
          Where ID = r_Needstop.Id;
        
          Insert Into 病人医嘱状态
            (医嘱id, 操作类型, 操作人员, 操作时间)
            Select ID, 8, v_人员姓名, 发送时间_In From 病人医嘱记录 Where ID = r_Needstop.Id;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Needstop.Id;
        End Loop;
        --已停止未确认的长嘱,终止时间在医嘱开始后的,调前其终止时间(同时多个特殊医嘱的情况)
        For r_Havestop In c_Havestop(r_Advice.病人id, r_Advice.主页id, v_婴儿, r_Advice.开始执行时间) Loop
          Update 病人医嘱记录
          Set 执行终止时间 = Decode(Sign(开始执行时间 - r_Advice.开始执行时间), 1, 开始执行时间, r_Advice.开始执行时间), 停嘱时间 = 发送时间_In,
              停嘱医生 = r_Advice.开嘱医生
          Where ID = r_Havestop.Id;
        
          --不修改停止医嘱的操作人员，因为停止时，医生可能已进行电子签名
          Update 病人医嘱状态 Set 操作时间 = 发送时间_In Where 医嘱id = r_Havestop.Id And 操作类型 = 8;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Havestop.Id;
        End Loop;
        --处理长期备用医嘱(没有执行（发送）过的标记未用）,同时处理临嘱
        Update 病人医嘱记录
        Set 执行标记 = -1
        Where 病人id = r_Advice.病人id And 主页id = r_Advice.主页id And
              (医嘱期效 = 0 And 执行频次 = '必要时' And 上次执行时间 Is Null And 医嘱状态 In (3, 5, 6, 7) Or
              医嘱期效 = 1 And 执行频次 = '需要时' And 医嘱状态 = 3) And 执行标记 <> -1;
      End If;
    
      --具体的特殊处理
      If Nvl(r_Advice.婴儿, 0) = 0 Then
        If r_Advice.操作类型 = '3' And 执行部门id_In Is Not Null And r_Advice.病人科室id Is Not Null And
           Nvl(r_Advice.病人科室id, 0) <> Nvl(执行部门id_In, 0) Then
          --转科医嘱,将病人登记转科到"执行科室ID"(在院病人且当前科室与转入科室不同才处理)
          Select Count(1)
          Into v_Temp
          From 病案主页
          Where 病人id = r_Advice.病人id And 主页id = r_Advice.主页id And 出院科室id <> 执行部门id_In;
          If v_Temp = '1' Then
            Zl_病人变动记录_Change(r_Advice.病人id, r_Advice.主页id, 执行部门id_In, v_人员编号, v_人员姓名);
          End If;
        Elsif r_Advice.操作类型 In ('5', '6', '11') Then
          --出院、转院、死亡医嘱,将病人标记为预出院
          Begin
            Select 开始时间
            Into v_Date
            From 病人变动记录
            Where 开始时间 Is Not Null And 终止时间 Is Null And 病人id = r_Advice.病人id And 主页id = r_Advice.主页id;
          Exception
            When Others Then
              v_Date := To_Date('1900-01-01', 'YYYY-MM-DD');
          End;
          If r_Advice.开始执行时间 <= v_Date Then
            v_Error := '医嘱"' || r_Advice.医嘱内容 || '"的开始时间应大于该病人上次变动时间 ' || To_Char(v_Date, 'YYYY-MM-DD HH24:Mi') || ' 。';
            Raise Err_Custom;
          End If;
          Zl_病人变动记录_Preout(r_Advice.病人id, r_Advice.主页id, r_Advice.开始执行时间);
        End If;
      Else
        If r_Advice.操作类型 = '11' Then
          Update 病人新生儿记录
          Set 死亡时间 = r_Advice.开始执行时间
          Where 病人id = r_Advice.病人id And Nvl(主页id, 0) = Nvl(r_Advice.主页id, 0) And Nvl(序号, 0) = Nvl(r_Advice.婴儿, 0);
        End If;
      End If;
    End If;
    --12小时未执行的备用临嘱处理为标记未用
    If r_Advice.医嘱期效 = 1 Then
      Update 病人医嘱记录
      Set 执行标记 = -1
      Where 病人id = r_Advice.病人id And 主页id = r_Advice.主页id And 执行标记 <> -1 And 医嘱期效 = 1 And 执行频次 = '需要时' And
            Sysdate - 开始执行时间 > 0.5 And 医嘱状态 = 3;
    End If;
  End If;

  --填写发送记录
  ---------------------------------------------------------------------------------------
  n_执行状态 := 执行状态_In;
  If 执行状态_In = 1 Then
    v_Temp := zl_GetSysParameter(186);
    If v_Temp = '11' Then
      If r_Advice.诊疗类别 = 'E' And r_Advice.操作类型 In ('1', '8') Or r_Advice.诊疗类别 = 'K' Then
        n_执行状态 := 0;
      End If;
    Elsif v_Temp = '01' Then
      If r_Advice.诊疗类别 = 'E' And r_Advice.操作类型 = '1' Then
        n_执行状态 := 0;
      End If;
    Elsif v_Temp = '10' Then
      If r_Advice.诊疗类别 = 'E' And r_Advice.操作类型 = '8' Or r_Advice.诊疗类别 = 'K' Then
        n_执行状态 := 0;
      End If;
    End If;
  End If;

  If 原液皮试_In Is Not Null Then
    v_Count      := Instr(原液皮试_In, ',');
    n_皮试医嘱id := Substr(原液皮试_In, 1, v_Count - 1);
    n_皮试标号   := Substr(原液皮试_In, v_Count + 1);
    Update 病人医嘱发送 Set 标本发送批号 = n_皮试标号 Where 医嘱id = n_皮试医嘱id;
  End If;

  Insert Into 病人医嘱发送
    (医嘱id, 发送号, 记录性质, NO, 记录序号, 发送数次, 发送人, 发送时间, 执行状态, 执行部门id, 计费状态, 首次时间, 末次时间, 样本条码, 门诊记帐, 标本发送批号)
  Values
    (医嘱id_In, 发送号_In, 记录性质_In, No_In, 记录序号_In, 发送数次_In, v_人员姓名, 发送时间_In, n_执行状态, 执行部门id_In, 计费状态_In,
     Nvl(首次时间_In, d_开始时间), Nvl(末次时间_In, d_开始时间), 样本条码_In, 门诊记帐_In, n_皮试标号);

  --手术和检查医嘱同步更新主医嘱的计费状态
  If 计费状态_In = 1 And r_Advice.组id <> 医嘱id_In And (r_Advice.诊疗类别 = 'D' Or r_Advice.诊疗类别 = 'F') Then
    Update 病人医嘱发送 Set 计费状态 = 1 Where 医嘱id = r_Advice.组id And 发送号 = 发送号_In;
  End If;

  --领药号的填写
  If 领药号_In Is Not Null Then
    Update 未发药品记录 Set 领药号 = 领药号_In Where NO = No_In And 单据 = 9 And 领药号 Is Null;
    Update 药品收发记录 Set 产品合格证 = 领药号_In Where NO = No_In And 单据 = 9 And 产品合格证 Is Null;
  End If;

  --自动填为已执行时，需要同步处理费用执行状态及审核划价状态
  If 执行状态_In = 1 Then
    Zl_病人医嘱执行_Finish(医嘱id_In, 发送号_In, Null, Null, v_人员编号, v_人员姓名, 执行部门id_In);
  End If;

  --产生医嘱执行时间记录(只产生主记录的)
  If Nvl(分解时间_In, To_Char(d_开始时间, 'yyyy-mm-dd hh24:mi:ss')) Is Not Null Then
    If r_Advice.相关id Is Null Then
      Insert Into 医嘱执行时间
        (要求时间, 医嘱id, 发送号)
        Select To_Date(Column_Value, 'yyyy-mm-dd hh24:mi:ss'), 医嘱id_In, 发送号_In
        From Table(f_Str2list(Nvl(分解时间_In, To_Char(d_开始时间, 'yyyy-mm-dd hh24:mi:ss'))));
    End If;
  End If;

  --病历书写时机的填写
  If r_Advice.诊疗类别 = 'F' Then
    --一组手术只调一次
    If r_Advice.相关id Is Null Then
      If Not r_Advice.标本部位 Is Null Then
        v_Date := To_Date(r_Advice.标本部位, 'yyyy-mm-dd hh24:mi:ss');
      Else
        v_Date := r_Advice.开始执行时间;
      End If;
      Zl_电子病历时机_Insert(r_Advice.病人id, r_Advice.主页id, 2, '手术', r_Advice.开嘱科室id, r_Advice.开嘱医生, v_Date, v_Date,
                       r_Advice.执行科室id);
    End If;
  Elsif r_Advice.诊疗类别 = 'Z' And r_Advice.操作类型 = '7' Then
    Zl_电子病历时机_Insert(r_Advice.病人id, r_Advice.主页id, 2, '会诊', r_Advice.开嘱科室id, r_Advice.开嘱医生, r_Advice.开始执行时间,
                     r_Advice.开始执行时间, r_Advice.执行科室id);
  Elsif r_Advice.诊疗类别 = 'Z' And r_Advice.操作类型 = '8' Then
    Zl_电子病历时机_Insert(r_Advice.病人id, r_Advice.主页id, 2, '抢救', r_Advice.开嘱科室id, r_Advice.开嘱医生, r_Advice.开始执行时间,
                     r_Advice.开始执行时间, r_Advice.执行科室id);
  Elsif r_Advice.诊疗类别 = 'Z' And r_Advice.操作类型 = '11' Then
    Zl_电子病历时机_Insert(r_Advice.病人id, r_Advice.主页id, 2, '死亡', r_Advice.开嘱科室id, r_Advice.开嘱医生, r_Advice.开始执行时间,
                     r_Advice.开始执行时间, r_Advice.执行科室id);
  End If;
  --额外调用(知情文件允许的诊疗类别才调用)
  If Instr('C,D,E,F,G,K,L', r_Advice.诊疗类别) > 0 Then
    Zl_电子病历时机_Insert(r_Advice.病人id, r_Advice.主页id, 2, '知情文书', r_Advice.开嘱科室id, r_Advice.开嘱医生, r_Advice.开始执行时间,
                     r_Advice.开始执行时间, r_Advice.执行科室id, r_Advice.诊疗项目id, r_Advice.医嘱内容);
  End If;
  --医嘱停止消息的处理
  If v_Stopadviceids Is Not Null Then
    v_Stopadviceids := Substr(v_Stopadviceids, 2);
    b_Message.Zlhis_Cis_002(r_Advice.病人id, r_Advice.主页id, Null, v_Stopadviceids);
    Select Max(a.Id)
    Into n_标记
    From 病人医嘱记录 A
    Where a.Id In (Select Column_Value From Table(f_Num2list(v_Stopadviceids))) And a.医嘱期效 = 0 And a.医嘱状态 = 8 And
          Nvl(a.执行标记, 0) <> -1 And a.病人来源 <> 3;
    If n_标记 Is Not Null Then
      Select Max(a.Id)
      Into n_Adviceid
      From 病人医嘱记录 A
      Where a.Id In (Select Column_Value From Table(f_Num2list(v_Stopadviceids))) And a.紧急标志 = 1 And a.医嘱期效 = 0 And
            a.医嘱状态 = 8 And Nvl(a.执行标记, 0) <> -1 And a.病人来源 <> 3;
      If n_Adviceid Is Not Null Then
        Select Nvl(Max(0), 2)
        Into n_标记
        From 业务消息清单 A
        Where a.病人id = r_Advice.病人id And a.就诊id = r_Advice.主页id And a.类型编码 = 'ZLHIS_CIS_002' And a.优先程度 = 2 And
              a.是否已阅 = 0;
      Else
        n_Adviceid := n_标记;
        Select Nvl(Max(0), 1)
        Into n_标记
        From 业务消息清单 A
        Where a.病人id = r_Advice.病人id And a.就诊id = r_Advice.主页id And a.类型编码 = 'ZLHIS_CIS_002' And a.是否已阅 = 0;
      End If;
      If n_标记 > 0 Then
        For R In (Select a.病人性质 As 性质, a.出院科室id As 科室id, a.当前病区id As 病区id
                  From 病案主页 A
                  Where a.病人id = r_Advice.病人id And a.主页id = r_Advice.主页id) Loop
          Zl_业务消息清单_Insert(r_Advice.病人id, r_Advice.主页id, r.科室id, r.病区id, r.性质, '有新停止医嘱。', '0010', 'ZLHIS_CIS_002',
                           n_Adviceid, n_标记, 0, Null, r.病区id);
        End Loop;
      End If;
    End If;
  End If;

  If r_Advice.诊疗类别 = 'E' And r_Advice.操作类型 = '6' Then
    --检验项目
    b_Message.Zlhis_Cis_016(r_Advice.病人id, r_Advice.主页id, r_Advice.挂号单, 发送号_In, r_Advice.组id, 2);
  Elsif r_Advice.诊疗类别 = 'D' And r_Advice.相关id Is Null Then
    b_Message.Zlhis_Cis_017(r_Advice.病人id, r_Advice.主页id, r_Advice.挂号单, 发送号_In, r_Advice.组id, 2);
  Elsif r_Advice.诊疗类别 = 'F' And r_Advice.相关id Is Null Then
    b_Message.Zlhis_Cis_018(r_Advice.病人id, r_Advice.主页id, r_Advice.挂号单, 发送号_In, r_Advice.组id);
  Elsif r_Advice.诊疗类别 = 'K' Then
    b_Message.Zlhis_Cis_019(r_Advice.病人id, r_Advice.主页id, r_Advice.挂号单, 发送号_In, r_Advice.组id);
  Elsif r_Advice.诊疗类别 = 'Z' Then
    If r_Advice.操作类型 = '7' Then
      b_Message.Zlhis_Cis_020(r_Advice.病人id, r_Advice.主页id, 发送号_In, r_Advice.组id);
    Elsif r_Advice.操作类型 = '8' Then
      b_Message.Zlhis_Cis_021(r_Advice.病人id, r_Advice.主页id, 发送号_In, r_Advice.组id);
    Elsif r_Advice.操作类型 = '11' Then
      b_Message.Zlhis_Cis_022(r_Advice.病人id, r_Advice.主页id, 发送号_In, r_Advice.组id);
    End If;
  Elsif r_Advice.诊疗类别 = 'E' And r_Advice.操作类型 = '5' Then
    b_Message.Zlhis_Cis_023(r_Advice.病人id, r_Advice.主页id, 发送号_In, r_Advice.组id);
  End If;
  
  --病区执行医嘱发送
  Select Count(1) Into n_标记 From 部门性质说明 B Where b.部门id = r_Advice.执行科室id And b.工作性质 = '护理';
  If n_标记 > 0 Then
    b_Message.Zlhis_Cis_026(r_Advice.病人id, r_Advice.主页id, 发送号_In, 医嘱id_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱发送_Insert;
/

--125261:胡俊勇,2018-05-08,转科医嘱校对发送处理自动停长嘱
CREATE OR REPLACE Procedure Zl_病人医嘱记录_校对
(
  --功能：校对指定的医嘱 
  --参数：医嘱ID_IN=Nvl(相关ID,ID) 
  --      状态_IN=校对通过3或校对疑问2 
  --      自动校对_IN=保存之后调用自动校对,自动填写计价内容 
  --说明：一组医嘱只能调用一次,过程同时完成处理一组医嘱的校对 
  医嘱id_In     In 病人医嘱记录.Id%Type,
  状态_In       In 病人医嘱记录.医嘱状态%Type,
  校对时间_In   In 病人医嘱状态.操作时间%Type,
  校对说明_In   In 病人医嘱状态.操作说明%Type := Null,
  自动校对_In   In Number := Null,
  操作员编号_In In 人员表.编号%Type := Null,
  操作员姓名_In In 人员表.姓名%Type := Null
) Is
  --用于医嘱检查 
  v_状态       病人医嘱记录.医嘱状态%Type;
  v_期效       病人医嘱记录.医嘱期效%Type;
  v_病人id     病人医嘱记录.病人id%Type;
  v_主页id     病人医嘱记录.主页id%Type;
  v_婴儿       病人医嘱记录.婴儿%Type;
  v_医嘱内容   病人医嘱记录.医嘱内容%Type;
  v_开嘱时间   病人医嘱记录.开嘱时间%Type;
  v_开始时间   病人医嘱记录.开始执行时间%Type;
  v_开嘱医生   病人医嘱记录.开嘱医生%Type;
  v_前提id     病人医嘱记录.前提id%Type;
  v_执行标记   病人医嘱记录.执行标记%Type;
  v_执行科室id 病人医嘱记录.执行科室id%Type;
  v_标本部位   病人医嘱记录.标本部位%Type;
  v_停止时间   病人医嘱记录.开嘱时间%Type;
  v_开嘱科室id 病人医嘱记录.开嘱科室id%Type;
  n_病人科室id 病人医嘱记录.病人科室id%Type;

  --用于变更护理等级 
  v_诊疗类别   病人医嘱记录.诊疗类别%Type;
  v_诊疗项目id 病人医嘱记录.诊疗项目id%Type;
  v_操作类型   诊疗项目目录.操作类型%Type;
  v_护理等级id 病案主页.护理等级id%Type;
  v_紧急标志   病人医嘱记录.紧急标志%Type;
  v_入院方式   入院方式.名称%Type;

  v_药品等级   收费价格等级.名称%Type;
  v_卫材等级   收费价格等级.名称%Type;
  v_普通等级   收费价格等级.名称%Type;
  v_Pricegrade Varchar2(1000);
  v_站点       部门表.站点%Type;

  v_Stopadviceids 病人医嘱记录.医嘱内容%Type;
  n_Adviceid      病人医嘱记录.病人id%Type;
  n_标记          Number(18);
  --与该项目同一自动停止互斥组的项目:组中应该都是长嘱(包括当前医嘱),程序应已检查。 
  --注意应加婴儿条件,同时也应停止除当前医嘱外的其它相同诊疗项目的医嘱。 
  Cursor c_Exclude Is
    Select Distinct b.Id As 医嘱id, b.开始执行时间, b.执行终止时间, b.上次执行时间, b.开嘱医生, b.执行时间方案, b.频率间隔, b.频率次数, b.间隔单位
    From 诊疗互斥项目 A, 病人医嘱记录 B
    Where a.类型 = 3 And a.项目id = b.诊疗项目id And b.Id <> 医嘱id_In And Nvl(b.医嘱期效, 0) = 0 And b.医嘱状态 In (3, 5, 6, 7) And
          b.病人id = v_病人id And Nvl(b.主页id, 0) = Nvl(v_主页id, 0) And Nvl(b.婴儿, 0) = Nvl(v_婴儿, 0) And
          a.组编号 In (Select Distinct 组编号 From 诊疗互斥项目 Where 类型 = 3 And 项目id = v_诊疗项目id)
    Order By b.Id;
  v_终止时间 病人医嘱记录.执行终止时间%Type;

  --护理等级互斥 
  Cursor c_Nurse Is
    Select a.Id As 医嘱id, a.开始执行时间, a.执行终止时间, a.上次执行时间, a.开嘱医生
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.诊疗项目id = b.Id And a.诊疗类别 = 'H' And b.操作类型 = '1' And a.病人id = v_病人id And Nvl(a.主页id, 0) = Nvl(v_主页id, 0) And
          Nvl(a.婴儿, 0) = Nvl(v_婴儿, 0) And Nvl(a.医嘱期效, 0) = 0 And a.医嘱状态 In (3, 5, 6, 7) And a.Id <> 医嘱id_In;

  --记录入出量互斥 
  Cursor c_Patiio Is
    Select a.Id As 医嘱id, a.开始执行时间, a.执行终止时间, a.上次执行时间, a.开嘱医生
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.诊疗项目id = b.Id And a.诊疗类别 = 'Z' And b.操作类型 = '12' And a.病人id = v_病人id And Nvl(a.主页id, 0) = Nvl(v_主页id, 0) And
          Nvl(a.婴儿, 0) = Nvl(v_婴儿, 0) And Nvl(a.医嘱期效, 0) = 0 And a.医嘱状态 In (3, 5, 6, 7) And a.Id <> 医嘱id_In;

  --记录病情互斥 
  Cursor c_Patistate Is
    Select a.Id As 医嘱id, a.开始执行时间, a.执行终止时间, a.上次执行时间, a.开嘱医生
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.诊疗项目id = b.Id And a.诊疗类别 = 'Z' And b.操作类型 In ('9', '10') And a.病人id = v_病人id And
          Nvl(a.主页id, 0) = Nvl(v_主页id, 0) And Nvl(a.婴儿, 0) = Nvl(v_婴儿, 0) And Nvl(a.医嘱期效, 0) = 0 And
          a.医嘱状态 In (3, 5, 6, 7) And a.Id <> 医嘱id_In;
  --变动有效记录 
  Cursor c_Oldinfo Is
    Select b.*
    From (Select c.*
           From 病人变动记录 C
           Where c.病人id = v_病人id And c.主页id = v_主页id And
                 c.开始时间 = (Select Min(y.开始时间)
                           From 病人变动记录 Y
                           Where y.病人id = v_病人id And y.主页id = v_主页id And y.开始时间 > v_开始时间) And
                 Nvl(c.终止时间 || '', '空') =
                 (Select Nvl(Min(x.终止时间) || '', '空')
                  From 病人变动记录 X
                  Where x.病人id = v_病人id And x.主页id = v_主页id And x.开始时间 > v_开始时间)) A, 病人变动记录 B
    Where b.病人id = v_病人id And b.主页id = v_主页id And a.开始时间 = b.终止时间 And a.开始原因 = b.终止原因 And a.附加床位 = b.附加床位
    Union
    Select a.*
    From 病人变动记录 A
    Where a.病人id = v_病人id And a.主页id = v_主页id And a.终止时间 Is Null And a.开始时间 <= v_开始时间;

  Cursor c_Endinfo Is
    Select * From 病人变动记录 Where 病人id = v_病人id And 主页id = v_主页id And 终止时间 Is Null;
  r_Oldinfo      c_Oldinfo%RowType;
  r_Endinfo      c_Endinfo%RowType;
  v_变动终止原因 病人变动记录.终止原因%Type;
  v_变动终止时间 病人变动记录.终止时间%Type;
  v_变动终止人员 病人变动记录.终止人员%Type;

  --包含病人(婴儿)的所有未停长嘱(含配方长嘱) 
  Cursor c_Needstop
  (
    v_病人id   病人医嘱记录.病人id%Type,
    v_主页id   病人医嘱记录.主页id%Type,
    v_婴儿     病人医嘱记录.婴儿%Type,
    v_Stoptime Date
  ) Is
    Select a.Id, a.诊疗类别, b.操作类型, b.执行频率
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.诊疗项目id = b.Id(+) And a.病人id = v_病人id And a.主页id = v_主页id And (v_婴儿 = -1 Or Nvl(a.婴儿, 0) = Nvl(v_婴儿, 0)) And
          Nvl(a.医嘱期效, 0) = 0 And a.医嘱状态 Not In (1, 2, 4, 8, 9) And a.开始执行时间 <= v_Stoptime
    Order By a.序号;
  --包含病人(婴儿)的已停但未确认的长嘱,终止执行时间在指定时间之后 
  Cursor c_Havestop
  (
    v_病人id   病人医嘱记录.病人id%Type,
    v_主页id   病人医嘱记录.主页id%Type,
    v_婴儿     病人医嘱记录.婴儿%Type,
    v_Stoptime Date
  ) Is
    Select ID
    From 病人医嘱记录
    Where 病人id = v_病人id And 主页id = v_主页id And (v_婴儿 = -1 Or Nvl(婴儿, 0) = Nvl(v_婴儿, 0)) And Nvl(医嘱期效, 0) = 0 And
          医嘱状态 = 8 And 执行终止时间 > v_Stoptime And 开始执行时间 < v_Stoptime
    Order By 序号;

  --取一组医嘱的计价内容 
  Cursor c_Price Is
    Select a.Id, b.收费项目id, b.收费数量, b.从属项目, b.费用性质, b.收费方式, c.类别 As 收费类别, a.诊疗类别, e.操作类型, e.试管编码,
           Sum(Decode(Nvl(c.是否变价, 0), 1, Nvl(d.缺省价格, d.原价), Null)) As 单价
    From 病人医嘱记录 A, 诊疗收费关系 B, 收费项目目录 C, 收费价目 D, 诊疗项目目录 E
    Where a.诊疗项目id = b.诊疗项目id And b.收费项目id = c.Id And c.Id = d.收费细目id And
          ((Instr(';5;6;7;', ';' || c.类别 || ';') > 0 And d.价格等级 = v_药品等级) Or
          (Instr(';4;', ';' || c.类别 || ';') > 0 And d.价格等级 = v_卫材等级) Or
          (Instr(';4;5;6;7;', ';' || c.类别 || ';') = 0 And d.价格等级 = v_普通等级) Or
          (d.价格等级 Is Null And Not Exists
           (Select 1
             From 收费价目
             Where c.Id = 收费细目id And ((Instr(';5;6;7;', ';' || c.类别 || ';') > 0 And 价格等级 = v_药品等级) Or
                   (Instr(';4;', ';' || c.类别 || ';') > 0 And 价格等级 = v_卫材等级) Or
                   (Instr(';4;5;6;7;', ';' || c.类别 || ';') = 0 And 价格等级 = v_普通等级))))) And
          (a.相关id Is Null And a.执行标记 In (1, 2) And b.费用性质 = 1 Or
          a.标本部位 = b.检查部位 And a.检查方法 = b.检查方法 And Nvl(b.费用性质, 0) = 0 Or
          a.检查方法 Is Null And Nvl(b.费用性质, 0) = 0 And b.检查部位 Is Null And b.检查方法 Is Null) And
          a.诊疗类别 Not In ('5', '6', '7') And Nvl(a.计价特性, 0) = 0 And Nvl(a.执行性质, 0) Not In (0, 5) And c.服务对象 In (2, 3) And
          (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) And Sysdate Between d.执行日期 And
          Nvl(d.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(b.收费数量, 0) <> 0 And
          Not (Nvl(c.是否变价, 0) = 1 And Nvl(Nvl(d.缺省价格, d.原价), 0) = 0) And a.诊疗项目id = e.Id And
          (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In)
    Group By a.Id, b.收费项目id, b.收费数量, b.从属项目, b.费用性质, b.收费方式, c.类别, a.诊疗类别, e.操作类型, e.试管编码;

  Cursor c_Pati(v_病人id 病人信息.病人id%Type) Is
    Select * From 病人信息 Where 病人id = v_病人id;
  r_Pati c_Pati%RowType;

  v_材料id 采血管类型.材料id%Type;

  --其它临时变量 
  v_Count    Number;
  v_Date     Date;
  v_Temp     Varchar2(255);
  v_Par停嘱  Varchar2(255);
  v_人员编号 人员表.编号%Type;
  v_人员姓名 人员表.姓名%Type;
  v_叮嘱执行 Varchar2(5);

  v_Error Varchar2(255);
  Err_Custom Exception;

  Function Getadvicetext(v_医嘱id 病人医嘱记录.Id%Type) Return Varchar2 Is
    v_Text 病人医嘱记录.医嘱内容%Type;
    v_类别 病人医嘱记录.诊疗类别%Type;
    v_配方 Number;
  Begin
    Select 诊疗类别, 医嘱内容 Into v_类别, v_Text From 病人医嘱记录 Where ID = v_医嘱id;
    If v_类别 = 'E' Then
      --西药，中成药的医嘱内容 
      Begin
        Select 诊疗类别, Decode(诊疗类别, '7', v_Text, 医嘱内容)
        Into v_类别, v_Text
        From 病人医嘱记录
        Where 相关id = v_医嘱id And 诊疗类别 In ('5', '6', '7') And Rownum = 1;
      Exception
        When Others Then
          Null;
      End;
      If v_类别 = '7' Then
        v_配方 := 1;
      End If;
    End If;
    If Length(v_Text) > 30 Then
      v_Text := Substr(v_Text, 1, 30) || '...';
    End If;
    If Length(v_Text) > 20 Then
      v_Text := '"' || v_Text || '"' || Chr(13) || Chr(10);
    Else
      v_Text := '"' || v_Text || '"';
    End If;
    If v_配方 = 1 Then
      v_Text := '中药配方' || v_Text;
    End If;
    Return(v_Text);
  End;
Begin
  --检查医嘱状态是否正确:并发操作 
  Begin
    Select a.医嘱期效, a.医嘱状态, a.开嘱时间, a.开嘱医生, a.开始执行时间, a.病人id, a.主页id, a.婴儿, a.医嘱内容, a.诊疗类别, a.诊疗项目id, a.前提id,
           Nvl(b.操作类型, '0'), Nvl(a.执行标记, 0), a.执行科室id, a.标本部位, a.开嘱科室id, Nvl(a.紧急标志, 0) As 紧急标志, a.病人科室id
    Into v_期效, v_状态, v_开嘱时间, v_开嘱医生, v_开始时间, v_病人id, v_主页id, v_婴儿, v_医嘱内容, v_诊疗类别, v_诊疗项目id, v_前提id, v_操作类型, v_执行标记,
         v_执行科室id, v_标本部位, v_开嘱科室id, v_紧急标志, n_病人科室id
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.诊疗项目id = b.Id(+) And a.Id = 医嘱id_In;
  Exception
    When Others Then
      Begin
        v_Error := '医嘱已被删除，不能进行校对。' || Chr(13) || Chr(10) || '这可能是并发操作引起的，请重新读取校对数据。';
        Raise Err_Custom;
      End;
  End;
  If v_状态 <> 1 Then
    v_Error := '医嘱"' || Getadvicetext(医嘱id_In) || '"不是新开的医嘱，不能通过校对。' || Chr(13) || Chr(10) || '这可能是并发操作引起的，请重新读取校对数据。';
    Raise Err_Custom;
  End If;
  --再次检查校对时间的有效性:并发操作 
  If To_Char(v_开嘱时间, 'YYYY-MM-DD HH24:MI') <= To_Char(v_开始时间, 'YYYY-MM-DD HH24:MI') Then
    If To_Char(校对时间_In, 'YYYY-MM-DD HH24:MI') < To_Char(v_开嘱时间, 'YYYY-MM-DD HH24:MI') Then
      v_Error := '医嘱"' || Getadvicetext(医嘱id_In) || '"的校对时间不能小于开嘱时间 ' || To_Char(v_开嘱时间, 'YYYY-MM-DD HH24:MI') || '。' ||
                 Chr(13) || Chr(10) || '这可能是并发操作引起的，请重新读取校对数据。';
      Raise Err_Custom;
    End If;
  Else
    If To_Char(校对时间_In, 'YYYY-MM-DD HH24:MI') < To_Char(v_开始时间, 'YYYY-MM-DD HH24:MI') Then
      v_Error := '医嘱"' || Getadvicetext(医嘱id_In) || '"的校对时间不能小于开始执行时间 ' || To_Char(v_开始时间, 'YYYY-MM-DD HH24:MI') || '。' ||
                 Chr(13) || Chr(10) || '这可能是并发操作引起的，请重新读取校对数据。';
      Raise Err_Custom;
    End If;
  End If;

  --如果要求签名，检查校对时是否有签名(并发取消签名) 
  If 状态_In = 3 Then
    Select Zl_Fun_Getsignpar(Decode(v_前提id, Null, 1, 3), v_开嘱科室id) Into v_Count From Dual;
    If v_Count = 1 Then
      --证书停用或未注册证书不进入签名环节只判断一条数据即可 
      For C In (Select a.是否停用
                From 人员证书记录 A, 人员表 B
                Where a.人员id = b.Id And b.姓名 = v_开嘱医生
                Order By a.注册时间 Desc) Loop
        If Nvl(c.是否停用, 0) = 0 Then
          Select Count(*)
          Into v_Count
          From 病人医嘱状态 A
          Where 操作类型 = 1 And 医嘱id = 医嘱id_In And
                (签名id Is Null And Exists
                 (Select 1
                  From 人员表 R, 人员性质说明 X
                  Where r.Id = x.人员id And r.姓名 = a.操作人员 And x.人员性质 = '护士') And Not Exists
                 (Select 1
                  From 人员表 R, 人员性质说明 Y
                  Where r.Id = y.人员id And r.姓名 = a.操作人员 And y.人员性质 = '医生') Or 签名id Is Not Null Or a.操作人员 <> v_开嘱医生);
          If Nvl(v_Count, 0) = 0 Then
            v_Error := '医嘱"' || Getadvicetext(医嘱id_In) || '"还没有电子签名，不能通过校对。';
            Raise Err_Custom;
          End If;
        End If;
        Exit;
      End Loop;
    End If;
  End If;

  --当前操作人员 
  If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null Then
    v_人员编号 := 操作员编号_In;
    v_人员姓名 := 操作员姓名_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  --因为可能同时：新开->自动校对->互斥自动停止,因此分别-2,-1秒 
  Select Sysdate - 1 / 60 / 60 / 24 Into v_Date From Dual;

  Update 病人医嘱记录
  Set 医嘱状态 = 状态_In, 校对护士 = v_人员姓名, 校对时间 = 校对时间_In
  Where ID = 医嘱id_In Or 相关id = 医嘱id_In;

  Insert Into 病人医嘱状态
    (医嘱id, 操作类型, 操作人员, 操作时间, 操作说明)
    Select ID, 状态_In, v_人员姓名, v_Date, 校对说明_In From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;

  --校对通过时的其它处理 
  If 状态_In = 3 Then
    --自动校对时，自动填写缺省的计价内容 
    If Nvl(自动校对_In, 0) = 1 Then
      --1.变价的计价项目,如果最低限价不为0,则缺省为最低限价,否则不加入;可再手工计价. 
      --2.对于非药嘱药品和在用卫材未定执行科室,发送时会取缺省的,可再手工设置。 
      Select Min(站点) Into v_站点 From 部门表 Where ID = n_病人科室id;
    
      v_Pricegrade := Zl_Get_Pricegrade(v_站点, v_病人id, v_主页id);
      v_药品等级   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);
      v_Pricegrade := Substr(v_Pricegrade, Instr(v_Pricegrade, '|') + 1);
      v_卫材等级   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);
      v_Pricegrade := Substr(v_Pricegrade, Instr(v_Pricegrade, '|') + 1);
      v_普通等级   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);
      For r_Price In c_Price Loop
        --取(检验)医嘱的管码和材料,采集方式以检验项目的为准 
        v_材料id := Null;
        If r_Price.诊疗类别 = 'E' And r_Price.操作类型 = '6' Then
          Begin
            Select c.材料id
            Into v_材料id
            From 病人医嘱记录 A, 诊疗项目目录 B, 采血管类型 C
            Where a.诊疗项目id = b.Id And b.试管编码 = c.编码 And a.相关id = r_Price.Id And Rownum = 1;
          Exception
            When Others Then
              Null;
          End;
        Elsif r_Price.诊疗类别 = 'C' And r_Price.试管编码 Is Not Null Then
          Begin
            Select 材料id Into v_材料id From 采血管类型 Where 编码 = r_Price.试管编码;
          Exception
            When Others Then
              Null;
          End;
        End If;
      
        --判断处理检验试管费用的收取 
        If (Nvl(r_Price.收费方式, 0) = 1 And r_Price.收费类别 = '4' And r_Price.收费项目id = Nvl(v_材料id, 0) Or
           Not (Nvl(r_Price.收费方式, 0) = 1 And r_Price.收费类别 = '4' And Nvl(v_材料id, 0) <> 0)) Then
          Insert Into 病人医嘱计价
            (医嘱id, 收费细目id, 数量, 单价, 从项, 执行科室id, 费用性质, 收费方式)
          Values
            (r_Price.Id, r_Price.收费项目id, r_Price.收费数量, r_Price.单价, r_Price.从属项目, Null, r_Price.费用性质, r_Price.收费方式);
        End If;
      End Loop;
    End If;
  
    --自由录入的临嘱医嘱标记为停止 
    If Nvl(v_期效, 0) = 1 And v_诊疗项目id Is Null Then
      Update 病人医嘱记录
      Set 医嘱状态 = 8, 停嘱时间 = 校对时间_In, 停嘱医生 = v_开嘱医生
      Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
    
      Insert Into 病人医嘱状态
        (医嘱id, 操作类型, 操作人员, 操作时间)
        Select ID, 8, v_人员姓名, Sysdate From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
    End If;
  
    --判断是否开启叮嘱需要执行 
    v_叮嘱执行 := zl_GetSysParameter(288);
    If v_叮嘱执行 = 1 And v_诊疗项目id Is Null Then
      Insert Into 病人医嘱发送
        (医嘱id, 发送号, 记录性质, NO, 记录序号, 发送数次, 发送人, 发送时间, 执行状态, 执行部门id, 计费状态, 首次时间, 末次时间)
      Values
        (医嘱id_In, Nextno('10', '0', '', '1'), '2', Nextno('14', '0', '', '1'), '1', '1', v_人员姓名, Sysdate, '0', v_执行科室id,
         '0', Sysdate, Sysdate);
    End If;
  
    v_Par停嘱 := zl_GetSysParameter(271);
  
    --将同一自动停止互斥组中的病人其它医嘱停止(如果尚未停止) 
    For r_Exclude In c_Exclude Loop
      Select Decode(Sign(r_Exclude.开始执行时间 - v_开始时间), 1, r_Exclude.开始执行时间, v_开始时间)
      Into v_终止时间
      From Dual;
      Select Decode(Sign(r_Exclude.执行终止时间 - v_开始时间), -1, r_Exclude.执行终止时间, v_开始时间)
      Into v_终止时间
      From Dual;
      If v_Par停嘱 = '1' Then
        v_Temp := '自动停止：医嘱互斥。';
      Else
        v_Temp := Null;
      End If;
      Zl_病人医嘱记录_停止(r_Exclude.医嘱id, v_终止时间, v_开嘱医生, 1, 0, 0, Null, Null, v_Temp);
      v_Stopadviceids := v_Stopadviceids || ',' || r_Exclude.医嘱id;
    End Loop;
  
    --对一些特殊医嘱的处理 
    If v_诊疗类别 = 'H' And v_操作类型 = '1' And Nvl(v_期效, 0) = 0 Then
      --校对护理等级时,同步更改病人护理等级 
      If Nvl(v_婴儿, 0) = 0 Then
        --病人当前应处于正常住院状态 
        v_Temp := Null;
        Begin
          Select Decode(状态, 1, '等待入科', 2, '正在转科', 3, '已预出院', Null)
          Into v_Temp
          From 病案主页
          Where 病人id = v_病人id And 主页id = v_主页id;
        Exception
          When Others Then
            Null;
        End;
        If v_Temp Is Not Null Then
          v_Error := '病人当前处于' || v_Temp || '状态,医嘱"' || v_医嘱内容 || '"不能通过校对。';
          Raise Err_Custom;
        End If;
      
        Begin
          --根据收费对照处理，当前医嘱计价表还没有填写 
          --未设置时,不处理；相同时,不处理；有多个时,只取一个。 
          Select a.收费项目id
          Into v_护理等级id
          From 诊疗收费关系 A, 收费项目目录 B
          Where a.收费项目id = b.Id And b.类别 = 'H' And Nvl(b.项目特性, 0) <> 0 And a.诊疗项目id = v_诊疗项目id And Rownum = 1 And
                Not Exists
           (Select 1 From 病案主页 Where 病人id = v_病人id And 主页id = v_主页id And 护理等级id = a.收费项目id);
        Exception
          When Others Then
            Null;
        End;
      End If;
    
      --变动记录的时间加上秒，以便回退操作时区分同一分种的校对、停止等操作 
      v_开始时间 := To_Date(To_Char(v_开始时间, 'yyyy-mm-dd hh24:mi') || To_Char(Sysdate, 'ss'), 'yyyy-mm-dd hh24:mi:ss');
      If v_护理等级id Is Not Null Then
        Zl_病人变动记录_Nurse(v_病人id, v_主页id, v_护理等级id, v_开始时间, v_人员编号, v_人员姓名);
      End If;
    
      --并停止其它护理等级医嘱(护理等级应该都为"持续性"长嘱,且只有一个未停) 
      For r_Nurse In c_Nurse Loop
        Select Decode(Sign(r_Nurse.开始执行时间 - v_开始时间), 1, r_Nurse.开始执行时间, v_开始时间)
        Into v_终止时间
        From Dual;
        Select Decode(Sign(r_Nurse.执行终止时间 - v_开始时间), -1, r_Nurse.执行终止时间, v_开始时间)
        Into v_终止时间
        From Dual;
        If v_Par停嘱 = '1' Then
          v_Temp := '自动停止：护理等级。';
        Else
          v_Temp := Null;
        End If;
        Zl_病人医嘱记录_停止(r_Nurse.医嘱id, v_终止时间, v_开嘱医生, 1, 0, 0, Null, Null, v_Temp);
        Zl_病人医嘱记录_确认停止(r_Nurse.医嘱id, v_终止时间, v_人员姓名, 1);
        v_Stopadviceids := v_Stopadviceids || ',' || r_Nurse.医嘱id;
      End Loop;
    Elsif v_诊疗类别 = 'Z' And v_操作类型 In ('9', '10') And Nvl(v_期效, 0) = 0 And Nvl(v_婴儿, 0) = 0 Then
      --病重病危医嘱：9-病重;10-病危 
      --停止相同医嘱 
      For r_Patistate In c_Patistate Loop
        Select Decode(Sign(r_Patistate.开始执行时间 - v_开始时间), 1, r_Patistate.开始执行时间, v_开始时间)
        Into v_终止时间
        From Dual;
        Select Decode(Sign(r_Patistate.执行终止时间 - v_开始时间), -1, r_Patistate.执行终止时间, v_开始时间)
        Into v_终止时间
        From Dual;
        If v_Par停嘱 = '1' Then
          If v_操作类型 = '9' Then
            v_Temp := '自动停止：病重医嘱。';
          Else
            v_Temp := '自动停止：病危医嘱。';
          End If;
        Else
          v_Temp := Null;
        End If;
        Zl_病人医嘱记录_停止(r_Patistate.医嘱id, v_终止时间, v_开嘱医生, 1, 0, 0, Null, Null, v_Temp);
        v_Stopadviceids := v_Stopadviceids || ',' || r_Patistate.医嘱id;
      End Loop;
    
      b_Message.Zlhis_Patient_005(v_病人id, v_主页id);
    
      --产生病情变动 
      Open c_Oldinfo; --必须在处理之前先打开 
      Fetch c_Oldinfo
        Into r_Oldinfo;
      Open c_Endinfo;
      Fetch c_Endinfo
        Into r_Endinfo;
      If c_Endinfo%RowCount = 0 Then
        Close c_Endinfo;
        v_Error := '未发现该病人当前有效的变动记录！';
        Raise Err_Custom;
      End If;
      Select Count(*)
      Into v_Count
      From 病人变动记录
      Where 病人id = v_病人id And 主页id = v_主页id And 开始时间 Is Null And 终止时间 Is Null;
      If v_Count > 0 Then
        v_Error := '病人当前处于转科状态，请先办理转科确认或者取消转科状态。';
        Raise Err_Custom;
      End If;
    
      Update 病案主页
      Set 当前病况 = Decode(v_操作类型, '9', '重', '10', '危')
      Where 病人id = v_病人id And 主页id = v_主页id;
    
      --取消上次变动 
      If r_Oldinfo.终止时间 Is Not Null Then
        v_变动终止时间 := r_Oldinfo.终止时间;
        v_变动终止原因 := r_Oldinfo.终止原因;
        v_变动终止人员 := r_Oldinfo.终止人员;
        --取消上次变动 
        Update 病人变动记录
        Set 终止时间 = v_开始时间, 终止原因 = 13, 终止人员 = v_人员姓名, 上次计算时间 = Null
        Where 病人id = v_病人id And 主页id = v_主页id And 终止时间 = v_变动终止时间 And 终止原因 = v_变动终止原因;
        --更新将来的记录如果有停止到将来的则删除上次计算时间 
        Update 病人变动记录
        Set 病情 = Decode(v_操作类型, '9', '重', '10', '危'), 上次计算时间 = Null
        Where 病人id = v_病人id And 主页id = v_主页id And 开始时间 > v_开始时间;
      Else
        Update 病人变动记录
        Set 终止时间 = v_开始时间, 终止原因 = 13, 终止人员 = v_人员姓名,
            上次计算时间 = Decode(Sign(Nvl(上次计算时间, v_开始时间) - v_开始时间), 1, Null, 上次计算时间)
        Where 病人id = v_病人id And 主页id = v_主页id And 终止时间 Is Null;
      End If;
    
      While c_Oldinfo%Found Loop
        Insert Into 病人变动记录
          (ID, 病人id, 主页id, 开始时间, 开始原因, 附加床位, 病区id, 科室id, 护理等级id, 床位等级id, 床号, 责任护士, 经治医师, 主治医师, 主任医师, 病情, 操作员编号, 操作员姓名,
           终止时间, 终止原因, 终止人员)
        Values
          (病人变动记录_Id.Nextval, v_病人id, v_主页id, v_开始时间, 13, r_Oldinfo.附加床位, r_Oldinfo.病区id, r_Oldinfo.科室id,
           r_Oldinfo.护理等级id, r_Oldinfo.床位等级id, r_Oldinfo.床号, r_Oldinfo.责任护士, r_Oldinfo.经治医师, r_Oldinfo.主治医师,
           r_Oldinfo.主任医师, Decode(v_操作类型, '9', '重', '10', '危'), v_人员编号, v_人员姓名, v_变动终止时间, v_变动终止原因, v_变动终止人员);
      
        Fetch c_Oldinfo
          Into r_Oldinfo;
      End Loop;
    
      Close c_Oldinfo;
      Close c_Endinfo;
    Elsif v_诊疗类别 = 'Z' And v_操作类型 = '12' And Nvl(v_期效, 0) = 0 And Nvl(v_婴儿, 0) = 0 Then
      --记录入出量的医嘱，互斥 
      For r_Patiio In c_Patiio Loop
        Select Decode(Sign(r_Patiio.开始执行时间 - v_开始时间), 1, r_Patiio.开始执行时间, v_开始时间)
        Into v_终止时间
        From Dual;
        Select Decode(Sign(r_Patiio.执行终止时间 - v_开始时间), -1, r_Patiio.执行终止时间, v_开始时间)
        Into v_终止时间
        From Dual;
        Zl_病人医嘱记录_停止(r_Patiio.医嘱id, v_终止时间, v_开嘱医生, 1);
        v_Stopadviceids := v_Stopadviceids || ',' || r_Patiio.医嘱id;
      End Loop;
    Elsif (v_诊疗类别 = 'Z' And v_操作类型 In ('3', '4', '5', '6', '11', '14') And
          (v_操作类型 <> '14' Or v_操作类型 = '14' And v_执行标记 = 1)) Or (v_诊疗类别 = 'F' And v_执行标记 = 1) Then
      v_Count := 0;
      If v_操作类型 = '4' Or v_操作类型 = '14' Or v_诊疗类别 = 'F' Then
        --保持与以前校对时相同的处理 
        If Nvl(v_婴儿, 0) = 0 Then
          v_Count := 1;
        End If;
      Else
        --这几个特殊医嘱在校对中停止医嘱是新加的内容，保持与发送中相同的处理 
        v_Count := 1;
        If Nvl(v_婴儿, 0) = 0 Then
          v_婴儿 := -1;
        Else
          v_婴儿 := Nvl(v_婴儿, 0);
        End If;
      End If;
      If v_Count = 1 Then
        If v_诊疗类别 = 'F' And v_执行标记 = 1 Then
          --在手术当天(取整)停止 
          v_开始时间 := Trunc(To_Date(v_标本部位, 'yyyy-mm-dd hh24:mi:ss'));
        End If;
      
        --几个特殊医嘱校对时停止前面的长嘱,在医嘱开始时终止：3-转科;4-术后;5-出院;6-转院,11-死亡,14-术前 
        For r_Needstop In c_Needstop(v_病人id, v_主页id, v_婴儿, v_开始时间) Loop
          Select Decode(Sign(开始执行时间 - v_开始时间), 1, 开始执行时间, v_开始时间)
          Into v_停止时间
          From 病人医嘱记录
          Where ID = r_Needstop.Id;
          Update 病人医嘱记录
          Set 医嘱状态 = 8, 执行终止时间 = v_停止时间, 停嘱时间 = 校对时间_In, 停嘱医生 = v_开嘱医生
          Where ID = r_Needstop.Id;
        
          Insert Into 病人医嘱状态
            (医嘱id, 操作类型, 操作人员, 操作时间)
            Select ID, 8, v_人员姓名, 校对时间_In From 病人医嘱记录 Where ID = r_Needstop.Id;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Needstop.Id;
        End Loop;
        --已停止未确认的长嘱,终止时间在医嘱开始后的,调前其终止时间(同时多个特殊医嘱的情况) 
        For r_Havestop In c_Havestop(v_病人id, v_主页id, v_婴儿, v_开始时间) Loop
          Update 病人医嘱记录
          Set 执行终止时间 = Decode(Sign(开始执行时间 - v_开始时间), 1, 开始执行时间, v_开始时间), 停嘱时间 = 校对时间_In, 停嘱医生 = v_开嘱医生
          Where ID = r_Havestop.Id;
        
          --不修改停止医嘱的操作人员，因为停止时，医生可能已进行电子签名 
          Update 病人医嘱状态 Set 操作时间 = 校对时间_In Where 医嘱id = r_Havestop.Id And 操作类型 = 8;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Havestop.Id;
        End Loop;
        --处理长期备用医嘱(没有执行（发送）过的标记未用） 
        Update 病人医嘱记录
        Set 执行标记 = -1
        Where 病人id = v_病人id And 主页id = v_主页id And 医嘱期效 = 0 And 执行频次 = '必要时' And 上次执行时间 Is Null And 医嘱状态 In (3, 5, 6, 7) And
              执行标记 <> -1;
        --如果是转院转科死亡出院医嘱同时处理临时备用医嘱。 
        If v_操作类型 In ('3', '5', '6', '11') Then
          Update 病人医嘱记录
          Set 执行标记 = -1
          Where 病人id = v_病人id And 主页id = v_主页id And 医嘱期效 = 1 And 执行频次 = '需要时' And 医嘱状态 = 3 And 执行标记 <> -1;
        End If;
      End If;
    Elsif v_诊疗类别 = 'Z' And v_操作类型 = '2' Then
      --对留观病人下达入院通知; 
      --预约登记的条件：1.当前无预约,2.当前是门诊留观病人（在院时也允许，因为需要先预约,入院接收时检查了必须出院后才能接收） 
      Select Count(*) Into v_Count From 病案主页 Where 病人id = v_病人id And Nvl(主页id, 0) = 0;
      If v_Count = 0 Then
        Select Count(*) Into v_Count From 病案主页 Where 病人id = v_病人id And 主页id = v_主页id And 病人性质 <> 1;
      End If;
      If v_Count = 0 Then
        Open c_Pati(v_病人id);
        Fetch c_Pati
          Into r_Pati;
        Close c_Pati;
      
        v_入院方式 := Null;
        If v_紧急标志 = 1 Then
          v_入院方式 := '急诊';
        End If;
      
        Zl_入院病案主页_Insert(1, 0, r_Pati.病人id, r_Pati.住院号, Null, r_Pati.姓名, r_Pati.性别, r_Pati.年龄, r_Pati.费别, r_Pati.出生日期,
                         r_Pati.国籍, r_Pati.民族, r_Pati.学历, r_Pati.婚姻状况, r_Pati.职业, r_Pati.身份, r_Pati.身份证号, r_Pati.出生地点,
                         r_Pati.家庭地址, r_Pati.家庭地址邮编, r_Pati.家庭电话, r_Pati.户口地址, r_Pati.户口地址邮编, r_Pati.联系人姓名, r_Pati.联系人关系,
                         r_Pati.联系人地址, r_Pati.联系人电话, r_Pati.工作单位, r_Pati.合同单位id, r_Pati.单位电话, r_Pati.单位邮编, r_Pati.单位开户行,
                         r_Pati.单位帐号, r_Pati.担保人, r_Pati.担保额, r_Pati.担保性质, v_执行科室id, Null, Null, v_入院方式, Null, Null,
                         v_开嘱医生, r_Pati.籍贯, r_Pati.区域, v_开始时间, Null, Null, r_Pati.医疗付款方式, Null, Null, Null, Null, Null,
                         Null, r_Pati.险类, v_人员编号, v_人员姓名, 0, Null, Null, 0);
      End If;
    End If;
    --医嘱停止消息的处理 
    If v_Stopadviceids Is Not Null Then
      v_Stopadviceids := Substr(v_Stopadviceids, 2);
      b_Message.Zlhis_Cis_002(v_病人id, v_主页id, Null, v_Stopadviceids);
      Select Max(a.Id)
      Into n_标记
      From 病人医嘱记录 A
      Where a.Id In (Select Column_Value From Table(f_Num2list(v_Stopadviceids))) And a.医嘱期效 = 0 And a.医嘱状态 = 8 And
            Nvl(a.执行标记, 0) <> -1 And a.病人来源 <> 3;
      If n_标记 Is Not Null Then
        Select Max(a.Id)
        Into n_Adviceid
        From 病人医嘱记录 A
        Where a.Id In (Select Column_Value From Table(f_Num2list(v_Stopadviceids))) And a.紧急标志 = 1 And a.医嘱期效 = 0 And
              a.医嘱状态 = 8 And Nvl(a.执行标记, 0) <> -1 And a.病人来源 <> 3;
        If n_Adviceid Is Not Null Then
          n_Adviceid := n_标记;
          Select Nvl(Max(0), 2)
          Into n_标记
          From 业务消息清单 A
          Where a.病人id = v_病人id And a.就诊id = v_主页id And a.类型编码 = 'ZLHIS_CIS_002' And a.优先程度 = 2 And a.是否已阅 = 0;
        Else
          Select Nvl(Max(0), 1)
          Into n_标记
          From 业务消息清单 A
          Where a.病人id = v_病人id And a.就诊id = v_主页id And a.类型编码 = 'ZLHIS_CIS_002' And a.是否已阅 = 0;
        End If;
        If n_标记 > 0 Then
          For R In (Select a.病人性质 As 性质, a.出院科室id As 科室id, a.当前病区id As 病区id
                    From 病案主页 A
                    Where a.病人id = v_病人id And a.主页id = v_主页id) Loop
            Zl_业务消息清单_Insert(v_病人id, v_主页id, r.科室id, r.病区id, r.性质, '有新停止医嘱。', '0010', 'ZLHIS_CIS_002', n_Adviceid, n_标记,
                             0, Null, r.病区id);
          End Loop;
        End If;
      End If;
    End If;
  End If;

  --病区执行医嘱校对消息
  For R In (Select a.Id, a.病人id, a.主页id
            From 病人医嘱记录 A
            Where (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In) And Exists
             (Select 1 From 部门性质说明 B Where b.部门id = a.执行科室id And b.工作性质 = '护理')) Loop
    b_Message.Zlhis_Cis_012(r.病人id, r.主页id, r.Id);
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱记录_校对;
/

------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0010' Where 编号=&n_System;
Commit;