--- TABLE
CREATE TABLE 临床业务提醒(
    ID	Number(18),
    病人ID	Number(18),
    主页ID	Number(5),
    类别	Number(2),
    原因	Number(3),
    业务ID	Number(18),
    提醒标题	Varchar2(100),
    提醒内容	Varchar2(2000),
    登记时间	Date,
    登记人	Varchar2(20))
    TABLESPACE zl9CisRec
    initrans 20;

--- Constraint 
Alter Table 临床业务提醒 Add Constraint 临床业务提醒_PK  Primary Key (ID)  Using Index Tablespace zl9indexhis;
Alter Table 临床业务提醒 Add Constraint 临床业务提醒_FK_病人ID Foreign Key (病人ID) References 病人信息(病人ID) On Delete Cascade;
Alter Table 临床业务提醒 Add Constraint 临床业务提醒_FK_主页ID Foreign Key (病人ID,主页ID) References 病案主页(病人ID,主页ID) On Delete Cascade;

---- INDEX
Create Index 临床业务提醒_IX_登记时间 On 临床业务提醒(登记时间) Tablespace zl9indexhis
/
--- Sequence
Create Sequence 临床业务提醒_ID Start With 1;


----- Program
Create Or Replace Procedure Zl_临床业务提醒_Edit
(
  操作_In   In Number, ---- 1-增加
  Id_In     In 临床业务提醒.Id%Type,
  病人id_In In 临床业务提醒.病人id%Type := Null,
  主页id_In In 临床业务提醒.主页id%Type := Null,
  类别_In   In 临床业务提醒.类别%Type := Null, ---- 1-医嘱，2-病历，3-医技报告，4-医技安排
  原因_In   In 临床业务提醒.原因%Type := Null, ---- 101-新开，102-新停,103-作废，201-病历审核，301-检验危急指标，401-手术安排
  业务id_In In 临床业务提醒.业务id%Type := Null,
  标题_In   In 临床业务提醒.提醒标题%Type := Null,
  内容_In   In 临床业务提醒.提醒内容%Type := Null,
  登记人_In In 临床业务提醒.登记人%Type := Null
) Is
Begin
  If 操作_In = 1 Then
    Insert Into 临床业务提醒
      (ID, 病人id, 主页id, 类别, 原因, 业务id, 提醒标题, 提醒内容, 登记人, 登记时间)
    Values
      (Id_In, 病人id_In, 主页id_In, 类别_In, 原因_In, 业务id_In, 标题_In, 内容_In, 登记人_In, Sysdate);
  End If;

  If 操作_In = 2 Then
    Delete 临床业务提醒 Where ID = Id_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_临床业务提醒_Edit;
/