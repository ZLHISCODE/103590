--驳回记录问题，用触发器
-- Create table
create table 报告驳回记录
(
  报告ID   number(18) not null,
  驳回时间 date not null,
  驳回人   varchar2(10),
  驳回原因 varchar2(100)
)
;
-- Create/Recreate primary, unique and foreign key constraints 
alter table 报告驳回记录
  add constraint 报告驳回记录_FK_报告ID foreign key (报告ID)
  references 病人病历记录 (ID) on delete cascade;
alter table 报告驳回记录
  add constraint 报告驳回记录_PK_报告ID primary key (报告ID, 驳回时间);
-- Create/Recreate indexes 
create index 报告驳回记录_IX_驳回时间 on 报告驳回记录 (驳回时间);

Create Or Replace Trigger Zltg_产生驳回记录
  Before Update On 病人医嘱发送
  For Each Row
Declare
  -- local variables here
  v_诊疗类别 Varchar2(10);
  v_Temp     Varchar2(255);
  v_人员姓名 报告驳回记录.驳回人%Type;
Begin
  Begin
    Select A.诊疗类别 Into v_诊疗类别 From 病人医嘱记录 A Where A.ID = :Old.医嘱id;
  Exception
    When Others Then
      v_诊疗类别 := Null;
  End;
  If v_诊疗类别 <> 'D' Or :New.执行状态 <> 3 Or :New.执行过程 <> 5 Then
    Return;
  End If;

  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  Begin
    Insert Into 报告驳回记录
      (报告id, 驳回时间, 驳回人, 驳回原因)
    Values
      (:Old.报告id, To_Date(To_Char(Sysdate, 'yyyy-mm-dd hh24:mi'), 'yyyy-mm-dd hh24:mi:ss'), v_人员姓名, Null);
  Exception
    When Others Then
      Return;
  End;
End Zltg_产生驳回记录;
/
