--���ؼ�¼���⣬�ô�����
-- Create table
create table ���沵�ؼ�¼
(
  ����ID   number(18) not null,
  ����ʱ�� date not null,
  ������   varchar2(10),
  ����ԭ�� varchar2(100)
)
;
-- Create/Recreate primary, unique and foreign key constraints 
alter table ���沵�ؼ�¼
  add constraint ���沵�ؼ�¼_FK_����ID foreign key (����ID)
  references ���˲�����¼ (ID) on delete cascade;
alter table ���沵�ؼ�¼
  add constraint ���沵�ؼ�¼_PK_����ID primary key (����ID, ����ʱ��);
-- Create/Recreate indexes 
create index ���沵�ؼ�¼_IX_����ʱ�� on ���沵�ؼ�¼ (����ʱ��);

Create Or Replace Trigger Zltg_�������ؼ�¼
  Before Update On ����ҽ������
  For Each Row
Declare
  -- local variables here
  v_������� Varchar2(10);
  v_Temp     Varchar2(255);
  v_��Ա���� ���沵�ؼ�¼.������%Type;
Begin
  Begin
    Select A.������� Into v_������� From ����ҽ����¼ A Where A.ID = :Old.ҽ��id;
  Exception
    When Others Then
      v_������� := Null;
  End;
  If v_������� <> 'D' Or :New.ִ��״̬ <> 3 Or :New.ִ�й��� <> 5 Then
    Return;
  End If;

  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  Begin
    Insert Into ���沵�ؼ�¼
      (����id, ����ʱ��, ������, ����ԭ��)
    Values
      (:Old.����id, To_Date(To_Char(Sysdate, 'yyyy-mm-dd hh24:mi'), 'yyyy-mm-dd hh24:mi:ss'), v_��Ա����, Null);
  Exception
    When Others Then
      Return;
  End;
End Zltg_�������ؼ�¼;
/
