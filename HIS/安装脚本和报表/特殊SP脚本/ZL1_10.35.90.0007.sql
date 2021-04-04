----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--124650:��ҵ��,2018-04-20,�������������Ƿ��������
Alter Table �������� Add �Ƿ���� Number(1) Default 0;

------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------

--123971:������,2018-04-16,ѪҺ���պ������ִ�еǼ�
Insert Into Zlparameters
  (Id, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
Values
  (Zlparameters_Id.Nextval, &n_System, -null, -null, -null, -null, -null, 0, 0, 301, 'ѪҺ���պ������ִ�еǼ�', 1, 1,
   '����Ѫ��ϵͳʱҽ����ԱȡѪ���Һ��Ƿ���Ҫ����ѪҺ���պ˶Ի��ڲ����������Ѫִ������Ǽ�', '0-������н��ջ��ڼ��ɽ���ִ������Ǽ�,1-�������ѪҺ���պ˶Ի��ڲ��������ִ������Ǽ�',
   'ֻ������236�Ų���[����Ѫ�����ϵͳ]�����������ô˲������Լ����ݴ˲��������Ƿ���Ҫ����ѪҺ���ջ���', 'ҽԺ�ɸ��ݾ���ҵ�����ģʽ��������', Null);


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------





-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--122998:������,2018-04-20,����ƽ̨��Ϣê��
Create Or Replace Procedure Zl_�ֵ����_Execute(Sql_In In Varchar2) Is
  --һ��������SQL��䣬ע�����ǰһ��Ҫ�������߼��ϡ�
  --��UPDATE ZLHIS.���㷽ʽ SET ȱʡ��־=0
  v_Rulesql Varchar2(8000);
  n_Pos     Number;
  v_Tmp     Varchar2(4000);
  v_Tab     Varchar2(100);
  v_Sql     Varchar2(8000);
  n_Count   Number;
  v_Owner   Varchar2(100);
  v_Code    Varchar2(100);
  v_Tmp1    Varchar2(8000);

  v_Err Varchar2(500);
  Err_Custom Exception;
Begin
  -------------------------
  --SQLУ��
  ----------------------
  --1.��ʽ��SQL���
  v_Rulesql := Upper(Sql_In);
  v_Rulesql := Trim(Replace(v_Rulesql, Chr(10), ' '));
  v_Rulesql := Trim(Replace(v_Rulesql, Chr(13), ' '));
  --��˫�ո��滻Ϊ���ո�
  While Instr(v_Rulesql, '  ', 1) > 0 Loop
    v_Rulesql := Trim(Replace(v_Rulesql, '  ', ''));
  End Loop;
  v_Rulesql := Trim(v_Rulesql);
  --2�������Ǳ�׼��Insert,uPdate,Delete���
  n_Pos := Instr(v_Rulesql, ' ');
  --���ֱ�׼��DML���һ�������ո񣬲��ҿո��λ���ǵ���λ
  If n_Pos = 0 Or n_Pos <> 7 Then
    v_Err := '�﷨���ʧ�ܣ��﷨�������䲻��DML��䣡';
    Raise Err_Custom;
  End If;
  v_Tmp := Trim(Substr(v_Rulesql, 1, n_Pos));
  v_Sql := Trim(Substr(v_Rulesql, n_Pos));

  If v_Tmp = 'INSERT' Or v_Tmp = 'DELETE' Or v_Tmp = 'UPDATE' Then
    --Insert ��������Insert into tableName(col1,col2,...) values(val1,val2,...)
    If v_Tmp = 'INSERT' Then
      --Insert �����Insert into tableName(col1,col2,...) values(val1,val2,...)
      If v_Rulesql Like 'INSERT INTO %(%)%VALUES%(%)' Or v_Rulesql Like 'INSERT INTO %(%)%SELECT % FROM DUAL' Then
        --��ȡINTO TableName �ֶ�
        n_Pos := Instr(v_Sql, '(');
        v_Tab := Trim(Substr(v_Sql, 1, n_Pos - 1));
        --��ȡOWNER.Table�ֶ�
        n_Pos := Instr(v_Tab, ' ');
        v_Tab := Trim(Substr(v_Tab, n_Pos));
      Else
        v_Err := '�﷨���ʧ�ܣ�Insert����﷨����';
        Raise Err_Custom;
      End If;
    Elsif v_Tmp = 'UPDATE' Then
      --Update ��������Update tableName Set COl1=val1,.....
      If v_Rulesql Like 'UPDATE % SET %' Then
        --��ȡOWNER.Table�ֶ�
        n_Pos := Instr(v_Sql, 'SET');
        v_Tab := Trim(Substr(v_Sql, 1, n_Pos - 1));
      Else
        v_Err := '�﷨���ʧ�ܣ�UPDATE����﷨����';
        Raise Err_Custom;
      End If;
    Elsif v_Tmp = 'DELETE' Then
      --DELETE ��������DELETE [From] tableName ,DELETE [From] tableName Where ..........
      If v_Rulesql Like 'DELETE % WHERE %' Then
        --delete��京FROM
        If v_Rulesql Like 'DELETE FROM % WHERE %' Then
          --��ȡFROM TableName �ֶ�
          n_Pos := Instr(v_Sql, 'WHERE');
          v_Tab := Trim(Substr(v_Sql, 1, n_Pos - 1));
          --��ȡOWNER.Table�ֶ�
          n_Pos := Instr(v_Tab, ' ');
          v_Tab := Trim(Substr(v_Tab, n_Pos));
          --delete��䲻��FROM
        Else
          --��ȡOWNER.Table�ֶ�
          n_Pos := Instr(v_Sql, 'WHERE');
          v_Tab := Trim(Substr(v_Sql, 1, n_Pos - 1));
        End If;
      Elsif v_Rulesql Like 'DELETE % ' Then
        --delete��京FROM
        If v_Rulesql Like 'DELETE FROM %' Then
          --��ȡOWNER.Table�ֶ�
          n_Pos := Instr(v_Tab, ' ');
          v_Tab := Trim(Substr(v_Sql, n_Pos));
          --delete��䲻��FROM
        Else
          --��ȡOWNER.Table�ֶ�
          v_Tab := v_Sql;
        End If;
      Else
        v_Err := '�﷨���ʧ�ܣ�DELETE����﷨����';
        Raise Err_Custom;
      End If;
    End If;
  Else
    v_Err := '�﷨���ʧ�ܣ���������DML��䡣';
    Raise Err_Custom;
  End If;
  --��ȡ�������Լ�ϵͳ��
  --û�д�������ʱĬ��Ϊ��׼��
  v_Tab := Trim(v_Tab);
  If v_Tab || ' ' <> ' ' Then
    n_Pos := Instr(v_Tab, '.');
    If n_Pos <> 0 Then
      v_Owner := Substr(v_Tab, 1, n_Pos - 1);
      v_Tab   := Substr(v_Tab, n_Pos + 1);
    Else
      Select Max(a.������) Into v_Owner From zlSystems A Where a.��� = 100;
    End If;
  End If;

  --DML�������ı������ZLBASECODE�еķǹ̶���
  Select Count(1)
  Into n_Count
  From zlBaseCode
  Where �̶� = 0 And ���� = v_Tab And ϵͳ In (Select a.��� From zlSystems A Where a.������ = v_Owner);

  If n_Count = 0 Then
    v_Err := '��' || v_Tab || '���ǵ�ǰϵͳ���еķǹ̶���';
    Raise Err_Custom;
  End If;

  If v_Tab = '���Ƽ������' Then
    --��������ֵ
    If v_Tmp = 'INSERT' Then
      n_Pos  := Instr(v_Sql, 'VALUES');
      v_Tmp1 := Substr(v_Sql, n_Pos);
      n_Pos  := Instr(v_Tmp1, ',');
      v_Tmp1 := Substr(v_Tmp1, 1, n_Pos - 1);
      n_Pos  := Instr(v_Tmp1, '(');
      v_Tmp1 := Substr(v_Tmp1, n_Pos + 1);
      v_Code := Trim(Replace(v_Tmp1, '''', ''));
    Else
      n_Pos  := Instr(v_Sql, 'WHERE');
      v_Tmp1 := Substr(v_Sql, n_Pos);
      n_Pos  := Instr(v_Tmp1, '=');
      v_Tmp1 := Substr(v_Tmp1, n_Pos + 1);
      v_Code := Trim(Replace(v_Tmp1, '''', ''));
    End If;
  End If;

  --��Ϊ���ܶ�����װ������ֻ�ܶ�ִ̬��  
  If v_Tmp = 'DELETE' Then
    If v_Tab = '���Ƽ������' Then
      --ɾ����¼
      For R In (Select a.����, a.����, a.����, a.������ From ���Ƽ������ A Where a.���� = v_Code) Loop
        --b_Message.Zlhis_Dictpacs_003(r.����, r.����, r.����, r.������);      
        Begin
          Execute Immediate 'call b_Message.Zlhis_Dictpacs_003(:1,:2,:3,:4)'
            Using r.����, r.����, r.����, r.������;
        Exception
          When Others Then
            Null;
        End;
      End Loop;
    Elsif v_Tab = '���Ƽ���걾' Then
      --ɾ����¼
      For R In (Select a.����, a.����, a.����, a.�����Ա� From ���Ƽ���걾 A Where a.���� = v_Code) Loop
        --b_Message.Zlhis_Dictlis_006(r.����, r.����, r.����, r.�����Ա�);          
        Begin
          Execute Immediate 'call b_Message.Zlhis_Dictlis_006(:1,:2,:3,:4)'
            Using r.����, r.����, r.����, r.�����Ա�;
        Exception
          When Others Then
            Null;
        End;
      End Loop;
    End If;
  End If;

  Execute Immediate v_Rulesql;

  If v_Tab = '���Ƽ������' Then
    If v_Tmp = 'INSERT' Or v_Tmp = 'UPDATE' Then
      For R In (Select a.����, a.����, a.����, a.������ From ���Ƽ������ A Where a.���� = v_Code) Loop
        If v_Tmp = 'INSERT' Then
          --b_Message.Zlhis_Dictpacs_001(r.����, r.����, r.����, r.������);         
          Begin
            Execute Immediate 'call b_Message.Zlhis_Dictpacs_001(:1,:2,:3,:4)'
              Using r.����, r.����, r.����, r.������;
          Exception
            When Others Then
              Null;
          End;
        Else
          --b_Message.Zlhis_Dictpacs_002(r.����, r.����, r.����, r.������);          
          Begin
            Execute Immediate 'call b_Message.Zlhis_Dictpacs_002(:1,:2,:3,:4)'
              Using r.����, r.����, r.����, r.������;
          Exception
            When Others Then
              Null;
          End;
        End If;
      End Loop;
    End If;
  Elsif v_Tab = '���Ƽ���걾' Then
    If v_Tmp = 'INSERT' Or v_Tmp = 'UPDATE' Then
      For R In (Select a.����, a.����, a.����, a.�����Ա� From ���Ƽ���걾 A Where a.���� = v_Code) Loop
        If v_Tmp = 'INSERT' Then
          --b_Message.Zlhis_Dictlis_004(r.����, r.����, r.����, r.�����Ա�);        
          Begin
            Execute Immediate 'call b_Message.Zlhis_Dictlis_004(:1,:2,:3,:4)'
              Using r.����, r.����, r.����, r.�����Ա�;
          Exception
            When Others Then
              Null;
          End;
        Else
          --b_Message.Zlhis_Dictlis_005(r.����, r.����, r.����, r.�����Ա�);        
          Begin
            Execute Immediate 'call b_Message.Zlhis_Dictlis_005(:1,:2,:3,:4)'
              Using r.����, r.����, r.����, r.�����Ա�;
          Exception
            When Others Then
              Null;
          End;
        End If;
      End Loop;
    End If;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ֵ����_Execute;
/

--124359:������,2018-04-17,�����Ŀҽ������λִ�л�ȡ��ִ��
Create Or Replace Procedure Zl_סԺҽ��ִ��_Finish
(
  ҽ��id_In     In ����ҽ��ִ��.ҽ��id%Type,
  ���ͺ�_In     In ����ҽ��ִ��.���ͺ�%Type,
  ����ִ��_In   In Number,
  ����Ա���_In In ��Ա��.���%Type,
  ����Ա����_In In ��Ա��.����%Type,
  ��id_In       In ����ҽ��ִ��.ҽ��id%Type,
  �������_In   In ����ҽ����¼.�������%Type,
  ִ�в���id_In In סԺ���ü�¼.ִ�в���id%Type
  --������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID��
  --      ����ִ��_In=����ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ
  --ִ�в���id_In=������ָ��ִ�в��ŵķ��ã���������0ʱ������ִ�в���
) Is
  --ҽ����صķ��õ���
  Cursor c_No Is
    Select a.No || ':' || a.��¼����
    From ����ҽ������ A, ����ҽ����¼ B
    Where a.���ͺ� + 0 = ���ͺ�_In And a.ҽ��id = b.Id And b.������� = �������_In And (b.Id = ��id_In Or b.���id = ��id_In)
    Union
    Select NO || ':' || ��¼����
    From ����ҽ������
    Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;

  Cursor c_Noone Is
    Select NO || ':' || ��¼����
    From ����ҽ������
    Where ���ͺ� + 0 = ���ͺ�_In And ҽ��id = ҽ��id_In
    Union
    Select NO || ':' || ��¼����
    From ����ҽ������
    Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;

  r_No       t_Strlist;
  r_No_Stuff t_Strlist;
  r_Finish   t_Numlist;

  Cursor c_Finish(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select Distinct a.Id, a.�շ����, a.�շ�ϸĿid
           From סԺ���ü�¼ A, ����ҽ����¼ B,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
                  From Table(r_No)) N
           Where a.ҽ����� = b.Id And (b.Id = ��id_In Or b.���id = ��id_In) And a.No = n.No And a.��¼���� = n.��¼���� And
                 a.��¼״̬ In (0, 1, 3) And a.ִ��״̬ <> 1 And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)) A, �������� B
    Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);

  Cursor c_Finishone(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select a.Id, a.�շ����, a.�շ�ϸĿid
           From סԺ���ü�¼ A,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
                  From Table(r_No)) N
           Where a.ҽ����� = ҽ��id_In And a.No = n.No And a.��¼���� = n.��¼���� And a.��¼״̬ In (0, 1, 3) And a.ִ��״̬ <> 1 And
                 (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)) A, �������� B
    Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);

  --ִ���а����������õ�δ������ʱ�����ݲ��������Ƿ��Զ�����
  --��������ҽ��Ŀǰ�����ڵ��������ִ�е����
  Cursor c_Stuff(r_No t_Strlist) Is
    Select /*+ RULE */
     b.Id, Decode(d.��ֵ����, 1, a.ִ�в���id, b.�ⷿid) As �ⷿid
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ����ҽ����¼ C, �������� D,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
           From Table(r_No)) N
    Where d.����id = a.�շ�ϸĿid And a.Id = b.����id And Mod(b.��¼״̬, 3) = 1 And b.����� Is Null And b.�ⷿid Is Not Null And
          a.�շ���� = '4' And a.��¼״̬ = 1 And a.ҽ����� = c.Id And c.������� = �������_In And (c.Id = ��id_In Or c.���id = ��id_In) And
          a.No = n.No And a.��¼���� = n.��¼���� And b.���� In (25, 26) And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And
          (c.Id = ҽ��id_In And ����ִ��_In = 1 Or Nvl(����ִ��_In, 0) <> 1)
    Order By b.�ⷿid, b.ҩƷid;

  --ִ���а�����δ��ҩƷ������ִ�е��Զ���ҩ
  Cursor c_Drug(r_No t_Strlist) Is
    Select /*+ RULE */
     b.Id, b.�ⷿid
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ����ҽ����¼ C, ������ҳ D,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
           From Table(r_No)) N
    Where a.Id = b.����id And Mod(b.��¼״̬, 3) = 1 And b.����� Is Null And b.�ⷿid = d.��ǰ����id And a.�շ���� In ('5', '6', '7') And
          a.��¼״̬ = 1 And a.ҽ����� = c.Id And c.������� = �������_In And c.����id = d.����id And c.��ҳid = d.��ҳid And
          (c.Id = ��id_In Or c.���id = ��id_In) And a.No = n.No And a.��¼���� = n.��¼���� And
          (c.Id = ҽ��id_In And ����ִ��_In = 1 Or Nvl(����ִ��_In, 0) <> 1)
    Order By b.�ⷿid, b.ҩƷid;

  --δ��˵ķ�����(����ҩƷ������)
  Cursor c_Verify(r_No t_Strlist) Is
    Select /*+ RULE */
    Distinct a.No, a.���
    From סԺ���ü�¼ A, ����ҽ����¼ C,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
           From Table(r_No)) N
    Where a.���ʷ��� = 1 And a.��¼״̬ = 0 And a.�۸񸸺� Is Null And a.ҽ����� = c.Id And (c.Id = ��id_In Or c.���id = ��id_In) And
          a.No = n.No And a.��¼���� = n.��¼���� And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)
    Order By NO, ���;

  Cursor c_Verifyone(r_No t_Strlist) Is
    Select /*+ RULE */
     a.No, a.���
    From סԺ���ü�¼ A,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
           From Table(r_No)) N
    Where a.���ʷ��� = 1 And a.��¼״̬ = 0 And a.�۸񸸺� Is Null And a.ҽ����� + 0 = ҽ��id_In And a.No = n.No And a.��¼���� = n.��¼���� And
          (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)
    Order By NO, ���;

  v_No   ����ҽ������.No%Type;
  v_��� Varchar2(1000);

  v_���Ϻ�  ҩƷ�շ���¼.���ܷ�ҩ��%Type;
  v_�ⷿid  ҩƷ�շ���¼.�ⷿid%Type;
  v_�շ�ids Varchar2(4000);

  v_ҽ����Ч ����ҽ����¼.ҽ����Ч%Type;
Begin
  Open c_Noone;
  Fetch c_Noone Bulk Collect
    Into r_No_Stuff;
  Close c_Noone;

  If Nvl(����ִ��_In, 0) = 0 Then
    Open c_No;
    Fetch c_No Bulk Collect
      Into r_No;
    Close c_No;
  Else
    r_No := r_No_Stuff;
  End If;

  --�����ÿ�����Ҫ����ҽ�����
  --������ҩƷ�͸������õ����ģ���Ϊ��Щ��Ҫ���Ųű�ʾִ��
  If Nvl(����ִ��_In, 0) = 0 Then
    Open c_Finish(r_No);
    Fetch c_Finish Bulk Collect
      Into r_Finish;
    Close c_Finish;
  Else
    Open c_Finishone(r_No);
    Fetch c_Finishone Bulk Collect
      Into r_Finish;
    Close c_Finishone;
  End If;
  Forall I In 1 .. r_Finish.Count
    Update סԺ���ü�¼ Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = ����Ա����_In Where ID = r_Finish(I);

  --ִ��ʱ�Զ���˶�Ӧ�ļ��ʻ��۵�����
  --����ҽ����Ӧ��ҩƷ�����ķ��ã���Ϊҽ����ִ�У�����Ӧ����Ч��
  If Nvl(����ִ��_In, 0) = 0 Then
    For r_Verify In c_Verify(r_No) Loop
      If r_Verify.No <> v_No And v_��� Is Not Null Then
        Zl_סԺ���ʼ�¼_Verify(v_No, ����Ա���_In, ����Ա����_In, Substr(v_���, 2));
        v_��� := Null;
      End If;
      v_No   := r_Verify.No;
      v_��� := v_��� || ',' || r_Verify.���;
    End Loop;
  Else
    For r_Verify In c_Verifyone(r_No) Loop
      If r_Verify.No <> v_No And v_��� Is Not Null Then
        Zl_סԺ���ʼ�¼_Verify(v_No, ����Ա���_In, ����Ա����_In, Substr(v_���, 2));
        v_��� := Null;
      End If;
      v_No   := r_Verify.No;
      v_��� := v_��� || ',' || r_Verify.���;
    End Loop;
  End If;
  If v_��� Is Not Null Then
    Zl_סԺ���ʼ�¼_Verify(v_No, ����Ա���_In, ����Ա����_In, Substr(v_���, 2));
  End If;

  --����������������Զ�����
  For r_Stuff In c_Stuff(r_No_Stuff) Loop
    If v_���Ϻ� Is Null Then
      v_���Ϻ� := Nextno(20);
    End If;
  
    If r_Stuff.�ⷿid <> Nvl(v_�ⷿid, 0) Then
      If Nvl(v_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
        v_�շ�ids := Substr(v_�շ�ids, 2);
        Zl_ҩƷ�շ���¼_��������(v_�շ�ids, v_�ⷿid, ����Ա����_In, Sysdate, 1, ����Ա����_In, v_���Ϻ�, ����Ա����_In);
      End If;
      v_�ⷿid  := r_Stuff.�ⷿid;
      v_�շ�ids := Null;
    End If;
  
    v_�շ�ids := v_�շ�ids || '|' || r_Stuff.Id || ',0';
  End Loop;
  If Nvl(v_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
    v_�շ�ids := Substr(v_�շ�ids, 2);
    Zl_ҩƷ�շ���¼_��������(v_�շ�ids, v_�ⷿid, ����Ա����_In, Sysdate, 1, ����Ա����_In, v_���Ϻ�, ����Ա����_In);
  End If;

  --����ҩƷ�Զ���ҩ(ֻ�ڻ�ʿվ������ҩƷ�Ŵ���,�����ɲ������α��ж�)
  Select ҽ����Ч Into v_ҽ����Ч From ����ҽ����¼ Where ID = ҽ��id_In;
  If Substr(zl_GetSysParameter('����ִ���Զ����', 1254), v_ҽ����Ч + 1, 1) = '1' Then
    v_���Ϻ�  := Null;
    v_�շ�ids := Null;
    For r_Drug In c_Drug(r_No_Stuff) Loop
      If v_���Ϻ� Is Null Then
        v_���Ϻ� := Nextno(20);
      End If;
    
      If r_Drug.�ⷿid <> Nvl(v_�ⷿid, 0) Then
        If Nvl(v_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
          v_�շ�ids := Substr(v_�շ�ids, 2);
          Zl_ҩƷ�շ���¼_������ҩ(v_�շ�ids, v_�ⷿid, ����Ա����_In, Sysdate, 1, ����Ա����_In, v_���Ϻ�);
        End If;
        v_�ⷿid  := r_Drug.�ⷿid;
        v_�շ�ids := Null;
      End If;
    
      v_�շ�ids := v_�շ�ids || '|' || r_Drug.Id || ',0';
    End Loop;
    If Nvl(v_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
      v_�շ�ids := Substr(v_�շ�ids, 2);
      Zl_ҩƷ�շ���¼_������ҩ(v_�շ�ids, v_�ⷿid, ����Ա����_In, Sysdate, 1, ����Ա����_In, v_���Ϻ�);
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_סԺҽ��ִ��_Finish;
/

--124359:������,2018-04-17,�����Ŀҽ������λִ�л�ȡ��ִ��
Create Or Replace Procedure Zl_סԺҽ��ִ��_Cancel
(
  ҽ��id_In     In ����ҽ��ִ��.ҽ��id%Type,
  ���ͺ�_In     In ����ҽ��ִ��.���ͺ�%Type,
  ����ִ��_In   In Number,
  ����Ա���_In In ��Ա��.���%Type,
  ����Ա����_In In ��Ա��.����%Type,
  ��id_In       In ����ҽ��ִ��.ҽ��id%Type,
  �������_In   In ����ҽ����¼.�������%Type,
  ִ�в���id_In In ������ü�¼.ִ�в���id%Type
  --������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID��
  --      ����ִ��_In=����ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ
) Is
  --ҽ����صķ��õ���
  Cursor c_No Is
    Select a.No || ':' || a.��¼����
    From ����ҽ������ A, ����ҽ����¼ B
    Where a.���ͺ� + 0 = ���ͺ�_In And a.ҽ��id = b.Id And b.������� = �������_In And (b.Id = ��id_In Or b.���id = ��id_In)
    Union
    Select NO || ':' || ��¼����
    From ����ҽ������
    Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;

  Cursor c_Noone Is
    Select NO || ':' || ��¼����
    From ����ҽ������
    Where ���ͺ� + 0 = ���ͺ�_In And ҽ��id = ҽ��id_In
    Union
    Select NO || ':' || ��¼����
    From ����ҽ������
    Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;
  r_No       t_Strlist;
  r_No_Stuff t_Strlist;
  r_Finish   t_Numlist;

  n_ִ�д��� Number;
  n_ʣ����� Number;
  n_ִ��״̬ Number;
  n_Count    Number;
  --Ҫȡ��ִ�еķ�����(������ҩƷ�͸������õ�����)
  Cursor c_Finish(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select Distinct a.Id, a.�շ����, a.�շ�ϸĿid
           From סԺ���ü�¼ A, ����ҽ����¼ B,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
                  From Table(r_No)) N
           Where a.ҽ����� = b.Id And (b.Id = ��id_In Or b.���id = ��id_In) And a.No = n.No And a.��¼���� = n.��¼���� And
                 a.��¼״̬ In (0, 1, 3) And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)) A, �������� B
    Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);

  Cursor c_Finishone(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select a.Id, a.�շ����, a.�շ�ϸĿid
           From סԺ���ü�¼ A,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
                  From Table(r_No)) N
           Where a.ҽ����� = ҽ��id_In And a.No = n.No And a.��¼���� = n.��¼���� And a.��¼״̬ In (0, 1, 3) And
                 (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)) A, �������� B
    Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);

  --ȡ��ִ���а����������õķ�������ʱ�����ݲ��������Ƿ��Զ�����
  --��������ҽ��Ŀǰ�����ڵ��������ִ�е����
  Cursor c_Stuff(r_No t_Strlist) Is
    Select /*+ RULE */
     b.Id
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ����ҽ����¼ C,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
           From Table(r_No)) N
    Where a.Id = b.����id And (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0) And b.����� Is Not Null And a.�շ���� = '4' And a.��¼״̬ = 1 And
          a.ҽ����� = c.Id And c.������� = �������_In And (c.Id = ��id_In Or c.���id = ��id_In) And a.No = n.No And a.��¼���� = n.��¼���� And
          b.���� In (25, 26) And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And
          (c.Id = ҽ��id_In And ����ִ��_In = 1 Or Nvl(����ִ��_In, 0) <> 1)
    Order By b.ҩƷid;

  --ȡ��ִ���а���ҩƷʱ������ִ�е��Զ���ҩ
  Cursor c_Drug(r_No t_Strlist) Is
    Select /*+ RULE */
     b.Id
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ����ҽ����¼ C, ������ҳ D,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
           From Table(r_No)) N
    Where a.Id = b.����id And (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0) And b.����� Is Not Null And b.�ⷿid = ִ�в���id_In And
          a.�շ���� In ('5', '6', '7') And a.��¼״̬ = 1 And a.ҽ����� = c.Id And c.������� = �������_In And c.����id = d.����id And
          c.��ҳid = d.��ҳid And (c.Id = ��id_In Or c.���id = ��id_In) And a.No = n.No And a.��¼���� = n.��¼���� And
          (c.Id = ҽ��id_In And ����ִ��_In = 1 Or Nvl(����ִ��_In, 0) <> 1)
    Order By b.ҩƷid;

  v_ҽ����Ч ����ҽ����¼.ҽ����Ч%Type;
Begin
  Open c_Noone;
  Fetch c_Noone Bulk Collect
    Into r_No_Stuff;
  Close c_Noone;

  If Nvl(����ִ��_In, 0) = 0 Then
    Open c_No;
    Fetch c_No Bulk Collect
      Into r_No;
    Close c_No;
  Else
    r_No := r_No_Stuff;
  End If;

  --�����ÿ�����Ҫ����ҽ�����
  --������ҩƷ�͸������õ����ģ���Ϊ��Щ��Ҫ���Ųű�ʾִ��
  If Nvl(����ִ��_In, 0) = 0 Then
    Open c_Finish(r_No);
    Fetch c_Finish Bulk Collect
      Into r_Finish;
    Close c_Finish;
  Else
    Open c_Finishone(r_No);
    Fetch c_Finishone Bulk Collect
      Into r_Finish;
    Close c_Finishone;
  End If;

  Select Decode(a.ִ��״̬, 1, a.��������, c.�ǼǴ���), Decode(a.ִ��״̬, 1, 0, a.�������� - c.�ǼǴ���)
  Into n_ִ�д���, n_ʣ�����
  From ����ҽ������ A,
       (Select ҽ��id_In ҽ��id, ���ͺ�_In ���ͺ�, Nvl(Sum(b.��������), 0) As �ǼǴ���
         From ����ҽ��ִ�� B
         Where b.ҽ��id = ҽ��id_In And b.���ͺ� = ���ͺ�_In And Nvl(b.ִ�н��, 1) <> 0) C
  Where a.ҽ��id = c.ҽ��id And a.���ͺ� = c.���ͺ� And a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In;

  --���ȫ��ִ����״̬Ϊ1��δִ��״̬Ϊ0������ִ��״̬Ϊ2
  Select Decode(n_ʣ�����, 0, 1, Decode(n_ִ�д���, 0, 0, 2)) Into n_ִ��״̬ From Dual;

  Forall I In 1 .. r_Finish.Count
    Update סԺ���ü�¼
    Set ִ��״̬ = n_ִ��״̬, ִ��ʱ�� = Decode(n_ִ��״̬, 0, Null, ִ��ʱ��), ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����)
    Where ID = r_Finish(I);

  --����������������Զ�����
  For r_Stuff In c_Stuff(r_No_Stuff) Loop
    Zl_�����շ���¼_��������(r_Stuff.Id, ����Ա����_In, Sysdate, Null, Null, Null, Null, 0, ����Ա����_In);
  End Loop;

  --����ҩƷ�Զ���ҩ(ֻ�ڻ�ʿվ������ҩƷ�Ŵ���,�����ɲ������α��ж�)
  Select Max(a.ҽ����Ч), Max(Decode(b.����id, ִ�в���id_In, 1, 0))
  Into v_ҽ����Ч, n_Count
  From ����ҽ����¼ A, ���˱䶯��¼ B
  Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.Id = ҽ��id_In;

  If Substr(zl_GetSysParameter('����ִ���Զ����', 1254), v_ҽ����Ч + 1, 1) = '1' And n_Count = 1 Then
    For r_Drug In c_Drug(r_No_Stuff) Loop
      Zl_ҩƷ�շ���¼_������ҩ(r_Drug.Id, ����Ա����_In, Sysdate);
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_סԺҽ��ִ��_Cancel;
/

--124359:������,2018-04-17,�����Ŀҽ������λִ�л�ȡ��ִ��
Create Or Replace Procedure Zl_����ҽ��ִ��_Finish
(
  ҽ��id_In     In ����ҽ��ִ��.ҽ��id%Type,
  ���ͺ�_In     In ����ҽ��ִ��.���ͺ�%Type,
  ����ִ��_In   In Number,
  ����Ա���_In In ��Ա��.���%Type,
  ����Ա����_In In ��Ա��.����%Type,
  ��id_In       In ����ҽ��ִ��.ҽ��id%Type,
  �������_In   In ����ҽ����¼.�������%Type,
  ִ�в���id_In In ������ü�¼.ִ�в���id%Type
  --������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID��
  --      ����ִ��_In=����ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ
  --ִ�в���id_In=������ָ��ִ�в��ŵķ��ã���������0ʱ������ִ�в���
) Is
  --ҽ����صķ��õ���
  Cursor c_No Is
    Select a.No || ':' || a.��¼����
    From ����ҽ������ A, ����ҽ����¼ B
    Where a.���ͺ� + 0 = ���ͺ�_In And a.ҽ��id = b.Id And b.������� = �������_In And (b.Id = ��id_In Or b.���id = ��id_In)
    Union
    Select NO || ':' || ��¼����
    From ����ҽ������
    Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;

  Cursor c_Noone Is
    Select NO || ':' || ��¼����
    From ����ҽ������
    Where ���ͺ� + 0 = ���ͺ�_In And ҽ��id = ҽ��id_In
    Union
    Select NO || ':' || ��¼����
    From ����ҽ������
    Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;

  r_No       t_Strlist;
  r_No_Stuff t_Strlist;
  r_Finish   t_Numlist;
  n_Cnt      Number;
  v_Error    Varchar2(2000);
  Err_Custom Exception;
  v_ִ��ǰ�Ƚ��� Varchar2(500);

  Cursor c_Finish(r_No t_Strlist) Is
    Select a.Id
    From (Select Distinct a.Id, a.�շ����, a.�շ�ϸĿid
           From ������ü�¼ A, ����ҽ����¼ B,
                (Select /*+cardinality(f,10)*/
                   Substr(f.Column_Value, 1, Instr(f.Column_Value, ':') - 1) As NO,
                   To_Number(Substr(f.Column_Value, Instr(f.Column_Value, ':') + 1)) As ��¼����
                  From Table(r_No) F) N
           Where a.ҽ����� = b.Id And (b.Id = ��id_In Or b.���id = ��id_In) And a.No = n.No And Mod(a.��¼����, 10) = n.��¼���� And
                 a.��¼״̬ In (0, 1, 3) And a.ִ��״̬ <> 1 And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)) A, �������� B
    Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);

  Cursor c_Finishone(r_No t_Strlist) Is
    Select a.Id
    From (Select a.Id, a.�շ����, a.�շ�ϸĿid
           From ������ü�¼ A,
                (Select /*+cardinality(f,10)*/
                   Substr(f.Column_Value, 1, Instr(f.Column_Value, ':') - 1) As NO,
                   To_Number(Substr(f.Column_Value, Instr(f.Column_Value, ':') + 1)) As ��¼����
                  From Table(r_No) F) N
           Where a.ҽ����� = ҽ��id_In And a.No = n.No And Mod(a.��¼����, 10) = n.��¼���� And a.��¼״̬ In (0, 1, 3) And a.ִ��״̬ <> 1 And
                 (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)) A, �������� B
    Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);

  --ִ���а����������õ�δ������ʱ�����ݲ��������Ƿ��Զ�����
  --��������ҽ��Ŀǰ�����ڵ��������ִ�е����
  Cursor c_Stuff(r_No t_Strlist) Is
    Select b.Id, Decode(d.��ֵ����, 1, a.ִ�в���id, b.�ⷿid) As �ⷿid
    From ������ü�¼ A, ҩƷ�շ���¼ B, ����ҽ����¼ C, �������� D,
         (Select /*+cardinality(f,10)*/
            Substr(f.Column_Value, 1, Instr(f.Column_Value, ':') - 1) As NO,
            To_Number(Substr(f.Column_Value, Instr(f.Column_Value, ':') + 1)) As ��¼����
           From Table(r_No) F) N
    Where d.����id = a.�շ�ϸĿid And a.Id = b.����id And Mod(b.��¼״̬, 3) = 1 And b.����� Is Null And b.�ⷿid Is Not Null And
          a.�շ���� = '4' And a.��¼״̬ = 1 And a.ҽ����� = c.Id And c.������� = �������_In And (c.Id = ��id_In Or c.���id = ��id_In) And
          a.No = n.No And Mod(a.��¼����, 10) = n.��¼���� And b.���� In (24, 25, 26) And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And
          (c.Id = ҽ��id_In And ����ִ��_In = 1 Or Nvl(����ִ��_In, 0) <> 1)
    Order By b.�ⷿid, b.ҩƷid;

  --δ��˵ķ�����(����ҩƷ������)
  Cursor c_Verify
  (
    r_No        t_Strlist,
    ���ʷ���_In Number := 1
  ) Is
    Select Distinct a.No, a.���
    From ������ü�¼ A, ����ҽ����¼ C,
         (Select /*+cardinality(f,10)*/
            Substr(f.Column_Value, 1, Instr(f.Column_Value, ':') - 1) As NO,
            To_Number(Substr(f.Column_Value, Instr(f.Column_Value, ':') + 1)) As ��¼����
           From Table(r_No) F) N
    Where Nvl(a.���ʷ���, 0) = ���ʷ���_In And a.��¼״̬ = 0 And a.�۸񸸺� Is Null And a.ҽ����� = c.Id And
          (c.Id = ��id_In Or c.���id = ��id_In) And a.No = n.No And Mod(a.��¼����, 10) = n.��¼���� And
          (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)
    Order By NO, ���;

  Cursor c_Verifyone
  (
    r_No        t_Strlist,
    ���ʷ���_In Number := 1
  ) Is
    Select a.No, a.���
    From ������ü�¼ A,
         (Select /*+cardinality(f,10)*/
            Substr(f.Column_Value, 1, Instr(f.Column_Value, ':') - 1) As NO,
            To_Number(Substr(f.Column_Value, Instr(f.Column_Value, ':') + 1)) As ��¼����
           From Table(r_No) F) N
    Where Nvl(a.���ʷ���, 0) = ���ʷ���_In And a.��¼״̬ = 0 And a.�۸񸸺� Is Null And a.ҽ����� + 0 = ҽ��id_In And a.No = n.No And
          Mod(a.��¼����, 10) = n.��¼���� And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)
    Order By NO, ���;

  v_No   ����ҽ������.No%Type;
  v_��� Varchar2(1000);

  v_���Ϻ�  ҩƷ�շ���¼.���ܷ�ҩ��%Type;
  v_�ⷿid  ҩƷ�շ���¼.�ⷿid%Type;
  v_�շ�ids Varchar2(4000);
Begin
  Open c_Noone;
  Fetch c_Noone Bulk Collect
    Into r_No_Stuff;
  Close c_Noone;

  If Nvl(����ִ��_In, 0) = 0 Then
    Open c_No;
    Fetch c_No Bulk Collect
      Into r_No;
    Close c_No;
  Else
    r_No := r_No_Stuff;
  End If;

  --�����ÿ�����Ҫ����ҽ�����
  --������ҩƷ�͸������õ����ģ���Ϊ��Щ��Ҫ���Ųű�ʾִ��
  If Nvl(����ִ��_In, 0) = 0 Then
    Open c_Finish(r_No);
    Fetch c_Finish Bulk Collect
      Into r_Finish;
    Close c_Finish;
  Else
    Open c_Finishone(r_No);
    Fetch c_Finishone Bulk Collect
      Into r_Finish;
    Close c_Finishone;
  End If;
  Select Count(1)
  Into n_Cnt
  From ������ü�¼ A,
       (Select /*+cardinality(f,10)*/
          To_Number(f.Column_Value) As ����id
         From Table(r_Finish) F) B
  Where a.Id = b.����id And a.����״̬ = 1;

  If n_Cnt > 0 Then
    v_Error := '��ǰִ�е�ҽ����Ӧ�ķ��õ����д����쳣���ݡ�';
    Raise Err_Custom;
  End If;

  Select zl_GetSysParameter(163) Into v_ִ��ǰ�Ƚ��� From Dual;
  Forall I In 1 .. r_Finish.Count
    Update ������ü�¼ Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = ����Ա����_In Where ID = r_Finish(I);

  --ִ��ʱ�Զ���˶�Ӧ�ļ��ʻ��۵�����
  --����ҽ����Ӧ��ҩƷ�����ķ��ã���Ϊҽ����ִ�У�����Ӧ����Ч��
  If Nvl(����ִ��_In, 0) = 0 Then
    If Nvl(v_ִ��ǰ�Ƚ���, '0') <> '0' Then
      For r_Verify In c_Verify(r_No, 0) Loop
        v_Error := '��ǰִ�е�ҽ��������δ��ȡ�ķ��á�';
        Raise Err_Custom;
      End Loop;
    End If;
    For r_Verify In c_Verify(r_No) Loop
      If r_Verify.No <> v_No And v_��� Is Not Null Then
        Zl_������ʼ�¼_Verify(v_No, ����Ա���_In, ����Ա����_In, Substr(v_���, 2));
        v_��� := Null;
      End If;
      v_No   := r_Verify.No;
      v_��� := v_��� || ',' || r_Verify.���;
    End Loop;
  Else
    If Nvl(v_ִ��ǰ�Ƚ���, '0') <> '0' Then
      For r_Verify In c_Verifyone(r_No, 0) Loop
        v_Error := '��ǰִ�е�ҽ��������δ��ȡ�ķ��á�';
        Raise Err_Custom;
      End Loop;
    End If;
    For r_Verify In c_Verifyone(r_No) Loop
      If r_Verify.No <> v_No And v_��� Is Not Null Then
        Zl_������ʼ�¼_Verify(v_No, ����Ա���_In, ����Ա����_In, Substr(v_���, 2));
        v_��� := Null;
      End If;
      v_No   := r_Verify.No;
      v_��� := v_��� || ',' || r_Verify.���;
    End Loop;
  End If;
  If v_��� Is Not Null Then
    Zl_������ʼ�¼_Verify(v_No, ����Ա���_In, ����Ա����_In, Substr(v_���, 2));
  End If;

  --����������������Զ�����
  For r_Stuff In c_Stuff(r_No_Stuff) Loop
    If v_���Ϻ� Is Null Then
      v_���Ϻ� := Nextno(20);
    End If;
  
    If r_Stuff.�ⷿid <> Nvl(v_�ⷿid, 0) Then
      If Nvl(v_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
        v_�շ�ids := Substr(v_�շ�ids, 2);
        Zl_ҩƷ�շ���¼_��������(v_�շ�ids, v_�ⷿid, ����Ա����_In, Sysdate, 1, ����Ա����_In, v_���Ϻ�, ����Ա����_In);
      End If;
      v_�ⷿid  := r_Stuff.�ⷿid;
      v_�շ�ids := Null;
    End If;
  
    v_�շ�ids := v_�շ�ids || '|' || r_Stuff.Id || ',0';
  End Loop;
  If Nvl(v_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
    v_�շ�ids := Substr(v_�շ�ids, 2);
    Zl_ҩƷ�շ���¼_��������(v_�շ�ids, v_�ⷿid, ����Ա����_In, Sysdate, 1, ����Ա����_In, v_���Ϻ�, ����Ա����_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ��ִ��_Finish;
/

--124359:������,2018-04-17,�����Ŀҽ������λִ�л�ȡ��ִ��
Create Or Replace Procedure Zl_����ҽ��ִ��_Cancel
(
  ҽ��id_In     In ����ҽ��ִ��.ҽ��id%Type,
  ���ͺ�_In     In ����ҽ��ִ��.���ͺ�%Type,
  ����ִ��_In   In Number,
  ����Ա���_In In ��Ա��.���%Type,
  ����Ա����_In In ��Ա��.����%Type,
  ��id_In       In ����ҽ��ִ��.ҽ��id%Type,
  �������_In   In ����ҽ����¼.�������%Type,
  ִ�в���id_In In ������ü�¼.ִ�в���id%Type
  --������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID��
  --      ����ִ��_In=����ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ
) Is
  --ҽ����صķ��õ���
  Cursor c_No Is
    Select a.No || ':' || a.��¼����
    From ����ҽ������ A, ����ҽ����¼ B
    Where a.���ͺ� + 0 = ���ͺ�_In And a.ҽ��id = b.Id And b.������� = �������_In And (b.Id = ��id_In Or b.���id = ��id_In)
    Union
    Select NO || ':' || ��¼����
    From ����ҽ������
    Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;

  Cursor c_Noone Is
    Select NO || ':' || ��¼����
    From ����ҽ������
    Where ���ͺ� + 0 = ���ͺ�_In And ҽ��id = ҽ��id_In
    Union
    Select NO || ':' || ��¼����
    From ����ҽ������
    Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;

  r_No       t_Strlist;
  r_No_Stuff t_Strlist;
  r_Finish   t_Numlist;

  n_ִ�д��� Number;
  n_ʣ����� Number;
  n_ִ��״̬ Number;

  --Ҫȡ��ִ�еķ�����(������ҩƷ�͸������õ�����)
  Cursor c_Finish(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select Distinct a.Id, a.�շ����, a.�շ�ϸĿid
           From ������ü�¼ A, ����ҽ����¼ B,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
                  From Table(r_No)) N
           Where a.ҽ����� = b.Id And (b.Id = ��id_In Or b.���id = ��id_In) And a.No = n.No And
                 (a.��¼���� = n.��¼���� Or a.��¼���� = 11 And n.��¼���� = 1) And a.��¼״̬ In (0, 1, 3) And
                 (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)) A, �������� B
    Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);

  Cursor c_Finishone(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select a.Id, a.�շ����, a.�շ�ϸĿid
           From ������ü�¼ A,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
                  From Table(r_No)) N
           Where a.ҽ����� = ҽ��id_In And a.No = n.No And (a.��¼���� = n.��¼���� Or a.��¼���� = 11 And n.��¼���� = 1) And
                 a.��¼״̬ In (0, 1, 3) And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In)) A, �������� B
    Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);

  --ȡ��ִ���а����������õķ�������ʱ�����ݲ��������Ƿ��Զ�����
  --��������ҽ��Ŀǰ�����ڵ��������ִ�е����
  Cursor c_Stuff(r_No t_Strlist) Is
    Select /*+ RULE */
     b.Id
    From ������ü�¼ A, ҩƷ�շ���¼ B, ����ҽ����¼ C,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As ��¼����
           From Table(r_No)) N
    Where a.Id = b.����id And (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0) And b.����� Is Not Null And a.�շ���� = '4' And a.��¼״̬ = 1 And
          a.ҽ����� = c.Id And c.������� = �������_In And (c.Id = ��id_In Or c.���id = ��id_In) And a.No = n.No And
          (a.��¼���� = n.��¼���� Or a.��¼���� = 11 And n.��¼���� = 1) And b.���� In (24, 25, 26) And
          (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And (c.Id = ҽ��id_In And ����ִ��_In = 1 Or Nvl(����ִ��_In, 0) <> 1)
    Order By b.ҩƷid;
Begin
  Open c_Noone;
  Fetch c_Noone Bulk Collect
    Into r_No_Stuff;
  Close c_Noone;

  If Nvl(����ִ��_In, 0) = 0 Then
    Open c_No;
    Fetch c_No Bulk Collect
      Into r_No;
    Close c_No;
  Else
    r_No := r_No_Stuff;
  End If;

  --�����ÿ�����Ҫ����ҽ�����
  --������ҩƷ�͸������õ����ģ���Ϊ��Щ��Ҫ���Ųű�ʾִ��
  If Nvl(����ִ��_In, 0) = 0 Then
    Open c_Finish(r_No);
    Fetch c_Finish Bulk Collect
      Into r_Finish;
    Close c_Finish;
  Else
    Open c_Finishone(r_No);
    Fetch c_Finishone Bulk Collect
      Into r_Finish;
    Close c_Finishone;
  End If;

  Select Decode(a.ִ��״̬, 1, a.��������, c.�ǼǴ���), Decode(a.ִ��״̬, 1, 0, a.�������� - c.�ǼǴ���)
  Into n_ִ�д���, n_ʣ�����
  From ����ҽ������ A,
       (Select ҽ��id_In ҽ��id, ���ͺ�_In ���ͺ�, Nvl(Sum(b.��������), 0) As �ǼǴ���
         From ����ҽ��ִ�� B
         Where b.ҽ��id = ҽ��id_In And b.���ͺ� = ���ͺ�_In And Nvl(b.ִ�н��, 1) <> 0) C
  Where a.ҽ��id = c.ҽ��id And a.���ͺ� = c.���ͺ� And a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In;

  --���ȫ��ִ����״̬Ϊ1��δִ��״̬Ϊ0������ִ��״̬Ϊ2
  Select Decode(n_ʣ�����, 0, 1, Decode(n_ִ�д���, 0, 0, 2)) Into n_ִ��״̬ From Dual;

  --�������ﵥ�ݣ������������շѣ�����ִ�У�2������ȫִ�У�1��,ִ��ʱ��Ϊִ����ɵ�ִ��ʱ�䣬ִ����Ϊִ����ɵ�ִ����
  Forall I In 1 .. r_Finish.Count
    Update ������ü�¼
    Set ִ��״̬ = n_ִ��״̬, ִ��ʱ�� = Decode(n_ִ��״̬, 0, Null, ִ��ʱ��), ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����)
    Where ID = r_Finish(I);

  --����������������Զ�����
  For r_Stuff In c_Stuff(r_No_Stuff) Loop
    Zl_�����շ���¼_��������(r_Stuff.Id, ����Ա����_In, Sysdate, Null, Null, Null, Null, 0, ����Ա����_In);
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ��ִ��_Cancel;
/


--124306:����,2018-04-17,����Oracle����Zl_Third_Charge_Del��Zl_Third_Registdel
Create Or Replace Procedure Zl_Third_Charge_Del
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  -------------------------------------------------------------------------------------------------- 
  --����:�����˷ѽ��� 
  --���:Xml_In: 
  --<IN> 
  --    <BRID>����ID</BRID> 
  --    <XM>����</XM> 
  --    <SFZH>���֤��</SFZH> 
  --    <JE></JE> //�˿��ܽ�� 
  --    <JSKLB></JSKLB>     //���㿨��� 
  --    <TFZY>�˷�ժҪ</TFZY> 
  --    <JCFP>1</JCFP>      //��鷢Ʊ 
  --    <FYLIST> 
  --        <FY> 
  --           <DJH>�˿�ݺ�</DJH> 
  --           <XH>�˿����(��ʽ:1,2,3..Ϊ�մ�����ʣ������)</DJH> 
  --        <FY> 
  --    </FYLIST> 
  --    <TKLIST> 
  --        <TK> 
  --            <TKKLB>�˿���</TKKLB> 
  --            <TKKH>�˿��</TKKH> 
  --            <TKFS>�˿ʽ</TKFS> //�˿ʽ:�ֽ�;֧Ʊ,�����������,���Դ��� 
  --            <TKJE>֧�����</TKJE> 
  --            <JYLSH>������ˮ��</JYLSH> 
  --            <TKZY>ժҪ</TKZY> 
  --            <TYJK>�˻�Ԥ����</TYJK> //�ʳ�Ԥ��ʱ,ֻ��JSJE�ڵ�:1-��Ԥ�� 
  --            <SFXFK>�Ƿ����ѿ�</SFXFK>   //(1-�����ѿ�),���ѿ�ʱ,������㿨���,���㿨��,������Ƚӵ� 
  --            <EXPENDLIST>  //��չ������Ϣ 
  --                <EXPEND> 
  --                    <JYMC>��������</JYMC> 
  --                    <JYLR>��������</JYLR> 
  --                </EXPEND> 
  --            </EXPENDLIST> 
  --        </TK> 
  --    </TKLIST> 
  --</IN> 

  --����:Xml_Out 
  --  <OUT> 
  --    <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ�� 
  --    <YJZID>ԭ����ID</YJZID>       //ԭ����ID 
  --    <CXID>����ID</CXID>          //����ID 
  --    �D�D�������д�������˵����ȷִ�� 
  --    <ERROR> 
  --      <MSG>������Ϣ</MSG> 
  --    </ERROR> 
  --  </OUT> 
  -------------------------------------------------------------------------------------------------- 
  n_�˿��ܶ� ������ü�¼.ʵ�ս��%Type;
  n_�����id ҽ�ƿ����.Id%Type;
  v_���㷽ʽ Varchar2(2000);

  n_����id     ������ü�¼.����id%Type;
  v_����       ������Ϣ.����%Type;
  v_���֤��   ������Ϣ.���֤��%Type;
  n_���ݲ���id ������ü�¼.����id%Type;
  v_����Ա���� ������ü�¼.����Ա���%Type;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  n_����id     ������ü�¼.����id%Type;
  n_����id     ������ü�¼.����id%Type;
  n_���ʽ��   ������ü�¼.���ʽ��%Type;
  n_����     ����Ԥ����¼.��Ԥ��%Type;
  n_ԭ������� ����Ԥ����¼.�������%Type;
  l_�Һŵ�     t_Strlist := t_Strlist();
  n_Column     Number(18);
  n_��鷢Ʊ   Number(3);
  n_�Ƿ��ӡ   Number(3);
  v_���㿨��� Varchar2(100);
  v_����ids    Varchar2(1000);

  n_���ѿ�id ���ѿ���Ϣ.Id%Type;
  v_ժҪ     ������ü�¼.ժҪ%Type;
  n_Count    Number(18);

  d_�˷�ʱ�� ����Ԥ����¼.�տ�ʱ��%Type;

  v_�˷ѽ��� Varchar2(2000);
  v_��ͨ���� Varchar2(4000);
  n_Temp     Number(18);

  v_Temp    Varchar2(32767); --��ʱXML 
  x_Templet Xmltype; --ģ��XML 

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  Procedure Third_Cardbalance_Modfiy
  (
    ����id_In     ����Ԥ����¼.����id%Type,
    �����_In     Varchar2,
    ����_In       ����Ԥ����¼.����%Type,
    �˿���_In   ����Ԥ����¼.��Ԥ��%Type,
    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    ����˵��_In   ����Ԥ����¼.����˵��%Type,
    ժҪ_In       ����Ԥ����¼.ժҪ%Type,
    Xmlexpned_In  Xmltype
  ) Is
    n_�����id ҽ�ƿ����.Id%Type;
    v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  
  Begin
    v_Err_Msg := Null;
    Begin
      n_�����id := To_Number(�����_In);
    Exception
      When Others Then
        n_�����id := 0;
    End;
    If n_�����id = 0 Then
      Begin
        Select ID, ���㷽ʽ, Decode(Nvl(�Ƿ�����, 0), 1, Null, ���� || 'δ����,��������нɷ�!')
        Into n_�����id, v_���㷽ʽ, v_Err_Msg
        From ҽ�ƿ����
        Where ���� = �����_In;
      Exception
        When Others Then
          n_�����id := -1;
          v_Err_Msg  := �����_In || '������!';
      End;
    Else
      Begin
        Select ID, ���㷽ʽ, Decode(Nvl(�Ƿ�����, 0), 1, Null, ���� || 'δ����,��������нɷ�!')
        Into n_�����id, v_���㷽ʽ, v_Err_Msg
        From ҽ�ƿ����
        Where ID = n_�����id;
      Exception
        When Others Then
          n_�����id := -1;
          v_Err_Msg  := 'δ�ҵ�ָ���Ľ���֧����Ϣ!';
      End;
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
  
    v_�˷ѽ��� := v_���㷽ʽ || '|' || �˿���_In || '|' || ' |' || Nvl(ժҪ_In, ' ');
    --   2.�������˷ѽ���: 
    --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:���㷽ʽ|������|�������|����ժҪ 
    --     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ���� 
    --���㷽ʽ|������|�������|����ժҪ 
    Zl_�����˷ѽ���_Modify(2, n_����id, ����id_In, v_�˷ѽ���, 0, n_�����id, ����_In, ������ˮ��_In, ����˵��_In, 0, 0, 0, 0);
  
    --������չ������Ϣ 
    For c_��չ In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_�������㽻��_Insert(n_�����id, 0, ����_In, ����id_In, c_��չ.Jymc || '|' || c_��չ.Jylr, 0);
    End Loop;
  End Third_Cardbalance_Modfiy;

  Procedure Square_Cardbalance_Modfiy
  (
    ����id_In     ����Ԥ����¼.����id%Type,
    �����_In     Varchar2,
    ����_In       ����Ԥ����¼.����%Type,
    �˿���_In   ����Ԥ����¼.��Ԥ��%Type,
    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    ����˵��_In   ����Ԥ����¼.����˵��%Type,
    ժҪ_In       ����Ԥ����¼.ժҪ%Type,
    
    Xmlexpned_In Xmltype
  ) Is
    n_�����id ҽ�ƿ����.Id%Type;
    v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  
  Begin
    v_Err_Msg := Null;
    Begin
      n_�����id := To_Number(�����_In);
    Exception
      When Others Then
        n_�����id := 0;
    End;
  
    If n_�����id = 0 Then
      Begin
        Select ���, ���㷽ʽ, Decode(Nvl(����, 0), 1, Null, ���� || 'δ����,��������нɷ�!')
        Into n_�����id, v_���㷽ʽ, v_Err_Msg
        From ���ѿ����Ŀ¼
        Where ���� = �����_In;
      Exception
        When Others Then
          v_Err_Msg := '����:' || �����_In || '������!';
      End;
    
    Else
    
      Begin
        Select ���, ���㷽ʽ, Decode(Nvl(����, 0), 1, Null, ���� || 'δ����,��������нɷ�!')
        Into n_�����id, v_���㷽ʽ, v_Err_Msg
        From ���ѿ����Ŀ¼
        Where ��� = �����_In;
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ľ���֧����Ϣ!';
      End;
    
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
  
    v_�˷ѽ��� := v_���㷽ʽ || '|' || �˿���_In || '|' || ' |' || Nvl(ժҪ_In, ' ');
    --   4-���ѿ�����: 
    --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��|| 
    --     ����֧Ʊ��_In:������ 
    Select ID
    Into n_���ѿ�id
    From ���ѿ���Ϣ
    Where �ӿڱ�� = n_�����id And ���� = ����_In And
          ��� = (Select Max(���) From ���ѿ���Ϣ Where �ӿڱ�� = n_�����id And ���� = ����_In);
  
    --�����ID|����|���ѿ�ID|���ѽ��||. 
    v_�˷ѽ��� := n_�����id || '|' || ����_In || '|' || n_���ѿ�id || '|' || �˿���_In;
    Zl_�����˷ѽ���_Modify(4, n_����id, ����id_In, v_�˷ѽ���, 0, Null, ����_In, ������ˮ��_In, ����˵��_In, 0, 0, 0, 0);
  
    --������չ������Ϣ 
    For c_��չ In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_�������㽻��_Insert(n_�����id, 1, ����_In, ����id_In, c_��չ.Jymc || '|' || c_��չ.Jylr, 0);
    End Loop;
  End Square_Cardbalance_Modfiy;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  --0.��ȡ����еĲ���ID����Ϣ 
  Select To_Number(Extractvalue(Value(A), 'IN/BRID')), To_Number(Extractvalue(Value(A), 'IN/JE')),
         Extractvalue(Value(A), 'IN/TFZY'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM')
  Into n_����id, n_�˿��ܶ�, v_ժҪ, n_��鷢Ʊ, v_���㿨���, v_���֤��, v_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_����id, 0) = 0 And Not v_���֤�� Is Null And Not v_���� Is Null Then
    n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
  End If;

  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '�޷�ȷ��������Ϣ,����!';
    Raise Err_Item;
  End If;

  --��Աid,��Ա���,��Ա���� 
  v_Temp       := Zl_Identity(1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := v_Temp;
  v_Err_Msg    := Null;
  v_����ids    := Null;

  If v_���㿨��� Is Not Null Then
    Begin
      n_�����id := To_Number(v_���㿨���);
    Exception
      When Others Then
        n_�����id := 0;
    End;
    If n_�����id = 0 Then
      Begin
        Select ID Into n_�����id From ҽ�ƿ���� Where ���� = v_���㿨���;
      Exception
        When Others Then
          v_Err_Msg := '�޷�ȷ�ϴ���Ľ��㿨��';
          Raise Err_Item;
      End;
    End If;
  Else
    n_�����id := 0;
  End If;

  If Nvl(n_�����id, 0) <> 0 Then
    Select ���㷽ʽ Into v_���㷽ʽ From ҽ�ƿ���� Where ID = n_�����id;
  End If;

  --1.�Ƚ����˷� 

  Select ���˽��ʼ�¼_Id.Nextval, Sysdate Into n_����id, d_�˷�ʱ�� From Dual;

  n_Count      := 0;
  n_ԭ������� := 0;
  For c_���� In (Select Extractvalue(b.Column_Value, '/FY/DJH') As ���ݺ�, Extractvalue(b.Column_Value, '/FY/XH') As �˿����
               From Table(Xmlsequence(Extract(Xml_In, '/IN/FYLIST/FY'))) B) Loop
    Begin
      Select �������, ����id, ����id
      Into n_Temp, n_����id, n_���ݲ���id
      From ����Ԥ����¼
      Where ����id In (Select ����id
                     From ������ü�¼
                     Where NO = c_����.���ݺ� And ��¼���� = 1 And Nvl(����״̬, 0) = 0 And ��¼״̬ In (1, 3) And Rownum < 2) And
            Rownum < 2;
    Exception
      When Others Then
        n_Temp := Null;
    End;
    If Instr(',' || v_����ids || ',', ',' || n_����id || ',') = 0 Then
      v_����ids := v_����ids || ',' || n_����id;
    End If;
  
    If n_Temp Is Null Then
      v_Err_Msg := 'ָ���ĵ��ݺ�:' || c_����.���ݺ� || 'δ�ҵ�,�����˷�!';
      Raise Err_Item;
    End If;
  
    For c_�Һ� In (Select b.No As �Һŵ�, b.�շѵ�
                 From ������ü�¼ A, ���˹Һż�¼ B
                 Where a.No = c_����.���ݺ� And a.��¼���� = 1 And Nvl(����״̬, 0) = 0 And a.��¼״̬ In (1, 3) And a.�Һ�id = b.Id And
                       Instr(',' || b.�շѵ� || ',', ',' || c_����.���ݺ� || ',') > 0 And Rownum < 2) Loop
      Select /*+ cardinality(b, 10) */
       Count(1)
      Into n_Column
      From ������ü�¼ A
      Where a.��¼���� = 1 And a.No In (Select Column_Value From Table(f_Str2list(c_�Һ�.�շѵ�))) And a.��¼״̬ = 1 And
            a.No <> c_����.���ݺ� And ��� = 1;
      If n_Column = 0 Then
        If Not c_�Һ�.�Һŵ� Is Null Then
          l_�Һŵ�.Extend;
          l_�Һŵ�(l_�Һŵ�.Count) := c_�Һ�.�Һŵ�;
        End If;
      End If;
    End Loop;
  
    If Nvl(n_���ݲ���id, 0) = 0 Then
      Begin
        Select ����id
        Into n_���ݲ���id
        From ������ü�¼
        Where NO = c_����.���ݺ� And ��¼���� = 1 And Nvl(����״̬, 0) = 0 And ��¼״̬ In (1, 3) And Rownum < 2;
      Exception
        When Others Then
          n_���ݲ���id := 0;
      End;
    End If;
  
    If Nvl(n_����id, 0) <> Nvl(n_���ݲ���id, 0) Then
      v_Err_Msg := '�����˷ѵ��շѵ�:' || c_����.���ݺ� || '���ǵ�ǰ���˵��շѵ�,�����˷�!';
      Raise Err_Item;
    End If;
  
    If n_ԭ������� <> 0 And n_ԭ������� <> n_Temp Then
      v_Err_Msg := '�����˷ѵĵ��ݺŲ���һ���շѽ���,�����˷�!';
      Raise Err_Item;
    End If;
    n_ԭ������� := n_Temp;
  
    Select Count(*) Into n_Temp From ���ò����¼ Where �շѽ���id = n_����id;
    If Nvl(n_Temp, 0) <> 0 Then
      v_Err_Msg := '�����˷ѵĵ��ݺ��Ѿ������˱��ղ������,�����˷�!';
      Raise Err_Item;
    End If;
  
    If v_���㿨��� Is Not Null Then
      Select Count(*) Into n_Temp From ����Ԥ����¼ Where ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
      If Nvl(n_Temp, 0) = 0 Then
        v_Err_Msg := '�����˷ѵĵ��ݲ���' || v_���㷽ʽ || '�����,�����˷�!';
        Raise Err_Item;
      End If;
    End If;
  
    If Nvl(n_��鷢Ʊ, 0) = 1 Then
      Select Max(Decode(a.ʵ��Ʊ��, Null, 0, 1))
      Into n_�Ƿ��ӡ
      From ������ü�¼ A
      Where NO = c_����.���ݺ� And ��¼���� = 1;
      If Nvl(n_�Ƿ��ӡ, 0) = 1 Then
        v_Err_Msg := '�����˷ѵĵ��ݺ��ѿ���Ʊ,�����˷�!';
        Raise Err_Item;
      End If;
    End If;
  
    Zl_�����շѼ�¼_����(c_����.���ݺ�, v_����Ա����, v_����Ա����, c_����.�˿����, d_�˷�ʱ��, v_ժҪ, n_����id);
    n_Count := n_Count + 1;
  End Loop;
  If n_Count = 0 Then
    v_Err_Msg := 'δȷ��������Ҫ�˷ѵĵ���,�����˷�!';
    Raise Err_Item;
  End If;

  --2.�����˷ѵĽ�����Ϣ 

  n_���ʽ�� := 0;

  --����ܽ���Ƿ���ȷ 
  Select Sum(���ʽ��) Into n_���ʽ�� From ������ü�¼ Where ����id = n_����id;

  n_���� := -1 * Nvl(n_���ʽ��, 0) - Nvl(n_�˿��ܶ�, 0);
  If Abs(n_����) > 1.00 Then
    v_Err_Msg := '���ݽɿ�����ʵ�ʽ�������̫��!';
    Raise Err_Item;
  End If;

  --2.ȷ��֧����ʽ 
  n_Count := 0;
  For c_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/TK/TKKLB') As �����, Extractvalue(b.Column_Value, '/TK/TKKH') As ����,
                        Extractvalue(b.Column_Value, '/TK/TKFS') As ���㷽ʽ,
                        -1 * To_Number(Extractvalue(b.Column_Value, '/TK/TKJE')) As �˿���,
                        Extractvalue(b.Column_Value, '/TK/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/TK/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/TK/TKZY') As ժҪ,
                        Extractvalue(b.Column_Value, '/TK/TYJK') As �Ƿ���Ԥ��,
                        Extractvalue(b.Column_Value, '/TK/SFXFK') As �Ƿ����ѿ�,
                        Extract(b.Column_Value, '/TK/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/TKLIST/TK'))) B) Loop
  
    --1.�˻������� 
    If c_���㷽ʽ.����� Is Not Null And Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0) = 0 Then
      --1.���������� 
      Third_Cardbalance_Modfiy(n_����id, c_���㷽ʽ.�����, c_���㷽ʽ.����, c_���㷽ʽ.�˿���, c_���㷽ʽ.������ˮ��, c_���㷽ʽ.����˵��, c_���㷽ʽ.ժҪ,
                               c_���㷽ʽ.Expend);
    Elsif c_���㷽ʽ.����� Is Not Null And Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0) = 1 Then
      --2.���ѿ����� 
      Square_Cardbalance_Modfiy(n_����id, c_���㷽ʽ.�����, c_���㷽ʽ.����, c_���㷽ʽ.�˿���, c_���㷽ʽ.������ˮ��, c_���㷽ʽ.����˵��, c_���㷽ʽ.ժҪ,
                                c_���㷽ʽ.Expend);
    Elsif Nvl(c_���㷽ʽ.�Ƿ���Ԥ��, 0) = 1 Then
      --3.��Ԥ���� 
      Zl_�����˷ѽ���_Modify(4, n_����id, n_����id, Null, c_���㷽ʽ.�˿���, Null, Null, Null, Null, 0, 0, 0, 0);
    Else
      --4.��ͨ���� 
      If c_���㷽ʽ.���㷽ʽ Is Null Then
        v_Err_Msg := 'δָ��ָ����ʽ�����ʽɿ�!';
        Raise Err_Item;
      End If;
      --���㷽ʽ|������|�������|����ժҪ||.. 
      v_�˷ѽ��� := c_���㷽ʽ.���㷽ʽ || '|' || c_���㷽ʽ.�˿��� || '| |' || Nvl(c_���㷽ʽ.ժҪ, '  ');
      v_��ͨ���� := Nvl(v_��ͨ����, '') || '||' || v_�˷ѽ���;
    End If;
    n_Count := n_Count + 1;
  End Loop;

  --   0-ԭ���� 
  --      ԭ������һ��ȫ��,����У�Ա�־��Ϊ1,ҽ�����óɹ���,����Ϊ2,��ɺ���0 
  --   1-��ͨ�˷ѷ�ʽ: 
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������. 
  --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���) 
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.. 
  --     ����֧Ʊ��_In:������ 
  If n_Count = 0 Then
    v_Err_Msg := '������Чȷ�ϵ�ǰ��֧����ʽ!';
    Raise Err_Item;
  End If;

  --5.��ͨ���㼰��ɽ� 
  If v_��ͨ���� Is Not Null Then
    v_��ͨ���� := Substr(v_��ͨ����, 3);
  End If;
  Zl_�����˷ѽ���_Modify(1, n_����id, n_����id, v_��ͨ����, 0, Null, Null, Null, Null, 0, 0, n_����, 2);

  If v_����ids Is Not Null Then
    v_����ids := Substr(v_����ids, 2);
  End If;

  If l_�Һŵ�.Count <> 0 Then
    For I In 0 .. l_�Һŵ�.Count Loop
      x_Templet := Xmltype('<IN></IN>');
      v_Temp    := '<GHDH>' || l_�Һŵ�(I) || '</GHDH>';
      Select Appendchildxml(x_Templet, '/IN', Xmltype(v_Temp)) Into x_Templet From Dual;
      v_Temp := '<JSKLB>' || v_���㿨��� || '</JSKLB>';
      Select Appendchildxml(x_Templet, '/IN', Xmltype(v_Temp)) Into x_Templet From Dual;
      v_Temp := '<GHJE>' || 0 || '</GHJE>';
      Select Appendchildxml(x_Templet, '/IN', Xmltype(v_Temp)) Into x_Templet From Dual;
      Zl_Third_Registdel(x_Templet, Xml_Out);
    End Loop;
  End If;
  v_Temp := '<CZSJ>' || To_Char(d_�˷�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<YJZID>' || v_����ids || '</YJZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<CXID>' || n_����id || '</CXID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Charge_Del;
/

--124306:����,2018-04-17,����Oracle����Zl_Third_Charge_Del��Zl_Third_Registdel
CREATE OR REPLACE Procedure Zl_Third_Registdel
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --����:HIS�˺� 
  --���:Xml_In: 
  --<IN> 
  --  <GHDH>A000001</GHDH>    //�Һŵ��� 
  --  <JSKLB>֧����</JSKLB>      //���㿨��� 
  --  <JCFP>1</JCFP>            //��鷢Ʊ 
  --  <GHJE>20</GHJE>            //�ҺŽ�� 
  --  <LSH>34563</LSH>           //������ˮ�� 
  --  <JKFS>0</JKFS>             //�ɿʽ,0-�ҺŻ�ԤԼ�ɿ�;1-ԤԼ���ɿ� 
  --  <YYFS></YYFS>              //�ɿʽ=1ʱ���룬ԤԼ��ԤԼ��ʽ 
  --</IN> 

  --����:Xml_Out 
  --<OUTPUT> 
  -- <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ�� 
  -- <YJZID>ԭ����ID</YJZID> 
  -- <CXID>����ID</CXID> 
  -- <ERROR><MSG></MSG></ERROR> //Ϊ�ձ�ʾȡ���Һųɹ� 
  --</OUTPUT> 
  -------------------------------------------------------------------------------------------------- 
  v_�����     Varchar2(100);
  v_No         ���˹Һż�¼.No%Type;
  n_�ҺŽ��   ������ü�¼.ʵ�ս��%Type;
  v_����Ա��� ������ü�¼.����Ա���%Type;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  v_���㷽ʽ   ҽ�ƿ����.���㷽ʽ%Type;
  n_ʵ�ս��   ������ü�¼.ʵ�ս��%Type;
  v_������ˮ�� ����Ԥ����¼.������ˮ��%Type;
  n_����       Number(3);
  v_Type       Varchar2(50);
  v_Temp       Varchar2(32767); --��ʱXML 
  x_Templet    Xmltype; --ģ��XML 
  v_Err_Msg    Varchar2(200);
  n_�ѿ�ҽ��   Number(2);
  n_��鷢Ʊ   Number(3);
  n_�Ƿ��ӡ   Number(3);
  n_�ɿʽ   Number(3);
  n_����id     ������ü�¼.����id%Type;
  n_����id     ������ü�¼.����id%Type;
  d_�Ǽ�ʱ��   Date;
  v_ԤԼ��ʽ   ���˹Һż�¼.ԤԼ��ʽ%Type;
  v_�շѵ�     ������ü�¼.No%Type;
  n_����id     ������ü�¼.����id%Type;
  n_�����id   ҽ�ƿ����.Id%Type;
  v_�˷ѽ���   Varchar2(1000);
  n_Column     Number(18);
  Err_Item Exception;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/GHJE'),
         Extractvalue(Value(A), 'IN/LSH'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         To_Number(Extractvalue(Value(A), 'IN/JKFS')), Extractvalue(Value(A), 'IN/YYFS')
  Into v_No, v_�����, n_�ҺŽ��, v_������ˮ��, n_��鷢Ʊ, n_�ɿʽ, v_ԤԼ��ʽ
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Max(�շѵ�) Into v_�շѵ� From ���˹Һż�¼ Where NO = v_No;

  n_�ɿʽ := Nvl(n_�ɿʽ, 0);

  If n_�ɿʽ = 1 Then
    Begin
      Select 1 Into n_���� From ������ü�¼ Where NO = v_No And ��¼���� = 4 And ����id Is Not Null And Rownum < 2;
      Select 1
      Into n_����
      From ������ü�¼
      Where NO In (Select Column_Value From Table(f_Str2list(v_�շѵ�)) B) And ��¼���� = 1 And ����id Is Not Null And
            Rownum < 2;
    Exception
      When Others Then
        n_���� := 0;
    End;
    If n_���� = 1 Then
      v_Err_Msg := '����ĹҺŵ��ݲ���ԤԼ�Һŵ�,�޷�ȡ��ԤԼ!';
      Raise Err_Item;
    End If;
    Begin
      Select 1 Into n_���� From ���˹Һż�¼ A Where a.No = v_No And a.ԤԼ��ʽ = v_ԤԼ��ʽ And Rownum < 2;
    Exception
      When Others Then
        n_���� := 0;
    End;
    If n_���� = 0 Then
      v_Err_Msg := '����ĹҺŵ��ݲ���' || v_ԤԼ��ʽ || 'ԤԼ��,�޷�ȡ��ԤԼ!';
      Raise Err_Item;
    End If;
  End If;

  If v_����� Is Not Null And n_�ɿʽ = 0 Then
    Select Nvl2(Translate(v_�����, '\1234567890', '\'), 'Char', 'Num') Into v_Type From Dual;
    If v_Type = 'Num' Then
      --������ǿ����ID 
      Select ���㷽ʽ, ID Into v_���㷽ʽ, n_�����id From ҽ�ƿ���� Where ID = To_Number(v_�����);
    Else
      --������ǿ�������� 
      Select ���㷽ʽ, ID Into v_���㷽ʽ, n_�����id From ҽ�ƿ���� Where ���� = v_�����;
    End If;
  
    Select Sum(ʵ�ս��) Into n_ʵ�ս�� From ������ü�¼ Where NO = v_No And ��¼���� = 4;
  
    If Nvl(n_�ɿʽ, 0) = 0 Then
      --Ҫ�˵ĵ��ݲ����Ըý��㿨����ģ����ֹ�˺� 
      Begin
        Select 1
        Into n_����
        From ����Ԥ����¼ A,
             (Select Distinct ����id
               From ������ü�¼
               Where NO = v_No And ��¼���� = 4
               Union
               Select Distinct ����id
               From סԺ���ü�¼
               Where NO = v_No And ��¼���� = 5
               Union
               Select Distinct ����id
               From ������ü�¼
               Where NO In (Select Column_Value From Table(f_Str2list(v_�շѵ�)) B) And ��¼���� = 1) B
        Where a.����id = b.����id And ���㷽ʽ = v_���㷽ʽ And Rownum < 2;
      Exception
        When Others Then
          n_���� := 0;
      End;
      If n_���� = 0 Then
        v_Err_Msg := '����ĹҺŵ��ݲ���' || v_���㷽ʽ || '�����,�޷��˺�!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  --��������飬�Ѵ��ڲ��������ݵģ������˺� 
  Begin
    Select 1
    Into n_����
    From ���ò����¼ A,
         (Select Distinct ����id
           From ������ü�¼
           Where NO = v_No And ��¼���� = 4
           Union
           Select Distinct ����id
           From סԺ���ü�¼
           Where NO = v_No And ��¼���� = 5
           Union
           Select Distinct ����id
           From ������ü�¼
           Where NO In (Select Column_Value From Table(f_Str2list(v_�շѵ�)) B) And ��¼���� = 1) B
    Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 1 And Nvl(a.����״̬, 0) <> 2 And Rownum < 2;
  Exception
    When Others Then
      n_���� := 0;
  End;
  If n_���� = 1 Then
    v_Err_Msg := '����ĹҺŵ����Ѿ������˶��ν���,�޷��˺�!';
    Raise Err_Item;
  End If;
  --ҽ����飬�Ѿ�����ҽ���ģ������˺� 
  Begin
    Select Distinct 1 Into n_�ѿ�ҽ�� From ����ҽ����¼ Where �Һŵ� = v_No;
  Exception
    When Others Then
      n_�ѿ�ҽ�� := 0;
  End;
  If n_�ѿ�ҽ�� = 1 Then
    v_Err_Msg := '����ĹҺŵ����Ѿ�����ҽ��,�޷��˺�!';
    Raise Err_Item;
  End If;
  If Nvl(n_��鷢Ʊ, 0) = 1 Then
    Select Max(Decode(a.ʵ��Ʊ��, Null, 0, 1)) Into n_�Ƿ��ӡ From ������ü�¼ A Where NO = v_No And ��¼���� = 4;
    If Nvl(n_�Ƿ��ӡ, 0) = 1 Then
      v_Err_Msg := '�����˺ŵĵ����ѿ���Ʊ,�����˷�!';
      Raise Err_Item;
    End If;
    Select Max(Decode(a.ʵ��Ʊ��, Null, 0, 1))
    Into n_�Ƿ��ӡ
    From ������ü�¼ A
    Where NO In (Select Column_Value From Table(f_Str2list(v_�շѵ�)) B) And ��¼���� = 1;
    If Nvl(n_�Ƿ��ӡ, 0) = 1 Then
      v_Err_Msg := '�����˺ŵĵ����ѿ���Ʊ,�����˷�!';
      Raise Err_Item;
    End If;
  End If;
  --��ȡ����Ա��Ϣ 
  v_Temp := Zl_Identity(1);
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_����Ա��� From Dual;
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_����Ա���� From Dual;
  d_�Ǽ�ʱ�� := Sysdate;

  Zl_���������Һ�_Delete(v_No, v_������ˮ��, '�ƶ�ƽ̨�˺�', d_�Ǽ�ʱ��);

  --ͬ�������۵� 
  If v_�շѵ� Is Not Null Then
  
    n_Column := 0;
    For c_�Һ� In (Select NO, Max(��¼״̬) As ��¼״̬, Max(����id) As ����id, Max(Decode(��¼״̬, 2, 0, ����id)) As ԭ����id,
                        Max(Decode(��¼״̬, 2, ����id, 0)) As ����id
                 From ������ü�¼
                 Where NO In (Select * From Table(f_Str2list(v_�շѵ�)) B) And ��¼���� = 1) Loop
      If Nvl(c_�Һ�.��¼״̬, 0) = 0 Then
        Zl_���ﻮ�ۼ�¼_Delete(c_�Һ�.No);
        n_����id := c_�Һ�.ԭ����id;
        n_����id := c_�Һ�.����id;
      Elsif Nvl(c_�Һ�.��¼״̬, 0) = 1 Then
        If v_���㷽ʽ Is Null Then
          v_Err_Msg := '���ιҺŵ����˿�ʧ��,����!';
          Raise Err_Item;
        End If;
        Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
        Zl_�����շѼ�¼_����(c_�Һ�.No, v_����Ա���, v_����Ա����, Null, d_�Ǽ�ʱ��, Null, n_����id);
        v_�˷ѽ��� := v_���㷽ʽ || '|' || -1 * n_�ҺŽ�� || '|' || ' |' || ' ';
        Zl_�����˷ѽ���_Modify(2, n_����id, n_����id, v_�˷ѽ���, 0, n_�����id, Null, v_������ˮ��, Null, 0, 0, 0, 2);
        n_����id := c_�Һ�.ԭ����id;
        n_����id := c_�Һ�.����id;
        n_Column := n_Column + 1;
      Else
        n_����id := c_�Һ�.ԭ����id;
        n_����id := c_�Һ�.����id;
      End If;
    
    End Loop;
    If n_Column > 1 Then
      v_Err_Msg := '���ιҺŴ��ڶ���շѣ������˷Ѻ����˺�!';
      Raise Err_Item;
    End If;
  
  Else
  
    Select Max(����id) Into n_����id From ������ü�¼ Where NO = v_No And ��¼���� = 4 And ��¼״̬ = 3;
    Select Max(����id) Into n_����id From ������ü�¼ Where NO = v_No And ��¼���� = 4 And ��¼״̬ = 2;
  
  End If;

  v_Temp := '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<YJZID>' || n_����id || '</YJZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<CXID>' || n_����id || '</CXID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Registdel;
/

--124650:��ҵ��,2018-04-20,�������������Ƿ��������
Create Or Replace Procedure Zl_��������_Insert
(
  ����id_In         In ��������.����id%Type,
  ����id_In         In ��������.����id%Type,
  ����_In           In �շ���ĿĿ¼.����%Type,
  ���_In           In �շ���ĿĿ¼.���%Type,
  ����_In           In �շ���ĿĿ¼.����%Type := Null,
  ��ʶ����_In       In �շ���ĿĿ¼.��ʶ����%Type := Null,
  ��ʶ����_In       In �շ���ĿĿ¼.��ʶ����%Type := Null,
  ��ѡ��_In         In �շ���ĿĿ¼.��ѡ��%Type := Null,
  ������Դ_In       In ��������.������Դ%Type := Null,
  ��Դ���_In       In ��������.��Դ���%Type := Null,
  ɢװ��λ_In       In �շ���ĿĿ¼.���㵥λ%Type := Null,
  ��װ��λ_In       In ��������.��װ��λ%Type := Null,
  ����ϵ��_In       In ��������.����ϵ��%Type := Null,
  �Ƿ���_In       In �շ���ĿĿ¼.�Ƿ���%Type := Null,
  ָ��������_In     In ��������.ָ��������%Type := Null,
  ����_In           In ��������.����%Type := 95,
  ָ�����ۼ�_In     In ��������.ָ�����ۼ�%Type := Null,
  ָ�������_In     In ��������.ָ�������%Type := Null,
  ��������_In       In �շ���ĿĿ¼.��������%Type := Null,
  �������_In       In �շ���ĿĿ¼.�������%Type := Null,
  ���ηѱ�_In       In �շ���ĿĿ¼.���ηѱ�%Type := 0,
  �ⷿ����_In       In ��������.�ⷿ����%Type := Null,
  ���÷���_In       In ��������.���÷���%Type := Null,
  ���Ч��_In       In ��������.���Ч��%Type := Null,
  ���Ч��_In       In ��������.���Ч��%Type := Null,
  �޾��Բ���_In     In ��������.�޾��Բ���%Type := Null,
  һ���Բ���_In     In ��������.һ���Բ���%Type := Null,
  ԭ����_In         In ��������.ԭ����%Type := Null,
  ���������_In     In ��������.���������%Type := 0,
  �ɱ���_In         In ��������.�ɱ���%Type := 0,
  ��������_In       In ��������.��������%Type := Null,
  �������_In       In ��������.�������%Type := 0,
  ��ǰ�ۼ�_In       In �շѼ�Ŀ.�ּ�%Type := 0,
  ����id_In         In �շѼ�Ŀ.������Ŀid%Type := Null,
  ��׼�ĺ�_In       In ��������.��׼�ĺ�%Type := Null,
  ע���̱�_In       In ��������.ע���̱�%Type := Null,
  ע��֤��_In       In ��������.ע��֤��%Type := Null,
  ���֤��_In       In ��������.���֤��%Type := Null,
  ���֤��Ч��_In   In ��������.���֤��Ч��%Type := Null,
  ���ʷ���_In       In ��������.���ʷ���%Type := Null,
  �洢����_In       In ��������.�洢����%Type := Null,
  ���ٲ���_In       In ��������.���ٲ���%Type := 0,
  վ��_In           In �շ���ĿĿ¼.վ��%Type := Null,
  Ʒ��_In           In �շ���Ŀ����.����%Type := Null,
  ƴ��_In           In �շ���Ŀ����.����%Type := Null,
  ���_In           In �շ���Ŀ����.����%Type := Null,
  ��ֵ˰��_In       In ��������.��ֵ˰��%Type := Null,
  ˵��_In           In �շ���ĿĿ¼.˵��%Type := Null,
  ��ֵ����_In       In ��������.��ֵ����%Type := Null,
  �������_In       In ��������.�Ƿ��������%Type := Null,
  ������Ŀ_In       In �շ���ĿĿ¼.������Ŀ%Type := Null,
  ��е�����ĵ���_In In ��������.��е�����ĵ���%Type := 0,
  ע��֤��Ч��_In   In ��������.ע��֤��Ч��%Type := Null,
  �Ƿ�ֲ��Ĳ�_In   In ��������.�Ƿ�ֲ��Ĳ�%Type := 0,
  �ӳ���_In         In ��������.�ӳ���%Type := Null,
  ����ʹ��_In       In ��������.�Ƿ����%Type := 0
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(100);

  v_No       �շѼ�Ŀ.No%Type;
  v_����     ������ĿĿ¼.����%Type;
  v_Temp     �շ���ĿĿ¼.������Ŀ%Type;
  v_������Ŀ �շ���ĿĿ¼.������Ŀ%Type;

  Cursor c_Item Is
    Select ID
    From ���ű� D
    Where ID In (Select Distinct ����id
                 From ��������˵�� A
                 Where �������� In ('���ϲ���', '���ʿⷿ', '���Ŀ�', '�Ƽ���', '����ⷿ'));
Begin
  v_Err_Msg := 'NO';

  v_������Ŀ := ������Ŀ_In;
  --�жϲ�����Ŀ 
  If v_������Ŀ Is Null Then
    If ����id_In Is Not Null Then
      Begin
        Select ������Ŀ Into v_Temp From ������Ŀ Where ID = ����id_In;
      Exception
        When Others Then
          v_Temp := Null;
      End;
      If v_Temp Is Not Null Then
        v_������Ŀ := v_Temp;
      End If;
    End If;
  End If;
  Begin
    Select ����
    Into v_����
    From ������ĿĿ¼
    Where ID = ����id_In And (����ʱ�� Is Null Or To_Char(����ʱ��, 'yyyy-mm-dd') = '3000-01-01');
  Exception
    When Others Then
      v_Err_Msg := 'Err';
  End;
  If v_Err_Msg = 'Err' Then
    v_Err_Msg := '[ZLSOFT]δ�ҵ�ָ���Ĳ���Ʒ�֣����ܸ�Ʒ���ѱ������û�ɾ����ͣ�ã�[ZLSOFT]';
    Raise Err_Item;
  End If;
  --�����Ϣ 
  Insert Into �շ���ĿĿ¼
    (���, ID, ����, ����, ���, ����, ��ʶ����, ��ʶ����, ��ѡ��, ���㵥λ, ��������, �������, ���ηѱ�, �Ƿ���, վ��, ����ʱ��, ����ʱ��, ˵��, ������Ŀ)
  Values
    (4, ����id_In, ����_In, v_����, ���_In, ����_In, ��ʶ����_In, ��ʶ����_In, ��ѡ��_In, ɢװ��λ_In, ��������_In, �������_In, ���ηѱ�_In, �Ƿ���_In,
     վ��_In, Sysdate, To_Date('3000-01-01', 'YYYY-MM-DD'), ˵��_In, v_������Ŀ);

  --�������� 
  Insert Into ��������
    (����id, ����id, ���Ч��, ���Ч��, �޾��Բ���, һ���Բ���, ԭ����, ��Դ���, ��װ��λ, ����ϵ��, ָ��������, ָ�����ۼ�, ָ�������, ����, �ⷿ����, ���÷���, ������Դ, ���������, �ɱ���,
     ��������, �������, ��׼�ĺ�, ע���̱�, ע��֤��, ע��֤��Ч��, ���֤��, ���֤��Ч��, ���ʷ���, �洢����, ���ٲ���, ��ֵ˰��, ��ֵ����, �Ƿ��������, ��е�����ĵ���, �Ƿ�ֲ��Ĳ�, �ӳ���,
     �Ƿ����)
  Values
    (����id_In, ����id_In, ���Ч��_In, ���Ч��_In, �޾��Բ���_In, һ���Բ���_In, ԭ����_In, ��Դ���_In, ��װ��λ_In, ����ϵ��_In, ָ��������_In, ָ�����ۼ�_In,
     ָ�������_In, ����_In, �ⷿ����_In, ���÷���_In, ������Դ_In, ���������_In, �ɱ���_In, ��������_In, �������_In, ��׼�ĺ�_In, ע���̱�_In, ע��֤��_In,
     ע��֤��Ч��_In, ���֤��_In, ���֤��Ч��_In, ���ʷ���_In, �洢����_In, ���ٲ���_In, ��ֵ˰��_In, ��ֵ����_In, �������_In, ��е�����ĵ���_In, �Ƿ�ֲ��Ĳ�_In, �ӳ���_In,
     ����ʹ��_In);

  --�����Ĵ��� 
  Insert Into �շ���Ŀ����
    (�շ�ϸĿid, ����, ����, ����, ����)
    Select ����id_In, ����, ����, ����, ���� From ������Ŀ���� Where ������Ŀid = ����id_In;
  If (Ʒ��_In Is Not Null) And (ƴ��_In Is Not Null) Then
    Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (����id_In, Ʒ��_In, 3, ƴ��_In, 1);
  End If;
  If (Ʒ��_In Is Not Null) And (���_In Is Not Null) Then
    Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (����id_In, Ʒ��_In, 3, ���_In, 2);
  End If;

  For r_Item In c_Item Loop
    Insert Into ���ϴ����޶� (�ⷿid, ����id, ����, ����, �̵�����) Values (r_Item.Id, ����id_In, 0, 0, '1111');
  End Loop;
  --������Ϣ 
  If ����id_In Is Not Null Then
    v_No := Nextno(9);
    --�Ǹ������õ�ʱ�����ģ�������ʱ�൱��һ����շ���Ŀ���ڵ���ʱӦ��������޼ۡ�����޼ۡ�ȱʡ�۸񡱽������á� 
    Insert Into �շѼ�Ŀ
      (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ȱʡ�۸�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
    Values
      (�շѼ�Ŀ_Id.Nextval, Null, ����id_In, 0, ��ǰ�ۼ�_In, Decode(��������_In, 0, Decode(�Ƿ���_In, 1, ��ǰ�ۼ�_In, Null), Null), ����id_In,
       1, '��������', User, Sysdate, To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  End If;

  --���������̱Ƚ����� 
  If ����_In Is Not Null Then
    Update ���������� Set ���� = ����_In Where ���� = ����_In;
    If Sql%RowCount = 0 Then
      Insert Into ����������
        (����, ����, ����)
        Select Nvl(Max(To_Number(����)), 0) + 1, ����_In, zlSpellCode(����_In) From ����������;
    End If;
  End If;

  --������ϵķ������ 
  Insert Into �շ�ִ�п���
    (�շ�ϸĿid, ������Դ, ��������id, ִ�п���id)
    Select ����id_In, ������Դ, ��������id, ִ�п���id From ����ִ�п��� Where ������Ŀid = ����id_In;

  b_Message.Zlhis_Dict_043(����id_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��������_Insert;
/

--124650:��ҵ��,2018-04-20,�������������Ƿ��������
Create Or Replace Procedure Zl_��������_Update
(
  ����id_In         In ��������.����id%Type,
  ����id_In         In ��������.����id%Type,
  ����_In           In �շ���ĿĿ¼.����%Type,
  ���_In           In �շ���ĿĿ¼.���%Type,
  ����_In           In �շ���ĿĿ¼.����%Type := Null,
  ��ʶ����_In       In �շ���ĿĿ¼.��ʶ����%Type := Null,
  ��ʶ����_In       In �շ���ĿĿ¼.��ʶ����%Type := Null,
  ��ѡ��_In         In �շ���ĿĿ¼.��ѡ��%Type := Null,
  ������Դ_In       In ��������.������Դ%Type := Null,
  ��Դ���_In       In ��������.��Դ���%Type := Null,
  ɢװ��λ_In       In �շ���ĿĿ¼.���㵥λ%Type := Null,
  ��װ��λ_In       In ��������.��װ��λ%Type := Null,
  ����ϵ��_In       In ��������.����ϵ��%Type := Null,
  �Ƿ���_In       In �շ���ĿĿ¼.�Ƿ���%Type := Null,
  ָ��������_In     In ��������.ָ��������%Type := Null,
  ����_In           In ��������.����%Type := 95,
  ָ�����ۼ�_In     In ��������.ָ�����ۼ�%Type := Null,
  ָ�������_In     In ��������.ָ�������%Type := Null,
  ��������_In       In �շ���ĿĿ¼.��������%Type := Null,
  �������_In       In �շ���ĿĿ¼.�������%Type := Null,
  ���ηѱ�_In       In �շ���ĿĿ¼.���ηѱ�%Type := 0,
  �ⷿ����_In       In ��������.�ⷿ����%Type := Null,
  ���÷���_In       In ��������.���÷���%Type := Null,
  ���Ч��_In       In ��������.���Ч��%Type := Null,
  ���Ч��_In       In ��������.���Ч��%Type := Null,
  �޾��Բ���_In     In ��������.�޾��Բ���%Type := Null,
  һ���Բ���_In     In ��������.һ���Բ���%Type := Null,
  ԭ����_In         In ��������.ԭ����%Type := Null,
  ���������_In     In ��������.���������%Type := 0,
  �ɱ���_In         In ��������.�ɱ���%Type := 0,
  ��������_In       In ��������.��������%Type := Null,
  �������_In       In ��������.�������%Type := 0,
  ��ǰ�ۼ�_In       In �շѼ�Ŀ.�ּ�%Type := 0,
  ����id_In         In �շѼ�Ŀ.������Ŀid%Type := Null,
  ��׼�ĺ�_In       In ��������.��׼�ĺ�%Type := Null,
  ע���̱�_In       In ��������.ע���̱�%Type := Null,
  ע��֤��_In       In ��������.ע��֤��%Type := Null,
  ���֤��_In       In ��������.���֤��%Type := Null,
  ���֤��Ч��_In   In ��������.���֤��Ч��%Type := Null,
  ���ʷ���_In       In ��������.���ʷ���%Type := Null,
  �洢����_In       In ��������.�洢����%Type := Null,
  ���ٲ���_In       In ��������.���ٲ���%Type := 0,
  վ��_In           In �շ���ĿĿ¼.վ��%Type := Null,
  Ʒ��_In           In �շ���Ŀ����.����%Type := Null,
  ƴ��_In           In �շ���Ŀ����.����%Type := Null,
  ���_In           In �շ���Ŀ����.����%Type := Null,
  ��ֵ˰��_In       In ��������.��ֵ˰��%Type := Null,
  ˵��_In           In �շ���ĿĿ¼.˵��%Type := Null,
  ��ֵ����_In       In ��������.��ֵ����%Type := Null,
  �������_In       In ��������.�Ƿ��������%Type := Null,
  ������Ŀ_In       In �շ���ĿĿ¼.������Ŀ%Type := Null,
  ��е�����ĵ���_In In ��������.��е�����ĵ���%Type := 0,
  ע��֤��Ч��_In   In ��������.ע��֤��Ч��%Type := Null,
  �Ƿ�ֲ��Ĳ�_In   In ��������.�Ƿ�ֲ��Ĳ�%Type := 0,
  �޸�����_In       In Number := 0, --1-ͬ���޸�Ʒ��������ע��֤�ź�ע��֤��Ч��
  �ӳ���_In         In ��������.�ӳ���%Type := Null,
  ����ʹ��_In       In ��������.�Ƿ����%Type := 0
) Is
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  v_����     Integer;
  v_�������� Integer;
  v_Count    Integer;
  v_No       �շѼ�Ŀ.No%Type;
  v_����     ������ĿĿ¼.����%Type;
  v_Temp     �շ���ĿĿ¼.������Ŀ%Type;
  v_������Ŀ �շ���ĿĿ¼.������Ŀ%Type;

Begin
  v_Err_Msg := '��';

  v_������Ŀ := ������Ŀ_In;
  --�жϲ�����Ŀ 
  If v_������Ŀ Is Null Then
    If ����id_In Is Not Null Then
      Begin
        Select ������Ŀ Into v_Temp From ������Ŀ Where ID = ����id_In;
      Exception
        When Others Then
          v_Temp := Null;
      End;
      If v_Temp Is Not Null Then
        v_������Ŀ := v_Temp;
      End If;
    End If;
  End If;
  --�޸�������Ŀ 
  Begin
    Select �������� Into v_�������� From �������� Where ����id = ����id_In;
  Exception
    When Others Then
      v_Err_Msg := '[ZLSOFT]�����ڹ�����,���ܱ������û�ɾ����,����![ZLSOFT]';
  End;
  If v_Err_Msg <> '��' Then
    Raise Err_Item;
  End If;

  Begin
    Select ���� Into v_���� From ������ĿĿ¼ Where ID = ����id_In;
  Exception
    When Others Then
      v_Err_Msg := 'Err';
  End;

  If v_Err_Msg = 'Err' Then
    v_Err_Msg := '[ZLSOFT]δ�ҵ�ָ���Ĳ���Ʒ�֣������ѱ������û�ɾ����[ZLSOFT]';
    Raise Err_Item;
  End If;

  --�������ǰ�Ĳ���Ϊ��������,�����Ϊ�˲����������жϿ�� 
  If v_�������� = 1 And ��������_In <> 1 Then
    Begin
      Select Count(*)
      Into v_Count
      From ҩƷ���
      Where ҩƷid = ����id_In And (Nvl(��������, 0) <> 0 Or Nvl(ʵ������, 0) <> 0 Or Nvl(ʵ�ʽ��, 0) <> 0 Or Nvl(ʵ�ʲ��, 0) <> 0);
      If v_Count <> 0 Then
        v_Err_Msg := '[ZLSOFT]���������ϴ��ڿ��,����ȡ��������������,����![ZLSOFT]';
      End If;
    Exception
      When Others Then
        Null;
    End;
  End If;

  If v_Err_Msg <> '��' Then
    Raise Err_Item;
  End If;

  --�����Ϣ 
  Update �շ���ĿĿ¼
  Set ���� = ����_In, ���� = v_����, ��� = ���_In, ��ʶ���� = ��ʶ����_In, ��ʶ���� = ��ʶ����_In, ��ѡ�� = ��ѡ��_In, ���� = ����_In, �Ƿ��� = �Ƿ���_In,
      ���㵥λ = ɢװ��λ_In, �������� = ��������_In, ������� = �������_In, ���ηѱ� = ���ηѱ�_In, վ�� = վ��_In, ˵�� = ˵��_In, ������Ŀ = v_������Ŀ
  Where ID = ����id_In;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]���������Ͽ��ܱ������û�ɾ����,����![ZLSOFT]';
    Raise Err_Item;
  End If;

  --�������� 
  Update ��������
  Set ���Ч�� = ���Ч��_In, ���Ч�� = ���Ч��_In, �޾��Բ��� = �޾��Բ���_In, һ���Բ��� = һ���Բ���_In, ԭ���� = ԭ����_In, ��Դ��� = ��Դ���_In, ��װ��λ = ��װ��λ_In,
      ����ϵ�� = ����ϵ��_In, ָ�������� = ָ��������_In, ָ�����ۼ� = ָ�����ۼ�_In, ָ������� = ָ�������_In, ���� = ����_In, �ⷿ���� = �ⷿ����_In, ���÷��� = ���÷���_In,
      ������Դ = ������Դ_In, ��������� = ���������_In, �ɱ��� = �ɱ���_In, �������� = ��������_In, ������� = �������_In, ��׼�ĺ� = ��׼�ĺ�_In, ע���̱� = ע���̱�_In,
      ע��֤�� = ע��֤��_In, ע��֤��Ч�� = ע��֤��Ч��_In, ���ʷ��� = ���ʷ���_In, �洢���� = �洢����_In, ���֤�� = ���֤��_In, ���֤��Ч�� = ���֤��Ч��_In,
      ����id = ����id_In, ���ٲ��� = ���ٲ���_In, ��ֵ˰�� = ��ֵ˰��_In, ��ֵ���� = ��ֵ����_In, �Ƿ�������� = �������_In, ��е�����ĵ��� = ��е�����ĵ���_In,
      �Ƿ�ֲ��Ĳ� = �Ƿ�ֲ��Ĳ�_In, �ӳ��� = �ӳ���_In, �Ƿ���� = ����ʹ��_In
  Where ����id = ����id_In;

  --ͬ���޸ĸ�Ʒ�������й��
  If �޸�����_In = 1 Then
    Update �������� Set ע��֤�� = ע��֤��_In, ע��֤��Ч�� = ע��֤��Ч��_In Where ����id = ����id_In;
  End If;

  --�����Ĵ��� 
  Delete �շ���Ŀ���� Where �շ�ϸĿid = ����id_In And ���� = 1;

  Insert Into �շ���Ŀ����
    (�շ�ϸĿid, ����, ����, ����, ����)
    Select ����id_In, ����, ����, ����, ���� From ������Ŀ���� Where ������Ŀid = ����id_In And ���� = 1;

  If Ʒ��_In Is Null Then
    Delete �շ���Ŀ���� Where �շ�ϸĿid = ����id_In And ���� = 3;
  Else
    If ƴ��_In Is Null Then
      Delete �շ���Ŀ���� Where �շ�ϸĿid = ����id_In And ���� = 3 And ���� = 1;
    Else
      Update �շ���Ŀ���� Set ���� = Ʒ��_In, ���� = ƴ��_In Where �շ�ϸĿid = ����id_In And ���� = 3 And ���� = 1;
      If Sql%RowCount = 0 Then
        Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (����id_In, Ʒ��_In, 3, ƴ��_In, 1);
      End If;
    End If;
    If ���_In Is Null Then
      Delete �շ���Ŀ���� Where �շ�ϸĿid = ����id_In And ���� = 3 And ���� = 2;
    Else
      Update �շ���Ŀ���� Set ���� = Ʒ��_In, ���� = ���_In Where �շ�ϸĿid = ����id_In And ���� = 3 And ���� = 2;
      If Sql%RowCount = 0 Then
        Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (����id_In, Ʒ��_In, 3, ���_In, 2);
      End If;
    End If;
  End If;

  --������Ϣ������Ѿ��з�����������ֱ�Ӹ�����Щ��Ϣ 
  Select Nvl(Count(*), 0) Into v_���� From ҩƷ�շ���¼ Where ҩƷid = ����id_In And Rownum < 2;

  If v_���� = 0 Then
    Update �շ���ĿĿ¼ Set �Ƿ��� = �Ƿ���_In Where ID = ����id_In;
    Update �������� Set �ɱ��� = �ɱ���_In Where ����id = ����id_In;
  
    If ����id_In Is Not Null Then
      Update �շѼ�Ŀ
      Set �ּ� = ��ǰ�ۼ�_In, ȱʡ�۸� = Decode(��������_In, 0, Decode(�Ƿ���_In, 1, ��ǰ�ۼ�_In, ȱʡ�۸�), ȱʡ�۸�), ������Ŀid = ����id_In, �䶯ԭ�� = 1,
          ����˵�� = '�޸Ķ���', ������ = User
      Where �շ�ϸĿid = ����id_In
           --And (��ֹ���� Is Null Or ��ֹ����=to_date('3000-01-01','YYYY-MM-DD')); 
            And (Sysdate Between ִ������ And ��ֹ���� Or Sysdate >= ִ������ And ��ֹ���� Is Null) And �䶯ԭ�� = 1;
    
      If Sql%RowCount = 0 Then
        v_No := Nextno(9);
        Insert Into �շѼ�Ŀ
          (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ȱʡ�۸�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
        Values
          (�շѼ�Ŀ_Id.Nextval, Null, ����id_In, 0, ��ǰ�ۼ�_In, Decode(��������_In, 0, Decode(�Ƿ���_In, 1, ��ǰ�ۼ�_In, Null), Null),
           ����id_In, 1, '��������', User, Sysdate, To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
      End If;
    End If;
  Else
    --��ҵ�񵥾ݺ���ֱ���޸ļ۸񣬵��ǿ����޸�������Ŀ 
    Update �շѼ�Ŀ
    Set ������Ŀid = ����id_In
    Where �շ�ϸĿid = ����id_In And (Sysdate Between ִ������ And ��ֹ���� Or Sysdate >= ִ������ And ��ֹ���� Is Null) And �䶯ԭ�� = 1;
  End If;

  --���������̱Ƚ����� 
  If ����_In Is Not Null Then
    Update ���������� Set ���� = ����_In Where ���� = ����_In;
    If Sql%RowCount = 0 Then
      Insert Into ����������
        (����, ����, ����)
        Select Nvl(Max(To_Number(����)), 0) + 1, ����_In, zlSpellCode(����_In) From ����������;
    End If;
  End If;

  b_Message.Zlhis_Dict_044(����id_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��������_Update;
/





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0007' Where ���=&n_System;
Commit;