----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------



------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--125974:������,2018-05-18,��ҽ��¼����ҽ���
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1252, 0, 0, 0, 0, 0, 0, 68, '������ҽ������¼����ҽ���', '0', '0',
         '�����ò������������������ޡ���ҽ�ơ�����ʱ��д���ʱ�����´���ҽ���', '0-��ʾ������,1-��ʾ����', '���������������ﳡ��', '��������ҽ����Ҫ¼����ҽ��ϵ����', Null
  From Dual;  

--119329:Ƚ����,2018-05-16,�����ӿڻ�ȡ�ɹҺſ��ҹ��̵���
Declare
  --���ܣ������ٴ�����Һſ���
  Cursor c_���� Is
    Select Rowid From �ٴ�����Һſ��� Where ���Ʒ�ʽ = 3 And ��� = 0 And ���� = 0;

  c_Rowid      t_Strlist := t_Strlist();
  n_Array_Size Number := 10000; --ÿ��һ��,���˿���PGA����
  I            Number(8) := 0; --ÿ����10������¼�ύһ��,���˿���Undo����,�����ύ����Ƶ��
  J            Number(16) := 0;
Begin
  Open c_����();
  Loop
    Fetch c_���� Bulk Collect
      Into c_Rowid Limit n_Array_Size;
    Exit When c_Rowid.Count = 0;
  
    Forall K In 1 .. c_Rowid.Count
      Update �ٴ�����Һſ��� Set ���Ʒ�ʽ = 4 Where Rowid = c_Rowid(K);
  
    J := J + c_Rowid.Count;
    If I = 10 Then
      Commit;
      I := 0;
    Else
      I := I + 1;
    End If;
  End Loop;
  Close c_����;
  Commit;
End;
/

--119329:Ƚ����,2018-05-16,�����ӿڻ�ȡ�ɹҺſ��ҹ��̵���
Declare
  --���ܣ������ٴ�����Һſ��Ƽ�¼
  Cursor c_��¼ Is
    Select Rowid From �ٴ�����Һſ��Ƽ�¼ Where ���Ʒ�ʽ = 3 And ��� = 0 And ���� = 0;

  c_Rowid      t_Strlist := t_Strlist();
  n_Array_Size Number := 10000; --ÿ��һ��,���˿���PGA����
  I            Number(8) := 0; --ÿ����10������¼�ύһ��,���˿���Undo����,�����ύ����Ƶ��
  J            Number(16) := 0;
Begin
  Open c_��¼();
  Loop
    Fetch c_��¼ Bulk Collect
      Into c_Rowid Limit n_Array_Size;
    Exit When c_Rowid.Count = 0;
  
    Forall K In 1 .. c_Rowid.Count
      Update �ٴ�����Һſ��Ƽ�¼ Set ���Ʒ�ʽ = 4 Where Rowid = c_Rowid(K);
  
    J := J + c_Rowid.Count;
    If I = 10 Then
      Commit;
      I := 0;
    Else
      I := I + 1;
    End If;
  End Loop;
  Close c_��¼;
  Commit;
End;
/




-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--125932:������,2018-05-17,����Ȩ��ȱʧ
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1252,'ҽ���´�',User,'Zl_Fun_Bloodapplyrate','EXECUTE' From Dual;

--125588:����,2018-05-16,�ų�����Ԥ����¼
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1333, '����', User, 'Zl_Fun_Getbatchpro', 'EXECUTE'
  From Dual;






-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--122403:����,2018-06-01,����ҽ������ʱ�ֽ���Һ�����εĴ���
Create Or Replace Procedure Zl_��Һ��ҩ��¼_�˲�
(
  ����id_In   In ��Һ��ҩ��¼.����id%Type,
  ҽ��id_In   In Varchar2, --��Һҽ����ҩ;����Ӧ��ҽ��ID:ҽ��ID1,ҽ��ID2...
  ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
  �˲���_In   In ��Һ��ҩ״̬.������Ա%Type,
  �˲�ʱ��_In In ��Һ��ҩ״̬.����ʱ��%Type
) Is
  v_Count    Number;
  v_���     Number;
  v_ִ��ʱ�� Date;

  v_���id      Number;
  v_New���id   Number;
  v_Old���id   Number;
  v_���ͺ�      Number;
  v_Tmp         Varchar2(200);
  I             Number;
  v_��ҩid      Number;
  v_����        Number;
  v_Maxno       Varchar2(4000);
  v_Lableno     Varchar2(200);
  v_Maxbatch    Number;
  v_Curdose     Number;
  v_Sumdose     Number;
  v_Drugcount   Number;
  v_Currdate    Date;
  n_Needcheck   Number;
  n_Lngid       ҩƷ�շ���¼.Id%Type;
  n_Count       Number(3);
  n_����        ҩƷ�շ���¼.����%Type;
  v_No          ҩƷ�շ���¼.No%Type;
  n_���ʹ���    Number(5);
  n_����id      ������Ϣ.����id%Type := 0;
  b_Change      Boolean := True;
  n_Sum         Number;
  n_��������    Number(1);
  n_Cur         Number(5);
  v_�ϴη��ͺ�  ����ҽ������.���ͺ�%Type;
  v_ҽ��ids     Varchar2(4000);
  v_Tansid      Varchar2(12);
  v_��ǰ����    Varchar2(20);
  n_Num         Number(8);
  d_Oldִ��ʱ�� Date;
  n_�Ƿ���    Number(1);
  n_���        Number(1);
  n_��ҩ��      Number(2);
  --���Ʋ���
  v_ҽ������       Number;
  v_��Һ����       Number;
  v_����Һ����     Varchar2(2000);
  v_����Һ��ҩ;�� Varchar2(2000);
  v_��Դ����       Varchar2(4000);
  v_Continue       Number := 1;
  v_Nodosage       Number := 0;
  v_�����ϴ�����   Number := 0;
  d_�ֹ����ʱ��   Date;
  n_Tpn���÷�ʽ    Number := 0;
  v_ҩƷ����       Varchar2(20);
  n_���ҩƷ����   Number(1);
  n_����ҩƷ����   Number(1);
  n_���ȼ�         Number := 999;
  n_�Զ�����       Number := 0;
  n_����id         Number := 0;
  n_Row            Number(2);
  n_��������       Number := 0;
  n_ʣ������       Number := 0;
  n_��������       Number := 0;
  n_�ۼ�����       Number := 0;
  n_ҽ��id         Number := 0;
  n_��д����       Number := 0;
  v_��ҩ����       Varchar2(20);
  v_ʱ�䴮         Varchar2(100);
  v_ʱ��ֵ         Date;
  v_Fields         Varchar2(100);
  v_�Ƿ�ı�       Varchar2(20);
  v_ʱ�䴮1        Varchar2(100);
  Err_Item Exception;
  n_��ͨ���С�� Number;

  Cursor c_ҽ����¼ Is
    Select /*+rule */
    Distinct e.ҽ��id As ���id, e.���ͺ�, b.Ƶ�ʼ��, b.�����λ, b.ִ��ʱ�䷽��, a.����, a.�Ա�, a.����, a.סԺ��, a.��ǰ���� As ����, a.��ǰ����id As ���˲���id,
             a.��ǰ����id As ���˿���id, e.�״�ʱ��, e.ĩ��ʱ��, b.��ʼִ��ʱ��, Nvl(e.��������, 0) As ����, e.����ʱ��, Decode(b.ҽ����Ч, 0, 1, 2) As ҽ������,
             b.������Ŀid As ��ҩ;��, b.����id, Nvl(c.ִ�б��, 0) As �Ƿ�tpn
    From ����ҽ������ E, ����ҽ����¼ B, ������Ϣ A, ������ĿĿ¼ C, Table(f_Num2list(ҽ��id_In)) D
    Where e.ҽ��id = b.Id And b.����id = a.����id And c.��� = 'E' And c.�������� = '2' And c.ִ�з��� = 1 And b.������Ŀid = c.Id And
          e.ҽ��id = d.Column_Value And e.���ͺ� = ���ͺ�_In
    Order By b.����id, e.ҽ��id, e.���ͺ�;

  Cursor c_����ҽ����¼ Is
    Select /*+rule */
    Distinct e.ҽ��id, e.���ͺ�, b.Ƶ�ʼ��, b.�����λ, b.ִ��ʱ�䷽��, a.����, a.�Ա�, a.����, a.סԺ��, a.��ǰ���� As ����, a.��ǰ����id As ���˲���id,
             a.��ǰ����id As ���˿���id, e.�״�ʱ��, e.ĩ��ʱ��, b.��ʼִ��ʱ��, Nvl(e.��������, 0) As ����, e.����ʱ��, Decode(b.ҽ����Ч, 0, 1, 2) As ҽ������,
             b.������Ŀid As ��ҩ;��, b.����id
    From ����ҽ������ E, ����ҽ����¼ B, ������Ϣ A, ������ĿĿ¼ C
    Where e.ҽ��id = b.Id And b.����id = a.����id And b.������Ŀid = c.Id And b.���id = v_���id And e.���ͺ� = ���ͺ�_In
    Order By e.ҽ��id, e.���ͺ�;

  Cursor c_�շ���¼ Is
    Select Distinct c.Id As �շ�id, c.���, c.ʵ������ As ����, Nvl(e.�Ƿ�������, 0) As �Ƿ�������, c.����, c.No
    From ����ҽ����¼ A, ����ҽ������ B, ҩƷ�շ���¼ C, סԺ���ü�¼ D, ��ҺҩƷ���� E
    Where a.Id = b.ҽ��id And c.����id = d.Id And a.Id = d.ҽ����� And b.No = c.No And c.ҩƷid = e.ҩƷid(+) And
          b.ִ�в���id + 0 = ����id_In And c.���� = 9 And c.������� Is Null And a.Id = n_ҽ��id And b.���ͺ� = ���ͺ�_In And c.��� < 1000
    Order By c.No, c.���;

  Cursor c_ԭʼ�շ���¼ Is
    Select Distinct c.Id As �շ�id, c.���, c.ʵ������ As ����, Nvl(e.�Ƿ�������, 0) As �Ƿ�������, c.����, c.No
    From ����ҽ����¼ A, ����ҽ������ B, ҩƷ�շ���¼ C, סԺ���ü�¼ D, ��ҺҩƷ���� E
    Where a.Id = b.ҽ��id And c.����id = d.Id And a.Id = d.ҽ����� And b.No = c.No And c.ҩƷid = e.ҩƷid(+) And
          b.ִ�в���id + 0 = ����id_In And c.���� = 9 And c.������� Is Null And a.���id = v_���id And b.���ͺ� = ���ͺ�_In And c.��� < 1000
    Order By c.No, c.���;

  Cursor c_��Һ����¼ Is
    Select a.Id, a.ִ��ʱ��, a.��ҩ����, a.ҽ��id, d.����ʱ��
    From ��Һ��ҩ��¼ A, ����ҽ����¼ B, ��ҩ�������� C, ����ҽ������ D
    Where a.ҽ��id = b.Id And a.��ҩ���� = c.���� And d.ҽ��id = a.ҽ��id And a.���ͺ� = d.���ͺ� And c.���� <> 0 And c.ҩƷ���� Is Null And
          b.����id = n_����id And a.����״̬ < 2 And a.ִ��ʱ�� Between Trunc(v_ʱ��ֵ) And Trunc(v_ʱ��ֵ + 1) - 1 / 24 / 60 / 60;

  v_��Һ����¼   c_��Һ����¼%RowType;
  v_ҽ����¼     c_ҽ����¼%RowType;
  v_�շ���¼     c_�շ���¼%RowType;
  v_����ҽ����¼ c_����ҽ����¼%RowType;

  Function Zl_Getpivaworkbatch
  (
    ִ��ʱ��_In In Date,
    ����ʱ��_In In Date,
    ҩƷ����_In In Varchar2 := Null
  ) Return Number As
  
    v_Exetime   Date;
    v_Starttime Date;
    v_Endtime   Date;
    v_Maxbatch  Number(2);
    v_Batch     Number;
    Cursor c_��ҩ���� Is
      Select ����, ��ҩʱ��, ��ҩʱ��, ���, ҩƷ����
      From ��ҩ��������
      Where ���� = 1 And ��������id = ����id_In
      Order By ҩƷ����, ����;
  
    v_��ҩ���� c_��ҩ����%RowType;
  Begin
    v_Exetime := To_Date(Substr(To_Char(ִ��ʱ��_In, 'yyyy-mm-dd hh24:mi'), 12), 'hh24:mi');
  
    Select Max(Nvl(����, 0)) + 1 Into v_Maxbatch From ��ҩ�������� Where ���� = 1 And ��������id = ����id_In;
  
    For v_��ҩ���� In c_��ҩ���� Loop
      v_Batch := 0;
    
      --���췢�͵�ҽ�����͵���������
      If (Trunc(ִ��ʱ��_In) >= Trunc(v_Currdate) And Trunc(����ʱ��_In) < Trunc(ִ��ʱ��_In)) Or n_�������� = 0 Then
        If v_��ҩ����.���� <> '0' And
           ((Nvl(v_��ҩ����.ҩƷ����, '0') <> '0' And v_��ҩ����.ҩƷ���� = ҩƷ����_In) Or Nvl(v_��ҩ����.ҩƷ����, '0') = '0') Then
          v_Starttime := To_Date(Substr(v_��ҩ����.��ҩʱ��, 1, Instr(v_��ҩ����.��ҩʱ��, '-') - 1), 'hh24:mi');
          v_Endtime   := To_Date(Substr(v_��ҩ����.��ҩʱ��, Instr(v_��ҩ����.��ҩʱ��, '-') + 1), 'hh24:mi');
        
          If v_Exetime >= v_Starttime And v_Exetime <= v_Endtime Then
            v_Batch := v_��ҩ����.����;
            n_���  := v_��ҩ����.���;
            Exit When v_Batch > 0;
          End If;
        End If;
      End If;
    End Loop;
  
    If v_Batch = 0 And (n_���ҩƷ���� <> 1 Or n_�������� = 1) Then
      v_Batch := v_Maxbatch;
    End If;
    Return(v_Batch);
  End;

  Function Zl_Getfirst
  (
    ��ҩid_In In Number,
    ����id_In In Number
  ) Return Number As
    n_First  Number;
    n_����id Number;
    Cursor c_���ȼ� Is
      Select ����id, ��ҩ����, ���ȼ�, Ƶ��
      From ��ҺҩƷ���ȼ�
      Where (����id = ����id_In Or ����id = 0)
      Order By ����id, ���ȼ� Desc;
  
    r_���ȼ� c_���ȼ�%RowType;
  Begin
    n_First := 0;
    For r_���ȼ� In c_���ȼ� Loop
      If n_����id <> 0 And r_���ȼ�.����id = 0 Then
        Exit;
      End If;
      n_����id := r_���ȼ�.����id;
    
      For r_��ҩ��¼ In (Select Distinct d.��ҩ����, e.ִ��Ƶ��
                     From ��Һ��ҩ��¼ A, ��Һ��ҩ���� B, ҩƷ�շ���¼ C, ��ҺҩƷ���� D, ����ҽ����¼ E
                     Where a.ҽ��id = e.Id And a.Id = b.��¼id And b.�շ�id = c.Id And c.ҩƷid = d.ҩƷid And a.Id = ��ҩid_In) Loop
        If Instr(r_��ҩ��¼.��ҩ����, r_���ȼ�.��ҩ����, 1) > 0 And (Instr(r_���ȼ�.Ƶ��, r_��ҩ��¼.ִ��Ƶ��, 1) > 0 Or r_���ȼ�.Ƶ�� = '����Ƶ��') Then
          n_First := r_���ȼ�.���ȼ�;
          Exit;
        End If;
      End Loop;
    End Loop;
  
    If n_First = 0 Then
      n_First := 999;
    End If;
    Return(n_First);
  End;
Begin
  n_Count          := 0;
  v_ҽ������       := Zl_To_Number(Nvl(zl_GetSysParameter('ҽ������', 1345), 1));
  v_��Һ����       := Zl_To_Number(Nvl(zl_GetSysParameter('ͬ������Һ����', 1345), 0));
  v_����Һ����     := Nvl(zl_GetSysParameter('����ҺҩƷ����', 1345), '');
  v_����Һ��ҩ;�� := Nvl(zl_GetSysParameter('��Һ��ҩ;��', 1345), '');
  v_��Դ����       := Nvl(zl_GetSysParameter('��Դ����', 1345), '');
  v_�����ϴ�����   := Zl_To_Number(Nvl(zl_GetSysParameter('�����ϴ�����', 1345), 0));
  n_Tpn���÷�ʽ    := Zl_To_Number(Nvl(zl_GetSysParameter('����Ӫ��ҩ�ﴦ�÷�ʽ', 1345), 0));
  n_���ҩƷ����   := Zl_To_Number(Nvl(zl_GetSysParameter('����ҩƷ����������ҩƷ�����ݸ�ҩʱ��û����ҩ���ε���Һ��Ĭ��Ϊ0���β����', 1345), 0));
  n_����ҩƷ����   := Zl_To_Number(Nvl(zl_GetSysParameter('����ҩƷ��ҩƷ����ָ������', 1345), 0));
  n_�Զ�����       := Zl_To_Number(Nvl(zl_GetSysParameter('�����Զ�����', 1345), 0));
  n_��������       := Zl_To_Number(Nvl(zl_GetSysParameter('���췢�͵�ҽ����������Һ��ȫ������������', 1345), 0));
  v_ҽ��ids        := ҽ��id_In;
  v_��ǰ����       := '';
  n_���ʹ���       := 0;

  --ȡ��ͨҵ�񾫶�λ��
  --���:1-ҩƷ 2-����
  --���ݣ�2-���ۼ� 4-���
  --��λ��ҩƷ:1-�ۼ� 5-��λ
  Select ���� Into n_��ͨ���С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 4 And ��λ = 5;

  Select Trunc(Sysdate) Into v_Currdate From Dual;

  Select Max(Nvl(����, 0)) + 1 Into v_Maxbatch From ��ҩ��������;

  --��鵱ǰ���˵�ҽ���Ƿ��н�����Ҫִ�е���Һ��������״̬��
  If Instr(v_ҽ��ids, ',') = 0 Then
    v_Tansid := v_ҽ��ids;
  Else
    v_Tansid := Substr(v_Tmp, 1, Instr(v_ҽ��ids, ',') - 1);
  End If;

  Select Count(ID)
  Into n_Num
  From ��Һ��ҩ��¼
  Where �Ƿ����� = 1 And ִ��ʱ�� Between Trunc(v_Currdate) And Trunc(v_Currdate + 1) - 1 / 24 / 60 / 60 And
        ҽ��id In
        (Select ���id
         From ����ҽ����¼
         Where ����id = (Select ����id From ����ҽ����¼ Where ���id = v_Tansid And Rownum < 2) And (������� = '5' Or ������� = '6')) And
        Rownum < 2;

  If n_Num > 0 Then
    Select ����
    Into v_��ǰ����
    From ��Һ��ҩ��¼
    Where �Ƿ����� = 1 And ִ��ʱ�� Between Trunc(v_Currdate) And Trunc(v_Currdate + 1) - 1 / 24 / 60 / 60 And
          ҽ��id In
          (Select ���id
           From ����ҽ����¼
           Where ����id = (Select ����id From ����ҽ����¼ Where ���id = v_Tansid And Rownum < 2) And (������� = '5' Or ������� = '6')) And
          Rownum < 2;
    Raise Err_Item;
  End If;

  --�Ƚ�ԭ�շ���¼����������µ��շ���¼��������ɾ��
  --Update ҩƷ�շ���¼
  --Set ��� = ��� + 10000
  --Where ID In (Select \*+rule *\
  --             Distinct c.Id
  --             From ����ҽ����¼ A, ����ҽ������ B, ҩƷ�շ���¼ C, סԺ���ü�¼ D, Table(f_Num2list(ҽ��id_In)) F
  --             Where a.Id = b.ҽ��id And c.����id = d.Id And a.Id = d.ҽ����� And b.No = c.No And b.ִ�в���id + 0 = ����id_In And
  --                   c.���� = 9 And c.������� Is Null And a.���id = f.Column_Value And b.���ͺ� = ���ͺ�_In And c.��� < 10000);

  For v_ҽ����¼ In c_ҽ����¼ Loop
    v_Continue := 1;
    n_����id   := v_ҽ����¼.����id;
    n_����id   := v_ҽ����¼.���˿���id;
  
    Select Count(1)
    Into v_Continue
    From ����ҽ����¼ A, ��Һ������ҩƷ B, סԺ���ü�¼ C
    Where c.�շ�ϸĿid = b.ҩƷid And c.ҽ����� = a.Id And a.���id = v_ҽ����¼.���id;
    If v_Continue = 0 Then
      v_Continue := 1;
    Else
      v_Continue := 0;
    End If;
  
    --�������Ʋ�����Һ��
    If (v_ҽ������ = 1 And v_ҽ����¼.ҽ������ <> 1) Or (v_ҽ������ = 2 And v_ҽ����¼.ҽ������ <> 2) Then
      v_Continue := 0;
    End If;
  
    If Not v_����Һ��ҩ;�� Is Null Then
      If Instr(',' || v_����Һ��ҩ;�� || ',', ',' || v_ҽ����¼.��ҩ;�� || ',') = 0 Then
        v_Continue := 0;
      End If;
    End If;
  
    If Not v_��Դ���� Is Null Then
      If Instr(',' || v_��Դ���� || ',', ',' || v_ҽ����¼.���˿���id || ',') = 0 Then
        v_Continue := 0;
      End If;
    End If;
  
    v_ҩƷ���� := Null;
    For r_ҩƷ���� In (Select Decode(Nvl(d.������, 0), 0, Decode(Nvl(d.�Ƿ�����ҩ, 0), 0, '', '����ҩ'), '������') ҩƷ����
                   From ����ҽ����¼ A, ҩƷ��� B, סԺ���ü�¼ C, ҩƷ���� D
                   Where c.�շ�ϸĿid = b.ҩƷid And b.ҩ��id = d.ҩ��id And c.ҽ����� = a.Id And a.���id = v_ҽ����¼.���id) Loop
      If r_ҩƷ����.ҩƷ���� Is Not Null Then
        v_ҩƷ���� := r_ҩƷ����.ҩƷ����;
      End If;
    End Loop;
  
    If v_ҩƷ���� Is Null Then
      If v_ҽ����¼.�Ƿ�tpn = 2 Then
        v_ҩƷ���� := 'Ӫ��ҩ';
      End If;
    End If;
  
    If v_Continue = 1 Then
      v_Old���id := v_New���id;
      v_���id    := v_ҽ����¼.���id;
      v_New���id := v_���id;
      v_���ͺ�    := v_ҽ����¼.���ͺ�;
      v_���      := 0;
    
      If v_Continue = 1 Then
        --v_Count := Zl_Gettransexenumber(v_ҽ����¼.��ʼִ��ʱ��, v_ҽ����¼.�״�ʱ��, v_ҽ����¼.ĩ��ʱ��, v_ҽ����¼.Ƶ�ʼ��, v_ҽ����¼.�����λ, v_ҽ����¼.ִ��ʱ�䷽��);
        Select Count(ҽ��id)
        Into v_Count
        From ҽ��ִ��ʱ��
        Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ�;
      
        v_Nodosage := 0;
      
        For I In 1 .. v_Count Loop
          Select ��Һ��ҩ��¼_Id.Nextval Into v_��ҩid From Dual;
          v_��� := v_��� + 1;
        
          If I > 1 Then
            --��ҽ��ִ��ʱ�����ȡҽ����ִ��ʱ��
            Select Ҫ��ʱ��
            Into v_ִ��ʱ��
            From ҽ��ִ��ʱ��
            Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ� And Ҫ��ʱ�� > v_ִ��ʱ�� And Rownum = 1
            Order By Ҫ��ʱ��;
          Else
            Select Min(Ҫ��ʱ��)
            Into v_ִ��ʱ��
            From ҽ��ִ��ʱ��
            Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ� And Rownum = 1
            Order By Ҫ��ʱ��;
          End If;
        
          v_���� := 0;
        
          If d_Oldִ��ʱ�� <> Trunc(v_ִ��ʱ��) Or d_Oldִ��ʱ�� Is Null Then
            b_Change := True;
          End If;
        
          If b_Change = True Then
            If d_Oldִ��ʱ�� <> Trunc(v_ִ��ʱ��) Or d_Oldִ��ʱ�� Is Null Then
              d_Oldִ��ʱ�� := v_ִ��ʱ��;
            
              Select Count(Distinct a.��ҩ����)
              Into n_��ҩ��
              From ��Һ��ҩ��¼ A
              Where a.ҽ��id In (Select ID From ����ҽ����¼ Where ����id = v_ҽ����¼.����id And ���id Is Null) And
                    a.ִ��ʱ�� Between Trunc(v_ִ��ʱ�� - 1) And Trunc(v_ִ��ʱ��) - 1 / 24 / 60 / 60 And ����״̬ >= 2 And ����״̬ < 9;
            
              If n_��ҩ�� > 1 Then
                Update ��Һ��ҩ��¼
                Set �Ƿ�������� = 1
                Where ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id) And
                     
                      ִ��ʱ�� Between Trunc(v_ִ��ʱ��) And Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
                b_Change := False;
              
              End If;
            End If;
          End If;
        
          If b_Change = True Then
            n_����id := v_ҽ����¼.����id;
            Select Count(ID)
            
            Into n_Sum
            From ��Һ��ҩ��¼
            Where ҽ��id = v_ҽ����¼.���id And ִ��ʱ�� Between Trunc(v_ִ��ʱ�� - 1) And Trunc(v_ִ��ʱ��) - 1 / 24 / 60 / 60;
            If n_Sum = 0 Then
              Update ��Һ��ҩ��¼
              Set �Ƿ�������� = 1
              Where ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id) And
                   
                    ִ��ʱ�� Between Trunc(v_ִ��ʱ��) And Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
              b_Change := False;
            
            End If;
          
            If b_Change = True Then
              --�����Һ���Ƿ���������״̬
              Select Count(a.Id)
              Into n_Sum
              From ��Һ��ҩ��¼ A, ��Һ��ҩ���� B, ҩƷ�շ���¼ C
              Where a.Id = b.��¼id And b.�շ�id = c.Id And
                    a.ҽ��id In (Select ID
                               From ����ҽ����¼
                               Where ����id = (Select ����id From ����ҽ����¼ Where ID = v_ҽ����¼.���id And Rownum < 2)) And
                    a.ִ��ʱ�� Between Trunc(v_ִ��ʱ�� - 1) And Trunc(v_ִ��ʱ��) - 1 / 24 / 60 / 60 And a.���ʱ�� Is Not Null;
              If n_Sum <> 0 Then
                Update ��Һ��ҩ��¼
                Set �Ƿ�������� = 1
                Where ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id) And ִ��ʱ�� Between Trunc(v_ִ��ʱ��) And
                      Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
                b_Change := False;
              End If;
            
              Select Count(ҽ��id)
              Into n_Cur
              From ҽ��ִ��ʱ��
              Where ҽ��id = v_ҽ����¼.���id And Ҫ��ʱ�� Between Trunc(v_ִ��ʱ��) And Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60;
            
              Select Count(ҽ��id)
              Into n_Sum
              From ҽ��ִ��ʱ��
              Where ҽ��id = v_ҽ����¼.���id And Ҫ��ʱ�� Between Trunc(v_ִ��ʱ�� - 1) And Trunc(v_ִ��ʱ��) - 1 / 24 / 60 / 60;
            
              If n_Sum <> n_Cur Then
                Update ��Һ��ҩ��¼
                Set �Ƿ�������� = 1
                Where ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id) And ִ��ʱ�� Between Trunc(v_ִ��ʱ��) And
                      Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
                b_Change := False;
              End If;
            End If;
          End If;
        
          If v_ʱ�䴮 <> Trunc(Sysdate) || ';false\' Or v_ʱ�䴮 Is Null Then
            If Trunc(v_ִ��ʱ��) = Trunc(Sysdate) Then
              If b_Change = False Then
                v_ʱ�䴮 := Trunc(v_ִ��ʱ��) || ';false\';
              Else
                v_ʱ�䴮 := Trunc(v_ִ��ʱ��) || ';true\';
              End If;
            End If;
          End If;
        
          If v_ʱ�䴮1 <> Trunc(Sysdate + 1) || ';false\' Or v_ʱ�䴮1 Is Null Then
            If Trunc(v_ִ��ʱ��) = Trunc(Sysdate + 1) Then
              If b_Change = False Then
                v_ʱ�䴮1 := Trunc(v_ִ��ʱ��) || ';false\';
              Else
                v_ʱ�䴮1 := Trunc(v_ִ��ʱ��) || ';true\';
              End If;
            End If;
          End If;
        
          If v_ҩƷ���� Is Null Or n_����ҩƷ���� = 0 Then
            v_���� := Zl_Getpivaworkbatch(v_ִ��ʱ��, Sysdate);
          Else
            --ҩƷ���Ͳ�Ϊ�գ�ֱ�Ӹ���ҩƷ����ƥ������
            v_���� := Zl_Getpivaworkbatch(v_ִ��ʱ��, Sysdate, v_ҩƷ����);
          End If;
        
          Select Count(ҽ��id)
          Into n_���ʹ���
          From ҽ��ִ��ʱ��
          Where ҽ��id = v_ҽ����¼.���id And Ҫ��ʱ�� <= v_ִ��ʱ��
          Order By Ҫ��ʱ��;
        
          If n_���ʹ��� > 99 Then
            n_���ʹ��� := Mod(n_���ʹ���, 99);
          End If;
        
          If Length(v_ҽ����¼.���id) > 9 Then
            If n_���ʹ��� < 10 Then
              Select '91' || Substr(To_Char(v_ҽ����¼.���id), Length(v_ҽ����¼.���id) - 8) || To_Char(v_ҽ����¼.���id) || '0' ||
                      To_Char(n_���ʹ���)
              Into v_Maxno
              From Dual;
            Else
              Select '91' || Substr(To_Char(v_ҽ����¼.���id), Length(v_ҽ����¼.���id) - 8) || To_Char(v_ҽ����¼.���id) ||
                      To_Char(n_���ʹ���)
              Into v_Maxno
              From Dual;
            End If;
          Else
            If n_���ʹ��� < 10 Then
              Select '91' || Substr('000000000', Length(v_ҽ����¼.���id) + 1) || To_Char(v_ҽ����¼.���id) || '0' ||
                      To_Char(n_���ʹ���)
              Into v_Maxno
              From Dual;
            Else
              Select '91' || Substr('000000000', Length(v_ҽ����¼.���id) + 1) || To_Char(v_ҽ����¼.���id) || To_Char(n_���ʹ���)
              Into v_Maxno
              From Dual;
            End If;
          End If;
          n_�������� := 0;
          If b_Change = False Then
            n_�������� := 1;
          End If;
        
          If v_���� <> 0 Then
            Select Nvl(Max(���), 0), Max(ҩƷ����)
            Into n_���, v_��ҩ����
            From ��ҩ��������
            Where ���� = v_���� And ��������id = ����id_In;
          End If;
        
          If (Trunc(v_ִ��ʱ��) <= v_Currdate Or n_��� <> 0) And v_��ҩ���� Is Null Then
            n_�Ƿ���     := 1;
            d_�ֹ����ʱ�� := Null;
          Else
            n_�Ƿ���     := 0;
            d_�ֹ����ʱ�� := Null;
          End If;
        
          --�����TPN��������������ζ�����Ϊ����
          If v_ҽ����¼.�Ƿ�tpn = 2 Then
            n_�Ƿ���     := 0;
            d_�ֹ����ʱ�� := Null;
          End If;
        
          If v_���� = 0 Then
            n_�Ƿ��� := 1;
          End If;
          --������ҩ��¼
          Insert Into ��Һ��ҩ��¼
            (ID, ����id, ���, ����, �Ա�, ����, סԺ��, ����, ���˲���id, ���˿���id, ִ��ʱ��, ҽ��id, ���ͺ�, ��ҩ����, ƿǩ��, �Ƿ��������, �Ƿ���, ���ʱ��, ����״̬,
             ������Ա, ����ʱ��)
          Values
            (v_��ҩid, ����id_In, v_���, v_ҽ����¼.����, v_ҽ����¼.�Ա�, v_ҽ����¼.����, v_ҽ����¼.סԺ��, v_ҽ����¼.����, v_ҽ����¼.���˲���id,
             v_ҽ����¼.���˿���id, v_ִ��ʱ��, v_ҽ����¼.���id, v_ҽ����¼.���ͺ�, v_����, v_Maxno, n_��������, n_�Ƿ���, d_�ֹ����ʱ��, 1, �˲���_In, �˲�ʱ��_In);
        
          Insert Into ��Һ��ҩ״̬ (��ҩid, ��������, ������Ա, ����ʱ��) Values (v_��ҩid, 1, �˲���_In, �˲�ʱ��_In);
        
          For v_����ҽ����¼ In c_����ҽ����¼ Loop
            n_ҽ��id   := v_����ҽ����¼.ҽ��id;
            n_�ۼ����� := 0;
            n_ʣ������ := 0;
          
            Select Sum(c.ʵ������)
            Into n_Sum
            From ����ҽ����¼ A, ����ҽ������ B, ҩƷ�շ���¼ C, סԺ���ü�¼ D
            Where a.Id = b.ҽ��id And c.����id = d.Id And a.Id = d.ҽ����� And b.No = c.No And b.ִ�в���id + 0 = ����id_In And
                  c.���� = 9 And c.������� Is Null And a.Id = n_ҽ��id And b.���ͺ� = v_ҽ����¼.���ͺ� And c.��� < 1000;
          
            --������ҩ��¼��Ӧ��ҩƷ��¼
            For v_�շ���¼ In c_�շ���¼ Loop
              If v_�շ���¼.�Ƿ������� = 1 Then
                v_Nodosage := 1;
              End If;
            
              Select ҩƷ�շ���¼_Id.Nextval Into n_Lngid From Dual;
              n_�ۼ����� := n_�ۼ����� + v_�շ���¼.����;
            
              If n_ʣ������ = 0 Then
                n_ʣ������ := n_Sum / v_Count;
              End If;
              n_�������� := n_Sum / v_Count;
            
              If n_�ۼ����� >= n_Sum / v_Count * I Then
                n_Count := n_Count + 1;
                Insert Into ҩƷ�շ���¼
                  (ID, ��¼״̬, ����, NO, ���, �ⷿid, ��ҩ��λid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����, ��������, Ч��, ����, ��д����, ʵ������,
                   �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��, ���, ժҪ, ������, ��������, ��ҩ��, ��ҩ����, �����, �������, �۸�id, ����id, ����, Ƶ��, �÷�, ���, �������,
                   ���Ч��, ��Ʒ�ϸ�֤, ��ҩ��ʽ, ��ҩ����, ������, ��׼�ĺ�, ���ܷ�ҩ��, ע��֤��, �ⷿ��λ, ��Ʒ����, �ڲ�����, �˲���, �˲�����, ǩ��ȷ����, ǩ��ʱ��)
                  Select n_Lngid, ��¼״̬, ����, NO, n_Count + 1000, �ⷿid, ��ҩ��λid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����,
                         ��������, Ч��, ����, n_ʣ������, n_ʣ������, �ɱ���, Round(�ɱ��� * n_ʣ������, n_��ͨ���С��), ����, ���ۼ�,
                         Round(���ۼ� * n_ʣ������, n_��ͨ���С��), Round(��� * (ʵ������ / n_ʣ������), n_��ͨ���С��), '����', ������, ��������, ��ҩ��,
                         ��ҩ����, �����, �������, �۸�id, ����id, ����, Ƶ��, �÷�, ���, �������, ���Ч��, ��Ʒ�ϸ�֤, ��ҩ��ʽ, ��ҩ����, ������, ��׼�ĺ�, ���ܷ�ҩ��,
                         ע��֤��, �ⷿ��λ, ��Ʒ����, �ڲ�����, �˲���, �˲�����, ǩ��ȷ����, ǩ��ʱ��
                  From ҩƷ�շ���¼
                  Where ID = v_�շ���¼.�շ�id;
              
                Zl_δ��ҩƷ��¼_Insert(n_Lngid);
              
                Insert Into ��Һ��ҩ���� (��¼id, �շ�id, ����) Values (v_��ҩid, n_Lngid, n_ʣ������);
              
                n_ʣ������ := 0;
                Exit;
              Elsif n_�ۼ����� > (n_Sum / v_Count * (I - 1)) Then
                n_Count    := n_Count + 1;
                n_��д���� := n_�ۼ����� - (n_Sum / v_Count * (I - 1)) - (n_�������� - n_ʣ������);
                Insert Into ҩƷ�շ���¼
                  (ID, ��¼״̬, ����, NO, ���, �ⷿid, ��ҩ��λid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����, ��������, Ч��, ����, ��д����, ʵ������,
                   �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��, ���, ժҪ, ������, ��������, ��ҩ��, ��ҩ����, �����, �������, �۸�id, ����id, ����, Ƶ��, �÷�, ���, �������,
                   ���Ч��, ��Ʒ�ϸ�֤, ��ҩ��ʽ, ��ҩ����, ������, ��׼�ĺ�, ���ܷ�ҩ��, ע��֤��, �ⷿ��λ, ��Ʒ����, �ڲ�����, �˲���, �˲�����, ǩ��ȷ����, ǩ��ʱ��)
                  Select n_Lngid, ��¼״̬, ����, NO, n_Count + 1000, �ⷿid, ��ҩ��λid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����,
                         ��������, Ч��, ����, n_��д����, n_��д����, �ɱ���, Round(�ɱ��� * n_��д����, n_��ͨ���С��), ����, ���ۼ�,
                         Round(���ۼ� * n_��д����, n_��ͨ���С��), Round(��� * (ʵ������ / n_��д����), n_��ͨ���С��), '����', ������, ��������, ��ҩ��,
                         ��ҩ����, �����, �������, �۸�id, ����id, ����, Ƶ��, �÷�, ���, �������, ���Ч��, ��Ʒ�ϸ�֤, ��ҩ��ʽ, ��ҩ����, ������, ��׼�ĺ�, ���ܷ�ҩ��,
                         ע��֤��, �ⷿ��λ, ��Ʒ����, �ڲ�����, �˲���, �˲�����, ǩ��ȷ����, ǩ��ʱ��
                  From ҩƷ�շ���¼
                  Where ID = v_�շ���¼.�շ�id;
              
                Zl_δ��ҩƷ��¼_Insert(n_Lngid);
              
                Insert Into ��Һ��ҩ���� (��¼id, �շ�id, ����) Values (v_��ҩid, n_Lngid, n_��д����);
              
                n_ʣ������ := n_ʣ������ - n_��д����;
              End If;
            End Loop;
          End Loop;
          n_���ȼ� := Zl_Getfirst(v_��ҩid, v_ҽ����¼.���˿���id);
          Update ��Һ��ҩ��¼ Set ���ȼ� = n_���ȼ� Where ID = v_��ҩid;
        
        End Loop;
      
        For v_�շ���¼ In c_ԭʼ�շ���¼ Loop
          n_���� := v_�շ���¼.����;
        
          v_No := v_�շ���¼.No;
          Delete From ҩƷ�շ���¼ Where ID = v_�շ���¼.�շ�id;
        End Loop;
      
        --����ҩƷ���߲������õ�ҩƷĬ��Ϊ0����
        Select Count(�շ�id) Into n_Row From ��Һ��ҩ���� Where ��¼id = v_��ҩid;
        If (v_Nodosage = 1 Or n_Row = 1) And n_���ҩƷ���� = 1 Then
          Update ��Һ��ҩ��¼
          Set ��ҩ���� = 0, �Ƿ��� = 1
          Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ� And ����״̬ < 2;
        End If;
        --������ڡ��������á����Ե�ҩƷ��Ҳ����Ϊ���
        If v_Nodosage = 1 Then
          Update ��Һ��ҩ��¼
          Set �Ƿ��� = 1
          Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ� And ����״̬ < 2;
        End If;
      End If;
    End If;
  End Loop;

  For v_�շ���¼ In (Select ID From ҩƷ�շ���¼ Where ��� < 1000 And ���� = n_���� And NO = v_No) Loop
    n_Count := n_Count + 1;
    Update ҩƷ�շ���¼ Set ��� = n_Count + 1000, ժҪ = '����' Where ID = v_�շ���¼.Id;
  End Loop;

  Update ҩƷ�շ���¼
  Set ��� = ��� - 1000, ժҪ = 'ҽ������'
  Where ժҪ = '����' And ��� > 1000 And ���� = n_���� And NO = v_No;

  If n_�������� = 1 Then
  
    Select Count(a.Id)
    Into n_Sum
    From ��Һ��ҩ��¼ A, ����ҽ������ B
    Where a.ҽ��id = b.ҽ��id And a.���ͺ� = b.���ͺ� And
          a.ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id And ���id Is Null) And b.����ʱ�� Between Trunc(Sysdate) And
          Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And a.ִ��ʱ�� Between Trunc(Sysdate) And
          Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And ����״̬ < 9;
    If n_Sum <> 0 Then
      b_Change  := False;
      v_ʱ�䴮1 := Trunc(Sysdate + 1) || ';false\';
    
      Update ��Һ��ҩ��¼
      Set �Ƿ�������� = 1
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id) And ִ��ʱ�� Between Trunc(Sysdate + 1) And
            Trunc(Sysdate + 2) - 1 / 24 / 60 / 60 And ����״̬ < 2;
    End If;
  End If;
  If v_ʱ�䴮 Is Null Then
    v_ʱ�䴮 := v_ʱ�䴮1;
  Else
    v_ʱ�䴮 := v_ʱ�䴮 || v_ʱ�䴮1;
  End If;

  While v_ʱ�䴮 Is Not Null Loop
    --�ֽⵥ��ID��
    v_Fields   := Substr(v_ʱ�䴮, 1, Instr(v_ʱ�䴮, '\') - 1);
    v_ʱ��ֵ   := Substr(v_Fields, 1, Instr(v_Fields, ';') - 1);
    v_�Ƿ�ı� := Substr(v_Fields, Instr(v_Fields, ';') + 1);
  
    v_ʱ�䴮 := Replace('\' || v_ʱ�䴮, '\' || v_Fields || '\');
  
    If v_�Ƿ�ı� = 'true' Then
      b_Change := True;
    Else
      b_Change := False;
    End If;
  
    If b_Change = True Then
      Select Count(ҽ��id)
      Into n_Cur
      From (Select Distinct a.Ҫ��ʱ��, a.ҽ��id
             From ҽ��ִ��ʱ�� A, ��Һ��ҩ��¼ B
             Where a.Ҫ��ʱ�� = b.ִ��ʱ�� And a.ҽ��id = b.ҽ��id And a.Ҫ��ʱ�� Between Trunc(v_ʱ��ֵ) And
                   Trunc(v_ʱ��ֵ + 1) - 1 / 24 / 60 / 60 And
                   a.ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id And ���id Is Null));
    
      Select Count(ҽ��id)
      Into n_Sum
      From (Select Distinct a.Ҫ��ʱ��, a.ҽ��id
             From ҽ��ִ��ʱ�� A, ��Һ��ҩ��¼ B
             Where a.Ҫ��ʱ�� = b.ִ��ʱ�� And a.ҽ��id = b.ҽ��id And a.Ҫ��ʱ�� Between Trunc(v_ʱ��ֵ - 1) And
                   Trunc(v_ʱ��ֵ) - 1 / 24 / 60 / 60 And
                   a.ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id And ���id Is Null));
    
      If n_Cur <> n_Sum Then
        Update ��Һ��ҩ��¼
        Set �Ƿ�������� = 1
        Where ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id) And ִ��ʱ�� Between Trunc(v_ʱ��ֵ) And
              Trunc(v_ʱ��ֵ + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
        b_Change := False;
      End If;
    End If;
  
    If v_�����ϴ����� = 1 And b_Change = True Then
      For v_��Һ����¼ In c_��Һ����¼ Loop
        Begin
          Select Distinct ��ҩ����
          Into v_����
          From ��Һ��ҩ��¼ A, ��Һ��ҩ���� B, ҩƷ�շ���¼ C
          Where a.Id = b.��¼id And b.�շ�id = c.Id And a.ҽ��id = v_��Һ����¼.ҽ��id And
                To_Char(a.ִ��ʱ��, 'hh24:mi:ss') = To_Char(v_��Һ����¼.ִ��ʱ��, 'hh24:mi:ss') And
                a.ִ��ʱ�� Between Trunc(v_��Һ����¼.ִ��ʱ�� - 1) And Trunc(v_��Һ����¼.ִ��ʱ��) - 1 / 24 / 60 / 60 And Rownum = 1;
        Exception
          When Others Then
            Begin
              Select Distinct ��ҩ����
              Into v_����
              From ��Һ��ҩ��¼ A
              Where a.ҽ��id = v_��Һ����¼.ҽ��id And To_Char(a.ִ��ʱ��, 'hh24:mi:ss') = To_Char(v_��Һ����¼.ִ��ʱ��, 'hh24:mi:ss') And
                    a.����״̬ <> 12 And a.ִ��ʱ�� Between Trunc(v_��Һ����¼.ִ��ʱ�� - 1) And Trunc(v_��Һ����¼.ִ��ʱ��) - 1 / 24 / 60 / 60 And
                    Rownum = 1;
            Exception
              When Others Then
                v_���� := v_��Һ����¼.��ҩ����;
            End;
        End;
      
        Update ��Һ��ҩ��¼
        Set �Ƿ�ȷ�ϵ��� = 0, �Ƿ�������� = 0
        Where ҽ��id In (Select ID From ����ҽ����¼ Where ����id = n_����id) And ִ��ʱ�� Between Trunc(v_��Һ����¼.ִ��ʱ��) And
              Trunc(v_��Һ����¼.ִ��ʱ�� + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
      
        If v_��Һ����¼.��ҩ���� <> v_���� Then
          Update ��Һ��ҩ��¼ Set ��ҩ���� = v_���� Where ID = v_��Һ����¼.Id;
          Select Nvl(Max(���), 0) Into n_��� From ��ҩ�������� Where ���� = v_���� And ��������id = ����id_In;
          If n_��� <> 0 Then
            Update ��Һ��ҩ��¼ Set �Ƿ��� = n_��� Where ID = v_��Һ����¼.Id;
          Else
            Select Nvl(Max(���), 0)
            Into n_���
            From ��ҩ��������
            Where ���� = v_��Һ����¼.��ҩ���� And ��������id = ����id_In;
          
            If n_��� <> 0 Then
              Update ��Һ��ҩ��¼ Set �Ƿ��� = 0 Where ID = v_��Һ����¼.Id;
            End If;
          End If;
        End If;
      End Loop;
    End If;
  
    If n_�Զ����� = 1 And (b_Change = False Or v_�����ϴ����� = 0) Then
      For v_��Һ����¼ In c_��Һ����¼ Loop
        v_���� := Zl_Getpivaworkbatch(v_��Һ����¼.ִ��ʱ��, v_��Һ����¼.����ʱ��);
        Update ��Һ��ҩ��¼ Set ��ҩ���� = v_���� Where ID = v_��Һ����¼.Id;
      End Loop;
      Zl_��Һ��ҩ��¼_�Զ�����(n_����id, n_����id, ����id_In, v_ʱ��ֵ);
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]����' || v_��ǰ���� || '����Һ���������б���������Һ��������ʧ�ܣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_�˲�;
/

--125588:����,2018-05-17,������ʵ������Ϊ0������
Create Or Replace Procedure Zl_�����շ���¼_Adjust
(
  ����id_In   In Number, --���ۼ�¼��ID
  ����_In     In Number := 0, --�Ƿ�תΪ�������ۣ����²������ԡ��շ�ϸĿ�еı�ۣ�
  ����id_In   In Number := 0, --����Ϊ0ʱ��ʾ�ǳɱ��۵��ۣ��������ۼ��������
  Billinfo_In In Varchar2 := Null --����ʱ�����İ����ε��ۡ���ʽ:"����1,�ּ�1|����2,�ּ�2|....."
) As
  n_������id ҩƷ�շ���¼.������id%Type; --������
  v_���۵��ݺ� ҩƷ�շ���¼.No%Type; --���۵���
  d_��Ч����   Date; --������Чʱ��
  n_ִ�е���   Number(1); --����ʱ�̵���
  n_ʵ�۲���   Number(1); --ʱ��ҩƷ
  n_�շ�ϸĿid Number(18); --�շ�ϸĿID
  d_�������   ҩƷ�շ���¼.�������%Type;
  n_���۽��   ҩƷ���.ʵ�ʽ��%Type;
  n_���ۼ�     ҩƷ���.���ۼ�%Type;
  n_���       Integer(8);
  v_Infotmp    Varchar2(4000);
  v_Fields     Varchar2(4000);
  n_����       Number(18);
  n_�ּ�       �շѼ�Ŀ.�ּ�%Type;
  n_ԭ��       �շѼ�Ŀ.ԭ��%Type;
  n_�շ�id     ҩƷ�շ���¼.Id%Type;
  n_ʱ�۷���   Number(1);
  v_Lngid      ҩƷ�շ���¼.Id%Type; --�շ�ID
  n_�۸�id     �շѼ�Ŀ.Id%Type;

  Cursor c_Price --��ͨ����
  Is
    Select 1 ��¼״̬, 13 ����, v_���۵��ݺ� NO, Rownum ���, n_������id ������id, m.����id ҩƷid, s.���� ����, Null ����, s.Ч��,
           Decode(s.�ϴβ���, Null, q.����, s.�ϴβ���) ����, 1 ����, s.ʵ������ ��д����, 0 ʵ������, a.ԭ�� �ɱ���, 0 �ɱ����, a.�ּ� ���ۼ�, 0 ����,
           Nvl(s.���ۼ�, 0) As ������ۼ�, s.ʵ�ʽ�� As �����, s.ʵ�ʲ�� As �����, '���ĵ���' ժҪ, User ������, Sysdate ��������, s.�ⷿid �ⷿid,
           1 ���ϵ��, a.Id �۸�id, s.�ϴ���������, s.���Ч��, s.��׼�ĺ�, s.�ϴι�Ӧ��id,
           Nvl(s.���ۼ�, Decode(Nvl(s.ʵ������, 0), 0, a.ԭ��, Nvl(s.ʵ�ʽ��, 0) / s.ʵ������)) As ԭ�ۼ�
    From ҩƷ��� S, �������� M, �շѼ�Ŀ A, �շ���ĿĿ¼ Q
    Where s.ҩƷid = m.����id And m.����id = q.Id And m.����id = a.�շ�ϸĿid And s.���� = 1 And a.�䶯ԭ�� = 0 And a.Id = ����id_In And
          a.ִ������ <= Sysdate;

  Cursor c_ʱ�۰����ε��� --ʱ�����İ����ε���
  Is
    Select 1 ��¼״̬, 13 ����, v_���۵��ݺ� NO, n_��� + Rownum ���, n_������id ������id, s.ҩƷid ҩƷid, s.���� ����, Null ����, s.Ч��,
           Decode(s.�ϴβ���, Null, b.����, s.�ϴβ���) ����, 1 ����, Nvl(s.ʵ������, 0) ��д����, 0 ʵ������, a.ԭ�� �ɱ���, 0 �ɱ����, n_�ּ� ���ۼ�, 0 ����,
           '���ĵ���' ժҪ, User ������, Sysdate ��������, s.�ⷿid �ⷿid, 1 ���ϵ��, a.Id �۸�id, Nvl(b.�Ƿ���, 0) As ʱ��, s.ʵ�ʽ�� As �����,
           s.ʵ�ʲ�� As �����, Nvl(s.���ۼ�, Decode(Nvl(s.ʵ������, 0), 0, a.ԭ��, Nvl(s.ʵ�ʽ��, 0) / s.ʵ������)) As ԭ�ۼ�
    From ҩƷ��� S, �������� M, �շѼ�Ŀ A, �շ���ĿĿ¼ B
    Where s.ҩƷid = m.����id And m.����id = a.�շ�ϸĿid And a.�շ�ϸĿid = b.Id And s.���� = 1 And a.�䶯ԭ�� = 0 And a.Id = ����id_In And
          a.ִ������ <= Sysdate And Nvl(s.����, 0) = n_����;
Begin

  If ����id_In <> 0 Then
    --�ɱ��۵���
    Zl_�����շ���¼_�ɱ��۵���(����id_In);
    Return;
  End If;

  --ȡ������ID
  Select ���id Into n_������id From ҩƷ�������� Where ���� = 13;

  --ȡ����
  Select Nextno(147) Into v_���۵��ݺ� From Dual;
  --ȡ���ۼ�¼��Ч����
  Select �շ�ϸĿid, ִ������ Into n_�շ�ϸĿid, d_��Ч���� From �շѼ�Ŀ Where ID = ����id_In;
  --ȡ�ò����Ƿ���ʱ��ҩƷ
  Select Nvl(�Ƿ���, 0) Into n_ʵ�۲��� From �շ���ĿĿ¼ Where ID = n_�շ�ϸĿid;

  If Sysdate >= d_��Ч���� Then
    n_ִ�е��� := 1;
  Else
    n_ִ�е��� := 0;
  End If;

  If n_ִ�е��� = 1 Then
    d_������� := Sysdate;
    --��ͨ���۴���
    If Billinfo_In = '' Or Billinfo_In Is Null Then
      --��ʱ��ҩƷ����
      For c_���� In c_Price Loop
        n_�۸�id := c_����.�۸�id;
        /*If Nvl(c_����.��д����, 0) = 0 And Nvl(c_����.�����, 0) = 0 And Nvl(c_����.�����, 0) = 0 Then
          Null;
        Elsif Nvl(c_����.��д����, 0) = 0 And (Nvl(c_����.�����, 0) <> 0 Or Nvl(c_����.�����, 0) <> 0) Then
          --����=0 ������<>0ʱֻ���¿����ж�Ӧ�����ۼ�,�������ۼ��������ݵ��ǽ���=0��ֻ��¼�����ۼۣ�����Ͳ�۲������

        
        
          --��������Ӱ���¼
          Select ҩƷ�շ���¼_Id.Nextval Into v_Lngid From Dual;
          Insert Into ҩƷ�շ���¼
            (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ժҪ, ������, ��������,
             �ⷿid, ���ϵ��, �۸�id, �����, �������, ��������, ���Ч��, ��׼�ĺ�, ��ҩ��λid, ����, Ƶ��)
          Values
            (v_Lngid, c_����.��¼״̬, c_����.����, c_����.No, c_����.���, c_����.������id, c_����.ҩƷid, c_����.����, c_����.����, c_����.Ч��, c_����.����,
             c_����.����, c_����.��д����, c_����.ʵ������, Decode(n_ʵ�۲���, 1, c_����.ԭ�ۼ�, c_����.�ɱ���), c_����.�ɱ����, c_����.���ۼ�, c_����.����, c_����.ժҪ,
             c_����.������, c_����.��������, c_����.�ⷿid, c_����.���ϵ��, c_����.�۸�id, User, d_�������, c_����.�ϴ���������, c_����.���Ч��, c_����.��׼�ĺ�,
             c_����.�ϴι�Ӧ��id, c_����.�����, c_����.�����);
        
          Zl_δ��ҩƷ��¼_Insert(v_Lngid);
          --���²��Ͽ�� ��ֻ��ʱ�����ĲŸ������ۼ�
          Update ҩƷ���
          Set ���ۼ� = Decode(n_ʵ�۲���, 1, Decode(Nvl(c_����.����, 0), 0, Null, c_����.���ۼ�), Null)
          Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.ҩƷid And ���� = 1 And Nvl(����, 0) = Nvl(c_����.����, 0);
        
          Zl_δ��ҩƷ��¼_Delete(v_Lngid);
        Else*/
        If n_ʵ�۲��� = 1 Then
          If c_����.������ۼ� = 0 Then
            n_���ۼ� := c_����.ԭ�ۼ�;
          Else
            n_���ۼ� := c_����.������ۼ�;
          End If;
        Else
          n_���ۼ� := c_����.�ɱ���;
        End If;
        n_���۽�� := Round((c_����.���ۼ� - n_���ۼ�) * c_����.��д����, 2);
      
        --��������Ӱ���¼
        Select ҩƷ�շ���¼_Id.Nextval Into v_Lngid From Dual;
        Insert Into ҩƷ�շ���¼
          (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ���۽��, ���, ժҪ, ������,
           ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ��������, ���Ч��, ��׼�ĺ�, ��ҩ��λid, ����, Ƶ��)
        Values
          (v_Lngid, c_����.��¼״̬, c_����.����, c_����.No, c_����.���, c_����.������id, c_����.ҩƷid, c_����.����, c_����.����, c_����.Ч��, c_����.����,
           c_����.����, c_����.��д����, c_����.ʵ������, Decode(n_ʵ�۲���, 1, c_����.ԭ�ۼ�, c_����.�ɱ���), c_����.�ɱ����, c_����.���ۼ�, c_����.����, n_���۽��,
           n_���۽��, c_����.ժҪ, c_����.������, c_����.��������, c_����.�ⷿid, c_����.���ϵ��, c_����.�۸�id, User, d_�������, c_����.�ϴ���������, c_����.���Ч��,
           c_����.��׼�ĺ�, c_����.�ϴι�Ӧ��id, c_����.�����, c_����.�����);
      
        Zl_δ��ҩƷ��¼_Insert(v_Lngid);
        --���²��Ͽ��
        Update ҩƷ���
        Set ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���۽��,
            ���ۼ� = Decode(n_ʵ�۲���, 1, Decode(Nvl(c_����.����, 0), 0, Null, c_����.���ۼ�), Null)
        Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.ҩƷid And ���� = 1 And Nvl(����, 0) = Nvl(c_����.����, 0);
      
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ��׼�ĺ�, ���ۼ�)
          Values
            (c_����.�ⷿid, c_����.ҩƷid, c_����.����, 1, 0, 0, n_���۽��, n_���۽��, c_����.Ч��, c_����. ���Ч��, c_����.�ϴι�Ӧ��id, c_����.�ɱ���,
             c_����.����, c_����.�ϴ���������, c_����.����, c_����.��׼�ĺ�,
             Decode(n_ʵ�۲���, 1, Decode(Nvl(c_����.����, 0), 0, Null, c_����.���ۼ�), Null));
        End If;
      
        Zl_δ��ҩƷ��¼_Delete(v_Lngid);
        --End If;
      End Loop;
    
      --��Ϣ����
      b_Message.Zlhis_Drug_011(n_�۸�id, 0);
    Else
      --ʱ�۷������۴���
      n_��� := 0;
      --ʱ��ҩƷ�����ε���
      v_Infotmp := Billinfo_In || '|';
      While v_Infotmp Is Not Null Loop
        --�ֽⵥ��ID��
        v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
        n_����    := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
        n_�ּ�    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
        v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
      
        For v_ʱ�۰����ε��� In c_ʱ�۰����ε��� Loop
          If v_ʱ�۰����ε���.��д���� <> 0 Then
            n_ԭ�� := Nvl(v_ʱ�۰����ε���.�����, 0) / v_ʱ�۰����ε���.��д����;
          Else
            n_ԭ�� := v_ʱ�۰����ε���.�ɱ���;
          End If;
        
          Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
          /*If Nvl(v_ʱ�۰����ε���.��д����, 0) = 0 And Nvl(v_ʱ�۰����ε���.�����, 0) = 0 And Nvl(v_ʱ�۰����ε���.�����, 0) = 0 Then
            Null;
            n_�۸�id := Null;
          Elsif Nvl(v_ʱ�۰����ε���.��д����, 0) = 0 And (Nvl(v_ʱ�۰����ε���.�����, 0) <> 0 Or Nvl(v_ʱ�۰����ε���.�����, 0) <> 0) Then
            --����=0 ������<>0ʱֻ���¿����ж�Ӧ�����ۼ�,�������ۼ��������ݵ��ǽ���=0��ֻ��¼�����ۼۣ�����Ͳ�۲������

          
          
            --��������Ӱ���¼
            Insert Into ҩƷ�շ���¼
              (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ժҪ, ������, ��������,
               �ⷿid, ���ϵ��, �۸�id, �����, �������, ����, Ƶ��)
            Values
              (n_�շ�id, v_ʱ�۰����ε���.��¼״̬, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.No, v_ʱ�۰����ε���.���, v_ʱ�۰����ε���.������id, v_ʱ�۰����ε���.ҩƷid,
               v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.Ч��, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.��д����, v_ʱ�۰����ε���.ʵ������,
               Decode(n_ʵ�۲���, 1, v_ʱ�۰����ε���.ԭ�ۼ�, v_ʱ�۰����ε���.�ɱ���), v_ʱ�۰����ε���.�ɱ����, v_ʱ�۰����ε���.���ۼ�, v_ʱ�۰����ε���.����,
               v_ʱ�۰����ε���.ժҪ, v_ʱ�۰����ε���.������, v_ʱ�۰����ε���.��������, v_ʱ�۰����ε���.�ⷿid, v_ʱ�۰����ε���.���ϵ��, v_ʱ�۰����ε���.�۸�id, User, d_�������,
               v_ʱ�۰����ε���.�����, v_ʱ�۰����ε���.�����);
            n_��� := n_��� + 1;
          
            Zl_δ��ҩƷ��¼_Insert(n_�շ�id);
            --������
            --���¿�����ۼ�,ֻ��ʱ�۷���ҩƷ���ܸ������ۼ��ֶ�
            Update ҩƷ���
            Set ���ۼ� = Decode(v_ʱ�۰����ε���.ʱ��, 1, Decode(Nvl(v_ʱ�۰����ε���.����, 0), 0, Null, v_ʱ�۰����ε���.���ۼ�), Null)
            Where �ⷿid = v_ʱ�۰����ε���.�ⷿid And ҩƷid = v_ʱ�۰����ε���.ҩƷid And ���� = 1 And Nvl(����, 0) = Nvl(v_ʱ�۰����ε���.����, 0);
          
            Zl_δ��ҩƷ��¼_Delete(n_�շ�id);
          
            n_�۸�id := n_�շ�id;
          Else*/
          n_���ۼ�   := v_ʱ�۰����ε���.ԭ�ۼ�;
          n_���۽�� := Round((n_�ּ� - n_���ۼ�) * v_ʱ�۰����ε���.��д����, 2);
          --��������Ӱ���¼
          Insert Into ҩƷ�շ���¼
            (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ���۽��, ���, ժҪ, ������,
             ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ����, Ƶ��)
          Values
            (n_�շ�id, v_ʱ�۰����ε���.��¼״̬, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.No, v_ʱ�۰����ε���.���, v_ʱ�۰����ε���.������id, v_ʱ�۰����ε���.ҩƷid,
             v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.Ч��, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.��д����, v_ʱ�۰����ε���.ʵ������,
             Decode(n_ʵ�۲���, 1, v_ʱ�۰����ε���.ԭ�ۼ�, v_ʱ�۰����ε���.�ɱ���), v_ʱ�۰����ε���.�ɱ����, v_ʱ�۰����ε���.���ۼ�, v_ʱ�۰����ε���.����, n_���۽��,
             n_���۽��, v_ʱ�۰����ε���.ժҪ, v_ʱ�۰����ε���.������, v_ʱ�۰����ε���.��������, v_ʱ�۰����ε���.�ⷿid, v_ʱ�۰����ε���.���ϵ��, v_ʱ�۰����ε���.�۸�id, User,
             d_�������, v_ʱ�۰����ε���.�����, v_ʱ�۰����ε���.�����);
          n_��� := n_��� + 1;
        
          Zl_δ��ҩƷ��¼_Insert(n_�շ�id);
          --������
          If v_ʱ�۰����ε���.ʱ�� = 1 And Nvl(v_ʱ�۰����ε���.����, 0) > 0 Then
            n_ʱ�۷��� := 1;
          Else
            n_ʱ�۷��� := 0;
          End If;
        
          If Nvl(v_ʱ�۰����ε���.����, 0) = 0 Then
            Update ҩƷ���
            Set ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���۽��
            Where �ⷿid = v_ʱ�۰����ε���.�ⷿid And ҩƷid = v_ʱ�۰����ε���.ҩƷid And ���� = 1 And (���� Is Null Or ���� = 0);
          Else
            Update ҩƷ���
            Set ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���۽��, ���ۼ� = Decode(n_ʱ�۷���, 1, v_ʱ�۰����ε���.���ۼ�, ���ۼ�)
            Where �ⷿid = v_ʱ�۰����ε���.�ⷿid And ҩƷid = v_ʱ�۰����ε���.ҩƷid And ���� = 1 And ���� = v_ʱ�۰����ε���.����;
          End If;
        
          If Sql%RowCount = 0 Then
            Insert Into ҩƷ���
              (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, ���ۼ�)
            Values
              (v_ʱ�۰����ε���.�ⷿid, v_ʱ�۰����ε���.ҩƷid, v_ʱ�۰����ε���.����, 1, 0, 0, n_���۽��, n_���۽��,
               Decode(n_ʱ�۷���, 1, v_ʱ�۰����ε���.���ۼ�, Null));
          End If;
        
          Zl_δ��ҩƷ��¼_Delete(n_�շ�id);
        
          n_�۸�id := n_�շ�id;
          --End If;
        
          --��Ϣ����
          If n_�۸�id Is Not Null Then
            b_Message.Zlhis_Drug_011(n_�۸�id, 1);
          End If;
        End Loop;
      End Loop;
    End If;
  
    Update ҩƷ�շ���¼ Set ����� = User, ������� = Sysdate Where �۸�id = ����id_In;
    Update �շѼ�Ŀ Set �䶯ԭ�� = 1 Where ID = ����id_In;
  
    --����ҩƷĿ¼���շ�ϸĿ�еı��
    If ����_In = 1 Then
      Update �շ���ĿĿ¼ Set �Ƿ��� = 0 Where ID = n_�շ�ϸĿid;
    End If;
    --�ɱ��۵���
    Zl_�����շ���¼_�ɱ��۵���(n_�շ�ϸĿid);
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շ���¼_Adjust;
/

--125925:������,2018-05-17,�ƶ�����ӿڷ�װ
Create Or Replace Procedure Zl_Third_Advicecheck
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ�����ҽ���˶�/ȡ���˶ԣ�����д��
  --��Σ�xml_in
  --<IN>
  -- <TYPE>1</TYPE>   --�������ͣ�1���˶ԣ�0��ȡ���˶�
  -- <YZID>1162695</YZID>   --ҽ��id
  -- <FSH>202704</FSH>   --���ͺ�    
  -- <ZXSJ>2017-12-05 16:26:54</ZXSJ>   --ִ��ʱ��

  --���½ڵ�ȡ���˶�ʱ����
  -- <HDSJ>2017-12-05 10:00:00</HDSJ>   --�˶�ʱ��  
  -- <HDR></HDR>   --�˶���   
  --</IN>

  --���Σ�Xml_Out
  --<OUTPUT>
  --   <RESULT>true</RESULT>
  --</OUTPUT>

  --ʧ�ܣ�
  --<OUTPUT>
  --   <RESULT>false</RESULT>
  --   <ERROR>
  --     <MSG>��ϸ������ʾ</MSG>
  --   </ERROR>
  --</OUTPUT>

  n_Type     Number;
  n_ҽ��id   ����ҽ����¼.Id%Type;
  n_���ͺ�   ����ҽ������.���ͺ�%Type;
  d_ִ��ʱ�� ����ҽ��ִ��.Ҫ��ʱ��%Type;
  d_�˶�ʱ�� ����ҽ��ִ��.Ҫ��ʱ��%Type;
  v_�˶���   ����ҽ��ִ��.�˶���%Type;
Begin
  Select Extractvalue(Value(A), 'IN/TYPE'), Extractvalue(Value(A), 'IN/YZID') As ҽ��id,
         Extractvalue(Value(A), 'IN/FSH') As ���ͺ�,
         To_Date(Extractvalue(Value(A), 'IN/ZXSJ'), 'yyyy-mm-dd hh24:mi:ss') As ִ��ʱ��,
         To_Date(Extractvalue(Value(A), 'IN/HDSJ'), 'yyyy-mm-dd hh24:mi:ss') As d_�˶�ʱ��,
         Extractvalue(Value(A), 'IN/HDR') As �˶���
  Into n_Type, n_ҽ��id, n_���ͺ�, d_ִ��ʱ��, d_�˶�ʱ��, v_�˶���
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  If n_Type = 1 Then
    Zl_����ҽ���˶�_Insert(n_ҽ��id, n_���ͺ�, v_�˶���, d_ִ��ʱ��, d_�˶�ʱ��);  
  Else
    Zl_����ҽ���˶�_Delete(n_ҽ��id, n_���ͺ�, d_ִ��ʱ��);  
  End If;
  Xml_Out := Xmltype('<OUTPUT><RESULT>true</RESULT></OUTPUT>');
Exception
  When Others Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || SQLCode || '***' || SQLErrM ||
                       '</MSG></ERROR></OUTPUT>');
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Advicecheck;
/

--124431:���ϴ�,2018-05-17,�Һ����ռ�ú�����ȡ���
Create Or Replace Procedure Zl_���˹Һż�¼_����_Insert
(
  �����¼id_In    �ٴ������¼.Id%Type,
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���ݺ�_In        ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ���_In          ������ü�¼.���%Type,
  �۸񸸺�_In      ������ü�¼.�۸񸸺�%Type,
  ��������_In      ������ü�¼.��������%Type,
  �շ����_In      ������ü�¼.�շ����%Type,
  �շ�ϸĿid_In    ������ü�¼.�շ�ϸĿid%Type,
  ����_In          ������ü�¼.����%Type,
  ��׼����_In      ������ü�¼.��׼����%Type,
  ������Ŀid_In    ������ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In      ������ü�¼.�վݷ�Ŀ%Type,
  ���㷽ʽ_In      Varchar2,
  Ӧ�ս��_In      ������ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In      ������ü�¼.ʵ�ս��%Type,
  ���˿���id_In    ������ü�¼.���˿���id%Type,
  ��������id_In    ������ü�¼.��������id%Type,
  ִ�в���id_In    ������ü�¼.ִ�в���id%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
  ҽ������_In      �ҺŰ���.ҽ������%Type,
  ҽ��id_In        �ҺŰ���.ҽ��id%Type,
  ������_In        Number, --������¼�Ƿ���������
  ����_In          Number,
  �ű�_In          �ҺŰ���.����%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --ˢ���Һ�ʱʹ�õ�Ԥ�����,���Ϊ1����.
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��,���Ϊ1����.
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����,,���Ϊ1����.
  ���մ���id_In    ������ü�¼.���մ���id%Type,
  ������Ŀ��_In    ������ü�¼.������Ŀ��%Type,
  ͳ����_In      ������ü�¼.ͳ����%Type,
  ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
  ԤԼ�Һ�_In      Number := 0, --ԤԼ�Һ�ʱ��(��¼״̬=0,����ʱ��ΪԤԼʱ��),��ʱ����Ҫ���������ز���
  �շ�Ʊ��_In      Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
  ���ձ���_In      ������ü�¼.���ձ���%Type,
  ����_In          ���˹Һż�¼.����%Type := 0,
  ����_In          �Һ����״̬.���%Type := Null, --ԤԼʱ������ü�¼�ķ�ҩ�����ֶ�,�Һ�ʱ����Һż�¼
  ����_In          ���˹Һż�¼.����%Type := Null,
  ԤԼ����_In      Number := 0,
  ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
  ���ɶ���_In      Number := 0,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In      ����Ԥ����¼.������λ%Type := Null,
  ��������_In      Number := 0,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  �˺�����_In      Number := 1,
  ��Ԥ������ids_In Varchar2 := Null,
  �������˷ѱ�_In  Number := 0,
  ԤԼ˳���_In    �ٴ�������ſ���.ԤԼ˳���%Type := Null,
  ������������_In  Number := 0,
  �շѵ�_In        ���˹Һż�¼.�շѵ�%Type := Null,
  ���½������_In  Number := 1 --�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�����������
) As
  ---------------------------------------------------------------------------
  --
  --����:
  --     ��������_in:0-�����ҺŻ���ԤԼ 1-����Աӵ�мӺ�Ȩ�޼Ӻ�
  --     �������˷ѱ�_In:0-���޸Ĳ��˷ѱ� 1-�޸Ĳ��˷ѱ�
  ----------------------------------------------------------------------------
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select ����id, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1
     Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By NO, ����id
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ����id, NO;
  --���ܣ�����һ�в��˹Һŷ��ã�����������ܵ�����Ԥ����¼
  --       ͬʱ������صĻ��ܱ�(���˹ҺŻ��ܡ����û���)
  --       ��һ�з��ô���Ʊ��ʹ�����(����ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ���%Type;
  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
  v_�������� �ŶӽкŶ���.��������%Type;

  n_��ʱ��       Number;
  n_ԭʼ��ʱ��   Number;
  n_ʱ���޺�     Number;
  n_ʱ����Լ     Number;
  d_ʱ��ʱ��     Date;
  d_������ʱ�� Date;
  n_׷�Ӻ�       Number := 0; --����ʱ�ι��� ׷�ӹҺŵ����
  n_��Լ��       ���˹ҺŻ���.��Լ��%Type;
  n_ԤԼ��Чʱ�� Number;
  n_ʧЧ��       Number;
  n_ʧԼ�Һ�     Number := 0;
  n_��������     Number;
  n_����         Number := 0;
  n_ռ��         Number;
  d_����ʱ��     ������ü�¼.����ʱ��%Type;
  
  n_����id        ������ü�¼.Id%Type;
  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_��ǰ���      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  n_Ԥ��id        ����Ԥ����¼.Id%Type;
  n_�Һ�id        ���˹Һż�¼.Id%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_��id           ����ɿ����.Id%Type;
  n_�����         ������Ϣ.�����%Type;
  n_���           �Һ����״̬.���%Type;
  n_�������       �Һ����״̬.���%Type;
  n_��ſ���       �ҺŰ���.��ſ���%Type;
  n_����̨ǩ���Ŷ� Number;
  n_Count          Number;
  n_�޺���         Number(18);
  d_�Ŷ�ʱ��       Date;
  v_���㷽ʽ��¼   Varchar2(1000);
  d_���ʱ��       Date;
  v_�ѱ�           �ѱ�.����%Type;
  n_ʱ�����       Number := -1;
  n_ԤԼ���ɶ���   Number;
  v_���㷽ʽ       ���㷽ʽ.����%Type;
  v_��������       Varchar2(1000);
  v_��ǰ����       Varchar2(200);
  v_�������       ����Ԥ����¼.�������%Type;
  n_������       ����Ԥ����¼.��Ԥ��%Type;
  n_��������־     Number(2);
  n_ԤԼ˳���     �ٴ�������ſ���.ԤԼ˳���%Type;
  n_��Լ��         Number(18);
  n_�ѹ���         Number(4) := 0;
  n_Exists         Number;
  n_�ҳ��������� Number(4) := 0;
  n_��ʱ����ʾ     Number(3);
  n_����ģʽ       ������Ϣ.����ģʽ%Type;
  v_�Ŷ����       �ŶӽкŶ���.�Ŷ����%Type;
  v_������         �Һ����״̬.������%Type;
  v_��Ų���Ա     �Һ����״̬.����Ա����%Type;
  v_��Ż�����     �Һ����״̬.������%Type;
  v_���ʽ       ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  n_״̬           �ٴ�������ſ���.�Һ�״̬%Type;
Begin
  --��¼�����ж�
  If �����¼id_In Is Not Null Then
    Begin
      Select 1
      Into n_Exists
      From �ٴ������¼
      Where ID = �����¼id_In And Nvl(�Ƿ񷢲�, 0) = 1 And Nvl(�Ƿ�����, 0) = 0;
    Exception
      When Others Then
        v_Err_Msg := '�޷�ȷ�������¼����������¼�Ƿ���ڻ�������';
        Raise Err_Item;
    End;
  End If;

  --��ȡ��ǰ��������
  Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);

  If �ѱ�_In Is Null Then
    Begin
      Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
      Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
    Exception
      When Others Then
        v_Err_Msg := '�޷�ȷ�����˷ѱ�����ȱʡ�ѱ��Ƿ���ȷ���ã�';
        Raise Err_Item;
    End;
  Else
    v_�ѱ� := �ѱ�_In;
    If Nvl(�������˷ѱ�_In, 0) = 1 Then
      Begin
        Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
      Exception
        When Others Then
          v_Err_Msg := 'û���ҵ���Ӧ�Ĳ��ˣ�';
          Raise Err_Item;
      End;
    End If;
  End If;

  If Nvl(������������_In, 0) = 1 Then
    Begin
      Update ������Ϣ Set ���� = ����_In Where ����id = ����id_In;
    Exception
      When Others Then
        v_Err_Msg := 'û���ҵ���Ӧ�Ĳ��ˣ�';
        Raise Err_Item;
    End;
  End If;

  If �����_In Is Not Null Then
    Begin
      Select Nvl(�����, 0) Into n_����� From ������Ϣ Where ����id = ����id_In;
    Exception
      When Others Then
        n_����� := 0;
    End;
    If n_����� = 0 Then
      Update ������Ϣ Set ����� = �����_In Where ����id = ����id_In;
    End If;
  End If;

  Begin
    Update �ٴ�������ſ���
    Set �Һ�״̬ = 0
    Where ��¼id = �����¼id_In And ��� = ����_In And Nvl(�Һ�״̬, 0) = 3 And ����Ա���� = ����Ա����_In;
  Exception
    When Others Then
      Null;
  End;

  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    --�ҺŻ���ԤԼ����
    --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
    Select Count(*)
    Into n_Count
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ In (1, 3) And ��� = ���_In And NO = ���ݺ�_In;
    If n_Count <> 0 Then
      v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
      Raise Err_Item;
    End If;
  
    --��ȡ���㷽ʽ����
    Begin
      Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
    Exception
      When Others Then
        v_�ֽ� := '�ֽ�';
    End;
    Begin
      Select ���� Into v_�����ʻ� From ���㷽ʽ Where ���� = 3;
    Exception
      When Others Then
        v_�����ʻ� := '�����ʻ�';
    End;
  End If;

  IF Nvl(����_In, 0) <> 0 then
    Begin
      Select 1 Into n_ռ�� From �ٴ�������ſ���
      Where ��¼ID = �����¼id_In And ��� = ����_In And (�Һ�״̬ In (1,2,4) Or �Һ�״̬ in (3, 5) And (����Ա����_In <> v_��Ų���Ա Or v_������ <> v_��Ż�����));
    Exception
      When Others Then
        n_ռ�� := 0;
    End;
  End IF;
  IF Nvl(n_ռ��, 0) = 1 And ���_In = 1 And n_��ʱ�� > 0 And Nvl(n_��ſ���, 0) = 1 Then
    Begin
      Select ���, ��ʼʱ��
      Into n_���, d_����ʱ��
      From (Select ���, ��ʼʱ�� From �ٴ�������ſ��� Where ��� > ����_In And Nvl(�Һ�״̬, 0) = 0 Order By ���)
      Where Rownum < 2;
    Exception
      When Others Then
        n_��� := Null;
        d_����ʱ�� := ����ʱ��_In;
    End;
  Else
    n_��� := ����_In;
    d_����ʱ�� := ����ʱ��_In;
  End IF;

  --��ȡ�Ƿ��ʱ��
  Begin
    Select Nvl(�Ƿ��ʱ��, 0), Nvl(�Ƿ���ſ���, 0), �޺���, ��Լ��
    Into n_��ʱ��, n_��ſ���, n_�޺���, n_��Լ��
    From �ٴ������¼
    Where ID = �����¼id_In;
    n_ԭʼ��ʱ�� := n_��ʱ��;
  Exception
    When Others Then
      n_��ʱ��     := 0;
      n_ԭʼ��ʱ�� := n_��ʱ��;
      n_��ſ���   := 0;
      n_�޺���     := Null;
      n_��Լ��     := Null;
  End;

  If n_��� Is Null And n_��ʱ�� = 1 And n_��ſ��� = 0 Then
    Begin
      Select ���
      Into n_���
      From �ٴ�������ſ���
      Where ��¼id = �����¼id_In And ��ʼʱ�� = d_����ʱ�� And Rownum < 2;
    Exception
      When Others Then
        n_��� := Null;
    End;
  End If;

  --��ʱ�ιҺ�ʱ�ж��Ƿ��ǹ��ڹҺ� Ҳ����׷�Ӻŵ���� Ŀǰֻ���ר�Һŷ�ʱ�ν��д���
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 And Nvl(n_��ſ���, 0) = 1 And ����_In Is Null And Nvl(��������_In, 0) = 0 Then
    Begin
      Select Max(To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_������ʱ��
      From �ٴ�������ſ���
      Where ��¼id = �����¼id_In And Nvl(����, 0) <> 0;
    
      n_׷�Ӻ� := Case Sign(d_����ʱ�� - d_������ʱ��)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_׷�Ӻ� := 0;
    End;
  End If;
  d_ʱ��ʱ�� := d_����ʱ��;

  If ���_In = 1 And n_��ʱ�� > 0 Then
    If Nvl(n_��ſ���, 0) = 1 Then
      Begin
        Select Nvl(���, 0),
               To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               ����, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ����)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �ٴ�������ſ���
        Where ��¼id = �����¼id_In And ��� = n_���;
      Exception
        When Others Then
          n_ʱ����� := -1;
          n_��ʱ��   := 0;
          d_ʱ��ʱ�� := d_����ʱ��;
          n_ʱ���޺� := 0;
          n_ʱ����Լ := 0;
      End;
    Else
      --�Һ�ʱ��� �Ƿ����ʱ��,����ʱ��,��ʱ�ε�����������ȡ����
      Begin
        Select Nvl(���, 0),
               To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               ����, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ����)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �ٴ�������ſ���
        Where ��¼id = �����¼id_In And ��� = n_��� And ԤԼ˳��� Is Null;
      Exception
        When Others Then
          n_ʱ����� := -1;
          n_��ʱ��   := 0;
          d_ʱ��ʱ�� := d_����ʱ��;
          n_ʱ���޺� := 0;
          n_ʱ����Լ := 0;
      End;
    End If;
  End If;

  If ���_In = 1 Then
    --��ȡ��ǰδʹ�õ����
    If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
      n_ԤԼ��Чʱ�� := Zl_To_Number(zl_GetSysParameter('ԤԼ��Чʱ��', 1111));
      n_ʧԼ�Һ�     := Zl_To_Number(zl_GetSysParameter('ʧԼ���ڹҺ�', 1111));
    End If;
    If Nvl(n_��ſ���, 0) = 1 And n_��ʱ�� = 0 Then
      --<��ſ��� δ����ʱ�� ��ȡ���õ�������,�Լ��Ѿ�ʹ�õ�����>
      Begin
        --������
        Select Count(1) Into n_�������� From ���˹Һż�¼ Where �����¼id = �����¼id_In And ��¼״̬ = 1;
        Select Max(���) Into n_������� From �ٴ�������ſ��� Where ��¼id = �����¼id_In;
      Exception
        When Others Then
          n_������� := 0;
          n_�������� := 0;
      End;
      Begin
        --������
        Select Sum(Nvl(����, 0))
        
        Into n_��Լ��
        From �ٴ�������ſ���
        Where ��¼id = �����¼id_In And Nvl(�Һ�״̬, 0) = 2;
      Exception
        When Others Then
          n_��Լ�� := 0;
      End;
    
      n_ʧЧ�� := 0;
      If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        If Nvl(n_ԤԼ��Чʱ��, 0) <> 0 And Nvl(n_ʧԼ�Һ�, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_ԤԼ��Чʱ�� / 24 / 60) - ԤԼʱ��), 1, 1, 0))
            Into n_ʧЧ��
            From ���˹Һż�¼
            Where �����¼id = �����¼id_In And ��¼״̬ = 1 And ��¼���� = 2;
          Exception
            When Others Then
              n_ʧЧ�� := 0;
          End;
        End If;
      End If;
    
      If n_ԭʼ��ʱ�� = 0 Then
        Begin
          Select Min(���) Into n_������� From �ٴ�������ſ��� Where ��¼id = �����¼id_In And Nvl(�Һ�״̬, 0) = 0;
          If n_��� Is Null Then
            n_��� := Nvl(n_�������, 0);
          End If;
        Exception
          When Others Then
            Select Max(���)
            Into n_�������
            From �ٴ�������ſ���
            Where ��¼id = �����¼id_In And Nvl(�Һ�״̬, 0) <> 0;
            If n_��� Is Null Then
              n_��� := Nvl(n_�������, 0) + 1;
            End If;
        End;
      Else
        Select Max(���) Into n_������� From �ٴ�������ſ��� Where ��¼id = �����¼id_In;
        If n_��� Is Null Then
          n_��� := Nvl(n_�������, 0) + 1;
        End If;
      End If;
      --<��ſ��� δ����ʱ�� ��ȡ���õ������� �Լ��Ѿ�ʹ�õ����� --end>
    
      --�ǼӺŵ������Ҫ����Ƿ񳬹�������
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�ʱ���
          --������ſ���δ��ʱ�� �ﵽ������
          If n_�޺��� <= n_�������� And n_�޺��� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(d_����ʱ��), 'yyyy-mm-dd ') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        
        Else
          --ԤԼʱ���
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
        
          If n_��Լ�� <= n_��Լ�� And n_��Լ�� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(d_����ʱ��), 'yyyy-mm-dd') || '�Ѵﵽ�����Լ����';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --������ſ���,δ��ʱ�� �Ӻ����   ������,����Ժ������������Ժ󲹳�
      End If;
    
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 0 Then
      --<--��ͨ�ŷ�ʱ�� ����ֻ��ԤԼһ�����-->
      If ��������_In = 0 Then
        --<����ԤԼ�Һ�-->
        Begin
          Select Count(0) As �ѹ���, Nvl(Sum(Decode(Nvl(Sign(a.��ʼʱ�� - d_ʱ��ʱ��), 0), 0, 1, 0)), 0) As ��Լ��
          Into n_�ѹ���, n_��Լ��
          From �ٴ�������ſ��� A
          Where ��¼id = �����¼id_In And �Һ�״̬ Not In (0, 4, 5);
        Exception
          When Others Then
            n_�ѹ��� := 0;
            n_��Լ�� := 0;
        End;
      
        n_ʱ����Լ := n_ʱ���޺�; --��ͨ�ŷ�ʱ�ε����,29��n_ʱ����Լʼ����0 �������⴦��
        --�����������
        If n_��Լ�� = 0 Then
          n_��Լ�� := n_�޺���;
        End If;
        If n_ʱ����Լ <= n_��Լ�� Or n_��Լ�� <= n_��Լ�� Then
          v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����Լ����';
          Raise Err_Item;
        End If;
        If n_�޺��� <= n_�ѹ��� Then
          v_Err_Msg := '�ű�' || �ű�_In || To_Char(d_����ʱ��, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
          Raise Err_Item;
        End If;
      End If;
    
      --û�дﵽʱ�ε��޺��� �����ڵ�ǰʱ������׷��
    
      --��ȡ����ҳ���������
      Select Nvl(Max(���), 0)
      Into n_�ҳ���������
      From �ٴ�������ſ��� A
      Where ��¼id = �����¼id_In And ԤԼ˳��� Is Null And �Һ�״̬ Not In (0, 5);
      If ԤԼ˳���_In Is Not Null Then
        n_ԤԼ˳��� := ԤԼ˳���_In;
      Else
        Begin
          Select Nvl(Max(ԤԼ˳���), 0) + 1
          Into n_ԤԼ˳���
          From �ٴ�������ſ���
          Where ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� Is Not Null;
        Exception
          When Others Then
            n_ԤԼ˳��� := Null;
        End;
      End If;
      --�������
      n_��� := RPad(Nvl(n_ʱ�����, 0), Length(n_�޺���) + Length(Nvl(n_ʱ�����, 0)), 0) + n_ԤԼ˳���;
      If n_ԤԼ˳��� Is Null Then
        n_��� := Nvl(n_�ҳ���������, 0) + 1;
      End If;
    
      --<--��ͨ�ŷ�ʱ��--End>
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 1 Then
      --<������ſ��� ����ʱ��
      --ר�Һŷ�ʱ��
      Begin
        Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(��ʼʱ�� - d_ʱ��ʱ��), 0, 1, 0))
        Into n_�������, n_�ѹ���, n_��������
        From �ٴ�������ſ���
        Where ��¼id = �����¼id_In And �Һ�״̬ Not In (0, 4, 5);
      Exception
        When Others Then
          n_������� := 0;
          n_�������� := 0;
          n_�ѹ���   := 0;
      End;
    
      n_ʧЧ�� := 0;
      If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        If Nvl(n_ԤԼ��Чʱ��, 0) <> 0 And Nvl(n_ʧԼ�Һ�, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_ԤԼ��Чʱ�� / 24 / 60) - ��ʼʱ��), 1, 1, 0))
            Into n_ʧЧ��
            From �ٴ�������ſ���
            Where ��¼id = �����¼id_In And ��ʼʱ�� Between Trunc(Sysdate) And Sysdate And Nvl(�Һ�״̬, 0) = 2;
          Exception
            When Others Then
              n_ʧЧ�� := 0;
          End;
        End If;
      End If;
    
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        
          --�Һ� ׷�Ӻ���ʱ�����ʱ���޺���
          If n_ʱ���޺� <= n_�������� And Nvl(n_׷�Ӻ�, 0) = 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� - n_ʧЧ�� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(d_����ʱ��, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        
        Else
          --�Һ�
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
          If n_��Լ�� <= n_�������� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(d_����ʱ��, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_��� Is Null Then
        --�������
        If Nvl(n_�������, 0) < Nvl(n_ʱ�����, 0) Then
          n_������� := Nvl(n_ʱ�����, 0);
        End If;
        n_��� := Nvl(n_�������, 0) + 1;
      End If;
    Elsif Nvl(n_��ʱ��, 0) = 0 And Nvl(n_��ſ���, 0) = 0 And Nvl(������_In, 0) = 0 And Nvl(�ű�_In, 0) > 0 Then
      ---<--��ͨ��  -->
      Begin
        Select �ѹ���, ��Լ�� Into n_��������, n_��Լ�� From �ٴ������¼ Where ID = �����¼id_In;
      Exception
        When Others Then
          n_�������� := 0;
          n_��Լ��   := 0;
      End;
      If Nvl(��������_In, 0) = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�
          If Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(d_����ʱ��, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        Else
          --ԤԼ
          If (Nvl(n_��Լ��, 0) > 0 And Nvl(n_��Լ��, 0) >= Nvl(n_��Լ��, 0)) Or
             (Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0) Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(d_����ʱ��, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--��ͨ��  -->
    End If;
  End If;

  --���¹Һ����״̬
  If ���_In = 1 And Not n_��� Is Null Then
    If n_��ʱ�� = 1 Then
      d_���ʱ�� := d_����ʱ��;
    Else
      d_���ʱ�� := Trunc(d_����ʱ��);
    End If;
    --������ŵĴ���
    Begin
      If n_ԤԼ˳��� Is Null Then
        Select ����Ա����, ����վ����
        Into v_��Ų���Ա, v_��Ż�����
        From �ٴ�������ſ���
        Where Nvl(�Һ�״̬, 0) = 5 And ��¼id = �����¼id_In And ��� = n_���;
      Else
        Select ����Ա����, ����վ����
        Into v_��Ų���Ա, v_��Ż�����
        From �ٴ�������ſ���
        Where Nvl(�Һ�״̬, 0) = 5 And ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� = n_ԤԼ˳���;
      End If;
      n_���� := 1;
    Exception
      When Others Then
        v_��Ų���Ա := Null;
        v_��Ż����� := Null;
        n_����       := 0;
    End;
    If n_���� = 0 Then
      If n_ԤԼ˳��� Is Null Then
        Update �ٴ�������ſ���
        Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
        Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) = 3 And ����Ա���� = ����Ա����_In;
      Else
        Update �ٴ�������ſ���
        Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
        Where ��¼id = �����¼id_In And ��� = n_��� And ԤԼ˳��� = n_ԤԼ˳��� And Nvl(�Һ�״̬, 0) = 3 And ����Ա���� = ����Ա����_In;
      End If;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_��ʱ��, 0) > 0 Then
            If Nvl(n_��ſ���, 0) = 1 Then
              --��ʱ�κ�ר�Һ� ʧԼ��ԤԼ������Һ�
              Update �ٴ�������ſ���
              Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա���� = ����Ա����_In
              Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) In (0, 2);
              If Sql%NotFound Then
                Begin
                  Select �Һ�״̬ Into n_״̬ From �ٴ�������ſ��� Where ��¼id = �����¼id_In And ��� = n_���;
                Exception
                  When Others Then
                    n_״̬ := -1;
                End;
                If n_״̬ = -1 Then
                  Insert Into �ٴ�������ſ���
                    (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ��ע)
                    Select �����¼id_In, n_���, d_���ʱ��, d_���ʱ��, 1, Decode(ԤԼ�Һ�_In, 1, 1, 0), Decode(ԤԼ�Һ�_In, 1, 2, 1), Null,
                           Null, Null, ����Ա����_In, '׷�Ӻ�'
                    From Dual;
                Else
                  v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.';
                  Raise Err_Item;
                End If;
              End If;
            Else
              If Nvl(ԤԼ����_In, 0) = 1 Then
                Insert Into �ٴ�������ſ���
                  (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ��ע, ԤԼ˳���)
                  Select ��¼id, ���, ��ʼʱ��, ��ֹʱ��, 1, 1, Decode(ԤԼ�Һ�_In, 1, 2, 1), Null, Null, Null, ����Ա����_In, n_���, n_ԤԼ˳���
                  From �ٴ�������ſ���
                  Where ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� Is Null;
              End If;
            End If;
          Else
            If Nvl(n_��ſ���, 0) = 1 Then
              Update �ٴ�������ſ���
              Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա���� = ����Ա����_In
              Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) = 0;
            
              If Sql%RowCount = 0 Then
                Begin
                  Select �Һ�״̬ Into n_״̬ From �ٴ�������ſ��� Where ��¼id = �����¼id_In And ��� = n_���;
                Exception
                  When Others Then
                    n_״̬ := -1;
                End;
                If n_״̬ = -1 Then
                  Insert Into �ٴ�������ſ���
                    (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ��ע)
                    Select �����¼id_In, n_���, d_����ʱ��, d_����ʱ��, 1, Decode(ԤԼ�Һ�_In, 1, 1, 0), Decode(ԤԼ�Һ�_In, 1, 2, 1),
                           Null, Null, Null, ����Ա����_In, '׷�Ӻ�'
                    From Dual;
                Else
                  v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.';
                  Raise Err_Item;
                End If;
              End If;
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.';
            Raise Err_Item;
        End;
      End If;
    Else
      If ����Ա����_In <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
        v_Err_Msg := '���' || n_��� || '�ѱ�������' || v_������ || '����,������ѡ��һ�����.';
        Raise Err_Item;
      Else
        If n_ԤԼ˳��� Is Null Then
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
          Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) = 5 And ����Ա���� = ����Ա����_In And ����վ���� = v_������;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
          Where ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� = n_ԤԼ˳��� And Nvl(�Һ�״̬, 0) = 5 And ����Ա���� = ����Ա����_In And
                ����վ���� = v_������;
        End If;
      End If;
    End If;
  End If;

  --�������˹Һŷ���(���ܵ����ǻ������������)
  Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual; --Ӧ��ͨ������õ�

  Insert Into ������ü�¼
    (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id, �շ����,
     ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����,
     ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
  Values
    (n_����id, 4, Decode(ԤԼ�Һ�_In, 1, 0, 1), ���_In, Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), ��������_In, ���ݺ�_In, Ʊ�ݺ�_In, 1, ����_In,
     ������_In, Decode(ԤԼ�Һ�_In, 1, To_Char(n_���), ����_In), Decode(����id_In, 0, Null, ����id_In),
     Decode(�����_In, 0, Null, �����_In), ���ʽ_In, ����_In, Decode(����_In, Null, Null, �Ա�_In), Decode(����_In, Null, Null, ����_In),
     v_�ѱ�, ���˿���id_In, �շ����_In, �ű�_In, �շ�ϸĿid_In, ������Ŀid_In, �վݷ�Ŀ_In, 1, ����_In, ��׼����_In, Ӧ�ս��_In, ʵ�ս��_In,
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��_In)),
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In)), Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ��������id_In,
     ����Ա����_In, Decode(ԤԼ�Һ�_In, 1, ����Ա����_In, Null), ִ�в���id_In, ҽ������_In, ����Ա���_In, ����Ա����_In, d_����ʱ��, �Ǽ�ʱ��_In, ���մ���id_In,
     ������Ŀ��_In, ���ձ���_In, ͳ����_In, Decode(�շѵ�_In, Null, ժҪ_In, '����:' || �շѵ�_In), ԤԼ��ʽ_In, Decode(ԤԼ�Һ�_In, 1, Null, n_��id));

  --���ܽ��㵽����Ԥ����¼
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 0 Then
    If Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0 And ���_In = 1 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
         ��������)
      Values
        (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), v_�ֽ�, 0, �Ǽ�ʱ��_In, ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�',
         n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4);
    End If;
    If Nvl(�ֽ�֧��_In, 0) <> 0 And ���_In = 1 Then
      v_��������     := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
      v_���㷽ʽ��¼ := '';
      While v_�������� Is Not Null Loop
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        v_������� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_��������־ := To_Number(v_��ǰ����);
      
        If Instr('|' || v_���㷽ʽ��¼ || '|', '|' || Nvl(v_���㷽ʽ, v_�ֽ�) || '|') <> 0 Then
          v_Err_Msg := 'ʹ�����ظ��Ľ��㷽ʽ,����!';
          Raise Err_Item;
        Else
          v_���㷽ʽ��¼ := v_���㷽ʽ��¼ || '|' || Nvl(v_���㷽ʽ, v_�ֽ�);
        End If;
      
        If n_��������־ = 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, Null, Null, Null, Null, Null, ������λ_In, 4, v_�������);
        Else
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4,
             v_�������);
        
          If Nvl(���㿨���_In, 0) <> 0 And Nvl(�ֽ�֧��_In, 0) <> 0 Then
            Zl_���˿������¼_֧��(���㿨���_In, ����_In, 0, Nvl(n_������, 0), n_Ԥ��id, ����Ա���_In, ����Ա����_In, �Ǽ�ʱ��_In);
          End If;
        End If;
      
        If Nvl(���½������_In, 1) = 1 Then
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) + n_������
          Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(v_���㷽ʽ, v_�ֽ�)
          Returning ��� Into n_����ֵ;
        
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (����Ա����_In, Nvl(v_���㷽ʽ, v_�ֽ�), 1, n_������);
            n_����ֵ := n_������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From ��Ա�ɿ����
            Where �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(v_���㷽ʽ, v_�ֽ�) And ���� = 1 And Nvl(���, 0) = 0;
          End If;
        End If;
      
        v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
      End Loop;
    End If;
  
    --����ҽ���Һ�
    If Nvl(����֧��_In, 0) <> 0 And ���_In = 1 Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), v_�����ʻ�, ����֧��_In, �Ǽ�ʱ��_In, ����Ա���_In,
         ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id, 4);
    End If;
  
    --���ھ��￨ͨ��Ԥ����Һ�
    If Nvl(Ԥ��֧��_In, 0) <> 0 And ���_In = 1 Then
      n_Ԥ����� := Ԥ��֧��_In;
      For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
        n_��ǰ��� := Case
                    When r_Deposit.��� - n_Ԥ����� < 0 Then
                     r_Deposit.���
                    Else
                     n_Ԥ�����
                  End;
      
        If r_Deposit.����id = 0 Then
          --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
          Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 4 Where ID = r_Deposit.ԭԤ��id;
        
        End If;
        --���ϴ�ʣ���
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
           ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                 �Ǽ�ʱ��_In, ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
      
        --���²���Ԥ�����
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
        Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2);
        --����Ƿ��Ѿ�������
        If r_Deposit.��� < n_Ԥ����� Then
          n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
        Else
          n_Ԥ����� := 0;
        End If;
      
        If n_Ԥ����� = 0 Then
          Exit;
        End If;
      End Loop;
      If n_Ԥ����� > 0 Then
        v_Err_Msg := 'Ԥ���಻��֧������֧�����,���ܼ���������';
        Raise Err_Item;
      
      End If;
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
    End If;
  
    --��ػ��ܱ�Ĵ���
    --��Ա�ɿ����
    If ���_In = 1 And Nvl(���½������_In, 1) = 1 Then
      If Nvl(����֧��_In, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + ����֧��_In
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
        Returning ��� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
          n_����ֵ := ����֧��_In;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_�����ʻ� And Nvl(���, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --���˹ҺŻ���(ֻ����һ��,�ҵ�����ȡ�����Ѳ�����)
  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    --���˱��ξ���(�Է���ʱ��Ϊ׼)
    If Nvl(����id_In, 0) <> 0 And ���_In = 1 Then
      Update ������Ϣ Set ����ʱ�� = d_����ʱ��, ����״̬ = 1, �������� = ����_In Where ����id = ����id_In;
    End If;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
  
    --�������
    Update �������
    Set ������� = Nvl(�������, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (����id_In, 1, 1, Nvl(ʵ�ս��_In, 0), 0);
    End If;
  
    --����δ�����
    Update ����δ�����
    Set ��� = Nvl(���, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(���˿���id_In, 0) And
          Nvl(��������id, 0) = Nvl(��������id_In, 0) And Nvl(ִ�в���id, 0) = Nvl(ִ�в���id_In, 0) And ������Ŀid + 0 = ������Ŀid_In And
          ��Դ;�� + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (����id_In, Null, Null, ���˿���id_In, ��������id_In, ִ�в���id_In, ������Ŀid_In, 1, Nvl(ʵ�ս��_In, 0));
    End If;
  End If;

  --���˹Һż�¼
  If �ű�_In Is Not Null And ���_In = 1 Then
    --And Nvl(ԤԼ�Һ�_In, 0) = 0
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
    Begin
      Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
    Exception
      When Others Then
        v_���ʽ := Null;
    End;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���,
       ����Ա����, ����, ����, ����, ԤԼ, ԤԼ��ʽ, ժҪ, ������ˮ��, ����˵��, ������λ, ����ʱ��, ������, ԤԼ����Ա, ԤԼ����Ա���, ����, ҽ�Ƹ��ʽ, �����¼id, �շѵ�)
    Values
      (n_�Һ�id, ���ݺ�_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, 2, 1), 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, �ű�_In, ����_In, ����_In,
       Null, ִ�в���id_In, ҽ������_In, 0, Null, �Ǽ�ʱ��_In, d_����ʱ��, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, d_����ʱ��, Null), ����Ա���_In,
       ����Ա����_In, ����_In, n_���, ����_In, Decode(ԤԼ����_In, 1, 1, 0), ԤԼ��ʽ_In, ժҪ_In, ������ˮ��_In, ����˵��_In, ������λ_In,
       Decode(Nvl(ԤԼ�Һ�_In, 0), 0, �Ǽ�ʱ��_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 0, ����Ա����_In, Null),
       Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա����_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա���_In, Null),
       Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, �����¼id_In, �շѵ�_In);
  
    If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
      Update ���˹Һż�¼
      Set ԤԼ = 1, ԤԼʱ�� = d_����ʱ��, ԤԼ����Ա = ����Ա����_In, ԤԼ����Ա��� = ����Ա���_In
      Where ID = n_�Һ�id;
    End If;
  
    n_ԤԼ���ɶ��� := 0;
    If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
      n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    End If;
  
    --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
    If Nvl(���ɶ���_In, 0) <> 0 And Nvl(ԤԼ�Һ�_In, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
        n_��ʱ����ʾ := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
        If Nvl(ԤԼ�Һ�_In, 0) = 1 And n_��ʱ����ʾ = 1 And n_��ʱ�� = 1 Then
          n_��ʱ����ʾ := 1;
        Else
          n_��ʱ����ʾ := Null;
        End If;
        --��������
        --.����ִ�в��š� �ķ�ʽ���ɶ���
        v_�������� := ִ�в���id_In;
        v_�ŶӺ��� := Zlgetnextqueue(ִ�в���id_In, n_�Һ�id, �ű�_In || '|' || n_���);
        v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, �ű�_In, n_���, d_����ʱ��);
        --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
        Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, ִ�в���id_In, v_�ŶӺ���, Null, ����_In, ����id_In, ����_In, ҽ������_In, d_�Ŷ�ʱ��, ԤԼ��ʽ_In,
                         n_��ʱ����ʾ, v_�Ŷ����);
      
        --�Һ������Ŷ�
        If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
          Update ���˹Һż�¼ Set ��¼��־ = 1 Where ID = n_�Һ�id;
        End If;
      End If;
    End If;
  End If;
  --���˵�����Ϣ
  If ����id_In Is Not Null And ���_In = 1 Then
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(����ģʽ, 0) Into n_����ģʽ From ������Ϣ Where ����id = ����id_In;
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ĳ�����Ϣ,������Һ�';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If n_����ģʽ = 1 And Nvl(����ģʽ_In, 0) = 0 Then
        --�����Ѿ���"�����ƺ�����",������"�Ƚ�������Ƶ�",�����Ƿ����δ������
        Select Count(1)
        Into n_Count
        From ����δ�����
        Where ����id = ����id_In And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --����δ�������ݣ������Ƚ���������ִ��
          v_Err_Msg := '��ǰ���˵ľ���ģʽΪ�����ƺ�����Ҵ���δ����ã�����������ò��˵ľ���ģʽ,������ȶ�δ����ý��ʺ��ٹҺŻ򲻵������˵ľ���ģʽ!';
          Raise Err_Item;
        End If;
        --���
        --δ����ҽ��ҵ��ģ�����ʱ�͹Һŵ�,��Ҫ��֤ͬһ�εľ���ģʽ��һ����(�����Ѿ���飬�����ٴ���)
      End If;
      Update ������Ϣ Set ����ģʽ = ����ģʽ_In Where ����id = ����id_In;
    End If;
  
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Exists (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
  
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = Sysdate
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, Sysdate) > Sysdate;
    End If;
  End If;
  If ���_In = 1 Then
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 1, n_�Һ�id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_�Һ�id, ���ݺ�_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_����_Insert;
/

--124431:���ϴ�,2018-05-17,�Һ����ռ�ú�����ȡ���
Create Or Replace Procedure Zl_���˹Һż�¼_Insert
(
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���ݺ�_In        ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ���_In          ������ü�¼.���%Type,
  �۸񸸺�_In      ������ü�¼.�۸񸸺�%Type,
  ��������_In      ������ü�¼.��������%Type,
  �շ����_In      ������ü�¼.�շ����%Type,
  �շ�ϸĿid_In    ������ü�¼.�շ�ϸĿid%Type,
  ����_In          ������ü�¼.����%Type,
  ��׼����_In      ������ü�¼.��׼����%Type,
  ������Ŀid_In    ������ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In      ������ü�¼.�վݷ�Ŀ%Type,
  ���㷽ʽ_In      ����Ԥ����¼.���㷽ʽ%Type, --�ֽ�Ľ�������
  Ӧ�ս��_In      ������ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In      ������ü�¼.ʵ�ս��%Type,
  ���˿���id_In    ������ü�¼.���˿���id%Type,
  ��������id_In    ������ü�¼.��������id%Type,
  ִ�в���id_In    ������ü�¼.ִ�в���id%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
  ҽ������_In      �ҺŰ���.ҽ������%Type,
  ҽ��id_In        �ҺŰ���.ҽ��id%Type,
  ������_In        Number, --������¼�Ƿ���������
  ����_In          Number,
  �ű�_In          �ҺŰ���.����%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --ˢ���Һ�ʱʹ�õ�Ԥ�����,���Ϊ1����.
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��,���Ϊ1����.
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����,,���Ϊ1����.
  ���մ���id_In    ������ü�¼.���մ���id%Type,
  ������Ŀ��_In    ������ü�¼.������Ŀ��%Type,
  ͳ����_In      ������ü�¼.ͳ����%Type,
  ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
  ԤԼ�Һ�_In      Number := 0, --ԤԼ�Һ�ʱ��(��¼״̬=0,����ʱ��ΪԤԼʱ��),��ʱ����Ҫ���������ز���
  �շ�Ʊ��_In      Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
  ���ձ���_In      ������ü�¼.���ձ���%Type,
  ����_In          ���˹Һż�¼.����%Type := 0,
  ����_In          �Һ����״̬.���%Type := Null, --ԤԼʱ������ü�¼�ķ�ҩ�����ֶ�,�Һ�ʱ����Һż�¼
  ����_In          ���˹Һż�¼.����%Type := Null,
  ԤԼ����_In      Number := 0,
  ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
  ���ɶ���_In      Number := 0,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In      ����Ԥ����¼.������λ%Type := Null,
  ��������_In      Number := 0,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  �˺�����_In      Number := 1,
  ��Ԥ������ids_In Varchar2 := Null,
  �������˷ѱ�_In  Number := 0,
  ������������_In  Number := 0,
  �շѵ�_In        ���˹Һż�¼.�շѵ�%Type := Null,
  ���½������_In  Number := 1
) As
  ---------------------------------------------------------------------------
  --
  --����:
  --     ��������_in:0-�����ҺŻ���ԤԼ 1-����Աӵ�мӺ�Ȩ�޼Ӻ�
  --     �������˷ѱ�_In:0-���޸Ĳ��˷ѱ� 1-�޸Ĳ��˷ѱ�
  --     ���½������_In:0-��zl_��Ա�ɿ����_Update �и��� 1-�ڱ������и���
  ----------------------------------------------------------------------------
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select ����id, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1
     Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By NO, ����id
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ����id, NO;
  --���ܣ�����һ�в��˹Һŷ��ã�����������ܵ�����Ԥ����¼
  --       ͬʱ������صĻ��ܱ�(���˹ҺŻ��ܡ����û���)
  --       ��һ�з��ô���Ʊ��ʹ�����(����ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ���%Type;
  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
  v_�������� �ŶӽкŶ���.��������%Type;

  n_��ʱ��       Number;
  n_ʱ���޺�     Number;
  n_ʱ����Լ     Number;
  d_ʱ��ʱ��     Date;
  d_������ʱ�� Date;
  n_׷�Ӻ�       Number := 0; --����ʱ�ι��� ׷�ӹҺŵ����
  n_��Լ��       ���˹ҺŻ���.��Լ��%Type;
  n_ԤԼ��Чʱ�� Number;
  n_ʧЧ��       Number;
  n_ʧԼ�Һ�     Number := 0;
  n_��������     Number;
  n_����         Number := 0;

  n_����id        ������ü�¼.Id%Type;
  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_��ǰ���      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  n_Ԥ��id        ����Ԥ����¼.Id%Type;
  n_�Һ�id        ���˹Һż�¼.Id%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_��id           ����ɿ����.Id%Type;
  n_�����         ������Ϣ.�����%Type;
  n_���           �Һ����״̬.���%Type;
  n_�������       �Һ����״̬.���%Type;
  n_��ſ���       �ҺŰ���.��ſ���%Type;
  n_����̨ǩ���Ŷ� Number;
  n_Count          Number;
  n_�޺���         Number(18);
  d_�Ŷ�ʱ��       Date;
  d_���ʱ��       Date;
  v_�ѱ�           �ѱ�.����%Type;
  n_ʱ�����       Number := -1;
  n_ԤԼ���ɶ���   Number;
  n_����id         �ҺŰ���.Id%Type;
  n_�ƻ�id         �ҺŰ��żƻ�.Id%Type := 0;
  v_����           �ҺŰ�������.������Ŀ%Type;
  n_��Լ��         Number(18);
  n_�ѹ���         Number(4) := 0;
  n_ռ��           Number;
  d_����ʱ��       ������ü�¼.����ʱ��%Type;
  
  n_�ҳ��������� Number(4) := 0;
  n_����ģʽ       ������Ϣ.����ģʽ%Type;
  v_�Ŷ����       �ŶӽкŶ���.�Ŷ����%Type;
  v_������         �Һ����״̬.������%Type;
  v_��Ų���Ա     �Һ����״̬.����Ա����%Type;
  v_��Ż�����     �Һ����״̬.������%Type;
  v_���ʽ       ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  v_Temp           Varchar2(3000);
  v_ʱ���         ʱ���.ʱ���%Type;
  d_��鿪ʼʱ��   ʱ���.��ʼʱ��%Type;
  d_������ʱ��   ʱ���.��ֹʱ��%Type;
  n_�����¼id     �ٴ������¼.Id%Type;
  n_��ʱ����ʾ     Number(3);
  d_����ʱ��       Date;
Begin
  --��ȡ��ǰ��������
  Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);

  If �ѱ�_In Is Null Then
    Begin
      Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
      Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
    Exception
      When Others Then
        v_Err_Msg := '�޷�ȷ�����˷ѱ�����ȱʡ�ѱ��Ƿ���ȷ���ã�';
        Raise Err_Item;
    End;
  Else
    v_�ѱ� := �ѱ�_In;
    If Nvl(�������˷ѱ�_In, 0) = 1 Then
      Begin
        Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
      Exception
        When Others Then
          v_Err_Msg := 'û���ҵ���Ӧ�Ĳ��ˣ�';
          Raise Err_Item;
      End;
    End If;
  End If;

  If Nvl(������������_In, 0) = 1 Then
    Begin
      Update ������Ϣ Set ���� = ����_In Where ����id = ����id_In;
    Exception
      When Others Then
        v_Err_Msg := 'û���ҵ���Ӧ�Ĳ��ˣ�';
        Raise Err_Item;
    End;
  End If;

  If �����_In Is Not Null Then
    Begin
      Select Nvl(�����, 0) Into n_����� From ������Ϣ Where ����id = ����id_In;
    Exception
      When Others Then
        n_����� := 0;
    End;
    If n_����� = 0 Then
      Update ������Ϣ Set ����� = �����_In Where ����id = ����id_In;
    End If;
  End If;

  Begin
    Delete From �Һ����״̬
    Where ���� = �ű�_In And ���� = ����ʱ��_In And ��� = ����_In And ״̬ = 3 And ����Ա���� = ����Ա����_In;
  Exception
    When Others Then
      Null;
  End;
  v_Temp := zl_GetSysParameter(256);
  If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
    Null;
  Else
    Begin
      d_����ʱ�� := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
    If d_����ʱ�� Is Not Null Then
      If ����ʱ��_In > d_����ʱ�� Then
        v_Err_Msg := '��ǰ�Һŵķ���ʱ��' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�Ѿ������˳�����Ű�ģʽ,������ʹ�üƻ��Ű�ģʽ�Һ�!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    --�ҺŻ���ԤԼ����
    --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
    Select Count(*)
    Into n_Count
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ In (1, 3) And ��� = ���_In And NO = ���ݺ�_In;
    If n_Count <> 0 Then
      v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
      Raise Err_Item;
    End If;
  
    --��ȡ���㷽ʽ����
    Begin
      Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
    Exception
      When Others Then
        v_�ֽ� := '�ֽ�';
    End;
    Begin
      Select ���� Into v_�����ʻ� From ���㷽ʽ Where ���� = 3;
    Exception
      When Others Then
        v_�����ʻ� := '�����ʻ�';
    End;
  End If;
  
  Select Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null)
  Into v_����
  From Dual;

  --�ҺŻ�ȡ����
  Begin
    Select a.Id, a.��ſ���, Nvl(b.�޺���, 0), Nvl(b.��Լ��, 0)
    Into n_����id, n_��ſ���, n_�޺���, n_��Լ��
    From �ҺŰ��� A, �ҺŰ������� B
    Where a.Id = b.����id(+) And b.������Ŀ(+) = v_���� And a.���� = �ű�_In;
  
  Exception
    When Others Then
      n_����id := -1;
  End;

  --����ǲ����ѻ��ߺű�Ϊ��ʱ�����
  If Nvl(������_In, 0) = 0 Or �ű�_In Is Not Null Then
    If n_����id = -1 Then
      v_Err_Msg := '������Ӧ�ĹҺŰ�������,����';
      Raise Err_Item;
    End If;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
    --���Ȼ�ȡ�ƻ�
    Begin
      Select ID
      Into n_�ƻ�id
      From �ҺŰ��żƻ�
      Where ����id = n_����id And ���ʱ�� Is Not Null And
            Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) =
            (Select Max(a.��Чʱ��) As ��Ч
             From �ҺŰ��żƻ� A
             Where a.���ʱ�� Is Not Null And ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                   Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.����id = n_����id) And
            ����ʱ��_In Between Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
            Nvl(ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd'));
    
    Exception
      When Others Then
        n_�ƻ�id := 0;
    End;
    If Nvl(n_�ƻ�id, 0) <> 0 Then
      Begin
        --��ȡ�ƻ�������
        Select a.Id, a.��ſ���, Nvl(b.�޺���, 0) As �޺���, Nvl(b.��Լ��, 0) As ��Լ��
        Into n_�ƻ�id, n_��ſ���, n_�޺���, n_��Լ��
        From �ҺŰ��żƻ� A, �Һżƻ����� B
        Where a.���� = �ű�_In And a.Id = n_�ƻ�id And a.���ʱ�� Is Not Null And a.Id = b.�ƻ�id(+) And b.������Ŀ(+) = v_����;
      Exception
        When Others Then
          v_Err_Msg := '������Ӧ�ĹҺŰ��Ż�ƻ�����,����';
          Raise Err_Item;
      End;
    End If;
  End If;

  --��ȡ�Ƿ��ʱ��
  Begin
    If Nvl(n_�ƻ�id, 0) = 0 Then
      Select Count(Rownum) Into n_��ʱ�� From �ҺŰ���ʱ�� Where ���� = v_���� And ����id = n_����id And Rownum <= 1;
      Select Decode(To_Char(����ʱ��_In, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ���
      Where ID = n_����id;
    Else
      Select Count(Rownum) Into n_��ʱ�� From �Һżƻ�ʱ�� Where ���� = v_���� And �ƻ�id = n_�ƻ�id And Rownum <= 1;
      Select Decode(To_Char(����ʱ��_In, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ��żƻ�
      Where ID = n_�ƻ�id;
    End If;
  Exception
    When Others Then
      v_ʱ��� := Null;
  End;

  If v_ʱ��� Is Not Null And d_����ʱ�� Is Not Null And ���_In = 1 Then
    --����Ƿ��ģʽ�ҺŰ���
    Select To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
           To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
    Into d_��鿪ʼʱ��, d_������ʱ��
    From ʱ���
    Where ʱ��� = v_ʱ��� And վ�� Is Null And ���� Is Null;
    If d_��鿪ʼʱ�� > d_������ʱ�� Then
      d_������ʱ�� := d_������ʱ�� + 1;
    End If;
    If d_��鿪ʼʱ�� < d_����ʱ�� And d_������ʱ�� > d_����ʱ�� Then
      --��ȡ�����¼id
      Begin
        Select a.Id
        Into n_�����¼id
        From �ٴ������¼ A, �ٴ������Դ B
        Where a.��Դid = b.Id And b.���� = �ű�_In And �ϰ�ʱ�� = v_ʱ��� And ����ʱ��_In Between ��ʼʱ�� And ��ֹʱ��;
      Exception
        When Others Then
          n_�����¼id := Null;
      End;
    End If;
  End If;
  
  --�����ǰ��Ų�Ϊ0����Ҫ�����Ҫ�Ƿ�ռ�ã������ռ�ã���˳��ȥ��һ����ż�����ʱ��
  IF Nvl(����_In, 0) <> 0 then
    Begin
      IF �˺�����_In = 1 Then
        Select 1
        Into n_ռ��
        From �Һ����״̬
        Where ���� = �ű�_In And Trunc(����) = Trunc(����ʱ��_In) And ��� = ����_In And
              (״̬ In (1, 2) Or ״̬ In (3, 5) And (����Ա����_In <> v_��Ų���Ա Or v_������ <> v_��Ż�����));
      Else
        Select 1
        Into n_ռ��
        From �Һ����״̬
        Where ���� = �ű�_In And Trunc(����) = Trunc(����ʱ��_In) And ��� = ����_In And
              (״̬ In (1, 2, 4) Or ״̬ In (3, 5) And (����Ա����_In <> v_��Ų���Ա Or v_������ <> v_��Ż�����));
      End IF;
    Exception
      When Others Then
        n_ռ�� := 0;
    End;
  End IF;
  IF Nvl(n_ռ��, 0) = 1 And ���_In = 1 And n_��ʱ�� > 0 And Nvl(n_��ſ���, 0) = 1 Then
    Begin
      If Nvl(n_�ƻ�id, 0) = 0 Then
        Select ���, ʱ��ʱ��
        Into n_���, d_����ʱ��
        From (Select Nvl(���, 0) As ���,
                      To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'),
                               'yyyy-mm-dd hh24:mi:ss') As ʱ��ʱ��
               From �ҺŰ���ʱ��
               Where ����id = n_����id And ���� = v_���� And ��� > Nvl(����_In, 0) And
                     ��� Not In (Select ���
                                From �Һ����״̬
                                Where ���� = �ű�_In And ״̬ <> 0 And ״̬ <> Decode(�˺�����_In, 1, 4, 0) And
                                      ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60)
               Order By ���)
        Where Rownum < 2;
      Else
        Select ���, ʱ��ʱ��
        Into n_���, d_����ʱ��
        From (Select Nvl(���, 0) As ���,
                      To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'),
                               'yyyy-mm-dd hh24:mi:ss') As ʱ��ʱ��
               From �Һżƻ�ʱ��
               Where �ƻ�id = n_�ƻ�id And ���� = v_���� And ��� > Nvl(����_In, 0) And
                     ��� Not In (Select ���
                                From �Һ����״̬
                                Where ���� = �ű�_In And ״̬ <> 0 And ״̬ <> Decode(�˺�����_In, 1, 4, 0) And
                                      ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60)
               Order By ���)
        Where Rownum < 2;
      End IF;
    Exception
      When Others Then
        n_��� := Null;
        d_����ʱ�� := ����ʱ��_In;
    End;
  Else
    n_��� := ����_In;
    d_����ʱ�� := ����ʱ��_In;
  End IF;
  
  --��ʱ�ιҺ�ʱ�ж��Ƿ��ǹ��ڹҺ� Ҳ����׷�Ӻŵ���� Ŀǰֻ���ר�Һŷ�ʱ�ν��д���
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 And Nvl(n_��ſ���, 0) = 1 And ����_In Is Null And Nvl(��������_In, 0) = 0 Then
    --����ʱ��_in>Sysdate ����ʱ��>����ʱ��ʱ��--����_in is null
    Begin
      Select Max(To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_������ʱ��
      From �ҺŰ���ʱ��
      Where ����id = n_����id And ���� = v_���� And Nvl(��������, 0) <> 0;
      n_׷�Ӻ� := Case Sign(d_����ʱ�� - d_������ʱ��)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_׷�Ӻ� := 0;
    End;
  End If;
  d_ʱ��ʱ�� := d_����ʱ��;

  If ���_In = 1 And Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 Then
    --�Һ�ʱ��� �Ƿ����ʱ��,����ʱ��,��ʱ�ε�����������ȡ����
    Begin
      Select Nvl(���, 0),
             To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������)
      Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
      From �ҺŰ���ʱ��
      Where ����id = n_����id And ���� = v_���� And
            (���, ����id, ����) In (Select Nvl(Max(���), -1), ����id, ����
                               From �ҺŰ���ʱ��
                               Where ����id = n_����id And ���� = v_���� And
                                     Decode(��������_In + n_׷�Ӻ�, 0, To_Char(d_����ʱ��, 'hh24:mi'),
                                            To_Char(Nvl(��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')) =
                                     To_Char(Nvl(��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
                               Group By ����id, ����);
    Exception
      When Others Then
        n_ʱ����� := -1;
        n_��ʱ��   := 0;
        d_ʱ��ʱ�� := d_����ʱ��;
        n_ʱ���޺� := 0;
        n_ʱ����Լ := 0;
    End;
  End If;

  If ���_In = 1 And Nvl(ԤԼ�Һ�_In, 0) = 1 And n_��ʱ�� > 0 Then
    --ԤԼ��,ȡ�ƻ�
    Begin
      If Nvl(n_�ƻ�id, 0) = 0 Then
        --û�ƻ���Ч,ȡ���ŵ�����
        Select Nvl(���, 0),
               To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || ' ' || To_Char(c.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ʱ��ʱ��,
               ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �ҺŰ���ʱ�� C
        Where ����id = n_����id And ���� = v_���� And
              (���, ����id, ����) In
              (Select Nvl(Max(c.���), -1), ����id, ����
               From �ҺŰ���ʱ�� C
               Where ����id = n_����id And c.���� = v_���� And
                     Decode(��������_In, 1, To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(d_����ʱ��, 'hh24:mi')) =
                     To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By ����id, ����);
      Else
        --�мƻ���Чȡ�ƻ�
        --û��Ч�������ǴӹҺżƻ�ʱ�β�ѯ
        Select Nvl(���, -1),
               To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || ' ' || To_Char(c.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ʱ��ʱ��,
               ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �Һżƻ�ʱ�� C
        Where �ƻ�id = n_�ƻ�id And ���� = v_���� And
              (���, �ƻ�id, ����) In
              (Select Nvl(Max(c.���), -1), �ƻ�id, ����
               From �Һżƻ�ʱ�� C
               Where �ƻ�id = n_�ƻ�id And c.���� = v_���� And
                     Decode(��������_In, 1, To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(d_����ʱ��, 'hh24:mi')) =
                     To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By �ƻ�id, ����);
      End If;
    Exception
      When Others Then
        n_ʱ����� := -1;
        n_��ʱ��   := 0;
        d_ʱ��ʱ�� := d_����ʱ��;
        n_ʱ���޺� := 0;
        n_ʱ����Լ := 0;
    End;
  End If;

  If ���_In = 1 Then
  
    --��ȡ��ǰδʹ�õ����
    If Nvl(n_��ſ���, 0) = 1 And n_��ʱ�� = 0 Then
      --<��ſ��� δ����ʱ�� ��ȡ���õ�������,�Լ��Ѿ�ʹ�õ�����>
      Begin
        --������
        If �˺�����_In = 1 Then
          Select Max(Nvl(���, 0)), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Nvl(ԤԼ, 0))
          Into n_�������, n_��������, n_��Լ��
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(d_����ʱ��) And Trunc(d_����ʱ��) + 1 - 1 / 24 / 60 / 60 And ״̬ Not In (4, 5);
        Else
          Select Max(Nvl(���, 0)), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Nvl(ԤԼ, 0))
          Into n_�������, n_��������, n_��Լ��
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(d_����ʱ��) And Trunc(d_����ʱ��) + 1 - 1 / 24 / 60 / 60 And ״̬ <> 5;
        End If;
      Exception
        When Others Then
          n_������� := 0;
          n_�������� := 0;
      End;
      If n_��� Is Null Then
        n_��� := Nvl(n_�������, 0) + 1;
      End If;
      --<��ſ��� δ����ʱ�� ��ȡ���õ������� �Լ��Ѿ�ʹ�õ����� --end>
    
      --�ǼӺŵ������Ҫ����Ƿ񳬹�������
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�ʱ���
          --������ſ���δ��ʱ�� �ﵽ������
          If n_�޺��� <= n_�������� And n_�޺��� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(d_����ʱ��), 'yyyy-mm-dd ') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        Else
          --ԤԼʱ���
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
        
          If n_��Լ�� <= n_��Լ�� And n_��Լ�� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(d_����ʱ��), 'yyyy-mm-dd') || '�Ѵﵽ�����Լ����';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --������ſ���,δ��ʱ�� �Ӻ����   ������,����Ժ������������Ժ󲹳�
      End If;
    
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 0 Then
      --<--��ͨ�ŷ�ʱ�� ����ֻ��ԤԼһ�����-->
      If ��������_In = 0 Then
        --<����ԤԼ�Һ�-->
        Begin
          Select Count(0) As �ѹ���, Nvl(Sum(Decode(Nvl(Sign(a.���� - d_ʱ��ʱ��), 0), 0, 1, 0)), 0) As ��Լ��
          Into n_�ѹ���, n_��Լ��
          From �Һ����״̬ A
          Where a.���� = �ű�_In And ���� Between Trunc(d_����ʱ��) And Trunc(d_����ʱ��) + 1 - 1 / 24 / 60 / 60 And
                ״̬ Not In (4, 5);
        Exception
          When Others Then
            n_�ѹ��� := 0;
            n_��Լ�� := 0;
        End;
      
        n_ʱ����Լ := n_ʱ���޺�; --��ͨ�ŷ�ʱ�ε����,29��n_ʱ����Լʼ����0 �������⴦��
        --�����������
        If n_��Լ�� = 0 Then
          n_��Լ�� := n_�޺���;
        End If;
        If n_ʱ����Լ <= n_��Լ�� Or n_��Լ�� <= n_��Լ�� Then
          v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����Լ����';
          Raise Err_Item;
        End If;
        If n_�޺��� <= n_�ѹ��� Then
          v_Err_Msg := '�ű�' || �ű�_In || To_Char(d_����ʱ��, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
          Raise Err_Item;
        End If;
      End If;
    
      --û�дﵽʱ�ε��޺��� �����ڵ�ǰʱ������׷��
    
      --��ȡ����ҳ���������
      If �˺�����_In = 1 Then
        Select Nvl(Max(���), 0)
        Into n_�ҳ���������
        From �Һ����״̬ A
        Where a.���� = d_ʱ��ʱ�� And ���� = �ű�_In And ״̬ Not In (4, 5);
      Else
        Select Nvl(Max(���), 0)
        Into n_�ҳ���������
        From �Һ����״̬ A
        Where a.���� = d_ʱ��ʱ�� And ���� = �ű�_In And ״̬ <> 5;
      End If;
    
      --�������
      n_��� := RPad(Nvl(n_ʱ�����, 0), Length(n_�޺���) + Length(Nvl(n_ʱ�����, 0)), 0) + n_��Լ�� + 1;
      If n_��� <= Nvl(n_�ҳ���������, 0) Then
        n_��� := Nvl(n_�ҳ���������, 0) + 1;
      End If;
    
      --<--��ͨ�ŷ�ʱ��--End>
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 1 Then
      --<������ſ��� ����ʱ��
      --ר�Һŷ�ʱ��
      Begin
        If �˺�����_In = 1 Then
          Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(���� - d_ʱ��ʱ��), 0, 1, 0))
          Into n_�������, n_�ѹ���, n_��������
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(d_����ʱ��) And Trunc(d_����ʱ��) + 1 - 1 / 24 / 60 / 60 And ״̬ Not In (4, 5);
        Else
          Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(���� - d_ʱ��ʱ��), 0, 1, 0))
          Into n_�������, n_�ѹ���, n_��������
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(d_����ʱ��) And Trunc(d_����ʱ��) + 1 - 1 / 24 / 60 / 60 And ״̬ <> 5;
        End If;
      Exception
        When Others Then
          n_������� := 0;
          n_�������� := 0;
          n_�ѹ���   := 0;
      End;
    
      n_ʧЧ�� := 0;
      If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        n_ԤԼ��Чʱ�� := Zl_To_Number(zl_GetSysParameter('ԤԼ��Чʱ��', 1111));
        n_ʧԼ�Һ�     := Zl_To_Number(zl_GetSysParameter('ʧԼ���ڹҺ�', 1111));
        If Nvl(n_ԤԼ��Чʱ��, 0) <> 0 And Nvl(n_ʧԼ�Һ�, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_ԤԼ��Чʱ�� / 24 / 60) - ����), 1, 1, 0))
            Into n_ʧЧ��
            From �Һ����״̬
            Where ���� = �ű�_In And ���� Between Trunc(Sysdate) And Sysdate And Nvl(ԤԼ, 0) = 1 And ״̬ = 2;
          Exception
            When Others Then
              n_ʧЧ�� := 0;
          End;
        End If;
      End If;
    
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        
          --�Һ� ׷�Ӻ���ʱ�����ʱ���޺���
          If n_ʱ���޺� <= n_�������� And Nvl(n_׷�Ӻ�, 0) = 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� - n_ʧЧ�� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(d_����ʱ��, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        
        Else
          --�Һ�
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
          If n_��Լ�� <= n_�������� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(d_����ʱ��, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_��� Is Null Then
        --�������
        If Nvl(n_�������, 0) < Nvl(n_ʱ�����, 0) Then
          n_������� := Nvl(n_ʱ�����, 0);
        End If;
        n_��� := Nvl(n_�������, 0) + 1;
      End If;
    Elsif Nvl(n_��ʱ��, 0) = 0 And Nvl(n_��ſ���, 0) = 0 And Nvl(������_In, 0) = 0 And Nvl(�ű�_In, 0) > 0 Then
      ---<--��ͨ��  -->
      Begin
        Select �ѹ���, ��Լ��
        Into n_��������, n_��Լ��
        From ���˹ҺŻ���
        Where ���� = Trunc(d_����ʱ��) And ���� = �ű�_In;
      Exception
        When Others Then
          n_�������� := 0;
          n_��Լ��   := 0;
      End;
      If Nvl(��������_In, 0) = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�
          If Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(d_����ʱ��, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        Else
          --ԤԼ
          If (Nvl(n_��Լ��, 0) > 0 And Nvl(n_��Լ��, 0) >= Nvl(n_��Լ��, 0)) Or
             (Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0) Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(d_����ʱ��, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--��ͨ��  -->
    End If;
  End If;

  --���¹Һ����״̬
  If ���_In = 1 And Not n_��� Is Null Then
    If n_��ʱ�� = 1 Then
      d_���ʱ�� := d_����ʱ��;
    Else
      d_���ʱ�� := Trunc(d_����ʱ��);
    End If;
    --������ŵĴ���
    Begin
      Select ����Ա����, ������
      Into v_��Ų���Ա, v_��Ż�����
      From �Һ����״̬
      Where ״̬ = 5 And ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
      n_���� := 1;
    Exception
      When Others Then
        v_��Ų���Ա := Null;
        v_��Ż����� := Null;
        n_����       := 0;
    End;
    If n_���� = 0 Then
      Update �Һ����״̬
      Set ״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ԤԼ = Decode(ԤԼ����_In, 1, 1, 0), �Ǽ�ʱ�� = Sysdate
      Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_��� And ״̬ = 3 And ����Ա���� = ����Ա����_In;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_��ʱ��, 0) = 0 Or Nvl(ԤԼ�Һ�_In, 0) = 1 Or (Nvl(n_��ſ���, 0) = 0 And Nvl(����_In, 0) = 0) Then
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, ԤԼ, �Ǽ�ʱ��)
            Values
              (�ű�_In, d_���ʱ��, n_���, Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա����_In, Decode(ԤԼ����_In, 1, 1, 0), Sysdate);
            If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
              Update �Һ����״̬ Set ԤԼ = 1 Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
            End If;
          Elsif Nvl(n_��ʱ��, 0) > 0 Then
            --��ʱ�κ�ר�Һ� ʧԼ��ԤԼ������Һ�
            Update �Һ����״̬
            Set ״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա���� = ����Ա����_In, ԤԼ = Decode(ԤԼ����_In, 1, 1, 0), �Ǽ�ʱ�� = Sysdate
            Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_��� And ״̬ = 2;
            If Sql%NotFound Then
              Insert Into �Һ����״̬
                (����, ����, ���, ״̬, ����Ա����, ԤԼ, �Ǽ�ʱ��)
              Values
                (�ű�_In, d_���ʱ��, n_���, Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա����_In, Decode(ԤԼ����_In, 1, 1, 0), Sysdate);
            End If;
            If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
              Update �Һ����״̬ Set ԤԼ = 1 Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.';
            Raise Err_Item;
        End;
      End If;
    Else
      If ����Ա����_In <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
        v_Err_Msg := '���' || n_��� || '�ѱ�������' || v_������ || '����,������ѡ��һ�����.';
        Raise Err_Item;
      Else
        Update �Һ����״̬
        Set ״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ԤԼ = Decode(ԤԼ����_In, 1, 1, 0), �Ǽ�ʱ�� = Sysdate
        Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_��� And ״̬ = 5 And ����Ա���� = ����Ա����_In And ������ = v_������;
        If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
          Update �Һ����״̬ Set ԤԼ = 1 Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
        End If;
      End If;
    End If;
  End If;

  If n_�����¼id Is Not Null Then
    Update �ٴ�������ſ���
    Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա���� = ����Ա����_In
    Where ��¼id = n_�����¼id And ��� = n_���;
    If ԤԼ�Һ�_In = 1 Then
      Update �ٴ������¼ Set ��Լ�� = ��Լ�� + 1 Where ID = n_�����¼id;
    Else
      If ԤԼ����_In = 1 Then
        Update �ٴ������¼
        Set ��Լ�� = ��Լ�� + 1, �ѹ��� = �ѹ��� + 1, �����ѽ��� = �����ѽ��� + 1
        Where ID = n_�����¼id;
      Else
        Update �ٴ������¼ Set �ѹ��� = �ѹ��� + 1 Where ID = n_�����¼id;
      End If;
    End If;
  End If;

  --�������˹Һŷ���(���ܵ����ǻ������������)
  Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual; --Ӧ��ͨ������õ�

  Insert Into ������ü�¼
    (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id, �շ����,
     ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����,
     ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
  Values
    (n_����id, 4, Decode(ԤԼ�Һ�_In, 1, 0, 1), ���_In, Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), ��������_In, ���ݺ�_In, Ʊ�ݺ�_In, 1, ����_In,
     ������_In, Decode(ԤԼ�Һ�_In, 1, To_Char(n_���), ����_In), Decode(����id_In, 0, Null, ����id_In),
     Decode(�����_In, 0, Null, �����_In), ���ʽ_In, ����_In, Decode(����_In, Null, Null, �Ա�_In), Decode(����_In, Null, Null, ����_In),
     v_�ѱ�, ���˿���id_In, �շ����_In, �ű�_In, �շ�ϸĿid_In, ������Ŀid_In, �վݷ�Ŀ_In, 1, ����_In, ��׼����_In, Ӧ�ս��_In, ʵ�ս��_In,
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��_In)),
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In)), Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ��������id_In,
     ����Ա����_In, Decode(ԤԼ�Һ�_In, 1, ����Ա����_In, Null), ִ�в���id_In, ҽ������_In, ����Ա���_In, ����Ա����_In, d_����ʱ��, �Ǽ�ʱ��_In, ���մ���id_In,
     ������Ŀ��_In, ���ձ���_In, ͳ����_In, Decode(�շѵ�_In, Null, ժҪ_In, '����:' || �շѵ�_In), ԤԼ��ʽ_In, Decode(ԤԼ�Һ�_In, 1, Null, n_��id));

  --���ܽ��㵽����Ԥ����¼
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 0 Then
  
    If (Nvl(�ֽ�֧��_In, 0) <> 0 Or (Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0)) And ���_In = 1 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
         ��������)
      Values
        (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(���㷽ʽ_In, v_�ֽ�), Nvl(�ֽ�֧��_In, 0), �Ǽ�ʱ��_In,
         ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4);
    
      If Nvl(���㿨���_In, 0) <> 0 And Nvl(�ֽ�֧��_In, 0) <> 0 Then
        Zl_���˿������¼_֧��(���㿨���_In, ����_In, 0, �ֽ�֧��_In, n_Ԥ��id, ����Ա���_In, ����Ա����_In, �Ǽ�ʱ��_In);
      End If;
    End If;
  
    --����ҽ���Һ�
    If Nvl(����֧��_In, 0) <> 0 And ���_In = 1 Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), v_�����ʻ�, ����֧��_In, �Ǽ�ʱ��_In, ����Ա���_In,
         ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id, 4);
    End If;
  
    --���ھ��￨ͨ��Ԥ����Һ�
    If Nvl(Ԥ��֧��_In, 0) <> 0 And ���_In = 1 Then
      n_Ԥ����� := Ԥ��֧��_In;
      For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
        n_��ǰ��� := Case
                    When r_Deposit.��� - n_Ԥ����� < 0 Then
                     r_Deposit.���
                    Else
                     n_Ԥ�����
                  End;
      
        If r_Deposit.����id = 0 Then
          --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
          Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 4 Where ID = r_Deposit.ԭԤ��id;
        
        End If;
        --���ϴ�ʣ���
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
           ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                 �Ǽ�ʱ��_In, ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
      
        --���²���Ԥ�����
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
        Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2);
        --����Ƿ��Ѿ�������
        If r_Deposit.��� < n_Ԥ����� Then
          n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
        Else
          n_Ԥ����� := 0;
        End If;
      
        If n_Ԥ����� = 0 Then
          Exit;
        End If;
      End Loop;
      If n_Ԥ����� > 0 Then
        v_Err_Msg := 'Ԥ���಻��֧������֧�����,���ܼ���������';
        Raise Err_Item;
      
      End If;
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
    End If;
  
    --��ػ��ܱ�Ĵ���
    --��Ա�ɿ����
    If ���_In = 1 And Nvl(���½������_In, 1) = 1 Then
      If Nvl(�ֽ�֧��_In, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + �ֽ�֧��_In
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�)
        Returning ��� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (����Ա����_In, Nvl(���㷽ʽ_In, v_�ֽ�), 1, �ֽ�֧��_In);
          n_����ֵ := �ֽ�֧��_In;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�) And ���� = 1 And Nvl(���, 0) = 0;
        End If;
      End If;
    
      If Nvl(����֧��_In, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + ����֧��_In
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
        Returning ��� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
          n_����ֵ := ����֧��_In;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_�����ʻ� And Nvl(���, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --���˹ҺŻ���(ֻ����һ��,�ҵ�����ȡ�����Ѳ�����)
  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    --���˱��ξ���(�Է���ʱ��Ϊ׼)
    If Nvl(����id_In, 0) <> 0 And ���_In = 1 Then
      Update ������Ϣ Set ����ʱ�� = d_����ʱ��, ����״̬ = 1, �������� = ����_In Where ����id = ����id_In;
    End If;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
  
    --�������
    Update �������
    Set ������� = Nvl(�������, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (����id_In, 1, 1, Nvl(ʵ�ս��_In, 0), 0);
    End If;
  
    --����δ�����
    Update ����δ�����
    Set ��� = Nvl(���, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(���˿���id_In, 0) And
          Nvl(��������id, 0) = Nvl(��������id_In, 0) And Nvl(ִ�в���id, 0) = Nvl(ִ�в���id_In, 0) And ������Ŀid + 0 = ������Ŀid_In And
          ��Դ;�� + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (����id_In, Null, Null, ���˿���id_In, ��������id_In, ִ�в���id_In, ������Ŀid_In, 1, Nvl(ʵ�ս��_In, 0));
    End If;
  End If;

  --���˹Һż�¼
  If �ű�_In Is Not Null And ���_In = 1 Then
    --And Nvl(ԤԼ�Һ�_In, 0) = 0
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
    Begin
      Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
    Exception
      When Others Then
        v_���ʽ := Null;
    End;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���,
       ����Ա����, ����, ����, ����, ԤԼ, ԤԼ��ʽ, ժҪ, ������ˮ��, ����˵��, ������λ, ����ʱ��, ������, ԤԼ����Ա, ԤԼ����Ա���, ����, ҽ�Ƹ��ʽ, �շѵ�)
    Values
      (n_�Һ�id, ���ݺ�_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, 2, 1), 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, �ű�_In, ����_In, ����_In,
       Null, ִ�в���id_In, ҽ������_In, 0, Null, �Ǽ�ʱ��_In, d_����ʱ��, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, d_����ʱ��, Null), ����Ա���_In,
       ����Ա����_In, ����_In, n_���, ����_In, Decode(ԤԼ����_In, 1, 1, 0), ԤԼ��ʽ_In, ժҪ_In, ������ˮ��_In, ����˵��_In, ������λ_In,
       Decode(Nvl(ԤԼ�Һ�_In, 0), 0, �Ǽ�ʱ��_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 0, ����Ա����_In, Null),
       Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա����_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա���_In, Null),
       Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, �շѵ�_In);
    If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
      Update ���˹Һż�¼
      Set ԤԼ = 1, ԤԼʱ�� = d_����ʱ��, ԤԼ����Ա = ����Ա����_In, ԤԼ����Ա��� = ����Ա���_In
      Where ID = n_�Һ�id;
    End If;
    n_ԤԼ���ɶ��� := 0;
    If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
      n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    End If;
  
    --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
    If Nvl(���ɶ���_In, 0) <> 0 And Nvl(ԤԼ�Һ�_In, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
        n_��ʱ����ʾ := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
        If Nvl(ԤԼ�Һ�_In, 0) = 1 And n_��ʱ����ʾ = 1 And n_��ʱ�� = 1 Then
          n_��ʱ����ʾ := 1;
        Else
          n_��ʱ����ʾ := Null;
        End If;
      
        --��������
        --.����ִ�в��š� �ķ�ʽ���ɶ���
        v_�������� := ִ�в���id_In;
        v_�ŶӺ��� := Zlgetnextqueue(ִ�в���id_In, n_�Һ�id, �ű�_In || '|' || n_���);
        v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, �ű�_In, n_���, d_����ʱ��);
        --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
        Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, ִ�в���id_In, v_�ŶӺ���, Null, ����_In, ����id_In, ����_In, ҽ������_In, d_�Ŷ�ʱ��, ԤԼ��ʽ_In,
                         n_��ʱ����ʾ, v_�Ŷ����);
      
        --�Һ������Ŷ�
        If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
          Update ���˹Һż�¼ Set ��¼��־ = 1 Where ID = n_�Һ�id;
        End If;
      End If;
    End If;
  End If;
  --���˵�����Ϣ
  If ����id_In Is Not Null And ���_In = 1 Then
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(����ģʽ, 0) Into n_����ģʽ From ������Ϣ Where ����id = ����id_In;
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ĳ�����Ϣ,������Һ�';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If n_����ģʽ = 1 And Nvl(����ģʽ_In, 0) = 0 Then
        --�����Ѿ���"�����ƺ�����",������"�Ƚ�������Ƶ�",�����Ƿ����δ������
        Select Count(1)
        Into n_Count
        From ����δ�����
        Where ����id = ����id_In And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --����δ�������ݣ������Ƚ���������ִ��
          v_Err_Msg := '��ǰ���˵ľ���ģʽΪ�����ƺ�����Ҵ���δ����ã�����������ò��˵ľ���ģʽ,������ȶ�δ����ý��ʺ��ٹҺŻ򲻵������˵ľ���ģʽ!';
          Raise Err_Item;
        End If;
        --���
        --δ����ҽ��ҵ��ģ�����ʱ�͹Һŵ�,��Ҫ��֤ͬһ�εľ���ģʽ��һ����(�����Ѿ���飬�����ٴ���)
      End If;
      Update ������Ϣ Set ����ģʽ = ����ģʽ_In Where ����id = ����id_In;
    End If;
  
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Nvl(��Ժ, 0) = 0 And Exists
     (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
  
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = Sysdate
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, Sysdate) >= Sysdate;
    End If;
  End If;
  If ���_In = 1 Then
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 1, n_�Һ�id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_�Һ�id, ���ݺ�_In);
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_Insert;
/

--125588:����,2018-05-17,�������Ŀ��ʵ������Ϊ0������
Create Or Replace Procedure Zl_�����շ���¼_�ɱ��۵���(����id_In In ҩƷ�շ���¼.ҩƷid%Type) As
  v_No         ҩƷ�շ���¼.No%Type;
  v_Ӧ��id     Ӧ����¼.Id%Type; --Ӧ����¼��ID 
  v_Ӧ�����ݺ� Ӧ����¼.No%Type;
  d_����ʱ��   Date;
  n_���       Number(8);
  n_�ⷿid     ҩƷ�շ���¼.�ⷿid%Type;
  n_������id ҩƷ�շ���¼.������id%Type;
  n_���ϵ��   ҩƷ�շ���¼.���ϵ��%Type;
  n_�շ�id     ҩƷ�շ���¼.Id%Type;
  n_������     ҩƷ�շ���¼.���۽��%Type;
  n_ԭ�ɱ���   ҩƷ�շ���¼.�ɱ���%Type;
  n_�³ɱ���   ҩƷ�շ���¼.�ɱ���%Type;
  n_ƽ���ɱ��� ҩƷ���.ƽ���ɱ���%Type;
  v_����id     �ɱ��۵�����Ϣ.Id%Type;
  v_���ۻ��ܺ� �ɱ��۵�����Ϣ.���ۻ��ܺ�%Type;
  n_Count      Number(1) := 0;

  Cursor c_Stock Is --��ǰ��� 
    Select �ϴι�Ӧ��id, a.�ⷿid, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, a.�ϴ�����, a.Ч��, a.�ϴβ���, a.���Ч��,
           Decode(Sign(Nvl(a.����, 0)), 1, a.�ϴβɹ���, a.ƽ���ɱ���) As ԭ�ɱ���
    From ҩƷ��� A
    Where a.���� = 1 And Nvl(a.ʵ������, 0) <> 0 And a.ҩƷid = ����id_In
    Order By a.�ⷿid;

  v_Stock c_Stock%RowType;
Begin
  d_����ʱ�� := Sysdate;
  n_�ⷿid   := 0;

  --�ж��Ƿ�����޿����� 
  Begin
    Select ID, �³ɱ���, ���ۻ��ܺ�
    Into v_����id, n_�³ɱ���, v_���ۻ��ܺ�
    From �ɱ��۵�����Ϣ
    Where ִ������ Is Null And Nvl(�ⷿid, 0) = 0 And ҩƷid = ����id_In;
  Exception
    When Others Then
      v_����id   := 0;
      n_�³ɱ��� := Null;
  End;

  --�޿����� 
  If v_����id > 0 Then
    --���ݵ�ǰ������²���������Ϣ 
    For v_Stock In c_Stock Loop
      Zl_���ϳɱ�����_Insert(v_Stock.�ϴι�Ӧ��id, v_Stock.�ⷿid, v_Stock.����id, v_Stock.����, v_Stock.�ϴ�����, v_Stock.ԭ�ɱ���, n_�³ɱ���,
                       Null, Null, 0, 0, v_���ۻ��ܺ�);
      n_Count := n_Count + 1;
    End Loop;
  
    If n_Count > 0 Then
      --�����ǰ�п���¼����ɾ���޿����ۼ�¼ 
      Delete �ɱ��۵�����Ϣ Where ID = v_����id;
    Else
      Update �ɱ��۵�����Ϣ Set ִ������ = d_����ʱ�� Where ID = v_����id;
    
      Update �������� Set �ɱ��� = n_�³ɱ��� Where ����id = ����id_In And �ɱ��� <> n_�³ɱ���;
    End If;
  End If;

  --ȡ����۵�����������ID 
  Select b.Id, b.ϵ��
  Into n_������id, n_���ϵ��
  From ҩƷ�������� A, ҩƷ������ B
  Where a.���id = b.Id And a.���� = 33 And Rownum < 2;

  For c_�ɱ����� In (Select a.�ⷿid, a.ҩƷid As ����id, Nvl(a.����, 0) ����, a.�ϴι�Ӧ��id, a.ʵ������, a.ʵ�ʽ��, a.ʵ�ʲ��, a.�ϴβ��� As ����,
                        a.�ϴ����� As ����, a.���Ч��, a.Ч��, a.�ϴ��������� As ��������, a.��׼�ĺ�, Nvl(a.ƽ���ɱ���, 0) As ԭ�ɱ���, b.�³ɱ���, b.��Ʊ��,
                        b.��Ʊ����, b.��Ʊ���, Nvl(a.�ϴβɹ���, 0) As �ϴβɹ���, b.Id As ����id
                 From ҩƷ��� A, �ɱ��۵�����Ϣ B
                 Where a.ҩƷid = b.ҩƷid And Nvl(a.�ϴι�Ӧ��id, 0) = Nvl(b.��ҩ��λid, 0) And a.�ⷿid = b.�ⷿid And
                       Nvl(a.����, 0) = Nvl(b.����, 0) And a.���� = 1 And b.ִ������ Is Null And a.ҩƷid = ����id_In
                 Order By a.�ⷿid) Loop
    If n_�ⷿid <> c_�ɱ�����.�ⷿid Then
      n_���   := 1;
      n_�ⷿid := c_�ɱ�����.�ⷿid;
      v_No     := Nextno(71, n_�ⷿid);
    Else
      n_��� := n_��� + 1;
    End If;
  
    Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
    /*    If Nvl(c_�ɱ�����.ʵ������, 0) = 0 And Nvl(c_�ɱ�����.ʵ�ʽ��, 0) = 0 And Nvl(c_�ɱ�����.ʵ�ʲ��, 0) = 0 Then
      --����,����۶�Ϊ0�����ʾ��������¿���������������ĵ��ݣ��˵��ݻ�û����ˣ����ֻ��Ҫ���µ�����Ϣ������������
      Update �������� Set �ɱ��� = c_�ɱ�����.�³ɱ��� Where ����id = c_�ɱ�����.����id;
    
      Update �ɱ��۵�����Ϣ
      Set �շ�id = n_�շ�id, ִ������ = d_����ʱ��, Ч�� = c_�ɱ�����.Ч��, ���Ч�� = c_�ɱ�����.���Ч��, ���� = c_�ɱ�����.����, ���� = c_�ɱ�����.����
      Where ID = c_�ɱ�����.����id;
    Elsif Nvl(c_�ɱ�����.ʵ������, 0) = 0 And (Nvl(c_�ɱ�����.ʵ�ʽ��, 0) <> 0 Or Nvl(c_�ɱ�����.ʵ�ʲ��, 0) <> 0) Then
      --����=0 ������<>0ʱֻ���¿����ж�Ӧ��ƽ���ɱ��ۺ����Ա��гɱ��ۣ��������ɱ����������ݵ��ǲ�۲�=0��ֻ��¼���³ɱ��� 
      --�������ۼ�¼��ֻ��¼���³ɱ��� 
      Insert Into ҩƷ�շ���¼
        (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ��ҩ��λid, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ���ۼ�, �ɱ���, ���, ժҪ, ������, ��������, �����,
         �������, ��������, ��׼�ĺ�, ����, ��ҩ��ʽ, ����)
      Values
        (n_�շ�id, 1, 18, v_No, n_���, c_�ɱ�����.�ⷿid, n_������id, c_�ɱ�����.�ϴι�Ӧ��id, n_���ϵ��, c_�ɱ�����.����id, c_�ɱ�����.����, c_�ɱ�����.����,
         c_�ɱ�����.����, c_�ɱ�����.Ч��, 0, c_�ɱ�����.ʵ�ʽ��, c_�ɱ�����.ʵ�ʲ��, 0, '�������ϳɱ��۵���', Zl_Username, d_����ʱ��, Zl_Username, d_����ʱ��,
         c_�ɱ�����.��������, c_�ɱ�����.��׼�ĺ�, c_�ɱ�����.�³ɱ���, 1, c_�ɱ�����.ԭ�ɱ���);
    
      Zl_δ��ҩƷ��¼_Insert(n_�շ�id);
      --���¿�� 
      Update ҩƷ���
      Set ƽ���ɱ��� = c_�ɱ�����.�³ɱ���, �ϴβɹ��� = c_�ɱ�����.�³ɱ���
      Where �ⷿid = c_�ɱ�����.�ⷿid And ҩƷid = c_�ɱ�����.����id And Nvl(����, 0) = c_�ɱ�����.���� And ���� = 1;
      Update �������� Set �ɱ��� = c_�ɱ�����.�³ɱ��� Where ����id = c_�ɱ�����.����id;
    
      Update �ɱ��۵�����Ϣ
      Set �շ�id = n_�շ�id, ִ������ = d_����ʱ��, Ч�� = c_�ɱ�����.Ч��, ���Ч�� = c_�ɱ�����.���Ч��, ���� = c_�ɱ�����.����, ���� = c_�ɱ�����.����
      Where ID = c_�ɱ�����.����id;*/
    --Else
    --������Ӧ�Ŀ��:ԭ�ɱ����-ʵ�³ɱ���� 
    n_������   := Round(c_�ɱ�����.ԭ�ɱ��� * c_�ɱ�����.ʵ������, 2) - Round(c_�ɱ�����.�³ɱ��� * c_�ɱ�����.ʵ������, 2);
    n_ԭ�ɱ��� := c_�ɱ�����.ԭ�ɱ���;
  
    If n_ԭ�ɱ��� <= 0 Then
      n_ԭ�ɱ��� := c_�ɱ�����.�ϴβɹ���;
    End If;
  
    --Ŀǰ���շ���¼��Ӧ: 
    -- ����--> ԭ�ɱ��� 
    -- ����-->�³ɱ��� 
    -- ��д����-->���ʵ������ 
    -- ���ۼ�-->���ʵ�ʽ�� 
    -- �ɱ���-->���ʵ�ʲ�� 
    -- ���-->���ε����� 
  
    Insert Into ҩƷ�շ���¼
      (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ��ҩ��λid, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ���ۼ�, �ɱ���, ���, ժҪ, ������, ��������, �����,
       �������, ��������, ��׼�ĺ�, ����, ��ҩ��ʽ, ����)
    Values
      (n_�շ�id, 1, 18, v_No, n_���, c_�ɱ�����.�ⷿid, n_������id, c_�ɱ�����.�ϴι�Ӧ��id, n_���ϵ��, c_�ɱ�����.����id, c_�ɱ�����.����, c_�ɱ�����.����,
       c_�ɱ�����.����, c_�ɱ�����.Ч��, c_�ɱ�����.ʵ������, c_�ɱ�����.ʵ�ʽ��, c_�ɱ�����.ʵ�ʲ��, n_������, '�������ϳɱ��۵���', Zl_Username, d_����ʱ��, Zl_Username,
       d_����ʱ��, c_�ɱ�����.��������, c_�ɱ�����.��׼�ĺ�, c_�ɱ�����.�³ɱ���, 1, n_ԭ�ɱ���);
  
    Zl_δ��ҩƷ��¼_Insert(n_�շ�id);
    --���¿�� 
    Update ҩƷ���
    Set ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_������
    Where �ⷿid = c_�ɱ�����.�ⷿid And ҩƷid = c_�ɱ�����.����id And Nvl(����, 0) = Nvl(c_�ɱ�����.����, 0) And ���� = 1;
  
    If Sql%NotFound Then
      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, ����, ʵ�ʲ��, �ϴ�����, Ч��, �ϴβ���, �ϴι�Ӧ��id, �ϴ���������, ��׼�ĺ�, ���Ч��)
      Values
        (c_�ɱ�����.�ⷿid, c_�ɱ�����.����id, c_�ɱ�����.����, 1, n_������, c_�ɱ�����.����, c_�ɱ�����.Ч��, c_�ɱ�����.����, c_�ɱ�����.�ϴι�Ӧ��id, c_�ɱ�����.��������,
         c_�ɱ�����.��׼�ĺ�, c_�ɱ�����.���Ч��);
    End If;
  
    Update ҩƷ���
    Set �ϴβɹ��� = c_�ɱ�����.�³ɱ���
    Where ҩƷid = c_�ɱ�����.����id And �ϴβɹ��� <> c_�ɱ�����.�³ɱ���;
  
    Update ��������
    Set �ɱ��� = c_�ɱ�����.�³ɱ���
    Where ����id = c_�ɱ�����.����id And �ɱ��� <> c_�ɱ�����.�³ɱ���;
  
    --���¼�������е�ƽ���ɱ��� 
    Update ҩƷ���
    Set ƽ���ɱ��� = Decode(Nvl(����, 0), 0, Decode((ʵ�ʽ�� - ʵ�ʲ��) / ʵ������, 0, �ϴβɹ���, (ʵ�ʽ�� - ʵ�ʲ��) / ʵ������), �ϴβɹ���)
    Where ҩƷid = c_�ɱ�����.����id And Nvl(����, 0) = Nvl(c_�ɱ�����.����, 0) And �ⷿid = c_�ɱ�����.�ⷿid And ���� = 1 And Nvl(ʵ������, 0) <> 0;
    If Sql%NotFound Then
      Select �ɱ��� Into n_ƽ���ɱ��� From �������� Where ����id = c_�ɱ�����.����id;
      Update ҩƷ���
      Set ƽ���ɱ��� = n_ƽ���ɱ���
      Where ҩƷid = c_�ɱ�����.����id And �ⷿid = c_�ɱ�����.�ⷿid And Nvl(����, 0) = Nvl(c_�ɱ�����.����, 0) And ���� = 1;
    End If;
  
    --���³ɱ��۵�����Ϣ 
    Update �ɱ��۵�����Ϣ
    Set �շ�id = n_�շ�id, ִ������ = d_����ʱ��, ԭ�ɱ��� = n_ԭ�ɱ���, Ч�� = c_�ɱ�����.Ч��, ���Ч�� = c_�ɱ�����.���Ч��, ���� = c_�ɱ�����.����, ���� = c_�ɱ�����.����
    Where ID = c_�ɱ�����.����id;
    --End If;
  
    --��Ϣ����
    b_Message.Zlhis_Drug_010(c_�ɱ�����.����id);
  End Loop;

  --����Ӧ����¼ 
  For c_Ӧ�� In (Select Distinct a.��ҩ��λid, a.ҩƷid, a.��Ʊ��, a.��Ʊ����, a.��Ʊ���, b.����, b.���㵥λ, b.���
               From �ɱ��۵�����Ϣ A, �շ���ĿĿ¼ B
               Where a.ҩƷid = b.Id And Nvl(a.Ӧ����䶯, 0) = 1 And Nvl(a.��ҩ��λid, 0) <> 0 And a.ҩƷid = ����id_In
               Order By a.��ҩ��λid) Loop
  
    v_Ӧ�����ݺ� := Nextno(67);
  
    Select Ӧ����¼_Id.Nextval Into v_Ӧ��id From Dual;
  
    Insert Into Ӧ����¼
      (ID, ��¼����, ��¼״̬, ��λid, NO, ϵͳ��ʶ, ��Ʊ��, ��Ʊ����, ��Ʊ���, Ʒ��, ���, ������, ��������, �����, �������, ժҪ)
    Values
      (v_Ӧ��id, 1, 1, c_Ӧ��.��ҩ��λid, v_Ӧ�����ݺ�, 5, c_Ӧ��.��Ʊ��, c_Ӧ��.��Ʊ����, c_Ӧ��.��Ʊ���, c_Ӧ��.����, c_Ӧ��.���, Zl_Username, d_����ʱ��,
       Zl_Username, d_����ʱ��, '�ɱ��۵����Զ�����Ӧ����䶯��¼');
  
    If Nvl(c_Ӧ��.��ҩ��λid, 0) <> 0 Then
      Update Ӧ����� Set ��� = Nvl(���, 0) + Nvl(c_Ӧ��.��Ʊ���, 0) Where ��λid = c_Ӧ��.��ҩ��λid And ���� = 1;
      If Sql%NotFound Then
        Insert Into Ӧ����� (��λid, ����, ���) Values (c_Ӧ��.��ҩ��λid, 1, Nvl(c_Ӧ��.��Ʊ���, 0));
      End If;
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շ���¼_�ɱ��۵���;
/

--125781:����,2018-05-17,����oracle����Zl_Third_Getvisitinfo,ɾ�����õ�������
CREATE OR REPLACE Procedure Zl_Third_Getvisitinfo
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --����:���ݹҺŵ��Ż�ȡ�ôξ�������(ҽ��Ϊ��Ҫ��ʾ) 
  --���:Xml_In: 
  --<IN> 
  --    <GHDH>�Һŵ���</GHDH> 
  --    <JSKLB>���㿨���</JSKLB> 
  --    <MXGL>��ϸ����</MXGL> 0-������,��ϸ�������� 1-����,��ϸ����������,Ĭ��Ϊ1 
  --</IN> 
  --����:Xml_Out 
  --<OUTPUT> 
  --  <GH> 
  --     <GHDH>�Һŵ���</GHDH> //���β�ѯ�ĹҺŵ��� 
  --     <YYSJ>ԤԼʱ��</YYSJ> //yyyy-mm-dd hh24:mi:ss 
  --     <JZSJ></JZSJ>      //ʵ�ʾ���ʱ�� 
  --     <DJH></DJH>        //���ݺ� 
  --     <JE></JE>          //��� 
  --     <DJLX></DJLX>      //��������,1-�շѵ���4-�Һŵ� 
  --     <KDSJ></KDSJ>      //����ʱ�� 
  --     <JKFS></JKFS>      //�ɿʽ,0-�ҺŻ�ԤԼ�ɿ�;1-ԤԼ���ɿ� 
  --     <ZFZT></ZFZT>  //֧��״̬,0-��֧����1-��֧����2-���˷� 
  --     <SFJSK></SFJSK>    //�Ƿ���㿨֧����0-��1-�� 
  --  </GH> 
  --  <YZLIST> 
  --     <YZ>                   //ҽ��������HIS����ʾ��������ͬ 
  --        <YZID><YZID>        //ҽ��ID��������ҽ��ID 
  --        <YZLX><YZLX>        //ҽ������,�紦������顢���� 
  --         <YZMC></YZMC>        //ҽ������ 
  --        <ZXKS></ZXKS>       //ִ�п��� 
  --        <ZXKSID></ZXKSID>   //ִ�п���ID 
  --        <FYCK></FYCK>       //��ҩ���� 
  --        <YZMX> 
  --           <MX> 
  --              <YZNR></YZNR>        //ҽ������ 
  --              <ZXZT></ZXZT>        //ҽ��ִ��״̬ 
  --              <SFFY>�Ƿ�ҩ</SFFY> // 0-�� ��1-�� 
  --              <GG>���</GG> 
  --              <SL>����</SL> 
  --              <DW>���㵥λ</DW> 
  --              <BZDJ>��׼����</BZDJ> 
  --              <YSJE>Ӧ�ս��</YSJE> 
  --              <SSJE>ʵ�ս��</SSJE> 
  --           </MX> 
  --           <MX/> 
  --        </YZMX> 
  --        <BG></BG>                   //�Ƿ��ѳ����棬�Ƿ�ǩ�� 
  --        <BGLY></BGLY>               //�Ƿ������Ŀ,1-Ժ����Ŀ��2-�����Ŀ 
  --        <BGLYSM></BGLYSM>           //�����Ŀ˵�� 
  --        <JZBG></JZBG>                //��ֹ��ʾ���档0-����1-��ֹ 
  --        <JZTS></JZTS>                 //��ʾ���֡����ڽ�ֹ�鿴�ı��棬�ɷ���������ʾ���˵���Ϣ 
  --        <BLID></BLID>              //����ID�����<BG>�ֶ�Ϊ1����ֵ��Ϊ�� 
  --        <DJLIST> 
  --           <DJ>                //���õ�����Ϣ 
  --              <DJH></DJH>      //���õ��ݺ� 
  --              <DJLX></DJLX>    //�������� 
  --              <JE></JE>        //�����ܽ�� 
  --              <KDSJ></KDSJ>    //����ʱ�� 
  --              <ZFZT></ZFZT>    //֧��״̬,0-��֧����1-��֧����2-���˷�,3-�˷�������,4-���ͨ��,5-���δͨ�� 
  --              <SHSM></SHSM>    //���˵��,���δͨ��ԭ�� 
  --              <SFJSK></SFJSK>  //�Ƿ���㿨֧����0-��1-�� 
  --           </DJ> 
  --           <DJ/> 
  --        </DJLIST> 
  --     </YZ> 
  --  </YZLIST> 
  --    <ERROR><MSG></MSG></ERROR>                      //������󷵻� 
  --</OUTPUT> 

  -------------------------------------------------------------------------------------------------- 
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  x_Templet Xmltype; --ģ��XML 

  v_�����   Varchar2(100);
  n_�����id Number(18);
  v_�Һŵ�   Varchar2(10);
  v_�ŶӺ��� Varchar2(10);
  n_Temp     Number(18);
  v_�������� �ŶӽкŶ���.��������%Type;

  n_Count Number(18);

  v_Temp       Varchar2(32767); --��ʱXML 
  v_����       Varchar2(32767);
  v_No         Varchar2(50);
  n_Add_Djlist Number(1); --�Ƿ�������DJLIST�� 
  n_����       Number(2);
  n_��ҽ��id   Number(18);
  n_����ҽ��   Number(8);
  n_ִ�п���id Number(18);
  v_ִ�п���   Varchar2(50);
  n_�˿���   ����Ԥ����¼.��Ԥ��%Type;
  n_��ϸ����   Number(3);
  n_�˷�״̬   �����˷�����.״̬%Type;
  v_����ԭ��   �����˷�����.����ԭ��%Type;
  v_���ԭ��   �����˷�����.���ԭ��%Type;
  v_��ҩ����   ������ü�¼.��ҩ����%Type;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/MXGL')
  Into v_�Һŵ�, v_�����, n_��ϸ����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If v_�Һŵ� Is Null Then
    v_Err_Msg := '�����ҵ�ָ���ĹҺŵ���(��ǰ�Һŵ���Ϊ��)';
    Raise Err_Item;
  End If;
  If n_��ϸ���� Is Null Then
    n_��ϸ���� := 1;
  End If;
  n_Add_Djlist := 0;

  v_Err_Msg := Null;
  If v_����� Is Not Null Then
    Begin
      n_�����id := To_Number(v_�����);
    Exception
      When Others Then
        n_�����id := 0;
    End;
  
    If n_�����id = 0 Then
      Begin
        Select ID, Decode(Nvl(�Ƿ�����, 0), 1, Null, ���� || 'δ����,��������нɷ�!')
        Into n_�����id, v_Err_Msg
        From ҽ�ƿ����
        Where ���� = v_�����;
      Exception
        When Others Then
          v_Err_Msg := '�����:' || v_����� || '������!';
      End;
    
    Else
    
      Begin
        Select ID, Decode(Nvl(�Ƿ�����, 0), 1, Null, ���� || 'δ����,��������нɷ�!')
        Into n_�����id, v_Err_Msg
        From ҽ�ƿ����
        Where ID = n_�����id;
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ľ���֧����Ϣ!';
      End;
    
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
  End If;
  n_���� := 4;
  --1.��ȡ�Һ����� 

  Select Max(�շѵ�) Into v_No From ���˹Һż�¼ Where NO = v_�Һŵ�;

  If v_No Is Not Null Then
    Select Count(*) Into n_Count From ������ü�¼ Where NO = v_No And ��¼���� = 1;
    If n_Count <> 0 Then
      n_���� := 1;
    End If;
  End If;
  If n_���� = 4 Then
    v_No := v_�Һŵ�;
  End If;

  n_Count := 0;
  For c_�Һ� In (Select a.Id, v_No As NO, n_���� As ��¼����, a.ִ�в���id, c.���� As ִ�в���,
                      To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, To_Char(a.ԤԼʱ��, 'yyyy-mm-dd hh24:mi:ss') As ԤԼʱ��,
                      a.����ʱ��, To_Char(a.����ʱ��, 'yyyy-mm-dd HH24:mi:ss') As ����ʱ��, a.�ű�, a.����, b.���, a.��¼״̬,
                      Decode(Nvl(a.ִ��״̬, 0), 0, '�ȴ�����', 1, '��ɾ���', 2, '���ھ���', -1, 'ȡ������') As ִ��״̬,
                      Decode(Nvl(b.����id, 0), 0, 0, 1) As ֧����־, Decode(Nvl(a.��¼����, 0), 2, 1, 0) As �ɿʽ, b.����id As ����id
               From ���˹Һż�¼ A,
                    (Select Max(Decode(��¼״̬, 0, 0, 2, 0, Nvl(����id, 0))) As ����id, Sum(ʵ�ս��) As ���
                      From ������ü�¼ B
                      Where ��¼���� = n_���� And NO = v_No) B, ���ű� C
               Where a.No = v_�Һŵ� And a.ִ�в���id = c.Id) Loop
  
    If Nvl(c_�Һ�.��¼״̬, 0) <> 1 Then
      v_Err_Msg := '���ݺ�:' || v_�Һŵ� || '�Ѿ����˺�!';
      Raise Err_Item;
    End If;
  
    Select Max(�ŶӺ���), Max(��������)
    Into v_�ŶӺ���, v_��������
    From �ŶӽкŶ���
    Where ҵ��id = c_�Һ�.Id And Nvl(ҵ������, 0) = 0;
  
    If v_�ŶӺ��� Is Not Null Then
      --ҵ��id_In ,ҵ������_In �ŶӺ���_In Number := Null 
      n_Temp := Zl_Getsequencebeforperons(c_�Һ�.Id, 0, v_�ŶӺ���, v_��������);
      v_���� := v_���� || '<DL><XH>' || v_�ŶӺ��� || '</XH><QMRS>' || n_Temp || '</QMRS></DL>';
    End If;
    n_Temp := 0;
    If Nvl(n_�����id, 0) <> 0 Then
      Begin
        Select 1
        Into n_Temp
        From ����Ԥ����¼
        Where ����id = c_�Һ�.����id And ��¼���� = 4 And ��¼״̬ In (1, 3) And �����id = n_�����id And Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    End If;
  
    v_Temp := '<GHDH>' || v_�Һŵ� || '</GHDH>';
    v_Temp := v_Temp || '<DJH>' || c_�Һ�.No || '</DJH>';
    v_Temp := v_Temp || '<YYSJ>' || c_�Һ�.ԤԼʱ�� || '</YYSJ>';
    v_Temp := v_Temp || '<JZSJ>' || c_�Һ�.����ʱ�� || '</JZSJ>';
    v_Temp := v_Temp || '<KDSJ>' || c_�Һ�.�Ǽ�ʱ�� || '</KDSJ>';
    v_Temp := v_Temp || '<JKFS>' || c_�Һ�.�ɿʽ || '</JKFS>';
    v_Temp := v_Temp || '<JE>' || c_�Һ�.��� || '</JE>';
    v_Temp := v_Temp || '<DJLX>' || n_���� || '</DJLX>';
    v_Temp := v_Temp || '<ZFZT>' || c_�Һ�.֧����־ || '</ZFZT>';
    v_Temp := v_Temp || '<SFJSK>' || n_Temp || '</SFJSK>';
    If v_���� Is Not Null Then
      v_Temp := v_Temp || v_����;
    End If;
    v_Temp := '<GH>' || v_Temp || '</GH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    n_Count := n_Count + 1;
  End Loop;

  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := 'δ�ҵ�ָ���ĹҺŵ���:' || v_�Һŵ� || '!';
    Raise Err_Item;
  End If;

  --2.�齨ҽ��������������� 
  n_��ҽ��id := 0;

  For c_ҽ�� In (With ҽ������ As
                  (Select ҽ��id, ���ͺ�, ��¼����, NO, Max(Nvl(ִ��״̬, 0)) As ִ��״̬
                  From (Select b.ҽ��id, b.���ͺ�, b.��¼����, b.No, Nvl(b.ִ��״̬, 0) As ִ��״̬
                         From ����ҽ����¼ A, ����ҽ������ B
                         Where a.�Һŵ� = v_�Һŵ� And a.Id = b.ҽ��id
                         Union All
                         Select b.ҽ��id, b.���ͺ�, b.��¼����, b.No, Nvl(c.ִ��״̬, 0) As ִ��״̬
                         From ����ҽ����¼ A, ����ҽ������ C, ����ҽ������ B
                         Where a.�Һŵ� = v_�Һŵ� And a.Id = b.ҽ��id And b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ�)
                  Group By ҽ��id, ���ͺ�, ��¼����, NO)
                 
                 Select Nvl(a.���id, a.Id) As ��id, Decode(a.���id, Null, 0, 1) As ��ҽ��, a.Id, a.���id, e.��ҩ����,
                        Max(Decode(a.�������, 'E', Decode(q.��������, '2', '����', '4', '����', '6', '����', m.����), m.����)) As ҽ������,
                        a.ִ�п���id, d.���� As ִ�п���, Decode(a.���id, Null, a.ҽ������, Null) As ��ҽ������,
                        Max(Decode(a.�������, '5', 1, '6', 1, '7', 1, 0) * Decode(Nvl(e.ִ��״̬, 0), 1, 1, 3, 1, 0)) As ��ҩ״̬,
                        Decode(a.���id, Null, Null, q.����) As ��ϸҽ������, s.���, (e.���� * e.����) As ����, e.���㵥λ As ��λ,
                        Decode(Nvl(b.ִ��״̬, 0), 0, 'δִ��', 1, '��ȫִ��', 2, '�ܾ�ִ��', 3, '����ִ��', '����ִ��') As ִ��״̬,
                        Max(Decode(p.���ʱ��, Null, Decode(C1.���ʱ��, Null, 0, 1), 1)) As �Ƿ��ѳ�����, c.����id, e.No, e.��¼���� As ��������,
                        Max(e.��׼����) As ��׼����, Sum(e.Ӧ�ս��) As Ӧ�ս��, Sum(e.ʵ�ս��) As ʵ�ս��,
                        To_Char(e.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Decode(Nvl(e.��¼״̬, 0), 0, 0, 3, 2, 1) As ֧��״̬,
                        a.����id
                 
                 From ����ҽ����¼ A, ҽ������ B, ����ҽ������ C, ���Ӳ�����¼ C1, ���ű� D, ������ü�¼ E, ������Ŀ��� M, ������ĿĿ¼ Q, �շ���ĿĿ¼ S, ����걾��¼ P
                 Where a.Id = b.ҽ��id And a.ִ�п���id = d.Id And c.����id = C1.Id(+) And a.Id = c.ҽ��id(+) And a.Id = p.ҽ��id(+) And
                       b.ҽ��id = e.ҽ����� And e.�շ�ϸĿid = s.Id And b.No = e.No And b.��¼���� = e.��¼���� And e.��¼״̬ <> 2 And
                       a.�Һŵ� = v_�Һŵ� And a.������� = m.���� And a.������Ŀid = q.Id And a.ҽ��״̬ In (3, 8)
                 Group By a.Id, a.Ӥ��, a.���, a.���id, e.��ҩ����, a.�������, a.ִ�п���id, d.����, a.ҽ������, q.����, s.���, e.���� * e.����,
                          e.���㵥λ, Decode(Nvl(b.ִ��״̬, 0), 0, 'δִ��', 1, '��ȫִ��', 2, '�ܾ�ִ��', 3, '����ִ��', '����ִ��'), C1.���ʱ��,
                          Decode(c.����id, Null, 0, 1), c.����id, e.No, e.��¼����, e.�Ǽ�ʱ��, Decode(Nvl(e.��¼״̬, 0), 0, 0, 3, 2, 1),
                          p.���ʱ��, a.����id
                 Order By ��id, ��ҽ��, Nvl(a.Ӥ��, 0), a.���) Loop
    If Nvl(n_Add_Djlist, 0) = 0 Then
      --����DJList�ڵ� 
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<YZLIST></YZLIST>')) Into x_Templet From Dual;
      n_Add_Djlist := 1;
    End If;
  
    If n_��ҽ��id <> Nvl(c_ҽ��.��id, 0) Then
      n_��ҽ��id := Nvl(c_ҽ��.��id, 0);
    
      Zl_Third_Custom_Getdeptinfo(n_��ҽ��id, n_ִ�п���id, v_ִ�п���);
    
      If Nvl(n_ִ�п���id, 0) = 0 Then
        If c_ҽ��.ҽ������ = '����' Then
          --����ҽ������ʾ�ɼ����� 
          n_ִ�п���id := c_ҽ��.ִ�п���id;
          v_ִ�п���   := c_ҽ��.ִ�п���;
        Else
          Begin
            Select b.Id, b.����, c.��ҩ����
            Into n_ִ�п���id, v_ִ�п���, v_��ҩ����
            From ����ҽ����¼ A, ���ű� B, ������ü�¼ C
            Where a.Id = c.ҽ����� And a.���id = n_��ҽ��id And a.ִ�п���id = b.Id And Rownum <= 1;
          Exception
            When Others Then
              n_ִ�п���id := c_ҽ��.ִ�п���id;
              v_ִ�п���   := c_ҽ��.ִ�п���;
              v_��ҩ����   := c_ҽ��.��ҩ����;
          End;
        End If;
      End If;
    
      v_Temp := '<YZID>' || n_��ҽ��id || '</YZID>';
      v_Temp := v_Temp || '<YZLX>' || c_ҽ��.ҽ������ || '</YZLX>';
      v_Temp := v_Temp || '<YZMC>' || c_ҽ��.��ҽ������ || '</YZMC>';
      v_Temp := v_Temp || '<ZXKS>' || v_ִ�п��� || '</ZXKS>';
      v_Temp := v_Temp || '<ZXKSID>' || n_ִ�п���id || '</ZXKSID>';
      v_Temp := v_Temp || '<FYCK>' || v_��ҩ���� || '</FYCK>';
      v_Temp := v_Temp || '<BG>' || c_ҽ��.�Ƿ��ѳ����� || '</BG>';
      v_Temp := v_Temp || Zl_Third_Custom_Getrptfrom(n_��ҽ��id);
      v_Temp := v_Temp || Zl_Third_Custom_Rptlimit(c_ҽ��.����id, n_��ҽ��id);
      If Nvl(c_ҽ��.�Ƿ��ѳ�����, 0) = 1 And c_ҽ��.����id Is Not Null Then
        v_Temp := v_Temp || '<BLID>' || c_ҽ��.����id || '</BLID>';
      End If;
      v_Temp := '<YZ ҽ��ID="' || n_��ҽ��id || '">' || v_Temp || '<YZMX></YZMX><DJLIST></DJLIST></YZ>';
      Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      For v_���� In (
                   
                   Select a.No, Mod(a.��¼����, 10) As ��������, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��,
                           Max(Decode(Nvl(a.��¼״̬, 0), 0, 0, 3, 2, 1)) As ֧��״̬, Sum(a.ʵ�ս��) As ���ݽ��, Max(a.����id) As ���㿨֧��
                   From ������ü�¼ A
                   Where (a.No, a.��¼����) In
                         (Select Distinct q.No, q.��¼����
                          From ����ҽ����¼ M, ����ҽ������ Q
                          Where m.Id = q.ҽ��id And (m.Id = n_��ҽ��id Or m.���id = n_��ҽ��id)
                          Union All
                          Select Distinct q.No, q.��¼����
                          From ����ҽ����¼ M, ����ҽ������ Q
                          Where m.Id = q.ҽ��id And (m.Id = n_��ҽ��id Or m.���id = n_��ҽ��id)) And Nvl(a.��¼״̬, 0) In (0, 1, 3)
                   Group By a.No, Mod(a.��¼����, 10), To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss')) Loop
        Begin
          Select 1
          Into n_Temp
          From ����Ԥ����¼ A, ������ü�¼ B
          Where a.����id = b.����id And b.No = v_����.No And Mod(b.��¼����, 10) = 1 And b.��¼״̬ In (1, 3) And a.�����id = n_�����id And
                Rownum < 2;
        Exception
          When Others Then
            n_Temp := 0;
        End;
        Begin
          Select -1 * Sum(���ʽ��)
          Into n_�˿���
          From ������ü�¼ B
          Where b.No = v_����.No And Mod(b.��¼����, 10) = 1 And b.��¼״̬ = 2;
        Exception
          When Others Then
            n_�˿��� := 0;
        End;
        Begin
          Select ״̬, ����ԭ��, ���ԭ��
          Into n_�˷�״̬, v_����ԭ��, v_���ԭ��
          From �����˷�����
          Where NO = v_����.No And Mod(��¼����, 10) = Mod(v_����.��������, 10);
        Exception
          When Others Then
            n_�˷�״̬ := -1;
            v_����ԭ�� := '';
            v_���ԭ�� := '';
        End;
      
        v_Temp := '<DJH>' || v_����.No || '</DJH>';
        v_Temp := v_Temp || '<DJLX>' || v_����.�������� || '</DJLX>';
        v_Temp := v_Temp || '<JE>' || v_����.���ݽ�� || '</JE>';
        v_Temp := v_Temp || '<KDSJ>' || v_����.����ʱ�� || '</KDSJ>';
        If n_�˷�״̬ = -1 Then
          v_Temp := v_Temp || '<ZFZT>' || v_����.֧��״̬ || '</ZFZT>';
        Else
          If n_�˷�״̬ = 0 Then
            v_Temp := v_Temp || '<ZFZT>3</ZFZT>';
          End If;
          If n_�˷�״̬ = 1 Then
            If v_����.֧��״̬ = 2 Then
              v_Temp := v_Temp || '<ZFZT>2</ZFZT>';
            Else
              v_Temp := v_Temp || '<ZFZT>4</ZFZT>';
            End If;
          End If;
          If n_�˷�״̬ = 2 Then
            v_Temp := v_Temp || '<ZFZT>5</ZFZT>';
          End If;
        End If;
      
        If n_�˷�״̬ = -1 Then
          v_Temp := v_Temp || '<SHSM>' || '' || '</SHSM>';
        Else
          If n_�˷�״̬ = 0 Then
            v_Temp := v_Temp || '<SHSM>' || v_����ԭ�� || '</SHSM>';
          End If;
          If n_�˷�״̬ = 1 Then
            v_Temp := v_Temp || '<SHSM>' || v_���ԭ�� || '</SHSM>';
          End If;
          If n_�˷�״̬ = 2 Then
            v_Temp := v_Temp || '<SHSM>' || v_���ԭ�� || '</SHSM>';
          End If;
        End If;
      
        v_Temp := v_Temp || '<YTJE>' || Nvl(n_�˿���, 0) || '</YTJE>';
        v_Temp := v_Temp || '<SFJSK>' || n_Temp || '</SFJSK>';
        v_Temp := '<DJ>' || v_Temp || '</DJ>';
        Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST/YZ[@ҽ��ID="' || n_��ҽ��id || '"]/DJLIST', Xmltype(v_Temp))
        Into x_Templet
        From Dual;
      End Loop;
    End If;
  
    --ֻ��һ����¼��ҽ��������ϸ�����Ӹ���ҽ�����Ի�ȡִ��״̬ 
    Select Decode(Count(*), 0, 1, 0) Into n_����ҽ�� From ����ҽ����¼ Where ���id = n_��ҽ��id;
    If n_����ҽ�� = 1 Then
      v_Temp := '<YZNR>' || c_ҽ��.��ҽ������ || '</YZNR>';
      v_Temp := v_Temp || '<GG>' || c_ҽ��.��� || '</GG>';
      v_Temp := v_Temp || '<SFFY>' || c_ҽ��.��ҩ״̬ || '</SFFY>';
      v_Temp := v_Temp || '<SL>' || c_ҽ��.���� || '</SL>';
      v_Temp := v_Temp || '<DW>' || c_ҽ��.��λ || '</DW>';
      v_Temp := v_Temp || '<BZDJ>' || Nvl(c_ҽ��.��׼����, 0) || '</BZDJ>';
      v_Temp := v_Temp || '<YSJE>' || Nvl(c_ҽ��.Ӧ�ս��, 0) || '</YSJE>';
      v_Temp := v_Temp || '<SSJE>' || Nvl(c_ҽ��.ʵ�ս��, 0) || '</SSJE>';
      v_Temp := v_Temp || '<ZXZT>' || c_ҽ��.ִ��״̬ || '</ZXZT>';
      v_Temp := '<MX>' || v_Temp || '</MX>';
      Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST/YZ[@ҽ��ID="' || n_��ҽ��id || '"]/YZMX', Xmltype(v_Temp))
      Into x_Templet
      From Dual;
    End If;
  
    If Nvl(c_ҽ��.��ҽ��, 0) = 1 Then
      If n_��ϸ���� = 0 Or (n_��ϸ���� = 1 And c_ҽ��.ҽ������ <> '����') Then
        v_Temp := '<YZNR>' || c_ҽ��.��ϸҽ������ || '</YZNR>';
        v_Temp := v_Temp || '<GG>' || c_ҽ��.��� || '</GG>';
        v_Temp := v_Temp || '<SL>' || c_ҽ��.���� || '</SL>';
        v_Temp := v_Temp || '<DW>' || c_ҽ��.��λ || '</DW>';
        v_Temp := v_Temp || '<SFFY>' || c_ҽ��.��ҩ״̬ || '</SFFY>';
        v_Temp := v_Temp || '<ZXZT>' || c_ҽ��.ִ��״̬ || '</ZXZT>';
        v_Temp := v_Temp || '<BZDJ>' || Nvl(c_ҽ��.��׼����, 0) || '</BZDJ>';
        v_Temp := v_Temp || '<YSJE>' || Nvl(c_ҽ��.Ӧ�ս��, 0) || '</YSJE>';
        v_Temp := v_Temp || '<SSJE>' || Nvl(c_ҽ��.ʵ�ս��, 0) || '</SSJE>';
        v_Temp := '<MX>' || v_Temp || '</MX>';
        Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST/YZ[@ҽ��ID="' || n_��ҽ��id || '"]/YZMX', Xmltype(v_Temp))
        Into x_Templet
        From Dual;
      End If;
    End If;
  
  End Loop;
  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getvisitinfo;
/

--119329:Ƚ����,2018-05-16,�����ӿڻ�ȡ�ɹҺſ��ҹ��̵���
Create Or Replace Procedure Zl_Third_Getdeptlist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:��ȡ�ɹҺſ���
  --���:Xml_In:
  --<IN>
  --  <CXTS>14</CXTS>        //��ѯ����
  --  <HZDW>֧����</HZDW>    //������λ
  --  <ZD></ZD>              //վ��
  --</IN>
  --����:Xml_Out
  --<OUTPUT>
  -- <KSLIST>
  --  <KS>
  --    <ID>����ID</ID>       //����ID
  --    <MC>��������</MC>     //��������
  --    <ZDYYTS>����ԤԼ����</ZDYYTS>     //����ԤԼ����
  --  </KS>
  --  <KS>
  --    ...
  --  </KS>
  -- </KSLIST>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  x_Templet Xmltype; --ģ��XML
  v_Temp    Varchar2(5000); --��ʱXML

  n_��ѯ���� Number(5);
  v_������λ ������λ���ſ���.������λ%Type;
  v_վ��     ���ű�.վ��%Type;

  v_Para     Varchar2(4000);
  n_ԤԼ���� Number(5);
  n_�������� Number(5);

  n_�Һ�ģʽ Number(3);
  d_����ʱ�� Date;
Begin
  x_Templet := Xmltype('<OUTPUT><KSLIST></KSLIST></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/CXTS'), Extractvalue(Value(A), 'IN/HZDW'), Nvl(Extractvalue(Value(A), 'IN/ZD'), '-')
  Into n_��ѯ����, v_������λ, v_վ��
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  v_Para     := zl_GetSysParameter(256);
  n_�Һ�ģʽ := To_Number(Substr(v_Para, 1, 1));
  If n_�Һ�ģʽ = 1 Then
    Begin
      d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
  End If;
  n_ԤԼ���� := zl_GetSysParameter(66);

  If n_�Һ�ģʽ = 0 Then
    If v_������λ Is Null Then
      For r_Dept In (Select a.����id, d.����, Max(Nvl(a.ԤԼ����, n_ԤԼ����)) As ԤԼ����
                     From �ҺŰ��� A, ���ű� D
                     Where a.����id = d.Id And a.ͣ������ Is Null And Nvl(d.վ��, v_վ��) = v_վ�� And
                           (d.����ʱ�� Is Null Or d.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd'))
                     Group By a.����id, d.����) Loop
        v_Temp := '<KS>' || '<ID>' || r_Dept.����id || '</ID>' || '<MC>' || r_Dept.���� || '</MC>';
        v_Temp := v_Temp || '<ZDYYTS>' || r_Dept.ԤԼ���� || '</ZDYYTS>' || '</KS>';
        Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
      End Loop;
    Else
      For r_Dept In (Select ����id, ����, Max(ԤԼ����) As ԤԼ����
                     From (
                            --1.�ƻ�
                            Select a.����id, d.����, Nvl(a.ԤԼ����, n_ԤԼ����) As ԤԼ����
                            From �ҺŰ��� A, �ҺŰ��żƻ� C, ���ű� D
                            Where a.Id = c.����id And a.����id = d.Id And a.ͣ������ Is Null And c.���ʱ�� Is Not Null And
                                  Not (c.ʧЧʱ�� < Sysdate Or c.��Чʱ�� > Sysdate + Nvl(n_��ѯ����, Nvl(a.ԤԼ����, n_ԤԼ����)))
                                 
                                  And
                                  (Not Exists (Select 1 From ������λ�ƻ����� Where �ƻ�id = c.Id And ������λ = v_������λ) Or Exists
                                   (Select 1
                                    From ������λ�ƻ�����
                                    Where �ƻ�id = c.Id And ������λ = v_������λ And ���� <> 0))
                                 
                                  And Nvl(d.վ��, v_վ��) = v_վ�� And
                                  (d.����ʱ�� Is Null Or d.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd'))
                            --2.����
                            Union All
                            Select a.����id, d.����, Nvl(a.ԤԼ����, n_ԤԼ����) As ԤԼ����
                            From �ҺŰ��� A, ���ű� D
                            Where a.����id = d.Id And a.ͣ������ Is Null And Not Exists
                             (Select 1 From �ҺŰ��żƻ� Where ����id = a.Id)
                                 
                                  And
                                  (Not Exists (Select 1 From ������λ���ſ��� Where ����id = a.Id And ������λ = v_������λ) Or Exists
                                   (Select 1
                                    From ������λ���ſ���
                                    Where ����id = a.Id And ������λ = v_������λ And ���� <> 0))
                                 
                                  And Nvl(d.վ��, v_վ��) = v_վ�� And
                                  (d.����ʱ�� Is Null Or d.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')))
                     Group By ����id, ����) Loop
        v_Temp := '<KS>' || '<ID>' || r_Dept.����id || '</ID>' || '<MC>' || r_Dept.���� || '</MC>';
        v_Temp := v_Temp || '<ZDYYTS>' || r_Dept.ԤԼ���� || '</ZDYYTS>' || '</KS>';
        Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
      End Loop;
    End If;
  Else
    --������Ű�ģʽ
    n_�������� := Zl_Fun_Getappointmentdays;
    If v_������λ Is Null Then
      For r_Dept In (Select ����id, ����, Max(ԤԼ����) As ԤԼ����
                     From (
                            --����ǰ
                            Select a.����id, d.����, Nvl(a.ԤԼ����, n_ԤԼ����) As ԤԼ����
                            From �ҺŰ��� A, ���ű� D
                            Where a.����id = d.Id And a.ͣ������ Is Null And Nvl(d.վ��, v_վ��) = v_վ�� And
                                  (d.����ʱ�� Is Null Or d.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')) And Sysdate < d_����ʱ��
                            --���ú�
                            Union All
                            Select a.����id, d.����, Nvl(c.ԤԼ����, n_ԤԼ����) + n_�������� As ԤԼ����
                            From �ٴ������¼ A, �ٴ������Դ C, ���ű� D
                            Where a.��Դid = c.Id And a.����id = d.Id And a.�������� Between Trunc(Sysdate) And
                                  Trunc(Sysdate) + Nvl(n_��ѯ����, Nvl(c.ԤԼ����, n_ԤԼ����) + n_��������) And a.��ʼʱ�� > d_����ʱ�� And
                                  Nvl(a.�Ƿ񷢲�, 0) = 1
                                 --�ų�ȫʱ��ͣ���˵�
                                  And (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) And Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) > Sysdate Or
                                  a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��) And a.��ֹʱ�� > Sysdate Or Exists
                                   (Select 1
                                        From �ٴ�������ſ���
                                        Where ��¼id = a.Id And Nvl(�Ƿ�ͣ��, 0) = 0 And Nvl(a.�Ƿ���ſ���, 0) = 1 And
                                              Nvl(a.�Ƿ��ʱ��, 0) = 1 And ��ʼʱ�� <> ��ֹʱ�� And ��ʼʱ�� >= Sysdate))
                                 --
                                  And (c.����ʱ�� Is Null Or c.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')) And
                                  Nvl(d.վ��, v_վ��) = v_վ�� And
                                  (d.����ʱ�� Is Null Or d.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')))
                     Group By ����id, ����) Loop
        v_Temp := '<KS>' || '<ID>' || r_Dept.����id || '</ID>' || '<MC>' || r_Dept.���� || '</MC>';
        v_Temp := v_Temp || '<ZDYYTS>' || r_Dept.ԤԼ���� || '</ZDYYTS>' || '</KS>';
        Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
      End Loop;
    Else
      For r_Dept In (Select ����id, ����, Max(ԤԼ����) As ԤԼ����
                     From (
                            --1.����ǰ
                            --1.1.�ƻ�
                            Select a.����id, d.����, Nvl(a.ԤԼ����, n_ԤԼ����) As ԤԼ����
                            From �ҺŰ��� A, �ҺŰ��żƻ� C, ���ű� D
                            Where a.Id = c.����id And a.����id = d.Id And a.ͣ������ Is Null And c.���ʱ�� Is Not Null And
                                  Not (c.ʧЧʱ�� < Sysdate Or c.��Чʱ�� > Sysdate + Nvl(n_��ѯ����, Nvl(a.ԤԼ����, n_ԤԼ����)))
                                 
                                  And
                                  (Not Exists (Select 1 From ������λ�ƻ����� Where �ƻ�id = c.Id And ������λ = v_������λ) Or Exists
                                   (Select 1
                                    From ������λ�ƻ�����
                                    Where �ƻ�id = c.Id And ������λ = v_������λ And ���� <> 0))
                                 
                                  And Nvl(d.վ��, v_վ��) = v_վ�� And
                                  (d.����ʱ�� Is Null Or d.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')) And Sysdate < d_����ʱ��
                            --1.2.����
                            Union All
                            Select a.����id, d.����, Nvl(a.ԤԼ����, n_ԤԼ����) As ԤԼ����
                            From �ҺŰ��� A, ���ű� D
                            Where a.����id = d.Id And a.ͣ������ Is Null And Not Exists
                             (Select 1 From �ҺŰ��żƻ� Where ����id = a.Id)
                                 
                                  And
                                  (Not Exists (Select 1 From ������λ���ſ��� Where ����id = a.Id And ������λ = v_������λ) Or Exists
                                   (Select 1
                                    From ������λ���ſ���
                                    Where ����id = a.Id And ������λ = v_������λ And ���� <> 0))
                                 
                                  And Nvl(d.վ��, v_վ��) = v_վ�� And
                                  (d.����ʱ�� Is Null Or d.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')) And Sysdate < d_����ʱ��
                            --2.���ú�
                            Union All
                            Select a.����id, d.����, Nvl(c.ԤԼ����, n_ԤԼ����) + n_�������� As ԤԼ����
                            From �ٴ������¼ A, �ٴ������Դ C, ���ű� D
                            Where a.��Դid = c.Id And a.����id = d.Id And a.�������� Between Trunc(Sysdate) And
                                  Trunc(Sysdate) + Nvl(n_��ѯ����, Nvl(c.ԤԼ����, n_ԤԼ����) + n_��������) And a.��ʼʱ�� >= d_����ʱ�� And
                                  Nvl(a.�Ƿ񷢲�, 0) = 1
                                 --�ų�ȫʱ��ͣ���˵�
                                  And (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) And Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) > Sysdate Or
                                  a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��) And a.��ֹʱ�� > Sysdate Or Exists
                                   (Select 1
                                        From �ٴ�������ſ���
                                        Where ��¼id = a.Id And Nvl(�Ƿ�ͣ��, 0) = 0 And Nvl(a.�Ƿ���ſ���, 0) = 1 And
                                              Nvl(a.�Ƿ��ʱ��, 0) = 1 And ��ʼʱ�� <> ��ֹʱ�� And ��ʼʱ�� >= Sysdate))
                                 --�ٴ������¼.ԤԼ���ƣ�0-����ԤԼ����;1-�úű��ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
                                  And
                                  (Not Exists (Select 1
                                               From �ٴ�����Һſ��Ƽ�¼
                                               Where ��¼id = a.Id And ���� = 1 And ���� = 1 And ���� = v_������λ) Or Exists
                                   (Select 1
                                    From �ٴ�����Һſ��Ƽ�¼
                                    Where ��¼id = a.Id And ���� = 1 And ���� = 1 And ���� = v_������λ
                                         --�ٴ�����Һſ��Ƽ�¼.���Ʒ�ʽ��0-��ֹԤԼ;1-����������ԤԼ;2-����������ԤԼ;3-����ſ���ԤԼ;4-��������
                                          And (���Ʒ�ʽ In (1, 2, 3) And ���� <> 0 Or ���Ʒ�ʽ = 4)))
                                 
                                  And (c.����ʱ�� Is Null Or c.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')) And
                                  Nvl(d.վ��, v_վ��) = v_վ�� And
                                  (d.����ʱ�� Is Null Or d.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')))
                     Group By ����id, ����) Loop
        v_Temp := '<KS>' || '<ID>' || r_Dept.����id || '</ID>' || '<MC>' || r_Dept.���� || '</MC>';
        v_Temp := v_Temp || '<ZDYYTS>' || r_Dept.ԤԼ���� || '</ZDYYTS>' || '</KS>';
        Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
      End Loop;
    End If;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getdeptlist;
/

--125688:����,2018-05-16,ҩƷ������ⷿ���δ���
Create Or Replace Procedure Zl_ҩƷ����_Insert
(
  No_In           In ҩƷ�շ���¼.No%Type,
  ���_In         In ҩƷ�շ���¼.���%Type,
  �ⷿid_In       In ҩƷ�շ���¼.�ⷿid%Type,
  �Է�����id_In   In ҩƷ�շ���¼.�Է�����id%Type,
  ҩƷid_In       In ҩƷ�շ���¼.ҩƷid%Type,
  ����_In         In ҩƷ�շ���¼.����%Type,
  ��д����_In     In ҩƷ�շ���¼.��д����%Type,
  ʵ������_In     In ҩƷ�շ���¼.ʵ������%Type,
  �ɱ���_In       In ҩƷ�շ���¼.�ɱ���%Type,
  �ɱ����_In     In ҩƷ�շ���¼.�ɱ����%Type,
  ���ۼ�_In       In ҩƷ�շ���¼.���ۼ�%Type,
  ���۽��_In     In ҩƷ�շ���¼.���۽��%Type,
  ���_In         In ҩƷ�շ���¼.���%Type,
  ������_In       In ҩƷ�շ���¼.������%Type,
  ����_In         In ҩƷ�շ���¼.����%Type := Null,
  ����_In         In ҩƷ�շ���¼.����%Type := Null,
  Ч��_In         In ҩƷ�շ���¼.Ч��%Type := Null,
  ժҪ_In         In ҩƷ�շ���¼.ժҪ%Type := Null,
  ��������_In     In ҩƷ�շ���¼.��������%Type := Null,
  �ϴι�Ӧ��id_In In ҩƷ�շ���¼.��ҩ��λid%Type := Null,
  ��׼�ĺ�_In     In ҩƷ�շ���¼.��׼�ĺ�%Type := Null,
  ���췽ʽ_In     In ҩƷ�շ���¼.����%Type := 0,
  ����ʱ��_In     In ҩƷ�շ���¼.Ƶ��%Type := Null,
  ԭ����_In       In ҩƷ�շ���¼.ԭ����%Type := Null,
  �޸���_In       In ҩƷ�շ���¼.�޸���%Type,
  �޸�����_In     In ҩƷ�շ���¼.�޸�����%Type := Null
) Is
  v_Lngid        ҩƷ�շ���¼.Id%Type; --�շ�ID 
  n_�����շ�id   ҩƷ�շ���¼.Id%Type; --����ⷿ�շ�id 
  v_������id   ҩƷ�շ���¼.������id%Type; --������ID 
  v_�������id   ҩƷ�շ���¼.������id%Type; --������ID 
  d_�ϴ��������� ҩƷ���.�ϴ���������%Type;
  v_����         ҩƷ�շ���¼.����%Type := Null; --��Ҫ��������ʵ��ҩ�������ҩƷ
  v_�Ƿ����     Integer; --�ж�����Ƿ�ҩ�����   1:������0��������
  v_ҩ�����     Integer; --�ж�����Ƿ�ҩ�����   1:������0��������
  v_ҩ������     Integer; --�ж�����Ƿ�ҩ�����   1:������0��������
Begin
  --�����ҳ���ͳ������ID 
  Select b.Id 
  Into v_������id 
  From ҩƷ�������� A, ҩƷ������ B 
  Where a.���id = b.Id And a.���� = 6 And b.ϵ�� = 1 And Rownum < 2; 
  
  Select b.Id 
  Into v_�������id 
  From ҩƷ�������� A, ҩƷ������ B 
  Where a.���id = b.Id And a.���� = 6 And b.ϵ�� = -1 And Rownum < 2; 

  Select ҩƷ�շ���¼_Id.Nextval Into v_Lngid From Dual;

  Begin
    Select �ϴ���������
    Into d_�ϴ���������
    From ҩƷ���
    Where ���� = 1 And �ⷿid = �ⷿid_In And ҩƷid = ҩƷid_In And Nvl(����, 0) = Nvl(����_In, 0);
  Exception
    When Others Then
      d_�ϴ��������� := Null;
  End;

  Select ҩƷ�շ���¼_Id.Nextval Into n_�����շ�id From Dual;
  --�������Ϊ������һ�� 
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ԭ����,����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ,
     ������, ��������, ��ҩ��ʽ, ��ҩ��λid, ��׼�ĺ�, ��������, ����, Ƶ��, �޸���, �޸�����)
  Values
    (n_�����շ�id, 1, 6, No_In, ���_In, �ⷿid_In, �Է�����id_In, v_�������id, -1, ҩƷid_In, ����_In, ����_In, ԭ����_In,����_In, Ч��_In, ��д����_In,
     ʵ������_In, �ɱ���_In, �ɱ����_In, ���ۼ�_In, ���۽��_In, ���_In, ժҪ_In, ������_In, ��������_In, 1, �ϴι�Ӧ��id_In, ��׼�ĺ�_In, d_�ϴ���������, ���췽ʽ_In,
     ����ʱ��_In, �޸���_In, �޸�����_In);
  
  Zl_δ��ҩƷ��¼_Insert(n_�����շ�id);

  --������
  Zl_ҩƷ���_Update(n_�����շ�id, 0);

  Select Nvl(ҩ�����, 0), Nvl(ҩ������, 0) Into v_ҩ�����, v_ҩ������ From ҩƷ��� Where ҩƷid = ҩƷid_In;

  v_�Ƿ���� := 0;
  If v_ҩ������ = 0 Then
    If v_ҩ����� = 1 Then
      Begin
        Select Distinct 0
        Into v_�Ƿ����
        From ��������˵��
        Where ((�������� Like '%ҩ��') Or (�������� Like '�Ƽ���')) And ����id = �Է�����id_In;
      Exception
        When Others Then
          v_�Ƿ���� := 1;
      End;
    End If;
  Else
    v_�Ƿ���� := 1;
  End If;

  If v_�Ƿ���� = 1 And Nvl(����_In, 0) = 0 Then
    --�������ҳ��ⲻ����
    v_���� := Zl_Fun_Getbatchnum(ҩƷid_In, ����_In, ����_In, �ɱ���_In, ���ۼ�_In, v_Lngid, �ϴι�Ӧ��id_In);
  Elsif v_�Ƿ���� = 0 Then
    --��ⲻ����
    v_���� := 0;
  Elsif Nvl(����_In, 0) <> 0 Then
    --�������ҳ���Ҳ����
    v_���� := ����_In;
  End If;

  --�������Ϊ�����һ�� 
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����,ԭ����, ����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ,
     ������, ��������, ��ҩ��ʽ, ��ҩ��λid, ��׼�ĺ�, ��������, ����, Ƶ��, �޸���, �޸�����)
  Values
    (v_Lngid, 1, 6, No_In, ���_In + 1, �Է�����id_In, �ⷿid_In, v_������id, 1, ҩƷid_In, v_����, ����_In,ԭ����_In, ����_In, Ч��_In, ��д����_In,
     ʵ������_In, �ɱ���_In, �ɱ����_In, ���ۼ�_In, ���۽��_In, ���_In, ժҪ_In, ������_In, ��������_In, 1, �ϴι�Ӧ��id_In, ��׼�ĺ�_In, d_�ϴ���������, ���췽ʽ_In,
     ����ʱ��_In, �޸���_In, �޸�����_In);
  
  Zl_δ��ҩƷ��¼_Insert(v_Lngid);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ����_Insert;
/

--113688:��ΰ��,2018-05-14,���ȡ������Ժ���������۲��˵ĵǼǺ�����Ϣ�е����Ժ��Ϣû�и��µ�����
--97423:��ΰ��,2018-05-14,��������Ժ����ȡ���Ǽ�ʱ�Ҳ������ݵ�����
Create Or Replace Procedure Zl_��Ժ������ҳ_Delete
(
  ����id_In     ������ҳ.����id%Type,
  ��ҳid_In     ������ҳ.��ҳid%Type,
  ת����_In     Number := 0,
  ���סԺ��_In Number := 0
  --���ܣ�ȡ��������Ժ/ԤԼ�Ǽ�
  --     ��ҳID_IN:Ϊ0ʱ��ʾȡ��ԤԼ�Ǽ�
  --     ת����_IN:��������Ժ�Ǽǲ���תΪסԺ���۲���
  --     ���סԺ��_In:��һ��סԺ�Ĳ���ת����ʱ�Ƿ����סԺ��
) As
  v_��Ժʱ��   ������ҳ.��Ժ����%Type;
  v_��Ժ����   ������ҳ.��Ժ����id%Type;
  v_��Ժʱ��   ������ҳ.��Ժ����%Type;
  v_סԺ��     ������ҳ.סԺ��%Type;
  v_����Ժ     ������ҳ.����Ժ%Type;
  v_��Ժ����id ������ҳ.��Ժ����id%Type;
  n_��������   ������ҳ.��������%Type;
  n_��ҳid     ������ҳ.��ҳid%Type;

  v_Count Number;
  v_Error Varchar2(255);
  Err_Custom Exception;

  Function Checkpatiadvice
  (
    ����id_In ������ҳ.����id%Type,
    ��ҳid_In ������ҳ.��ҳid%Type
  ) Return Varchar2 Is
    --����סԺ����ҽ����¼��������
    v_Err Varchar2(255);
  Begin
    v_Err := Null;
  
    For r_Row In (Select ����ҽ��, Decode(ҽ��״̬, -1, '�ݴ�', 1, '�¿�', 2, 'У������', 'δ����') As ״̬, ҽ������
                  From ����ҽ����¼
                  Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(ҽ��״̬, 0) <> 4 And Rownum < 2) Loop
      v_Err := '��' || r_Row.����ҽ�� || '��ҽ����' || r_Row.״̬ || '��ҽ��û�д���,������ȡ���Ǽǣ�';
    End Loop;
    Return v_Err;
  End Checkpatiadvice;
Begin
  Select Nvl(״̬, 0), Nvl(��������, 0)
  Into v_Count, n_��������
  From ������ҳ
  Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  If v_Count <> 1 Then
    v_Error := '�ò����Ѿ����,���Ƚ����˳�������Ժ״̬��';
    Raise Err_Custom;
  End If;

  --ɾ�����Ӳ���ʱ��
  Select ��Ժ����id, ����Ժ Into v_��Ժ����id, v_����Ժ From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  If v_����Ժ = 0 Then
    Zl_���Ӳ���ʱ��_Delete(����id_In, ��ҳid_In, '��Ժ', v_��Ժ����id);
  Else
    Zl_���Ӳ���ʱ��_Delete(����id_In, ��ҳid_In, '�ٴ���Ժ', v_��Ժ����id);
  End If;

  --��ȡ���һ�β�Ϊ�յ�סԺ��
  Begin
    If ��ҳid_In = 0 Then
      Select סԺ��
      Into v_סԺ��
      From ������ҳ
      Where ����id = ����id_In And
            ��ҳid =
            (Select Max(��ҳid) From ������ҳ Where ����id = ����id_In And Nvl(��ҳid, 0) <> 0 And Nvl(סԺ��, 0) <> 0);
    Else
      Select סԺ��
      Into v_סԺ��
      From ������ҳ
      Where ����id = ����id_In And
            ��ҳid =
            (Select Max(��ҳid) From ������ҳ Where ����id = ����id_In And ��ҳid < ��ҳid_In And Nvl(סԺ��, 0) <> 0);
    End If;
  Exception
    When Others Then
      Null;
  End;

  b_Message.ZLHIS_PATIENT_006(����id_In,��ҳid_In,'��Ժ�Ǽ�');

  If ת����_In = 1 And Nvl(��ҳid_In, 0) <> 0 Then
    Update ������ҳ
    Set �������� = 2, סԺ�� = Decode(���סԺ��_In, 1, Null, סԺ��)
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(��������, 0) = 0;
  
    --����סԺ����
    Update ������Ϣ Set סԺ���� = Decode(Sign(סԺ���� - 1), 1, סԺ���� - 1, Null) Where ����id = ����id_In;
    If ���סԺ��_In = 1 Then
      Update ������Ϣ Set סԺ�� = v_סԺ�� Where ����id = ����id_In;
    End If;
  Else
    Begin
      Select b.��Ժ����, b.��Ժ����, b.��Ժ����id
      Into v_��Ժʱ��, v_��Ժʱ��, v_��Ժ����
      From ������Ϣ A, ������ҳ B
      Where a.����id = ����id_In And a.����id = b.����id And a.��ҳid = b.��ҳid And Nvl(b.��ҳid, 0) <> 0;
    Exception
      When Others Then
        Null;
    End;
    --����ԤԼ�Ǽǲ��˲����סԺ�ձ�
    If Nvl(��ҳid_In, 0) <> 0 Then
      Select Zl_סԺ�ձ�_Count(v_��Ժ����, v_��Ժʱ��) Into v_Count From Dual;
      If v_Count > 0 Then
        v_Error := '�Ѳ���ҵ��ʱ���ڵ�סԺ�ձ�,���ܰ����ҵ��!';
        Raise Err_Custom;
      End If;
    End If;
    --�������۲����´���Ժ֪ͨ�����������Ч�Ĳ�����ҳ��¼��36549��
    Select Count(*) Into v_Count From ������ҳ Where ����id = ����id_In And ��Ժ���� Is Not Null And ��Ժ���� Is Null;
    If Not v_Count > 1 Then
      v_Count := 0;
      If Nvl(��ҳid_In, 0) <> 0 And Nvl(n_��������, 0) = 0 Then
        v_Count := 1;
      End If;
      --����Ժ����,ȡ����Ժ�Ǽ�ʱ,������Ϣ����Ժʱ��ͳ�Ժʱ��Ӧ�û��˵���һ����Ժ���ںͳ�Ժ����
      If v_����Ժ = 1 Then
        Begin
          Select ��Ժ����, ��Ժ����
          Into v_��Ժʱ��, v_��Ժʱ��
          From ������ҳ
          Where ����id = ����id_In And
                ��ҳid = (Select Max(��ҳid)
                        From ������ҳ
                        Where ����id = ����id_In And ��ҳid < ��ҳid_In);
        Exception
          When Others Then
            --�쳣������Ϊ������ȡ�������ݵ��쳣���
            Null;
        End;
      End If;
    
      Update ������Ϣ
      Set סԺ�� = v_סԺ��, סԺ���� = Decode(v_Count, 0, סԺ����, Decode(Sign(סԺ���� - 1), 1, סԺ���� - 1, Null)), ��ǰ����id = Null,
          ��ǰ����id = Null, ��ǰ���� = Null, ��Ժʱ�� = v_��Ժʱ��, ��Ժʱ�� = v_��Ժʱ��, ������ = Null, ������ = Null, �������� = Null, ��Ժ = Null
      Where ����id = ����id_In;
      Delete From ��Ժ���� Where ����id = ����id_In;
    End If;
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    Delete From �����Զ����� Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    Delete From ������ϼ�¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼��Դ = 2;
  
    --����סԺ�������Ԥ����,��Ϊ�������ｻ��
    Update ����Ԥ����¼ Set ��ҳid = Null Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  
    --���η�����,�ı����﷢��
    Update סԺ���ü�¼ Set ��ҳid = Null Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 5;
  
    --����סԺ�����з��ü�¼�޽�������ȫ���������򽫶�Ӧ���ü�¼�е�"��ҳID"�����
    v_Count := 0;
    Select Nvl(Count(*), 0)
    Into v_Count
    From סԺ���ü�¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ���ʷ��� = 1 And ����id Is Not Null;
  
    If v_Count = 0 Then
      Begin
        Select Nvl(Count(*), 0)
        Into v_Count
        From סԺ���ü�¼
        Where ����id = ����id_In And ��ҳid = ��ҳid_In And ���ʷ��� = 1
        Group By NO, ��¼����, ���
        Having Nvl(Sum(ʵ�ս��), 0) <> 0;
      Exception
        When Others Then
          v_Count := 0;
      End;
    
      If v_Count = 0 Then
        Delete ����δ����� Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��� = 0;
        Update סԺ���ü�¼ Set ��ҳid = Null Where ����id = ����id_In And ��ҳid = ��ҳid_In And ���ʷ��� = 1;
      End If;
    End If;
  
    --����סԺ����ҽ����¼��������
    v_Count := 0;
    Select Nvl(Count(*), 0)
    Into v_Count
    From ����ҽ����¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(ҽ��״̬, 0) <> 4;
    If v_Count = 0 Then
      Delete From ����ҽ����¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    Else
      v_Error := Checkpatiadvice(����id_In, ��ҳid_In);
      If v_Error Is Not Null Then
        Raise Err_Custom;
      End If;
    End If;
  
    --���±�,û�н�������ҳ(����ID,��ҳID)�����,��Ϊ����ҳID�����ǹҺ�ID
    Delete From ���˹�����¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ������ϼ�¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ������������¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ���Ӳ�����¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ���Ӳ�����ӡ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    --�����Ժ�����˾��￨,��ɾ����ʧ��(���˷��ü�¼��ҳID�����Լ��)
    Delete From ������ҳ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    --�޸Ĳ�����Ϣ����ҳID��סԺ����
    Select Max(��ҳid) Into n_��ҳid From ������ҳ Where ����id = ����id_In And Nvl(��ҳid, 0) <> 0;
    Update ������Ϣ Set ��ҳid = n_��ҳid Where ����id = ����id_In;
    If n_��ҳid Is Null Then
      Update ������Ϣ Set סԺ���� = Null Where ����id = ����id_In;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Ժ������ҳ_Delete;
/

--125233:����,2018-05-14,��������λ�ҺŻ��ܱ����������ʱû�������ŵ������ʧ��
Create Or Replace Procedure Zl_���������Һ�_Delete
(
  ���ݺ�_In     ������ü�¼.No%Type,
  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
  ����˵��_In   ����Ԥ����¼.����˵��%Type,
  �˺�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
  Ԥ��id_In     ����Ԥ����¼.Id%Type := Null
) As
  v_Error Varchar2(255);
  Err_Custom Exception;

  --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ���
  Cursor c_Registinfo
  (
    v_״̬     ���˹Һż�¼.��¼״̬%Type,
    v_����     ���˹Һż�¼.��¼����%Type,
    v_��Ч���� Number := 0
  ) Is
    Select a.����ʱ��, a.�Ǽ�ʱ��, b.��Ŀid, b.����id, b.ҽ������, b.ҽ��id, b.����
    From ���˹Һż�¼ A, �ҺŰ��� B
    Where a.��¼���� = Decode(v_��Ч����, 0, v_����, a.��¼����) And a.��¼״̬ = v_״̬ And a.No = ���ݺ�_In And a.�ű� = b.���� And Rownum = 1;

  r_Registrow c_Registinfo%RowType;

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
  Cursor c_Opermoney Is
    Select Distinct b.���㷽ʽ, b.��Ԥ��
    From ������ü�¼ A, ����Ԥ����¼ B
    Where a.����id = b.����id And a.No = ���ݺ�_In And a.��¼���� = 4 And a.��¼״̬ = 3 And b.��¼���� = 4 And b.��¼״̬ = 3 And
          Nvl(b.��Ԥ��, 0) <> 0;

  n_ִ��״̬       ���˹Һż�¼.ִ��״̬%Type;
  n_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;
  n_����id         ������ü�¼.����id%Type;
  n_ԭ����id       ����Ԥ����¼.����id%Type;
  n_����id         ������Ϣ.����id%Type;
  n_����ֵ         �������.Ԥ�����%Type;
  n_����̨ǩ���Ŷ� Number;
  n_ԤԼ�Һ�       Number;
  n_��Ч����       Number; --��Ч����û�в������õ���
  n_�Һ����ɶ���   Number;
  n_Count          Number;
  n_��id           ����ɿ����.Id%Type;
  d_�˺�ʱ��       Date;
  v_����Ա���     ��Ա��.���%Type;
  v_����Ա����     ��Ա��.����%Type;
  v_������λ       ������λ�ҺŻ���.������λ%Type;
  n_ԤԼ״̬       ���˹Һż�¼.ԤԼ%Type;
  v_Temp           Varchar2(100);
  d_�Ǽ�ʱ��       ���˹Һż�¼.�Ǽ�ʱ��%Type;
  v_�ű�           ���˹Һż�¼.�ű�%Type;
  n_����           ���˹Һż�¼.����%Type;
  d_ԤԼʱ��       ���˹Һż�¼.ԤԼʱ��%Type;
  n_������λ����   Number(18);
  n_ԤԼ���ɶ���   Number;
  n_��¼����       Number;
  n_״̬           Number;
  n_�˺�����       Number(3);
  n_�Һ��Ű�ģʽ   Number;
  n_����           ������ü�¼.���ʷ���%Type;
  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_�ѽ���         Number;
  n_���ض�         �������.�������%Type;
  n_Ԥ��֧��       Number(3);
  n_����֧��       Number(3);
  n_����id         �ҺŰ���.Id%Type;
  n_�ƻ�id         �ҺŰ��żƻ�.Id%Type;
  v_ʱ���         ʱ���.ʱ���%Type;
  d_��鿪ʼʱ��   Date;
  d_������ʱ��   Date;
  d_����ʱ��       Date;
  n_�����¼id     Number(18);
  Function Zl_����Ա
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-��ȡȱʡ����ID;1-��ȡ����Ա���;2-��ȡ����Ա����
    -- SplitStr:��ʽΪ:����ID,��������;��ԱID,��Ա���,��Ա����(��Zl_Identity��ȡ��)
  Begin
    If Type_In = 0 Then
      --ȱʡ����
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      Return v_Sub;
    End If;
  End;

  Procedure Zl_���������Һ�_����_Delete
  (
    ���ݺ�_In     ������ü�¼.No%Type,
    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    ����˵��_In   ����Ԥ����¼.����˵��%Type,
    �˺�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
    Ԥ��id_In     ����Ԥ����¼.Id%Type := Null
  ) As
    v_Error Varchar2(255);
    Err_Custom Exception;
  
    --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ���
    Cursor c_Registinfo
    (
      v_״̬     ���˹Һż�¼.��¼״̬%Type,
      v_����     ���˹Һż�¼.��¼����%Type,
      v_��Ч���� Number := 0
    ) Is
      Select a.����ʱ��, a.�Ǽ�ʱ��, b.��Ŀid, b.����id, b.ҽ������, b.ҽ��id, b.Id As ��¼id, a.�ű� As ����
      From ���˹Һż�¼ A, �ٴ������¼ B
      Where a.��¼���� = Decode(v_��Ч����, 0, v_����, a.��¼����) And a.��¼״̬ = v_״̬ And a.No = ���ݺ�_In And a.�����¼id = b.Id And
            Rownum < 2;
  
    r_Registrow c_Registinfo%RowType;
  
    --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
    Cursor c_Opermoney Is
      Select Distinct b.���㷽ʽ, b.��Ԥ��
      From ������ü�¼ A, ����Ԥ����¼ B
      Where a.����id = b.����id And a.No = ���ݺ�_In And a.��¼���� = 4 And a.��¼״̬ = 3 And b.��¼���� = 4 And b.��¼״̬ = 3 And
            Nvl(b.��Ԥ��, 0) <> 0;
  
    n_ִ��״̬       ���˹Һż�¼.ִ��״̬%Type;
    n_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;
    n_����id         ������ü�¼.����id%Type;
    n_ԭ����id       ����Ԥ����¼.����id%Type;
    n_����id         ������Ϣ.����id%Type;
    n_����ֵ         �������.Ԥ�����%Type;
    n_����̨ǩ���Ŷ� Number;
    n_ԤԼ�Һ�       Number;
    n_��Ч����       Number; --��Ч����û�в������õ���
    n_�Һ����ɶ���   Number;
    n_Count          Number;
    n_��id           ����ɿ����.Id%Type;
    d_�˺�ʱ��       Date;
    v_����Ա���     ��Ա��.���%Type;
    v_����Ա����     ��Ա��.����%Type;
    v_������λ       ������λ�ҺŻ���.������λ%Type;
    n_ԤԼ״̬       ���˹Һż�¼.ԤԼ%Type;
    v_Temp           Varchar2(100);
    d_�Ǽ�ʱ��       ���˹Һż�¼.�Ǽ�ʱ��%Type;
    v_�ű�           ���˹Һż�¼.�ű�%Type;
    n_����           ���˹Һż�¼.����%Type;
    d_ԤԼʱ��       ���˹Һż�¼.ԤԼʱ��%Type;
    n_������λ����   Number(18);
    n_ԤԼ���ɶ���   Number;
    n_��¼����       Number;
    n_״̬           Number;
    n_�˺�����       Number(3);
    n_�Һ�id         ���˹Һż�¼.Id%Type;
    n_����           ������ü�¼.���ʷ���%Type;
    n_��¼id         �ٴ������¼.Id%Type;
  
    n_�ѽ���   Number;
    n_���ض�   �������.�������%Type;
    n_Ԥ��֧�� Number(3);
    n_����֧�� Number(3);
    Function Zl_����Ա
    (
      Type_In     Integer,
      Splitstr_In Varchar2
    ) Return Varchar2 As
      n_Step Number(18);
      v_Sub  Varchar2(1000);
      --Type_In:0-��ȡȱʡ����ID;1-��ȡ����Ա���;2-��ȡ����Ա����
      -- SplitStr:��ʽΪ:����ID,��������;��ԱID,��Ա���,��Ա����(��Zl_Identity��ȡ��)
    Begin
      If Type_In = 0 Then
        --ȱʡ����
        n_Step := Instr(Splitstr_In, ',');
        v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 1 Then
        --����Ա����
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 2 Then
        --����Ա����
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        Return v_Sub;
      End If;
    End;
  Begin
    --����ID,��������;��ԱID,��Ա���,��Ա����
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Error := '��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
      Raise Err_Custom;
    End If;
    v_����Ա��� := Zl_����Ա(1, v_Temp);
    v_����Ա���� := Zl_����Ա(2, v_Temp);
  
    n_��id := Zl_Get��id(v_����Ա����);
  
    d_�˺�ʱ�� := �˺�ʱ��_In;
    If d_�˺�ʱ�� Is Null Then
      d_�˺�ʱ�� := Sysdate;
    End If;
  
    --�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ����
    Begin
      Select Decode(��¼����, 2, 1, 0), ��¼����, �Ǽ�ʱ��, �ű�, ����, Nvl(ԤԼʱ��, ����ʱ��), ������λ, Nvl(ԤԼ, 0), Decode(��¼״̬, 0, 1, 0), �����¼id
      Into n_ԤԼ�Һ�, n_��¼����, d_�Ǽ�ʱ��, v_�ű�, n_����, d_ԤԼʱ��, v_������λ, n_ԤԼ״̬, n_��Ч����, n_��¼id
      From ���˹Һż�¼
      Where NO = ���ݺ�_In And ��¼״̬ In (0, 1) And Rownum < 2;
    Exception
      When Others Then
        n_ԤԼ�Һ� := -1;
    End;
  
    If n_ԤԼ�Һ� = -1 Then
      v_Error := '���ݿ����Ѿ����˺Ż򵥾��������!';
      Raise Err_Custom;
    End If;
  
    Begin
      Select Nvl(���ʷ���, 0), Decode(Sign(Nvl(����id, 0)), 0, 0, 1)
      Into n_����, n_�ѽ���
      From ������ü�¼
      Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ In (1, 3) And Rownum < 2;
    Exception
      When Others Then
        n_����   := 0;
        n_�ѽ��� := 0;
    End;
  
    --ԤԼ����Ƿ���Ӻ�����λ����
    --��������˺�����λ���� ��
    Select Count(0) Into n_������λ���� From �ٴ�����Һſ��Ƽ�¼ Where ���� = 1 And ���� = 1 And Rownum < 2;
    --���¹Һ����״̬
    n_�˺����� := Zl_To_Number(zl_GetSysParameter('�����������Һ�', 1111));
    If n_�˺����� = 0 Then
      Update �ٴ�������ſ��� Set �Һ�״̬ = 4 Where ��¼id = n_��¼id And (��� = n_���� Or ��ע = To_Char(n_����));
    Else
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 0, ���� = Null, ���� = Null, ����Ա���� = Null, ����վ���� = Null
      Where ��¼id = n_��¼id And (��� = n_���� Or ��ע = To_Char(n_����));
    End If;
    If Nvl(n_ԤԼ�Һ�, 0) = 1 Or Nvl(n_��Ч����, 0) = 1 Then
      If Nvl(n_��Ч����, 0) = 0 Then
        --N���ڲ���ȡ��ԤԼ��
        n_Count := Zl_To_Number(zl_GetSysParameter('N���ڲ���ȡ��ԤԼ��', 1111));
        If n_Count <> 0 Then
          If Trunc(Sysdate - n_Count) < d_�Ǽ�ʱ�� Then
            v_Error := '�����˵�ԤԼ��' || To_Char(Trunc(Sysdate - n_Count), 'yyyy-mm-dd') || '��ǰ��ԤԼ��!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      n_״̬ := Case n_��Ч����
                When 1 Then
                 0
                Else
                 1
              End;
      --������Լ��
      Open c_Registinfo(n_״̬, 2, n_��Ч����);
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update ���˹ҺŻ���
      Set ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬, �ѹ��� = Nvl(�ѹ���, 0) - Decode(n_ԤԼ״̬, 0, 1, 0)
      Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
            Nvl(ҽ������, 'ҽ��') = Nvl(r_Registrow.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
            (���� = r_Registrow.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��, �ѹ���)
        Values
          (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
           Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -n_ԤԼ״̬, Decode(n_ԤԼ״̬, 0, 1, 0));
      End If;
    
      Update �ٴ������¼
      Set ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬, �ѹ��� = Nvl(�ѹ���, 0) - Decode(n_ԤԼ״̬, 0, 1, 0)
      Where ID = n_��¼id;
      Close c_Registinfo;
    
      If Nvl(n_��Ч����, 0) = 0 Then
        --ɾ��������ü�¼
        Delete From ������ü�¼ Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 0;
        --���ԤԼ���ɶ���ʱ��Ҫ�������
        n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
        If Nvl(n_�Һ����ɶ���, 0) = 1 Then
          n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
          If Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
            --Ҫɾ������
            For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
              Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
            End Loop;
          End If;
        End If;
      End If;
    Else
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    
      --���¹Һ����״̬
    
      --���˾���״̬
      Select ����id
      Into n_����id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And ��� = 1;
    
      If n_����id Is Not Null Then
        Update ������Ϣ Set ����״̬ = 0, �������� = Null Where ����id = n_����id;
        --ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ��
      End If;
    
      --������ü�¼
      Insert Into ������ü�¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ,
         ����, ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��,
         �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
        Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
               �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
               v_����Ա���, v_����Ա����, ����ʱ��, d_�˺�ʱ��, n_����id,
               Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����, ժҪ,
               Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
    
      --ԭʼ��¼
      If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
        Update ������ü�¼
        Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
        Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
      Else
        Update ������ü�¼ Set ��¼״̬ = 3 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
      End If;
    
      n_ԭ����id := 0;
      If n_���� = 0 Then
        --��ȡ����ID
        Select Nvl(����id, 0)
        Into n_ԭ����id
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And Rownum < 2;
      End If;
    
      If n_���� = 1 Then
        --����
        For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                     From ������ü�¼
                     Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And Nvl(���ʷ���, 0) = 1) Loop
          --�������
          Update �������
          Set ������� = Nvl(�������, 0) - Nvl(c_����.ʵ�ս��, 0)
          Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1
          Returning ������� Into n_���ض�;
        
          If Sql%RowCount = 0 Then
            Insert Into �������
              (����id, ����, ����, �������, Ԥ�����)
            Values
              (n_����id, 1, 1, -1 * Nvl(c_����.ʵ�ս��, 0), 0);
            n_���ض� := Nvl(c_����.ʵ�ս��, 0);
          End If;
          If Nvl(n_���ض�, 0) = 0 Then
            Delete �������
            Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
          --����δ�����
          Update ����δ�����
          Set ��� = Nvl(���, 0) - Nvl(c_����.ʵ�ս��, 0)
          Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
                Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And
                ������Ŀid + 0 = c_����.������Ŀid And ��Դ;�� + 0 = 1;
        
          If Sql%RowCount = 0 Then
            Insert Into ����δ�����
              (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
            Values
              (n_����id, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, -1 * Nvl(c_����.ʵ�ս��, 0));
          End If;
        End Loop;
        Delete ����δ�����
        Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���, 0) = 0 And ��Դ;�� + 0 = 1;
      End If;
    
      If n_���� = 0 Then
        Begin
          Select 1
          Into n_Ԥ��֧��
          From ����Ԥ����¼
          Where Mod(��¼����, 10) = 1 And ��¼״̬ = 1 And ����id = n_ԭ����id And Rownum < 2;
        Exception
          When Others Then
            n_Ԥ��֧�� := 0;
        End;
        Begin
          Select 1
          Into n_����֧��
          From ����Ԥ����¼
          Where Mod(��¼����, 10) = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id And Rownum < 2;
        Exception
          When Others Then
            n_����֧�� := 0;
        End;
        If n_Ԥ��֧�� = 1 And n_����֧�� = 1 Then
          v_Error := '���ܴ�����ֽ��㷽ʽ,���鴫����˺ŵ����Ƿ���ȷ!';
          Raise Err_Custom;
        End If;
        If n_Ԥ��֧�� = 1 Then
          --ԭ���˻�Ԥ��
          If Nvl(Ԥ��id_In, 0) = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
               ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                     d_�˺�ʱ��, v_����Ա����, v_����Ա���, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� In (1, 11) And ����id = n_ԭ����id And Nvl(��Ԥ��, 0) <> 0;
          Else
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
               ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select Ԥ��id_In, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_�˺�ʱ��,
                     v_����Ա����, v_����Ա���, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� In (1, 11) And ����id = n_ԭ����id And Nvl(��Ԥ��, 0) <> 0;
          End If;
          --������Ԥ�����
          For c_Ԥ�� In (Select ����id, Ԥ�����, -1 * Sum(Nvl(��Ԥ��, 0)) As ��Ԥ��
                       From ����Ԥ����¼
                       Where ��¼���� In (1, 11) And ����id = n_����id
                       Group By ����id, Ԥ�����) Loop
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(c_Ԥ��.��Ԥ��, 0)
            Where ����id = c_Ԥ��.����id And ���� = Nvl(c_Ԥ��.Ԥ�����, 2) And ���� = 1
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into �������
                (����id, Ԥ�����, ����, ����)
              Values
                (c_Ԥ��.����id, Nvl(c_Ԥ��.��Ԥ��, 0), 1, Nvl(c_Ԥ��.Ԥ�����, 2));
              n_����ֵ := Nvl(c_Ԥ��.��Ԥ��, 0);
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = c_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
          End Loop;
        Else
          If Nvl(Ԥ��id_In, 0) = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ������ˮ��, ����˵��,
               ������λ, �������, �����id, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_�˺�ʱ��, v_����Ա���, v_����Ա����, -��Ԥ��,
                     n_����id, n_��id, ������ˮ��_In, ����˵��_In, ������λ, n_����id, �����id, ��������
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
          Else
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ������ˮ��, ����˵��,
               ������λ, �������, �����id, ��������)
              Select Ԥ��id_In, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_�˺�ʱ��, v_����Ա���, v_����Ա����, -��Ԥ��, n_����id,
                     n_��id, ������ˮ��_In, ����˵��_In, ������λ, n_����id, �����id, ��������
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
          End If;
          Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
        End If;
        --�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�)
        --�����һ�εĴ�ӡ������ȡ
        Select Max(ID)
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And b.�������� = 4 And b.No = ���ݺ�_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
        If n_��ӡid Is Not Null Then
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_�˺�ʱ��, v_����Ա����, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And ���� = 1;
        End If;
      End If;
    
      --��ػ��ܱ�Ĵ���
    
      --���˹ҺŻ���
      Open c_Registinfo(1, 1);
      Fetch c_Registinfo
        Into r_Registrow;
    
      If c_Registinfo%RowCount = 0 Then
        --ֻ�ղ�����ʱ�޺ű�,������
        Close c_Registinfo;
      Else
      
        --��Ҫȷ���Ƿ�ԤԼ�Һ�
        --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ���
        --2.����������Һ�,��ֻ���ѹ���
        Begin
          Select Decode(ԤԼ, Null, 0, 0, 0, 1), ִ��״̬
          Into n_ԤԼ�Һ�, n_ִ��״̬
          From ���˹Һż�¼
          Where NO = ���ݺ�_In And ��¼״̬ = 1 And Rownum = 1;
        Exception
          When Others Then
            n_ԤԼ�Һ� := 0;
        End;
        --0-�ȴ�����,1-��ɾ���,2-���ھ���,-1���Ϊ������
        If n_ִ��״̬ > 0 Then
          If n_ִ��״̬ = 1 Then
            v_Error := '�ò����Ѿ���ɾ���,�������˺�!';
          Else
            v_Error := '�ò������ھ���, �����˺�!';
          End If;
          Raise Err_Custom;
        End If;
      
        Update ���˹ҺŻ���
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ״̬, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬
        Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
              Nvl(ҽ������, 'ҽ��') = Nvl(r_Registrow.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
              (���� = r_Registrow.���� Or ���� Is Null);
      
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���)
          Values
            (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
             Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1 * n_ԤԼ�Һ�);
        End If;
      
        Update �ٴ������¼
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ״̬, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬
        Where ID = n_��¼id;
        Close c_Registinfo;
      End If;
    
      --��Ա�ɿ����(���������ʻ��ȵĽ�����,�����˳�Ԥ����)
      If n_���� = 0 Then
        For r_Opermoney In c_Opermoney Loop
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) + (-1 * r_Opermoney.��Ԥ��)
          Where �տ�Ա = v_����Ա���� And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1
          Returning ��� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (v_����Ա����, r_Opermoney.���㷽ʽ, 1, -1 * r_Opermoney.��Ԥ��);
            n_����ֵ := r_Opermoney.��Ԥ��;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From ��Ա�ɿ����
            Where �տ�Ա = v_����Ա���� And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
          End If;
        End Loop;
      End If;
    
      n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
      If n_�Һ����ɶ��� <> 0 Then
        n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
        If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
          --Ҫɾ������
          For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
            Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
          End Loop;
        End If;
      End If;
    
      --ҽ�������ľ���ǼǼ�¼
      Delete From ����ǼǼ�¼
      Where (����id, ��ҳid, ����ʱ��) In (Select ����id, ��ҳid, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In);
    End If;
  
    If Nvl(n_��Ч����, 0) = 0 Then
      Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼״̬ = 1;
      If Sql%NotFound Then
        v_Error := 'δ�ҵ��Һŵ���,����!';
        Raise Err_Custom;
      End If;
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, �����¼id)
        Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_�˺�ʱ��, ����ʱ��,
               v_����Ա���, v_����Ա����, ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��_In, ����˵��_In, ������λ, �����¼id
        From ���˹Һż�¼
        Where NO = ���ݺ�_In And ��¼״̬ = 3;
    End If;
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 2, ���ݺ�_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_�Һ�id, ���ݺ�_In);
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Error || '+[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

Begin
  n_�Һ��Ű�ģʽ := To_Number(Substr(Nvl(zl_GetSysParameter('�Һ��Ű�ģʽ'), 0), 1, 1));
  If n_�Һ��Ű�ģʽ = 1 Then
    --������Ű�ģʽ
    Zl_���������Һ�_����_Delete(���ݺ�_In, ������ˮ��_In, ����˵��_In, �˺�ʱ��_In, Ԥ��id_In);
  Else
    v_Temp := zl_GetSysParameter(256);
    If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
      Null;
    Else
      Begin
        d_����ʱ�� := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
      Exception
        When Others Then
          d_����ʱ�� := Null;
      End;
    End If;
    --����ID,��������;��ԱID,��Ա���,��Ա����
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Error := '��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
      Raise Err_Custom;
    End If;
    v_����Ա��� := Zl_����Ա(1, v_Temp);
    v_����Ա���� := Zl_����Ա(2, v_Temp);
  
    n_��id := Zl_Get��id(v_����Ա����);
  
    d_�˺�ʱ�� := �˺�ʱ��_In;
    If d_�˺�ʱ�� Is Null Then
      d_�˺�ʱ�� := Sysdate;
    End If;
  
    --�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ����
    Begin
      Select Decode(��¼����, 2, 1, 0), ��¼����, �Ǽ�ʱ��, �ű�, ����, Nvl(ԤԼʱ��, ����ʱ��), ������λ, Nvl(ԤԼ, 0), Decode(��¼״̬, 0, 1, 0)
      Into n_ԤԼ�Һ�, n_��¼����, d_�Ǽ�ʱ��, v_�ű�, n_����, d_ԤԼʱ��, v_������λ, n_ԤԼ״̬, n_��Ч����
      From ���˹Һż�¼
      Where NO = ���ݺ�_In And ��¼״̬ In (0, 1) And Rownum <= 1;
    Exception
      When Others Then
        n_ԤԼ�Һ� := -1;
    End;
  
    If n_ԤԼ�Һ� = -1 Then
      v_Error := '���ݿ����Ѿ����˺Ż򵥾��������!';
      Raise Err_Custom;
    End If;
  
    Begin
      Select Nvl(���ʷ���, 0), Decode(Sign(Nvl(����id, 0)), 0, 0, 1)
      Into n_����, n_�ѽ���
      From ������ü�¼
      Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ In (1, 3) And Rownum < 2;
    Exception
      When Others Then
        n_����   := 0;
        n_�ѽ��� := 0;
    End;
  
    Begin
      Select a.Id Into n_����id From �ҺŰ��� A Where a.���� = v_�ű�;
    Exception
      When Others Then
        n_����id := -1;
    End;
  
    Begin
      Select ID
      Into n_�ƻ�id
      From �ҺŰ��żƻ�
      Where ����id = n_����id And ���ʱ�� Is Not Null And
            Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) =
            (Select Max(a.��Чʱ��) As ��Ч
             From �ҺŰ��żƻ� A
             Where a.���ʱ�� Is Not Null And d_ԤԼʱ�� Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                   Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.����id = n_����id) And
            d_ԤԼʱ�� Between Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
            Nvl(ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd'));
    Exception
      When Others Then
        n_�ƻ�id := 0;
    End;
    If Nvl(n_�ƻ�id, 0) = 0 Then
      Select Decode(To_Char(d_ԤԼʱ��, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ���
      Where ID = n_����id;
    Else
      Select Decode(To_Char(d_ԤԼʱ��, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ��żƻ�
      Where ID = n_�ƻ�id;
    End If;
  
    If v_ʱ��� Is Not Null And d_����ʱ�� Is Not Null Then
      --����Ƿ��ģʽ�ҺŰ���
      Select To_Date(To_Char(d_ԤԼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             To_Date(To_Char(d_ԤԼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
      Into d_��鿪ʼʱ��, d_������ʱ��
      From ʱ���
      Where ʱ��� = v_ʱ��� And վ�� Is Null And ���� Is Null;
      If d_��鿪ʼʱ�� > d_������ʱ�� Then
        d_������ʱ�� := d_������ʱ�� + 1;
      End If;
      If d_������ʱ�� > d_����ʱ�� Then
        --��ȡ�����¼id
        Select Max(a.Id)
        Into n_�����¼id
        From �ٴ������¼ A, �ٴ������Դ B
        Where a.��Դid = b.Id And b.���� = v_�ű� And �ϰ�ʱ�� = v_ʱ��� And d_ԤԼʱ�� Between ��ʼʱ�� And ��ֹʱ��;
      End If;
    End If;
  
    --ԤԼ����Ƿ���Ӻ�����λ����
    --��������˺�����λ���� ��
    Select Count(0) Into n_������λ���� From ������λ���ſ��� Where Rownum = 1;
    --���¹Һ����״̬
    n_�˺����� := Zl_To_Number(zl_GetSysParameter('�����������Һ�', 1111));
    If n_�˺����� = 0 Then
      Update �Һ����״̬
      Set ״̬ = 4
      Where ���� = v_�ű� And ��� = n_���� And ���� Between Trunc(d_ԤԼʱ��) And Trunc(d_ԤԼʱ�� + 1) - 1 / 24 / 60 / 60;
    Else
      Delete �Һ����״̬
      Where ���� = v_�ű� And ��� = n_���� And ���� Between Trunc(d_ԤԼʱ��) And Trunc(d_ԤԼʱ�� + 1) - 1 / 24 / 60 / 60;
    End If;
    If Nvl(n_ԤԼ�Һ�, 0) = 1 Or Nvl(n_��Ч����, 0) = 1 Then
      If Nvl(n_��Ч����, 0) = 0 Then
        --N���ڲ���ȡ��ԤԼ��
        n_Count := Zl_To_Number(zl_GetSysParameter('N���ڲ���ȡ��ԤԼ��', 1111));
        If n_Count <> 0 Then
          If Trunc(Sysdate - n_Count) < d_�Ǽ�ʱ�� Then
            v_Error := '�����˵�ԤԼ��' || To_Char(Trunc(Sysdate - n_Count), 'yyyy-mm-dd') || '��ǰ��ԤԼ��!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      n_״̬ := Case n_��Ч����
                When 1 Then
                 0
                Else
                 1
              End;
      --������Լ��
      Open c_Registinfo(n_״̬, 2, n_��Ч����);
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update ���˹ҺŻ���
      Set ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬, �ѹ��� = Nvl(�ѹ���, 0) - Decode(n_ԤԼ״̬, 0, 1, 0)
      Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
            Nvl(ҽ������, 'ҽ��') = Nvl(r_Registrow.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
            (���� = r_Registrow.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��, �ѹ���)
        Values
          (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
           Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -n_ԤԼ״̬, Decode(n_ԤԼ״̬, 0, 1, 0));
      End If;
    
      If n_�����¼id Is Not Null Then
        Update �ٴ������¼
        Set ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬, �ѹ��� = Nvl(�ѹ���, 0) - Decode(n_ԤԼ״̬, 0, 1, 0)
        Where ID = n_�����¼id And Nvl(��Լ��, 0) > 0;
        Update �ٴ�������ſ��� Set �Һ�״̬ = Null, ����Ա���� = Null Where ��¼id = n_�����¼id And ��� = n_����;
      End If;
      If Nvl(n_������λ����, 0) <> 0 And v_������λ Is Not Null And Nvl(n_ԤԼ״̬, 0) <> 0 Then
        Update ������λ�ҺŻ���
        Set ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬, �ѽ��� = Nvl(�ѽ���, 0) - Decode(n_ԤԼ״̬, 0, 1, 0)
        Where ���� = Trunc(r_Registrow.����ʱ��) And (���� = r_Registrow.���� Or ���� Is Null) And ������λ = v_������λ And
              ��� = Nvl(n_����, 0);
        If Sql%RowCount = 0 Then
          Insert Into ������λ�ҺŻ���
            (����, ����, ��Լ��, ������λ, ���, �ѽ���)
          Values
            (Trunc(r_Registrow.����ʱ��), r_Registrow.����, -n_ԤԼ״̬, v_������λ, Nvl(n_����, 0), -decode(n_ԤԼ״̬, 0, 1, 0));
        End If;
      End If;
      Close c_Registinfo;
    
      If Nvl(n_��Ч����, 0) = 0 Then
        --ɾ��������ü�¼
        Delete From ������ü�¼ Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 0;
        --���ԤԼ���ɶ���ʱ��Ҫ�������
        n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
        If Nvl(n_�Һ����ɶ���, 0) = 1 Then
          n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
          If Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
            --Ҫɾ������
            For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
              Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
            End Loop;
          End If;
        End If;
      End If;
    Else
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    
      --���¹Һ����״̬
    
      --���˾���״̬
      Select ����id
      Into n_����id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And ��� = 1;
    
      If n_����id Is Not Null Then
        Update ������Ϣ Set ����״̬ = 0, �������� = Null Where ����id = n_����id;
        --ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ��
      End If;
    
      --������ü�¼
      Insert Into ������ü�¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ,
         ����, ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��,
         �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
        Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
               �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
               v_����Ա���, v_����Ա����, ����ʱ��, d_�˺�ʱ��, n_����id,
               Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����, ժҪ,
               Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
    
      --ԭʼ��¼
      If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
        Update ������ü�¼
        Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
        Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
      Else
        Update ������ü�¼ Set ��¼״̬ = 3 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
      End If;
    
      n_ԭ����id := 0;
      If n_���� = 0 Then
        --��ȡ����ID
        Select Nvl(����id, 0)
        Into n_ԭ����id
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And Rownum < 2;
      End If;
    
      If n_���� = 1 Then
        --����
        For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                     From ������ü�¼
                     Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And Nvl(���ʷ���, 0) = 1) Loop
          --�������
          Update �������
          Set ������� = Nvl(�������, 0) - Nvl(c_����.ʵ�ս��, 0)
          Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1
          Returning ������� Into n_���ض�;
        
          If Sql%RowCount = 0 Then
            Insert Into �������
              (����id, ����, ����, �������, Ԥ�����)
            Values
              (n_����id, 1, 1, -1 * Nvl(c_����.ʵ�ս��, 0), 0);
            n_���ض� := Nvl(c_����.ʵ�ս��, 0);
          End If;
          If Nvl(n_���ض�, 0) = 0 Then
            Delete �������
            Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
          --����δ�����
          Update ����δ�����
          Set ��� = Nvl(���, 0) - Nvl(c_����.ʵ�ս��, 0)
          Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
                Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And
                ������Ŀid + 0 = c_����.������Ŀid And ��Դ;�� + 0 = 1;
        
          If Sql%RowCount = 0 Then
            Insert Into ����δ�����
              (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
            Values
              (n_����id, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, -1 * Nvl(c_����.ʵ�ս��, 0));
          End If;
        End Loop;
        Delete ����δ�����
        Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���, 0) = 0 And ��Դ;�� + 0 = 1;
      End If;
    
      If n_���� = 0 Then
        Begin
          Select 1
          Into n_Ԥ��֧��
          From ����Ԥ����¼
          Where Mod(��¼����, 10) = 1 And ��¼״̬ = 1 And ����id = n_ԭ����id And Rownum < 2;
        Exception
          When Others Then
            n_Ԥ��֧�� := 0;
        End;
      
        Begin
          Select 1
          Into n_����֧��
          From ����Ԥ����¼
          Where Mod(��¼����, 10) = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id And Rownum < 2;
        Exception
          When Others Then
            n_����֧�� := 0;
        End;
      
        If n_Ԥ��֧�� = 1 And n_����֧�� = 1 Then
          v_Error := '���ܴ�����ֽ��㷽ʽ,���鴫����˺ŵ����Ƿ���ȷ!';
          Raise Err_Custom;
        End If;
      
        If n_Ԥ��֧�� = 1 Then
          --ԭ���˻�Ԥ��
          If Nvl(Ԥ��id_In, 0) = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
               ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                     d_�˺�ʱ��, v_����Ա����, v_����Ա���, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� In (1, 11) And ����id = n_ԭ����id And Nvl(��Ԥ��, 0) <> 0;
          Else
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
               ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select Ԥ��id_In, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_�˺�ʱ��,
                     v_����Ա����, v_����Ա���, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� In (1, 11) And ����id = n_ԭ����id And Nvl(��Ԥ��, 0) <> 0;
          End If;
          --������Ԥ�����
          For c_Ԥ�� In (Select ����id, Ԥ�����, -1 * Sum(Nvl(��Ԥ��, 0)) As ��Ԥ��
                       From ����Ԥ����¼
                       Where ��¼���� In (1, 11) And ����id = n_����id
                       Group By ����id, Ԥ�����) Loop
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(c_Ԥ��.��Ԥ��, 0)
            Where ����id = c_Ԥ��.����id And ���� = Nvl(c_Ԥ��.Ԥ�����, 2) And ���� = 1
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into �������
                (����id, Ԥ�����, ����, ����)
              Values
                (c_Ԥ��.����id, Nvl(c_Ԥ��.��Ԥ��, 0), 1, Nvl(c_Ԥ��.Ԥ�����, 2));
              n_����ֵ := Nvl(c_Ԥ��.��Ԥ��, 0);
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = c_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
          End Loop;
        Else
          If Nvl(Ԥ��id_In, 0) = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ������ˮ��, ����˵��,
               ������λ, �������, �����id, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_�˺�ʱ��, v_����Ա���, v_����Ա����, -��Ԥ��,
                     n_����id, n_��id, ������ˮ��_In, ����˵��_In, ������λ, n_����id, �����id, ��������
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
          Else
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ������ˮ��, ����˵��,
               ������λ, �������, �����id, ��������)
              Select Ԥ��id_In, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_�˺�ʱ��, v_����Ա���, v_����Ա����, -��Ԥ��, n_����id,
                     n_��id, ������ˮ��_In, ����˵��_In, ������λ, n_����id, �����id, ��������
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
          End If;
        
          Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
        End If;
        --�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�)
        --�����һ�εĴ�ӡ������ȡ
        Select Max(ID)
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And b.�������� = 4 And b.No = ���ݺ�_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
        If n_��ӡid Is Not Null Then
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_�˺�ʱ��, v_����Ա����, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And ���� = 1;
        End If;
      End If;
    
      --��ػ��ܱ�Ĵ���
    
      --���˹ҺŻ���
      Open c_Registinfo(1, 1);
      Fetch c_Registinfo
        Into r_Registrow;
    
      If c_Registinfo%RowCount = 0 Then
        --ֻ�ղ�����ʱ�޺ű�,������
        Close c_Registinfo;
      Else
      
        --��Ҫȷ���Ƿ�ԤԼ�Һ�
        --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ���
        --2.����������Һ�,��ֻ���ѹ���
        Begin
          Select Decode(ԤԼ, Null, 0, 0, 0, 1), ִ��״̬
          Into n_ԤԼ�Һ�, n_ִ��״̬
          From ���˹Һż�¼
          Where NO = ���ݺ�_In And ��¼״̬ = 1 And Rownum = 1;
        Exception
          When Others Then
            n_ԤԼ�Һ� := 0;
        End;
        --0-�ȴ�����,1-��ɾ���,2-���ھ���,-1���Ϊ������
        If n_ִ��״̬ > 0 Then
          If n_ִ��״̬ = 1 Then
            v_Error := '�ò����Ѿ���ɾ���,�������˺�!';
          Else
            v_Error := '�ò������ھ���, �����˺�!';
          End If;
          Raise Err_Custom;
        End If;
      
        Update ���˹ҺŻ���
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ״̬, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬
        Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
              Nvl(ҽ������, 'ҽ��') = Nvl(r_Registrow.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
              (���� = r_Registrow.���� Or ���� Is Null);
      
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���)
          Values
            (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
             Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1 * n_ԤԼ�Һ�);
        End If;
        If n_�����¼id Is Not Null Then
          Update �ٴ������¼
          Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ״̬, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬
          Where ID = n_�����¼id And Nvl(��Լ��, 0) > 0;
          Update �ٴ�������ſ��� Set �Һ�״̬ = Null, ����Ա���� = Null Where ��¼id = n_�����¼id And ��� = n_����;
        End If;
        If Nvl(n_������λ����, 0) <> 0 And v_������λ Is Not Null And Nvl(n_ԤԼ״̬, 0) <> 0 Then
          Update ������λ�ҺŻ���
          Set �ѽ��� = Nvl(�ѽ���, 0) - 1, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
          Where ���� = Trunc(r_Registrow.����ʱ��) And (���� = r_Registrow.���� Or ���� Is Null) And ������λ = v_������λ And
                ��� = Nvl(n_����, 0);
          If Sql%RowCount = 0 Then
            Insert Into ������λ�ҺŻ���
              (����, ����, ��Լ��, ������λ, �ѽ���, ���)
            Values
              (Trunc(r_Registrow.����ʱ��), r_Registrow.����, -1, v_������λ, -1 * n_ԤԼ�Һ�, Nvl(n_����, 0));
          End If;
        End If;
        Close c_Registinfo;
      End If;
    
      --��Ա�ɿ����(���������ʻ��ȵĽ�����,�����˳�Ԥ����)
      If n_���� = 0 Then
        For r_Opermoney In c_Opermoney Loop
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) + (-1 * r_Opermoney.��Ԥ��)
          Where �տ�Ա = v_����Ա���� And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1
          Returning ��� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (v_����Ա����, r_Opermoney.���㷽ʽ, 1, -1 * r_Opermoney.��Ԥ��);
            n_����ֵ := r_Opermoney.��Ԥ��;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From ��Ա�ɿ����
            Where �տ�Ա = v_����Ա���� And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
          End If;
        End Loop;
      End If;
    
      n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
      If n_�Һ����ɶ��� <> 0 Then
        n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
        If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
          --Ҫɾ������
          For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
            Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
          End Loop;
        End If;
      End If;
    
      --ҽ�������ľ���ǼǼ�¼
      Delete From ����ǼǼ�¼
      Where (����id, ��ҳid, ����ʱ��) In (Select ����id, ��ҳid, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In);
    End If;
  
    If Nvl(n_��Ч����, 0) = 0 Then
      Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼״̬ = 1;
      If Sql%NotFound Then
        v_Error := 'δ�ҵ��Һŵ���,����!';
        Raise Err_Custom;
      End If;
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ)
        Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_�˺�ʱ��, ����ʱ��,
               v_����Ա���, v_����Ա����, ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��_In, ����˵��_In, ������λ
        From ���˹Һż�¼
        Where NO = ���ݺ�_In And ��¼״̬ = 3;
    End If;
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 2, ���ݺ�_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_�Һ�id, ���ݺ�_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Error || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���������Һ�_Delete;
/

--125588:����,2018-05-14,������ʵ������Ϊ0������
Create Or Replace Procedure Zl_��ҩ���_Update
(
  ҩƷid_In         In ҩƷ���.ҩƷid%Type,
  ����_In           In �շ���ĿĿ¼.����%Type,
  ���_In           In �շ���ĿĿ¼.���%Type,
  ����_In           In �շ���ĿĿ¼.����%Type := Null,
  Ʒ��_In           In �շ���Ŀ����.����%Type := Null,
  ƴ��_In           In �շ���Ŀ����.����%Type := Null,
  ���_In           In �շ���Ŀ����.����%Type := Null,
  ������_In         In �շ���Ŀ����.����%Type := Null,
  ��ʶ��_In         In ҩƷ���.��ʶ��%Type := Null,
  ҩƷ��Դ_In       In ҩƷ���.ҩƷ��Դ%Type := Null,
  ��׼�ĺ�_In       In ҩƷ���.��׼�ĺ�%Type := Null,
  ע���̱�_In       In ҩƷ���.ע���̱�%Type := Null,
  �ۼ۵�λ_In       In �շ���ĿĿ¼.���㵥λ%Type := Null,
  ����ϵ��_In       In ҩƷ���.����ϵ��%Type := Null,
  ���ﵥλ_In       In ҩƷ���.���ﵥλ%Type := Null,
  �����װ_In       In ҩƷ���.�����װ%Type := Null,
  ҩ�ⵥλ_In       In ҩƷ���.ҩ�ⵥλ%Type := Null,
  ҩ���װ_In       In ҩƷ���.ҩ���װ%Type := Null,
  ���쵥λ_In       In ҩƷ���.���쵥λ%Type := 1,
  ���췧ֵ_In       In ҩƷ���.���췧ֵ%Type := Null,
  �Ƿ���_In       In �շ���ĿĿ¼.�Ƿ���%Type := Null,
  ָ��������_In     In ҩƷ���.ָ��������%Type := Null,
  ����_In           In ҩƷ���.����%Type := 95,
  ָ�����ۼ�_In     In ҩƷ���.ָ�����ۼ�%Type := Null,
  �ӳ���_In         In ҩƷ���.�ӳ���%Type := Null,
  ����ѱ���_In     In ҩƷ���.����ѱ���%Type := Null,
  ҩ�ۼ���_In       In ҩƷ���.ҩ�ۼ���%Type := Null,
  ��������_In       In �շ���ĿĿ¼.��������%Type := Null,
  �������_In       In �շ���ĿĿ¼.�������%Type := Null,
  Gmp��֤_In        In ҩƷ���.Gmp��֤%Type := 0,
  �б�ҩƷ_In       In ҩƷ���.�б�ҩƷ%Type := 0,
  ���ηѱ�_In       In �շ���ĿĿ¼.���ηѱ�%Type := 0,
  סԺ�ɷ����_In   In ҩƷ���.סԺ�ɷ����%Type := 0,
  ҩ�����_In       In ҩƷ���.ҩ�����%Type := Null,
  ҩ������_In       In ҩƷ���.ҩ������%Type := Null,
  ���Ч��_In       In ҩƷ���.���Ч��%Type := Null,
  ���������_In     In ҩƷ���.���������%Type := 0,
  �ɱ���_In         In ҩƷ���.�ɱ���%Type := 0,
  ��ǰ�ۼ�_In       In �շѼ�Ŀ.�ּ�%Type := 0,
  ����id_In         In �շѼ�Ŀ.������Ŀid%Type := Null,
  ��ͬ��λid_In     In ҩƷ���.��ͬ��λid%Type := Null,
  ˵��_In           In �շ���ĿĿ¼.˵��%Type := Null,
  ��̬����_In       In ҩƷ���.��̬����%Type := 0,
  ��ҩ����_In       In ҩƷ���.��ҩ����%Type := Null,
  ��ѡ��_In         In �շ���ĿĿ¼.��ѡ��%Type := Null,
  ��ֵ˰��_In       In ҩƷ���.��ֵ˰��%Type := Null,
  ����ҩ��_In       In ҩƷ���.����ҩ��%Type := Null,
  ��ҩ��̬_In       In ҩƷ���.��ҩ��̬%Type := Null,
  վ��_In           In �շ���ĿĿ¼.վ��%Type := Null,
  �Ƿ񳣱�_In       In ҩƷ���.�Ƿ񳣱�%Type := Null,
  ������Ŀ_In       In �շ���ĿĿ¼.������Ŀ%Type := Null,
  ����ɷ����_In   In ҩƷ���.����ɷ����%Type := 0,
  �ͻ���λ_In       ҩƷ���.�ͻ���λ%Type := Null,
  �ͻ���װ_In       ҩƷ���.�ͻ���װ%Type := Null,
  �Ƿ��ҩ_In       ҩƷ���.�Ƿ��ҩ%Type := Null,
  �Ƿ����۹���_In In ҩƷ���.�Ƿ����۹���%Type := Null,
  ��λ��_In         In ҩƷ���.��λ��%Type := Null,
  ԭ����_In         In ҩƷ���.ԭ����%Type := Null
) Is
  v_ҩ��id   ������ĿĿ¼.Id%Type;
  v_����     ������ĿĿ¼.����%Type;
  v_�Ƿ��� �շ���ĿĿ¼.�Ƿ���%Type; --������ҩƷ��ʱ��Ϊʱ�ۣ�ʱ��ҩƷֻ����δ������������޸�Ϊ���ۣ���������������޸Ķ������� 
  v_����     Number(2);
  Err_Notfind Exception;
  v_No           �շѼ�Ŀ.No%Type;
  v_Temp         �շ���ĿĿ¼.������Ŀ%Type;
  v_������Ŀ     �շ���ĿĿ¼.������Ŀ%Type;
  n_ָ�������   ҩƷ���.ָ�������%Type;
  n_ҩƷ�ϴ��ۼ� ҩƷ���.�ϴ��ۼ�%Type;
  n_���۽��     ҩƷ���.ʵ�ʽ��%Type;
  n_�շ�id       ҩƷ�շ���¼.Id%Type;
  n_��ͨ���С�� Number;
  n_���         Number(8);
  Classid        Number(18); --������
  v_Billno       ҩƷ�շ���¼.No%Type; --���۵���
  n_�۸�id       �շѼ�Ŀ.Id%Type;
  n_�շѼ�Ŀ�ּ� �շѼ�Ŀ.�ּ�%Type;
  n_ԭ��         ҩƷ�۸��¼.ԭ��%Type;
  n_ҩƷ�۸��¼ Number(1);
  v_���         �շ���ĿĿ¼.���%Type;
  --����->ʱ�ۺ����ҩƷ�۸��¼��ֵ

  Cursor c_Priceadjust Is
    Select s.ҩƷid, s.�ⷿid, Nvl(s.����, 0) As ����, s.�ϴι�Ӧ��id As ��Ӧ��id, s.�ϴ����� As ����, s.Ч��, s.�ϴβ��� As ����,
           Nvl(s.ʵ������, 0) As ʵ������, s.�ϴο��� As ����, Nvl(s.ʵ�ʽ��, 0) As ʵ�ʽ��, Nvl(s.ʵ�ʲ��, 0) As ʵ�ʲ��, Nvl(s.���ۼ�, 0) As ���ۼ�,
           s.ƽ���ɱ���, s.���Ч��, s.��׼�ĺ�, s.�ϴ��������� As ��������
    From ҩƷ��� S
    Where s.ҩƷid = ҩƷid_In And s.���� = 1 
    Order By s.ҩƷid, s.����, s.�ⷿid;

  r_Priceadjust c_Priceadjust%RowType;

Begin
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
  --ͨ������ 
  Select ID, ����
  Into v_ҩ��id, v_����
  From ������ĿĿ¼
  Where ID = (Select ҩ��id From ҩƷ��� Where ҩƷid = ҩƷid_In);
  --ȡԭʼ�Ķ������� 
  Select �Ƿ��� Into v_�Ƿ��� From �շ���ĿĿ¼ Where ID = ҩƷid_In;
  --�����Ϣ 
  Update �շ���ĿĿ¼
  Set ���� = ����_In, ���� = v_����, ��� = ���_In, ���� = ����_In, ���㵥λ = �ۼ۵�λ_In, �������� = ��������_In, ������� = �������_In, ���ηѱ� = ���ηѱ�_In,
      ������Ŀ = v_������Ŀ, ˵�� = ˵��_In, ��ѡ�� = ��ѡ��_In, վ�� = վ��_In
  Where ID = ҩƷid_In
  Returning ��� Into v_���;
  If Sql%RowCount = 0 Then
    Raise Err_Notfind;
  End If;
  n_ָ������� := (1 - 1 / (1 + �ӳ���_In / 100)) * 100;
  Update ҩƷ���
  Set ��ʶ�� = ��ʶ��_In, ҩƷ��Դ = ҩƷ��Դ_In, ��׼�ĺ� = ��׼�ĺ�_In, ע���̱� = ע���̱�_In, ����ϵ�� = ����ϵ��_In, ���ﵥλ = ���ﵥλ_In, �����װ = �����װ_In,
      סԺ��λ = ���ﵥλ_In, סԺ��װ = �����װ_In, ҩ�ⵥλ = ҩ�ⵥλ_In, ҩ���װ = ҩ���װ_In, ���쵥λ = ���쵥λ_In, ���췧ֵ = ���췧ֵ_In, ָ�������� = ָ��������_In,
      ���� = ����_In, ָ�����ۼ� = ָ�����ۼ�_In, ָ������� = n_ָ�������, ����ѱ��� = ����ѱ���_In, ҩ�ۼ��� = ҩ�ۼ���_In, סԺ�ɷ���� = סԺ�ɷ����_In,
      ҩ����� = ҩ�����_In, ҩ������ = ҩ������_In, ���Ч�� = ���Ч��_In, �б�ҩƷ = �б�ҩƷ_In, Gmp��֤ = Gmp��֤_In, ��������� = ���������_In,
      ��ͬ��λid = ��ͬ��λid_In, ��̬���� = ��̬����_In, ��ҩ���� = ��ҩ����_In, ��ֵ˰�� = ��ֵ˰��_In, ����ҩ�� = ����ҩ��_In, ��ҩ��̬ = ��ҩ��̬_In, �Ƿ񳣱� = �Ƿ񳣱�_In,
      ����ɷ���� = ����ɷ����_In, �ͻ���λ = �ͻ���λ_In, �ͻ���װ = �ͻ���װ_In, �ӳ��� = �ӳ���_In, �Ƿ��ҩ = �Ƿ��ҩ_In, �Ƿ����۹��� = �Ƿ����۹���_In,
      ��λ�� = ��λ��_In, ԭ���� = ԭ����_In
  Where ҩƷid = ҩƷid_In;

  --�����޸ģ�����ҩƷ������ҩ���г�ҩ��ʱ��ȱʡ�������Ϊ�����סԺ������޸Ĺ��ҩƷʱ�����ٸ��ݹ��ҩƷ�ķ���������ҩƷ�ķ������ 
  --������Ŀ�������ĸ��� 
  --select nvl(sum(distinct I.�������),0) into v_���� 
  --from �շ���ĿĿ¼ I,ҩƷ��� S 
  --where I.ID=S.ҩƷID and S.ҩ��ID=v_ҩ��ID; 
  --update ������ĿĿ¼ 
  --set �������=decode(v_����,0,0,1,1,2,2,3) 
  --where ID=v_ҩ��ID; 

  --�����Ĵ��� 
  If ������_In Is Null Then
    Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 1 And ���� = 3;
  Else
    Update �շ���Ŀ���� Set ���� = v_����, ���� = ������_In Where �շ�ϸĿid = ҩƷid_In And ���� = 1 And ���� = 3;
    If Sql%RowCount = 0 Then
      Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, v_����, 1, ������_In, 3);
    End If;
  End If;
  If Ʒ��_In Is Null Then
    Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 3;
  Else
    If ƴ��_In Is Null Then
      Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 1;
    Else
      Update �շ���Ŀ���� Set ���� = Ʒ��_In, ���� = ƴ��_In Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 1;
      If Sql%RowCount = 0 Then
        Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, Ʒ��_In, 3, ƴ��_In, 1);
      End If;
    End If;
    If ���_In Is Null Then
      Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 2;
    Else
      Update �շ���Ŀ���� Set ���� = Ʒ��_In, ���� = ���_In Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 2;
      If Sql%RowCount = 0 Then
        Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, Ʒ��_In, 3, ���_In, 2);
      End If;
    End If;
  End If;

  --������Ϣ������Ѿ��з�����������ֱ�Ӹ�����Щ��Ϣ 
  Select Nvl(Count(*), 0) Into v_���� From ҩƷ�շ���¼ Where ҩƷid = ҩƷid_In And Rownum < 2;
  If v_���� = 0 Then
    Update ҩƷ��� Set �ɱ��� = �ɱ���_In Where ҩƷid = ҩƷid_In;
    If ����id_In Is Not Null Then
      Update �շѼ�Ŀ
      Set �ּ� = ��ǰ�ۼ�_In, ������Ŀid = ����id_In, �䶯ԭ�� = 1, ����˵�� = '�޸Ķ���', ������ = User
      Where �շ�ϸĿid = ҩƷid_In And (Sysdate Between ִ������ And ��ֹ���� Or Sysdate >= ִ������ And ��ֹ���� Is Null) And �䶯ԭ�� = 1;
      If Sql%RowCount = 0 Then
        v_No := Nextno(9);
        Insert Into �շѼ�Ŀ
          (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
        Values
          (�շѼ�Ŀ_Id.Nextval, Null, ҩƷid_In, 0, ��ǰ�ۼ�_In, ����id_In, 1, '��������', User, Sysdate,
           To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
      End If;
    End If;
  Else
    --����ҵ������ʱ�������޸ļ۸��ǿ����޸�������Ŀ 
    Update �շѼ�Ŀ
    Set ������Ŀid = ����id_In
    Where �շ�ϸĿid = ҩƷid_In And (Sysdate Between ִ������ And ��ֹ���� Or Sysdate >= ִ������ And ��ֹ���� Is Null) And �䶯ԭ�� = 1;
  End If;

  --ʱ��->����
  If v_�Ƿ��� = 1 And �Ƿ���_In = 0 Then
    Update �շ���ĿĿ¼ Set �Ƿ��� = �Ƿ���_In Where ID = ҩƷid_In;
  
    Begin
      Select �ּ�, ID As �۸�id
      Into n_�շѼ�Ŀ�ּ�, n_�۸�id
      From �շѼ�Ŀ
      Where �շ�ϸĿid = ҩƷid_In And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And �䶯ԭ�� = 1;
    Exception
      When Others Then
        n_�շѼ�Ŀ�ּ� := Null;
        n_�۸�id       := Null;
    End;
  
    Begin
      Select �ϴ��ۼ� Into n_ҩƷ�ϴ��ۼ� From ҩƷ��� Where ҩƷid = ҩƷid_In;
    Exception
      When Others Then
        n_ҩƷ�ϴ��ۼ� := Null;
    End;
  
    If n_ҩƷ�ϴ��ۼ� Is Null Then
      n_ҩƷ�ϴ��ۼ� := n_�շѼ�Ŀ�ּ�;
    End If;
  
    Zl_�շѼ�Ŀ_Stop(ҩƷid_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into �շѼ�Ŀ
      (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
    Values
      (�շѼ�Ŀ_Id.Nextval, n_�۸�id, ҩƷid_In, n_�շѼ�Ŀ�ּ�, n_ҩƷ�ϴ��ۼ�, ����id_In, 1, 'ʱ��ת����', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    Select ���� Into n_��ͨ���С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 4 And ��λ = 5;
  
    --ȡ������ID
    Select ���id Into Classid From ҩƷ�������� Where ���� = 13;
  
    n_���   := 0;
    v_Billno := Null;
  
    For r_Priceadjust In c_Priceadjust Loop
      If n_ҩƷ�ϴ��ۼ� <> r_Priceadjust.���ۼ� Then
        If v_Billno Is Null Then
          Select Nextno(147) Into v_Billno From Dual;
        End If;
        n_��� := n_��� + 1;
        Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
        n_���۽�� := Round(n_ҩƷ�ϴ��ۼ� * r_Priceadjust.ʵ������, n_��ͨ���С��) -
                  Round(r_Priceadjust.���ۼ� * r_Priceadjust.ʵ������, n_��ͨ���С��);
        --��������Ӱ���¼
        Insert Into ҩƷ�շ���¼
          (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ���۽��, ���, ժҪ, ������,
           ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ����, Ƶ��, ��ҩ��λid)
        Values
          (n_�շ�id, 1, 13, v_Billno, n_���, Classid, r_Priceadjust.ҩƷid, r_Priceadjust.����, r_Priceadjust.����,
           r_Priceadjust.Ч��, r_Priceadjust.����, 1, r_Priceadjust.ʵ������, 0, r_Priceadjust.���ۼ�, 0, n_ҩƷ�ϴ��ۼ�,
           r_Priceadjust.����, n_���۽��, n_���۽��, 'ʱ��ת����', Zl_Username, Sysdate, r_Priceadjust.�ⷿid, 1, n_�۸�id, Zl_Username,
           Sysdate, r_Priceadjust.ʵ�ʽ��, r_Priceadjust.ʵ�ʲ��, r_Priceadjust.��Ӧ��id);
      
        Zl_ҩƷ���_Update(n_�շ�id, 2, 0);
      End If;
    End Loop;
  
    --����->ʱ��
  Elsif v_�Ƿ��� = 0 And �Ƿ���_In = 1 Then
    Update �շ���ĿĿ¼ Set �Ƿ��� = �Ƿ���_In Where ID = ҩƷid_In;
    Begin
      Select �ּ�, ID As �۸�id
      Into n_�շѼ�Ŀ�ּ�, n_�۸�id
      From �շѼ�Ŀ
      Where �շ�ϸĿid = ҩƷid_In And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And �䶯ԭ�� = 1;
    Exception
      When Others Then
        n_�շѼ�Ŀ�ּ� := Null;
        n_�۸�id       := Null;
    End;
  
    Zl_�շѼ�Ŀ_Stop(ҩƷid_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into �շѼ�Ŀ
      (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
    Values
      (�շѼ�Ŀ_Id.Nextval, n_�۸�id, ҩƷid_In, n_�շѼ�Ŀ�ּ�, n_�շѼ�Ŀ�ּ�, ����id_In, 1, '����תʱ��', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    For r_Priceadjust In c_Priceadjust Loop
      n_ҩƷ�۸��¼ := 0;
      Begin
        Select 1, �ּ�
        Into n_ҩƷ�۸��¼, n_ԭ��
        From ҩƷ�۸��¼
        Where ҩƷid = r_Priceadjust.ҩƷid And �ⷿid = r_Priceadjust.�ⷿid And Nvl(����, 0) = r_Priceadjust.���� And ��¼״̬ = 1 And
              �۸����� = 1;
      Exception
        When Others Then
          n_ҩƷ�۸��¼ := 0;
          n_ԭ��         := n_�շѼ�Ŀ�ּ�;
      End;
    
      If n_ҩƷ�۸��¼ = 1 Then
        Zl_ҩƷ�۸��¼_Stop(1, r_Priceadjust.�ⷿid, r_Priceadjust.ҩƷid, r_Priceadjust.����, Sysdate - 1 / 24 / 60 / 60, 2);
      End If;
      Zl_ҩƷ�۸��¼_Insert(0, 1, r_Priceadjust.�ⷿid, r_Priceadjust.ҩƷid, r_Priceadjust.����, n_ԭ��, n_�շѼ�Ŀ�ּ�, Sysdate, '����תʱ��',
                       Zl_Username, Null, r_Priceadjust.��Ӧ��id, r_Priceadjust.����, r_Priceadjust.Ч��, r_Priceadjust.����,
                       r_Priceadjust.���Ч��, Null, Null, Null, Null, 1);
    
      Update ҩƷ���
      Set ���ۼ� = n_�շѼ�Ŀ�ּ�
      Where ���� = 1 And �ⷿid = r_Priceadjust.�ⷿid And ҩƷid = r_Priceadjust.ҩƷid And Nvl(����, 0) = r_Priceadjust.����;
    
    End Loop;
  End If;

  --ҩƷ�����̱Ƚ����� 
  If ����_In Is Not Null Then
    Update ҩƷ������ Set ���� = ����_In Where ���� = ����_In;
    If Sql%RowCount = 0 Then
      Insert Into ҩƷ������
        (����, ����, ����)
        Select Nvl(Max(To_Number(����)), 0) + 1, ����_In, zlSpellCode(����_In, 10) From ҩƷ������;
    End If;
  End If;

  --ԭ���ؽ����� 
  If ԭ����_In Is Not Null Then
    Update ҩƷ������ Set ���� = ԭ����_In Where ���� = ԭ����_In;
    If Sql%RowCount = 0 Then
      Insert Into ҩƷ������
        (����, ����, ����)
        Select Nvl(Max(To_Number(����)), 0) + 1, ԭ����_In, zlSpellCode(ԭ����_In, 10) From ҩƷ������;
    End If;
  End If;

  --ҩƷ���ȵ���(����ģʽʱ)
  Zl_ҩƷ���ľ���_���۵���;

  b_Message.Zlhis_Dict_036(v_���, ҩƷid_In);
Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]�ù�񲻴��ڣ������ѱ������û�ɾ����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��ҩ���_Update;
/

--125588:����,2018-05-14,������ʵ������Ϊ0������
Create Or Replace Procedure Zl_��ҩ���_Update
(
  ҩƷid_In         In ҩƷ���.ҩƷid%Type,
  ����_In           In �շ���ĿĿ¼.����%Type,
  ���_In           In �շ���ĿĿ¼.���%Type,
  ����_In           In �շ���ĿĿ¼.����%Type := Null,
  Ʒ��_In           In �շ���Ŀ����.����%Type := Null,
  ƴ��_In           In �շ���Ŀ����.����%Type := Null,
  ���_In           In �շ���Ŀ����.����%Type := Null,
  ������_In         In �շ���Ŀ����.����%Type := Null,
  ��ʶ��_In         In ҩƷ���.��ʶ��%Type := Null,
  ҩƷ��Դ_In       In ҩƷ���.ҩƷ��Դ%Type := Null,
  ��׼�ĺ�_In       In ҩƷ���.��׼�ĺ�%Type := Null,
  ע���̱�_In       In ҩƷ���.ע���̱�%Type := Null,
  �ۼ۵�λ_In       In �շ���ĿĿ¼.���㵥λ%Type := Null,
  ����ϵ��_In       In ҩƷ���.����ϵ��%Type := Null,
  ���ﵥλ_In       In ҩƷ���.���ﵥλ%Type := Null,
  �����װ_In       In ҩƷ���.�����װ%Type := Null,
  סԺ��λ_In       In ҩƷ���.סԺ��λ%Type := Null,
  סԺ��װ_In       In ҩƷ���.סԺ��װ%Type := Null,
  ҩ�ⵥλ_In       In ҩƷ���.ҩ�ⵥλ%Type := Null,
  ҩ���װ_In       In ҩƷ���.ҩ���װ%Type := Null,
  ���쵥λ_In       In ҩƷ���.���쵥λ%Type := 1,
  ���췧ֵ_In       In ҩƷ���.���췧ֵ%Type := Null,
  �Ƿ���_In       In �շ���ĿĿ¼.�Ƿ���%Type := Null,
  ָ��������_In     In ҩƷ���.ָ��������%Type := Null,
  ����_In           In ҩƷ���.����%Type := 95,
  ָ�����ۼ�_In     In ҩƷ���.ָ�����ۼ�%Type := Null,
  �ӳ���_In         In ҩƷ���.�ӳ���%Type := Null,
  ����ѱ���_In     In ҩƷ���.����ѱ���%Type := Null,
  ҩ�ۼ���_In       In ҩƷ���.ҩ�ۼ���%Type := Null,
  ��������_In       In �շ���ĿĿ¼.��������%Type := Null,
  �������_In       In �շ���ĿĿ¼.�������%Type := Null,
  Gmp��֤_In        In ҩƷ���.Gmp��֤%Type := 0,
  �б�ҩƷ_In       In ҩƷ���.�б�ҩƷ%Type := 0,
  ���ηѱ�_In       In �շ���ĿĿ¼.���ηѱ�%Type := 0,
  סԺ�ɷ����_In   In ҩƷ���.סԺ�ɷ����%Type := 0,
  ҩ�����_In       In ҩƷ���.ҩ�����%Type := Null,
  ҩ������_In       In ҩƷ���.ҩ������%Type := Null,
  ���Ч��_In       In ҩƷ���.���Ч��%Type := Null,
  ���������_In     In ҩƷ���.���������%Type := 0,
  �ɱ���_In         In ҩƷ���.�ɱ���%Type := 0,
  ��ǰ�ۼ�_In       In �շѼ�Ŀ.�ּ�%Type := 0,
  ����id_In         In �շѼ�Ŀ.������Ŀid%Type := Null,
  ��ͬ��λid_In     In ҩƷ���.��ͬ��λid%Type := Null,
  ˵��_In           In �շ���ĿĿ¼.˵��%Type := Null,
  ��̬����_In       In ҩƷ���.��̬����%Type := 0,
  ��ҩ����_In       In ҩƷ���.��ҩ����%Type := Null,
  ��ѡ��_In         In �շ���ĿĿ¼.��ѡ��%Type := Null,
  ��ֵ˰��_In       In ҩƷ���.��ֵ˰��%Type := Null,
  ����ҩ��_In       In ҩƷ���.����ҩ��%Type := Null,
  վ��_In           In �շ���ĿĿ¼.վ��%Type := Null,
  �Ƿ񳣱�_In       In ҩƷ���.�Ƿ񳣱�%Type := Null,
  �洢�¶�_In       In ��ҺҩƷ����.�洢�¶�%Type := Null,
  �洢����_In       In ��ҺҩƷ����.�洢����%Type := Null,
  ��ҩ����_In       In ��ҺҩƷ����.��ҩ����%Type := Null,
  �Ƿ�������_In   In ��ҺҩƷ����.�Ƿ�������%Type := Null,
  ����_In           In ҩƷ���.����%Type := Null,
  ������Ŀ_In       In �շ���ĿĿ¼.������Ŀ%Type := Null,
  ����ɷ����_In   In ҩƷ���.����ɷ����%Type := 0,
  Dddֵ_In          In ҩƷ���.Dddֵ%Type := 0,
  ��ΣҩƷ_In       ҩƷ���.��ΣҩƷ%Type := Null,
  �ͻ���λ_In       In ҩƷ���.�ͻ���λ%Type := Null,
  �ͻ���װ_In       In ҩƷ���.�ͻ���װ%Type := Null,
  ��Һע������_In   In ��ҺҩƷ����.��Һע������%Type := Null,
  �Ƿ��ҩ_In       In ҩƷ���.�Ƿ��ҩ%Type := Null,
  �Ƿ����۹���_In In ҩƷ���.�Ƿ����۹���%Type := Null,
  ��λ��_In         In ҩƷ���.��λ��%Type := Null
) Is
  v_ҩ��id   ������ĿĿ¼.Id%Type;
  v_����     ������ĿĿ¼.����%Type;
  v_�Ƿ��� �շ���ĿĿ¼.�Ƿ���%Type; --������ҩƷ��ʱ��Ϊʱ�ۣ�ʱ��ҩƷֻ����δ������������޸�Ϊ���ۣ���������������޸Ķ������� 
  v_����     Number(2);
  Err_Notfind Exception;
  v_No           �շѼ�Ŀ.No%Type;
  v_Temp         �շ���ĿĿ¼.������Ŀ%Type;
  v_������Ŀ     �շ���ĿĿ¼.������Ŀ%Type;
  n_ָ�������   ҩƷ���.ָ�������%Type;
  n_ҩƷ�ϴ��ۼ� ҩƷ���.�ϴ��ۼ�%Type;
  n_���۽��     ҩƷ���.ʵ�ʽ��%Type;
  n_�շ�id       ҩƷ�շ���¼.Id%Type;
  n_��ͨ���С�� Number;
  n_���         Number(8);
  Classid        Number(18); --������
  v_Billno       ҩƷ�շ���¼.No%Type; --���۵���
  n_�۸�id       �շѼ�Ŀ.Id%Type;
  n_�շѼ�Ŀ�ּ� �շѼ�Ŀ.�ּ�%Type;
  n_ԭ��         ҩƷ�۸��¼.ԭ��%Type;
  n_ҩƷ�۸��¼ Number(1);
  v_���         �շ���ĿĿ¼.���%Type;
  --����->ʱ�ۺ����ҩƷ�۸��¼��ֵ

  Cursor c_Priceadjust Is
    Select s.ҩƷid, s.�ⷿid, Nvl(s.����, 0) As ����, s.�ϴι�Ӧ��id As ��Ӧ��id, s.�ϴ����� As ����, s.Ч��, s.�ϴβ��� As ����,
           Nvl(s.ʵ������, 0) As ʵ������, s.�ϴο��� As ����, Nvl(s.ʵ�ʽ��, 0) As ʵ�ʽ��, Nvl(s.ʵ�ʲ��, 0) As ʵ�ʲ��, Nvl(s.���ۼ�, 0) As ���ۼ�,
           s.ƽ���ɱ���, s.���Ч��, s.��׼�ĺ�, s.�ϴ��������� As ��������
    From ҩƷ��� S
    Where s.ҩƷid = ҩƷid_In And s.���� = 1 
    Order By s.ҩƷid, s.����, s.�ⷿid;

  r_Priceadjust c_Priceadjust%RowType;

Begin
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
  --ͨ������ 
  Select ID, ����
  Into v_ҩ��id, v_����
  From ������ĿĿ¼
  Where ID = (Select ҩ��id From ҩƷ��� Where ҩƷid = ҩƷid_In);
  --ȡԭʼ�Ķ������� 
  Select �Ƿ��� Into v_�Ƿ��� From �շ���ĿĿ¼ Where ID = ҩƷid_In;
  --�����Ϣ 
  Update �շ���ĿĿ¼
  Set ���� = ����_In, ���� = v_����, ��� = ���_In, ���� = ����_In, ���㵥λ = �ۼ۵�λ_In, �������� = ��������_In, ������� = �������_In, ���ηѱ� = ���ηѱ�_In,
      ������Ŀ = v_������Ŀ, ˵�� = ˵��_In, ��ѡ�� = ��ѡ��_In, վ�� = վ��_In
  Where ID = ҩƷid_In
  Returning ��� Into v_���;
  If Sql%RowCount = 0 Then
    Raise Err_Notfind;
  End If;
  n_ָ������� := (1 - 1 / (1 + �ӳ���_In / 100)) * 100;
  Update ҩƷ���
  Set ��ʶ�� = ��ʶ��_In, ҩƷ��Դ = ҩƷ��Դ_In, ��׼�ĺ� = ��׼�ĺ�_In, ע���̱� = ע���̱�_In, ����ϵ�� = ����ϵ��_In, ���ﵥλ = ���ﵥλ_In, �����װ = �����װ_In,
      סԺ��λ = סԺ��λ_In, סԺ��װ = סԺ��װ_In, ҩ�ⵥλ = ҩ�ⵥλ_In, ҩ���װ = ҩ���װ_In, ���쵥λ = ���쵥λ_In, ���췧ֵ = ���췧ֵ_In, ָ�������� = ָ��������_In,
      ���� = ����_In, ָ�����ۼ� = ָ�����ۼ�_In, ָ������� = n_ָ�������, ����ѱ��� = ����ѱ���_In, ҩ�ۼ��� = ҩ�ۼ���_In, סԺ�ɷ���� = סԺ�ɷ����_In,
      ҩ����� = ҩ�����_In, ҩ������ = ҩ������_In, ���Ч�� = ���Ч��_In, �б�ҩƷ = �б�ҩƷ_In, Gmp��֤ = Gmp��֤_In, ��������� = ���������_In,
      ��ͬ��λid = ��ͬ��λid_In, ��̬���� = ��̬����_In, ��ҩ���� = ��ҩ����_In, ��ֵ˰�� = ��ֵ˰��_In, ����ҩ�� = ����ҩ��_In, �Ƿ񳣱� = �Ƿ񳣱�_In, ���� = ����_In,
      ����ɷ���� = ����ɷ����_In, Dddֵ = Dddֵ_In, ��ΣҩƷ = ��ΣҩƷ_In, �ͻ���λ = �ͻ���λ_In, �ͻ���װ = �ͻ���װ_In, �ӳ��� = �ӳ���_In, �Ƿ��ҩ = �Ƿ��ҩ_In,
      �Ƿ����۹��� = �Ƿ����۹���_In, ��λ�� = ��λ��_In
  Where ҩƷid = ҩƷid_In;

  --�����޸ģ�����ҩƷ������ҩ���г�ҩ��ʱ��ȱʡ�������Ϊ�����סԺ������޸Ĺ��ҩƷʱ�����ٸ��ݹ��ҩƷ�ķ���������ҩƷ�ķ������ 
  --������Ŀ�������ĸ��� 
  --select nvl(sum(distinct I.�������),0) into v_���� 
  --from �շ���ĿĿ¼ I,ҩƷ��� S 
  --where I.ID=S.ҩƷID and S.ҩ��ID=v_ҩ��ID; 
  --update ������ĿĿ¼ 
  --set �������=decode(v_����,0,0,1,1,2,2,3) 
  --where ID=v_ҩ��ID; 

  --�����Ĵ��� 
  If ������_In Is Null Then
    Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 1 And ���� = 3;
  Else
    Update �շ���Ŀ���� Set ���� = v_����, ���� = ������_In Where �շ�ϸĿid = ҩƷid_In And ���� = 1 And ���� = 3;
    If Sql%RowCount = 0 Then
      Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, v_����, 1, ������_In, 3);
    End If;
  End If;
  If Ʒ��_In Is Null Then
    Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 3;
  Else
    If ƴ��_In Is Null Then
      Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 1;
    Else
      Update �շ���Ŀ���� Set ���� = Ʒ��_In, ���� = ƴ��_In Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 1;
      If Sql%RowCount = 0 Then
        Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, Ʒ��_In, 3, ƴ��_In, 1);
      End If;
    End If;
    If ���_In Is Null Then
      Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 2;
    Else
      Update �շ���Ŀ���� Set ���� = Ʒ��_In, ���� = ���_In Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 2;
      If Sql%RowCount = 0 Then
        Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, Ʒ��_In, 3, ���_In, 2);
      End If;
    End If;
  End If;

  --������Ϣ������Ѿ��з�����������ֱ�Ӹ�����Щ��Ϣ 
  Select Nvl(Count(*), 0) Into v_���� From ҩƷ�շ���¼ Where ҩƷid = ҩƷid_In And Rownum < 2;
  If v_���� = 0 Then
    Update ҩƷ��� Set �ɱ��� = �ɱ���_In Where ҩƷid = ҩƷid_In;
    If ����id_In Is Not Null Then
      Update �շѼ�Ŀ
      Set �ּ� = ��ǰ�ۼ�_In, ������Ŀid = ����id_In, �䶯ԭ�� = 1, ����˵�� = '�޸Ķ���', ������ = User
      Where �շ�ϸĿid = ҩƷid_In And (Sysdate Between ִ������ And ��ֹ���� Or Sysdate >= ִ������ And ��ֹ���� Is Null) And �䶯ԭ�� = 1;
      If Sql%RowCount = 0 Then
        v_No := Nextno(9);
        Insert Into �շѼ�Ŀ
          (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
        Values
          (�շѼ�Ŀ_Id.Nextval, Null, ҩƷid_In, 0, ��ǰ�ۼ�_In, ����id_In, 1, '��������', User, Sysdate,
           To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
      End If;
    End If;
  Else
    --����ҵ������ʱ�������޸ļ۸��ǿ����޸�������Ŀ 
    Update �շѼ�Ŀ
    Set ������Ŀid = ����id_In
    Where �շ�ϸĿid = ҩƷid_In And (Sysdate Between ִ������ And ��ֹ���� Or Sysdate >= ִ������ And ��ֹ���� Is Null) And �䶯ԭ�� = 1;
  End If;

  --ʱ��->����
  If v_�Ƿ��� = 1 And �Ƿ���_In = 0 Then
    Update �շ���ĿĿ¼ Set �Ƿ��� = �Ƿ���_In Where ID = ҩƷid_In;
  
    Begin
      Select �ּ�, ID As �۸�id
      Into n_�շѼ�Ŀ�ּ�, n_�۸�id
      From �շѼ�Ŀ
      Where �շ�ϸĿid = ҩƷid_In And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And �䶯ԭ�� = 1;
    Exception
      When Others Then
        n_�շѼ�Ŀ�ּ� := Null;
        n_�۸�id       := Null;
    End;
  
    Begin
      Select �ϴ��ۼ� Into n_ҩƷ�ϴ��ۼ� From ҩƷ��� Where ҩƷid = ҩƷid_In;
    Exception
      When Others Then
        n_ҩƷ�ϴ��ۼ� := Null;
    End;
  
    If n_ҩƷ�ϴ��ۼ� Is Null Then
      n_ҩƷ�ϴ��ۼ� := n_�շѼ�Ŀ�ּ�;
    End If;
  
    Zl_�շѼ�Ŀ_Stop(ҩƷid_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into �շѼ�Ŀ
      (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
    Values
      (�շѼ�Ŀ_Id.Nextval, n_�۸�id, ҩƷid_In, n_�շѼ�Ŀ�ּ�, n_ҩƷ�ϴ��ۼ�, ����id_In, 1, 'ʱ��ת����', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    Select ���� Into n_��ͨ���С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 4 And ��λ = 5;
  
    --ȡ������ID
    Select ���id Into Classid From ҩƷ�������� Where ���� = 13;
  
    n_���   := 0;
    v_Billno := Null;
  
    For r_Priceadjust In c_Priceadjust Loop
      If n_ҩƷ�ϴ��ۼ� <> r_Priceadjust.���ۼ� Then
        If v_Billno Is Null Then
          Select Nextno(147) Into v_Billno From Dual;
        End If;
        n_��� := n_��� + 1;
        Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
        n_���۽�� := Round(n_ҩƷ�ϴ��ۼ� * r_Priceadjust.ʵ������, n_��ͨ���С��) -
                  Round(r_Priceadjust.���ۼ� * r_Priceadjust.ʵ������, n_��ͨ���С��);
        --��������Ӱ���¼
        Insert Into ҩƷ�շ���¼
          (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ���۽��, ���, ժҪ, ������,
           ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ����, Ƶ��, ��ҩ��λid)
        Values
          (n_�շ�id, 1, 13, v_Billno, n_���, Classid, r_Priceadjust.ҩƷid, r_Priceadjust.����, r_Priceadjust.����,
           r_Priceadjust.Ч��, r_Priceadjust.����, 1, r_Priceadjust.ʵ������, 0, r_Priceadjust.���ۼ�, 0, n_ҩƷ�ϴ��ۼ�,
           r_Priceadjust.����, n_���۽��, n_���۽��, 'ʱ��ת����', Zl_Username, Sysdate, r_Priceadjust.�ⷿid, 1, n_�۸�id, Zl_Username,
           Sysdate, r_Priceadjust.ʵ�ʽ��, r_Priceadjust.ʵ�ʲ��, r_Priceadjust.��Ӧ��id);
      
        Zl_ҩƷ���_Update(n_�շ�id, 2, 0);
      End If;
    End Loop;
  
    --����->ʱ��
  Elsif v_�Ƿ��� = 0 And �Ƿ���_In = 1 Then
    Update �շ���ĿĿ¼ Set �Ƿ��� = �Ƿ���_In Where ID = ҩƷid_In;
    Begin
      Select �ּ�, ID As �۸�id
      Into n_�շѼ�Ŀ�ּ�, n_�۸�id
      From �շѼ�Ŀ
      Where �շ�ϸĿid = ҩƷid_In And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And �䶯ԭ�� = 1;
    Exception
      When Others Then
        n_�շѼ�Ŀ�ּ� := Null;
        n_�۸�id       := Null;
    End;
  
    Zl_�շѼ�Ŀ_Stop(ҩƷid_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into �շѼ�Ŀ
      (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
    Values
      (�շѼ�Ŀ_Id.Nextval, n_�۸�id, ҩƷid_In, n_�շѼ�Ŀ�ּ�, n_�շѼ�Ŀ�ּ�, ����id_In, 1, '����תʱ��', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    For r_Priceadjust In c_Priceadjust Loop
      n_ҩƷ�۸��¼ := 0;
      Begin
        Select 1, �ּ�
        Into n_ҩƷ�۸��¼, n_ԭ��
        From ҩƷ�۸��¼
        Where ҩƷid = r_Priceadjust.ҩƷid And �ⷿid = r_Priceadjust.�ⷿid And Nvl(����, 0) = r_Priceadjust.���� And ��¼״̬ = 1 And
              �۸����� = 1;
      Exception
        When Others Then
          n_ҩƷ�۸��¼ := 0;
          n_ԭ��         := n_�շѼ�Ŀ�ּ�;
      End;
    
      If n_ҩƷ�۸��¼ = 1 Then
        Zl_ҩƷ�۸��¼_Stop(1, r_Priceadjust.�ⷿid, r_Priceadjust.ҩƷid, r_Priceadjust.����, Sysdate - 1 / 24 / 60 / 60, 2);
      End If;
      Zl_ҩƷ�۸��¼_Insert(0, 1, r_Priceadjust.�ⷿid, r_Priceadjust.ҩƷid, r_Priceadjust.����, n_ԭ��, n_�շѼ�Ŀ�ּ�, Sysdate, '����תʱ��',
                       Zl_Username, Null, r_Priceadjust.��Ӧ��id, r_Priceadjust.����, r_Priceadjust.Ч��, r_Priceadjust.����,
                       r_Priceadjust.���Ч��, Null, Null, Null, Null, 1);
    
      Update ҩƷ���
      Set ���ۼ� = n_�շѼ�Ŀ�ּ�
      Where ���� = 1 And �ⷿid = r_Priceadjust.�ⷿid And ҩƷid = r_Priceadjust.ҩƷid And Nvl(����, 0) = r_Priceadjust.����;
    
    End Loop;
  End If;

  --ҩƷ�����̱Ƚ����� 
  If ����_In Is Not Null Then
    Update ҩƷ������ Set ���� = ����_In Where ���� = ����_In;
    If Sql%RowCount = 0 Then
      Insert Into ҩƷ������
        (����, ����, ����)
        Select Nvl(Max(To_Number(����)), 0) + 1, ����_In, zlSpellCode(����_In, 10) From ҩƷ������;
    End If;
  End If;

  --�޸���ҺҩƷ���� 
  Update ��ҺҩƷ����
  Set �洢�¶� = �洢�¶�_In, �洢���� = �洢����_In, ��ҩ���� = ��ҩ����_In, �Ƿ������� = �Ƿ�������_In, ��Һע������ = ��Һע������_In
  Where ҩƷid = ҩƷid_In;

  If Sql%NotFound Then
    Insert Into ��ҺҩƷ����
      (ҩƷid, �洢�¶�, �洢����, ��ҩ����, �Ƿ�������, ��Һע������)
    Values
      (ҩƷid_In, �洢�¶�_In, �洢����_In, ��ҩ����_In, �Ƿ�������_In, ��Һע������_In);
  End If;

  --ҩƷ���ȵ���(����ģʽʱ)
  Zl_ҩƷ���ľ���_���۵���;

  b_Message.Zlhis_Dict_036(v_���, ҩƷid_In);
Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]�ù�񲻴��ڣ������ѱ������û�ɾ����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��ҩ���_Update;
/

--125588:����,2018-05-14,������ʵ������Ϊ0������
Create Or Replace Procedure Zl_ҩƷ�շ���¼_Adjust
(
  ҩƷid_In   In Number, --ҩƷID,Ϊ0ʱ�������Ԥ����
  ���۷�ʽ_In In Number := 0 --0-����ۼۺͳɱ���Ԥ����,1-ֻ����ۼ�Ԥ����,2-ֻ���ɱ���Ԥ����
) As
  Classid          Number(18); --������
  v_Billno         ҩƷ�շ���¼.No%Type; --���۵���
  Adjustdate       Date; --����ʱ��
  n_����           Number(18);
  n_�ּ�           �շѼ�Ŀ.�ּ�%Type;
  n_ԭ��           �շѼ�Ŀ.ԭ��%Type;
  n_���           Number(8);
  n_ԭ��id         �շѼ�Ŀ.ԭ��id%Type;
  n_���۽��       ҩƷ���.ʵ�ʽ��%Type;
  n_�շ�id         ҩƷ�շ���¼.Id%Type;
  n_��ͨ���С��   Number;
  n_Stockid        ҩƷ�շ���¼.�ⷿid%Type;
  n_������id     ҩƷ�շ���¼.������id%Type;
  n_���ϵ��       ҩƷ�շ���¼.���ϵ��%Type;
  n_�۸�id         �շѼ�Ŀ.Id%Type;
  n_�޿�����ģʽ Number(1) := 0;
  n_��������       Number(1) := 0;
  n_��Ϣ����       Number(1) := 0;
  --�����ۼۣ�ʱ���ۼ�Ԥ���ۼ�¼
  --�۸����ͣ�0-�����ۼ�,1-ʱ���ۼ�
  Cursor c_Priceadjust Is
    Select 0 As �۸�����, p.Id As �۸�id, p.ԭ��id, p.ִ������, p.ԭ��, p.�ּ�, i.ҩƷid, s.�ⷿid As �ⷿid, Nvl(s.����, 0) As ����, s.�ϴ����� As ����,
           s.Ч��, s.�ϴβ��� As ����, s.�ϴι�Ӧ��id As ��Ӧ��id, Nvl(s.ʵ������, 0) As ʵ������, s.�ϴο��� As ����, Nvl(s.ʵ�ʽ��, 0) As ʵ�ʽ��,
           Nvl(s.ʵ�ʲ��, 0) As ʵ�ʲ��, Nvl(s.���ۼ�, 0) As ���ۼ�, s.ƽ���ɱ���, s.Rowid As ����¼, s.���Ч��, s.��׼�ĺ�, s.�ϴ��������� As ��������,
           p.���ۻ��ܺ�
    From �շѼ�Ŀ P, ҩƷ��� I, �շ���ĿĿ¼ A, ҩƷ��� S
    Where i.ҩƷid = a.Id And p.�շ�ϸĿid = i.ҩƷid And i.ҩƷid = s.ҩƷid(+) And s.����(+) = 1 And Nvl(a.�Ƿ���, 0) = 0 And
          Sysdate Between p.ִ������ And Nvl(p.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(p.�䶯ԭ��, 0) = 0 And
          p.�շ�ϸĿid = Decode(ҩƷid_In, 0, p.�շ�ϸĿid, ҩƷid_In) And ���۷�ʽ_In In (0, 1)
    Union All
    Select �۸�����, p.Id As �۸�id, p.ԭ��id, p.ִ������, p.ԭ��, p.�ּ�, i.ҩƷid, p.�ⷿid As �ⷿid, Nvl(p.����, 0) As ����, p.���� As ����, p.Ч��,
           p.���� As ����, s.�ϴι�Ӧ��id As ��Ӧ��id, Nvl(s.ʵ������, 0) As ʵ������, s.�ϴο��� As ����, Nvl(s.ʵ�ʽ��, 0) As ʵ�ʽ��,
           Nvl(s.ʵ�ʲ��, 0) As ʵ�ʲ��, Nvl(s.���ۼ�, 0) As ���ۼ�, s.ƽ���ɱ���, s.Rowid As ����¼, s.���Ч��, s.��׼�ĺ�, s.�ϴ��������� As ��������,
           p.���ۻ��ܺ�
    From ҩƷ�۸��¼ P, ҩƷ��� I, �շ���ĿĿ¼ A, ҩƷ��� S
    Where i.ҩƷid = a.Id And p.ҩƷid = i.ҩƷid And p.�ⷿid = s.�ⷿid(+) And p.ҩƷid = s.ҩƷid(+) And
          Nvl(p.����, 0) = Nvl(s.����(+), 0) And s.����(+) = 1 And Sysdate Between p.ִ������ And
          Nvl(p.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(p.��¼״̬, 0) = 0 And
          p.ҩƷid = Decode(ҩƷid_In, 0, p.ҩƷid, ҩƷid_In) And �۸����� = 1 And ���۷�ʽ_In In (0, 1)
    Order By �۸�����, ҩƷid, ����, �ⷿid;

  r_Priceadjust c_Priceadjust%RowType;

  --�ɱ���Ԥ���ۼ�¼
  Cursor c_Costadjust Is
    Select �۸�����, p.Id As �۸�id, p.ԭ��id, p.ִ������, p.ԭ��, p.�ּ�, i.ҩƷid, p.�ⷿid As �ⷿid, Nvl(p.����, 0) As ����, p.���� As ����, p.Ч��,
           p.���� As ����, s.�ϴι�Ӧ��id As ��Ӧ��id, Nvl(s.ʵ������, 0) As ʵ������, s.�ϴο��� As ����, Nvl(s.ʵ�ʽ��, 0) As ʵ�ʽ��,
           Nvl(s.ʵ�ʲ��, 0) As ʵ�ʲ��, Nvl(s.���ۼ�, 0) As ���ۼ�, s.ƽ���ɱ���, s.Rowid As ����¼, s.���Ч��, s.��׼�ĺ�, s.�ϴ��������� As ��������,
           p.���ۻ��ܺ�
    From ҩƷ�۸��¼ P, ҩƷ��� I, �շ���ĿĿ¼ A, ҩƷ��� S
    Where i.ҩƷid = a.Id And p.ҩƷid = i.ҩƷid And p.�ⷿid = s.�ⷿid(+) And p.ҩƷid = s.ҩƷid(+) And
          Nvl(p.����, 0) = Nvl(s.����(+), 0) And s.����(+) = 1 And Sysdate Between p.ִ������ And
          Nvl(p.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(p.��¼״̬, 0) = 0 And
          p.ҩƷid = Decode(ҩƷid_In, 0, p.ҩƷid, ҩƷid_In) And �۸����� = 2 And ���۷�ʽ_In In (0, 2)
    Order By ҩƷid, ����, �ⷿid;

  r_Costadjust c_Costadjust%RowType;

  --��ǰ��Ч�ļ۸������޿�����
  Cursor c_Nostockadjust
  (
    Drugid_In ҩƷ�۸��¼.ҩƷid%Type,
    Type_In   ҩƷ�۸��¼.�۸�����%Type
  ) Is
    Select a.�۸�����, a.Id As �۸�id, a.ԭ��, a.�ּ�, a.ҩƷid, a.�ⷿid, a.����, a.��ҩ��λid, a.����, a.Ч��, a.����
    From ҩƷ�۸��¼ A
    Where Sysdate Between a.ִ������ And Nvl(a.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And a.��¼״̬ = 1 And
          a.ҩƷid = Drugid_In And a.�۸����� = Type_In And a.�ⷿid Is Not Null
    Order By a.�ⷿid, a.ҩƷid, a.����;

  r_Nostockadjust c_Nostockadjust%RowType;
Begin
  --ȡ��ͨҵ�񾫶�λ��
  --���:1-ҩƷ 2-����
  --���ݣ�2-���ۼ� 4-���
  --��λ��ҩƷ:1-�ۼ� 5-��λ
  Select ���� Into n_��ͨ���С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 4 And ��λ = 5;

  --ȡ������ID
  Select ���id Into Classid From ҩƷ�������� Where ���� = 13;

  Adjustdate := Sysdate;

  --�ۼ۵��۴���
  If ���۷�ʽ_In = 0 Or ���۷�ʽ_In = 1 Then
  
    n_��� := 0;
  
    --ȡ����NOȡ
    Select Nextno(147) Into v_Billno From Dual;
  
    For r_Priceadjust In c_Priceadjust Loop
      If r_Priceadjust.�ⷿid Is Not Null Then
        --�пⷿid��������
      
        --ȡ��������
        n_�������� := Zl_Fun_Getbatchpro(r_Priceadjust.�ⷿid, r_Priceadjust.ҩƷid);
      
        --��������ӯ����¼��������1.Ҫ�п���¼��2.�������ԺͿ������һ��
        If r_Priceadjust.����¼ Is Not Null And ((n_�������� = 1 And r_Priceadjust.���� > 0) Or
           (n_�������� = 0 And r_Priceadjust.���� = 0)) Then
          n_��� := n_��� + 1;
        
          Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
        
          n_ԭ�� := r_Priceadjust.ԭ��;
        
          --ʱ�۵��ۣ����ԭ�ۺ͵�ǰ��治һ�£����Ե�ǰ���Ϊ׼
          If r_Priceadjust.�۸����� = 1 And r_Priceadjust.ԭ�� <> r_Priceadjust.���ۼ� And r_Priceadjust.����¼ Is Not Null Then
            n_ԭ�� := r_Priceadjust.���ۼ�;
          End If;
        
          n_���۽�� := Round(r_Priceadjust.�ּ� * r_Priceadjust.ʵ������, n_��ͨ���С��) - Round(n_ԭ�� * r_Priceadjust.ʵ������, n_��ͨ���С��);
        
          n_�۸�id := r_Priceadjust.�۸�id;
          If r_Priceadjust.�۸����� = 1 Then
            Select ID
            Into n_�۸�id
            From �շѼ�Ŀ
            Where Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And �շ�ϸĿid = r_Priceadjust.ҩƷid;
          End If;
        
          --��������Ӱ���¼
          Insert Into ҩƷ�շ���¼
            (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ���۽��, ���, ժҪ, ������,
             ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ����, Ƶ��, ��ҩ��λid)
          Values
            (n_�շ�id, 1, 13, v_Billno, n_���, Classid, r_Priceadjust.ҩƷid, r_Priceadjust.����, r_Priceadjust.����,
             r_Priceadjust.Ч��, r_Priceadjust.����, 1, r_Priceadjust.ʵ������, 0, n_ԭ��, 0, r_Priceadjust.�ּ�, r_Priceadjust.����,
             n_���۽��, n_���۽��, 'ҩƷ����', Zl_Username, Adjustdate, r_Priceadjust.�ⷿid, 1, n_�۸�id, Zl_Username, Adjustdate,
             r_Priceadjust.ʵ�ʽ��, r_Priceadjust.ʵ�ʲ��, r_Priceadjust.��Ӧ��id);
        
          Zl_δ��ҩƷ��¼_Insert(n_�շ�id);
        
          --����ҩƷ��棬�޿�治ִ��
          If r_Priceadjust.����¼ Is Not Null Then
            Zl_ҩƷ���_Update(n_�շ�id, 2, 0);
          End If;
        End If;
      
        --����ԭ�۸���Ϣ
        If r_Priceadjust.�۸����� = 1 Then
          Update ҩƷ�۸��¼ Set ��¼״̬ = 2 Where ID = r_Priceadjust.ԭ��id;
        End If;
      
        --ʱ�۵��۸��¼۸���е���Ϣ
        If r_Priceadjust.�۸����� = 1 Then
          --���µ�ǰ�۸���Ϣ
          If r_Priceadjust.����¼ Is Not Null Then
            Update ҩƷ�۸��¼
            Set ���� = r_Priceadjust.����, Ч�� = r_Priceadjust.Ч��, ���� = r_Priceadjust.����, ���Ч�� = r_Priceadjust.���Ч��,
                ��ҩ��λid = r_Priceadjust.��Ӧ��id, ԭ�� = n_ԭ��, �շ�id = n_�շ�id, ��¼״̬ = 1
            Where ID = r_Priceadjust.�۸�id;
          Else
            --�޿��ʱֻ���¼�¼״̬���շ�id
            Update ҩƷ�۸��¼ Set �շ�id = n_�շ�id, ��¼״̬ = 1 Where ID = r_Priceadjust.�۸�id;
          End If;
        End If;
      
        --�������Ŷ��ձ��ۼ�
        If r_Priceadjust.�۸����� = 1 Then
          --�����ʱ�ۣ�����¸�ҩƷ���ζ�Ӧ�ļ۸�
          Update ҩƷ���Ŷ���
          Set �ۼ� = r_Priceadjust.�ּ�
          Where ҩƷid = r_Priceadjust.ҩƷid And Nvl(����, 0) = r_Priceadjust.���� And �ۼ� <> r_Priceadjust.�ּ�;
        End If;
      
        --��Ϣ����
        --����ֻ����һ����Ϣ��ʱ�ۿɶ�ε���
        If (r_Priceadjust.�۸����� = 0 And n_��Ϣ���� = 0) Or r_Priceadjust.�۸����� = 1 Then
          n_��Ϣ���� := 1;
          b_Message.Zlhis_Drug_009(r_Priceadjust.�۸�id, r_Priceadjust.�۸�����);
        End If;
      Else
        --�޿�����ģʽ���۸���и�ҩƷ������Ч�ļ۸�Ҫ���޿�����ʱ�ļ۸����
      
        If r_Priceadjust.�۸����� = 1 Then
          --����ԭ�۸���Ϣ
          Update ҩƷ�۸��¼ Set ��¼״̬ = 2 Where ID = r_Priceadjust.ԭ��id;
        
          --�����ּ۸�״̬
          Update ҩƷ�۸��¼ Set ��¼״̬ = 1 Where ID = r_Priceadjust.�۸�id;
        End If;
      
        n_�޿�����ģʽ := 0;
        For r_Nostockadjust In c_Nostockadjust(r_Priceadjust.ҩƷid, 1) Loop
          If r_Priceadjust.�ּ� <> r_Nostockadjust.�ּ� Then
            Zl_ҩƷ�۸��¼_Stop(1, r_Nostockadjust.�ⷿid, r_Nostockadjust.ҩƷid, r_Nostockadjust.����,
                           Adjustdate - 2 / 24 / 60 / 60);
            Zl_ҩƷ�۸��¼_Insert(1, 1, r_Nostockadjust.�ⷿid, r_Nostockadjust.ҩƷid, r_Nostockadjust.����, Null,
                             r_Priceadjust.�ּ�, Adjustdate - 1 / 24 / 60 / 60, 'ҩƷ����', Zl_Username, r_Priceadjust.���ۻ��ܺ�,
                             r_Nostockadjust.��ҩ��λid, r_Nostockadjust.����, r_Nostockadjust.Ч��, r_Nostockadjust.����);
            n_�޿�����ģʽ := 1;
          End If;
        End Loop;
        If n_�޿�����ģʽ = 1 Then
          Zl_ҩƷ�շ���¼_Adjust(r_Priceadjust.ҩƷid, 1);
        End If;
      End If;
    
      --���¹��۸�
      If r_Priceadjust.�ּ� <> r_Priceadjust.ԭ�� Then
        Update ҩƷ���
        Set �ϴ��ۼ� = r_Priceadjust.�ּ�
        Where ҩƷid = r_Priceadjust.ҩƷid And �ϴ��ۼ� <> r_Priceadjust.�ּ�;
      End If;
    
      If r_Priceadjust.�۸����� = 0 Then
        n_�۸�id := r_Priceadjust.�۸�id;
      Else
        Begin
          Select ID
          Into n_�۸�id
          From �շѼ�Ŀ
          Where Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(�䶯ԭ��, 0) = 0 And
                �շ�ϸĿid = r_Priceadjust.ҩƷid;
        Exception
          When Others Then
            n_�۸�id := 0;
        End;
      End If;
    
      If n_�۸�id > 0 Then
        Update �շѼ�Ŀ Set �䶯ԭ�� = 1 Where Nvl(�䶯ԭ��, 0) = 0 And ID = n_�۸�id;
      End If;
    
      --�������Ŷ��ձ��ۼ�
      If r_Priceadjust.�۸����� = 0 Then
        --����Ƕ��ۣ�����¸�ҩƷ��Ӧ���������ε��ۼ�
        Update ҩƷ���Ŷ���
        Set �ۼ� = r_Priceadjust.�ּ�
        Where ҩƷid = r_Priceadjust.ҩƷid And �ۼ� <> r_Priceadjust.�ּ�;
      End If;
    End Loop;
  End If;

  --�ɱ��۵��۴���
  If ���۷�ʽ_In = 0 Or ���۷�ʽ_In = 2 Then
  
    n_���    := 0;
    n_Stockid := 0;
  
    Select b.Id, b.ϵ��
    Into n_������id, n_���ϵ��
    From ҩƷ�������� A, ҩƷ������ B
    Where a.���id = b.Id And a.���� = 5 And Rownum < 2;
  
    v_Billno := Nextno(25, n_Stockid);
  
    For r_Costadjust In c_Costadjust Loop
      If r_Costadjust.�ⷿid Is Not Null Then
        --�пⷿid��������
      
        --ȡ��������
        n_�������� := Zl_Fun_Getbatchpro(r_Costadjust.�ⷿid, r_Costadjust.ҩƷid);
      
        --��������ӯ����¼��������1.Ҫ�п���¼��2.�������ԺͿ������һ��
        If r_Costadjust.����¼ Is Not Null And ((n_�������� = 1 And r_Costadjust.���� > 0) Or
           (n_�������� = 0 And r_Costadjust.���� = 0)) Then
          n_��� := n_��� + 1;
        
          --��������۵�����
          Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
        
          --���ԭ�ۺ͵�ǰ��治һ�£����Ե�ǰ���Ϊ׼
          n_ԭ�� := r_Costadjust.ԭ��;
          If r_Costadjust.ԭ�� <> r_Costadjust.ƽ���ɱ��� And r_Costadjust.����¼ Is Not Null Then
            n_ԭ�� := r_Costadjust.ƽ���ɱ���;
          End If;
        
          n_���۽�� := Round(n_ԭ�� * r_Costadjust.ʵ������, n_��ͨ���С��) - Round(r_Costadjust.�ּ� * r_Costadjust.ʵ������, n_��ͨ���С��);
        
          Insert Into ҩƷ�շ���¼
            (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ��ҩ��λid, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ʵ������, ���ۼ�, ���۽��, �ɱ���, �ɱ����,
             ���, ժҪ, ������, ��������, �����, �������, ��������, ��׼�ĺ�, ����, ��ҩ��ʽ, ����, ���Ч��)
          Values
            (n_�շ�id, 1, 5, v_Billno, n_���, r_Costadjust.�ⷿid, n_������id, r_Costadjust.��Ӧ��id, n_���ϵ��, r_Costadjust.ҩƷid,
             r_Costadjust.����, r_Costadjust.����, r_Costadjust.����, r_Costadjust.Ч��, r_Costadjust.ʵ������, 0, r_Costadjust.ʵ�ʽ��,
             0, r_Costadjust.ʵ�ʲ��, 0, n_���۽��, '�ɱ��۵���', Zl_Username, Adjustdate, Zl_Username, Adjustdate,
             r_Costadjust.��������, r_Costadjust.��׼�ĺ�, r_Costadjust.�ּ�, 1, n_ԭ��, r_Costadjust.���Ч��);
        
          Zl_δ��ҩƷ��¼_Insert(n_�շ�id);
        
          Zl_ҩƷ���_Update(n_�շ�id, 2, 0);
        
          --���µ�ǰ�۸���Ϣ
          Update ҩƷ�۸��¼
          Set ���� = r_Costadjust.����, Ч�� = r_Costadjust.Ч��, ���� = r_Costadjust.����, ���Ч�� = r_Costadjust.���Ч��,
              ��ҩ��λid = r_Costadjust.��Ӧ��id, ԭ�� = n_ԭ��, �շ�id = n_�շ�id, ��¼״̬ = 1
          Where ID = r_Costadjust.�۸�id;
        Else
          --�޿��ʱֻ���¼�¼״̬
          Update ҩƷ�۸��¼ Set ��¼״̬ = 1 Where ID = r_Costadjust.�۸�id;
        End If;
      
        --����ԭ�۸���Ϣ
        Update ҩƷ�۸��¼ Set ��¼״̬ = 2 Where ID = r_Costadjust.ԭ��id;
      
        --�������Ŷ��ձ�ɱ���
        Update ҩƷ���Ŷ���
        Set �ɱ��� = r_Costadjust.�ּ�
        Where ҩƷid = r_Costadjust.ҩƷid And Nvl(����, 0) = r_Costadjust.���� And �ɱ��� <> r_Costadjust.�ּ�;
      Else
        --�޿�����ģʽ���۸���и�ҩƷ������Ч�ļ۸�Ҫ���޿�����ʱ�ļ۸����
      
        --����ԭ�۸���Ϣ
        Update ҩƷ�۸��¼ Set ��¼״̬ = 2 Where ID = r_Costadjust.ԭ��id;
      
        --�����ּ۸�״̬
        Update ҩƷ�۸��¼ Set ��¼״̬ = 1 Where ID = r_Costadjust.�۸�id;
      
        n_�޿�����ģʽ := 0;
        For r_Nostockadjust In c_Nostockadjust(r_Costadjust.ҩƷid, 2) Loop
          If r_Costadjust.�ּ� <> r_Nostockadjust.�ּ� Then
            Zl_ҩƷ�۸��¼_Stop(2, r_Nostockadjust.�ⷿid, r_Nostockadjust.ҩƷid, r_Nostockadjust.����,
                           Adjustdate - 2 / 24 / 60 / 60);
            Zl_ҩƷ�۸��¼_Insert(1, 2, r_Nostockadjust.�ⷿid, r_Nostockadjust.ҩƷid, r_Nostockadjust.����, Null,
                             r_Costadjust.�ּ�, Adjustdate - 1 / 24 / 60 / 60, '�ɱ��۵���', Zl_Username, r_Costadjust.���ۻ��ܺ�,
                             r_Nostockadjust.��ҩ��λid, r_Nostockadjust.����, r_Nostockadjust.Ч��, r_Nostockadjust.����);
            n_�޿�����ģʽ := 1;
          End If;
        End Loop;
        If n_�޿�����ģʽ = 1 Then
          Zl_ҩƷ�շ���¼_Adjust(r_Costadjust.ҩƷid, 2);
        End If;
      End If;
    
      --���¹��۸�
      If r_Costadjust.ԭ�� <> r_Costadjust.�ּ� Then
        Update ҩƷ���
        Set �ɱ��� = r_Costadjust.�ּ�
        Where ҩƷid = r_Costadjust.ҩƷid And �ɱ��� <> r_Costadjust.�ּ�;
      End If;
    
      --��Ϣ����
      b_Message.Zlhis_Drug_007(r_Costadjust.�۸�id);
    End Loop;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ���¼_Adjust;
/


--125656:��ҵ��,2018-05-14,�����˷�δɾ��δ��ҩƷ��¼
Create Or Replace Procedure Zl_ҩƷ�շ���¼_�����˷�
(
  ����id_In   In ������ü�¼.Id%Type,
  ��������_In In ҩƷ�շ���¼.ʵ������%Type := 0,
  ��ҩid_In   Varchar2 := Null,
  ��Ϣ_In     Number := 0 --�Ƿ���������Ϣ
) Is
  ----------------------------------
  --���ܣ�ɾ�������շѵ������ﻮ�۵��������շ�����ʱ��������ҩƷ��桢ҩƷ�շ���¼��δ��ҩ��¼�Ĺ���
  --������
  --      ����id_In��������ü�¼����סԺ���ü�¼��ɾ����������ʱ��ɾ�����ݵ�id
  --      ��������_In���������ʱ��Ҫ���ʵ�����
  --      ��ҩid_In���������ʱ��Һ����������Ҫ���ݵļ�¼id�����ַ������ݣ��ö��ŷָ�磺1001,1002,1003
  --      Ϊ�ձ�ʾ�������пɳ�����
  -----------------------------------
  --���α����ڴ���ҩƷ����������
  l_ҩƷ�շ�   t_Numlist := t_Numlist();
  n_����       ҩƷ�շ���¼.����%Type;
  v_No         ҩƷ�շ���¼.No%Type;
  n_ԭʼ����   Number;
  n_��������   Number;
  v_�շ����   �շ���ĿĿ¼.���%Type;
  n_����ⷿid ҩƷ�շ���¼.�ⷿid%Type;
  n_��������id ҩƷ�շ���¼.Id%Type;
  n_�ⷿid     ҩƷ�շ���¼.�ⷿid%Type;
  n_��������   Number;
  v_�շ�ids    Varchar2(4000); --�û���Ϣê�㷢�ͣ���ʽ:�շ�id,����|�շ�id,����...

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  --�ҳ���Ҫ�����ҩƷ����
  Cursor c_ҩƷ�շ���¼ Is
    Select ID, ����id, NO, ����, ҩƷid, �ⷿid, Nvl(����, 0) ����, ����, ����,
           Decode(��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(����, 1) * Nvl(ʵ������, 0) As ����, ʵ������, ����, ��ҩ��ʽ, ���Ч��, Ч��, ��Ʒ����, �ڲ�����
    From ҩƷ�շ���¼
    Where ���� In (8, 9, 10, 21, 24, 25, 26) And Mod(��¼״̬, 3) = 1 And ����� Is Null And ����id = ����id_In;

  Cursor c_������� Is
    Select /*+ rule*/
     a.Id, a.����id, a.No, a.����, a.ҩƷid, a.�ⷿid, Nvl(a.����, 0) ����, a.����, a.����,
     Decode(a.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(a.����, 1) * Nvl(a.ʵ������, 0) As ����, a.ʵ������, a.����, a.��ҩ��ʽ, a.���Ч��, a.Ч��, a.��Ʒ����,
     a.�ڲ�����
    From ҩƷ�շ���¼ A, Table(f_Str2list(��ҩid_In)) B, ��Һ��ҩ���� C
    Where a.���� In (9, 10, 25, 26) And Mod(a.��¼״̬, 3) = 1 And a.����� Is Null And a.����id = ����id_In And a.Id = c.�շ�id And
          c.��¼id = b.Column_Value
    Order By ��������;

  r_Row c_ҩƷ�շ���¼%RowType;
Begin
  n_���� := 0;
  v_No   := '';

  If ��������_In = 0 Then
  
    --����Ϊ�ձ�ʾ��ȫ��ɾ��
    --���α�
    Open c_ҩƷ�շ���¼;
  
    --�����α�
    Loop
      Fetch c_ҩƷ�շ���¼
        Into r_Row;
      Exit When c_ҩƷ�շ���¼%NotFound;
    
      Select ��� Into v_�շ���� From �շ���ĿĿ¼ Where ID = r_Row.ҩƷid;
    
      --����ҩƷ���
      If r_Row.�ⷿid Is Not Null Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + r_Row.����
        Where �ⷿid = r_Row.�ⷿid And ҩƷid = r_Row.ҩƷid And Nvl(����, 0) = Nvl(r_Row.����, 0) And ���� = 1;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����)
          Values
            (r_Row.�ⷿid, r_Row.ҩƷid, 1, Nvl(r_Row.����, 0), r_Row.Ч��, r_Row.����, r_Row.����, r_Row.����, r_Row.���Ч��,
             r_Row.��Ʒ����, r_Row.�ڲ�����);
        End If;
      
        --ɾ������Ŀ������
        Delete From ҩƷ���
        Where �ⷿid = r_Row.�ⷿid And ҩƷid = r_Row.ҩƷid And Nvl(����, 0) = Nvl(r_Row.����, 0) And ���� = 1 And Nvl(��������, 0) = 0 And
              Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
      
        Zl_ҩƷ���_���������쳣����(r_Row.�ⷿid, r_Row.ҩƷid, r_Row.����);
      End If;
    
      n_���� := r_Row.����;
      v_No   := r_Row.No;
      l_ҩƷ�շ�.Extend;
      l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := r_Row.Id;
    
      v_�շ�ids := v_�շ�ids || '|' || r_Row.Id || ',' || 0;
    End Loop;
  
    --�ر��α�
    Close c_ҩƷ�շ���¼;
  Else
    --������Ϊ�ձ�ʾ��������˲���
    If ��ҩid_In Is Not Null Then
      Open c_�������;
    Else
      Open c_ҩƷ�շ���¼;
    End If;
    n_�������� := ��������_In;
  
    --ֻ��סԺ���˴���Ż�����һ��
    Loop
      If ��ҩid_In Is Not Null Then
        Fetch c_�������
          Into r_Row;
        Exit When c_�������%NotFound;
      Else
        Fetch c_ҩƷ�շ���¼
          Into r_Row;
        Exit When c_ҩƷ�շ���¼%NotFound;
      End If;
    
      n_����ⷿid := Null;
      n_��������id := Null;
      Select ��� Into v_�շ���� From �շ���ĿĿ¼ Where ID = r_Row.ҩƷid;
      If v_�շ���� = '4' Then
        Begin
          Select 1, �ⷿid, ID
          Into n_��������, n_����ⷿid, n_��������id
          From ҩƷ�շ���¼
          Where ����id = ����id_In And ������� Is Null And ���� = 21 And Rownum = 1;
        Exception
          When Others Then
            n_�������� := 0;
        End;
      Else
        n_�������� := 0;
      End If;
    
      n_����     := r_Row.����;
      v_No       := r_Row.No;
      n_ԭʼ���� := r_Row.����;
    
      If n_�������� >= n_ԭʼ���� Then
        l_ҩƷ�շ�.Extend;
        l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := r_Row.Id;
        v_�շ�ids := v_�շ�ids || '|' || r_Row.Id || ',' || 0;
        If Nvl(n_��������id, 0) > 0 Then
          l_ҩƷ�շ�.Extend;
          l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := n_��������id;
          v_�շ�ids := v_�շ�ids || '|' || n_��������id || ',' || 0;
        End If;
        n_�������� := n_�������� - n_ԭʼ����;
      Else
        If v_�շ���� = '7' Then
          --��ǰ�е�����Ҫ��
          Update ҩƷ�շ���¼
          Set ���� = 1, ʵ������ = Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - n_��������,
              ��д���� = Decode(����, Null, 1, 0, 1, ����) * Nvl(��д����, 0) - n_��������,
              �ɱ���� =
               (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - n_��������) * �ɱ���,
              ���۽�� =
               (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - n_��������) * ���ۼ�,
              ��� = Round((Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - n_��������) * ���ۼ� -
                          (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - n_��������) * �ɱ���, 5)
          Where ID = r_Row.Id;
        Else
          Update ҩƷ�շ���¼
          Set ʵ������ = Nvl(ʵ������, 0) - n_��������, ��д���� = Nvl(��д����, 0) - n_��������,
              �ɱ���� =
               (Nvl(ʵ������, 0) - n_��������) * �ɱ���,
              ���۽�� =
               (Nvl(ʵ������, 0) - n_��������) * ���ۼ�,
              ��� = Round((Nvl(ʵ������, 0) - n_��������) * ���ۼ� - (Nvl(ʵ������, 0) - n_��������) * �ɱ���, 5)
          Where ID = r_Row.Id;
        End If;
      
        v_�շ�ids := v_�շ�ids || '|' || r_Row.Id || ',' || n_��������;
      
        --�����������ⵥ
        If Nvl(n_��������id, 0) <> 0 Then
          If v_�շ���� = '7' Then
            Update ҩƷ�շ���¼
            Set ���� = 1, ʵ������ = Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - n_��������,
                ��д���� = Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - n_��������,
                �ɱ���� =
                 (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - n_��������) * �ɱ���,
                ���۽�� =
                 (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - n_��������) * ���ۼ�,
                ��� = Round((Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - n_��������) * ���ۼ� -
                            (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - n_��������) * �ɱ���, 5)
            Where ID = Nvl(n_��������id, 0);
          Else
            Update ҩƷ�շ���¼
            Set ʵ������ = Nvl(ʵ������, 0) - n_��������, ��д���� = Nvl(ʵ������, 0) - n_��������,
                �ɱ���� =
                 (Nvl(ʵ������, 0) - n_��������) * �ɱ���,
                ���۽�� =
                 (Nvl(ʵ������, 0) - n_��������) * ���ۼ�,
                ��� = Round((Nvl(ʵ������, 0) - n_��������) * ���ۼ� - (Nvl(ʵ������, 0) - n_��������) * �ɱ���, 5)
            Where ID = Nvl(n_��������id, 0);
          End If;
        
          v_�շ�ids := v_�շ�ids || '|' || n_��������id || ',' || n_��������;
        End If;
        n_ԭʼ���� := n_��������;
        n_�������� := 0;
      End If;
      If Nvl(n_��������, 0) = 1 Then
        n_�ⷿid := n_����ⷿid;
      Else
        n_�ⷿid := r_Row.�ⷿid;
      End If;
    
      If n_�ⷿid Is Not Null Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + n_ԭʼ����
        Where �ⷿid = n_�ⷿid And ҩƷid = r_Row.ҩƷid And Nvl(����, 0) = Nvl(r_Row.����, 0) And ���� = 1;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��)
          Values
            (n_�ⷿid, r_Row.ҩƷid, 1, Nvl(r_Row.����, 0), r_Row.Ч��, n_ԭʼ����, r_Row.����, r_Row.����, r_Row.���Ч��);
        End If;
        Delete ҩƷ���
        Where �ⷿid = n_�ⷿid And ҩƷid = r_Row.ҩƷid And Nvl(����, 0) = Nvl(r_Row.����, 0) And ���� = 1 And Nvl(��������, 0) = 0 And
              Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
      
        Zl_ҩƷ���_���������쳣����(r_Row.�ⷿid, r_Row.ҩƷid, r_Row.����);
      End If;
    
      If Nvl(n_��������, 0) = 1 Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + n_ԭʼ����
        Where �ⷿid = r_Row.�ⷿid And ҩƷid = r_Row.ҩƷid And Nvl(����, 0) = Nvl(r_Row.����, 0) And ���� = 1;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��)
          Values
            (r_Row.�ⷿid, r_Row.ҩƷid, 1, Nvl(r_Row.����, 0), r_Row.Ч��, n_ԭʼ����, r_Row.����, r_Row.����, r_Row.���Ч��);
        End If;
      
        Delete ҩƷ���
        Where �ⷿid = r_Row.�ⷿid And ҩƷid = r_Row.ҩƷid And Nvl(����, 0) = Nvl(r_Row.����, 0) And ���� = 1 And Nvl(��������, 0) = 0 And
              Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
      
        Zl_ҩƷ���_���������쳣����(r_Row.�ⷿid, r_Row.ҩƷid, r_Row.����);
      End If;
    
      If n_�������� = 0 Then
        Exit;
      End If;
    End Loop;
  
    --���������ĵ�,�����:��Ϊ������Ļ�,������ҩƷ�շ���¼�д���
    If Nvl(n_��������, 0) <> 0 And Not (v_�շ���� = '4' And n_ԭʼ���� = 0) Then
      --δ�������,��ʾ��ҩƷ�����Ѿ�ִ��.
      v_Err_Msg := 'Ҫ���ʵķ����д����ѷ���ҩƷ�����ģ����ѱ����������ʣ�������ǲ�����������ġ�';
      Raise Err_Item;
    End If;
  End If;

  --ɾ��ҩƷ�շ���¼
  Forall I In 1 .. l_ҩƷ�շ�.Count
    Delete From ҩƷ�շ���¼ Where ID = l_ҩƷ�շ�(I) And ����� Is Null;

  --ɾ��δ��ҩƷ��¼
  Delete From δ��ҩƷ��¼ A
  Where NO = v_No And ���� = n_���� And Not Exists
   (Select 1
         From ҩƷ�շ���¼
         Where ���� = a.���� And Nvl(�ⷿid, 0) = Nvl(a.�ⷿid, 0) And NO = v_No And Mod(��¼״̬, 3) = 1 And ����� Is Null);

  --����������Ϣ
  If ��Ϣ_In = 1 Then
    b_Message.Zlhis_Charge_008(v_�շ����, ����id_In, v_�շ�ids);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ���¼_�����˷�;
/

--125779:��ҵ��,2018-05-15,��ҩ��ҩƷid������
Create Or Replace Procedure Zl_��Һ��ҩ��¼_�������
(
  ��ҩid_In   In Varchar2, --ID��:ID1,��˱�־1,ID2,��˱�־2....
  ������Ա_In In ��Һ��ҩ��¼.������Ա%Type,
  ����ʱ��_In In ��Һ��ҩ��¼.����ʱ��%Type
) Is
  v_Tansid     Varchar2(20);
  v_Tmp        Varchar2(4000);
  v_Usercode   Varchar2(100);
  v_��ҩid     ҩƷ�շ���¼.Id%Type;
  n_Count      Number(1);
  d_���ʱ��   ҩƷ�շ���¼.�������%Type;
  v_No         ҩƷ�շ���¼.No%Type;
  v_�ϴ�no     ҩƷ�շ���¼.No%Type;
  n_��˱�־   Number(1);
  n_����״̬   Number(2);
  v_�շ�ids    Varchar2(4000);
  v_��ҩ����id ҩƷ�շ���¼.Id%Type;
  v_ԭʼid     ҩƷ�շ���¼.Id%Type;
  v_Error      Varchar2(255);
  Err_Custom Exception;

  Cursor c_���ʼ�¼ Is
    Select Distinct a.����id, b.����ʱ��
    From ҩƷ�շ���¼ A, ��Һ��ҩ��¼ B, ��Һ��ҩ���� C
    Where a.Id = c.�շ�id And b.Id = c.��¼id And b.Id = v_Tansid And b.����״̬ = 9;

  v_���ʼ�¼ c_���ʼ�¼%RowType;

  Cursor c_��ҩ��¼ Is
    Select Distinct a.Id As ��ҩid, c.�շ�id, c.����, a.ҩƷid, a.����
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C
    Where c.��¼id = v_Tansid And b.Id = c.�շ�id And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And
          a.ҩƷid + 0 = b.ҩƷid And a.��� = b.��� And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)
    Order By a.ҩƷid, a.����;

  v_��ҩ��¼ c_��ҩ��¼%RowType;

  Cursor c_�������� Is
    Select a.No, a.��� || ':' || c.���� || ':' || c.��¼id As �������
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C
    Where a.Id = b.����id And b.Id = c.�շ�id And Mod(b.��¼״̬, 3) = 1 And c.��¼id = v_Tansid;

  v_�������� c_��������%RowType;

Begin
  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;

  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);

  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Tansid   := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
    n_��˱�־ := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
  
    v_�շ�ids := Null;
  
    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ����ҩ״̬
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Tansid;
    
      If n_����״̬ <> 9 Then
        v_Error := '�������ѱ����������ܽ���������ˣ�';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    If n_��˱�־ = 1 Then
      n_����״̬ := 10;
    Elsif n_��˱�־ = 2 Then
      n_����״̬ := 11;
    End If;
  
    --������Һ����Ӧ���շ�NO
    Begin
      Select NO
      Into v_No
      From ҩƷ�շ���¼
      Where ID In (Select �շ�id From ��Һ��ҩ���� Where ��¼id In (Select ID From ��Һ��ҩ��¼ Where ID = v_Tansid)) And Rownum < 2;
    Exception
      When Others Then
        v_No := '';
    End;
  
    --�շ�NO��ͬ����ҩID�����ʱ���Դ�����Ϊ�ӳ�1��
    If v_No = v_�ϴ�no Then
      d_���ʱ�� := d_���ʱ�� + 1 / 24 / 60 / 60;
    Else
      d_���ʱ�� := ����ʱ��_In;
      v_�ϴ�no   := v_No;
    End If;
  
    --���ʼ�¼����
    For v_���ʼ�¼ In c_���ʼ�¼ Loop
      Zl_���˷�������_Audit(v_���ʼ�¼.����id, v_���ʼ�¼.����ʱ��, ������Ա_In, d_���ʱ��, n_��˱�־);
    End Loop;
  
    Select Count(*) Into n_Count From ��Һ��ҩ״̬ Where ��ҩid = v_Tansid And ����ʱ�� = ����ʱ��_In;
  
    If n_Count <> 1 Then
      Insert Into ��Һ��ҩ״̬
        (��ҩid, ��������, ������Ա, ����ʱ��)
      Values
        (v_Tansid, n_����״̬, ������Ա_In, ����ʱ��_In);
    End If;
    Update ��Һ��ҩ��¼ Set ������Ա = ������Ա_In, ����ʱ�� = ����ʱ��_In, ����״̬ = n_����״̬ Where ID = v_Tansid;
  
    --��ҩ����
    If n_��˱�־ = 1 Then
      For v_��ҩ��¼ In c_��ҩ��¼ Loop
        Zl_ҩƷ�շ���¼_������ҩ(v_��ҩ��¼.��ҩid, ������Ա_In, ����ʱ��_In, Null, Null, Null, v_��ҩ��¼.����, Null, ������Ա_In);
      
        --ȡ��ҩ����id
        Select a.Id
        Into v_��ҩid
        From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
        Where b.Id = v_��ҩ��¼.��ҩid And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
              a.��� = b.��� And Mod(a.��¼״̬, 3) = 1 And a.������� Is Null;
      
        --��Һ��ҩ�����е��շ�ID����Ϊ��ҩ�������շ�ID
        Update ��Һ��ҩ���� Set �շ�id = v_��ҩid Where ��¼id = v_Tansid And �շ�id = v_��ҩ��¼.�շ�id;
      
        If v_�շ�ids Is Null Then
          v_�շ�ids := v_��ҩid;
        Else
          v_�շ�ids := v_�շ�ids || ',' || v_��ҩid;
        End If;
      
        --ȡԭʼid
        Select a.Id
        Into v_ԭʼid
        From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
        Where b.Id = v_��ҩ��¼.��ҩid And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
              a.��� = b.��� And Mod(a.��¼״̬, 3) = 0 And a.������� Is Not Null;
      
        Insert Into ��Һ��ҩ����
          (��¼id, �շ�id, ����)
          Select ��¼id, v_ԭʼid, ���� From ��Һ��ҩ���� Where ��¼id = v_Tansid And �շ�id = v_��ҩid;
      
        v_�շ�ids := v_�շ�ids || ',' || v_ԭʼid;
      End Loop;
    
      --��������
      For v_�������� In c_�������� Loop
        Zl_סԺ���ʼ�¼_Delete(v_��������.No, v_��������.�������, v_Usercode, Zl_Username, 2, 1, 1, d_���ʱ��);
      End Loop;
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_�������;
/




------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0011' Where ���=&n_System;
Commit;
