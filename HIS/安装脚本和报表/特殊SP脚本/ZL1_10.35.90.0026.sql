----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--125796:����,2018-08-22,���ӳ�����¼�ֶ�
alter table ҩƷ�շ���¼ add ����ԭ�� varchar2(200);

--128856:��¶¶,2018-08-23,���������Ŀʱϣ�����˶��סԺ���ž��͵�һ�β�����һ�µ�����
Alter Table סԺ������¼ Drop Constraint סԺ������¼_UQ_������ Cascade Drop Index;
Create Index סԺ������¼_IX_������ On סԺ������¼(������) PCTFREE 5 Tablespace zl9IndexMdr;
alter index סԺ������¼_IX_����ID rename to סԺ������¼_IX_������;
------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--130541:������,2018-08-23,����ƽ̨��Ϣ����
Insert Into Zlmsg_Lists(Bz_Type, Code, Name, Key_Define, Note) 
Select '�ٴ�', 'ZLHIS_CIS_059', 'ȷ��ֹͣ����ҽ��', '<root><����ID></����ID><��ҳID></��ҳID><ID></ID></root>', 'סԺ��ʿ����վ:ȷ��ֹͣ����ҽ��ʱ'  From Dual;

--130471:��ΰ��,2018-08-21,ԤԼ���Ų�ѯ
Insert Into ������������Ŀ¼ (ϵͳ��ʶ, ��������) Values ('ԤԼ����', 'ԤԼ���Ų�ѯ');

--130469:����,2018-08-21,������Ϣ
--ZLMSG_LISTS
Insert Into Zlmsg_Lists(Bz_Type, Code, Name, Key_Define, Note)
Select 'BLOOD', 'ZLHIS_BLOOD_003', '��Ѫ������', '<root><ҽ��ID></ҽ��ID><�շ�ID></�շ�ID></root>', '������Ѫ����:����Ѫ�����Ѫʱ'  From Dual Union All 
Select 'BLOOD', 'ZLHIS_BLOOD_004', 'ȡ����Ѫ���', '<root><ҽ��ID></ҽ��ID><�շ�ID></�շ�ID><�����></�����><���ʱ��></���ʱ��></root>', '������Ѫ����:����Ѫȡ����Ѫʱ'  From Dual Union All 
Select 'BLOOD', 'ZLHIS_BLOOD_005', '���ѪҺ����', '<root><ҽ��ID></ҽ��ID><�շ�ID></�շ�ID></root>', '���ҷ�Ѫ����:���ȡѪ�����ѪҺ����ʱ'  From Dual Union All 
Select 'BLOOD', 'ZLHIS_BLOOD_006', 'ȡ��ѪҺ����', '<root><ҽ��ID></ҽ��ID><�շ�ID></�շ�ID><ԭ�շ�ID></ԭ�շ�ID></root>', '���ҷ�Ѫ����:ȡ���ѷ�ѪҺʱ'  From Dual Union All 
Select 'BLOOD', 'ZLHIS_BLOOD_007', '��Ѫִ�еǼ�', '<root><ҽ��ID></ҽ��ID><�շ�ID></�շ�ID></root>', 'ҽ������վ:��Ѫҽ��ִ�еǼ�ʱ'  From Dual Union All 
Select 'BLOOD', 'ZLHIS_BLOOD_008', '��Ѫִ�еǼ�ɾ��', '<root><ҽ��ID></ҽ��ID><�շ�ID></�շ�ID><ִ����></ִ����><ִ��ʱ��></ִ��ʱ��><�˶���></�˶���><�˶�ʱ��></�˶�ʱ��><������></������></root>', 'ҽ������վ;סԺ��ʿ����վ:ȡ����Ѫҽ��ִ�еǼ�ʱ'  From Dual Union All 
Select 'BLOOD', 'ZLHIS_BLOOD_009', '��Ѫҽ������״̬', '<root><ҽ��ID></ҽ��ID><��������></��������></root>', '������Ѫ����:���յȴ���Ѫ��ҽ��ʱ;��������Ѫ������תΪΪ�ȴ���Ѫʱ'  From Dual  Union All 
Select 'BLOOD', 'ZLHIS_BLOOD_010', '��Ѫҽ��������', '<root><ҽ��ID></ҽ��ID></root>', '��Ѫ��˹���:��Ѫҽ����˻�ܾ����ʱ'  From Dual;



-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--130541:������,2018-08-23,����ƽ̨��Ϣ����
Create Or Replace Procedure Zl_����ҽ����¼_����
(
  ҽ��id_In     In ����ҽ����¼.Id%Type,
  Flag_In       In Number := 0,
  ҽ������_In   In ����ҽ����¼.ҽ������%Type := Null,
  ��������_In   In ����ҽ��״̬.��������%Type := Null,
  ����Ա���_In In ��Ա��.���%Type := Null,
  ����Ա����_In In ��Ա��.����%Type := Null
  --���ܣ�����סԺҽ����״̬�������Ͳ���(������������ͨ������Zl_����ҽ����¼_��������������) 
  --������ҽ��ID_IN=һ��ҽ��ID 
  --      FLAG_IN=�������ݡ�����ֹͣ��0=���ִ����ֹʱ��,1=�������е�ִ����ֹʱ�䡣 
  --      ҽ������_IN=�ù��̱��������˵���ʱ���ã����ڴ�����ʾ�� 
  --      ��������_IN=�ù��̱��������˵���ʱ���ã����ں˶Ի������ݡ�0-���˷���,n=���˾���ҽ������ 
) Is
  --����ָ��ҽ���Ĳ�����¼,��һ��ΪҪ���˵�����(״̬��������) 
  --���������˷��ͺ���Զ�ֹͣ,�ڻ��˷���ʱ�Զ�����ֹͣ���� 
  Cursor c_Rolladvice Is
    Select b.������Ա, b.����ʱ��, 0 As ���ͺ�, Null As NO, b.��������, 0 As ִ��״̬, Sysdate + Null As �״�ʱ��, Sysdate + Null As ĩ��ʱ��,
           a.�ϴ�ִ��ʱ��, a.ҽ����Ч, a.������� As ���, a.������Ŀid, Null As ����, a.����id, a.��ҳid, a.Ӥ��, 0 As ��¼����, 0 As �������, 0 As ��������id,
           a.��˱��, a.����ҽ��, a.ִ�п���id, Nvl(a.���id, a.Id) As ��id, a.���id, a.Id As ҽ��id, -null As ��������, Null As ��������
    From ����ҽ����¼ A, ����ҽ��״̬ B
    Where a.Id = b.ҽ��id And (a.Id = ҽ��id_In Or a.���id = ҽ��id_In) And
          (Nvl(a.ҽ����Ч, 0) = 0 And b.�������� Not In (1, 2, 3) Or Nvl(a.ҽ����Ч, 0) = 1 And b.�������� Not In (1, 2, 3, 8))
    Union
    Select b.������ As ������Ա, b.����ʱ�� As ����ʱ��, b.���ͺ�, b.No, -null As ��������, b.ִ��״̬, b.�״�ʱ��, b.ĩ��ʱ��, a.�ϴ�ִ��ʱ��, a.ҽ����Ч, c.���,
           a.������Ŀid, c.�������� As ����, a.����id, a.��ҳid, a.Ӥ��, b.��¼����, b.�������, a.��������id, a.��˱��, a.����ҽ��, a.ִ�п���id,
           Nvl(a.���id, a.Id) As ��id, a.���id, a.Id As ҽ��id, b.��������, b.��������
    From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C
    Where a.Id = b.ҽ��id And a.������Ŀid = c.Id And (a.Id = ҽ��id_In Or a.���id = ҽ��id_In)
    Order By ����ʱ�� Desc, ���ͺ�;
  r_Rolladvice c_Rolladvice%RowType;

  --��ʽͬc_Rolladvice��ֻȡ���Ͳ��������Զ����˴��� 
  Cursor c_Rollsend(v_���ͺ� ����ҽ������.���ͺ�%Type) Is
    Select Distinct b.ҽ��id, b.����ʱ�� As ����ʱ��, b.���ͺ�, b.ִ��״̬, a.������� As ���, c.��ǰ����id As ���˲���id, a.���˿���id,
                    b.ִ�в���id As ִ�п���id
    From ����ҽ����¼ A, ����ҽ������ B, ������ҳ C
    Where a.Id = b.ҽ��id And b.���ͺ� = v_���ͺ� And a.����id = c.����id And a.��ҳid = c.��ҳid And
          (a.Id = ҽ��id_In Or a.���id = ҽ��id_In)
    Order By b.����ʱ�� Desc, b.���ͺ�;

  --����ҽ��������NO������λ���Ҫ���ʵķ��ü�¼ 
  --һ��ҽ�������Ƕ���д�˷��ͼ�¼,�ҿ���NO��ͬ(ҩƷ��,�÷��巨��һ����) 
  --���ܷ��ͼ�¼�ļƷ�״̬(��������Ʒ�),�з��ü�¼��Ȼ�������� 
  --����ֻ��۸񸸺�Ϊ�յ�,�Ա�ȡ������� 
  --ֻ�ܼ�¼״̬Ϊ1�ķ���,���������ʻ򲿷����ʵļ�¼,���ٴ�������"��¼״̬=3"�Ķ�ȡ�����������жϣ������� 
  Cursor c_Rollmoneyout
  (
    v_���ͺ�    ����ҽ������.���ͺ�%Type,
    v_ҽ��id    ����ҽ����¼.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.��¼״̬, a.No, a.���, a.�շ����, a.ִ��״̬, d.��������, a.ִ�в���id, a.��¼����
    From ������ü�¼ A, Table(t_Adviceids) B, ����ҽ������ C, �������� D
    Where c.ҽ��id = b.Column_Value And c.���ͺ� = v_���ͺ� And a.ҽ����� = b.Column_Value And
          (a.ҽ����� = v_ҽ��id Or Nvl(v_ҽ��id, 0) = 0) And a.��¼״̬ In (0, 1, 3) And a.No = c.No And a.��¼���� = c.��¼���� And
          a.�۸񸸺� Is Null And a.�շ�ϸĿid = d.����id(+)
    Order By a.No, a.���;

  Cursor c_Rollmoneyin
  (
    v_���ͺ�    ����ҽ������.���ͺ�%Type,
    v_ҽ��id    ����ҽ����¼.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.��¼״̬, a.No, a.���, a.�շ����, a.ִ��״̬, d.��������, a.ִ�в���id, a.��¼����
    From סԺ���ü�¼ A, Table(t_Adviceids) B, ����ҽ������ C, �������� D
    Where c.ҽ��id = b.Column_Value And c.���ͺ� = v_���ͺ� And a.ҽ����� = b.Column_Value And
          (a.ҽ����� = v_ҽ��id Or Nvl(v_ҽ��id, 0) = 0) And a.��¼״̬ In (0, 1, 3) And a.No = c.No And a.��¼���� = c.��¼���� And
          a.�۸񸸺� Is Null And a.�շ�ϸĿid = d.����id(+)
    Order By a.No, a.���;

  --ȡ����סԺ����ʱ�Զ����ŵ�����(��û�����ϵ�) 
  Cursor c_Stuff_Drug(v_����id ҩƷ�շ���¼.����id%Type) Is
    Select ID
    From ҩƷ�շ���¼
    Where ����id = v_����id And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0) And ����� Is Not Null
    Order By ҩƷid;

  --���ڴ�������ҽ���Ļ��� 
  Cursor c_Patilog
  (
    v_����id ���˱䶯��¼.����id%Type,
    v_��ҳid ���˱䶯��¼.��ҳid%Type
  ) Is
    Select *
    From ���˱䶯��¼
    Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� Is Null
    Order By ��ʼʱ�� Desc;
  r_Patilog c_Patilog%RowType;

  Cursor c_Adviceids Is
    Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In;
  t_Adviceids t_Numlist;

  v_ҽ��״̬     ����ҽ����¼.ҽ��״̬%Type;
  v_ҽ����Ч     ����ҽ����¼.ҽ����Ч%Type;
  v_����no       ����ҽ������.No%Type;
  v_�������     Varchar2(255);
  v_ĩ��ʱ��     ����ҽ������.ĩ��ʱ��%Type;
  v_����ʱ��     ����ҽ��״̬.����ʱ��%Type;
  v_��������     ������ĿĿ¼.��������%Type;
  v_ִ��Ƶ��     ������ĿĿ¼.ִ��Ƶ��%Type;
  v_�ϴ�ʱ��     ����ҽ����¼.�ϴ�ִ��ʱ��%Type;
  v_ִ��ʱ��     ����ҽ����¼.ִ��ʱ�䷽��%Type;
  v_��ʼִ��ʱ�� ����ҽ����¼.��ʼִ��ʱ��%Type;
  v_�ϴδ�ӡʱ�� ����ҽ����¼.�ϴδ�ӡʱ��%Type;
  v_Ƶ�ʼ��     ����ҽ����¼.Ƶ�ʼ��%Type;
  v_�����λ     ����ҽ����¼.�����λ%Type;
  v_���ͺ�       ����ҽ������.���ͺ�%Type;
  n_����ȼ�id   ���˱䶯��¼.����ȼ�id%Type;
  d_��ʼʱ��     ���˱䶯��¼.��ʼʱ��%Type;
  d_����ʱ��     ����ҽ��״̬.����ʱ��%Type;
  v_Tmp���ͺ�    ����ҽ������.���ͺ�%Type;
  n_ִ��         Number;

  Intdigit   Number(3);
  v_Update   Number(1);
  v_Count    Number(5);
  v_Temp     Varchar2(2000);
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;
  v_Time     Varchar2(4000);
  n_Blndo    Number;

  v_Error Varchar2(2000);
  Err_Custom Exception;

  Function Checkmoneyundo
  (
    v_No       סԺ���ü�¼.No%Type,
    v_��¼���� סԺ���ü�¼.��¼����%Type,
    v_���     סԺ���ü�¼.���%Type,
    n_����     Number := 0 --0סԺ��1���� 
  ) Return Number Is
    n_Num      Number;
    n_ִ��״̬ Number;
  Begin
    n_Num := 0;
    If n_���� = 0 Then
      Select Nvl(Sum(Nvl(����, 1) * ����), 0) As ����
      Into n_Num
      From סԺ���ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ In (2, 3);
      Select Nvl(ִ��״̬, 0)
      Into n_ִ��״̬
      From סԺ���ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ = 3;
    Else
      Select Nvl(Sum(Nvl(����, 1) * ����), 0) As ����
      Into n_Num
      From ������ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ In (2, 3);
      Select Nvl(ִ��״̬, 0)
      Into n_ִ��״̬
      From ������ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ = 3;
    End If;
    If n_Num <> 0 Then
      n_Num := 1;
    End If;
    --�������¼����ִ�У�����ִ�еģ����Զ��ˡ� 
    If n_ִ��״̬ <> 0 Then
      n_Num := 0;
    End If;
    Return(n_Num);
  End;
Begin
  v_Tmp���ͺ� := -1;
  Open c_Rolladvice;
  Loop
    Fetch c_Rolladvice
      Into r_Rolladvice;
    If c_Rolladvice%RowCount = 0 Then
      Close c_Rolladvice;
      v_Error := Nvl(ҽ������_In, '��ҽ��') || '��ǰû�п��Ի��˵����ݡ�';
      Raise Err_Custom;
    End If;
    Exit When c_Rolladvice%NotFound;
    Exit When d_����ʱ�� <> r_Rolladvice.����ʱ�� And d_����ʱ�� Is Not Null;
    d_����ʱ�� := r_Rolladvice.����ʱ��;
  
    --�������˵���ʱ�ж� 
    If ҽ������_In Is Not Null Then
      If Nvl(r_Rolladvice.��������, 0) <> Nvl(��������_In, 0) Then
        v_Error := Nvl(ҽ������_In, '��ҽ��') || '�����뵱ǰҽ��һ����ˣ����ܸ�ҽ���Ѿ�ִ��������������';
        Raise Err_Custom;
      End If;
    End If;
  
    --һ�鷢�ͺ�ִֻ��һ�� 
    If v_Tmp���ͺ� <> r_Rolladvice.���ͺ� Then
      v_Tmp���ͺ� := r_Rolladvice.���ͺ�;
      n_ִ��      := 1;
    Else
      n_ִ�� := 0;
    End If;
  
    If n_ִ�� = 1 Then
      Open c_Adviceids;
      Fetch c_Adviceids Bulk Collect
        Into t_Adviceids;
      Close c_Adviceids;
    
      If r_Rolladvice.���ͺ� = 0 Then
        --����ҽ��״̬����(��ʱ��ؼ���) 
        --4-���ϣ�5-������6-��ͣ��7-���ã�8-ֹͣ��9-ȷ��ֹͣ��10-Ƥ�Խ��;13-ͣ������ 
        ------------------------------------------------------------------ 
        --���ֻ���˻ص�У��״̬ 
        If r_Rolladvice.�������� = 3 Then
          v_Error := Nvl(ҽ������_In, '��ҽ��') || '��ǰ����ͨ��У��״̬�������ٻ��ˡ�';
          Raise Err_Custom;
        Elsif r_Rolladvice.�������� = 4 And Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
          If r_Rolladvice.��� = 'H' Then
            Select ��������, ִ��Ƶ�� Into v_��������, v_ִ��Ƶ�� From ������ĿĿ¼ Where ID = r_Rolladvice.������Ŀid;
            If v_�������� = '1' And v_ִ��Ƶ�� = '2' Then
              v_Error := '����ȼ����Ϻ����ٻ��ˡ�';
              Raise Err_Custom;
            End If;
          End If;
        End If;
      
        --����Ƿ�������������֮ǰ�Ĳ��� 
        If r_Rolladvice.�������� <> 5 Then
          --ȡ�������ʱ�� 
          Select Nvl(ҽ������ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD'))
          Into v_����ʱ��
          From ������ҳ
          Where ����id = r_Rolladvice.����id And ��ҳid = r_Rolladvice.��ҳid;
        
          If r_Rolladvice.����ʱ�� < v_����ʱ�� Then
            v_Error := '�ò������������֮ǰ�Ĳ��������ٻ��ˡ�';
            Raise Err_Custom;
          End If;
        End If;
      
        --ɾ��(����ҽ��)�����״̬������¼ 
        Delete /*+ Rule*/
        From ����ҽ��״̬
        Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And ����ʱ�� = r_Rolladvice.����ʱ��;
      
        --ȡɾ����Ӧ�ָ���ҽ��״̬ 
        Select ��������
        Into v_ҽ��״̬
        From ����ҽ��״̬
        Where ����ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = ҽ��id_In) And ҽ��id = ҽ��id_In;
      
        --�ָ�(����ҽ��)���˺��״̬ 
        Update ����ҽ����¼ Set ҽ��״̬ = v_ҽ��״̬ Where ID = ҽ��id_In Or ���id = ҽ��id_In;
      
        --��������Ĵ��� 
        If r_Rolladvice.�������� = 8 Then
          --�����ڷ����ջع���ҽ�� ���������������ģʽ�����ж϶�Ӧ�ġ����˷������ʡ������Ƿ�ȡ��������������ˣ��������� 
          --                       ����ǲ�����������ģʽ���������ٻ��ˡ� 
          --���ܳ��ڷ����ջ�ʱ��ȫ���ջ�(���ϴ�ִ��ʱ��) 
          Select /*+ Rule*/
           Nvl(Count(*), 0)
          Into v_Count
          From ����ҽ����¼ A, ����ҽ������ B
          Where b.ҽ��id = a.Id And (a.Id = ҽ��id_In Or a.���id = ҽ��id_In) And
                b.���ͺ� =
                (Select Max(���ͺ�) From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(t_Adviceids))) And
                a.ִ����ֹʱ�� Is Not Null And ((a.�ϴ�ִ��ʱ�� < b.ĩ��ʱ��) Or (a.�ϴ�ִ��ʱ�� Is Null And b.ĩ��ʱ�� Is Not Null));
          If v_Count > 0 Then
            If zl_GetSysParameter('�����ջز�����������', 1254) = '1' Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '�ѱ����ڷ����ջأ������ٳ���ֹͣ������';
              Raise Err_Custom;
            Else
              --����Ѿ�ȡ���������룬���������. 
              Select Count(1)
              Into v_Count
              From ���˷������� A, סԺ���ü�¼ B, ����ҽ����¼ C
              Where a.����id = b.Id And c.Id = b.ҽ����� And (c.Id = ҽ��id_In Or c.���id = ҽ��id_In);
              If v_Count > 0 Then
                v_Error := Nvl(ҽ������_In, '��ҽ��') || '�ѱ����ڷ����ջأ������ٳ���ֹͣ������';
                Raise Err_Custom;
              Else
                --�õ��ϴ�ִ��ʱ�����Ϣ 
                Select �ϴ�ִ��ʱ��, ִ��ʱ�䷽��, ��ʼִ��ʱ��, �ϴδ�ӡʱ��, Ƶ�ʼ��, �����λ
                Into v_�ϴ�ʱ��, v_ִ��ʱ��, v_��ʼִ��ʱ��, v_�ϴδ�ӡʱ��, v_Ƶ�ʼ��, v_�����λ
                From ����ҽ����¼
                Where ID = ҽ��id_In;
                v_�ϴ�ʱ�� := To_Date(To_Char(v_�ϴ�ʱ�� + 1 / 24 / 60 / 60, 'yyyy-MM-dd hh24:mi:ss'), 'yyyy-MM-dd hh24:mi:ss');
              
                --�޸��ϴ�ִ��ʱ��Ϊ�ջغ��ĩ��ִ��ʱ�䡣 
                v_ĩ��ʱ�� := Null;
                Begin
                  --һ��ҽ���ķ�����ĩʱ����ͬ,һ����ҩ��ȡ��С�� 
                  --ȡ���IDΪNULL��ҽ���ķ��ͼ�¼��ʱ�� 
                  --����ҩ;������ҩ�÷�����δ��д���ͼ�¼ 
                  Select /*+ Rule*/
                   ĩ��ʱ��, ���ͺ�
                  Into v_ĩ��ʱ��, v_���ͺ�
                  From ����ҽ������
                  Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And
                        ���ͺ� = (Select Max(���ͺ�)
                               From ����ҽ������
                               Where ҽ��id In (Select Column_Value From Table(t_Adviceids))) And Rownum = 1;
                Exception
                  When Others Then
                    Null;
                End;
                Update ����ҽ����¼ Set �ϴ�ִ��ʱ�� = v_ĩ��ʱ�� Where ID = ҽ��id_In Or ���id = ҽ��id_In;
              
                Select Count(1) Into v_Count From ����ҽ������ Where ҽ��id = ҽ��id_In And ���ͺ� = v_���ͺ�;
                If v_Count > 0 Then
                  --��ԭҽ��ִ��ʱ�� 
                  Select Zl_Adviceexetimes(ҽ��id_In, v_�ϴ�ʱ��, v_ĩ��ʱ��, v_ִ��ʱ��, v_��ʼִ��ʱ��, v_�ϴδ�ӡʱ��, v_Ƶ�ʼ��, v_�����λ, 0)
                  Into v_Time
                  From Dual;
                  Insert Into ҽ��ִ��ʱ��
                    (Ҫ��ʱ��, ҽ��id, ���ͺ�)
                    Select To_Date(Column_Value, 'yyyy-mm-dd hh24:mi:ss'), ҽ��id_In, v_���ͺ�
                    From Table(f_Str2list(v_Time));
                End If;
              End If;
            End If;
          End If;
        
          --����ȼ��䶯�������������䶯ʱ����������� 
          If r_Rolladvice.��� = 'H' And Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
            Select ��������, ִ��Ƶ�� Into v_��������, v_ִ��Ƶ�� From ������ĿĿ¼ Where ID = r_Rolladvice.������Ŀid;
            If v_�������� = '1' And v_ִ��Ƶ�� = '2' Then
              Select Count(*), Max(a.����ȼ�id), Max(a.��ʼʱ��)
              Into v_Count, n_����ȼ�id, d_��ʼʱ��
              From ���˱䶯��¼ A
              Where a.����id = r_Rolladvice.����id And a.��ҳid = r_Rolladvice.��ҳid And a.��ʼԭ�� = 6 And a.��ֹʱ�� Is Null And
                    a.���Ӵ�λ = 0;
              --���û���ҵ����һ���ǻ���ȼ��䶯���ֹ 
              If v_Count = 0 Then
                --ҽ������ȼ�����סʱ��Ļ���ȼ�һ��ʱҪ�����ж� 
                Select Count(*)
                Into v_Count
                From ���˱䶯��¼ A
                Where a.����id = r_Rolladvice.����id And a.��ҳid = r_Rolladvice.��ҳid And a.��ʼԭ�� = 6;
                If v_Count > 0 Then
                  v_Error := '���ڻ���ȼ�ҽ��ֹͣ��ò����Ѿ������������䶯��¼,���ܻ��˸�ҽ����ֹͣ������';
                  Raise Err_Custom;
                End If;
              Else
                --���n_����ȼ�IDΪNull�������Ƿ��ǵ�ǰ���˵�ҽ����Ӧ�ı䶯��¼,Ŀ�����ж������ȼ�ҽ��ʱҪ��˳����ˡ� 
                --���n_����ȼ�ID��ΪNull�����п�����У����һ������ȼ�ʱ���Զ�ֹͣ�ģ�δ�����䶯��¼�� 
                --     ����Ҫ��鵱ǰ���һ���䶯�Ļ���ȼ�ID�Ƿ��ǵ�ǰҽ���Ļ���ȼ�ID,Ŀ�����ж������ȼ�ҽ��ʱҪ��˳����ˣ����������Ҫ�ٳ������һ�α䶯��ֱ�ӻ���ҽ�����ɡ� 
                If n_����ȼ�id Is Null Then
                  Select Count(*)
                  Into v_Count
                  From ���˱䶯��¼ B, ����ҽ���Ƽ� C
                  Where b.����id = r_Rolladvice.����id And b.��ҳid = r_Rolladvice.��ҳid And c.ҽ��id = ҽ��id_In And
                        c.�շ�ϸĿid = b.����ȼ�id And b.��ֹʱ�� = d_��ʼʱ�� And b.��ֹԭ�� = 6 And b.���Ӵ�λ = 0;
                Else
                  --��ʼʱ��ֻȡ���ӶԱȣ�У�Ե�ʱ����ȼ��Ŀ�ʼʱ����ҽ����ʼʱ��+��ǰʱ������� 
                  Select Count(*)
                  Into v_Count
                  From ����ҽ���Ƽ� C, ����ҽ����¼ A
                  Where a.Id = c.ҽ��id And a.Id = ҽ��id_In And c.�շ�ϸĿid = n_����ȼ�id And
                        a.��ʼִ��ʱ�� = To_Date(To_Char(d_��ʼʱ��, 'yyyy-mm-dd hh24:mi'), 'yyyy-mm-dd hh24:mi');
                End If;
                If v_Count = 0 Then
                  v_Error := '�����˵�ҽ���������һ������ȼ�ҽ�����뽫����Ļ���ȼ�ҽ�����Ϻ��ٻ��˱���ҽ����';
                  Raise Err_Custom;
                End If;
              
                If n_����ȼ�id Is Null Then
                  --��ǰ������Ա 
                  If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null Then
                    v_��Ա��� := ����Ա���_In;
                    v_��Ա���� := ����Ա����_In;
                  Else
                    v_Temp     := Zl_Identity;
                    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
                    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
                    v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
                    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
                  End If;
                
                  Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, '1', Null, Null, '����ȼ��䶯');
                End If;
              End If;
            End If;
          End If;
        
          If r_Rolladvice.��� = 'Z' And Instr(',9,10,', ',' || r_Rolladvice.���� || ',') > 0 And
             Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
            --���˲���ҽ��ʱ�����ñ䶯��¼���� 
            Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, Null, Null, Null, '�����䶯');
          End If;
        
          --����ҽ��ֹͣʱ,���ͣ��ҽ����ʱ��,�����ʵϰҽʦ�������˵ģ���ָ������״̬ 
          Update ����ҽ����¼
          Set ִ����ֹʱ�� = Decode(Flag_In, 1, ִ����ֹʱ��, Null), ͣ��ҽ�� = Null, ͣ��ʱ�� = Null,
              ��˱�� = Decode(r_Rolladvice.��˱��, 3, 2, r_Rolladvice.��˱��)
          Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        Elsif r_Rolladvice.�������� = 9 Then
          --����ҽ��ȷ��ֹͣʱ,����Ƿ��Ѵ�ӡͣ��ʱ�� 
          Select /*+ Rule*/
           Count(*)
          Into v_Count
          From ����ҽ����ӡ
          Where ��ӡ��� = 1 And ҽ��id In (Select Column_Value From Table(t_Adviceids));
          If v_Count > 0 Then
            v_Error := Nvl(ҽ������_In, '��ҽ��') || '��ͣ��ʱ���Ѿ���ӡ�������ٳ���ȷ��ֹͣ������';
            Raise Err_Custom;
          End If;
        
          --����ҽ��ȷ��ֹͣʱ,���ͣ��ҽ����ʱ�� 
          Update ����ҽ����¼ Set ȷ��ͣ��ʱ�� = Null, ȷ��ͣ����ʿ = Null Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        Elsif r_Rolladvice.�������� = 10 Then
          --���˱�עƤ�Խ��,ͬʱɾ�������Ǽ�(+)��(-),���ݼ�¼ʱ�� 
          --�����ļ�¼��ҽ�������޹ۣ�����Ҫ���� 
          Delete From ���˹�����¼
          Where ����id = r_Rolladvice.����id And Nvl(��ҳid, 0) = Nvl(r_Rolladvice.��ҳid, 0) And ��¼ʱ�� = r_Rolladvice.����ʱ�� And
                Nvl(���, 0) = 0;
        
          Update ����ҽ����¼ Set Ƥ�Խ�� = Null Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        Elsif r_Rolladvice.�������� = 13 Then
          If Instr(r_Rolladvice.����ҽ��, '/') > 0 Then
            Update ����ҽ����¼ Set ��˱�� = 1 Where ID = ҽ��id_In Or ���id = ҽ��id_In;
          Else
            Update ����ҽ����¼ Set ��˱�� = Null Where ID = ҽ��id_In Or ���id = ҽ��id_In;
          End If;
        End If;
        --��������ҽ�����ϲ��� 
        --����ҽ�� 
        If r_Rolladvice.�������� = 4 And r_Rolladvice.��� = 'Z' Then
          Select Count(1) Into v_Count From ������ĿĿ¼ Where ID = r_Rolladvice.������Ŀid And �������� = '4';
          If v_Count = 1 Then
            b_Message.Zlhis_Cis_004(r_Rolladvice.����id, r_Rolladvice.��ҳid, ҽ��id_In);
          End If;
        End If;
      Else
        --����ҽ������(�Է��ͺŹؼ���) 
        ------------------------------------------------------------------ 
        --��ǰ������Ա 
        v_Temp     := Zl_Identity;
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      
        --����Ƿ�����Һ��Һ��¼�����Ƿ��Ѿ������������ѯ������˵������Һ��¼ 
        Begin
          Select Decode(Max(�Ƿ�����), 1, 1, 0)
          Into v_Count
          From ��Һ��ҩ��¼
          Where ҽ��id = ҽ��id_In And ���ͺ� = r_Rolladvice.���ͺ�;
        Exception
          When Others Then
            v_Count := -1;
        End;
      
        If v_Count = 1 Then
          v_Error := 'ҽ��"' || ҽ������_In || '"����ҺҩƷ���Ѿ�����Һ�����������������ܻ��˷��͡�';
          Raise Err_Custom;
        Elsif v_Count = 0 Then
          Zl_��Һ��ҩ��¼_ҽ������(ҽ��id_In, r_Rolladvice.���ͺ�, v_��Ա����, Sysdate);
        End If;
      
        --����Ƿ����δ��˵��������� 
        Select Count(*)
        Into v_Count
        From ����ҽ����¼ A, ����ҽ������ B, סԺ���ü�¼ C, ���˷������� D
        Where (a.Id = ҽ��id_In Or a.���id = ҽ��id_In) And a.Id = b.ҽ��id And b.ҽ��id = c.ҽ����� And c.Id = d.����id And
              c.��¼״̬ In (0, 1, 3) And d.״̬ = 0;
      
        If v_Count > 0 Then
          v_Error := 'ҽ��"' || ҽ������_In || '"����δ��˵��������룬��ȡ�����������������ٻ��˷��͡�';
          Raise Err_Custom;
        End If;
      
        --���ҽ���Ƿ������Ч��ҽ������ 
        Select Count(*)
        Into v_Count
        From ����ҽ������ A, סԺ���ü�¼ B
        Where a.ҽ��id = b.ҽ����� And a.No = b.No And b.��¼״̬ = 1 And b.ʵ�ս�� <> 0 And a.���ͺ� = r_Rolladvice.���ͺ� And
              a.ҽ��id In (Select Column_Value From Table(t_Adviceids));
        If v_Count > 0 Then
          v_Error := '��ҽ���»����ڸ�����Ŀ�����ȳ�����';
          Raise Err_Custom;
        End If;
      
        --���Ʒ����Զ�ִ��ʱ������Ҳ�Զ�����ִ��(����ʿվ�д˹���) 
        --�Ǹ������õ�����ҽ����ͬ��ͨҽ��ִ�д��� 
        Select ҽ����Ч Into v_ҽ����Ч From ����ҽ����¼ Where ID = ҽ��id_In;
        If Substr(zl_GetSysParameter('����ִ���Զ����', 1254), v_ҽ����Ч + 1, 1) = '1' Then
          Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into Intdigit From Dual;
        
          For r_Rollsend In c_Rollsend(r_Rolladvice.���ͺ�) Loop
            If Nvl(r_Rollsend.ִ��״̬, 0) = 1 And
               (Nvl(r_Rollsend.ִ�п���id, 0) = Nvl(r_Rollsend.���˲���id, 0) Or
                Nvl(r_Rollsend.ִ�п���id, 0) = Nvl(r_Rollsend.���˿���id, 0)) Then
            
              --ҽ����ִ��״̬ 
              Update ����ҽ������ Set ִ��״̬ = 0 Where ���ͺ� = r_Rollsend.���ͺ� And ҽ��id = r_Rollsend.ҽ��id;
              v_Update := 1;
            
              If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 0 Then
                --���õ�ִ��״̬ 
                For r_Rollmoney In c_Rollmoneyin(r_Rollsend.���ͺ�, r_Rollsend.ҽ��id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.��¼״̬ <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1) And
                       Not r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --��ͨ����ֱ��ȡ��ִ��״̬������ҩƷ�͸������õ����� 
                      Update סԺ���ü�¼
                      Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
                      Where NO = r_Rollmoney.No And ��¼���� = r_Rolladvice.��¼���� And ��¼״̬ = r_Rollmoney.��¼״̬ And
                            Nvl(�۸񸸺�, ���) = r_Rollmoney.��� And ҽ����� = r_Rollsend.ҽ��id;
                    Elsif r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1 Then
                      --�������õ����ģ����Զ����� 
                      For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                        Zl_�����շ���¼_��������(r_Stuff.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, 0, v_��Ա����);
                      End Loop;
                    Elsif r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --סԺ���ҷ�ҩ��ҩƷ�Զ���ҩ 
                      If r_Rollmoney.ִ�в���id = r_Rollsend.���˲���id Or r_Rollmoney.ִ�в���id = r_Rollsend.���˿���id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_ҩƷ�շ���¼_������ҩ(r_Drug.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 2);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              Else
                --סԺ���˷��÷��͵�����������������Դ����סԺ�� 
                For r_Rollmoney In c_Rollmoneyout(r_Rollsend.���ͺ�, r_Rollsend.ҽ��id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.��¼״̬ <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���, 1);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1) And
                       Not r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --��ͨ����ֱ��ȡ��ִ��״̬������ҩƷ�͸������õ����� 
                      Update ������ü�¼
                      Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
                      Where NO = r_Rollmoney.No And ��¼���� = r_Rolladvice.��¼���� And ��¼״̬ = r_Rollmoney.��¼״̬ And
                            Nvl(�۸񸸺�, ���) = r_Rollmoney.��� And ҽ����� = r_Rollsend.ҽ��id;
                    Elsif r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1 Then
                      --�������õ����ģ����Զ����� 
                      For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                        Zl_�����շ���¼_��������(r_Stuff.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, 0, v_��Ա����);
                      End Loop;
                    Elsif r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --�����ҷ�ҩ��ҩƷ�Զ���ҩ 
                      If r_Rollmoney.ִ�в���id = r_Rollsend.���˲���id Or r_Rollmoney.ִ�в���id = r_Rollsend.���˿���id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_ҩƷ�շ���¼_������ҩ(r_Drug.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 1);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              End If;
            End If;
          End Loop;
        End If;
        ------------------------------------------------------------------ 
        --�������ջصĳ���ҩƷҽ�����������(���˷��þͶ�����) 
        If Nvl(r_Rolladvice.ҽ����Ч, 0) = 0 Then
          If r_Rolladvice.�ϴ�ִ��ʱ�� Is Not Null And r_Rolladvice.ĩ��ʱ�� Is Not Null Then
            If r_Rolladvice.�ϴ�ִ��ʱ�� < r_Rolladvice.ĩ��ʱ�� Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '������ڷ��͵������ѱ��ջأ������ٻ��ˡ�';
              Raise Err_Custom;
            End If;
          Elsif r_Rolladvice.�ϴ�ִ��ʱ�� Is Null And r_Rolladvice.ĩ��ʱ�� Is Not Null Then
            --�������ܱ�ȫ�������ջ� 
            v_Error := Nvl(ҽ������_In, '��ҽ��') || 'δ�����ͣ����͵������ѱ�ȫ�������ջأ������ٻ��ˡ�';
            Raise Err_Custom;
          End If;
        End If;
      
        If Nvl(r_Rolladvice.ִ��״̬, 0) In (1, 3) And v_Update <> 1 Then
          --1-��ȫִ��;3-����ִ�� 
          v_Error := Nvl(ҽ������_In, '��ҽ��') || '������͵������Ѿ�ִ�л�����ִ�У����ܻ��ˡ�';
          Raise Err_Custom;
        Else
          --������ҽ����ִ�У���ҲҪ���ƻ��ˣ����磺����Ĳɼ���ʽ�� 
          Select /*+ Rule*/
           Count(1)
          Into v_Count
          From ����ҽ������
          Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And ִ��״̬ In (1, 3) And
                ���ͺ� =
                (Select Max(���ͺ�) From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(t_Adviceids)));
          If v_Count > 0 Then
            v_Error := Nvl(ҽ������_In, '��ҽ��') || '������͵������Ѿ�ִ�л�����ִ�У����ܻ��ˡ�';
            Raise Err_Custom;
          End If;
        End If;
      
        ------------------------------------------------------------------ 
        --������ҽ���ķ�������(��һ��ҽ�������в�ͬNO����) 
        --���ԭʼ�����ѱ�����(�򲿷�����),���ù��������ж� 
        v_����no   := Null;
        v_������� := Null;
        If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 0 Then
          For r_Rollmoney In c_Rollmoneyin(r_Rolladvice.���ͺ�, Null, t_Adviceids) Loop
            --��Ӧ�ķ�����ִ�� 
            If Nvl(r_Rollmoney.ִ��״̬, 0) <> 0 Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '���͵ķ��õ���"' || r_Rollmoney.No || '"�е������ѱ����ֻ���ȫִ�У����ܻ��ˡ�';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.��¼״̬ <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���);
            End If;
            If n_Blndo > 0 Then
              --���ֽ������жϲ�����ҩ 
              If v_����no <> r_Rollmoney.No And v_������� Is Not Null Then
                Zl_סԺ���ʼ�¼_Delete(v_����no, Substr(v_�������, 2), v_��Ա���, v_��Ա����, 2, 0, 0);
                v_������� := Null;
              End If;
              v_����no   := r_Rollmoney.No;
              v_������� := v_������� || ',' || r_Rollmoney.���;
            End If;
          End Loop;
        Else
          For r_Rollmoney In c_Rollmoneyout(r_Rolladvice.���ͺ�, Null, t_Adviceids) Loop
            --��Ӧ�ķ�����ִ�� 
            If Nvl(r_Rollmoney.ִ��״̬, 0) <> 0 And Not (Nvl(r_Rollmoney.ִ��״̬, 0) = -1 And Nvl(r_Rollmoney.��¼״̬, 0) = 0) Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '���͵ķ��õ���"' || r_Rollmoney.No || '"�е������ѱ����ֻ���ȫִ�У����ܻ��ˡ�';
              Raise Err_Custom;
            End If;
            --�շѵ������շ� 
            If r_Rollmoney.��¼״̬ = 1 And Not (r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 1) Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '���͵����ﵥ��"' || r_Rollmoney.No || '"���շѣ����ܻ��ˡ�';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.��¼״̬ <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���, 1);
            End If;
            If n_Blndo > 0 Then
              --���ֽ������жϲ�����ҩ 
              If v_����no <> r_Rollmoney.No And v_������� Is Not Null Then
                If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 1 Then
                  --סԺ����Ϊ�������(���������ҽ������Ϊ������ʣ�����ҽ��û�л��˹���) 
                  Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����);
                Else
                  Zl_���ﻮ�ۼ�¼_Delete(v_����no, Substr(v_�������, 2));
                End If;
                v_������� := Null;
              End If;
              v_����no   := r_Rollmoney.No;
              v_������� := v_������� || ',' || r_Rollmoney.���;
            End If;
          End Loop;
        End If;
        If v_������� Is Not Null And v_����no Is Not Null Then
          v_������� := Substr(v_�������, 2);
          If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 0 Then
            Zl_סԺ���ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����, 2, 0, 0);
          Elsif r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 1 Then
            --סԺ����Ϊ�������(���������ҽ������Ϊ������ʣ�����ҽ��û�л��˹���) 
            Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����);
          Else
            Zl_���ﻮ�ۼ�¼_Delete(v_����no, v_�������);
          End If;
        End If;
      
        --ͨ�����ͼ�¼������Ϣ�����������ҽ��Ҫ��ǰ      
        For R In (Select a.����id, a.��ҳid, b.No, b.���ͺ�, b.��������, b.�״�ʱ��, b.ĩ��ʱ��, b.��������, a.Id, a.�������, c.��������, a.ִ�п���id
                  From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C
                  Where a.Id = b.ҽ��id And a.������Ŀid = c.Id And b.���ͺ� = r_Rolladvice.���ͺ� And
                        b.ҽ��id In (Select Column_Value From Table(t_Adviceids))) Loop
          If r.������� = 'E' And r.�������� = '6' Then
            --����
            b_Message.Zlhis_Cis_036(r.����id, r.��ҳid, Null, r.���ͺ�, r.Id, r.No, 2);
          End If;
          Select Count(1) Into v_Count From ��������˵�� A Where a.����id = r.ִ�п���id And a.�������� = '����';
          If v_Count > 0 Then
            --����ִ��ҽ�����˷���
            b_Message.Zlhis_Cis_044(r.����id, r.��ҳid, r.���ͺ�, r.Id, r.No, r.��������, r.�״�ʱ��, r.ĩ��ʱ��, r.��������);
          End If;
        End Loop;
      
        --��Ѫҽ����ɾ������ҽ������ 
        Delete From ����ҽ������ Where ���ͺ� = r_Rolladvice.���ͺ� And ҽ��id = ҽ��id_In;
      
        --ɾ��ҽ��ִ��ʱ�� (����ҽ��ID�Ų����˼�¼) 
        Delete From ҽ��ִ��ʱ�� Where ���ͺ� = r_Rolladvice.���ͺ� And ҽ��id = ҽ��id_In;
      
        --ɾ�����ͼ�¼(����ҽ����) 
        Delete /*+ Rule*/
        From ����ҽ������
        Where ���ͺ� = r_Rolladvice.���ͺ� And ҽ��id In (Select Column_Value From Table(t_Adviceids));
      
        --���(����ҽ��)�ϴ�ִ��ʱ��(���ϴη��͵�ĩ��ִ��ʱ��) 
        --���г���(���������Գ���)����ʱ����д��ĩ��ʱ�� 
        --��������û�У���ֻ���ܷ�����һ�Ρ� 
        v_ĩ��ʱ�� := Null;
        Begin
          --һ��ҽ���ķ�����ĩʱ����ͬ,һ����ҩ��ȡ��С�� 
          --ȡ���IDΪNULL��ҽ���ķ��ͼ�¼��ʱ�� 
          --����ҩ;������ҩ�÷�����δ��д���ͼ�¼ 
          Select /*+ Rule*/
           ĩ��ʱ��
          Into v_ĩ��ʱ��
          From ����ҽ������
          Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And
                ���ͺ� =
                (Select Max(���ͺ�) From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(t_Adviceids))) And
                Rownum = 1;
        Exception
          When Others Then
            Null;
        End;
        Update ����ҽ����¼ Set �ϴ�ִ��ʱ�� = v_ĩ��ʱ�� Where ID = ҽ��id_In Or ���id = ҽ��id_In;
      
        --������������ʱ��ͬʱ�Զ�����ֹͣ 
        If Nvl(r_Rolladvice.ҽ����Ч, 0) = 1 Then
          --ɾ��(��������)�����ֹͣ״̬������¼ 
          Delete /*+ Rule*/
          From ����ҽ��״̬
          Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And
                ����ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = ҽ��id_In) And �������� = 8;
          --r_RollAdvice.����ʱ��:����ʱ����ܲ����Զ�ֹͣʱ����ͬ�� 
        
          --ȡɾ����Ӧ�ָ���ҽ��״̬ 
          Select ��������
          Into v_ҽ��״̬
          From ����ҽ��״̬
          Where ����ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = ҽ��id_In) And ҽ��id = ҽ��id_In;
        
          --�ָ�(����ҽ��)���˺��״̬ 
          Update ����ҽ����¼
          Set ҽ��״̬ = v_ҽ��״̬, ִ����ֹʱ�� = Null, ͣ��ҽ�� = Null, ͣ��ʱ�� = Null
          Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        End If;
      
        --סԺ����ҽ�����ͺ�Ļ���(3-ת��;5-��Ժ;6-תԺ,11-����) 
        If r_Rolladvice.��� = 'Z' And Instr(',3,5,6,11,', ',' || r_Rolladvice.���� || ',') > 0 And
           Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
          Open c_Patilog(r_Rolladvice.����id, r_Rolladvice.��ҳid);
          Fetch c_Patilog
            Into r_Patilog;
          If c_Patilog%Found Then
            If r_Rolladvice.���� = '3' And r_Patilog.��ʼԭ�� = 3 Then
              --ȡ������ת��״̬ 
              If r_Patilog.��ʼʱ�� Is Null Then
                --ת��ҽ�������⴦����һ������������ת��ҽ��ʱ��ֻ�ܻ��������һ��,70443 
                Select Count(1)
                Into v_Count
                From ����ҽ����¼ A, ������ĿĿ¼ B
                Where a.������Ŀid = b.Id And a.����id = r_Rolladvice.����id And a.��ҳid = r_Rolladvice.��ҳid And a.������� = 'Z' And
                      b.�������� = '3' And a.ҽ��״̬ = 8 And
                      a.��ʼִ��ʱ�� > (Select ��ʼִ��ʱ�� From ����ҽ����¼ Where ID = ҽ��id_In);
                If v_Count = 0 Then
                  Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, Null, Null, Null, 'ת��');
                Else
                  v_Error := '����ת���Ѿ���ƣ������ٻ��ˡ�';
                  Raise Err_Custom;
                End If;
              Else
                v_Error := '����ת���Ѿ���ƣ������ٻ��ˡ�';
                Raise Err_Custom;
              End If;
            Elsif r_Rolladvice.���� In ('5', '6', '11') And r_Patilog.��ʼԭ�� = 10 Then
              --ȡ������Ԥ��Ժ״̬ 
              Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, Null, Null, Null, 'Ԥ��Ժ');
            End If;
          End If;
          Close c_Patilog;
        End If;
      
        --���˲���ʱ�� 
        --1.�����¼�(ֻ��һ��ҽ����¼)��������7-����,8-����,11-���� 
        If r_Rolladvice.��� = 'F' Or r_Rolladvice.��� = 'Z' And Instr(',7,8,11,', ',' || r_Rolladvice.���� || ',') > 0 Then
          Zl_���Ӳ���ʱ��_Delete(r_Rolladvice.����id, r_Rolladvice.��ҳid, 'ҽ��', r_Rolladvice.��������id, ҽ��id_In);
        End If;
      
        --2.���⴦��֪��ͬ����(������ص�֪��ͬ�����ٴε��ã���Ϊ����������������Ŀ�����й�����֪��ͬ����) 
        If Instr('C,D,E,F,G,K,L', r_Rolladvice.���) > 0 Then
          For R In (Select a.Id, a.������� From ����ҽ����¼ A Where a.Id = ҽ��id_In Or a.���id = ҽ��id_In) Loop
            --���id��һ��ҽ����һ����������ģ�����Ҫ���ж�һ����� 
            If Instr('C,D,E,F,G,K,L', r.�������) > 0 Then
              Zl_���Ӳ���ʱ��_Delete(r_Rolladvice.����id, r_Rolladvice.��ҳid, 'ҽ��', r_Rolladvice.��������id, r.Id);
            End If;
          End Loop;
        End If;
      
        --�˴������Ϣ����
        If r_Rolladvice.��� = 'D' And r_Rolladvice.���id Is Null Then
          --��� 
          b_Message.Zlhis_Cis_037(r_Rolladvice.����id, r_Rolladvice.��ҳid, Null, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No, 2);
        Elsif r_Rolladvice.��� = 'F' And r_Rolladvice.���id Is Null Then
          --���� 
          b_Message.Zlhis_Cis_038(r_Rolladvice.����id, r_Rolladvice.��ҳid, Null, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'K' And r_Rolladvice.���id Is Null Then
          --��Ѫ 
          b_Message.Zlhis_Cis_039(r_Rolladvice.����id, r_Rolladvice.��ҳid, Null, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'Z' And r_Rolladvice.�������� = '6' Then
          --���� 
          b_Message.Zlhis_Cis_040(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'Z' And r_Rolladvice.�������� = '8' Then
          --���� 
          b_Message.Zlhis_Cis_041(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'Z' And r_Rolladvice.�������� = '11' Then
          --���� 
          b_Message.Zlhis_Cis_042(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'E' And r_Rolladvice.�������� = '5' Then
          --�������� 
          b_Message.Zlhis_Cis_043(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'H' And Nvl(r_Rolladvice.��������, '0') = '0' Then
          --������ 
          b_Message.Zlhis_Cis_007(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        End If;
      End If;
    End If;
    Exit When r_Rolladvice.���ͺ� = 0;
  End Loop;
  Close c_Rolladvice;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_����;
/

--130541:������,2018-08-23,����ƽ̨��Ϣ����
Create Or Replace Procedure Zl_����ҽ����¼_ȷ��ֹͣ
(
  --���ܣ�ȷ��ָֹͣ����ҽ�� 
  --˵����һ����ҩ��ֻ�ܵ���һ�� 
  --������ҽ��ID=���IDΪNULL��ҽ����ID(��ҩ;��,��ҩ�÷�,�����Ŀ,��Ҫ����,������ҽ��) 
  ҽ��id_In           In ����ҽ����¼.Id%Type,
  ȷ��ʱ��_In         In ����ҽ����¼.ȷ��ͣ��ʱ��%Type,
  ����Ա����_In       In ��Ա��.����%Type := Null,
  �Զ�ȷ�ϻ���ȼ�_In In Number := 0
) Is
  v_״̬     ����ҽ����¼.ҽ��״̬%Type;
  v_ҽ������ ����ҽ����¼.ҽ������%Type;
  n_����id   ����ҽ����¼.����id%Type;
  n_��ҳid   ����ҽ����¼.��ҳid%Type;

  v_Temp     Varchar2(255);
  v_��Ա���� ����ҽ��״̬.������Ա%Type;
  n_Count    Number;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --���ҽ��״̬�Ƿ���ȷ:�������� 
  Select ҽ��״̬, ҽ������, ����id, ��ҳid
  Into v_״̬, v_ҽ������, n_����id, n_��ҳid
  From ����ҽ����¼
  Where ID = ҽ��id_In;
  If v_״̬ <> 8 Then
    v_Error := 'ҽ��"' || v_ҽ������ || '"��ǰ������ֹͣ״̬��';
    Raise Err_Custom;
  End If;
  --����Ƿ�����Һ��Һ��¼�����Ƿ��Ѿ����� 
  Select Count(1)
  Into n_Count
  From ��Һ��ҩ��¼ A, ����ҽ����¼ B
  Where a.ҽ��id = b.Id And ҽ��id = ҽ��id_In And a.ִ��ʱ�� > b.ִ����ֹʱ�� And a.�Ƿ����� = 1;
  If n_Count > 0 Then
    v_Error := 'ҽ��"' || v_ҽ������ || '"����ҺҩƷ���Ѿ�����Һ������������������ȷ��ֹͣ��';
    Raise Err_Custom;
  End If;

  --��ǰ������Ա 
  If ����Ա����_In Is Not Null Then
    v_��Ա���� := ����Ա����_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  Update ����ҽ����¼
  Set ҽ��״̬ = 9, ȷ��ͣ��ʱ�� = ȷ��ʱ��_In, ȷ��ͣ����ʿ = v_��Ա����
  Where ID = ҽ��id_In Or ���id = ҽ��id_In;

  Insert Into ����ҽ��״̬
    (ҽ��id, ��������, ������Ա, ����ʱ��)
    Select ID, 9, v_��Ա����, Sysdate + �Զ�ȷ�ϻ���ȼ�_In / 24 / 60 / 60
    From ����ҽ����¼
    Where ID = ҽ��id_In Or ���id = ҽ��id_In;

  b_Message.Zlhis_Cis_059(n_����id, n_��ҳid, ҽ��id_In);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_ȷ��ֹͣ;
/

--130471:��ΰ��,2018-08-22,ԤԼ����
CREATE OR REPLACE Procedure Zl_��Ժ������ҳ_Delete
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

  n_�������� ������ҳ.��������%Type;
  n_��ҳid   ������ҳ.��ҳid%Type;

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

  b_Message.Zlhis_Patient_006(����id_In, ��ҳid_In, '��Ժ�Ǽ�');

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
                ��ҳid = (Select Max(��ҳid) From ������ҳ Where ����id = ����id_In And ��ҳid < ��ҳid_In);
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

--125796:����,2018-08-22,���ӳ�����¼�ֶε�Ӧ��
Create Or Replace Procedure Zl_ҩƷ����_Strike
(
  �д�_In       In Integer,
  ԭ��¼״̬_In In ҩƷ�շ���¼.��¼״̬%Type,
  No_In         In ҩƷ�շ���¼.No%Type,
  ���_In       In ҩƷ�շ���¼.���%Type,
  ҩƷid_In     In ҩƷ�շ���¼.ҩƷid%Type,
  ��������_In   In ҩƷ�շ���¼.ʵ������%Type,
  ������_In     In ҩƷ�շ���¼.������%Type,
  ��������_In   In ҩƷ�շ���¼.��������%Type,
  ������ʽ_In   In Integer := 0, --0������������ʽ��1�������������뵥�ݣ�2������Ѳ����ĳ������뵥��  
  ����ԭ��_In   In ҩƷ�շ���¼.����ԭ��%Type := Null
) Is
  Err_Isstriked Exception;
  Err_Isoutstock Exception;
  Err_Isnonum Exception;
  Err_Isbatch Exception;
  v_Druginf Varchar2(50); --ԭ���������ڷ�����ҩƷ��Ϣ 

  v_�ⷿid       ҩƷ�շ���¼.�ⷿid%Type;
  v_�Է�����id   ҩƷ�շ���¼.�Է�����id%Type;
  v_������id   ҩƷ�շ���¼.������id%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_ԭ����       ҩƷ�շ���¼.ԭ����%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_Ч��         ҩƷ�շ���¼.Ч��%Type;
  v_�ɱ���       ҩƷ�շ���¼.�ɱ���%Type;
  v_�ɱ����     ҩƷ�շ���¼.�ɱ����%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_���ۼ�       ҩƷ�շ���¼.���ۼ�%Type;
  v_���۽��     ҩƷ�շ���¼.���۽��%Type;
  v_���         ҩƷ�շ���¼.���%Type;
  v_ժҪ         ҩƷ�շ���¼.ժҪ%Type;
  v_ʣ������     ҩƷ�շ���¼.ʵ������%Type;
  v_ʣ��ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
  v_ʣ�����۽�� ҩƷ�շ���¼.���۽��%Type;
  v_���ϵ��     ҩƷ�շ���¼.���ϵ��%Type;

  v_�շ�id   ҩƷ�շ���¼.Id%Type;
  v_������   ҩƷ�շ���¼.������%Type;
  v_��׼�ĺ� ҩƷ�շ���¼.��׼�ĺ�%Type;
  v_��ҩ��ʽ ҩƷ�շ���¼.��ҩ��ʽ%Type;

  v_�Ƿ���     �շ���ĿĿ¼.�Ƿ���%Type;
  Intdigit       Number;
  n_�ϴι�Ӧ��id ҩƷ���.�ϴι�Ӧ��id%Type;
  d_�ϴ��������� ҩƷ���.�ϴ���������%Type;
  v_������������ Varchar2(4000);
Begin
  --��ȡ���С��λ�� 
  Select Nvl(����, 2) Into Intdigit From ҩƷ���ľ��� Where ���� = 0 And ��� = 1 And ���� = 4 And ��λ = 5;
  Select Nvl(�Ƿ���, 0) Into v_�Ƿ��� From �շ���ĿĿ¼ Where Id = ҩƷid_In;
  Select Zl_Getsysparameter('������������', 1305) Into v_������������ From Dual;

  If ������ʽ_In = 1 Then
    --�����������뵥�ݣ�����д����ˡ�������ڣ������¿���¼ 
    If �д�_In = 1 Then
      Update ҩƷ�շ���¼
      Set ��¼״̬ = Decode(ԭ��¼״̬_In, 1, 3, ԭ��¼״̬_In + 3)
      Where No = No_In And ���� = 7 And ��¼״̬ = ԭ��¼״̬_In;
      If Sql%Rowcount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    Begin
      Select Distinct '(' || i.���� || ')' || Nvl(n.����, i.����) As ҩƷ��Ϣ
      Into v_Druginf
      From ҩƷ�շ���¼ a, ҩƷ��� b, �շ���ĿĿ¼ i, �շ���Ŀ���� n
      Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id And a.ҩƷid = n.�շ�ϸĿid(+) And n.����(+) = 3 And a.No = No_In And a.���� = 7 And
            Mod(a.��¼״̬, 3) = 0 And Nvl(a.����, 0) = 0 And a.ҩƷid + 0 = ҩƷid_In And
            ((Nvl(b.ҩ�����, 0) = 1 And
            a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '%ҩ��') Or (�������� Like '�Ƽ���'))) Or
            Nvl(b.ҩ������, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(ʵ������) As ʣ������, Sum(�ɱ����) As ʣ��ɱ����, Sum(���۽��) As ʣ�����۽��, �ⷿid, �Է�����id, ������id, ���ϵ��, ����, ����, ԭ����, ����, Ч��, �ɱ���,
           ����, ���ۼ�, ժҪ, ������, ��׼�ĺ�, ��ҩ��ʽ, ��ҩ��λid, ��������
    Into v_ʣ������, v_ʣ��ɱ����, v_ʣ�����۽��, v_�ⷿid, v_�Է�����id, v_������id, v_���ϵ��, v_����, v_����, v_ԭ����, v_����, v_Ч��, v_�ɱ���, v_����, v_���ۼ�,
         v_ժҪ, v_������, v_��׼�ĺ�, v_��ҩ��ʽ, n_�ϴι�Ӧ��id, d_�ϴ���������
    From ҩƷ�շ���¼
    Where No = No_In And ���� = 7 And ҩƷid = ҩƷid_In And ��� = ���_In
    Group By �ⷿid, �Է�����id, ������id, ���ϵ��, ����, ����, ԭ����, ����, Ч��, �ɱ���, ����, ���ۼ�, ժҪ, ������, ��׼�ĺ�, ��ҩ��ʽ, ��ҩ��λid, ��������;
  
    --������������ʣ�������������� 
    If v_ʣ������ < ��������_In Then
      Raise Err_Isnonum;
    End If;
  
    v_�ɱ���� := Round(��������_In / v_ʣ������ * v_ʣ��ɱ����, Intdigit);
    v_���۽�� := Round(��������_In / v_ʣ������ * v_ʣ�����۽��, Intdigit);
    v_���     := v_���۽�� - v_�ɱ����;
  
    Select ҩƷ�շ���¼_Id.Nextval Into v_�շ�id From Dual;
  
    Insert Into ҩƷ�շ���¼
      (Id, ��¼״̬, ����, No, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ԭ����, ����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ,
       ������, ��������, �����, �������, ������, ��׼�ĺ�, ��ҩ��ʽ, ��ҩ��λid, ��������, ����, ����ԭ��)
    Values
      (v_�շ�id, Decode(ԭ��¼״̬_In, 1, 2, ԭ��¼״̬_In + 2), 7, No_In, ���_In, v_�ⷿid, v_�Է�����id, v_������id, v_���ϵ��, ҩƷid_In, v_����,
       v_����, v_ԭ����, v_����, v_Ч��, -��������_In, -��������_In, v_�ɱ���, -v_�ɱ����, v_���ۼ�, -v_���۽��, -v_���, v_ժҪ, ������_In, ��������_In, Null, Null,
       v_������, v_��׼�ĺ�, v_��ҩ��ʽ, n_�ϴι�Ӧ��id, d_�ϴ���������, v_����, ����ԭ��_In);
    
    Zl_δ��ҩƷ��¼_Insert(v_�շ�id);
    
  Elsif ������ʽ_In = 2 Then
    --����Ѳ����ĳ������뵥�ݣ���д����ˡ�������ڣ����¿���¼ 
  
    --��д����ˡ�������� 
    Update ҩƷ�շ���¼
    Set ����� = ������_In, ������� = ��������_In
    Where ���� = 7 And No = No_In And ��� = ���_In And ��¼״̬ = ԭ��¼״̬_In;
  
    --��ѯ��ǰ�м�¼�Ķ�ӦID
    Select Id
    Into v_�շ�id
    From ҩƷ�շ���¼
    Where ���� = 7 And No = No_In And ��� = ���_In And ��¼״̬ = ԭ��¼״̬_In;
  
    --���¿����Ϣ ���ó����൱����� 
    Zl_ҩƷ���_Update(v_�շ�id, 3, 0);
    
    Zl_δ��ҩƷ��¼_Delete(v_�շ�id);

    --����ҩƷ���洦�� 
    If v_��ҩ��ʽ = 1 Then
      Update ҩƷ����
      Set �������� = Nvl(��������, 0) + ��������_In, ʵ������ = Nvl(ʵ������, 0) + ��������_In, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + v_���۽��
      Where �ڼ� = To_Char(Sysdate, Decode(v_������������, '1', 'yyyymm', 'yyyy')) And ����id = v_�Է�����id And �ⷿid = v_�ⷿid And
            ҩƷid = ҩƷid_In;
      --��������������0�ļ�¼ɾ���� 
      Delete From ҩƷ���� Where Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0;
    End If;
  
    --������ۺ���� 
    Zl_ҩƷ�շ���¼_��������(v_�շ�id);
  Else
    --����������ʽ������������¼����д����ˡ�������ڣ����¿���¼      
    If �д�_In = 1 Then
      Update ҩƷ�շ���¼
      Set ��¼״̬ = Decode(ԭ��¼״̬_In, 1, 3, ԭ��¼״̬_In + 3)
      Where No = No_In And ���� = 7 And ��¼״̬ = ԭ��¼״̬_In;
      If Sql%Rowcount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    Begin
      Select Distinct '(' || i.���� || ')' || Nvl(n.����, i.����) As ҩƷ��Ϣ
      Into v_Druginf
      From ҩƷ�շ���¼ a, ҩƷ��� b, �շ���ĿĿ¼ i, �շ���Ŀ���� n
      Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id And a.ҩƷid = n.�շ�ϸĿid(+) And n.����(+) = 3 And a.No = No_In And a.���� = 7 And
            Mod(a.��¼״̬, 3) = 0 And Nvl(a.����, 0) = 0 And a.ҩƷid + 0 = ҩƷid_In And
            ((Nvl(b.ҩ�����, 0) = 1 And
            a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '%ҩ��') Or (�������� Like '�Ƽ���'))) Or
            Nvl(b.ҩ������, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(ʵ������) As ʣ������, Sum(�ɱ����) As ʣ��ɱ����, Sum(���۽��) As ʣ�����۽��, �ⷿid, �Է�����id, ������id, ���ϵ��, ����, ����, ԭ����, ����, Ч��, �ɱ���,
           ����, ���ۼ�, ժҪ, ������, ��׼�ĺ�, ��ҩ��ʽ, ��ҩ��λid, ��������
    Into v_ʣ������, v_ʣ��ɱ����, v_ʣ�����۽��, v_�ⷿid, v_�Է�����id, v_������id, v_���ϵ��, v_����, v_����, v_ԭ����, v_����, v_Ч��, v_�ɱ���, v_����, v_���ۼ�,
         v_ժҪ, v_������, v_��׼�ĺ�, v_��ҩ��ʽ, n_�ϴι�Ӧ��id, d_�ϴ���������
    From ҩƷ�շ���¼
    Where No = No_In And ���� = 7 And ҩƷid = ҩƷid_In And ��� = ���_In
    Group By �ⷿid, �Է�����id, ������id, ���ϵ��, ����, ����, ԭ����, ����, Ч��, �ɱ���, ����, ���ۼ�, ժҪ, ������, ��׼�ĺ�, ��ҩ��ʽ, ��ҩ��λid, ��������;
  
    --������������ʣ�������������� 
    If v_ʣ������ < ��������_In Then
      Raise Err_Isnonum;
    End If;
  
    v_�ɱ���� := Round(��������_In / v_ʣ������ * v_ʣ��ɱ����, Intdigit);
    v_���۽�� := Round(��������_In / v_ʣ������ * v_ʣ�����۽��, Intdigit);
    v_���     := v_���۽�� - v_�ɱ����;
  
    Select ҩƷ�շ���¼_Id.Nextval Into v_�շ�id From Dual;
    Insert Into ҩƷ�շ���¼
      (Id, ��¼״̬, ����, No, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ԭ����, ����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ,
       ������, ��������, �����, �������, ������, ��׼�ĺ�, ��ҩ��ʽ, ��ҩ��λid, ��������, ����, ����ԭ��)
    Values
      (v_�շ�id, Decode(ԭ��¼״̬_In, 1, 2, ԭ��¼״̬_In + 2), 7, No_In, ���_In, v_�ⷿid, v_�Է�����id, v_������id, v_���ϵ��, ҩƷid_In, v_����,
       v_����, v_ԭ����, v_����, v_Ч��, -��������_In, -��������_In, v_�ɱ���, -v_�ɱ����, v_���ۼ�, -v_���۽��, -v_���, v_ժҪ, ������_In, ��������_In, ������_In,
       ��������_In, v_������, v_��׼�ĺ�, v_��ҩ��ʽ, n_�ϴι�Ӧ��id, d_�ϴ���������, v_����, ����ԭ��_In);
    
    --���¿����Ϣ ���ó����൱����� 
    Zl_ҩƷ���_Update(v_�շ�id, 3, 0);
    
    --����ҩƷ���洦�� 
    If v_��ҩ��ʽ = 1 Then
      Update ҩƷ����
      Set �������� = Nvl(��������, 0) + ��������_In, ʵ������ = Nvl(ʵ������, 0) + ��������_In, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + v_���۽��
      Where �ڼ� = To_Char(Sysdate, Decode(v_������������, '1', 'yyyymm', 'yyyy')) And ����id = v_�Է�����id And �ⷿid = v_�ⷿid And
            ҩƷid = ҩƷid_In;
      --��������������0�ļ�¼ɾ���� 
      Delete From ҩƷ���� Where Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0;
    End If;
  
    --������ۺ���� 
    Zl_ҩƷ�շ���¼_��������(v_�շ�id);
  End If;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102,
                            '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ[' || v_Druginf || ']�����ܳ�����[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_ҩƷ����_Strike;
/

--119442:��˼��,2018-08-22,ǩ����������ʱ��
CREATE OR REPLACE Procedure ZL_Ӱ�񱨸�ǩ��_����(
	�����ļ�ID_In  In   ���Ӳ�������.�ļ�ID%Type,
    ��ʼ��_In  In       ���Ӳ�������.��ʼ��%Type,
    ��ֹ��_In  In       ���Ӳ�������.��ֹ��%Type,
    ��������_In In      ���Ӳ�������.��������%Type,
    ����_In In          ���Ӳ�������.�����ı�%Type,
    ǰ������_In In      ���Ӳ�������.Ҫ������%Type,
    ʱ���_In  In       ���Ӳ�������.Ҫ�ص�λ%Type,
    ǩ������_In In      ���Ӳ�������.Ҫ�ر�ʾ%Type,
    ǩ����Ϣ_In In      ���Ӳ�������.Ҫ��ֵ��%Type
) Is
	 n_Nextid     ���Ӳ�������.Id%Type;
     n_���       ���Ӳ�������.�������%Type;
     n_������   ���Ӳ�������.������%Type;
Begin
     Select max(�������) +1 Into n_��� From ���Ӳ������� Where �ļ�ID = �����ļ�ID_In;
     Select nvl(Max(������),0)+1 Into n_������ From ���Ӳ������� Where �ļ�ID = �����ļ�ID_In And ��������=8;

     Select ���Ӳ�������_Id.Nextval Into n_Nextid From Dual;
     Insert Into ���Ӳ�������(ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������,
            �����д�, �����ı�, �Ƿ���, �������id, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id, �滻��, Ҫ������,
            Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��)
          Values (n_Nextid ,�����ļ�ID_In,��ʼ��_In,��ֹ��_In,Null,n_���,8,n_������,1,��������_In,Null,����_In,
                 0,Null,Null,Null,Null,Null,Null,ǰ������_In,1,50,0,ʱ���_In,ǩ������_In,0,ǩ����Ϣ_In);
     If n_������=1 Then
        Update ���Ӳ�����¼ Set ���ʱ�� = Sysdate ,ǩ������ = ǩ������_In Where id = �����ļ�ID_In;
     Else
        Update ���Ӳ�����¼ Set ǩ������ = ǩ������_In Where id = �����ļ�ID_In;
     End If;
	 Update ���Ӳ�����¼ Set ���ʱ�� = Sysdate, ǩ������ = ǩ������_In Where ID = �����ļ�id_In;
Exception
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End ZL_Ӱ�񱨸�ǩ��_����;
/

--130471:��ΰ��,2018-08-22,ԤԼ����
CREATE OR REPLACE Procedure Zl_Third_Outpatireg
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ����ڲ���Ԥ��Ժ��¼/ȡ��Ԥ��Ժ    ����д��
  --��Σ�xml_in
  --<IN>
  -- <TYPE>1</TYPE>   --�������ͣ�1-����Ԥ��Ժ��¼��0-ȡ��Ԥ��Ժ
  -- <GHID>1162695</GHID>       --�Һ�id
  -- <RYKSID>202704</RYKSID>    --��Ժ����ID
  -- <RYBQID>202704</RYBQID>    --��Ժ����ID
  -- <CH>5</CH>   --����
  -- <YZID>3</YZID> --ҽ��id
  -- <CZYBH></CZYBH> --����Ա���
  -- <CZYXM></CZYXM> --����Ա����
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

  n_ҽ��id ����ҽ����¼.Id%Type;
  Cursor c_Advice Is
    Select Nvl(a.���id, a.Id) As ��id, a.���id, a.���, a.����id, a.�Һŵ�, a.Ӥ��, a.����, c.��������, a.�������, a.ҽ��״̬, a.ҽ������, a.����ҽ��,
           a.��ʼִ��ʱ��, a.ִ��ʱ�䷽��, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, Nvl(a.������־, 0) As ������־, a.������Ŀid, a.�շ�ϸĿid
    From ����ҽ����¼ A, ������ĿĿ¼ C
    Where a.������Ŀid = c.Id And a.������� = 'Z' And c.�������� = '2' And a.Id = n_ҽ��id;
  r_Advice c_Advice%RowType;

  Cursor c_Pati(v_����id ������Ϣ.����id%Type) Is
    Select a.����id, a.סԺ��, a.����, a.�Ա�, a.����, a.�ѱ�, a.��������, a.����, a.����, a.ѧ��, a.����״��, a.ְҵ, a.���, a.���֤��, a.�����ص�, a.��ͥ��ַ,
           a.��ͥ��ַ�ʱ�, a.��ͥ�绰, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.������λ, a.��ͬ��λid, a.��λ�绰, a.��λ�ʱ�,
           a.��λ������, a.��λ�ʺ�, a.������, a.������, a.��������, a.����, a.����, a.ҽ�Ƹ��ʽ, a.����
    From ������Ϣ A
    Where a.����id = v_����id;
  r_Pati c_Pati%RowType;

  n_Type   Number;
  n_�Һ�id ����ҽ����¼.Id%Type;
  n_����id ����ҽ����¼.Id%Type;
  n_����id ����ҽ����¼.Id%Type;
  v_����   ������ҳ.��Ժ����%Type;

  n_����id ������ҳ.����id%Type;
  v_No     ���˹Һż�¼.No%Type;
  n_Count  Number;

  v_��Ժ��ʽ ������ҳ.��Ժ��ʽ%Type;
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;
  v_Temp     Varchar2(4000);
  v_Error    Varchar2(2000);

  Err_Custom Exception;
Begin
  Select Extractvalue(Value(A), 'IN/TYPE'), Extractvalue(Value(A), 'IN/GHID') As �Һ�id,
         Extractvalue(Value(A), 'IN/RYKSID') As ��Ժ����id, Extractvalue(Value(A), 'IN/RYBQID') As ��Ժ����id,
         Extractvalue(Value(A), 'IN/CH') As ����, Extractvalue(Value(A), 'IN/CZYBH') As ���,
         Extractvalue(Value(A), 'IN/CZYXM') As ����, Extractvalue(Value(A), 'IN/YZID') As ҽ��id
  Into n_Type, n_�Һ�id, n_����id, n_����id, v_����, v_��Ա���, v_��Ա����, n_ҽ��id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If n_Type = 1 Then
    --סԺԤԼ�Ǽ�
    Select a.����id, a.No, Decode(a.����, 1, '����', Null)
    Into n_����id, v_No, v_��Ժ��ʽ
    From ���˹Һż�¼ A
    Where a.Id = n_�Һ�id;
  
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
  
    If r_Advice.������־ = 1 Then
      v_��Ժ��ʽ := '����';
    End If;
  
    Open c_Pati(n_����id);
    Fetch c_Pati
      Into r_Pati;
  
    --��ǰ������Ա
    If v_��Ա��� Is Null Or v_��Ա���� Is Null Then
      v_Temp     := Zl_Identity;
      v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
      v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
      v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    End If;
  
    --ɾ�����ۼ�¼��סԺԤԼ��¼���ܲ���
    Begin
      Select Count(1) Into n_Count From ������ҳ Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = 0;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) > 0 Then
      Zl_��Ժ������ҳ_Delete(r_Advice.����id, 0, 0, 0);
      n_Count := 0;
    End If;
  
    If n_Count = 0 Then
      Select Count(1) Into n_Count From ������ҳ Where ����id = r_Advice.����id And ��Ժ���� Is Null;
    End If;
    If n_Count = 0 Then
      Select Count(1)
      Into n_Count
      From ������ҳ
      Where ����id = r_Advice.����id And (��Ժ���� >= r_Advice.��ʼִ��ʱ�� Or ��Ժ���� >= r_Advice.��ʼִ��ʱ��);
    End If;
  
    If n_Count = 0 Then
      Zl_��Ժ������ҳ_Insert(1, 0, r_Pati.����id, r_Pati.סԺ��, Null, r_Pati.����, r_Pati.�Ա�, r_Pati.����, r_Pati.�ѱ�, r_Pati.��������,
                       r_Pati.����, r_Pati.����, r_Pati.ѧ��, r_Pati.����״��, r_Pati.ְҵ, r_Pati.���, r_Pati.���֤��, r_Pati.�����ص�,
                       r_Pati.��ͥ��ַ, r_Pati.��ͥ��ַ�ʱ�, r_Pati.��ͥ�绰, r_Pati.���ڵ�ַ, r_Pati.���ڵ�ַ�ʱ�, r_Pati.��ϵ������, r_Pati.��ϵ�˹�ϵ,
                       r_Pati.��ϵ�˵�ַ, r_Pati.��ϵ�˵绰, r_Pati.������λ, r_Pati.��ͬ��λid, r_Pati.��λ�绰, r_Pati.��λ�ʱ�, r_Pati.��λ������,
                       r_Pati.��λ�ʺ�, r_Pati.������, r_Pati.������, r_Pati.��������, n_����id, Null, Null, v_��Ժ��ʽ, Null, Null,
                       r_Advice.����ҽ��, r_Pati.����, r_Pati.����, r_Advice.��ʼִ��ʱ��, Null, Null, r_Pati.ҽ�Ƹ��ʽ, Null, Null, Null,
                       Null, Null, Null, r_Pati.����, v_��Ա���, v_��Ա����, 0, Null, n_����id, 0, Null, Null, Null, Null, Null,
                       Null, Null, n_�Һ�id);
    End If;
  Else
    --ȡ���Ǽ�
    Select b.����id Into n_����id From ������ҳ B Where b.�Һ�id = n_�Һ�id;
    Zl_��Ժ������ҳ_Delete(n_����id, 0);
  End If;
  Xml_Out := Xmltype('<OUTPUT><RESULT>true</RESULT></OUTPUT>');
Exception
  When Err_Custom Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || v_Error || '</MSG></ERROR></OUTPUT>');
  When Others Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || SQLCode || '***' || SQLErrM ||
                       '</MSG></ERROR></OUTPUT>');
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Outpatireg;
/

--130471:��ΰ��,2018-08-21,ԤԼ���Ų�ѯ
Create Or Replace Procedure Zl_Third_Patiinfo_Update
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ�����ԤԼ���ĸ��²�����Ϣ
  --��Σ�xml_in
  --<IN>
  -- <REGID>1162695</REGID>   --�Һ�id
  -- <PATIID>5</PATIID>     --����ID
  -- <HOME_TEL>3</HOME_TEL>   --��ͥ�绰
  -- <CONTACT_TEL></CONTACT_TEL> --��ϵ�˵绰
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
  n_�Һ�id     ���˹Һż�¼.Id%Type;
  n_����id     ������Ϣ.����id%Type;
  v_��ͥ�绰   ������Ϣ.��ͥ�绰%Type;
  v_��ϵ�˵绰 ������Ϣ.��ϵ�˵绰%Type;

  v_Error Varchar2(2000);

  Err_Custom Exception;
Begin
  Select Extractvalue(Value(A), 'IN/REGID') As �Һ�id, Extractvalue(Value(A), 'IN/PATIID') As ����id,
         Extractvalue(Value(A), 'IN/HOME_TEL') As ��ͥ�绰, Extractvalue(Value(A), 'IN/CONTACT_TEL') As ��ϵ�˵绰
  Into n_�Һ�id, n_����id, v_��ͥ�绰, v_��ϵ�˵绰
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  If n_����id = 0 Then
    v_Error := '����ID������Ϊ��!';
    Raise Err_Custom;
  End If;
  Update ������Ϣ Set ��ͥ�绰 = v_��ͥ�绰, ��ϵ�˵绰 = v_��ϵ�˵绰 Where ����id = n_����id;
  Xml_Out := Xmltype('<OUTPUT><RESULT>true</RESULT></OUTPUT>');
Exception
  When Err_Custom Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || v_Error || '</MSG></ERROR></OUTPUT>');
  When Others Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || SQLCode || '***' || SQLErrM ||
                       '</MSG></ERROR></OUTPUT>');
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Patiinfo_Update;
/

--130469:����,2018-08-17,��Ѫ�������ê��(Zlhis_Blood_010)
Create Or Replace Procedure Zl_ҽ����˹���_Audit
(
  ҽ��id_In   ����ҽ��״̬.ҽ��id%Type,
  ���_In     Number,
  ������Ա_In ����ҽ��״̬.������Ա%Type,
  ����ʱ��_In ����ҽ��״̬.����ʱ��%Type,
  ����˵��_In ����ҽ��״̬.����˵��%Type := Null,
  ��˶���_In Number := 1 --1=����ҽ����2=��Ѫҽ�� 
) Is
  Err_Item Exception;
  v_Err_Msg  Varchar2(200);
  n_���״̬ Number;
  n_Count    Number;
  n_���     Number;
Begin
  Select Nvl(Max(���״̬), 0), Count(1) Into n_���״̬, n_Count From ����ҽ����¼ Where Id = ҽ��id_In;
  If n_���״̬ Not In (1, 7) And n_Count <> 0 Then
    v_Err_Msg := '��ҽ���Ѿ���˻������,���֤��';
    Raise Err_Item;
  Elsif n_Count = 0 Then
    v_Err_Msg := '��ҽ���Ѿ�ɾ��,���֤��';
    Raise Err_Item;
  End If;

  Update ����ҽ����¼ Set ���״̬ = ���_In + 1 Where Id = ҽ��id_In Or ���id = ҽ��id_In;
  If ��˶���_In = 2 And ���_In = 3 Then
    --����Ѫ��ϵͳʱ�����⴦�� 
    n_��� := 11;
  Elsif ��˶���_In = 2 And ���_In = 6 Then
    n_��� := 18;
  Else
    n_��� := ���_In + 10;
  End If;
  Insert Into ����ҽ��״̬
    (ҽ��id, ��������, ������Ա, ����ʱ��, ����˵��)
    Select Id, n_���, ������Ա_In, ����ʱ��_In, ����˵��_In
    From ����ҽ����¼
    Where Id = ҽ��id_In Or ���id = ҽ��id_In;
  --��Ѫҽ�������ɣ��׳�ê��
  If ��˶���_In = 2 And n_��� = 11 Then
    EXECUTE IMMEDIATE 'b_Message_Blood.Zlhis_Blood_010(:1)' USING ҽ��id_In;
  End If;	
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_ҽ����˹���_Audit;
/




------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0026' Where ���=&n_System;
Commit;
