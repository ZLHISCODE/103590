----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--123137:����,2018-03-21,�ڱ���֧����Ŀ���������ֶα��շ��õȼ�
alter table ����֧����Ŀ add ���շ��õȼ�  varchar2(50);




------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--123386:��˶,2018-03-23,�շѼ�Ŀ���շѶ���ê��
Insert Into Zlmsg_Lists(Bz_Type, Code, Name, Key_Define, Note)
Select '�ֵ�', 'ZLHIS_DICT_053', '�շѼ�Ŀ�䶯', '<root><ID></ID><�䶯����></�䶯����></root>', '�շ�ϸĿ����:����ʱ'  From Dual Union All 
Select '�ֵ�', 'ZLHIS_DICT_054', '�����շѶ��ձ䶯', '<root><ID></ID><ԭ����></ԭ����><�ֶ���></�ֶ���></root>', '������Ŀ����:���������շѶ���ʱ'  From Dual;

--123263:������,2018-03-22,����ƽ̨��Ϣê��
Insert Into Zlmsg_Lists(Bz_Type, Code, Name, Key_Define, Note)
Select '�ֵ�', 'ZLHIS_DICTLIS_004', '�������Ʊ걾����', '<root><����></����><����></����><����><����/><�����Ա�></�����Ա�></root>', '�ֵ������:�������Ʊ걾����'  From Dual Union All 
Select '�ֵ�', 'ZLHIS_DICTLIS_005', '�޸����Ʊ걾����', '<root><����></����><����></����><����><����/><�����Ա�></�����Ա�></root>', '�ֵ������:�޸����Ʊ걾����'  From Dual Union All 
Select '�ֵ�', 'ZLHIS_DICTLIS_006', 'ɾ�����Ʊ걾����', '<root><����></����><����></����><����><����/><�����Ա�></�����Ա�></root>', '�ֵ������:ɾ�����Ʊ걾����'  From Dual Union All 
Select '�ֵ�', 'ZLHIS_DICTLIS_007', '������Ѫ��', '<root><����></����><����></����><����></����><��Ӽ�></��Ӽ�><��Ѫ��></��Ѫ��><���></���><��ɫ></��ɫ><����ID></����ID><root>', '�����Ѫ������:������Ѫ��'  From Dual Union All 
Select '�ֵ�', 'ZLHIS_DICTLIS_008', '�޸Ĳ�Ѫ��', '<root><����></����><����></����><����></����><��Ӽ�></��Ӽ�><��Ѫ��></��Ѫ��><���></���><��ɫ></��ɫ><����ID></����ID><root>', '�����Ѫ������:�޸Ĳ�Ѫ��'  From Dual Union All 
Select '�ֵ�', 'ZLHIS_DICTLIS_009', 'ɾ����Ѫ��', '<root><����></����><����></����><����></����><��Ӽ�></��Ӽ�><��Ѫ��></��Ѫ��><���></���><��ɫ></��ɫ><����ID></����ID></root>', '�����Ѫ������:ɾ����Ѫ��'  From Dual ;


--122998:������,2018-03-21,����ƽ̨��Ϣê��
Insert Into Zlmsg_Lists(Bz_Type, Code, Name, Key_Define, Note)
Select '�ֵ�', 'ZLHIS_DICTPACS_001', '�������Ƽ������', '<root><����></����><����></����><����><����/><������></������></root>', '�ֵ������:�������Ƽ������'  From Dual Union All 
Select '�ֵ�', 'ZLHIS_DICTPACS_002', '�޸����Ƽ������', '<root><����></����><����></����><����><����/><������></������></root>', '�ֵ������:�޸����Ƽ������'  From Dual Union All 
Select '�ֵ�', 'ZLHIS_DICTPACS_003', 'ɾ�����Ƽ������', '<root><����></����><����></����><����><����/><������></������></root>', '�ֵ������:ɾ�����Ƽ������'  From Dual Union All 
Select '�ֵ�', 'ZLHIS_DICTPACS_004', '�������Ƽ�鲿λ', '<root><����></����><����></����><����></����><����></����><��ע></��ע><����></����><�����Ա�></�����Ա�><root>', '��鲿λ����:�������Ƽ�鲿λ'  From Dual Union All 
Select '�ֵ�', 'ZLHIS_DICTPACS_005', '�޸����Ƽ�鲿λ', '<root><����></����><����></����><����></����><����></����><��ע></��ע><����></����><�����Ա�></�����Ա�><root>', '��鲿λ����:�޸����Ƽ�鲿λ'  From Dual Union All 
Select '�ֵ�', 'ZLHIS_DICTPACS_006', 'ɾ�����Ƽ�鲿λ', '<root><����></����><����></����><����></����><����></����><��ע></��ע><����></����><�����Ա�></�����Ա�></root>', '��鲿λ����:ɾ�����Ƽ�鲿λ'  From Dual Union All 
Select '�ֵ�', 'ZLHIS_DICTPACS_007', '����������Ŀ��λ', '<root><ID></ID><��ĿID></��ĿID><����></����><��λ></��λ><����></����><Ĭ��></Ĭ��></root>', '������Ŀ����:����������Ŀ��λ'  From Dual Union All 
Select '�ֵ�', 'ZLHIS_DICTPACS_008', '�޸�������Ŀ��λ', '<root><ID></ID><��ĿID></��ĿID><����></����><��λ></��λ><����></����><Ĭ��></Ĭ��></root>', '������Ŀ����:�޸�������Ŀ��λ'  From Dual Union All 
Select '�ֵ�', 'ZLHIS_DICTPACS_009', 'ɾ��������Ŀ��λ', '<root><ID></ID><��ĿID></��ĿID><����></����><��λ></��λ><����></����><Ĭ��></Ĭ��></root>', '������Ŀ����:ɾ��������Ŀ��λ'  From Dual;


--123312:���Ʊ�,2018-03-22,�²���ϵͳ����ƽ̨��Ϣ
Insert Into Zlmsg_Lists(Bz_Type, Code, Name, Key_Define, Note)
Select '�ٴ�', 'ZLHIS_CIS_056', '�������뷢�ͺ��޸�', '<root><����ID></����ID><��ҳID></��ҳID><�Һŵ�></�Һŵ�><���ͺ�></���ͺ�><ID></ID><������Դ></������Դ></root>', 'ҽ��վ�޸Ĳ������뵥�ı걾��Ϣʱ'  From Dual;


--123098:���ϴ�,2018-03-20,������󶨿��Ƿ��Զ����������
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1107, 0, 0, 0, 0, 0, 0, 26, '�Զ������', NULL, '1',
         '�����˴˲������ڷ�����󶨿�ʱ����Ϊû������ŵĲ����Զ����������', '0-���Զ����������,1-�Զ����������', NULL, 'Ϊ�����Զ���������ţ�����������������', Null
  From Dual;

Update zlParameters Set ����ֵ = 1 where ϵͳ = &n_System And ģ�� = 1107 And ������ = '�Զ������' And Exists(Select 1��From zlParameters where ϵͳ = &n_System And ģ��= 1111 And ������ = '�Զ������' And ����ֵ = 1);





-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--123120:��͢��,2018-03-21,����޸Ĵ��״̬����Һ����Ȩ��
Insert Into zlProgFuncs
  (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ)
  Select &n_System, 1254, '�޸���Һ���״̬', 35, '�ٴ���ʿ�д�Ȩ��ʱ�����޸���Һ��¼�Ĵ��״̬���������޸ġ�', 1
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where ϵͳ = &n_System And ��� = 1254 And ���� = '�޸���Һ���״̬');

Insert Into zlProgFuncs
  (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ)
  Select &n_System, 1254, '�޸���Һ����', 36, '�ٴ���ʿ�д�Ȩ��ʱ�����޸���Һ��¼�����Σ��������޸ġ�', 1
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where ϵͳ = &n_System And ��� = 1254 And ���� = '�޸���Һ����');

Insert Into zlRoleGrant
  Select &n_System, 1254, b.��ɫ, '�޸���Һ���״̬'
  From zlRoleGrant B
  Where b.ϵͳ = &n_System And b.��� = 1254 And b.���� = '�޸Ĵ��״̬����Һ����' And Not Exists
   (Select 1
         From zlRoleGrant C
         Where c.ϵͳ = &n_System And c.��� = 1254 And c.��ɫ = b.��ɫ And c.���� = '�޸���Һ���״̬');

Insert Into zlRoleGrant
  Select &n_System, 1254, b.��ɫ, '�޸���Һ����'
  From zlRoleGrant B
  Where b.ϵͳ = &n_System And b.��� = 1254 And b.���� = '�޸Ĵ��״̬����Һ����' And Not Exists
   (Select 1
         From zlRoleGrant C
         Where c.ϵͳ = &n_System And c.��� = 1254 And c.��ɫ = b.��ɫ And c.���� = '�޸���Һ����');
         
Delete From zlProgFuncs Where ϵͳ = &n_System And ��� = 1254 And ���� = '�޸Ĵ��״̬����Һ����';






-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--123386:��˶,2018-03-23,�շѼ�Ŀ���շѶ���ê��
Create Or Replace Procedure Zl_�����շ�_Update
(
  ������Ŀid_In In �����շѹ�ϵ.������Ŀid%Type,
  �Ƽ�����_In   ������ĿĿ¼.�Ƽ�����%Type,
  �շ�����_In   In Varchar2, --��"|"�ָ��������շ����ݣ�ÿ����¼��"������ĿID^����^�̶�^����^����^��λ^��鷽��^�շѷ�ʽ"��֯
  �Ƿ�ɾ��_In   Number := 1,
  ���ÿ���id_In �����շѹ�ϵ.���ÿ���id%Type := Null,
  ������Դ_In   �����շѹ�ϵ.������Դ%Type := 0
) Is
  v_Records    Varchar2(4000);
  v_Currrec    Varchar2(1000);
  v_Fields     Varchar2(1000);
  v_�շ���Ŀid �����շѹ�ϵ.�շ���Ŀid%Type;
  v_�շ�����   �����շѹ�ϵ.�շ�����%Type;
  v_���ж���   �����շѹ�ϵ.���ж���%Type;
  v_������Ŀ   �����շѹ�ϵ.������Ŀ%Type;
  v_��������   �����շѹ�ϵ.��������%Type;
  v_��鲿λ   �����շѹ�ϵ.��鲿λ%Type;
  v_��鷽��   �����շѹ�ϵ.��鷽��%Type;
  v_�շѷ�ʽ   �����շѹ�ϵ.�շѷ�ʽ%Type;
  v_Old        Varchar2(4000);
Begin
  Update ������ĿĿ¼ Set �Ƽ����� = �Ƽ�����_In Where ID = ������Ŀid_In;
  If �Ƿ�ɾ��_In = 1 Then
    If Nvl(���ÿ���id_In, 0) = 0 And Nvl(������Դ_In, 0) = 0 Then
      Select f_List2str(Cast(Collect(�շ���Ŀid || '^' || �շ����� || '^' || ���ж��� || '^' || ������Ŀ || '^' || �������� || '^' || ��鲿λ || '^' || ��鷽�� || '^' || �շѷ�ʽ) As
                              t_Strlist), '|')
      Into v_Old
      From �����շѹ�ϵ
      Where ������Ŀid = ������Ŀid_In;
    End If;
    Delete �����շѹ�ϵ
    Where ������Ŀid = ������Ŀid_In And Nvl(���ÿ���id, 0) = Nvl(���ÿ���id_In, 0) And ������Դ = ������Դ_In;
  End If;
  If �շ�����_In Is Null Then
    v_Records := Null;
  Else
    v_Records := �շ�����_In || '|';
  End If;
  While v_Records Is Not Null Loop
    v_Currrec    := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
    v_Fields     := v_Currrec;
    v_�շ���Ŀid := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_�շ�����   := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_���ж���   := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_������Ŀ   := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_��������   := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_��鲿λ   := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_��鷽��   := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_�շѷ�ʽ   := To_Number(v_Fields);
    Insert Into �����շѹ�ϵ
      (������Ŀid, �շ���Ŀid, �շ�����, ���ж���, ������Ŀ, ��������, ��鲿λ, ��鷽��, �շѷ�ʽ, ���ÿ���id, ������Դ)
    Values
      (������Ŀid_In, v_�շ���Ŀid, v_�շ�����, v_���ж���, v_������Ŀ, v_��������, v_��鲿λ, v_��鷽��, v_�շѷ�ʽ, ���ÿ���id_In, ������Դ_In);
    v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
  End Loop;

  If Nvl(���ÿ���id_In, 0) = 0 And Nvl(������Դ_In, 0) = 0 Then
    b_Message.Zlhis_Dict_054(������Ŀid_In, v_Old, �շ�����_In);
  End If;
  --delete �����շѹ�ϵ where ������ĿID=������ĿID_IN and �շ�����=0;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շ�_Update;
/

--123386:��˶,2018-03-23,�շѼ�Ŀ���շѶ���ê��
Create Or Replace Procedure Zl_�շѼ�Ŀ_Update
(
  �շ�ϸĿid_In In �շѼ�Ŀ.�շ�ϸĿid%Type := Null,
  ������Ŀid_In In �շѼ�Ŀ.������Ŀid%Type := Null,
  ԭ��_In       In �շѼ�Ŀ.ԭ��%Type := Null,
  �ּ�_In       In �շѼ�Ŀ.�ּ�%Type := Null,
  �����շ���_In In �շѼ�Ŀ.�����շ���%Type := Null,
  �Ӱ�Ӽ���_In In �շѼ�Ŀ.�Ӱ�Ӽ���%Type := Null,
  ����˵��_In   In �շѼ�Ŀ.����˵��%Type := Null,
  ����id_In     In �շѼ�Ŀ.����id%Type := Null,
  ������_In     In �շѼ�Ŀ.������%Type := Null,
  ȱʡ�۸�_In   In �շѼ�Ŀ.ȱʡ�۸�%Type := Null,
  �۸�ȼ�_In   In �շѼ�Ŀ.�۸�ȼ�%Type := Null
) Is
  n_State Number(1);
Begin
  Update �շѼ�Ŀ
  Set ԭ�� = ԭ��_In, �ּ� = �ּ�_In, ������Ŀid = ������Ŀid_In, �Ӱ�Ӽ��� = �Ӱ�Ӽ���_In, �����շ��� = �����շ���_In, ����˵�� = ����˵��_In, ����id = ����id_In,
      ������ = ������_In, ȱʡ�۸� = ȱʡ�۸�_In
  Where �շ�ϸĿid = �շ�ϸĿid_In And Nvl(�۸�ȼ�, '-') = Nvl(�۸�ȼ�_In, '-') And
        Decode(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'), Null, ��ֹ����) Is Null;

  If Sql%NotFound Then
    --ֻ��ʱ�۲Ż�����������
    Insert Into �շѼ�Ŀ
      (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �Ӱ�Ӽ���, �����շ���, �䶯ԭ��, ����˵��, ����id, ������, ִ������, ��ֹ����, NO, ���, ȱʡ�۸�, ���ۻ��ܺ�, �۸�ȼ�)
    Values
      (�շѼ�Ŀ_Id.Nextval, Null, �շ�ϸĿid_In, ԭ��_In, �ּ�_In, ������Ŀid_In, �Ӱ�Ӽ���_In, �����շ���_In, 1, ����˵��_In, ����id_In, ������_In,
       Sysdate, To_Date('3000-01-01', 'yyyy-mm-dd'), Nextno(9), 1, ȱʡ�۸�_In, Null, �۸�ȼ�_In);
    n_State := 1;
  Else
    n_State := 2;
  End If;

  If �۸�ȼ�_In Is Null Then
    b_Message.Zlhis_Dict_053(�շ�ϸĿid_In, n_State);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�շѼ�Ŀ_Update;
/

--123386:��˶,2018-03-23,�շѼ�Ŀ���շѶ���ê��
Create Or Replace Procedure Zl_�շѼ�Ŀ_Stop
(
  �շ�ϸĿid_In In �շѼ�Ŀ.�շ�ϸĿid%Type,
  ��ֹ����_In   In �շѼ�Ŀ.��ֹ����%Type := Null,
  �۸�ȼ�_In   In �շѼ�Ŀ.�۸�ȼ�%Type := Null
) Is
Begin
  Update �շѼ�Ŀ
  Set ��ֹ���� = ��ֹ����_In
  Where Decode(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'), Null, ��ֹ����) Is Null And �շ�ϸĿid = �շ�ϸĿid_In And
        Nvl(�۸�ȼ�, '-') = Nvl(�۸�ȼ�_In, '-');

  If �۸�ȼ�_In Is Null Then
    b_Message.Zlhis_Dict_053(�շ�ϸĿid_In, 0);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�շѼ�Ŀ_Stop;
/

--123386:��˶,2018-03-23,�շѼ�Ŀ���շѶ���ê��
Create Or Replace Procedure Zl_�շѼ�Ŀ_Insert
(
  Id_In         In �շѼ�Ŀ.Id%Type,
  ԭ��id_In     In �շѼ�Ŀ.ԭ��id%Type := Null,
  �շ�ϸĿid_In In �շѼ�Ŀ.�շ�ϸĿid%Type := Null,
  ������Ŀid_In In �շѼ�Ŀ.������Ŀid%Type := Null,
  ԭ��_In       In �շѼ�Ŀ.ԭ��%Type := Null,
  �ּ�_In       In �շѼ�Ŀ.�ּ�%Type := Null,
  �����շ���_In In �շѼ�Ŀ.�����շ���%Type := Null,
  �Ӱ�Ӽ���_In In �շѼ�Ŀ.�Ӱ�Ӽ���%Type := Null,
  ����˵��_In   In �շѼ�Ŀ.����˵��%Type := Null,
  ����id_In     In �շѼ�Ŀ.����id%Type := Null,
  ������_In     In �շѼ�Ŀ.������%Type := Null,
  ִ������_In   In �շѼ�Ŀ.ִ������%Type := Null,
  �䶯ԭ��_In   In �շѼ�Ŀ.�䶯ԭ��%Type := 1,
  No_In         In �շѼ�Ŀ.No%Type := Null,
  ���_In       In �շѼ�Ŀ.���%Type := 1,
  ȱʡ�۸�_In   In �շѼ�Ŀ.ȱʡ�۸�%Type := Null,
  ���ۻ��ܺ�_In In �շѼ�Ŀ.���ۻ��ܺ�%Type := Null,
  �۸�ȼ�_In   In �շѼ�Ŀ.�۸�ȼ�%Type := Null
) Is
Begin
  Insert Into �շѼ�Ŀ
    (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �Ӱ�Ӽ���, �����շ���, �䶯ԭ��, ����˵��, ����id, ������, ִ������, ��ֹ����, NO, ���, ȱʡ�۸�, ���ۻ��ܺ�, �۸�ȼ�)
  Values
    (Id_In, ԭ��id_In, �շ�ϸĿid_In, ԭ��_In, �ּ�_In, ������Ŀid_In, �Ӱ�Ӽ���_In, �����շ���_In, �䶯ԭ��_In, ����˵��_In, ����id_In, ������_In, ִ������_In,
     To_Date('3000-01-01', 'yyyy-mm-dd'), No_In, ���_In, ȱʡ�۸�_In, ���ۻ��ܺ�_In, �۸�ȼ�_In);
  If �۸�ȼ�_In Is Null Then
    b_Message.Zlhis_Dict_053(�շ�ϸĿid_In, 1);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�շѼ�Ŀ_Insert;
/

--123263:������,2018-03-22,����ƽ̨��Ϣê��
Create Or Replace Procedure Zl_��Ѫ������_Clear
(
  Type_In  In Number := 0,
  Oldno_In In ��Ѫ������.����%Type := Null,
  Newno_In In ��Ѫ������.����%Type := Null
) Is
Begin
  If Type_In = 0 Then
    --- Ϊ���ּ����� 
    Delete ��Ѫ������;
  End If;

  If Type_In = 1 Then
    -- �ı��� 
    If Nvl(Oldno_In, 0) <> 0 Then
      If Nvl(Newno_In, 0) <> 0 Then
        Update ��Ѫ������ Set ���� = Newno_In Where ���� = Oldno_In;
      Else
        For R In (Select ����, ����, ����, ���, ��Ӽ�, ��Ѫ��, ��ɫ, ����id From ��Ѫ������ A Where a.���� = Oldno_In) Loop
          b_Message.Zlhis_Dictlis_009(r.����, r.����, r.����, r.���, r.��Ӽ�, r.��Ѫ��, r.��ɫ, r.����id);
        End Loop;
        Delete ��Ѫ������ Where ���� = Oldno_In;
      End If;    
      Update ������ĿĿ¼ Set �Թܱ��� = Newno_In Where �Թܱ��� = Oldno_In;
    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Ѫ������_Clear;
/

--123263:������,2018-03-22,����ƽ̨��Ϣê��
Create Or Replace Procedure Zl_��Ѫ������_Update
(
  ����_In   In ��Ѫ������.����%Type,
  ����_In   In ��Ѫ������.����%Type,
  ���_In   In ��Ѫ������.���%Type,
  ��Ӽ�_In In ��Ѫ������.��Ӽ�%Type,
  ��Ѫ��_In In ��Ѫ������.��Ѫ��%Type,
  ��ɫ_In   In ��Ѫ������.��ɫ%Type,
  ����id_In In ��Ѫ������.����id%Type := Null
) Is
  v_����id Number;
Begin
  If Nvl(����id_In, 0) <> 0 Then
    v_����id := ����id_In;
  Else
    v_����id := Null;
  End If;
  Update ��Ѫ������
  Set ���� = ����_In, ��� = ���_In, ��Ӽ� = ��Ӽ�_In, ��Ѫ�� = ��Ѫ��_In, ��ɫ = ��ɫ_In, ����id = v_����id
  Where ���� = ����_In;
  
  If Sql%NotFound Then
    Insert Into ��Ѫ������
      (����, ����, ���, ��Ӽ�, ��Ѫ��, ��ɫ, ����id)
    Values
      (����_In, ����_In, ���_In, ��Ӽ�_In, ��Ѫ��_In, ��ɫ_In, v_����id);
    b_Message.Zlhis_Dictlis_007(����_In, ����_In, Null, ���_In, ��Ӽ�_In, ��Ѫ��_In, ��ɫ_In, v_����id);
  else
    b_Message.Zlhis_Dictlis_008(����_In, ����_In, Null, ���_In, ��Ӽ�_In, ��Ѫ��_In, ��ɫ_In, v_����id);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Ѫ������_Update;
/

--122937:������,2018-03-21,����ӿ��޸�
Create Or Replace Procedure Zl_Third_Getadviceinfo
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ���ȡҽ��������Ϣ/��ѯ
  --������
  --����� Xml_In
  --<IN>
  --     <YZID>1156789</YZID>--��ҽ��ID
  --</IN>

  --���� Xml_Out
  --<OUTPUT>
  --    <YZ>
  --       <PATIID></PATIID>     --����ҽ����¼.����ID
  --       <PAGEID></PAGEID>     --����ҽ����¼.��ҳID
  --       <BABY></BABY>   --����ҽ����¼.Ӥ��
  --       <YZID>1145878</YZID>   --����ҽ����¼.ҽ��ID�� ��ҽ��ID
  --       <RELATEDID></RELATEDID>   --����ҽ����¼.���ID
  --       <ZXKSID></ZXKSID>   --����ҽ����¼.ִ�п���id
  --       <YZQX>0</YZQX>      --����ҽ����¼.ҽ����Ч
  --       <STATE>8</STATE>    --����ҽ����¼.ҽ��״̬
  --       <JJBZ>0</JJBZ>      --����ҽ����¼.������־
  --       <KZYS>����</KZYS>   --����ҽ����¼.����ҽ��
  --       <KZSJ>2015-03-25 16:37:00</KZSJ>   --����ҽ����¼.����ʱ��
  --       <ZLXMID></ZLXMID>   --������ĿĿ¼.ID
  --       <ZLLB>E</ZLLB>      --������ĿĿ¼.���
  --       <ZLXMMC></ZLXMMC>   --������ĿĿ¼.���� ����飬����(������ C)������(�������� F)����Ѫ(K)����ҩ�䷽(������ E)������(����)
  --       <ZLXMCZLX></<ZLXMCZLX>   --������ĿĿ¼.��������
  --       <ZLXMZXFL></ZLXMZXFL>   --������ĿĿ¼.ִ�з���
  --       <BZ>21</BZ> ������ĿĿ¼.��������||������ĿĿ¼.ִ�з���
  --       <YF>������ע</YF>   --����ҽ����¼.ҽ������ ����ҽ�����е�  ҽ������
  --       <PC>BID</PC>   --����Ƶ����Ŀ.Ӣ������
  --       <ZXSJFY>18-20</ZXSJFY>   --����ҽ����¼.ִ��ʱ�䷽��
  --       <PLCS>2</PLCS>   --����ҽ����¼.Ƶ�ʴ���
  --       <PLJG>1</PLJG>   --����ҽ����¼.Ƶ�ʼ��
  --       <PSJG></PSJG>   --����ҽ����¼.Ƥ�Խ��
  --       <YSZT></YSZT>   --����ҽ����¼.ҽ������
  --       <KSZXSJ>2015-03-25 16:35:00</KSZXSJ>  --����ҽ����¼.��ʼִ��ʱ��
  --       <ZXZZSJ></ZXZZSJ>   --����ҽ����¼.ִ����ֹʱ��
  --       <TZYS></TZYS>   --����ҽ����¼.ͣ��ҽ��
  --       <TZSJ></TZSJ>   --����ҽ����¼.ͣ��ʱ��
  --       <DW>��</DW>   --������ĿĿ¼.���㵥λ
  --       <DL></DL>   --����ҽ����¼.��������
  --       <ZL></ZL>   --����ҽ����¼.�ܸ�����

  --       <ITEMLIST> ����Ѫ��Ŀ����/��ҩҽ����Ŀ��ϸ�����Ϣ����Ѫ��Ѫ����Ϣ��ҩƷ����ϸ��Ϣ
  --        <ITEM>
  --         <YSZT></YSZT>   --����ҽ����¼.ҽ������
  --         <YZID>1145878</YZID>   --����ҽ����¼.ҽ��ID
  --         <RELATEDID></RELATEDID>   --����ҽ����¼.���ID
  --         <ZLXMID></ZLXMID>   --������ĿĿ¼.ID
  --         <SFXMID></SFXMID>   --�շ���ĿĿ¼.id
  --         <SFXMMC></SFXMMC>   --�շ���ĿĿ¼.����
  --         <SFXMGG></SFXMGG>   --�շ���ĿĿ¼.���
  --         <BM></BM>           --�շ���Ŀ����.���ƣ���Ʒ����
  --         <ZL></ZL>           --����ҽ����¼.�ܸ�����
  --         <DL>10</DL>         --����ҽ����¼.��������
  --         <DW>ml</DW>         --�շ���ĿĿ¼.���㵥λ
  --         <ZLDW>ml</ZLDW>   --������ĿĿ¼.���㵥λ
  --         <ZXXZ></ZXXZ>   --����ҽ����¼.ִ������
  --         <ZXKS></ZXKS>   --������ĿĿ¼.ִ�п���
  --         <XDBH></XDBH>   --ѪҺ�շ���¼.Ѫ�����
  --         <SXXH></SXXH>   --ѪҺ�շ���¼.���
  --        </ITEM>
  --        <ITEM/>...
  --       </ITEMLIST>
  --      </YZ>
  --</OUTPUT>

  n_ҽ��id  ����ҽ����¼.Id%Type;
  x_ҽ��    Xmltype;
  x_Item    Xmltype;
  v_Xtmp    Clob; --��ʱXML
  n_Cnt     Number;
  x_Templet Xmltype;

  v_Ӣ����     ����Ƶ����Ŀ.Ӣ������%Type;
  v_�Թ�����   ��Ѫ������.����%Type;
  v_��Ӽ�     ��Ѫ������.��Ӽ�%Type;
  v_�Թܹ��   ��Ѫ������.���%Type;
  n_�Թ���ɫ   ��Ѫ������.��ɫ%Type;
  v_�շ���Ʒ�� �շ���Ŀ����.����%Type;
  n_����Ѫ��   Number;
  v_SqlѪ��    Varchar2(4000);
  n_Ѫ������id Number(18);
  v_Tmp��Ѫ    Varchar2(4000);

  Type Bloodlist_Type Is Ref Cursor;
  Cbloodlist Bloodlist_Type;

  Type t_Code Is Record(
    ID       �շ���ĿĿ¼.Id%Type,
    ����     �շ���ĿĿ¼.����%Type,
    ���     �շ���ĿĿ¼.���%Type,
    ��λ     �շ���ĿĿ¼.���㵥λ%Type,
    Ѫ����� Varchar2(50),
    ���     Number(5));
  r_b t_Code;

Begin

  Select Extractvalue(Value(A), 'IN/YZID') Into n_ҽ��id From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_Cnt := 0;
  For R In (Select a.����id, a.��ҳid, a.Ӥ��, a.Id As ҽ��id, a.���id, a.ִ�п���id, a.ҽ����Ч, a.ҽ��״̬, a.������־, a.����ҽ��, a.����ʱ��, a.������Ŀid,
                   a.�������, a.ҽ������, a.ִ��ʱ�䷽��, a.ִ��Ƶ��, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.Ƥ�Խ��, a.ҽ������, a.��ʼִ��ʱ��, a.ִ����ֹʱ��, a.ͣ��ҽ��, a.ͣ��ʱ��,
                   b.���� As ��Ŀ����, b.��������, b.ִ�з���, b.���㵥λ As ���Ƶ�λ, a.��������, a.�ܸ�����, a.�걾��λ, a.��鷽��, a.�շ�ϸĿid, c.���� As �շ�����,
                   c.���, Null As �շ���Ʒ��, c.���㵥λ As �շѵ�λ, a.ִ������, b.ִ�п���, b.�Թܱ���, c.����, d.��ΣҩƷ
            From ����ҽ����¼ A, ������ĿĿ¼ B, �շ���ĿĿ¼ C, ҩƷ��� D
            Where a.������Ŀid = b.Id And a.�շ�ϸĿid = c.Id(+) And a.�շ�ϸĿid = d.ҩƷid(+) And (a.Id = n_ҽ��id Or a.���id = n_ҽ��id)
            Order By a.���) Loop
    n_Cnt := n_Cnt + 1;
    If n_Cnt = 1 Then
      Select Max(a.Ӣ������) Into v_Ӣ���� From ����Ƶ����Ŀ A Where a.���� = r.ִ��Ƶ��;
    End If;
    v_�Թ����� := Null;
    v_��Ӽ�   := Null;
    v_�Թܹ�� := Null;
    n_�Թ���ɫ := Null;
    If r.�Թܱ��� Is Not Null Then
      Select Max(a.����), Max(a.��Ӽ�), Max(a.���), Max(a.��ɫ)
      Into v_�Թ�����, v_��Ӽ�, v_�Թܹ��, n_�Թ���ɫ
      From ��Ѫ������ A
      Where a.���� = r.�Թܱ���;
    End If;
    --��ҽ��
    If r.���id Is Null Then
      v_Xtmp := '<YZ>';
      v_Xtmp := v_Xtmp || '<PATIID>' || r.����id || '</PATIID>'; --����ҽ����¼.����ID
      v_Xtmp := v_Xtmp || '<PAGEID>' || r.��ҳid || '</PAGEID>'; --����ҽ����¼.��ҳID
      v_Xtmp := v_Xtmp || '<BABY>' || r.Ӥ�� || '</BABY>'; --����ҽ����¼.Ӥ��
      v_Xtmp := v_Xtmp || '<YZID>' || r.ҽ��id || '</YZID>'; --����ҽ����¼.ҽ��ID
      v_Xtmp := v_Xtmp || '<RELATEDID>' || r.���id || '</RELATEDID>'; --����ҽ����¼.���ID
      v_Xtmp := v_Xtmp || '<ZXKSID>' || r.ִ�п���id || '</ZXKSID>'; --����ҽ����¼.ִ�п���id
      v_Xtmp := v_Xtmp || '<YZQX>' || r.ҽ����Ч || '</YZQX>'; --����ҽ����¼.ҽ����Ч
      v_Xtmp := v_Xtmp || '<STATE>' || r.ҽ��״̬ || '</STATE>'; --����ҽ����¼.ҽ��״̬
      v_Xtmp := v_Xtmp || '<JJBZ>' || r.������־ || '</JJBZ>'; --����ҽ����¼.������־
      v_Xtmp := v_Xtmp || '<KZYS>' || r.����ҽ�� || '</KZYS>'; --����ҽ����¼.����ҽ��
      v_Xtmp := v_Xtmp || '<KZSJ>' || To_Char(r.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</KZSJ>'; --����ҽ����¼.����ʱ��
      v_Xtmp := v_Xtmp || '<BZ>' || r.�������� || r.ִ�з��� || '</BZ>'; -- ������ĿĿ¼.��������||������ĿĿ¼.ִ�з���
      v_Xtmp := v_Xtmp || '<ZLXMID>' || r.������Ŀid || '</ZLXMID>'; --������ĿĿ¼.ID
      v_Xtmp := v_Xtmp || '<ZLLB>' || r.������� || '</ZLLB>'; --������ĿĿ¼.���
      v_Xtmp := v_Xtmp || '<YZNR>' || r.ҽ������ || '</YZNR>'; --ҽ������
      v_Xtmp := v_Xtmp || '<YF>' || r.��Ŀ���� || '</YF>'; --����ҽ����¼.ҽ������
      v_Xtmp := v_Xtmp || '<PC>' || v_Ӣ���� || '</PC>'; --����Ƶ����Ŀ.Ӣ������
      v_Xtmp := v_Xtmp || '<ZXSJFY>' || r.ִ��ʱ�䷽�� || '</ZXSJFY>'; --����ҽ����¼.ִ��ʱ�䷽��
      v_Xtmp := v_Xtmp || '<PLCS>' || r.Ƶ�ʴ��� || '</PLCS>'; --����ҽ����¼.Ƶ�ʴ���
      v_Xtmp := v_Xtmp || '<PLJG>' || r.Ƶ�ʼ�� || '</PLJG>'; --����ҽ����¼.Ƶ�ʼ��
      v_Xtmp := v_Xtmp || '<PSJG>' || r.Ƥ�Խ�� || '</PSJG>'; --����ҽ����¼.Ƥ�Խ��
      v_Xtmp := v_Xtmp || '<YSZT>' || r.ҽ������ || '</YSZT>'; --����ҽ����¼.ҽ������
      v_Xtmp := v_Xtmp || '<KSZXSJ>' || To_Char(r.��ʼִ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</KSZXSJ>'; --����ҽ����¼.��ʼִ��ʱ��
      v_Xtmp := v_Xtmp || '<ZXZZSJ>' || To_Char(r.ִ����ֹʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</ZXZZSJ>'; --����ҽ����¼.ִ����ֹʱ��
      v_Xtmp := v_Xtmp || '<TZYS>' || r.ͣ��ҽ�� || '</TZYS>'; --����ҽ����¼.ͣ��ҽ��
      v_Xtmp := v_Xtmp || '<TZSJ>' || To_Char(r.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</TZSJ>'; --����ҽ����¼.ͣ��ʱ��
      v_Xtmp := v_Xtmp || '<ZLXMMC>' || r.��Ŀ���� || '</ZLXMMC>'; --������ĿĿ¼.����
      v_Xtmp := v_Xtmp || '<ZLXMCZLX>' || r.�������� || '</ZLXMCZLX>'; --������ĿĿ¼.��������
      v_Xtmp := v_Xtmp || '<ZLXMZXFL>' || r.ִ�з��� || '</ZLXMZXFL>'; --������ĿĿ¼.ִ�з���
      --       (����Ѫ�ܷ���)
      v_Xtmp := v_Xtmp || '<CXGMC>' || v_�Թ����� || '</CXGMC>'; --��Ѫ������
      v_Xtmp := v_Xtmp || '<CXGTJJ>' || v_��Ӽ� || '</CXGTJJ>'; --��Ѫ����Ӽ�
      v_Xtmp := v_Xtmp || '<CXGGG>' || v_�Թܹ�� || '</CXGGG>'; --��Ѫ�ܹ��
      v_Xtmp := v_Xtmp || '<CXGYS>' || n_�Թ���ɫ || '</CXGYS>'; --��Ѫ����ɫ
      v_Xtmp := v_Xtmp || '<DW>' || r.���Ƶ�λ || '</DW>'; --������ĿĿ¼.���㵥λ
      v_Xtmp := v_Xtmp || '<DL>' || r.�������� || '</DL>'; --����ҽ����¼.��������
      v_Xtmp := v_Xtmp || '<ZL>' || r.�ܸ����� || '</ZL>'; --����ҽ����¼.�ܸ�����
      v_Xtmp := v_Xtmp || '</YZ>';
      x_ҽ�� := Xmltype(v_Xtmp);
    End If;
  
    --��Ѫ
    If r.������� = 'K' Then
      --�ж��Ƿ�װѪ��
      Select Zl_Checkobject(1, 'ѪҺ�շ���¼') Into n_����Ѫ�� From Dual;
      If n_����Ѫ�� > 0 Then
        n_Ѫ������id := r.ҽ��id;
        --ҽ������
        v_Xtmp    := '<YSZT>' || r.ҽ������ || '</YSZT>'; --����ҽ����¼.ҽ������
        v_Xtmp    := v_Xtmp || '<YZID>' || r.ҽ��id || '</YZID>'; --����ҽ����¼.ҽ��ID
        v_Xtmp    := v_Xtmp || '<RELATEDID>' || r.���id || '</RELATEDID>'; --����ҽ����¼.���ID
        v_Xtmp    := v_Xtmp || '<ZLXMID>' || r.������Ŀid || '</ZLXMID>'; --������ĿĿ¼.ID
        v_Xtmp    := v_Xtmp || '<ZL>' || r.�ܸ����� || '</ZL>'; --����ҽ����¼.�ܸ�����
        v_Xtmp    := v_Xtmp || '<DL>' || r.�������� || '</DL>'; --����ҽ����¼.��������
        v_Xtmp    := v_Xtmp || '<ZLDW>' || r.���Ƶ�λ || '</ZLDW>'; --������ĿĿ¼.���㵥λ
        v_Xtmp    := v_Xtmp || '<ZXXZ>' || r.ִ������ || '</ZXXZ>'; --����ҽ����¼.ִ������
        v_Xtmp    := v_Xtmp || '<ZXKS>' || r.ִ�п��� || '</ZXKS>'; --������ĿĿ¼.ִ�п���
        v_Tmp��Ѫ := v_Xtmp;
        If r.��鷽�� = '1' Then
          v_SqlѪ�� := 'Select d.Id,d.����,d.���,d.���㵥λ as ��λ, a.Ѫ�����,a.���
                       From ѪҺ�շ���¼ a,ѪҺ���ͼ�¼ b,ѪҺ��Ѫ��¼ c,�շ���ĿĿ¼ d
                       Where a.Id = b.�շ�id And b.�䷢id = c.Id and a.ѪҺid =d.id  And c.����id =:1';
        End If;
      End If;
    Elsif r.���id Is Not Null And r.������� = 'E' And r.�������� = '8' And Nvl(r.ִ�з���, 0) = 0 And n_����Ѫ�� = 1 And
          v_SqlѪ�� Is Null Then
      v_SqlѪ�� := 'Select b.Id,b.����,  b.���,b.���㵥λ as ��λ, a.Ѫ�����,a.���
                  From ѪҺ�շ���¼ a,�շ���ĿĿ¼ b
                  Where a.ѪҺid =b.id and a.�䷢id = (Select Id From ѪҺ��Ѫ��¼ Where ����id=:1)';
    Else
      v_SqlѪ�� := Null;
    End If;
  
    If v_SqlѪ�� Is Not Null And n_Ѫ������id Is Not Null Then
      --��Ѫҽ����ֻ�з�ҽ����ſ�����Ѫ����Ϣ
      x_Item := Xmltype('<ITEMLIST></ITEMLIST>');
      Open Cbloodlist For v_SqlѪ��
        Using n_Ѫ������id;
      Loop
        Fetch Cbloodlist
          Into r_b.Id, r_b.����, r_b.���, r_b.��λ, r_b.Ѫ�����, r_b.���;
        Exit When Cbloodlist%NotFound;
        v_�շ���Ʒ�� := Null;
        If r_b.Id Is Not Null Then
          For Z In (Select a.����, a.����
                    From �շ���Ŀ���� A
                    Where a.�շ�ϸĿid = r_b.Id
                    Group By a.����, a.����
                    Order By a.����) Loop
            v_�շ���Ʒ�� := z.����;
            If z.���� = 3 Then
              v_�շ���Ʒ�� := z.����;
              Exit;
            End If;
          End Loop;
        End If;
      
        v_Xtmp := '<ITEM jsonArray="True" >';
      
        v_Xtmp := v_Xtmp || v_Tmp��Ѫ;
      
        --Ѫ�ⲿ��
        v_Xtmp := v_Xtmp || '<SFXMID>' || r_b.Id || '</SFXMID>'; --�շ���ĿĿ¼.id
        v_Xtmp := v_Xtmp || '<SFXMMC>' || r_b.���� || '</SFXMMC>'; --�շ���ĿĿ¼.����
        v_Xtmp := v_Xtmp || '<SFXMGG>' || r_b.��� || '</SFXMGG>'; --�շ���ĿĿ¼.���
        v_Xtmp := v_Xtmp || '<BM>' || v_�շ���Ʒ�� || '</BM>'; --�շ���Ŀ����.���ƣ���Ʒ����
        v_Xtmp := v_Xtmp || '<DW>' || r_b.��λ || '</DW>'; --�շ���ĿĿ¼.���㵥λ
        v_Xtmp := v_Xtmp || '<XDBH>' || r_b.Ѫ����� || '</XDBH>'; --ѪҺ�շ���¼.Ѫ�����
        v_Xtmp := v_Xtmp || '<SXXH>' || r_b.��� || '</SXXH>'; --ѪҺ�շ���¼.���
      
        v_Xtmp := v_Xtmp || '</ITEM>';
        Select Appendchildxml(x_Item, '/ITEMLIST', Xmltype(v_Xtmp)) Into x_Item From Dual;
      End Loop;
      Close Cbloodlist;
    End If;
  
    --��ҩ��ҩҽ��
    If r.������� = '5' Or r.������� = '6' Then
      --��/�� ҩ
      If x_Item Is Null Then
        --ֻ��ʼ��һ��
        x_Item := Xmltype('<ITEMLIST></ITEMLIST>');
      End If;
      v_�շ���Ʒ�� := Null;
      If r.�շ�ϸĿid Is Not Null Then
        For Z In (Select a.����, a.����
                  From �շ���Ŀ���� A
                  Where a.�շ�ϸĿid = r.�շ�ϸĿid
                  Group By a.����, a.����
                  Order By a.����) Loop
          v_�շ���Ʒ�� := z.����;
          If z.���� = 3 Then
            v_�շ���Ʒ�� := z.����;
            Exit;
          End If;
        End Loop;
      End If;
    
      v_Xtmp := '<ITEM jsonArray="True" >';
      v_Xtmp := v_Xtmp || '<YSZT>' || r.ҽ������ || '</YSZT>'; --����ҽ����¼.ҽ������
      v_Xtmp := v_Xtmp || '<YZID>' || r.ҽ��id || '</YZID>'; --����ҽ����¼.ҽ��ID
      v_Xtmp := v_Xtmp || '<RELATEDID>' || r.���id || '</RELATEDID>'; --����ҽ����¼.���ID
      v_Xtmp := v_Xtmp || '<GW>' || nvl(r.��ΣҩƷ,0) || '</GW>'; --��Σҩ��ʶ��1��ʾ��Σҩ��0��ʾ��ͨ
      v_Xtmp := v_Xtmp || '<CDM>' || r.���� || '</CDM>'; --���������շ���ĿĿ¼.����    
      v_Xtmp := v_Xtmp || '<ZLXMID>' || r.������Ŀid || '</ZLXMID>'; --������ĿĿ¼.ID
      v_Xtmp := v_Xtmp || '<SFXMID>' || r.�շ�ϸĿid || '</SFXMID>'; --�շ���ĿĿ¼.id
      v_Xtmp := v_Xtmp || '<SFXMMC>' || r.�շ����� || '</SFXMMC>'; --�շ���ĿĿ¼.����
      v_Xtmp := v_Xtmp || '<SFXMGG>' || r.��� || '</SFXMGG>'; --�շ���ĿĿ¼.���
      v_Xtmp := v_Xtmp || '<BM>' || v_�շ���Ʒ�� || '</BM>'; --�շ���Ŀ����.���ƣ���Ʒ����
      v_Xtmp := v_Xtmp || '<ZL>' || r.�ܸ����� || '</ZL>'; --����ҽ����¼.�ܸ�����
      v_Xtmp := v_Xtmp || '<DL>' || r.�������� || '</DL>'; --����ҽ����¼.��������
      v_Xtmp := v_Xtmp || '<DW>' || r.�շѵ�λ || '</DW>'; --�շ���ĿĿ¼.���㵥λ
      v_Xtmp := v_Xtmp || '<ZLDW>' || r.���Ƶ�λ || '</ZLDW>'; --������ĿĿ¼.���㵥λ
      v_Xtmp := v_Xtmp || '<ZXXZ>' || r.ִ������ || '</ZXXZ>'; --����ҽ����¼.ִ������
      v_Xtmp := v_Xtmp || '<ZXKS>' || r.ִ�п��� || '</ZXKS>'; --������ĿĿ¼.ִ�п���
      v_Xtmp := v_Xtmp || '<XDBH></XDBH>'; --ѪҺ�շ���¼.Ѫ�����
      v_Xtmp := v_Xtmp || '<SXXH></SXXH>'; --ѪҺ�շ���¼.���
      v_Xtmp := v_Xtmp || '</ITEM>';
      Select Appendchildxml(x_Item, '/ITEMLIST', Xmltype(v_Xtmp)) Into x_Item From Dual;
    End If;
  End Loop;
  If x_Item Is Not Null Then
    Select Appendchildxml(x_ҽ��, '/YZ', x_Item) Into x_ҽ�� From Dual;
  End If;
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');
  Select Appendchildxml(x_Templet, '/OUTPUT', x_ҽ��) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getadviceinfo;
/
--123263:������,2018-03-22,����ƽ̨��Ϣê��
--122998:������,2018-03-21,����ƽ̨��Ϣê��
--122609:������,2018-03-08,����ƽ̨��Ϣ���
--123312:���Ʊ�,2018-03-22,�²���ϵͳ����ƽ̨��Ϣ
--123386:��˶,2018-03-23,�շѼ�Ŀ���շѶ���ê��
CREATE OR REPLACE Package b_Message Is
  Procedure p_Msg_Todo_Insert
  (
    Msg_Code_In  Zlmsg_Todo.Msg_Code%Type,
    Key_Value_In Zlmsg_Todo.Key_Value%Type
  );
  --����ƽ̨��������
  Procedure Set_Platform_Call(Platform_Call Number);
  --��������
  Procedure Zlhis_Dict_001(Id_In ���ű�.Id%Type);
  --�޸Ĳ���
  Procedure Zlhis_Dict_002(����id_In ���ű�.Id%Type);
  --ͣ�ò���
  Procedure Zlhis_Dict_003(����id_In ���ű�.Id%Type);
  --���ò���
  Procedure Zlhis_Dict_004(����id_In ���ű�.Id%Type);
  --������Ա
  Procedure Zlhis_Dict_005(��Աid_In ��Ա��.Id%Type);
  --�޸���Ա
  Procedure Zlhis_Dict_006(��Աid_In ��Ա��.Id%Type);
  --ͣ����Ա
  Procedure Zlhis_Dict_007(��Աid_In ��Ա��.Id%Type);
  --������Ա
  Procedure Zlhis_Dict_008(��Աid_In ��Ա��.Id%Type);
  --�����շ���Ŀ
  Procedure Zlhis_Dict_009(ϸĿid_In �շ���ĿĿ¼.Id%Type);
  --�޸��շ���Ŀ
  Procedure Zlhis_Dict_010(ϸĿid_In �շ���ĿĿ¼.Id%Type);
  --ͣ���շ���Ŀ
  Procedure Zlhis_Dict_011(ϸĿid_In �շ���ĿĿ¼.Id%Type);
  --�����շ���Ŀ
  Procedure Zlhis_Dict_012(ϸĿid_In �շ���ĿĿ¼.Id%Type);
  --����������Ŀ
  Procedure Zlhis_Dict_013(����id_In ������ĿĿ¼.Id%Type);
  --�޸�������Ŀ
  Procedure Zlhis_Dict_014(����id_In ������ĿĿ¼.Id%Type);
  --ͣ��������Ŀ
  Procedure Zlhis_Dict_015(����id_In ������ĿĿ¼.Id%Type);
  --����������Ŀ
  Procedure Zlhis_Dict_016(����id_In ������ĿĿ¼.Id%Type);
  --����������Ŀ
  Procedure Zlhis_Dict_017(����id_In ������ĿĿ¼.Id%Type);
  --�޸ļ�����Ŀ
  Procedure Zlhis_Dict_018(����id_In ������ĿĿ¼.Id%Type);
  --ɾ��������Ŀ
  Procedure Zlhis_Dict_019
  (
    ����id_In ������ĿĿ¼.Id%Type,
    ����_In   ����������Ŀ.����%Type,
    ������_In ����������Ŀ.������%Type,
    Ӣ����_In ����������Ŀ.Ӣ����%Type
  );

  --������������Ŀ¼
  Procedure Zlhis_Dict_021(����id_In ��������Ŀ¼.Id%Type);
  --�޸ļ�������Ŀ¼
  Procedure Zlhis_Dict_022(����id_In ��������Ŀ¼.Id%Type);
  --ͣ�ü�������Ŀ¼
  Procedure Zlhis_Dict_023(����id_In ��������Ŀ¼.Id%Type);
  --���ü�������Ŀ¼
  Procedure Zlhis_Dict_024(����id_In ��������Ŀ¼.Id%Type);
  --����ҩƷ����
  Procedure Zlhis_Dict_025
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --�޸�ҩƷ����
  Procedure Zlhis_Dict_026
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --ɾ��ҩƷ����
  Procedure Zlhis_Dict_027
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type
  );
  --ͣ��ҩƷ����
  Procedure Zlhis_Dict_028
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --����ҩƷ����
  Procedure Zlhis_Dict_029
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --����ҩƷƷ��
  Procedure Zlhis_Dict_030
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  );
  --�޸�ҩƷƷ��
  Procedure Zlhis_Dict_031
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  );
  --ɾ��ҩƷƷ��
  Procedure Zlhis_Dict_032
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type,
    ����_In ������ĿĿ¼.����%Type,
    ����_In ������ĿĿ¼.����%Type
  );
  --ͣ��ҩƷƷ��
  Procedure Zlhis_Dict_033
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  );
  --����ҩƷƷ��
  Procedure Zlhis_Dict_034
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  );
  --����ҩƷ���
  Procedure Zlhis_Dict_035
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --�޸�ҩƷ���
  Procedure Zlhis_Dict_036
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --ɾ��ҩƷ���
  Procedure Zlhis_Dict_037
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type,
    ����_In   �շ���ĿĿ¼.����%Type,
    ����_In   �շ���ĿĿ¼.����%Type,
    ���_In   �շ���ĿĿ¼.���%Type,
    ����_In   �շ���ĿĿ¼.����%Type
  );
  --ͣ��ҩƷ���
  Procedure Zlhis_Dict_038
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --����ҩƷ���
  Procedure Zlhis_Dict_039
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --����ҩƷ�洢�ⷿ
  Procedure Zlhis_Dict_040
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --����ҩƷ�����޶�
  Procedure Zlhis_Dict_041
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --��������Ʒ��
  Procedure Zlhis_Dict_042(Id_In ������ĿĿ¼.Id%Type);
  --�������Ĺ��
  Procedure Zlhis_Dict_043(Id_In �շ���ĿĿ¼.Id%Type);
  --�޸����Ĺ��
  Procedure Zlhis_Dict_044(Id_In �շ���ĿĿ¼.Id%Type);
  --ɾ�����Ĺ��
  Procedure Zlhis_Dict_045
  (
    Id_In �շ���ĿĿ¼.Id%Type,
    ����_In   �շ���ĿĿ¼.����%Type,
    ����_In   �շ���ĿĿ¼.����%Type,
    ���_In   �շ���ĿĿ¼.���%Type,
    ����_In   �շ���ĿĿ¼.����%Type
  );
  --ͣ�����Ĺ��
  Procedure Zlhis_Dict_046(Id_In �շ���ĿĿ¼.Id%Type);
  --�������Ĺ��
  Procedure Zlhis_Dict_047(Id_In �շ���ĿĿ¼.Id%Type);
  --ҽ������
  Procedure Zlhis_Dict_048
  (
    ����_In       In ����֧����Ŀ.����%Type,
    �շ�ϸĿid_In In ����֧����Ŀ.�շ�ϸĿid%Type
  );
  --ɾ��ҽ������
  Procedure Zlhis_Dict_049
  (
    ����_In       In ����֧����Ŀ.����%Type,
    �շ�ϸĿid_In In ����֧����Ŀ.�շ�ϸĿid%Type,
    ��Ŀ����_In   In �շ���ĿĿ¼.����%Type,
    ��Ŀ����_In   In �շ���ĿĿ¼.����%Type,
    ҽ������_In   In ����֧����Ŀ.��Ŀ����%Type,
    ҽ������_In   In ����֧����Ŀ.��Ŀ����%Type
  );
  --�������ķ���
  Procedure Zlhis_Dict_050
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --�޸����ķ���
  Procedure Zlhis_Dict_051
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --ɾ�����ķ���
  Procedure Zlhis_Dict_052
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type
  );
  --�շѼ�Ŀ�䶯
  Procedure Zlhis_Dict_053
  (
    Id_In       �շ���ĿĿ¼.Id%Type,
    �䶯����_In Number
  );
  --�����շѶ��ձ䶯
  Procedure Zlhis_Dict_054
  (
    Id_In     ���Ʒ���Ŀ¼.Id%Type,
    ԭ����_In Varchar2,
    �ֶ���_In Varchar2
  );
  --�������Ƽ������
  Procedure Zlhis_Dictpacs_001
  (
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ������_In ���Ƽ������.������%Type
  );
  --�޸����Ƽ������
  Procedure Zlhis_Dictpacs_002
  (
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ������_In ���Ƽ������.������%Type
  );
  --ɾ�����Ƽ������
  Procedure Zlhis_Dictpacs_003
  (
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ������_In ���Ƽ������.������%Type
  );
  --�������Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_004
  (
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ��ע_In     ���Ƽ�鲿λ.��ע%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    �����Ա�_In ���Ƽ�鲿λ.�����Ա�%Type
  );
  --�޸����Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_005
  (
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ��ע_In     ���Ƽ�鲿λ.��ע%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    �����Ա�_In ���Ƽ�鲿λ.�����Ա�%Type
  );
  --ɾ�����Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_006
  (
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ��ע_In     ���Ƽ�鲿λ.��ע%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    �����Ա�_In ���Ƽ�鲿λ.�����Ա�%Type
  );
  --����������Ŀ��λ
  Procedure Zlhis_Dictpacs_007
  (
    Id_In     ������Ŀ��λ.Id%Type,
    ��Ŀid_In ������Ŀ��λ.��Ŀid%Type,
    ����_In   ������Ŀ��λ.����%Type,
    ��λ_In   ������Ŀ��λ.��λ%Type,
    ����_In   ������Ŀ��λ.����%Type,
    Ĭ��_In   ������Ŀ��λ.Ĭ��%Type
  );
  --�޸�������Ŀ��λ
  Procedure Zlhis_Dictpacs_008
  (
    Id_In     ������Ŀ��λ.Id%Type,
    ��Ŀid_In ������Ŀ��λ.��Ŀid%Type,
    ����_In   ������Ŀ��λ.����%Type,
    ��λ_In   ������Ŀ��λ.��λ%Type,
    ����_In   ������Ŀ��λ.����%Type,
    Ĭ��_In   ������Ŀ��λ.Ĭ��%Type
  );
  --ɾ��������Ŀ��λ
  Procedure Zlhis_Dictpacs_009
  (
    Id_In     ������Ŀ��λ.Id%Type,
    ��Ŀid_In ������Ŀ��λ.��Ŀid%Type,
    ����_In   ������Ŀ��λ.����%Type,
    ��λ_In   ������Ŀ��λ.��λ%Type,
    ����_In   ������Ŀ��λ.����%Type,
    Ĭ��_In   ������Ŀ��λ.Ĭ��%Type
  );
    --�������Ƽ���걾
  Procedure Zlhis_DictLis_004
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In ���Ƽ���걾.����%Type,
    ����_In   ���Ƽ���걾.����%Type,
    �����Ա�_In   ���Ƽ���걾.�����Ա�%Type
  );
    --�޸����Ƽ���걾
  Procedure Zlhis_DictLis_005
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In ���Ƽ���걾.����%Type,
    ����_In   ���Ƽ���걾.����%Type,
    �����Ա�_In   ���Ƽ���걾.�����Ա�%Type
  );
    --ɾ��������Ŀ��λ
  Procedure Zlhis_DictLis_006
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In ���Ƽ���걾.����%Type,
    ����_In   ���Ƽ���걾.����%Type,
    �����Ա�_In   ���Ƽ���걾.�����Ա�%Type
  );
  --������Ѫ������
    Procedure Zlhis_DictLis_007
  (
    ����_In     ��Ѫ������.����%Type,
    ����_In ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In   ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In   ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����ID_In   ��Ѫ������.����ID%Type
  );
  --�޸Ĳ�Ѫ������
    Procedure Zlhis_DictLis_008
  (
    ����_In     ��Ѫ������.����%Type,
    ����_In ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In   ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In   ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����ID_In   ��Ѫ������.����ID%Type
  );
  --ɾ����Ѫ������
    Procedure Zlhis_DictLis_009
  (
    ����_In     ��Ѫ������.����%Type,
    ����_In ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In   ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In   ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����ID_In   ��Ѫ������.����ID%Type
  );
  
  --ҩƷ��ҩ����
  Procedure Zlhis_Drug_001(No_In ҩƷ�շ���¼.No%Type);
  --ȡ��ҩƷ��ҩ����
  Procedure Zlhis_Drug_002(No_In ҩƷ�շ���¼.No%Type);
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_003(No_In ҩƷ�շ���¼.No%Type);
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_004(No_In ҩƷ�շ���¼.No%Type);
  --���ŷ�ҩ
  Procedure Zlhis_Drug_005
  (
    �ⷿid_In ҩƷ�շ���¼.�ⷿid%Type,
    �շ�id_In ҩƷ�շ���¼.Id%Type
  );
  --������ҩ
  Procedure Zlhis_Drug_006
  (
    �����շ�id_In ҩƷ�շ���¼.Id%Type,
    �����շ�id_In ҩƷ�շ���¼.Id%Type,
    ����_In       ҩƷ�շ���¼.ʵ������%Type,
    ����id_In     ������ü�¼.Id%Type
  );
  --ҩƷ����
  Procedure Zlhis_Drug_007
  (
    �۸�id_In   ҩƷ�۸��¼.Id%Type
  );
  --���䷢��
  Procedure ZLHIS_DRUG_008
  (
    ��¼Ids_In Varchar2
  );
  --ҩƷ���ۼ�
  Procedure Zlhis_Drug_009
  (
    �۸�id_In   ҩƷ�۸��¼.Id%Type,
    ʱ��_In Number
  );
  --���ĵ��ɱ���
  Procedure Zlhis_Drug_010
  (
    �۸�id_In   �ɱ��۵�����Ϣ.ID%Type
  );
  --���ĵ��ۼ�
  Procedure Zlhis_Drug_011
  (
    �۸�id_In   �շѼ�Ŀ.Id%Type,
    ʱ��_In Number
  );
  --2.ֹͣ����ҽ����סԺ
  Procedure Zlhis_Cis_002
  (
    ����id_In  In ����ҽ����¼.����id%Type,
    ��ҳid_In  In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In  In ����ҽ����¼.Id%Type,
    ҽ��ids_In In Varchar2
  );
  --3.���ϻ���ҽ��������/סԺ
  Procedure Zlhis_Cis_003
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --4.��������ҽ����סԺ
  Procedure Zlhis_Cis_004
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --5.������������ҽ����סԺ
  Procedure Zlhis_Cis_005
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --6.���߻�����ҽ����סԺ
  Procedure Zlhis_Cis_006
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --7.�������߻�����ҽ����סԺ
  Procedure Zlhis_Cis_007
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --10.�´ﻼ����ϣ�����/סԺ
  Procedure Zlhis_Cis_010
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ���id_In In ������ϼ�¼.Id%Type
  );
  --11.�����������
  Procedure Zlhis_Cis_011
  (
    ����id_In   In ���˹Һż�¼.����id%Type,
    ����id_In   In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    Id_In       In ������ϼ�¼.Id%Type,
    ����id_In   In ������ϼ�¼.����id%Type,
    ���id_In   In ������ϼ�¼.���id%Type,
    �������_In In ������ϼ�¼.�������%Type
  );

  --����ִ��ҽ��У��
  Procedure Zlhis_Cis_012
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --13.����Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_014
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In ����ҽ����¼.Id%Type,
    ��Ϣid_In In ҵ����Ϣ�嵥.Id%Type
  );

  --15.���߼������룬����/סԺ
  Procedure Zlhis_Cis_016
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type
  );
  --16.���߼�����룬����/סԺ
  Procedure Zlhis_Cis_017
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type
  );
  --17.�����������룬����/סԺ
  Procedure Zlhis_Cis_018
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --18.������Ѫ���룬סԺ
  Procedure Zlhis_Cis_019
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --19.���߻������룬סԺ
  Procedure Zlhis_Cis_020
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --20.��������ҽ����סԺ
  Procedure Zlhis_Cis_021
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --21.��������ҽ����סԺ
  Procedure Zlhis_Cis_022
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --22.������������ҽ����סԺ
  Procedure Zlhis_Cis_023
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --24.���Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_025
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In ����ҽ����¼.Id%Type,
    ��Ϣid_In In ҵ����Ϣ�嵥.Id%Type
  );

  --����ִ��ҽ������
  Procedure Zlhis_Cis_026
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --�������߼�������
  Procedure Zlhis_Cis_036
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type
  );

  --�������߼������
  Procedure Zlhis_Cis_037
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type
  );

  --����������������
  Procedure Zlhis_Cis_038
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --����������Ѫ����
  Procedure Zlhis_Cis_039
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --�������߻�������
  Procedure Zlhis_Cis_040
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --������������ҽ��
  Procedure Zlhis_Cis_041
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --������������ҽ��
  Procedure Zlhis_Cis_042
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --������������ҽ��
  Procedure Zlhis_Cis_043
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --��������ִ��ҽ��
  Procedure Zlhis_Cis_044
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ��������_In In ����ҽ������.��������%Type,
    �״�ʱ��_In In ����ҽ������.�״�ʱ��%Type,
    ĩ��ʱ��_In In ����ҽ������.ĩ��ʱ��%Type,
    ��������_In In ����ҽ������.��������%Type
  );
  --����ҽ��ִ�еǼ�
  Procedure Zlhis_Cis_050
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    Ҫ��ʱ��_In In ����ҽ��ִ��.Ҫ��ʱ��%Type,
    ִ��ʱ��_In In ����ҽ��ִ��.ִ��ʱ��%Type
  );

  --����ҽ��ȡ��ִ�еǼ�
  Procedure Zlhis_Cis_051
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    Ҫ��ʱ��_In In ����ҽ��ִ��.Ҫ��ʱ��%Type,
    ִ��ʱ��_In In ����ҽ��ִ��.ִ��ʱ��%Type,
    ��������_In In ����ҽ��ִ��.��������%Type,
    ִ�н��_In In ����ҽ��ִ��.ִ�н��%Type,
    ִ��ժҪ_In In ����ҽ��ִ��.ִ��ժҪ%Type,
    ִ�п���_In In ����ҽ��ִ��.ִ�п���id%Type,
    ִ����_In   In ����ҽ��ִ��.ִ����%Type,
    �˶���_In   In ����ҽ��ִ��.�˶���%Type,
    ��¼��Դ_In In ����ҽ��ִ��.��¼��Դ%Type
  );
  --����ҽ��ִ�����
  Procedure Zlhis_Cis_052
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --����ҽ������ִ�����
  Procedure Zlhis_Cis_053
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --�������뷢�ͺ��޸�
  Procedure Zlhis_Cis_056
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type
  );

  --26.��鱨����ɣ�������ʱ
  Procedure Zlhis_Pacs_001
  (
    ҽ��id_In   In Ӱ�����¼.ҽ��id%Type,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  );
  --27.���״̬ͬ�������״̬�ı��
  Procedure Zlhis_Pacs_002
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԭ״̬_In In ����ҽ������.ִ�й���%Type,
    ��״̬_In In ����ҽ������.ִ�й���%Type
  );
  --28.���״̬���ˣ����״̬���˺�
  Procedure Zlhis_Pacs_003
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԭ״̬_In In ����ҽ������.ִ�й���%Type,
    ��״̬_In In ����ҽ������.ִ�й���%Type
  );
  --29.��鱨�泷��������������ʱ
  Procedure Zlhis_Pacs_004
  (
    ҽ��id_In   In Ӱ�����¼.ҽ��id%Type,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  );
  --30.���Σ��ֵ֪ͨ����鷢��Σ��ֵʱ
  Procedure Zlhis_Pacs_005(ҽ��id_In In Ӱ�����¼.ҽ��id%Type);
  -- ���ԤԼ֪ͨ�����ԤԼʱ
  Procedure Zlhis_Pacs_006
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԤԼid_In In Ris���ԤԼ.ԤԼid%Type
  );
  -- ȡ�����ԤԼ��ȡ��ԤԼʱ
  Procedure Zlhis_Pacs_007
  (
    ҽ��id_In       In Ӱ�����¼.ҽ��id%Type,
    ԤԼid_In       In Ris���ԤԼ.ԤԼid%Type,
    ԤԼ����_In     In Ris���ԤԼ.ԤԼ����%Type,
    ԤԼ���_In     In Ris���ԤԼ.���%Type,
    ����豸����_In In Ris���ԤԼ.����豸����%Type
  );


  --36.���߷���
  Procedure Zlhis_Patient_018
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ����_In     In ����ҽ�ƿ���Ϣ.����%Type
  );

  --37.�����˿�
  Procedure Zlhis_Patient_019
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ����_In     In ����ҽ�ƿ���Ϣ.����%Type
  );

  --38.�����˿�
  Procedure Zlhis_Patient_020
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ԭ����_In   In ����ҽ�ƿ���Ϣ.����%Type,
    �¿���_In   In ����ҽ�ƿ���Ϣ.����%Type
  );

  --39.���˹ҺŵǼǣ�����ԤԼ�Ǽ�)
  Procedure Zlhis_Regist_001
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type
  );

  --40.���˷���
  Procedure Zlhis_Regist_002
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type,
    ����_In   In ���˹Һż�¼.����%Type
  );

  --41.�����˺�
  Procedure Zlhis_Regist_003
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type
  );

  --42.�ٴ����ﰲ�ŵ���
  Procedure Zlhis_Regist_004
  (
    �䶯ԭ��_In In Integer, --1-ͣ��;2-����;3-���ұ䶯
    ��¼id_In   In �ٴ������¼.Id%Type,
    �䶯id_In   In �ٴ�����䶯��¼.Id%Type
  );

  --43.���ﻼ�߹ҺŻ��Ų���
  Procedure Zlhis_Regist_005
  (
    No_In         In ���˹Һż�¼.No%Type,
    �䶯ԭ��_In   Integer, --1-����;2-����;3-ԤԼ���ڱ䶯,
    ����䶯id_In ����䶯��¼.Id%Type
  );


  --���������շѼ��������
  --��������_In:1-�շѽ��㣬2-�������
  Procedure Zlhis_Charge_002
  (
    ��������_In In Number,
    ����id_In   In ������ü�¼.����id%Type
  );


  --46.�����˷ѵ���
  Procedure Zlhis_Charge_004
  (
    �˷�����_In In Number,
    ����id_In   In ������ü�¼.����id%Type
  );

  --47.��Ԥ����
  Procedure Zlhis_Charge_005
  (
    Ԥ��id_In In ����Ԥ����¼.Id%Type,
    ���ݺ�_In In ����Ԥ����¼.No%Type
  );

  --48.��Ԥ����(����������Ԥ�����)
  Procedure Zlhis_Charge_006
  (
    ��Ԥ��id_In In ����Ԥ����¼.Id%Type,
    ���ݺ�_In   In ����Ԥ����¼.No%Type
  );

  --סԺ���ʵ���
  Procedure Zlhis_Charge_007
  (
    �շ����_In In סԺ���ü�¼.�շ����%Type,
    ����id_In   In סԺ���ü�¼.Id%Type
  );

  --סԺ���ʵ�������
  Procedure Zlhis_Charge_008
  (
    �շ����_In In סԺ���ü�¼.�շ����%Type,
    ����id_In   In סԺ���ü�¼.Id%Type,
    �շ�ids_In  In Varchar2 := Null --���ܷ���ID��Ӧ����շ�id����Ӧ��ʽ���շ�id,����|�շ�id,��������ҩƷ����
  );

  --53.סԺ������Ժ�Ǽ�
  Procedure Zlhis_Patient_001
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --54.סԺ������Ժ���
  Procedure Zlhis_Patient_002
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --56.סԺ���ߴ�λ���
  Procedure Zlhis_Patient_004
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --57.סԺ���߲�����
  Procedure Zlhis_Patient_005
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --58.סԺ���߱������
  Procedure Zlhis_Patient_006
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    ������ʽ_In In Varchar2
  );
  --59.סԺ����ҽ�����
  Procedure Zlhis_Patient_007
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --סԺ���߻���ȼ����
  Procedure Zlhis_Patient_008
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --60.סԺ����Ԥ��Ժ
  Procedure Zlhis_Patient_009
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --61.סԺ���߳�Ժ
  Procedure Zlhis_Patient_010
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --62.סԺ�����������Ǽ�
  Procedure Zlhis_Patient_011
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    Ӥ�����_In ����ҽ����¼.Ӥ��%Type
  );
  --63.סԺ����ת�����
  Procedure Zlhis_Patient_012
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --64.�������Ǽ�����
  Procedure Zlhis_Patient_013
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    Ӥ�����_In ����ҽ����¼.Ӥ��%Type
  );
  --65.���ﻼ�ߵǼ�
  Procedure Zlhis_Patient_015(����id_In In ������ҳ.����id%Type);
  --66.������Ϣ�޸�
  Procedure Zlhis_Patient_016(����id_In In ������ҳ.����id%Type);

  --67.���ߺϲ�
  Procedure Zlhis_Patient_017
  (
    ����id_In   In ������ҳ.����id%Type,
    ԭ����id_In In ������ҳ.����id%Type
  );

  --69.����ת����ת��
  Procedure Zlhis_Patient_026
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );

  Procedure Zlhis_Patient_028(����id_In In ������ҳ.����id%Type);

  --Ѫ��:������Ѫ���
  Procedure Zlhis_Blood_001(ҽ��id_In In ����ҽ����¼.Id%Type);
  --Ѫ��:������Ѫ�ܾ�
  Procedure Zlhis_Blood_002(ҽ��id_In In ����ҽ����¼.Id%Type);

  --70.����걾���
  Procedure Zlhis_Lis_001(�걾id_In In ����걾��¼.Id%Type);
  --71.����걾��˳���
  Procedure Zlhis_Lis_002(�걾id_In In ����걾��¼.Id%Type);
  --73.����걾�����ӡ
  Procedure Zlhis_Lis_004
  (
    ��������_In In ����ҽ������.��������%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type,
    ҽ��ids_In  In Varchar2
  );
  --74.����걾�����ӡ����
  Procedure Zlhis_Lis_005
  (
    ��������_In In ����ҽ������.��������%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type,
    ҽ��ids_In  In Varchar2
  );
  --75.����걾����
  Procedure Zlhis_Lis_006(�걾id_In In ����걾��¼.Id%Type);
  --76.����걾���ճ���
  Procedure Zlhis_Lis_007(�걾id_In In ����걾��¼.Id%Type);
  --77.����걾����
  Procedure Zlhis_Lis_008(�걾id_In In ����걾��¼.Id%Type);
End b_Message;
/
CREATE OR REPLACE Package Body b_Message Is
  --�Ƿ���ƽ̨����
  Is_Platform_Call Number(1) := 0;
  --��Ϣ��������
  Message_Creator Zlmsg_Todo.Creator%Type := Null;
  --������Ϣ��ѯ���
  Type Tmap_Msg_Using Is Table Of Number(1) Index By Varchar2(30);
  Zlmsg_Map Tmap_Msg_Using;
  --��Ϣ�Ƿ�����
  Function p_Msg_Using(Msg_Code_In Zlmsg_Lists.Code%Type) Return Number As
    n_Using Zlmsg_Lists.Using%Type;
    v_Code  Zlmsg_Lists.Code%Type;
  Begin
    If Is_Platform_Call = 1 Then
      Return 0;
    End If;
    v_Code := Upper(Msg_Code_In);
    Begin
      n_Using := Zlmsg_Map(v_Code);
      Return n_Using;
    Exception
      When No_Data_Found Then
        --����ȡMax�ݴ��������൱�����,�û�����û�в�ȡͬ���޸Ļ��Լ���������Ϣ���͵���δע�ᵽZlmsg_Lists���������������ִ���


        Select Nvl(Using, 0) Into n_Using From Zlmsg_Lists A Where Code = v_Code;
        Zlmsg_Map(v_Code) := n_Using;
        --��ѯ������Ϣ����Ա�������������ִ�д���
        If Message_Creator Is Null Then
          Message_Creator := Zl_Username;
        End If;
        Return n_Using;
    End;
  Exception
    When Others Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || 'δ��Zlmsg_Lists���ҵ���Ϣ"' || v_Code || '"������ϵ����Ա���д���' || '[ZLSOFT]');
      Return 0;
  End;
  Procedure p_Msg_Todo_Insert
  (
    Msg_Code_In  Zlmsg_Todo.Msg_Code%Type,
    Key_Value_In Zlmsg_Todo.Key_Value%Type
  ) Is
  Begin
    If p_Msg_Using(Msg_Code_In) = 0 Then
      Return;
    End If;
    Insert Into Zlmsg_Todo
      (ID, Msg_Code, Key_Value, State, Create_Time, Creator)
    Values
      (Zlmsg_Todo_Id.Nextval, Upper(Msg_Code_In), Key_Value_In, 0, Sysdate, Message_Creator);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Msg_Todo_Insert;
  --���õ�ǰ�ỰΪƽ̨����
  Procedure Set_Platform_Call(Platform_Call Number) Is
  Begin
    Is_Platform_Call := Platform_Call;
  End Set_Platform_Call;
  --��ϢZlhis_Dict_001
  Procedure Zlhis_Dict_001(Id_In ���ű�.Id%Type) Is
    v_Define Xmltype;
    v_Value  Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_DICT_001') = 0 Then
      Return;
    End If;
    Begin
      Select Xmltype(Key_Define) Into v_Define From Zlmsg_Lists Where Code = 'ZLHIS_DICT_001';
    Exception
      When Others Then
        v_Define := Xmltype('<root><ID>NULL</ID></root>');
    End;
    Select Updatexml(v_Define, '/root/ID/text()', Id_In).Getstringval() Into v_Value From Dual;
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_001', v_Value);
  End Zlhis_Dict_001;
  --�޸Ĳ���
  Procedure Zlhis_Dict_002(����id_In ���ű�.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_002', v_Value);
  End Zlhis_Dict_002;
  --ͣ�ò���
  Procedure Zlhis_Dict_003(����id_In ���ű�.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_003', v_Value);
  End Zlhis_Dict_003;
  --���ò���
  Procedure Zlhis_Dict_004(����id_In ���ű�.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_004', v_Value);
  End Zlhis_Dict_004;
  --������Ա
  Procedure Zlhis_Dict_005(��Աid_In ��Ա��.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_005', v_Value);
  End Zlhis_Dict_005;
  --�޸���Ա
  Procedure Zlhis_Dict_006(��Աid_In ��Ա��.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_006', v_Value);
  End Zlhis_Dict_006;
  --ͣ����Ա
  Procedure Zlhis_Dict_007(��Աid_In ��Ա��.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_007', v_Value);
  End Zlhis_Dict_007;
  --������Ա
  Procedure Zlhis_Dict_008(��Աid_In ��Ա��.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_008', v_Value);
  End Zlhis_Dict_008;
  --�����շ���Ŀ
  Procedure Zlhis_Dict_009(ϸĿid_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_009', v_Value);
  End Zlhis_Dict_009;
  --�޸��շ���Ŀ
  Procedure Zlhis_Dict_010(ϸĿid_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_010', v_Value);
  End Zlhis_Dict_010;
  --ͣ���շ���Ŀ
  Procedure Zlhis_Dict_011(ϸĿid_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_011', v_Value);
  End Zlhis_Dict_011;
  --�����շ���Ŀ
  Procedure Zlhis_Dict_012(ϸĿid_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_012', v_Value);
  End Zlhis_Dict_012;
  --����������Ŀ
  Procedure Zlhis_Dict_013(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_013', v_Value);
  End Zlhis_Dict_013;
  --�޸�������Ŀ
  Procedure Zlhis_Dict_014(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_014', v_Value);
  End Zlhis_Dict_014;
  --ͣ��������Ŀ
  Procedure Zlhis_Dict_015(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_015', v_Value);
  End Zlhis_Dict_015;
  --����������Ŀ
  Procedure Zlhis_Dict_016(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_016', v_Value);
  End Zlhis_Dict_016;
  --����������Ŀ
  Procedure Zlhis_Dict_017(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><ϵͳ>1</ϵͳ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_017', v_Value);
  End Zlhis_Dict_017;
  --�޸ļ�����Ŀ
  Procedure Zlhis_Dict_018(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><ϵͳ>1</ϵͳ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_018', v_Value);
  End Zlhis_Dict_018;
  --ɾ��������Ŀ
  Procedure Zlhis_Dict_019
  (
    ����id_In ������ĿĿ¼.Id%Type,
    ����_In   ����������Ŀ.����%Type,
    ������_In ����������Ŀ.������%Type,
    Ӣ����_In ����������Ŀ.Ӣ����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID>' || '<����>' || ����_In || '</����>' || '<������>' || ������_In || '</������>' ||
               '<Ӣ����>' || Ӣ����_In || '</Ӣ����>' || '<ϵͳ>1</ϵͳ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_019', v_Value);
  End Zlhis_Dict_019;
  --������������Ŀ¼
  Procedure Zlhis_Dict_021(����id_In ��������Ŀ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_021', v_Value);
  End Zlhis_Dict_021;
  --�޸ļ�������Ŀ¼
  Procedure Zlhis_Dict_022(����id_In ��������Ŀ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_022', v_Value);
  End Zlhis_Dict_022;
  --ͣ�ü�������Ŀ¼
  Procedure Zlhis_Dict_023(����id_In ��������Ŀ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_023', v_Value);
  End Zlhis_Dict_023;
  --���ü�������Ŀ¼
  Procedure Zlhis_Dict_024(����id_In ��������Ŀ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_024', v_Value);
  End Zlhis_Dict_024;
  --����ҩƷ����
  Procedure Zlhis_Dict_025
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_025', v_Value);
  End Zlhis_Dict_025;
  --�޸�ҩƷ����
  Procedure Zlhis_Dict_026
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_026', v_Value);
  End Zlhis_Dict_026;
  --ɾ��ҩƷ����
  Procedure Zlhis_Dict_027
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID><����>' || ����_In || '</����><����>' || ����_In  || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_027', v_Value);
  End Zlhis_Dict_027;
  --ͣ��ҩƷ����
  Procedure Zlhis_Dict_028
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_028', v_Value);
  End Zlhis_Dict_028;
  --����ҩƷ����
  Procedure Zlhis_Dict_029
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_029', v_Value);
  End Zlhis_Dict_029;
  --����ҩƷƷ��
  Procedure Zlhis_Dict_030
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_030', v_Value);
  End Zlhis_Dict_030;
  --�޸�ҩƷƷ��
  Procedure Zlhis_Dict_031
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_031', v_Value);
  End Zlhis_Dict_031;
  --ɾ��ҩƷƷ��
  Procedure Zlhis_Dict_032
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type,
    ����_In ������ĿĿ¼.����%Type,
    ����_In ������ĿĿ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID><����>' || ����_In || '</����><����>' || ����_In  || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_032', v_Value);
  End Zlhis_Dict_032;
  --ͣ��ҩƷƷ��
  Procedure Zlhis_Dict_033
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_033', v_Value);
  End Zlhis_Dict_033;
  --����ҩƷƷ��
  Procedure Zlhis_Dict_034
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_034', v_Value);
  End Zlhis_Dict_034;
  --����ҩƷ���
  Procedure Zlhis_Dict_035
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_035', v_Value);
  End Zlhis_Dict_035;
  --�޸�ҩƷ���
  Procedure Zlhis_Dict_036
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_036', v_Value);
  End Zlhis_Dict_036;
  --ɾ��ҩƷ���
  Procedure Zlhis_Dict_037
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type,
    ����_In   �շ���ĿĿ¼.����%Type,
    ����_In   �շ���ĿĿ¼.����%Type,
    ���_In   �շ���ĿĿ¼.���%Type,
    ����_In   �շ���ĿĿ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID><����>' || ����_In || '</����><����>' || ����_In  || '</����><���>' ||
            ���_In || '</���><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_037', v_Value);
  End Zlhis_Dict_037;
  --ͣ��ҩƷ���
  Procedure Zlhis_Dict_038
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_038', v_Value);
  End Zlhis_Dict_038;
  --����ҩƷ���
  Procedure Zlhis_Dict_039
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_039', v_Value);
  End Zlhis_Dict_039;
  --����ҩƷ�洢�ⷿ
  Procedure Zlhis_Dict_040
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_040', v_Value);
  End Zlhis_Dict_040;
  --����ҩƷ�����޶�
  Procedure Zlhis_Dict_041
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_041', v_Value);
  End Zlhis_Dict_041;
  --��������Ʒ��
  Procedure Zlhis_Dict_042(Id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_042', v_Value);
  End Zlhis_Dict_042;
  --�������Ĺ��
  Procedure Zlhis_Dict_043(Id_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_043', v_Value);
  End Zlhis_Dict_043;
  --�޸����Ĺ��
  Procedure Zlhis_Dict_044(Id_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_044', v_Value);
  End Zlhis_Dict_044;
  --ɾ�����Ĺ��
  Procedure Zlhis_Dict_045
  (
   Id_In �շ���ĿĿ¼.Id%Type,
   ����_In   �շ���ĿĿ¼.����%Type,
   ����_In   �շ���ĿĿ¼.����%Type,
   ���_In   �շ���ĿĿ¼.���%Type,
   ����_In   �շ���ĿĿ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID><����>' || ����_In || '</����><����>' || ����_In  || '</����><���>' || ���_In || '</���><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_045', v_Value);
  End Zlhis_Dict_045;
  --ͣ�����Ĺ��
  Procedure Zlhis_Dict_046(Id_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_046', v_Value);
  End Zlhis_Dict_046;
  --�������Ĺ��
  Procedure Zlhis_Dict_047(Id_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_047', v_Value);
  End Zlhis_Dict_047;
  --ҽ������
  Procedure Zlhis_Dict_048
  (
    ����_In       In ����֧����Ŀ.����%Type,
    �շ�ϸĿid_In In ����֧����Ŀ.�շ�ϸĿid%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><�շ�ϸĿID>' || �շ�ϸĿid_In || '</�շ�ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_048', v_Value);
  End Zlhis_Dict_048;
  --ɾ��ҽ������
  Procedure Zlhis_Dict_049
  (
    ����_In       In ����֧����Ŀ.����%Type,
    �շ�ϸĿid_In In ����֧����Ŀ.�շ�ϸĿid%Type,
    ��Ŀ����_In   In �շ���ĿĿ¼.����%Type,
    ��Ŀ����_In   In �շ���ĿĿ¼.����%Type,
    ҽ������_In   In ����֧����Ŀ.��Ŀ����%Type,
    ҽ������_In   In ����֧����Ŀ.��Ŀ����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><�շ�ϸĿID>' || �շ�ϸĿid_In || '</�շ�ϸĿID><��Ŀ����>' || ��Ŀ����_In || '</��Ŀ����><��Ŀ����>' ||
               ��Ŀ����_In || '</��Ŀ����><ҽ������>' || ҽ������_In || '</ҽ������><ҽ������>' || ҽ������_In || '</ҽ������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_049', v_Value);
  End Zlhis_Dict_049;
  --�������ķ���
  Procedure Zlhis_Dict_050
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ID_In ���Ʒ���Ŀ¼.ID%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || ID_In ||  '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_050', v_Value);
  End Zlhis_Dict_050;
  --�޸����ķ���
  Procedure ZLHIS_DICT_051
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ID_In ���Ʒ���Ŀ¼.ID%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || ID_In ||  '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_051', v_Value);
  End ZLHIS_DICT_051;
  --ɾ�����ķ���
  Procedure ZLHIS_DICT_052
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ID_In ���Ʒ���Ŀ¼.ID%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID><����>' || ����_In || '</����><����>' || ����_In  || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_052', v_Value);
  End ZLHIS_DICT_052;
  --�շѼ�Ŀ�䶯
  Procedure Zlhis_Dict_053
  (
    Id_In       �շ���ĿĿ¼.Id%Type,
    �䶯����_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><�䶯����>' || �䶯����_In || '</�䶯����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_053', v_Value);
  End Zlhis_Dict_053;

  --�����շѶ��ձ䶯
  Procedure Zlhis_Dict_054
  (
    Id_In     ���Ʒ���Ŀ¼.Id%Type,
    ԭ����_In Varchar2,
    �ֶ���_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><ԭ����>' || ԭ����_In || '</ԭ����><�ֶ���>' || �ֶ���_In || '</�ֶ���></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_054', v_Value);
  End Zlhis_Dict_054;
  --�������Ƽ������
  Procedure Zlhis_Dictpacs_001
  (
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ������_In ���Ƽ������.������%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '<����/><������>' || ������_In ||
               '</������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_001', v_Value);
  End Zlhis_Dictpacs_001;

  --�޸����Ƽ������
  Procedure Zlhis_Dictpacs_002
  (
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ������_In ���Ƽ������.������%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '<����/><������>' || ������_In ||
               '</������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_002', v_Value);
  End Zlhis_Dictpacs_002;
  --ɾ�����Ƽ������
  Procedure Zlhis_Dictpacs_003
  (
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ����_In   ���Ƽ������.����%Type,
    ������_In ���Ƽ������.������%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '<����/><������>' || ������_In ||
               '</������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_003', v_Value);
  End Zlhis_Dictpacs_003;
  --�������Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_004
  (
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ��ע_In     ���Ƽ�鲿λ.��ע%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    �����Ա�_In ���Ƽ�鲿λ.�����Ա�%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In ||
               '</����><��ע>' || ��ע_In || '</��ע><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In || '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_004', v_Value);
  End Zlhis_Dictpacs_004;
  --�޸����Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_005
  (
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ��ע_In     ���Ƽ�鲿λ.��ע%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    �����Ա�_In ���Ƽ�鲿λ.�����Ա�%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In ||
               '</����><��ע>' || ��ע_In || '</��ע><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In || '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_005', v_Value);
  End Zlhis_Dictpacs_005;
  --ɾ�����Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_006
  (
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    ��ע_In     ���Ƽ�鲿λ.��ע%Type,
    ����_In     ���Ƽ�鲿λ.����%Type,
    �����Ա�_In ���Ƽ�鲿λ.�����Ա�%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In ||
               '</����><��ע>' || ��ע_In || '</��ע><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In || '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_006', v_Value);
  End Zlhis_Dictpacs_006;
  --����������Ŀ��λ
  Procedure Zlhis_Dictpacs_007
  (
    Id_In     ������Ŀ��λ.Id%Type,
    ��Ŀid_In ������Ŀ��λ.��Ŀid%Type,
    ����_In   ������Ŀ��λ.����%Type,
    ��λ_In   ������Ŀ��λ.��λ%Type,
    ����_In   ������Ŀ��λ.����%Type,
    Ĭ��_In   ������Ŀ��λ.Ĭ��%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><��ĿID>' || ��Ŀid_In || '</��ĿID><����>' || ����_In || '</����><��λ>' || ��λ_In ||
               '</��λ><����>' || ����_In || '</����><Ĭ��>' || Ĭ��_In || '</Ĭ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_007', v_Value);
  End Zlhis_Dictpacs_007;
  --�޸�������Ŀ��λ
  Procedure Zlhis_Dictpacs_008
  (
    Id_In     ������Ŀ��λ.Id%Type,
    ��Ŀid_In ������Ŀ��λ.��Ŀid%Type,
    ����_In   ������Ŀ��λ.����%Type,
    ��λ_In   ������Ŀ��λ.��λ%Type,
    ����_In   ������Ŀ��λ.����%Type,
    Ĭ��_In   ������Ŀ��λ.Ĭ��%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><��ĿID>' || ��Ŀid_In || '</��ĿID><����>' || ����_In || '</����><��λ>' || ��λ_In ||
               '</��λ><����>' || ����_In || '</����><Ĭ��>' || Ĭ��_In || '</Ĭ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_008', v_Value);
  End Zlhis_Dictpacs_008;
  --ɾ��������Ŀ��λ
  Procedure Zlhis_Dictpacs_009
  (
    Id_In     ������Ŀ��λ.Id%Type,
    ��Ŀid_In ������Ŀ��λ.��Ŀid%Type,
    ����_In   ������Ŀ��λ.����%Type,
    ��λ_In   ������Ŀ��λ.��λ%Type,
    ����_In   ������Ŀ��λ.����%Type,
    Ĭ��_In   ������Ŀ��λ.Ĭ��%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><��ĿID>' || ��Ŀid_In || '</��ĿID><����>' || ����_In || '</����><��λ>' || ��λ_In ||
               '</��λ><����>' || ����_In || '</����><Ĭ��>' || Ĭ��_In || '</Ĭ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_009', v_Value);
  End Zlhis_Dictpacs_009;
    --����������Ŀ��λ
  Procedure Zlhis_DictLis_004
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In ���Ƽ���걾.����%Type,
    ����_In   ���Ƽ���걾.����%Type,
    �����Ա�_In   ���Ƽ���걾.�����Ա�%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In ||
               '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_004', v_Value);
  End Zlhis_DictLis_004;
      --�޸�������Ŀ��λ
  Procedure Zlhis_DictLis_005
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In ���Ƽ���걾.����%Type,
    ����_In   ���Ƽ���걾.����%Type,
    �����Ա�_In   ���Ƽ���걾.�����Ա�%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In ||
               '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_005', v_Value);
  End Zlhis_DictLis_005;
      --ɾ��������Ŀ��λ
  Procedure Zlhis_DictLis_006
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In ���Ƽ���걾.����%Type,
    ����_In   ���Ƽ���걾.����%Type,
    �����Ա�_In   ���Ƽ���걾.�����Ա�%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In ||
               '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_006', v_Value);
  End Zlhis_DictLis_006;
   --������Ѫ������
  Procedure Zlhis_DictLis_007
  (
    ����_In     ��Ѫ������.����%Type,
    ����_In ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In   ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In   ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����ID_In   ��Ѫ������.����ID%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><��Ӽ�>' || ��Ӽ�_In ||
               '</��Ӽ�><��Ѫ��>' || ��Ѫ��_In || '</��Ѫ��><��ɫ>' || ��ɫ_In || '</��ɫ><���>' || ���_In || '</���><����ID_In>' || ����ID_In || '</����ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_007', v_Value);
  End Zlhis_DictLis_007;
    --������Ѫ������
  Procedure Zlhis_DictLis_008
  (
    ����_In     ��Ѫ������.����%Type,
    ����_In ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In   ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In   ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����ID_In   ��Ѫ������.����ID%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><��Ӽ�>' || ��Ӽ�_In ||
               '</��Ӽ�><��Ѫ��>' || ��Ѫ��_In || '</��Ѫ��><��ɫ>' || ��ɫ_In || '</��ɫ><���>' || ���_In || '</���><����ID_In>' || ����ID_In || '</����ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_008', v_Value);
  End Zlhis_DictLis_008;
     --������Ѫ������
  Procedure Zlhis_DictLis_009
  (
    ����_In     ��Ѫ������.����%Type,
    ����_In ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In   ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In   ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����ID_In   ��Ѫ������.����ID%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><��Ӽ�>' || ��Ӽ�_In ||
               '</��Ӽ�><��Ѫ��>' || ��Ѫ��_In || '</��Ѫ��><��ɫ>' || ��ɫ_In || '</��ɫ><���>' || ���_In || '</���><����ID_In>' || ����ID_In || '</����ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_009', v_Value);
  End Zlhis_DictLis_009;  
  --ҩƷ��ҩ����
  Procedure Zlhis_Drug_001(No_In ҩƷ�շ���¼.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_001', v_Value);
  End Zlhis_Drug_001;
  --ȡ��ҩƷ��ҩ����
  Procedure Zlhis_Drug_002(No_In ҩƷ�շ���¼.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_002', v_Value);
  End Zlhis_Drug_002;
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_003(No_In ҩƷ�շ���¼.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_003', v_Value);
  End Zlhis_Drug_003;
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_004(No_In ҩƷ�շ���¼.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_004', v_Value);
  End Zlhis_Drug_004;
  --���ŷ�ҩ
  Procedure Zlhis_Drug_005
  (
    �ⷿid_In ҩƷ�շ���¼.�ⷿid%Type,
    �շ�id_In ҩƷ�շ���¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�ⷿID>' || �ⷿid_In || '</�ⷿID><�շ�ID>' || �շ�id_In || '</�շ�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_005', v_Value);
  End Zlhis_Drug_005;
  --������ҩ
  Procedure Zlhis_Drug_006
  (
    �����շ�id_In ҩƷ�շ���¼.Id%Type,
    �����շ�id_In ҩƷ�շ���¼.Id%Type,
    ����_In       ҩƷ�շ���¼.ʵ������%Type,
    ����id_In     ������ü�¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><������¼ID>' || �����շ�id_In || '</������¼ID><������¼ID>' || �����շ�id_In || '</������¼ID><����>' || ����_In ||
               '</����><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_006', v_Value);
  End Zlhis_Drug_006;
  --ҩƷ����
  Procedure ZLHIS_DRUG_007
  (
    �۸�ID_In ҩƷ�۸��¼.ID%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�ID_In ||  '</�۸�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_007', v_Value);
  End ZLHIS_DRUG_007;
  --���䷢��
  Procedure ZLHIS_DRUG_008
  (
    ��¼Ids_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
    n_��¼id ��Һ��ҩ��¼.ID%Type;
    v_Tmp    varchar2(4000);
  Begin
    If ��¼Ids_In Is Null Then
      v_Tmp := Null;
    Else
      v_Tmp := ��¼Ids_In || ',';
    End If;

    v_Value := '<root><��¼IDS>';

    While v_Tmp Is Not Null Loop
      --�ֽⵥ��ID��
      n_��¼id :=to_number(Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1));
      v_Tmp    := Replace(',' || v_Tmp, ',' || n_��¼id || ',');

      v_Value:=v_Value || '<��¼ID>' || n_��¼id || '</��¼ID>';
    End Loop;

    v_Value:=v_Value || '</��¼IDS></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);
  End ZLHIS_DRUG_008;
  --ҩƷ���ۼ�
  Procedure ZLHIS_DRUG_009
  (
    �۸�ID_In ҩƷ�۸��¼.ID%Type,
    ʱ��_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�ID_In ||  '</�۸�ID><ʱ��>' || ʱ��_In || '</ʱ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_009', v_Value);
  End ZLHIS_DRUG_009;
  --���ĵ��ɱ���
  Procedure ZLHIS_DRUG_010
  (
    �۸�ID_In �ɱ��۵�����Ϣ.ID%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�ID_In ||  '</�۸�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_010', v_Value);
  End ZLHIS_DRUG_010;
  --���ĵ��ۼ�
  Procedure ZLHIS_DRUG_011
  (
    �۸�ID_In �շѼ�Ŀ.ID%Type,
    ʱ��_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�ID_In ||  '</�۸�ID><ʱ��>' || ʱ��_In || '</ʱ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_011', v_Value);
  End ZLHIS_DRUG_011;

  --2.ֹͣ����ҽ����סԺ
  Procedure Zlhis_Cis_002
  (
    ����id_In  In ����ҽ����¼.����id%Type,
    ��ҳid_In  In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In  In ����ҽ����¼.Id%Type,
    ҽ��ids_In In Varchar2
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_CIS_002') = 0 Then
      Return;
    End If;
    If ҽ��id_In Is Not Null Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002',
                                  '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                   '</ID></root>');
    Else
      For R In (Select '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ID || '</ID></root>' As Xml_Value
                From ����ҽ����¼
                Where ID In (Select Column_Value From Table(f_Num2list(ҽ��ids_In))) And ���id Is Null) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Cis_002;
  --3.���ϻ���ҽ��������/סԺ
  Procedure Zlhis_Cis_003
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_003', v_Value);
  End Zlhis_Cis_003;

  --4.��������ҽ����סԺ
  Procedure Zlhis_Cis_004
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_004',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_004;

  --5.������������ҽ����סԺ
  Procedure Zlhis_Cis_005
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_005',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_005;

  --6.���߻�����ҽ����סԺ
  Procedure Zlhis_Cis_006
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_006',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In ||
                                 '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_006;

  --7.�������߻�����ҽ����סԺ
  Procedure Zlhis_Cis_007
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_007', v_Value);
  End Zlhis_Cis_007;

  --10.�´ﻼ����ϣ�����/סԺ
  Procedure Zlhis_Cis_010
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ���id_In In ������ϼ�¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_010',
                                '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ID>' || ���id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_010;
  --11.�����������
  Procedure Zlhis_Cis_011
  (
    ����id_In   In ���˹Һż�¼.����id%Type,
    ����id_In   In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    Id_In       In ������ϼ�¼.Id%Type,
    ����id_In   In ������ϼ�¼.����id%Type,
    ���id_In   In ������ϼ�¼.���id%Type,
    �������_In In ������ϼ�¼.�������%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ID>' || Id_In || '</ID><����ID>' ||
               ����id_In || '</����ID><���ID>' || ���id_In || '</���ID><�������>' || �������_In || '</�������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_011', v_Value);
  End Zlhis_Cis_011;

  --����ִ��ҽ��У��
  Procedure Zlhis_Cis_012
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_012',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_012;

  --13.����Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_014
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In ����ҽ����¼.Id%Type,
    ��Ϣid_In In ҵ����Ϣ�嵥.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ID>' ||
               ��Ϣid_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_014', v_Value);
  End Zlhis_Cis_014;
  --15.���߼������룬����/סԺ
  Procedure Zlhis_Cis_016
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><������Դ>' || ������Դ_In || '</������Դ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_016', v_Value);
  End Zlhis_Cis_016;
  --16.���߼�����룬����/סԺ
  Procedure Zlhis_Cis_017
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
    v_�������� ������ĿĿ¼.��������%Type;
  Begin
    Select MAX(A.��������) Into v_�������� From ������ĿĿ¼ A,����ҽ����¼ B Where B.������ĿID = a.ID And B.ID = ҽ��id_In;
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><������Դ>' || ������Դ_In || '</������Դ></root>';
    If v_�������� = '����' Then
       b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_054', v_Value);
    Else
       b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_017', v_Value);
    End If;
  End Zlhis_Cis_017;
  --17.�����������룬����/סԺ
  Procedure Zlhis_Cis_018
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_018', v_Value);
  End Zlhis_Cis_018;
  --18.������Ѫ���룬סԺ
  Procedure Zlhis_Cis_019
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_019', v_Value);
  End Zlhis_Cis_019;
  --19.���߻������룬סԺ
  Procedure Zlhis_Cis_020
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_020', v_Value);
  End Zlhis_Cis_020;
  --20.��������ҽ����סԺ
  Procedure Zlhis_Cis_021
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_021', v_Value);
  End Zlhis_Cis_021;
  --21.��������ҽ����סԺ
  Procedure Zlhis_Cis_022
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_022', v_Value);
  End Zlhis_Cis_022;
  --22.������������ҽ����סԺ
  Procedure Zlhis_Cis_023
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_023', v_Value);
  End Zlhis_Cis_023;

  --24.���Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_025
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In ����ҽ����¼.Id%Type,
    ��Ϣid_In In ҵ����Ϣ�嵥.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ID>' ||
               ��Ϣid_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_025', v_Value);
  End Zlhis_Cis_025;

  --����ִ��ҽ������
  Procedure Zlhis_Cis_026
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_026',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In ||
                                 '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_026;

  --�������߼�������
  Procedure Zlhis_Cis_036
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO><������Դ>' || ������Դ_In ||
               '</������Դ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_036', v_Value);
  End Zlhis_Cis_036;

  --�������߼������
  Procedure Zlhis_Cis_037
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
    v_�������� ������ĿĿ¼.��������%Type;
  Begin
    Select MAX(A.��������) Into v_�������� From ������ĿĿ¼ A,����ҽ����¼ B Where B.������ĿID = a.ID And B.ID = ҽ��id_In;
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO><������Դ>' || ������Դ_In ||
               '</������Դ></root>';
    If v_�������� = '����' Then
       b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_055', v_Value);
    Else
       b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_037', v_Value);
    End If;
  End Zlhis_Cis_037;

  --����������������
  Procedure Zlhis_Cis_038
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_038', v_Value);
  End Zlhis_Cis_038;

  --����������Ѫ����
  Procedure Zlhis_Cis_039
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_039', v_Value);
  End Zlhis_Cis_039;

  --�������߻�������
  Procedure Zlhis_Cis_040
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_040', v_Value);
  End Zlhis_Cis_040;

  --������������ҽ��
  Procedure Zlhis_Cis_041
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_041', v_Value);
  End Zlhis_Cis_041;

  --������������ҽ��
  Procedure Zlhis_Cis_042
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_042', v_Value);
  End Zlhis_Cis_042;

  --������������ҽ��
  Procedure Zlhis_Cis_043
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_043', v_Value);
  End Zlhis_Cis_043;

  --��������ִ��ҽ��
  Procedure Zlhis_Cis_044
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ��������_In In ����ҽ������.��������%Type,
    �״�ʱ��_In In ����ҽ������.�״�ʱ��%Type,
    ĩ��ʱ��_In In ����ҽ������.ĩ��ʱ��%Type,
    ��������_In In ����ҽ������.��������%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO><��������>' || ��������_In || '</��������><�״�ʱ��>' ||
               To_Char(�״�ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</�״�ʱ��><ĩ��ʱ��>' ||
               To_Char(ĩ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</ĩ��ʱ��><��������>' || ��������_In || '</��������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_044', v_Value);
  End Zlhis_Cis_044;

  --����ҽ��ִ�еǼ�
  Procedure Zlhis_Cis_050
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    Ҫ��ʱ��_In In ����ҽ��ִ��.Ҫ��ʱ��%Type,
    ִ��ʱ��_In In ����ҽ��ִ��.ִ��ʱ��%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><Ҫ��ʱ��>' || To_Char(Ҫ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</Ҫ��ʱ��><ִ��ʱ��>' || To_Char(ִ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</ִ��ʱ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_050', v_Value);
  End Zlhis_Cis_050;

  --����ҽ��ȡ��ִ�еǼ�
  Procedure Zlhis_Cis_051
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    Ҫ��ʱ��_In In ����ҽ��ִ��.Ҫ��ʱ��%Type,
    ִ��ʱ��_In In ����ҽ��ִ��.ִ��ʱ��%Type,
    ��������_In In ����ҽ��ִ��.��������%Type,
    ִ�н��_In In ����ҽ��ִ��.ִ�н��%Type,
    ִ��ժҪ_In In ����ҽ��ִ��.ִ��ժҪ%Type,
    ִ�п���_In In ����ҽ��ִ��.ִ�п���id%Type,
    ִ����_In   In ����ҽ��ִ��.ִ����%Type,
    �˶���_In   In ����ҽ��ִ��.�˶���%Type,
    ��¼��Դ_In In ����ҽ��ִ��.��¼��Դ%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><Ҫ��ʱ��>' || To_Char(Ҫ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</Ҫ��ʱ��><ִ��ʱ��>' || To_Char(ִ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</ִ��ʱ��><��������>' || ��������_In ||
               '</��������><ִ�н��>' || ִ�н��_In || '</ִ�н��><ִ��ժҪ>' || ִ��ժҪ_In || '</ִ��ժҪ><ִ�п���ID>' || ִ�п���_In ||
               '</ִ�п���ID><ִ����>' || ִ����_In || '</ִ����><�˶���>' || �˶���_In || '</�˶���><��¼��Դ>' || ��¼��Դ_In || '</��¼��Դ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_051', v_Value);
  End Zlhis_Cis_051;
  --����ҽ��ִ�����
  Procedure Zlhis_Cis_052
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_052',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In ||
                                 '</�Һŵ�><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_052;
  --����ҽ������ִ�����
  Procedure Zlhis_Cis_053
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_053',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In ||
                                 '</�Һŵ�><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_053;

  --�������뷢�ͺ��޸�
  Procedure Zlhis_Cis_056
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
    v_�������� ������ĿĿ¼.��������%Type;
  Begin
    Select MAX(A.��������) Into v_�������� From ������ĿĿ¼ A,����ҽ����¼ B Where B.������ĿID = a.ID And B.ID = ҽ��id_In;
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><������Դ>' || ������Դ_In || '</������Դ></root>';
     b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_056', v_Value);
  End Zlhis_Cis_056;

  --26.��鱨����ɣ�������ʱ
  Procedure Zlhis_Pacs_001
  (
    ҽ��id_In   In Ӱ�����¼.ҽ��id%Type,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_PACS_001') = 0 Then
      Return;
    End If;
    For R In (Select '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><����ID>' || Column_Value || '</����ID><��������>' || ��������_In ||
                      '<��������></root>' As Xml_Value
              From Table(f_Str2list(����id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_001', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_001;
  --27.���״̬ͬ�������״̬�ı��
  Procedure Zlhis_Pacs_002
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԭ״̬_In In ����ҽ������.ִ�й���%Type,
    ��״̬_In In ����ҽ������.ִ�й���%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԭ״̬>' || ԭ״̬_In || '</ԭ״̬><��״̬>' || ��״̬_In || '</��״̬></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_002', v_Value);
  End Zlhis_Pacs_002;
  --28.���״̬���ˣ����״̬���˺�
  Procedure Zlhis_Pacs_003
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԭ״̬_In In ����ҽ������.ִ�й���%Type,
    ��״̬_In In ����ҽ������.ִ�й���%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԭ״̬>' || ԭ״̬_In || '</ԭ״̬><��״̬>' || ��״̬_In || '</��״̬></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_003', v_Value);
  End Zlhis_Pacs_003;
  --29.��鱨�泷��������������ʱ
  Procedure Zlhis_Pacs_004
  (
    ҽ��id_In   In Ӱ�����¼.ҽ��id%Type,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_PACS_004') = 0 Then
      Return;
    End If;
    For R In (Select '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><����ID>' || Column_Value || '</����ID><��������>' || ��������_In ||
                      '<��������></root>' As Xml_Value
              From Table(f_Str2list(����id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_004', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_004;
  --30.���Σ��ֵ֪ͨ����鷢��Σ��ֵʱ
  Procedure Zlhis_Pacs_005(ҽ��id_In In Ӱ�����¼.ҽ��id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_005', v_Value);
  End Zlhis_Pacs_005;
  -- ���ԤԼ֪ͨ�����ԤԼʱ
  Procedure Zlhis_Pacs_006
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԤԼid_In In Ris���ԤԼ.ԤԼid%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԤԼID>' || ԤԼid_In || '</ԤԼID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_006', v_Value);
  End Zlhis_Pacs_006;
  -- ȡ�����ԤԼ��ȡ��ԤԼʱ
  Procedure Zlhis_Pacs_007
  (
    ҽ��id_In       In Ӱ�����¼.ҽ��id%Type,
    ԤԼid_In       In Ris���ԤԼ.ԤԼid%Type,
    ԤԼ����_In     In Ris���ԤԼ.ԤԼ����%Type,
    ԤԼ���_In     In Ris���ԤԼ.���%Type,
    ����豸����_In In Ris���ԤԼ.����豸����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԤԼID>' || ԤԼid_In || '</ԤԼID><ԤԼ����>' || ԤԼ����_In || '</ԤԼ����><ԤԼ���>' ||
               ԤԼ���_In || '</ԤԼ���><����豸����>' || ����豸����_In || '</����豸����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_007', v_Value);
  End Zlhis_Pacs_007;


  --36.���߷�����󶨿�
  Procedure Zlhis_Patient_018
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ����_In     In ����ҽ�ƿ���Ϣ.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ID>' || �䶯id_In || '</�䶯ID><����ID>' || ����id_In || '</����ID><�����ID>' || �����id_In ||
               '</�����ID><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_018', v_Value);
  End;

  --37.�����˿�
  Procedure Zlhis_Patient_019
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ����_In     In ����ҽ�ƿ���Ϣ.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ID>' || �䶯id_In || '</�䶯ID><����ID>' || ����id_In || '</����ID><�����ID>' || �����id_In ||
               '</�����ID><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_019', v_Value);
  End;

  --38.���߲���/����
  Procedure Zlhis_Patient_020
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ԭ����_In   In ����ҽ�ƿ���Ϣ.����%Type,
    �¿���_In   In ����ҽ�ƿ���Ϣ.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ID>' || �䶯id_In || '</�䶯ID><����ID>' || ����id_In || '</����ID><�����ID>' || �����id_In ||
               '</�����ID><ԭ����>' || ԭ����_In || '</ԭ����><�¿���>' || �¿���_In || '</�¿���></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_020', v_Value);
  End;

  --39.���˹ҺŵǼǣ�����ԤԼ�Ǽ�)
  Procedure Zlhis_Regist_001
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�Һ�ID>' || �Һ�id_In || '</�Һ�ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_001', v_Value);
  End;

  --40.���˷���
  Procedure Zlhis_Regist_002
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type,
    ����_In   In ���˹Һż�¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�Һ�ID>' || �Һ�id_In || '</�Һ�ID><NO>' || No_In || '</NO><����>' || Nvl(����_In, '') || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_002', v_Value);
  End;

  --41.�����˺ţ���ȡ��ԤԼ)
  Procedure Zlhis_Regist_003
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�Һ�ID>' || �Һ�id_In || '</�Һ�ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_003', v_Value);
  End;

  --42.�ٴ����ﰲ�ŵ���
  Procedure Zlhis_Regist_004
  (
    �䶯ԭ��_In In Integer, --1-ͣ��;2-����;3-���ұ䶯
    ��¼id_In   In �ٴ������¼.Id%Type,
    �䶯id_In   In �ٴ�����䶯��¼.Id%Type

  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ԭ��>' || �䶯ԭ��_In || '</�䶯ԭ��><��¼ID>' || ��¼id_In || '</��¼ID><�䶯ID>' || �䶯id_In ||
               '</�䶯ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_004', v_Value);
  End;

  --43.���ﻼ�߹ҺŻ��Ų���
  Procedure Zlhis_Regist_005
  (
    No_In         In ���˹Һż�¼.No%Type,
    �䶯ԭ��_In   Integer, --1-����;2-����;3-ԤԼ���ڱ䶯,
    ����䶯id_In ����䶯��¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><NO>' || No_In || '</NO><�䶯ԭ��>' || �䶯ԭ��_In || '</�䶯ԭ��><����䶯ID>' || ����䶯id_In ||
               '</����䶯ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_005', v_Value);
  End;


  --���������շѼ��������
  Procedure Zlhis_Charge_002
  (
    ��������_In In Number,
    ����id_In   In ������ü�¼.����id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --��������_In:1-�շѽ��㣬2-�������
    v_Value := '<root><��������>' || ��������_In || '</��������><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_002', v_Value);
  End;


  --46.�����˷ѵ���
  Procedure Zlhis_Charge_004
  (
    �˷�����_In In Number,
    ����id_In   In ������ü�¼.����id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --�˷�����_In:1-�շѽ��㣬2-�������
    v_Value := '<root><�˷�����>' || �˷�����_In || '</�˷�����><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_004', v_Value);
  End;

  --47.��Ԥ����
  Procedure Zlhis_Charge_005
  (
    Ԥ��id_In In ����Ԥ����¼.Id%Type,
    ���ݺ�_In In ����Ԥ����¼.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><Ԥ��ID>' || Ԥ��id_In || '</Ԥ��ID><���ݺ�>' || ���ݺ�_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_005', v_Value);
  End;

  --48.��Ԥ����(����������Ԥ�����)
  Procedure Zlhis_Charge_006
  (
    ��Ԥ��id_In In ����Ԥ����¼.Id%Type,
    ���ݺ�_In   In ����Ԥ����¼.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��Ԥ��ID>' || ��Ԥ��id_In || '</��Ԥ��ID><���ݺ�>' || ���ݺ�_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_006', v_Value);
  End;

  --סԺ���ʵ���
  Procedure Zlhis_Charge_007
  (
    �շ����_In In סԺ���ü�¼.�շ����%Type,
    ����id_In   In סԺ���ü�¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�շ����>' || �շ����_In || '</�շ����><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_007', v_Value);
  End;

  --סԺ���ʵ�������
  Procedure Zlhis_Charge_008
  (
    �շ����_In In סԺ���ü�¼.�շ����%Type,
    ����id_In   In סԺ���ü�¼.Id%Type,
    �շ�ids_In  In Varchar2 := Null --���ܷ���ID��Ӧ����շ�id����Ӧ��ʽ���շ�id,����|�շ�id,��������ҩƷ����
  ) Is
    v_Value   Zlmsg_Todo.Key_Value%Type;
    v_Tmp     Varchar2(4000);
    v_Infotmp Varchar2(4000);
    v_Fields  Varchar2(4000);
    v_�շ�id  Varchar2(50);
    v_����    Varchar2(20);
  Begin
    If p_Msg_Using('ZLHIS_CHARGE_008') = 0 Then
      Return;
    End If;
    v_Value := '<root><�շ����>' || �շ����_In || '</�շ����><����ID>' || ����id_In || '</����ID>';

    If �շ�ids_In Is Null Then
      v_Infotmp := Null;
      v_Tmp     := '<�շ�IDS>' || '<�շ�ID>' || '</�շ�ID>' || '<����>' || '</����>' || '</�շ�IDS>';
    Else
      v_Infotmp := �շ�ids_In || '|';
      While v_Infotmp Is Not Null Loop
        --�ֽ��շ�ID��
        v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
        v_�շ�id  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
        v_����    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
        v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');

        v_Tmp := v_Tmp || '<�շ�IDS>' || '<�շ�ID>' || v_�շ�id || '</�շ�ID>' || '<����>' || v_���� || '</����>' || '</�շ�IDS>';
      End Loop;
    End If;

    v_Value := v_Value || v_Tmp || '</root>';

    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_008', v_Value);
  End;

  --53.סԺ������Ժ�Ǽ�
  Procedure Zlhis_Patient_001
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    n_�䶯id ���˱䶯��¼.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_001') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_�䶯id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 1 And Nvl(���Ӵ�λ, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_001',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_001;
  --54.סԺ������Ժ���
  Procedure Zlhis_Patient_002
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    n_�䶯id ���˱䶯��¼.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_002') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_�䶯id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_002',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_002;
  --56.סԺ���ߴ�λ���
  Procedure Zlhis_Patient_004
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    v_ԭ����   Varchar2(255);
    v_�´���   Varchar2(255);
    n_�䶯id   Number(18);
    n_��ʼԭ�� Number(3);
    d_��ʼʱ�� Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_004') = 0 Then
      Return;
    End If;
    Select ID, ����, ��ʼʱ��, ��ʼԭ��
    Into n_�䶯id, v_�´���, d_��ʼʱ��, n_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;

    Select Max(����)
    Into v_ԭ����
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = d_��ʼʱ�� And ��ֹԭ�� = n_��ʼԭ�� And Nvl(���Ӵ�λ, 0) = 0;

    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_004',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ԭ����>' ||
                                 v_ԭ���� || '</ԭ����>' || '<�´���>' || v_�´��� || '</�´���>' || '<�䶯ID>' || n_�䶯id || '</�䶯ID>' ||
                                 '</root>');
  End Zlhis_Patient_004;
  --57.סԺ���߲�����
  Procedure Zlhis_Patient_005
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    n_�䶯id ���˱䶯��¼.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_005') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_�䶯id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_005',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_005;
  --58.סԺ���߱������
  Procedure Zlhis_Patient_006
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    ������ʽ_In In Varchar2
  ) Is
    n_����id     ���˱䶯��¼.����id%Type;
    n_����id     ���˱䶯��¼.����id%Type;
    n_����ȼ�id ���˱䶯��¼.����ȼ�id%Type;
    n_ҽ��С��id ���˱䶯��¼.ҽ��С��id%Type;
    v_����       ���˱䶯��¼.����%Type;
    v_���λ�ʿ   ���˱䶯��¼.���λ�ʿ%Type;
    v_����ҽʦ   ���˱䶯��¼.����ҽʦ%Type;
    v_����ҽʦ   ���˱䶯��¼.����ҽʦ%Type;
    v_����ҽʦ   ���˱䶯��¼.����ҽʦ%Type;
    v_����       ���˱䶯��¼.����%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_006') = 0 Then
      Return;
    End If;
    Select Max(����id), Max(����id), Max(����ȼ�id), Max(ҽ��С��id), Max(����), Max(���λ�ʿ), Max(����ҽʦ), Max(����ҽʦ), Max(����ҽʦ), Max(����)
    Into n_����id, n_����id, n_����ȼ�id, n_ҽ��С��id, v_����, v_���λ�ʿ, v_����ҽʦ, v_����ҽʦ, v_����ҽʦ, v_����
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And (��ֹʱ�� Is Null Or ��ֹԭ�� = 1) And Nvl(���Ӵ�λ, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_006',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><������ʽ>' || ������ʽ_In ||
                                 '</������ʽ><����ID>' || n_����id || '</����ID>' || '<����ID>' || n_����id || '</����ID>' || '<����ȼ�ID>' ||
                                 n_����ȼ�id || '</����ȼ�ID>' || '<ҽ��С��ID>' || n_ҽ��С��id || '</ҽ��С��ID>' || '<����>' || v_���� ||
                                 '</����>' || '<���λ�ʿ>' || v_���λ�ʿ || '</���λ�ʿ>' || '<����ҽʦ>' || v_����ҽʦ || '</����ҽʦ>' ||
                                 '<����ҽʦ>' || v_����ҽʦ || '</����ҽʦ>' || '<����ҽʦ>' || v_����ҽʦ || '</����ҽʦ>' || '<����>' || v_���� ||
                                 '</����>' || '</root>');
  End Zlhis_Patient_006;
  --59.סԺ����ҽ�����
  Procedure Zlhis_Patient_007
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    v_ԭסԺҽ�� Varchar2(100);
    v_��סԺҽ�� Varchar2(100);
    v_ԭ����ҽ�� Varchar2(100);
    v_������ҽ�� Varchar2(100);
    v_ԭ����ҽ�� Varchar2(100);
    v_������ҽ�� Varchar2(100);
    v_ԭ���λ�ʿ Varchar2(100);
    v_�����λ�ʿ Varchar2(100);
    n_�䶯id     Number(18);
    n_��ʼԭ��   Number(3);
    d_��ʼʱ��   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_007') = 0 Then
      Return;
    End If;
    Select ID, ����ҽʦ, ����ҽʦ, ����ҽʦ, ���λ�ʿ, ��ʼʱ��, ��ʼԭ��
    Into n_�䶯id, v_��סԺҽ��, v_������ҽ��, v_������ҽ��, v_�����λ�ʿ, d_��ʼʱ��, n_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;

    Select Max(����ҽʦ), Max(����ҽʦ), Max(����ҽʦ), Max(���λ�ʿ)
    Into v_ԭסԺҽ��, v_ԭ����ҽ��, v_ԭ����ҽ��, v_ԭ���λ�ʿ
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = d_��ʼʱ�� And ��ֹԭ�� = n_��ʼԭ�� And Nvl(���Ӵ�λ, 0) = 0;

    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_007',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ԭסԺҽ��>' ||
                                 v_ԭסԺҽ�� || '</ԭסԺҽ��>' || '<��סԺҽ��>' || v_��סԺҽ�� || '</��סԺҽ��>' || '<ԭ����ҽ��>' || v_ԭ����ҽ�� ||
                                 '</ԭ����ҽ��>' || '<������ҽ��>' || v_������ҽ�� || '</������ҽ��>' || '<ԭ����ҽ��>' || v_ԭ����ҽ�� || '</ԭ����ҽ��>' ||
                                 '<������ҽ��>' || v_������ҽ�� || '</������ҽ��>' || '<ԭ���λ�ʿ>' || v_ԭ���λ�ʿ || '</ԭ���λ�ʿ>' || '<�����λ�ʿ>' ||
                                 v_�����λ�ʿ || '</�����λ�ʿ>' || '<�䶯ID>' || n_�䶯id || '</�䶯ID>' || '</root>');
  End Zlhis_Patient_007;
  --סԺ���߻���ȼ����
  Procedure Zlhis_Patient_008
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    v_ԭ����ȼ�id Number(18);
    v_�»���ȼ�id Number(18);
    n_�䶯id       Number(18);
    n_��ʼԭ��     Number(3);
    d_��ʼʱ��     Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_008') = 0 Then
      Return;
    End If;
    Select ID, ����ȼ�id, ��ʼʱ��, ��ʼԭ��
    Into n_�䶯id, v_�»���ȼ�id, d_��ʼʱ��, n_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;

    Select Max(����ȼ�id)
    Into v_ԭ����ȼ�id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = d_��ʼʱ�� And ��ֹԭ�� = n_��ʼԭ�� And Nvl(���Ӵ�λ, 0) = 0;

    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_008',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ԭ����ȼ�ID>' ||
                                 v_ԭ����ȼ�id || '</ԭ����ȼ�ID>' || '<�»���ȼ�ID>' || v_�»���ȼ�id || '</�»���ȼ�ID>' || '<�䶯ID>' ||
                                 n_�䶯id || '</�䶯ID>' || '</root>');
  End Zlhis_Patient_008;
  --60.סԺ����Ԥ��Ժ
  Procedure Zlhis_Patient_009
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    n_�䶯id ���˱䶯��¼.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_009') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_�䶯id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_009',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_009;
  --61.סԺ���߳�Ժ
  Procedure Zlhis_Patient_010
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_010',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID></root>');
  End Zlhis_Patient_010;
  --62.סԺ�����������Ǽ�
  Procedure Zlhis_Patient_011
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    Ӥ�����_In ����ҽ����¼.Ӥ��%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_011',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><Ӥ�����>' || Ӥ�����_In ||
                                 '</Ӥ�����></root>');
  End Zlhis_Patient_011;
  --63.סԺ����ת�����
  Procedure Zlhis_Patient_012
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    v_ת������id Number(18);
    v_ת�����id Number(18);
    n_�䶯id     Number(18);
    n_��ʼԭ��   Number(3);
    d_��ʼʱ��   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_012') = 0 Then
      Return;
    End If;
    Select ID, ����id, ��ʼʱ��, ��ʼԭ��
    Into n_�䶯id, v_ת�����id, d_��ʼʱ��, n_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;

    Select Max(����id)
    Into v_ת������id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = d_��ʼʱ�� And ��ֹԭ�� = n_��ʼԭ�� And Nvl(���Ӵ�λ, 0) = 0;

    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_012',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ת������ID>' ||
                                 v_ת������id || '</ת������ID>' || '<ת�����ID>' || v_ת�����id || '</ת�����ID>' || '<�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID>' || '</root>');
  End Zlhis_Patient_012;
  --64.�������Ǽ�����
  Procedure Zlhis_Patient_013
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    Ӥ�����_In ����ҽ����¼.Ӥ��%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_013',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><Ӥ�����>' || Ӥ�����_In ||
                                 '</Ӥ�����></root>');
  End Zlhis_Patient_013;
  --65.���ﻼ�ߵǼ�
  Procedure Zlhis_Patient_015(����id_In In ������ҳ.����id%Type) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_015', '<root><����ID>' || ����id_In || '</����ID></root>');
  End Zlhis_Patient_015;
  --66.������Ϣ�޸�
  Procedure Zlhis_Patient_016(����id_In In ������ҳ.����id%Type) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_016', '<root><����ID>' || ����id_In || '</����ID></root>');
  End Zlhis_Patient_016;

  --67.���ߺϲ�
  Procedure Zlhis_Patient_017
  (
    ����id_In   In ������ҳ.����id%Type,
    ԭ����id_In In ������ҳ.����id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_017',
                                '<root><����ID>' || ����id_In || '</����ID><ԭ����ID>' || ԭ����id_In || '</ԭ����ID></root>');
  End Zlhis_Patient_017;

  --69.סԺ����ת�벡��
  Procedure Zlhis_Patient_026
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    v_ת������id Number(18);
    v_ת�벡��id Number(18);
    n_�䶯id     Number(18);
    n_��ʼԭ��   Number(3);
    d_��ʼʱ��   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_026') = 0 Then
      Return;
    End If;
    Select ID, ����id, ��ʼʱ��, ��ʼԭ��
    Into n_�䶯id, v_ת�벡��id, d_��ʼʱ��, n_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;

    Select Max(����id)
    Into v_ת������id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = d_��ʼʱ�� And ��ֹԭ�� = n_��ʼԭ�� And Nvl(���Ӵ�λ, 0) = 0;

    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_026',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ת������ID>' ||
                                 v_ת������id || '</ת������ID>' || '<ת�벡��ID>' || v_ת�벡��id || '</ת�벡��ID>' || '<�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID>' || '</root>');
  End Zlhis_Patient_026;

  Procedure Zlhis_Patient_028(����id_In In ������ҳ.����id%Type) Is
    v_����     ������Ϣ.����%Type;
    v_�Ա�     ������Ϣ.�Ա�%Type;
    v_����     ������Ϣ.����%Type;
    v_�������� ������Ϣ.��������%Type;
    v_�����   ������Ϣ.�����%Type;
    v_���֤�� ������Ϣ.���֤��%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_028') = 0 Then
      Return;
    End If;
    Select ����, �Ա�, ����, ��������, �����, ���֤��
    Into v_����, v_�Ա�, v_����, v_��������, v_�����, v_���֤��
    From ������Ϣ
    Where ����id = ����id_In;

    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_028',
                                '<root><����ID>' || ����id_In || '</����ID><����>' || v_���� || '</����>' || '<�Ա�>' || v_�Ա� ||
                                 '</�Ա�>' || '<����>' || v_���� || '</����>' || '<��������>' || v_�������� || '</��������>' || '<�����>' ||
                                 v_����� || '</�����>' || '<���֤��>' || v_���֤�� || '</���֤��>' || '</root>');
  End Zlhis_Patient_028;

  --Ѫ��:������Ѫ���
  Procedure Zlhis_Blood_001(ҽ��id_In In ����ҽ����¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_001', v_Value);
    End If;
  End Zlhis_Blood_001;

  --Ѫ��:���Ҿܾ���Ѫ
  Procedure Zlhis_Blood_002(ҽ��id_In In ����ҽ����¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_002', v_Value);
    End If;
  End Zlhis_Blood_002;

  --70.���鱨�����
  Procedure Zlhis_Lis_001(�걾id_In In ����걾��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_001', v_Value);
    End If;
  End Zlhis_Lis_001;
  --71.���鱨����˳���
  Procedure Zlhis_Lis_002(�걾id_In In ����걾��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_002', v_Value);
    End If;
  End Zlhis_Lis_002;
  --73.����걾�����ӡ
  Procedure Zlhis_Lis_004
  (
    ��������_In In ����ҽ������.��������%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type,
    ҽ��ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_LIS_004') = 0 Then
      Return;
    End If;
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004', v_Value);
    Else
      For R In (Select '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ϵͳ>1</ϵͳ></root>' As Xml_Value
                From ����ҽ������
                Where ҽ��id In (Select Column_Value From Table(f_Num2list(ҽ��ids_In)))) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Lis_004;
  --74.����걾�����ӡ����
  Procedure Zlhis_Lis_005
  (
    ��������_In In ����ҽ������.��������%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type,
    ҽ��ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_LIS_005') = 0 Then
      Return;
    End If;
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005', v_Value);
    Else
      For R In (Select '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ϵͳ>1</ϵͳ></root>' As Xml_Value
                From ����ҽ������
                Where ҽ��id In (Select Column_Value From Table(f_Num2list(ҽ��ids_In)))) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Lis_005;
  --75.����걾����
  Procedure Zlhis_Lis_006(�걾id_In In ����걾��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_006', v_Value);
    End If;
  End Zlhis_Lis_006;
  --76.����걾���ճ���
  Procedure Zlhis_Lis_007(�걾id_In In ����걾��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_007', v_Value);
    End If;
  End Zlhis_Lis_007;
  --77.����걾����
  Procedure Zlhis_Lis_008(�걾id_In In ����걾��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_008', v_Value);
    End If;
  End Zlhis_Lis_008;

End b_Message;
/

--123263:������,2018-03-22,����ƽ̨��Ϣê��
--122998:������,2018-03-22,����ƽ̨��Ϣê��
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

  If v_Tmp = 'DELETE' Then
    If v_Tab = '���Ƽ������' Then
      --ɾ����¼
      For R In (Select a.����, a.����, a.����, a.������ From ���Ƽ������ A Where a.���� = v_Code) Loop
        b_Message.Zlhis_Dictpacs_003(r.����, r.����, r.����, r.������);
      End Loop;
    Elsif v_Tab = '���Ƽ���걾' Then
      --ɾ����¼
      For R In (Select a.����, a.����, a.����, a.�����Ա� From ���Ƽ���걾 A Where a.���� = v_Code) Loop
        b_Message.Zlhis_Dictlis_006(r.����, r.����, r.����, r.�����Ա�);
      End Loop;
    End If;
  End If;

  Execute Immediate v_Rulesql;

  If v_Tab = '���Ƽ������' Then
    If v_Tmp = 'INSERT' Or v_Tmp = 'UPDATE' Then
      For R In (Select a.����, a.����, a.����, a.������ From ���Ƽ������ A Where a.���� = v_Code) Loop
        If v_Tmp = 'INSERT' Then
          b_Message.Zlhis_Dictpacs_001(r.����, r.����, r.����, r.������);
        Else
          b_Message.Zlhis_Dictpacs_002(r.����, r.����, r.����, r.������);
        End If;
      End Loop;
    End If;
  Elsif v_Tab = '���Ƽ���걾' Then
    If v_Tmp = 'INSERT' Or v_Tmp = 'UPDATE' Then
      For R In (Select a.����, a.����, a.����, a.�����Ա� From ���Ƽ���걾 A Where a.���� = v_Code) Loop
        If v_Tmp = 'INSERT' Then
          b_Message.Zlhis_Dictlis_004(r.����, r.����, r.����, r.�����Ա�);
        Else
          b_Message.Zlhis_Dictlis_005(r.����, r.����, r.����, r.�����Ա�);
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

--122998:������,2018-03-22,����ƽ̨��Ϣê��
Create Or Replace Procedure Zl_���Ƽ�鲿λ_Edit
(
  ����_In     In Number, --1:����;2:�޸�;3:ɾ��
  ����_In     In ���Ƽ�鲿λ.����%Type,
  ԭ����_In   In ���Ƽ�鲿λ.����%Type,
  �±���_In   In ���Ƽ�鲿λ.����%Type := Null,
  ����_In     In ���Ƽ�鲿λ.����%Type := Null,
  ����_In     In ���Ƽ�鲿λ.����%Type := Null,
  ��ע_In     In ���Ƽ�鲿λ.��ע%Type := Null,
  ����_In     In ���Ƽ�鲿λ.����%Type := Null,
  �����Ա�_In In ���Ƽ�鲿λ.�����Ա�%Type := Null
) Is
  v_ԭ���� ���Ƽ�鲿λ.����%Type := Null;
  e_Notfind Exception;
  v_����   Varchar2(1000);
  v_Fields Varchar2(1000);
  v_Tmp    Varchar2(1000);
  n_Count  Number;
  n_��¼id ������Ŀ��λ.Id%Type;
Begin
  If ����_In = 1 Then
    Insert Into ���Ƽ�鲿λ
      (����, ����, ����, ����, ��ע, ����, �����Ա�)
    Values
      (����_In, �±���_In, ����_In, ����_In, ��ע_In, ����_In, �����Ա�_In);
    b_Message.Zlhis_Dictpacs_004(����_In, �±���_In, ����_In, ����_In, ��ע_In, ����_In, �����Ա�_In);
  Elsif ����_In = 2 Then
    Begin
      Select ���� Into v_ԭ���� From ���Ƽ�鲿λ Where ���� = ԭ����_In And ���� = ����_In;
    Exception
      When Others Then
        Null;
    End;
    If v_ԭ���� Is Null Then
      Raise e_Notfind;
    End If;
    Update ���Ƽ�鲿λ
    Set ���� = �±���_In, ���� = ����_In, ���� = ����_In, ��ע = ��ע_In, ���� = ����_In, �����Ա� = �����Ա�_In
    Where ���� = ԭ����_In And ���� = ����_In;
    b_Message.Zlhis_Dictpacs_005(����_In, �±���_In, ����_In, ����_In, ��ע_In, ����_In, �����Ա�_In);
  
    --�����޸�
    v_���� := ';' || ����_In;
    v_���� := Replace(v_����, ',', Chr(10));
    v_���� := Replace(v_����, Chr(9), ';');
    v_���� := Replace(v_����, ';0', Chr(10));
    v_���� := Replace(v_����, ';1', Chr(10));
    v_���� := Replace(v_����, Chr(10), ';');
    v_���� := Replace(v_����, ';;', ';');
    v_���� := v_���� || ';';
  
    v_���� := Substr(v_����, 2);
  
    --ԭ�еķ����������Ѿ�ɾ���˻�ԭ�еĲ�λ�������Ѿ��ı���
    For r_Used In (Select ID, ��Ŀid, ��λ, ����, ����, Ĭ�� From ������Ŀ��λ Where ��λ = v_ԭ���� And ���� = ����_In) Loop
      If Instr(';' || v_����, ';' || r_Used.���� || ';') = 0 Then
        b_Message.Zlhis_Dictpacs_009(r_Used.Id, r_Used.��Ŀid, r_Used.����, r_Used.��λ, r_Used.����, r_Used.Ĭ��);
        Delete ������Ŀ��λ
        Where ��Ŀid = r_Used.��Ŀid And ��λ = r_Used.��λ And ���� = r_Used.���� And ���� = r_Used.����;
      Else
        Update ������Ŀ��λ
        Set ��λ = ����_In
        Where ��Ŀid = r_Used.��Ŀid And ��λ = r_Used.��λ And ���� = r_Used.���� And ���� = r_Used.����;
        b_Message.Zlhis_Dictpacs_008(r_Used.Id, r_Used.��Ŀid, r_Used.����, r_Used.��λ, r_Used.����, r_Used.Ĭ��);
      End If;
    End Loop;
  
    --ԭ��û�еķ�����������
    v_Tmp := v_����;
    While v_Tmp Is Not Null Loop
      --����ȡÿ����Ŀ
      v_Fields := Substr(v_Tmp, 1, Instr(v_Tmp, ';') - 1);
      v_Tmp    := Substr(v_Tmp, Instr(v_Tmp, ';') + 1);
    
      If v_Fields Is Not Null Then
        For r_Used In (Select Distinct ��Ŀid From ������Ŀ��λ Where ��λ = ����_In And ���� = ����_In) Loop
          Select Count(ID)
          Into n_Count
          From ������Ŀ��λ
          Where ��Ŀid = r_Used.��Ŀid And ��λ = ����_In And ���� = ����_In And ���� = v_Fields;
        
          If n_Count = 0 Then
            Select ������Ŀ��λ_Id.Nextval Into n_��¼id From Dual;
            Insert Into ������Ŀ��λ
              (ID, ��Ŀid, ����, ��λ, ����)
            Values
              (n_��¼id, r_Used.��Ŀid, ����_In, ����_In, v_Fields);
            b_Message.Zlhis_Dictpacs_007(n_��¼id, r_Used.��Ŀid, ����_In, ����_In, v_Fields, Null);
          End If;
        End Loop;
      End If;
    End Loop;
  Elsif ����_In = 3 Then
    For R In (Select a.����, a.����, a.����, a.����, a.��ע, a.����, a.�����Ա�
              From ���Ƽ�鲿λ A
              Where a.���� = ԭ����_In And a.���� = ����_In) Loop
      b_Message.Zlhis_Dictpacs_006(r.����, r.����, r.����, r.����, r.��ע, r.����, r.�����Ա�);
    End Loop;
    Delete ���Ƽ�鲿λ Where ���� = ԭ����_In And ���� = ����_In;
  End If;

Exception
  When e_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]�ò�λ�����ڣ������ѱ������û�ɾ���޸ģ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���Ƽ�鲿λ_Edit;
/

--122998:������,2018-03-22,����ƽ̨��Ϣê��
Create Or Replace Procedure Zl_������Ŀ��λ_Insert
(
  ��Ŀid_In In ������Ŀ��λ.��Ŀid%Type,
  ����_In   In ������Ŀ��λ.����%Type,
  ��λ_In   In ������Ŀ��λ.��λ%Type,
  ����_In   In ������Ŀ��λ.����%Type,
  Ĭ��_In   In ������Ŀ��λ.Ĭ��%Type := Null
) As
  v_Code Varchar2(20); --����
  Err_Notfind Exception;
  n_��¼id ������Ŀ��λ.Id%Type;
Begin
  Select RTrim(����) Into v_Code From ������ĿĿ¼ Where ��� = 'D' And ID = ��Ŀid_In;
  If v_Code Is Null Then
    Raise Err_Notfind;
  End If;
  Select ������Ŀ��λ_Id.Nextval Into n_��¼id From Dual;
  Insert Into ������Ŀ��λ
    (ID, ��Ŀid, ����, ��λ, ����, Ĭ��)
  Values
    (n_��¼id, ��Ŀid_In, ����_In, ��λ_In, ����_In, Ĭ��_In);
  b_Message.Zlhis_Dictpacs_007(n_��¼id, ��Ŀid_In, ����_In, ��λ_In, ����_In, Ĭ��_In);
Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]����Ŀ�����ڣ������ѱ������û�ɾ����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������Ŀ��λ_Insert;
/

--122998:������,2018-03-21,����ƽ̨��Ϣê��
Create Or Replace Procedure Zl_������Ŀ��λ_Delete(��Ŀid_In In ������Ŀ��λ.��Ŀid%Type) As
Begin
  For R In (Select a.Id, a.��Ŀid, a.����, a.��λ, a.����, a.Ĭ�� From ������Ŀ��λ A Where a.��Ŀid = ��Ŀid_In) Loop
    b_Message.Zlhis_Dictpacs_009(r.Id, r.��Ŀid, r.����, r.��λ, r.����, r.Ĭ��);
  End Loop;
  Delete ������Ŀ��λ Where ��Ŀid = ��Ŀid_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������Ŀ��λ_Delete;
/

--123312:���Ʊ�,2018-03-22,�²���ϵͳ����ƽ̨��Ϣ
--123225:���Ʊ�,2018-03-20,����ϵͳ�޸���������
CREATE OR REPLACE Procedure Zl_ҽ�����뵥�ļ�_Edit
( 
  �ļ�id_In ҽ�����뵥�ļ�.�ļ�id%Type, 
  �ļ���_IN ҽ�����뵥�ļ�.�ļ���%Type, 
  ���_In   ҽ�����뵥�ļ�.���%Type, 
  ҽ��ID_In   ҽ�����뵥�ļ�.ҽ��ID%Type
) As 
   n_����ID ����ҽ����¼.����ID%Type;
   n_��ҳid ����ҽ����¼.��ҳid%Type;
   v_�Һŵ� ����ҽ����¼.�Һŵ�%Type;
   n_���ͺ� ����ҽ������.���ͺ�%Type;
   n_��ID  ����ҽ����¼.ID%TYPE;
   n_������Դ ����ҽ����¼.������Դ%Type;
Begin 
  Delete From ҽ�����뵥�ļ� Where �ļ�id = �ļ�id_In And ҽ��ID = ҽ��ID_In And ��� = ���_In;
  If Sql%Rowcount <> 0 And ���_In = 2 Then
    Select Max(a.����id), Max(a.��ҳid), Max(a.�Һŵ�), Max(b.���ͺ�), Max(Nvl(a.���id, a.Id)), Max(a.������Դ)
    Into n_����id, n_��ҳid, v_�Һŵ�, n_���ͺ�, n_��id, n_������Դ
    From ����ҽ����¼ A, ����ҽ������ B
    Where a.Id = b.ҽ��id And a.Id = ҽ��id_In;
    b_Message.Zlhis_Cis_056(n_����ID, n_��ҳid, v_�Һŵ�,n_���ͺ�, n_��ID, n_������Դ);
  End If;
  Insert Into ҽ�����뵥�ļ� 
      (�ļ�id,�ļ���, ҽ��ID, ���) 
  Values 
      (�ļ�id_In,�ļ���_IN, ҽ��ID_In, ���_In); 
Exception 
  When Others Then 
    Zl_Errorcenter(Sqlcode, Sqlerrm); 
End Zl_ҽ�����뵥�ļ�_Edit;
/







------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0003' Where ���=&n_System;
Commit;
