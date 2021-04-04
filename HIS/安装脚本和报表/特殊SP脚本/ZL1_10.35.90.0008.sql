----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--124269:������,2018-04-20,���ڻ��������´�ҽ�����
Alter Table ����ҽ����¼ Add ����ҽ��ID number(18);


------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--124187:����,2018-04-24,����Ѫ��ҵ����Ϣ
insert into ҵ����Ϣ����(����,����,˵��,��������) values('ZLHIS_BLOOD_007','Ѫ��������ʾ','��ʿִ����ɺ���ʾ��ʿվ��ҽ��վ����Ѫ��',7);

--122442:��͢��,2018-04-23,ҽ��Ĭ�϶�λ�����һ��
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1253, 1, 1, 0, 0, 0, 0, 68, 'ҽ�����Ĭ�϶�λ�����һ��', '1', '1',
         '���ò�����ҽ�����Ĭ�϶�λ�����һ��', '0-ҽ�����Ĭ�϶�λ����һ�У�1-ҽ�����Ĭ�϶�λ�����һ��', '', '�����ڻ�ʿ�鿴ҽ���б�������ҽ��', Null
  From Dual;

--124269:������,2018-04-20,���ڻ��������´�ҽ�����
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, -null, -null, -null, -null, -null, -null, -null, 302, '��������´�ҽ���ɻ���������Ҵ���', '0',
         '0', '�����ò��������������µ�ҽ�������ɻ���������ҵĻ�ʿ����У�Ի��߷���',
         '0-��ʾ������,1-��ʾ����', Null, '������ĳЩҽԺ��������´��ҽ��������������', Null
  From Dual;

--124189:����,2018-04-20,����Ѫ��ҵ����Ϣ
insert into ҵ����Ϣ����(����,����,˵��,��������) values('ZLHIS_BLOOD_006','������Ѫ��Ӧ','��ʿִ����Ѫʱ��������Ѫ��Ӧ����ʾ��Ӧҽ��վ',7);


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--124730:��͢��,2018-04-23,�����޸����Ʒ��õ�Ȩ�޿����ڷ���ҽ�����Ƿ����������ִ�еļ۸�
Insert Into zlProgFuncs
  (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ)
  Select &n_System, 1252, '�޸����Ʒ���', 31, '�޸����Ʒ��õĲ���Ȩ�ޡ��и�Ȩ��ʱ������Բ�����Ա�ڷ���ҽ��ʱ�޸ķǱ�����ִ�б����Ŀ�ļ۸���Ȩ��ʱ��ֻ���޸ı�����ִ�б����Ŀ�ļ۸�', 0
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where ϵͳ = &n_System And ��� = 1252 And ���� = '�޸����Ʒ���');






-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--124906:������,2018-04-28,ZLHISê����Ϣ���Ӽ���ص���
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
    �շ���ĿId_In       �շ���ĿĿ¼.Id%Type
  );
  --�����շѶ��ձ䶯
  Procedure Zlhis_Dict_054
  (
    ������ĿId_In     ���Ʒ���Ŀ¼.Id%Type
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

  --���ﻼ�߽���
  Procedure Zlhis_Cis_008
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
  );

  --���ﻼ��ȡ������
  Procedure Zlhis_Cis_009
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
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

  --���ﻼ����ɾ���
  Procedure Zlhis_Cis_057
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
  );

  --���ﻼ��ȡ����ɾ���
  Procedure Zlhis_Cis_058
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
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
    �շ���ĿId_In       �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�շ���ĿID>' || �շ���ĿId_In || '</�շ���ĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_053', v_Value);
  End Zlhis_Dict_053;

  --�����շѶ��ձ䶯
  Procedure Zlhis_Dict_054
  (
    ������ĿId_In     ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><������ĿID>' || ������ĿId_In || '</������ĿID></root>';
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
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><������>' || ������_In ||
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
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><������>' || ������_In ||
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
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><������>' || ������_In ||
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
	n_Length Number(18);
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
	  
      --�жϵ�ǰ�����Ƿ񼴽���������
      Select Lengthb(v_Value || '<��¼ID>' || n_��¼id || '</��¼ID>') Into n_Length From Dual;
      If n_Length > 950 Then
        v_Value := v_Value || '</��¼IDs></root>';
        b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);
        v_Value := '<root><��¼IDs>';
      End If;

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

  --���ﻼ�߽���
  Procedure Zlhis_Cis_008
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_008', '<root><����ID>' || ����id_In || '</����ID><NO>' || �Һŵ�_In || '</NO></root>');
  End Zlhis_Cis_008;

  --���ﻼ��ȡ������
  Procedure Zlhis_Cis_009
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_009', '<root><����ID>' || ����id_In || '</����ID><NO>' || �Һŵ�_In || '</NO></root>');
  End Zlhis_Cis_009;

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

  --���ﻼ����ɾ���
  Procedure Zlhis_Cis_057
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_057', '<root><����ID>' || ����id_In || '</����ID><NO>' || �Һŵ�_In || '</NO></root>');
  End Zlhis_Cis_057;

  --���ﻼ��ȡ����ɾ���
  Procedure Zlhis_Cis_058
  (
    ����id_In In ����ҽ����¼.����id%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_058', '<root><����ID>' || ����id_In || '</����ID><NO>' || �Һŵ�_In || '</NO></root>');
  End Zlhis_Cis_058;

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
--124890:������,2018-04-26,�ж��Ƿ�����˲ɼ���
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
  v_����id  Number;
  v_����    ��Ѫ������.����%Type;
  v_���    ��Ѫ������.���%Type;
  v_��Ӽ�  ��Ѫ������.��Ӽ�%Type;
  v_��Ѫ��  ��Ѫ������.��Ѫ��%Type;
  v_��ɫ    ��Ѫ������.��ɫ%Type;
  v_����id1 ��Ѫ������.����id%Type;
Begin
  If Nvl(����id_In, 0) <> 0 Then
    v_����id := ����id_In;
  Else
    v_����id := Null;
  End If;
Begin   
  Select ����, ���, ��Ӽ�, ��Ѫ��, ��ɫ, ����id
  Into v_����, v_���, v_��Ӽ�, v_��Ѫ��, v_��ɫ, v_����id
  From ��Ѫ������
  Where ���� = ����_In;
    Exception
     When Others Then
      Null;
  End;
  If v_���� <> ����_In Or v_��� <> ���_In Or v_��Ӽ� <> ��Ӽ�_In Or v_��Ѫ�� <> ��Ѫ��_In Or v_��ɫ <> ��ɫ_In Or v_����id <> v_����id1 Then
    Update ��Ѫ������
    Set ���� = ����_In, ��� = ���_In, ��Ӽ� = ��Ӽ�_In, ��Ѫ�� = ��Ѫ��_In, ��ɫ = ��ɫ_In, ����id = v_����id
    Where ���� = ����_In;
    b_Message.Zlhis_Dictlis_008(����_In, ����_In, Null, ���_In, ��Ӽ�_In, ��Ѫ��_In, ��ɫ_In, v_����id);
  End If;
  If Sql%NotFound Then
    Insert Into ��Ѫ������
      (����, ����, ���, ��Ӽ�, ��Ѫ��, ��ɫ, ����id)
    Values
      (����_In, ����_In, ���_In, ��Ӽ�_In, ��Ѫ��_In, ��ɫ_In, v_����id);
    b_Message.Zlhis_Dictlis_007(����_In, ����_In, Null, ���_In, ��Ӽ�_In, ��Ѫ��_In, ��ɫ_In, v_����id);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Ѫ������_Update;
/

--124826:��˶,2018-04-25,��Ϣê�����
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
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�շѼ�Ŀ_Stop;
/

--12426:��˶,2018-04-25,��Ϣê�����
Create Or Replace Procedure Zl_�շ�ϸĿ_Raisemass
(
  ��������_In       In Number,
  ����ֵ_In         In Number,
  ִ������_In       In Date,
  ��ֹ����_In       In Date,
  ������_In         In Varchar2,
  ����id_In         In �շ���ĿĿ¼.����id%Type := Null,
  ���_In           In �շ���ĿĿ¼.���%Type := Null,
  �۸�ȼ�_In       In �շѼ�Ŀ.�۸�ȼ�%Type := '����',
  վ��_In           In �շ���ĿĿ¼.վ��%Type := Null,
  ��ȱʡ�۸����_In Number := 0
) Is
  --�������� 
  --��Σ� 
  --     ��ȱʡ�۸����_In �۸�ȼ�δ���ü۸�ʱ���Ƿ���ȱʡ�۸��ԭ�ۻ����ϵ��� 
  Type t_�շѼ�Ŀ Is Ref Cursor Return �շѼ�Ŀ%RowType;

  c_�շѼ�Ŀ t_�շѼ�Ŀ;
  v_�շѼ�Ŀ �շѼ�Ŀ%RowType;
  v_�ּ�     �շѼ�Ŀ.�ּ�%Type;
  v_No       �շѼ�Ŀ.No%Type;
  v_���     �շѼ�Ŀ.���%Type;

  c_�շ�ϸĿ Sys_Refcursor;
  n_ϸĿid   �շ���ĿĿ¼.Id%Type;
  v_վ��     �շ���ĿĿ¼.վ��%Type;
Begin
  v_No   := Nextno(9);
  v_��� := 1;

  If Nvl(��ȱʡ�۸����_In, 0) = 1 Then
    If ��������_In = 1 Or ��������_In = 3 Then
      --ͬ����Ŀ 
      If ����id_In Is Null Or ����id_In = 0 Then
        Open c_�շ�ϸĿ For
          Select ID, վ�� From �շ���ĿĿ¼ Where ��� = ���_In And �Ƿ��� = 0;
      Else
        If ���_In Is Null Then
          Open c_�շ�ϸĿ For
            Select ID, վ�� From �շ���ĿĿ¼ Where ����id = ����id_In And �Ƿ��� = 0;
        Else
          Open c_�շ�ϸĿ For
            Select ID, վ�� From �շ���ĿĿ¼ Where ����id = ����id_In And ��� = ���_In And �Ƿ��� = 0;
        End If;
      End If;
    Else
      --������������Ŀ 
      If ����id_In Is Null Or ����id_In = 0 Then
        Open c_�շ�ϸĿ For
          Select c.Id, c.վ��
          From �շ���ĿĿ¼ C, (Select ID From �շѷ���Ŀ¼ Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id) D
          Where d.Id = c.����id And c.��� = ���_In And c.�Ƿ��� = 0;
      Else
        If ���_In Is Null Then
          Open c_�շ�ϸĿ For
            Select c.Id, c.վ��
            From �շ���ĿĿ¼ C, (Select ID From �շѷ���Ŀ¼ Start With ID = ����id_In Connect By Prior ID = �ϼ�id) D
            Where d.Id = c.����id And c.�Ƿ��� = 0;
        Else
          Open c_�շ�ϸĿ For
            Select c.Id, c.վ��
            From �շ���ĿĿ¼ C, (Select ID From �շѷ���Ŀ¼ Start With ID = ����id_In Connect By Prior ID = �ϼ�id) D
            Where d.Id = c.����id And c.��� = ���_In And c.�Ƿ��� = 0;
        End If;
      End If;
    End If;
  
    Loop
      Fetch c_�շ�ϸĿ
        Into n_ϸĿid, v_վ��;
      Exit When c_�շ�ϸĿ%NotFound;
    
      For c_�۸�ȼ� In (Select Distinct a.���� As �۸�ȼ�
                     From �շѼ۸�ȼ� A, �շѼ۸�ȼ�Ӧ�� B
                     Where a.���� = b.�۸�ȼ� And Nvl(a.�Ƿ�������ͨ��Ŀ, 0) = 1 And
                           (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                           (�۸�ȼ�_In = '����' Or a.���� = �۸�ȼ�_In) And Not Exists
                      (Select 1 From �շѼ�Ŀ Where �շ�ϸĿid = n_ϸĿid And �۸�ȼ� = a.����) And
                           (
                           --1.��"����Ժ��"Ȩ�ޣ��������м۸�ȼ�/ָ���۸�ȼ� 
                            (վ��_In Is Null And (�۸�ȼ�_In = '����' Or Nvl(a.����, '-') = Nvl(�۸�ȼ�_In, '-')))
                           --2.��"����Ժ��"Ȩ�� 
                            Or
                           --2.1��վ����Ŀ���������м۸�ȼ�/ָ���۸�ȼ� 
                            (v_վ�� = վ��_In And (�۸�ȼ�_In = '����' Or Nvl(a.����, '-') = Nvl(�۸�ȼ�_In, '-')))
                           --2.2ȫԺ��Ŀ�����м۸�ȼ�/ָ���۸�ȼ���ֻ�ܴ���վ�����Ч�۸�ȼ���Ŀ 
                            Or (v_վ�� Is Null
                            --2.2.1���м۸�ȼ� 
                            And (�۸�ȼ�_In = '����' And b.վ�� = վ��_In)
                            --2.2.2ָ���۸�ȼ� 
                            Or a.���� = �۸�ȼ�_In))) Loop
      
        For c_ȱʡ�۸� In (Select �շ�ϸĿid, �ּ�, ������Ŀid, �Ӱ�Ӽ���, �����շ���
                       From �շѼ�Ŀ
                       Where �շ�ϸĿid = n_ϸĿid And (��ֹ���� Is Null Or ��ֹ���� = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                             �۸�ȼ� Is Null) Loop
        
          If ��������_In = 1 Or ��������_In = 2 Then
            --������ 
            v_�ּ� := c_ȱʡ�۸�.�ּ� * (1 + ����ֵ_In);
          Else
            v_�ּ� := c_ȱʡ�۸�.�ּ� + ����ֵ_In;
          End If;
        
          If v_�ּ� > 0 Then
            --����ԭ��ID��������ɾ����δִ�м۸�ʱ���޷�ɾ�� 
            Insert Into �շѼ�Ŀ
              (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �Ӱ�Ӽ���, �����շ���, �䶯ԭ��, ����˵��, ����id, ������, ִ������, ��ֹ����, NO, ���, �۸�ȼ�)
            Values
              (�շѼ�Ŀ_Id.Nextval, Null, c_ȱʡ�۸�.�շ�ϸĿid, c_ȱʡ�۸�.�ּ�, v_�ּ�, c_ȱʡ�۸�.������Ŀid, c_ȱʡ�۸�.�Ӱ�Ӽ���, c_ȱʡ�۸�.�����շ���, 1,
               '��������', �շѼ�Ŀ_Id.Nextval, ������_In, ִ������_In, To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, v_���, c_�۸�ȼ�.�۸�ȼ�);
          
            v_��� := v_��� + 1;
            If c_�۸�ȼ�.�۸�ȼ� Is Null Then
              b_Message.Zlhis_Dict_053(c_ȱʡ�۸�.�շ�ϸĿid);
            End If;
          End If;
        End Loop;
      End Loop;
    End Loop;
    Close c_�շ�ϸĿ;
  End If;

  --���Ȱ���ͬ��Χ�򿪼�¼�� 
  --ȡֵΪ1��������������Χ;2����������ȫ��Χ;3����ֵ������Χ;4����ֵ��ȫ��Χ; 
  If ��������_In = 1 Or ��������_In = 3 Then
    --ͬ����Ŀ 
    If ����id_In Is Null Or ����id_In = 0 Then
      Open c_�շѼ�Ŀ For
        Select a.*
        From �շѼ�Ŀ A, (Select ID, վ�� From �շ���ĿĿ¼ Where ��� = ���_In And �Ƿ��� = 0) B, �շѼ۸�ȼ� P
        Where a.�շ�ϸĿid = b.Id And a.�۸�ȼ� = p.����(+) And
              (
              --1.��"����Ժ��"Ȩ�ޣ��������м۸�ȼ�/ָ���۸�ȼ� 
               (վ��_In Is Null And (�۸�ȼ�_In = '����' Or Nvl(a.�۸�ȼ�, '-') = Nvl(�۸�ȼ�_In, '-')))
              --2.��"����Ժ��"Ȩ�� 
               Or
              --2.1��վ����Ŀ���������м۸�ȼ�/ָ���۸�ȼ� 
               (b.վ�� = վ��_In And (�۸�ȼ�_In = '����' Or Nvl(a.�۸�ȼ�, '-') = Nvl(�۸�ȼ�_In, '-')))
              --2.2ȫԺ��Ŀ�����м۸�ȼ�/ָ���۸�ȼ���ֻ�ܴ���վ�����Ч�۸�ȼ���Ŀ 
               Or (b.վ�� Is Null
               --2.2.1���м۸�ȼ� 
               And (�۸�ȼ�_In = '����' And
               a.�۸�ȼ� In (Select m.����
                                   From �շѼ۸�ȼ� M, �շѼ۸�ȼ�Ӧ�� N
                                   Where m.���� = n.�۸�ȼ� And Nvl(m.�Ƿ�������ͨ��Ŀ, 0) = 1 And n.վ�� = վ��_In)
               --2.2.2ָ���۸�ȼ� 
               Or a.�۸�ȼ� = �۸�ȼ�_In))) And
              Nvl(a.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) = To_Date('3000-01-01', 'YYYY-MM-DD') And
              (p.����ʱ�� Is Null Or p.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(a.No, '-') <> v_No
        Order By a.�շ�ϸĿid;
    Else
      If ���_In Is Null Then
        Open c_�շѼ�Ŀ For
          Select a.*
          From �շѼ�Ŀ A, (Select ID, վ�� From �շ���ĿĿ¼ Where ����id = ����id_In And �Ƿ��� = 0) B, �շѼ۸�ȼ� P
          Where a.�շ�ϸĿid = b.Id And a.�۸�ȼ� = p.����(+) And
                (
                --1.��"����Ժ��"Ȩ�ޣ��������м۸�ȼ�/ָ���۸�ȼ� 
                 (վ��_In Is Null And (�۸�ȼ�_In = '����' Or Nvl(a.�۸�ȼ�, '-') = Nvl(�۸�ȼ�_In, '-')))
                --2.��"����Ժ��"Ȩ�� 
                 Or
                --2.1��վ����Ŀ���������м۸�ȼ�/ָ���۸�ȼ� 
                 (b.վ�� = վ��_In And (�۸�ȼ�_In = '����' Or Nvl(a.�۸�ȼ�, '-') = Nvl(�۸�ȼ�_In, '-')))
                --2.2ȫԺ��Ŀ�����м۸�ȼ�/ָ���۸�ȼ���ֻ�ܴ���վ�����Ч�۸�ȼ���Ŀ 
                 Or (b.վ�� Is Null
                 --2.2.1���м۸�ȼ� 
                 And (�۸�ȼ�_In = '����' And
                 a.�۸�ȼ� In (Select m.����
                                     From �շѼ۸�ȼ� M, �շѼ۸�ȼ�Ӧ�� N
                                     Where m.���� = n.�۸�ȼ� And Nvl(m.�Ƿ�������ͨ��Ŀ, 0) = 1 And n.վ�� = վ��_In)
                 --2.2.2ָ���۸�ȼ� 
                 Or a.�۸�ȼ� = �۸�ȼ�_In))) And
                Nvl(a.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) = To_Date('3000-01-01', 'YYYY-MM-DD') And
                (p.����ʱ�� Is Null Or p.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(a.No, '-') <> v_No
          Order By a.�շ�ϸĿid;
      Else
        Open c_�շѼ�Ŀ For
          Select a.*
          From �շѼ�Ŀ A, (Select ID, վ�� From �շ���ĿĿ¼ Where ����id = ����id_In And ��� = ���_In And �Ƿ��� = 0) B, �շѼ۸�ȼ� P
          Where a.�շ�ϸĿid = b.Id And a.�۸�ȼ� = p.����(+) And
                (
                --1.��"����Ժ��"Ȩ�ޣ��������м۸�ȼ�/ָ���۸�ȼ� 
                 (վ��_In Is Null And (�۸�ȼ�_In = '����' Or Nvl(a.�۸�ȼ�, '-') = Nvl(�۸�ȼ�_In, '-')))
                --2.��"����Ժ��"Ȩ�� 
                 Or
                --2.1��վ����Ŀ���������м۸�ȼ�/ָ���۸�ȼ� 
                 (b.վ�� = վ��_In And (�۸�ȼ�_In = '����' Or Nvl(a.�۸�ȼ�, '-') = Nvl(�۸�ȼ�_In, '-')))
                --2.2ȫԺ��Ŀ�����м۸�ȼ�/ָ���۸�ȼ���ֻ�ܴ���վ�����Ч�۸�ȼ���Ŀ 
                 Or (b.վ�� Is Null
                 --2.2.1���м۸�ȼ� 
                 And (�۸�ȼ�_In = '����' And
                 a.�۸�ȼ� In (Select m.����
                                     From �շѼ۸�ȼ� M, �շѼ۸�ȼ�Ӧ�� N
                                     Where m.���� = n.�۸�ȼ� And Nvl(m.�Ƿ�������ͨ��Ŀ, 0) = 1 And n.վ�� = վ��_In)
                 --2.2.2ָ���۸�ȼ� 
                 Or a.�۸�ȼ� = �۸�ȼ�_In))) And
                Nvl(a.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) = To_Date('3000-01-01', 'YYYY-MM-DD') And
                (p.����ʱ�� Is Null Or p.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(a.No, '-') <> v_No
          Order By a.�շ�ϸĿid;
      End If;
    End If;
  Else
    --������������Ŀ 
    If ����id_In Is Null Or ����id_In = 0 Then
      Open c_�շѼ�Ŀ For
        Select a.*
        From �շѼ�Ŀ A,
             (Select c.Id, c.վ��
               From �շ���ĿĿ¼ C, (Select ID From �շѷ���Ŀ¼ Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id) D
               Where d.Id = c.����id And c.��� = ���_In And c.�Ƿ��� = 0) B, �շѼ۸�ȼ� P
        Where a.�շ�ϸĿid = b.Id And a.�۸�ȼ� = p.����(+) And
              (
              --1.��"����Ժ��"Ȩ�ޣ��������м۸�ȼ�/ָ���۸�ȼ� 
               (վ��_In Is Null And (�۸�ȼ�_In = '����' Or Nvl(a.�۸�ȼ�, '-') = Nvl(�۸�ȼ�_In, '-')))
              --2.��"����Ժ��"Ȩ�� 
               Or
              --2.1��վ����Ŀ���������м۸�ȼ�/ָ���۸�ȼ� 
               (b.վ�� = վ��_In And (�۸�ȼ�_In = '����' Or Nvl(a.�۸�ȼ�, '-') = Nvl(�۸�ȼ�_In, '-')))
              --2.2ȫԺ��Ŀ�����м۸�ȼ�/ָ���۸�ȼ���ֻ�ܴ���վ�����Ч�۸�ȼ���Ŀ 
               Or (b.վ�� Is Null
               --2.2.1���м۸�ȼ� 
               And (�۸�ȼ�_In = '����' And
               a.�۸�ȼ� In (Select m.����
                                   From �շѼ۸�ȼ� M, �շѼ۸�ȼ�Ӧ�� N
                                   Where m.���� = n.�۸�ȼ� And Nvl(m.�Ƿ�������ͨ��Ŀ, 0) = 1 And n.վ�� = վ��_In)
               --2.2.2ָ���۸�ȼ� 
               Or a.�۸�ȼ� = �۸�ȼ�_In))) And
              Nvl(a.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) = To_Date('3000-01-01', 'YYYY-MM-DD') And
              (p.����ʱ�� Is Null Or p.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(a.No, '-') <> v_No
        Order By a.�շ�ϸĿid;
    Else
      If ���_In Is Null Then
        Open c_�շѼ�Ŀ For
          Select a.*
          From �շѼ�Ŀ A,
               (Select c.Id, c.վ��
                 From �շ���ĿĿ¼ C, (Select ID From �շѷ���Ŀ¼ Start With ID = ����id_In Connect By Prior ID = �ϼ�id) D
                 Where d.Id = c.����id And c.�Ƿ��� = 0) B, �շѼ۸�ȼ� P
          Where a.�շ�ϸĿid = b.Id And a.�۸�ȼ� = p.����(+) And
                (
                --1.��"����Ժ��"Ȩ�ޣ��������м۸�ȼ�/ָ���۸�ȼ� 
                 (վ��_In Is Null And (�۸�ȼ�_In = '����' Or Nvl(a.�۸�ȼ�, '-') = Nvl(�۸�ȼ�_In, '-')))
                --2.��"����Ժ��"Ȩ�� 
                 Or
                --2.1��վ����Ŀ���������м۸�ȼ�/ָ���۸�ȼ� 
                 (b.վ�� = վ��_In And (�۸�ȼ�_In = '����' Or Nvl(a.�۸�ȼ�, '-') = Nvl(�۸�ȼ�_In, '-')))
                --2.2ȫԺ��Ŀ�����м۸�ȼ�/ָ���۸�ȼ���ֻ�ܴ���վ�����Ч�۸�ȼ���Ŀ 
                 Or (b.վ�� Is Null
                 --2.2.1���м۸�ȼ� 
                 And (�۸�ȼ�_In = '����' And
                 a.�۸�ȼ� In (Select m.����
                                     From �շѼ۸�ȼ� M, �շѼ۸�ȼ�Ӧ�� N
                                     Where m.���� = n.�۸�ȼ� And Nvl(m.�Ƿ�������ͨ��Ŀ, 0) = 1 And n.վ�� = վ��_In)
                 --2.2.2ָ���۸�ȼ� 
                 Or a.�۸�ȼ� = �۸�ȼ�_In))) And
                Nvl(a.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) = To_Date('3000-01-01', 'YYYY-MM-DD') And
                (p.����ʱ�� Is Null Or p.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(a.No, '-') <> v_No
          Order By a.�շ�ϸĿid;
      Else
        Open c_�շѼ�Ŀ For
          Select a.*
          From �շѼ�Ŀ A,
               (Select c.Id, c.վ��
                 From �շ���ĿĿ¼ C, (Select ID From �շѷ���Ŀ¼ Start With ID = ����id_In Connect By Prior ID = �ϼ�id) D
                 Where d.Id = c.����id And c.��� = ���_In And c.�Ƿ��� = 0) B, �շѼ۸�ȼ� P
          Where a.�շ�ϸĿid = b.Id And a.�۸�ȼ� = p.����(+) And
                (
                --1.��"����Ժ��"Ȩ�ޣ��������м۸�ȼ�/ָ���۸�ȼ� 
                 (վ��_In Is Null And (�۸�ȼ�_In = '����' Or Nvl(a.�۸�ȼ�, '-') = Nvl(�۸�ȼ�_In, '-')))
                --2.��"����Ժ��"Ȩ�� 
                 Or
                --2.1��վ����Ŀ���������м۸�ȼ�/ָ���۸�ȼ� 
                 (b.վ�� = վ��_In And (�۸�ȼ�_In = '����' Or Nvl(a.�۸�ȼ�, '-') = Nvl(�۸�ȼ�_In, '-')))
                --2.2ȫԺ��Ŀ�����м۸�ȼ�/ָ���۸�ȼ���ֻ�ܴ���վ�����Ч�۸�ȼ���Ŀ 
                 Or (b.վ�� Is Null
                 --2.2.1���м۸�ȼ� 
                 And (�۸�ȼ�_In = '����' And
                 a.�۸�ȼ� In (Select m.����
                                     From �շѼ۸�ȼ� M, �շѼ۸�ȼ�Ӧ�� N
                                     Where m.���� = n.�۸�ȼ� And Nvl(m.�Ƿ�������ͨ��Ŀ, 0) = 1 And n.վ�� = վ��_In)
                 --2.2.2ָ���۸�ȼ� 
                 Or a.�۸�ȼ� = �۸�ȼ�_In))) And
                Nvl(a.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) = To_Date('3000-01-01', 'YYYY-MM-DD') And
                (p.����ʱ�� Is Null Or p.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(a.No, '-') <> v_No
          Order By a.�շ�ϸĿid;
      End If;
    End If;
  End If;

  --�����޸� 
  Loop
    Fetch c_�շѼ�Ŀ
      Into v_�շѼ�Ŀ;
    Exit When c_�շѼ�Ŀ%NotFound;
  
    If ��������_In = 1 Or ��������_In = 2 Then
      --������ 
      v_�ּ� := v_�շѼ�Ŀ.�ּ� * (1 + ����ֵ_In);
    Else
      v_�ּ� := v_�շѼ�Ŀ.�ּ� + ����ֵ_In;
    End If;
  
    If v_�ּ� > 0 Then
      Update �շѼ�Ŀ Set ��ֹ���� = ��ֹ����_In Where ID = v_�շѼ�Ŀ.Id;
    
      Insert Into �շѼ�Ŀ
        (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �Ӱ�Ӽ���, �����շ���, �䶯ԭ��, ����˵��, ����id, ������, ִ������, ��ֹ����, NO, ���, �۸�ȼ�)
      Values
        (�շѼ�Ŀ_Id.Nextval, v_�շѼ�Ŀ.Id, v_�շѼ�Ŀ.�շ�ϸĿid, v_�շѼ�Ŀ.�ּ�, v_�ּ�, v_�շѼ�Ŀ.������Ŀid, v_�շѼ�Ŀ.�Ӱ�Ӽ���, v_�շѼ�Ŀ.�����շ���, 1,
         '��������', �շѼ�Ŀ_Id.Nextval, ������_In, ִ������_In, To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, v_���, v_�շѼ�Ŀ.�۸�ȼ�);
    
      v_��� := v_��� + 1;
      If v_�շѼ�Ŀ.�۸�ȼ� Is Null Then
        b_Message.Zlhis_Dict_053(v_�շѼ�Ŀ.�շ�ϸĿid);
      End If;
    End If;
  End Loop;
  Close c_�շѼ�Ŀ;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�շ�ϸĿ_Raisemass;
/


--124269:������,2018-04-20,���ڻ��������´�ҽ�����
Create Or Replace Procedure Zl_����ҽ����¼_Insert
(
  Id_In           In ����ҽ����¼.Id%Type,
  ���id_In       In ����ҽ����¼.���id%Type,
  ���_In         In ����ҽ����¼.���%Type,
  ������Դ_In     In ����ҽ����¼.������Դ%Type,
  ����id_In       In ����ҽ����¼.����id%Type,
  ��ҳid_In       In ����ҽ����¼.��ҳid%Type,
  Ӥ��_In         In ����ҽ����¼.Ӥ��%Type,
  ҽ��״̬_In     In ����ҽ����¼.ҽ��״̬%Type,
  ҽ����Ч_In     In ����ҽ����¼.ҽ����Ч%Type,
  �������_In     In ����ҽ����¼.�������%Type,
  ������Ŀid_In   In ����ҽ����¼.������Ŀid%Type,
  �շ�ϸĿid_In   In ����ҽ����¼.�շ�ϸĿid%Type,
  ����_In         In ����ҽ����¼.����%Type,
  ��������_In     In ����ҽ����¼.��������%Type,
  �ܸ�����_In     In ����ҽ����¼.�ܸ�����%Type,
  ҽ������_In     In ����ҽ����¼.ҽ������%Type,
  ҽ������_In     In ����ҽ����¼.ҽ������%Type,
  �걾��λ_In     In ����ҽ����¼.�걾��λ%Type,
  ִ��Ƶ��_In     In ����ҽ����¼.ִ��Ƶ��%Type,
  Ƶ�ʴ���_In     In ����ҽ����¼.Ƶ�ʴ���%Type,
  Ƶ�ʼ��_In     In ����ҽ����¼.Ƶ�ʼ��%Type,
  �����λ_In     In ����ҽ����¼.�����λ%Type,
  ִ��ʱ�䷽��_In In ����ҽ����¼.ִ��ʱ�䷽��%Type,
  �Ƽ�����_In     In ����ҽ����¼.�Ƽ�����%Type,
  ִ�п���id_In   In ����ҽ����¼.ִ�п���id%Type,
  ִ������_In     In ����ҽ����¼.ִ������%Type,
  ������־_In     In ����ҽ����¼.������־%Type,
  ��ʼִ��ʱ��_In In ����ҽ����¼.��ʼִ��ʱ��%Type,
  ִ����ֹʱ��_In In ����ҽ����¼.ִ����ֹʱ��%Type,
  ���˿���id_In   In ����ҽ����¼.���˿���id%Type,
  ��������id_In   In ����ҽ����¼.��������id%Type,
  ����ҽ��_In     In ����ҽ����¼.����ҽ��%Type,
  ����ʱ��_In     In ����ҽ����¼.����ʱ��%Type,
  �Һŵ�_In       In ����ҽ����¼.�Һŵ�%Type := Null,
  ǰ��id_In       In ����ҽ����¼.ǰ��id%Type := Null,
  ��鷽��_In     In ����ҽ����¼.��鷽��%Type := Null,
  ִ�б��_In     In ����ҽ����¼.ִ�б��%Type := Null,
  �ɷ����_In     In ����ҽ����¼.�ɷ����%Type := Null,
  ժҪ_In         In ����ҽ����¼.ժҪ%Type := Null,
  ����Ա����_In   In ����ҽ��״̬.������Ա%Type := Null,
  ��Ѽ���_In     In ����ҽ����¼.��Ѽ���%Type := Null,
  ��ҩĿ��_In     In ����ҽ����¼.��ҩĿ��%Type := Null,
  ��ҩ����_In     In ����ҽ����¼.��ҩ����%Type := Null,
  ���״̬_In     In ����ҽ����¼.���״̬%Type := Null,
  �������_In     In ����ҽ����¼.�������%Type := Null,
  ����˵��_In     In ����ҽ����¼.����˵��%Type := Null,
  �״�����_In     In ����ҽ����¼.�״�����%Type := Null,
  �䷽id_In       In ����ҽ����¼.�䷽id%Type := Null,
  �������_In     In ����ҽ����¼.�������%Type := Null,
  �����Ŀid_In   In ����ҽ����¼.�����Ŀid%Type := Null,
  Ƥ�Խ��_In     In ����ҽ����¼.Ƥ�Խ��%Type := Null,
  �������_In     In ����ҽ����¼.�������%Type := Null,
  ����ҽ��id_In   In ����ҽ����¼.����ҽ��id%Type := Null
  --���ܣ�ҽ����ʿ�¿�,��¼ҽ��ʱ�²�����ҽ����¼�������������סԺ��
) Is
  v_Temp     Varchar2(255);
  v_��Ա���� ����ҽ��״̬.������Ա%Type;
  n_��Աid   ���ű�.Id%Type;
  v_����     ������Ϣ.����%Type;
  v_�Ա�     ������Ϣ.�Ա�%Type;
  v_����     ������Ϣ.����%Type;
  d_����ʱ�� ����ҽ����¼.����ʱ��%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --��ǰ������Ա
  If ����Ա����_In Is Not Null Then
    v_��Ա���� := ����Ա����_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  If Nvl(��ҳid_In, 0) <> 0 Then
    Select ����, �Ա�, ���� Into v_����, v_�Ա�, v_���� From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  Else
    Select ����, �Ա�, ���� Into v_����, v_�Ա�, v_���� From ������Ϣ Where ����id = ����id_In;
  End If;

  If Instr(',F,K,', �������_In) > 0 Then
    d_����ʱ�� := To_Date(�걾��λ_In, 'yyyy-mm-dd hh24:mi:ss');
  End If;

  --����ҽ����¼
  Insert Into ����ҽ����¼
    (ID, ���id, ���, ������Դ, ����id, ��ҳid, ����, �Ա�, ����, Ӥ��, ҽ��״̬, ҽ����Ч, �������, ������Ŀid, �շ�ϸĿid, ����, ��������, �ܸ�����, ҽ������, ҽ������, �걾��λ,
     ��鷽��, ִ�б��, ִ��Ƶ��, Ƶ�ʴ���, Ƶ�ʼ��, �����λ, ִ��ʱ�䷽��, �Ƽ�����, ִ�п���id, ִ������, ������־, �ɷ����, ��ʼִ��ʱ��, ִ����ֹʱ��, ���˿���id, ��������id, ����ҽ��,
     ����ʱ��, �Һŵ�, ǰ��id, ժҪ, ��Ѽ���, ����ʱ��, ��ҩĿ��, ��ҩ����, ���״̬, �������, ����˵��, �״�����, �䷽id, �������, �����Ŀid, Ƥ�Խ��, �������, ����ҽ��id)
  Values
    (Id_In, ���id_In, ���_In, ������Դ_In, ����id_In, ��ҳid_In, v_����, v_�Ա�, v_����, Ӥ��_In, ҽ��״̬_In, ҽ����Ч_In, �������_In, ������Ŀid_In,
     �շ�ϸĿid_In, ����_In, ��������_In, �ܸ�����_In, ҽ������_In, ҽ������_In, �걾��λ_In, ��鷽��_In, ִ�б��_In, ִ��Ƶ��_In, Ƶ�ʴ���_In, Ƶ�ʼ��_In, �����λ_In,
     ִ��ʱ�䷽��_In, �Ƽ�����_In, ִ�п���id_In, ִ������_In, ������־_In, �ɷ����_In, ��ʼִ��ʱ��_In, ִ����ֹʱ��_In, ���˿���id_In, ��������id_In, ����ҽ��_In,
     ����ʱ��_In, �Һŵ�_In, ǰ��id_In, ժҪ_In, ��Ѽ���_In, d_����ʱ��, ��ҩĿ��_In, ��ҩ����_In, ���״̬_In, �������_In, ����˵��_In, �״�����_In, �䷽id_In,
     �������_In, �����Ŀid_In, Ƥ�Խ��_In, �������_In, ����ҽ��id_In);

  --����ҽ��״̬
  If ҽ��״̬_In <> -1 Then
    Delete From ����ҽ��״̬ Where ҽ��id = Id_In And �������� = 1;
    If Sql%RowCount <> 0 Then
      v_Error := '��ͬID���¿�ҽ���Ѿ����ڡ�';
      Raise Err_Custom;
    End If;
    --��Ϊ����ͬʱ���¿�->�Զ�У��(סԺҽ������)->�����Զ�ֹͣ(סԺҽ����������ֹͣ),��˷ֱ�-2,-1��
    Insert Into ����ҽ��״̬
      (ҽ��id, ��������, ������Ա, ����ʱ��)
    Values
      (Id_In, 1, v_��Ա����, Sysdate - 2 / 60 / 60 / 24);
  End If;
  --���³�����Ŀ
  If Nvl(�Һŵ�_In, '��') <> '��' Then
    Select ID Into n_��Աid From ��Ա�� Where ���� = v_��Ա����;
    Zl_ҽ������ҽ��_Update(n_��Աid, ��������id_In, ������Ŀid_In, �շ�ϸĿid_In, �������_In, �걾��λ_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_Insert;
/

--124269:������,2018-04-20,���ڻ��������´�ҽ�����
Create Or Replace Procedure Zl_����ҽ����¼_Update
(
  Id_In           In ����ҽ����¼.Id%Type,
  ���id_In       In ����ҽ����¼.���id%Type,
  ���_In         In ����ҽ����¼.���%Type,
  ҽ��״̬_In     In ����ҽ����¼.ҽ��״̬%Type,
  ҽ����Ч_In     In ����ҽ����¼.ҽ����Ч%Type,
  ������Ŀid_In   In ����ҽ����¼.������Ŀid%Type,
  �շ�ϸĿid_In   In ����ҽ����¼.�շ�ϸĿid%Type,
  ����_In         In ����ҽ����¼.����%Type,
  ��������_In     In ����ҽ����¼.��������%Type,
  �ܸ�����_In     In ����ҽ����¼.�ܸ�����%Type,
  ҽ������_In     In ����ҽ����¼.ҽ������%Type,
  ҽ������_In     In ����ҽ����¼.ҽ������%Type,
  �걾��λ_In     In ����ҽ����¼.�걾��λ%Type,
  ִ��Ƶ��_In     In ����ҽ����¼.ִ��Ƶ��%Type,
  Ƶ�ʴ���_In     In ����ҽ����¼.Ƶ�ʴ���%Type,
  Ƶ�ʼ��_In     In ����ҽ����¼.Ƶ�ʼ��%Type,
  �����λ_In     In ����ҽ����¼.�����λ%Type,
  ִ��ʱ�䷽��_In In ����ҽ����¼.ִ��ʱ�䷽��%Type,
  �Ƽ�����_In     In ����ҽ����¼.�Ƽ�����%Type,
  ִ�п���id_In   In ����ҽ����¼.ִ�п���id%Type,
  ִ������_In     In ����ҽ����¼.ִ������%Type,
  ������־_In     In ����ҽ����¼.������־%Type,
  ��ʼִ��ʱ��_In In ����ҽ����¼.��ʼִ��ʱ��%Type,
  ִ����ֹʱ��_In In ����ҽ����¼.ִ����ֹʱ��%Type,
  ���˿���id_In   In ����ҽ����¼.���˿���id%Type,
  ��������id_In   In ����ҽ����¼.��������id%Type,
  ����ҽ��_In     In ����ҽ����¼.����ҽ��%Type,
  ����ʱ��_In     In ����ҽ����¼.����ʱ��%Type,
  ��鷽��_In     In ����ҽ����¼.��鷽��%Type := Null,
  ִ�б��_In     In ����ҽ����¼.ִ�б��%Type := Null,
  �ɷ����_In     In ����ҽ����¼.�ɷ����%Type := Null,
  ժҪ_In         In ����ҽ����¼.ժҪ%Type := Null,
  ��Ա������_In   In ����ҽ��״̬.������Ա%Type := Null,
  ��Ѽ���_In     In ����ҽ����¼.��Ѽ���%Type := Null,
  ��ҩĿ��_In     In ����ҽ����¼.��ҩĿ��%Type := Null,
  ��ҩ����_In     In ����ҽ����¼.��ҩ����%Type := Null,
  ���״̬_In     In ����ҽ����¼.���״̬%Type := Null,
  ����˵��_In     In ����ҽ����¼.����˵��%Type := Null,
  �״�����_In     In ����ҽ����¼.�״�����%Type := Null,
  �������_In     In ����ҽ����¼.�������%Type := Null,
  �����Ŀid_In   In ����ҽ����¼.�����Ŀid%Type := Null,
  Ƥ�Խ��_In     In ����ҽ����¼.Ƥ�Խ��%Type := Null,
  �������_In     In ����ҽ����¼.�������%Type := Null,
  ����ҽ��id_In   In ����ҽ����¼.����ҽ��id%Type := Null
  --���ܣ���ҽ����ʿ�޸��˲������ݵ�ҽ����¼�������������סԺ��
  --˵����Updateʱ֮�����漰������ĿID,�Ƽ����Ա仯,����Ϊ��ҩ;��,�÷��ı仯
  --      Updateʱ֮�����漰��Ч�仯,����Ϊ����¼��ҽ��������ı���Ч
) Is
  v_Count Number;

  v_Temp            Varchar2(255);
  v_��Ա����        ����ҽ��״̬.������Ա%Type;
  v_�����������ids Varchar2(4000);
  v_ҩƷ����        ����ҽ����¼.ҽ������%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --����ҽ��״̬:��������
  Begin
    Select ҽ��״̬ Into v_Count From ����ҽ����¼ Where ID = Id_In;
  Exception
    When Others Then
      Begin
        v_Error := 'ҽ��"' || ҽ������_In || '"�Ѿ�������,�����ѱ�������ɾ����';
        Raise Err_Custom;
      End;
  End;
  If v_Count Not In (-1, 1, 2) Then
    v_Error := 'ҽ��"' || ҽ������_In || '"�Ѿ�У�Ի���,�������޸ġ�';
    Raise Err_Custom;
  End If;

  Select Count(*) Into v_Count From ����ҽ��״̬ Where ҽ��id = Id_In And �������� = 1 And ǩ��id Is Not Null;
  If Nvl(v_Count, 0) > 0 Then
    v_Error := 'ҽ��"' || ҽ������_In || '"�Ѿ�����ǩ��,�������޸ġ�';
    Raise Err_Custom;
  End If;

  --������鳷��
  If ���id_In Is Null Then
    Zl_�������_Cancel(Id_In, v_�����������ids);
  End If;

  If v_�����������ids Is Not Null Then
    Select ҽ������ Into v_ҩƷ���� From ����ҽ����¼ Where ���id = Id_In And Rownum < 2;
    v_Error := 'ҽ��"' || v_ҩƷ���� || '"�����������ڽ��д�����飬�������޸ġ�';
    Raise Err_Custom;
  End If;

  --��ǰ������Ա
  If ��Ա������_In Is Not Null Then
    v_��Ա���� := ��Ա������_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  --����ҽ����¼
  Update ����ҽ����¼
  Set ���id = ���id_In,
      --����һ����ҩ���������ü�鲿λ����������ID�仯
      ��� = ���_In, ҽ��״̬ = ҽ��״̬_In,
      --!��Ϊֻ���޸�δУ��ҽ��������Ӧ��Ϊ�¿���У�����ʵ�ҽ���޸ĺ�Ϊ�¿�
      ҽ����Ч = ҽ����Ч_In, ������Ŀid = ������Ŀid_In, �շ�ϸĿid = �շ�ϸĿid_In, ���� = ����_In, �������� = ��������_In, �ܸ����� = �ܸ�����_In, ҽ������ = ҽ������_In,
      ҽ������ = ҽ������_In, �걾��λ = �걾��λ_In, ��鷽�� = ��鷽��_In, ִ�б�� = ִ�б��_In, ִ��Ƶ�� = ִ��Ƶ��_In, Ƶ�ʴ��� = Ƶ�ʴ���_In, Ƶ�ʼ�� = Ƶ�ʼ��_In,
      �����λ = �����λ_In, ִ��ʱ�䷽�� = ִ��ʱ�䷽��_In, �Ƽ����� = �Ƽ�����_In, ִ�п���id = ִ�п���id_In, ִ������ = ִ������_In, �ɷ���� = �ɷ����_In,
      --ҩƷ�����⹺ҩ,��Ժ��ҩ�ĵ���ʱ�ᷢ���仯
      ������־ = ������־_In, ��ʼִ��ʱ�� = ��ʼִ��ʱ��_In, ִ����ֹʱ�� = ִ����ֹʱ��_In,
      --!��������ֹʱ������޸�,����Ӧ��Ϊ��
      ���˿���id = ���˿���id_In,
      --�޸�ʱ����Ϊ���˵ĵ�ǰ����
      ��������id = ��������id_In,
      --�޸ĺ����ݵ�ǰ���ұ仯
      ����ҽ�� = ����ҽ��_In, ��˱�� = Decode(Nvl(Instr(����ҽ��_In, '/'), 0), 0, Decode(��˱��, 1, Null, ��˱��), 1),
      --��ʿ��ҽ��ʱ���Ը���
      ����ʱ�� = ����ʱ��_In,
      --��¼�Ŀ����޸�
      ժҪ = ժҪ_In, ��Ѽ��� = ��Ѽ���_In, ����ʱ�� = Decode(�������, 'F', To_Date(�걾��λ_In, 'yyyy-mm-dd hh24:mi:ss'), Null),
      ��ҩĿ�� = ��ҩĿ��_In, ��ҩ���� = ��ҩ����_In, ���״̬ = ���״̬_In, ����˵�� = ����˵��_In, �״����� = �״�����_In, ������� = �������_In, �����Ŀid = �����Ŀid_In,
      Ƥ�Խ�� = Ƥ�Խ��_In,
      --������ҩ���
      ������� = �������_In, ����ҽ��id = ����ҽ��id_In
  Where ID = Id_In;

  --����ҽ��״̬:����ҽ���¿�����
  --��Ϊ����ͬʱ���¿�(�޸�)->�Զ�У��(סԺҽ������)->�����Զ�ֹͣ(סԺҽ����������ֹͣ),��˷ֱ�-2,-1��
  If ҽ��״̬_In <> -1 Then
    Update ����ҽ��״̬
    Set ������Ա = v_��Ա����, ����ʱ�� = Sysdate - 2 / 60 / 60 / 24
    Where ҽ��id = Id_In And �������� = 1; --�¿�����ʼ����,У�����ʱ�����Ϊ��ʷ��¼
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_Update;
/  





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0008' Where ���=&n_System;
Commit;