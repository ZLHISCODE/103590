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


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--137701:������,2019-02-12,������Ϣ�ϲ���Ϣê���޸�
Create Or Replace Package b_Message Is
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
    Id_In   �շ���ĿĿ¼.Id%Type,
    ����_In �շ���ĿĿ¼.����%Type,
    ����_In �շ���ĿĿ¼.����%Type,
    ���_In �շ���ĿĿ¼.���%Type,
    ����_In �շ���ĿĿ¼.����%Type
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
  Procedure Zlhis_Dictlis_004
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    �����Ա�_In ���Ƽ���걾.�����Ա�%Type
  );
  --�޸����Ƽ���걾
  Procedure Zlhis_Dictlis_005
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    �����Ա�_In ���Ƽ���걾.�����Ա�%Type
  );
  --ɾ��������Ŀ��λ
  Procedure Zlhis_Dictlis_006
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    �����Ա�_In ���Ƽ���걾.�����Ա�%Type
  );
  --������Ѫ������
  Procedure Zlhis_Dictlis_007
  (
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����id_In ��Ѫ������.����id%Type
  );
  --�޸Ĳ�Ѫ������
  Procedure Zlhis_Dictlis_008
  (
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����id_In ��Ѫ������.����id%Type
  );
  --ɾ����Ѫ������
  Procedure Zlhis_Dictlis_009
  (
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����id_In ��Ѫ������.����id%Type
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
  Procedure Zlhis_Drug_007(�۸�id_In ҩƷ�۸��¼.Id%Type);
  --���䷢��
  Procedure Zlhis_Drug_008(��¼ids_In Varchar2);
  --ҩƷ���ۼ�
  Procedure Zlhis_Drug_009
  (
    �۸�id_In ҩƷ�۸��¼.Id%Type,
    ʱ��_In   Number
  );
  --���ĵ��ɱ���
  Procedure Zlhis_Drug_010(�۸�id_In �ɱ��۵�����Ϣ.Id%Type);
  --���ĵ��ۼ�
  Procedure Zlhis_Drug_011
  (
    �۸�id_In �շѼ�Ŀ.Id%Type,
    ʱ��_In   Number
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

  --ȷ��ֹͣ����ҽ�� 
  Procedure Zlhis_Cis_059
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
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
    ԭ����id_In In ������ҳ.����id%Type,
    �仯ids_In  In Varchar2 
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
Create Or Replace Package Body b_Message Is
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
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����></root>';
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
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����></root>';
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
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����><���>' || ���_In || '</���><����>' || ����_In || '</����></root>';
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
    Id_In   �շ���ĿĿ¼.Id%Type,
    ����_In �շ���ĿĿ¼.����%Type,
    ����_In �շ���ĿĿ¼.����%Type,
    ���_In �շ���ĿĿ¼.���%Type,
    ����_In �շ���ĿĿ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID><����>' || ����_In || '</����><����>' || ����_In || '</����><���>' || ���_In ||
               '</���><����>' || ����_In || '</����></root>';
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
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_050', v_Value);
  End Zlhis_Dict_050;
  --�޸����ķ���
  Procedure Zlhis_Dict_051
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_051', v_Value);
  End Zlhis_Dict_051;
  --ɾ�����ķ���
  Procedure Zlhis_Dict_052
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_052', v_Value);
  End Zlhis_Dict_052;
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
  Procedure Zlhis_Dictlis_004
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    �����Ա�_In ���Ƽ���걾.�����Ա�%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In ||
               '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_004', v_Value);
  End Zlhis_Dictlis_004;
  --�޸�������Ŀ��λ
  Procedure Zlhis_Dictlis_005
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    �����Ա�_In ���Ƽ���걾.�����Ա�%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In ||
               '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_005', v_Value);
  End Zlhis_Dictlis_005;
  --ɾ��������Ŀ��λ
  Procedure Zlhis_Dictlis_006
  (
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    ����_In     ���Ƽ���걾.����%Type,
    �����Ա�_In ���Ƽ���걾.�����Ա�%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In ||
               '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_006', v_Value);
  End Zlhis_Dictlis_006;
  --������Ѫ������
  Procedure Zlhis_Dictlis_007
  (
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����id_In ��Ѫ������.����id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><��Ӽ�>' || ��Ӽ�_In ||
               '</��Ӽ�><��Ѫ��>' || ��Ѫ��_In || '</��Ѫ��><��ɫ>' || ��ɫ_In || '</��ɫ><���>' || ���_In || '</���><����ID_In>' || ����id_In ||
               '</����ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_007', v_Value);
  End Zlhis_Dictlis_007;
  --������Ѫ������
  Procedure Zlhis_Dictlis_008
  (
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����id_In ��Ѫ������.����id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><��Ӽ�>' || ��Ӽ�_In ||
               '</��Ӽ�><��Ѫ��>' || ��Ѫ��_In || '</��Ѫ��><��ɫ>' || ��ɫ_In || '</��ɫ><���>' || ���_In || '</���><����ID_In>' || ����id_In ||
               '</����ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_008', v_Value);
  End Zlhis_Dictlis_008;
  --������Ѫ������
  Procedure Zlhis_Dictlis_009
  (
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ����_In   ��Ѫ������.����%Type,
    ��Ӽ�_In ��Ѫ������.��Ӽ�%Type,
    ��Ѫ��_In ��Ѫ������.��Ѫ��%Type,
    ���_In   ��Ѫ������.���%Type,
    ��ɫ_In   ��Ѫ������.��ɫ%Type,
    ����id_In ��Ѫ������.����id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><��Ӽ�>' || ��Ӽ�_In ||
               '</��Ӽ�><��Ѫ��>' || ��Ѫ��_In || '</��Ѫ��><��ɫ>' || ��ɫ_In || '</��ɫ><���>' || ���_In || '</���><����ID_In>' || ����id_In ||
               '</����ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_009', v_Value);
  End Zlhis_Dictlis_009;
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
  Procedure Zlhis_Drug_007(�۸�id_In ҩƷ�۸��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_007', v_Value);
  End Zlhis_Drug_007;
  --���䷢��
  Procedure Zlhis_Drug_008(��¼ids_In Varchar2) Is
    v_Value  Zlmsg_Todo.Key_Value%Type;
    n_��¼id ��Һ��ҩ��¼.Id%Type;
    v_Tmp    Varchar2(4000);
	n_Length Number(18);
  Begin
    If ��¼ids_In Is Null Then
      v_Tmp := Null;
    Else
      v_Tmp := ��¼ids_In || ',';
    End If;
  
    v_Value := '<root><��¼IDS>';
  
    While v_Tmp Is Not Null Loop
      --�ֽⵥ��ID��
      n_��¼id := To_Number(Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1));
      v_Tmp    := Replace(',' || v_Tmp, ',' || n_��¼id || ',');
      
      --�жϵ�ǰ�����Ƿ񼴽���������                                                                        
      Select Lengthb(v_Value || '<��¼ID>' || n_��¼id || '</��¼ID>') Into n_Length From Dual;            
      If n_Length > 950 Then								                   
        v_Value := v_Value || '</��¼IDs></root>';                                                         
        b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);                                            
        v_Value := '<root><��¼IDs>';                                                                      
      End If;

      v_Value := v_Value || '<��¼ID>' || n_��¼id || '</��¼ID>';
    End Loop;
  
    v_Value := v_Value || '</��¼IDS></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);
  End Zlhis_Drug_008;
  --ҩƷ���ۼ�
  Procedure Zlhis_Drug_009
  (
    �۸�id_In ҩƷ�۸��¼.Id%Type,
    ʱ��_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID><ʱ��>' || ʱ��_In || '</ʱ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_009', v_Value);
  End Zlhis_Drug_009;
  --���ĵ��ɱ���
  Procedure Zlhis_Drug_010(�۸�id_In �ɱ��۵�����Ϣ.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_010', v_Value);
  End Zlhis_Drug_010;
  --���ĵ��ۼ�
  Procedure Zlhis_Drug_011
  (
    �۸�id_In �շѼ�Ŀ.Id%Type,
    ʱ��_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID><ʱ��>' || ʱ��_In || '</ʱ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_011', v_Value);
  End Zlhis_Drug_011;

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
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_�������� ������ĿĿ¼.��������%Type;
  Begin
    Select Max(a.��������)
    Into v_��������
    From ������ĿĿ¼ A, ����ҽ����¼ B
    Where b.������Ŀid = a.Id And b.Id = ҽ��id_In;
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
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_�������� ������ĿĿ¼.��������%Type;
  Begin
    Select Max(a.��������)
    Into v_��������
    From ������ĿĿ¼ A, ����ҽ����¼ B
    Where b.������Ŀid = a.Id And b.Id = ҽ��id_In;
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
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_�������� ������ĿĿ¼.��������%Type;
  Begin
    Select Max(a.��������)
    Into v_��������
    From ������ĿĿ¼ A, ����ҽ����¼ B
    Where b.������Ŀid = a.Id And b.Id = ҽ��id_In;
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

  --ȷ��ֹͣ����ҽ�� 
  Procedure Zlhis_Cis_059
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_059','<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_059;

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
    ԭ����id_In In ������ҳ.����id%Type,
    �仯ids_In  In Varchar2
  ) Is 
  --������ 1����id,1��ҳid:1ԭ����id,1ԭ��ҳid; 2����id,2��ҳid:2ԭ����id,2ԭ��ҳid;��.
  Begin 
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_017', 
                                '<root><����ID>' || ����id_In || '</����ID><ԭ����ID>' || ԭ����id_In || '</ԭ����ID><CINFO>'||�仯ids_In||'</CINFO></root>'); 
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
    v_�����   ������Ϣ.�����%Type; 
    v_���֤�� ������Ϣ.���֤��%Type; 
    v_�������� varchar2(50); 
  Begin 
    If p_Msg_Using('ZLHIS_PATIENT_028') = 0 Then 
      Return; 
    End If; 
    Select ����, �Ա�, ����, To_Char(��������, 'yyyymmdd'), �����, ���֤�� 
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

--137701:������,2019-02-12,������Ϣ�ϲ���Ϣê���޸�
Create Or Replace Procedure Zl_������Ϣ_Merge
(
  A����id_In    ������Ϣ.����id%Type, --Ҫ�ϲ��Ĳ�����Ϣ
  B����id_In    ������Ϣ.����id%Type, --Ҫ�����Ĳ�����Ϣ
  �ϲ�ԭ��_In   ���˺ϲ���¼.�ϲ�ԭ��%Type,
  ����Ա����_In ��Ա��.����%Type,
  ǿ�Ʊ���_In   Number := 0
  --��׼��
  ----------------------------------------------------------------------------
  --������Ϣ,������ҳ,������ҳ�ӱ�,���˱䶯��¼,���ⲡ��
  --���ﲡ����¼,סԺ������¼,��λ״����¼
  --ҽ�����˵���,����ģ�����,���ս����¼,�ʻ������Ϣ
  --�������,����δ�����,סԺ���ü�¼,������ü�¼,����Ԥ����¼,���˽��ʼ�¼,δ��ҩƷ��¼
  --���˹Һż�¼,���˹���ҩ��,���˹�����¼,������ϼ�¼,������
  --����ҽ����¼,���������¼
  --����������Ϣ
  
  --�󱸱�
  --H���˽��ʼ�¼,H����Ԥ����¼,HסԺ���ü�¼,H������ü�¼
  --H����ҽ����¼,H������ϼ�¼,H���˹�����¼
  --H���˲�����¼,H���������¼
  
  --����ϵͳ
  ----------------------------------------------------------------------------
  --���˷���,�����¼,���ļ�¼
  --��������ϼ�¼,���˷�����Ϣ
  --��Ϸ������,�������ֽ��
  
) As
  --������ر�
  Cursor c_Patitable Is
    Select a.Table_Name, Max(Decode(b.Column_Name, '����ID', 1, 0)) As ����id,
           Max(Decode(b.Column_Name, '��ҳID', 1, 0)) As ��ҳid
    From User_Tables A, User_Tab_Columns B
    Where a.Table_Name = b.Table_Name And b.Column_Name In ('����ID', '��ҳID') And
          a.Table_Name Not In
          ('������Ϣ', '������ҳ', '������ҳ�ӱ�', '���˱䶯��¼','�����Զ�����', '���ⲡ��', '���ﲡ����¼', 'סԺ������¼', '��λ״����¼', 'ҽ�����˵���', 'ҽ�����˹�����', '����ģ�����',
           '�ʻ������Ϣ', '�������', '����δ�����', 'סԺ���ü�¼', '������ü�¼', '����Ԥ����¼', '���˽��ʼ�¼', 'δ��ҩƷ��¼', '���˹Һż�¼', '���˹���ҩ��', '���˹�����¼',
           '������ϼ�¼', '������', '����ҽ����¼', '���������¼', '���˷���', '�����¼', '���ļ�¼', '���˷�����Ϣ', '��Ϸ������', '�������ֽ��', '���˵�����¼', '����������Ϣ',
           '�������߼�¼', '������Ϣ�ӱ�', '����ҽ�ƿ�����') Having Max(Decode(b.Column_Name, '����ID', 1, 0)) <> 0
    Group By a.Table_Name;

  --���鶨��
  Type Array_Patitable Is Table Of Varchar2(100) Index By Binary_Integer;
  Arronbase Array_Patitable;
  Arronpage Array_Patitable;
  v_Loop    Number;
  n_Have    Number;

  -------------------------------------------------------
  --���ϲ��Ĳ���(סԺ�ſ���ÿ���²���,���סԺȡ���һ��)
  Cursor c_Infoa Is
    Select a.����id, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.����֤��, a.���,
           a.ְҵ, a.����, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.�໤��, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ,
           a.��ϵ�˵绰, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.Email, a.Qq, a.��ͬ��λid, a.������λ, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.������, a.������,
           a.��������, a.����ʱ��, a.����״̬, a.��������, a.סԺ����, a.��ǰ����id, a.��ǰ����id, a.��ǰ����, a.��Ժʱ��, a.��Ժʱ��, a.��Ժ, a.Ic����, a.������,
           a.ҽ����, a.����, a.��ѯ����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.����, a.��ϵ�����֤��, b.��ҳid, b.��Ժ����, b.��Ժ����
    From ������Ϣ A, ������ҳ B
    Where a.����id = b.����id(+) And a.����id = A����id_In
    Order By ��Ժ���� Desc, ��ҳid Desc;
  r_Infoa c_Infoa%RowType;

  --Ҫ�����Ĳ���(סԺ�ſ���ÿ���²���,���סԺȡ���һ��)
  Cursor c_Infob Is
    Select a.����id, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.����֤��, a.���,
           a.ְҵ, a.����, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.�໤��, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ,
           a.��ϵ�˵绰, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.Email, a.Qq, a.��ͬ��λid, a.������λ, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.������, a.������,
           a.��������, a.����ʱ��, a.����״̬, a.��������, a.סԺ����, a.��ǰ����id, a.��ǰ����id, a.��ǰ����, a.��Ժʱ��, a.��Ժʱ��, a.��Ժ, a.Ic����, a.������,
           a.ҽ����, a.����, a.��ѯ����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.����, a.��ϵ�����֤��, b.��ҳid, b.��Ժ����, b.��Ժ����
    From ������Ϣ A, ������ҳ B
    Where a.����id = b.����id(+) And a.����id = B����id_In
    Order By ��Ժ���� Desc, ��ҳid Desc;
  r_Infob c_Infob%RowType;

  --�ϲ������Ϣ
  Cursor c_Info(v_����id ������Ϣ.����id%Type) Is
    Select ����id, ��ҳid, (Select Nvl(Max(��ҳid), 0) From ������ҳ Where ����id = v_����id) �����ҳid, סԺ��, ��������, ҽ�Ƹ��ʽ, �ѱ�, ����Ժ,
           ��Ժ����id, ��Ժ����id, ҽ��С��id, ��Ժ����, ��Ժ����, ��Ժ��ʽ, ��Ժ����, ����Ժת��, סԺĿ��, ��Ժ����, �Ƿ����, ��ǰ����, ��ǰ����id, ����ȼ�id, ��Ժ����id, ��Ժ����,
           ��Ժ����, סԺ����, ��Ժ��ʽ, �Ƿ�ȷ��, ȷ������, �·�����, Ѫ��, ���ȴ���, �ɹ�����, �����־, ��������, ʬ���־, ����ҽʦ, ���λ�ʿ, סԺҽʦ, ������, ��ĿԱ���, ��ĿԱ����,
           ��Ŀ����, ״̬, ���ú�, ����, ���, ����, ����״��, ְҵ, ����, ѧ��, ��λ�绰, ��λ�ʱ�, ��λ��ַ, ����, ��ͥ��ַ, ��ͥ�绰, ��ͥ��ַ�ʱ�, ��ϵ������, ��ϵ�˹�ϵ, ��ϵ�˵�ַ,
           ��ϵ�˵绰, ��ϵ�����֤��, ���ڵ�ַ, ���ڵ�ַ�ʱ�, ��ҽ�������, ����, ����, ��˱�־, �����, �������, �Ƿ��ϴ�, ����ת��, �Ǽ���, �Ǽ�ʱ��, ��ע, ����״̬, ��������
    From ������ҳ
    Where ��ҳid = (Select Nvl(Max(��ҳid), 0)
                  From ������ҳ
                  Where ����id = v_����id And Not Exists (Select ��ҳid From ������ҳ Where ����id = v_����id And ��ҳid = 0)) And
          ����id = v_����id;
  r_Info c_Info%RowType;

  --�ϲ�����סԺ����
  Cursor c_Mergepati Is
    Select a.����, a.�����, a.סԺ�� ��ǰסԺ��, b.����id, b.��ҳid, b.סԺ��, b.���ۺ�, b.��������, b.ҽ�Ƹ��ʽ, b.�ѱ�, b.����Ժ, b.��Ժ����id, b.��Ժ����id,
           b.ҽ��С��id, b.��Ժ����, b.��Ժ����, b.��Ժ��ʽ, b.��Ժ����, b.����Ժת��, b.סԺĿ��, b.��Ժ����, b.�Ƿ����, b.��ǰ����, b.��ǰ����id, b.����ȼ�id,
           b.��Ժ����id, b.��Ժ����, b.��Ժ����, b.סԺ����, b.��Ժ��ʽ, b.�Ƿ�ȷ��, b.ȷ������, b.�·�����, b.Ѫ��, b.���ȴ���, b.�ɹ�����, b.�����־, b.��������,
           b.ʬ���־, b.����ҽʦ, b.���λ�ʿ, b.סԺҽʦ, b.������, b.��ĿԱ���, b.��ĿԱ����, b.��Ŀ����, b.״̬, b.���ú�, b.�Ա�, b.����, b.���, b.����, b.����״��,
           b.ְҵ, b.����, b.ѧ��, b.��λ�绰, b.��λ�ʱ�, b.��λ��ַ, b.����, b.��ͥ��ַ, b.��ͥ�绰, b.��ͥ��ַ�ʱ�, b.��ϵ������, b.��ϵ�˹�ϵ, b.��ϵ�˵�ַ, b.��ϵ�˵绰,
           b.��ϵ�����֤��, b.���ڵ�ַ, b.���ڵ�ַ�ʱ�, b.��ҽ�������, b.����, b.����, b.��˱�־, b.�����, b.�������, b.�Ƿ��ϴ�, b.����ת��, b.�Ǽ���, b.�Ǽ�ʱ��, b.��ע,
           b.����״̬, b.��������, b.���ʱ��, b.·��״̬, b.������, b.Ӥ������id, b.Ӥ������id, b.ĸӤת�Ʊ�־, b.ҽ������ʱ��
    From ������Ϣ A, ������ҳ B
    Where a.����id = b.����id And a.����id In (A����id_In, B����id_In)
    Order By b.��Ժ���� Desc, Nvl(b.��Ժ����, Sysdate) Desc;

  v_����id ������Ϣ.����id%Type;
  v_�ϲ�id ������Ϣ.����id%Type;
  v_����� ������Ϣ.�����%Type;
  v_סԺ�� ������Ϣ.סԺ��%Type;
  --����δ�����(���ﲿ��)
  Cursor c_Owe(v_����id ������Ϣ.����id%Type) Is
    Select ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, Sum(���) As ���
    From ����δ�����
    Where ��ҳid Is Null And ����id = v_����id
    Group By ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��;

  --�������
  Cursor c_Spare(v_����id ������Ϣ.����id%Type) Is
    Select ����, ����, Ԥ�����, ������� From ������� Where ����id = v_����id;

  --ҽ�����˵���
  Cursor c_Insure(v_����id ������Ϣ.����id%Type) Is
    Select * From �����ʻ� Where ����id = v_����id Order By ����;

  --Ҫ������ҽ�����˵���
  Cursor c_Keepinsure
  (
    v_����id ������Ϣ.����id%Type,
    v_����   ҽ�����˵���.����%Type
  ) Is
    Select * From �����ʻ� Where ����id = v_����id And ���� = v_����;
  r_Keepinsure c_Keepinsure%RowType;

  Cursor c_Year
  (
    v_����id ������Ϣ.����id%Type,
    v_����   ҽ�����˵���.����%Type
  ) Is
    Select * From �ʻ������Ϣ Where ����id = v_����id And ���� = v_����;

  v_ԭ��Ϣ   ���˺ϲ���¼.ԭ��Ϣ%Type;
  v_Count    Number;
  n_Readonly Number;
  v_Sql      Varchar2(1000);

  n_��ҳid       ������Ϣ.��ҳid%Type;
  v_Error        Varchar2(255);
  n_������       ���˵�����¼.������%Type;
  v_������       ������Ϣ.������%Type;
  n_��������     ���˵�����¼.��������%Type;
  n_Row          Number;
  n_��������     Number;
  n_ÿ����סԺ�� Number;
  n_Max��ҳid    Number;
  n_Cnt��ҳid    Number;
  n_Cur��ҳid    Number;
  n_CntסԺ����  Number;
  n_CurסԺ����  Number;
  n_MaxסԺ����  Number;
  n_Loop��ҳid   ������Ϣ.��ҳid%Type;
  v_Chgs         Varchar2(4000);

  n_Lengthb Number;
  Err_Custom Exception;
Begin
  Begin
    Select ֻ�� Into n_Readonly From zlBakSpaces Where ��ǰ = 1;
  Exception
    When Others Then
      Null;
  End;
  If n_Readonly = 1 Then
    n_Readonly := 0;
    For r_Bak In (Select a.���� Table_Name
                  From Zltools.Zlbaktables A, User_Constraints B
                  Where a.���� = b.Table_Name And b.r_Constraint_Name = '������Ϣ_PK' And b.Constraint_Type = 'R') Loop
      v_Sql := 'Select Count(����Id) From H' || r_Bak.Table_Name || ' Where ����Id In(:1,:2)';
      Execute Immediate v_Sql
        Into n_Readonly
        Using A����id_In, B����id_In;
      If n_Readonly > 0 Then
        v_Error := '������ֻ���ĵ�ǰת���ռ��������,���ܽ��кϲ�!';
        Raise Err_Custom;
      End If;
    End Loop;
  End If;

  --�����ﲡ�˺ϲ�������ZLHIS��һ��,��ֹ�����ﲡ����ZLHIS�ϲ����� 
  Select Count(1) Into v_Count From ���˹Һż�¼ A Where a.����id In (A����id_In, B����id_In) And a.���ӱ�־ = 3;
  If v_Count <> 0 Then
    v_Error := '���κϲ��Ĳ����к������ﲡ��,���ܽ��кϲ���';
    Raise Err_Custom;
  End If;

  --�������Ѽ�飺
  --1.ѡ����ͬһ������
  --2.����סԺ��������Ժ��ȴ��Ժ(������������Ժ)��
  --3.����סԺ���˵�סԺ�ڼ���ڽ�������
  --4.ҽ�����˴���δ�����

  --���������˲������������ҵ��
  Zl_������Ϣ_����(A����id_In, 1);
  Zl_������Ϣ_����(B����id_In, 1);

  Open c_Infoa;
  Fetch c_Infoa
    Into r_Infoa;
  If c_Infoa%RowCount = 0 Then
    Close c_Infoa;
    v_Error := 'û�з��ֱ��ϲ��Ĳ�����Ϣ��';
    Raise Err_Custom;
  End If;

  Open c_Infob;
  Fetch c_Infob
    Into r_Infob;
  If c_Infob%RowCount = 0 Then
    Close c_Infob;
    v_Error := 'û�з���Ҫ�����Ĳ�����Ϣ��';
    Raise Err_Custom;
  End If;

  --��ȡ������ز��˱�����
  For r_Patitable In c_Patitable Loop
    If r_Patitable.��ҳid = 0 Then
      Arronbase(Arronbase.Count + 1) := r_Patitable.Table_Name;
    Else
      Arronpage(Arronpage.Count + 1) := r_Patitable.Table_Name;
    End If;
  End Loop;

  --����סԺ���ȵǼǵĲ���ID��Ϊʵ����Ҫ�����Ĳ���ID
  If Nvl(ǿ�Ʊ���_In, 0) = 1 Then
    v_����id := B����id_In;
  Else
    Select ����id
    Into v_����id
    From (Select /*+ CHOOSE */
            a.����id
           From ������Ϣ A, ������ҳ B
           Where a.����id = b.����id(+) And a.����id In (A����id_In, B����id_In)
           Order By Nvl(b.��Ժ����, To_Date('3000-01-01', 'YYYY-MM-DD')), Nvl(b.��Ժ����, To_Date('3000-01-01', 'YYYY-MM-DD')),
                    a.�Ǽ�ʱ��, a.����id --סԺ��������
           )
    Where Rownum = 1;
  End If;

  --��ȷ�������ŵ�ģʽ
  Select Zl_To_Number(Nvl(zl_GetSysParameter(39), '0')) Into n_�������� From Dual;
  --סԺ��ģʽ
  Select Zl_To_Number(Nvl(zl_GetSysParameter(145), '0')) Into n_ÿ����סԺ�� From Dual;

  --����һ������ʵ�����Ҫɾ���Ĳ���ID
  If v_����id = A����id_In Then
    v_�ϲ�id := B����id_In;
    --����27445 ����ָ�����˵�����š�סԺ�š�ҽ����
    v_����� := Nvl(r_Infob.�����, r_Infoa.�����);
    v_סԺ�� := Nvl(r_Infob.סԺ��, r_Infoa.סԺ��);
  Else
    v_�ϲ�id := A����id_In;
    v_����� := Nvl(r_Infob.�����, r_Infoa.�����);
    v_סԺ�� := Nvl(r_Infob.סԺ��, r_Infoa.סԺ��);
  End If;

  ---��¼�ϲ�����,�ں�������r_PatiTable�Ѻϲ����˵ĺϲ���¼����Ϊ�������˵�
  v_ԭ��Ϣ := v_�ϲ�id || ',' || r_Infoa.����� || ',' || r_Infoa.סԺ�� || ',' || r_Infoa.���￨�� || ',' || r_Infoa.���� || ',' ||
           r_Infoa.�Ա� || ',' || r_Infoa.���� || ',' || To_Char(r_Infoa.��������, 'yyyy-mm-dd') || ',' || r_Infoa.���֤�� || ',' ||
           r_Infoa.����״�� || ',' || r_Infoa.ְҵ || ',' || r_Infoa.��ͥ��ַ;
  Insert Into ���˺ϲ���¼
    (����id, ԭ��Ϣ, �ϲ�ԭ��, ����Ա����, �ϲ�ʱ��)
  Values
    (v_����id, v_ԭ��Ϣ, �ϲ�ԭ��_In, ����Ա����_In, Sysdate);

  --��ʼ�ϲ�
  --84398�޸Ľ�סԺ��������������棬����Ҫ���������סԺ���˺ϲ�
  --10.34��ʼ,סԺ�������������ز���,�ϲ����סԺ����=��������סԺ����+�ϲ�����������Ժ�Ĵ���
  Select Nvl(סԺ����, 0) Into n_CurסԺ���� From ������Ϣ Where ����id = v_����id;
  Select Count(*) Into n_CntסԺ���� From ������ҳ Where ����id = v_�ϲ�id And ��ҳid <> 0 And �������� = 0;
  n_MaxסԺ���� := n_CurסԺ���� + n_CntסԺ����;
  --��������ҳ����(�漰����ID,��ҳID�ֶεı�)
  If (r_Infoa.��ҳid Is Not Null And r_Infob.��ҳid Is Not Null) Or (ǿ�Ʊ���_In = 1 And r_Infoa.��ҳid Is Not Null) Then
    If r_Infoa.��ҳid = 0 And r_Infob.��ҳid = 0 Then
      Close c_Infoa;
      Close c_Infob;
      v_Error := '����ԤԼ���˲��ܽ��в��˺ϲ�������';
      Raise Err_Custom;
    Elsif r_Infoa.��ҳid = 0 Then
      If r_Infob.��Ժ���� Is Not Null And r_Infob.��Ժ���� Is Null Then
        Close c_Infoa;
        Close c_Infob;
        v_Error := 'ԤԼ���˺���Ժ���˲��ܽ��в��˺ϲ�������';
        Raise Err_Custom;
      End If;
    Elsif r_Infob.��ҳid = 0 Then
      If r_Infoa.��Ժ���� Is Not Null And r_Infoa.��Ժ���� Is Null Then
        Close c_Infoa;
        Close c_Infob;
        v_Error := 'ԤԼ���˺���Ժ���˲��ܽ��в��˺ϲ�������';
        Raise Err_Custom;
      End If;
    End If;
    --�����������ܹ���סԺ�������
    Select Count(*) Into v_Count From ������ҳ Where ����id In (A����id_In, B����id_In) And ��ҳid <> 0;
    --��Ϊ10.19��ʼ����Ժʱ�����޸���ҳid�����������ҳID���ܴ����ܵ�סԺ�������
    Select Max(��ҳid) Into n_Max��ҳid From ������ҳ Where ����id = v_����id And ��ҳid <> 0;
    Select Count(*) Into n_Cnt��ҳid From ������ҳ Where ����id = v_�ϲ�id And ��ҳid <> 0;
    If n_Max��ҳid + n_Cnt��ҳid > v_Count Then
      v_Count := n_Max��ҳid + n_Cnt��ҳid;
    End If;
    --��ʵ��Ҫ���µ���ҳ����ֵ,��ǰ��v_Count >= n_Max��ҳid�жϴ���һ�����⣨�����������˶�ν�����Ժ�����ܵ���A,B���˲��־������û�и��£�
    Select Nvl(Max(��ҳid), 0)
    Into n_Loop��ҳid
    From ������ҳ A, (Select Min(��Ժ����) ��Ժ���� From ������ҳ Where ����id = v_�ϲ�id) B
    Where a.����id = v_����id And a.��Ժ���� < b.��Ժ����;
  
    For r_Merge In c_Mergepati Loop
      If Not (r_Merge.����id = v_����id And r_Merge.��ҳid = v_Count) And v_Count <> 0 Then
        --�ò�����ҳҪɾ��ʱ,�������ѱ�Ŀ�˵ġ�
        If r_Merge.��Ŀ���� Is Not Null Then
          Close c_Infoa;
          Close c_Infob;
          If r_Merge.��ǰסԺ�� Is Null Then
            v_Error := '����' || r_Merge.���� || '(����ID=' || r_Merge.����id || ')�����ѱ�Ŀ�Ĳ���,������ϲ��ò��ˡ�';
          Else
            v_Error := '����' || r_Merge.���� || '(����ID=' || r_Merge.����id || ',סԺ��=' || r_Merge.��ǰסԺ�� ||
                       ')�����ѱ�Ŀ�Ĳ���,������ϲ��ò��ˡ�';
          End If;
          Raise Err_Custom;
        End If;
        If v_Count >= Nvl(n_Loop��ҳid, 0) Then
          If r_Merge.��ҳid = 0 Then
            n_Cur��ҳid := 0;
            Update ������ҳ
            Set �������� = r_Merge.��������, ҽ�Ƹ��ʽ = r_Merge.ҽ�Ƹ��ʽ, �ѱ� = r_Merge.�ѱ�, ����Ժ = r_Merge.����Ժ,
                ��Ժ����id = r_Merge.��Ժ����id, ��Ժ����id = r_Merge.��Ժ����id, ��Ժ���� = r_Merge.��Ժ����, ��Ժ���� = r_Merge.��Ժ����,
                ��Ժ��ʽ = r_Merge.��Ժ��ʽ, ����Ժת�� = r_Merge.����Ժת��, סԺĿ�� = r_Merge.סԺĿ��, ��Ժ���� = r_Merge.��Ժ����,
                �Ƿ���� = r_Merge.�Ƿ����, ��ǰ���� = r_Merge.��ǰ����, ��ǰ����id = r_Merge.��ǰ����id, ����ȼ�id = r_Merge.����ȼ�id,
                ��Ժ����id = r_Merge.��Ժ����id, ��Ժ���� = r_Merge.��Ժ����, ��Ժ���� = r_Merge.��Ժ����, סԺ���� = r_Merge.סԺ����,
                ��Ժ��ʽ = r_Merge.��Ժ��ʽ, �Ƿ�ȷ�� = r_Merge.�Ƿ�ȷ��, ȷ������ = r_Merge.ȷ������, �·����� = r_Merge.�·�����, Ѫ�� = r_Merge.Ѫ��,
                ���ȴ��� = r_Merge.���ȴ���, �ɹ����� = r_Merge.�ɹ�����, �����־ = r_Merge.�����־, �������� = r_Merge.��������, ʬ���־ = r_Merge.ʬ���־,
                ����ҽʦ = r_Merge.����ҽʦ, ���λ�ʿ = r_Merge.���λ�ʿ, סԺҽʦ = r_Merge.סԺҽʦ, ��ĿԱ��� = r_Merge.��ĿԱ���,
                ��ĿԱ���� = r_Merge.��ĿԱ����, ��Ŀ���� = r_Merge.��Ŀ����, ״̬ = r_Merge.״̬, ���ú� = r_Merge.���ú�, ���� = r_Merge.����,
                �Ա� = r_Merge.�Ա�, ���� = r_Merge.����, ����״�� = r_Merge.����״��, ְҵ = r_Merge.ְҵ, ���� = r_Merge.����, ѧ�� = r_Merge.ѧ��,
                ��λ�绰 = r_Merge.��λ�绰, ��λ�ʱ� = r_Merge.��λ�ʱ�, ��λ��ַ = r_Merge.��λ��ַ, ���� = r_Merge.����, ��ͥ��ַ = r_Merge.��ͥ��ַ,
                ��ͥ�绰 = r_Merge.��ͥ�绰, ��ͥ��ַ�ʱ� = r_Merge.��ͥ��ַ�ʱ�, ���ڵ�ַ = r_Merge.���ڵ�ַ, ���ڵ�ַ�ʱ� = r_Merge.���ڵ�ַ�ʱ�,
                ��ϵ������ = r_Merge.��ϵ������, ��ϵ�˹�ϵ = r_Merge.��ϵ�˹�ϵ, ��ϵ�˵�ַ = r_Merge.��ϵ�˵�ַ, ��ϵ�˵绰 = r_Merge.��ϵ�˵绰,
                ��ҽ������� = r_Merge.��ҽ�������, �Ǽ��� = r_Merge.�Ǽ���, �Ǽ�ʱ�� = r_Merge.�Ǽ�ʱ��, ���� = r_Merge.����, ��˱�־ = r_Merge.��˱�־,
                �Ƿ��ϴ� = r_Merge.�Ƿ��ϴ�, ��ע = r_Merge.��ע, ����ת�� = r_Merge.����ת��, ������ = r_Merge.������,
                סԺ�� = Decode(n_ÿ����סԺ��, 1, r_Merge.סԺ��, v_סԺ��), ���ۺ� = r_Merge.���ۺ�, �������� = r_Merge.��������,
                ���ʱ�� = r_Merge.���ʱ��, ·��״̬ = r_Merge.·��״̬, ������ = r_Merge.������, Ӥ������id = r_Merge.Ӥ������id,
                Ӥ������id = r_Merge.Ӥ������id, ĸӤת�Ʊ�־ = r_Merge.ĸӤת�Ʊ�־, ҽ������ʱ�� = r_Merge.ҽ������ʱ��
            Where ����id = v_����id And ��ҳid = n_Cur��ҳid;
            If Sql%RowCount = 0 Then
              Insert Into ������ҳ
                (����id, ��ҳid, ��������, ҽ�Ƹ��ʽ, �ѱ�, ����Ժ, ��Ժ����id, ��Ժ����id, ��Ժ����, ��Ժ����, ��Ժ��ʽ, ����Ժת��, סԺĿ��, ��Ժ����, �Ƿ����, ��ǰ����,
                 ��ǰ����id, ����ȼ�id, ��Ժ����id, ��Ժ����, ��Ժ����, סԺ����, ��Ժ��ʽ, �Ƿ�ȷ��, ȷ������, �·�����, Ѫ��, ���ȴ���, �ɹ�����, �����־, ��������, ʬ���־,
                 ����ҽʦ, ���λ�ʿ, סԺҽʦ, ��ĿԱ���, ��ĿԱ����, ��Ŀ����, ״̬, ���ú�, ����, �Ա�, ����, ����״��, ְҵ, ����, ѧ��, ��λ�绰, ��λ�ʱ�, ��λ��ַ, ����, ��ͥ��ַ,
                 ��ͥ�绰, ��ͥ��ַ�ʱ�, ���ڵ�ַ, ���ڵ�ַ�ʱ�, ��ϵ������, ��ϵ�˹�ϵ, ��ϵ�˵�ַ, ��ϵ�˵绰, ��ҽ�������, �Ǽ���, �Ǽ�ʱ��, ����, ��˱�־, �Ƿ��ϴ�, ��ע, ����ת��,
                 ������, סԺ��, ���ۺ�, ��������, ���ʱ��, ·��״̬, ������, Ӥ������id, Ӥ������id, ĸӤת�Ʊ�־, ҽ������ʱ��)
              Values
                (v_����id, n_Cur��ҳid, r_Merge.��������, r_Merge.ҽ�Ƹ��ʽ, r_Merge.�ѱ�, r_Merge.����Ժ, r_Merge.��Ժ����id,
                 r_Merge.��Ժ����id, r_Merge.��Ժ����, r_Merge.��Ժ����, r_Merge.��Ժ��ʽ, r_Merge.����Ժת��, r_Merge.סԺĿ��, r_Merge.��Ժ����,
                 r_Merge.�Ƿ����, r_Merge.��ǰ����, r_Merge.��ǰ����id, r_Merge.����ȼ�id, r_Merge.��Ժ����id, r_Merge.��Ժ����, r_Merge.��Ժ����,
                 r_Merge.סԺ����, r_Merge.��Ժ��ʽ, r_Merge.�Ƿ�ȷ��, r_Merge.ȷ������, r_Merge.�·�����, r_Merge.Ѫ��, r_Merge.���ȴ���,
                 r_Merge.�ɹ�����, r_Merge.�����־, r_Merge.��������, r_Merge.ʬ���־, r_Merge.����ҽʦ, r_Merge.���λ�ʿ, r_Merge.סԺҽʦ,
                 r_Merge.��ĿԱ���, r_Merge.��ĿԱ����, r_Merge.��Ŀ����, r_Merge.״̬, r_Merge.���ú�, r_Merge.����, r_Merge.�Ա�, r_Merge.����,
                 r_Merge.����״��, r_Merge.ְҵ, r_Merge.����, r_Merge.ѧ��, r_Merge.��λ�绰, r_Merge.��λ�ʱ�, r_Merge.��λ��ַ, r_Merge.����,
                 r_Merge.��ͥ��ַ, r_Merge.��ͥ�绰, r_Merge.��ͥ��ַ�ʱ�, r_Merge.���ڵ�ַ, r_Merge.���ڵ�ַ�ʱ�, r_Merge.��ϵ������, r_Merge.��ϵ�˹�ϵ,
                 r_Merge.��ϵ�˵�ַ, r_Merge.��ϵ�˵绰, r_Merge.��ҽ�������, r_Merge.�Ǽ���, r_Merge.�Ǽ�ʱ��, r_Merge.����, r_Merge.��˱�־,
                 r_Merge.�Ƿ��ϴ�, r_Merge.��ע, r_Merge.����ת��, r_Merge.������, Decode(n_ÿ����סԺ��, 1, r_Merge.סԺ��, v_סԺ��),
                 r_Merge.���ۺ�, r_Merge.��������, r_Merge.���ʱ��, r_Merge.·��״̬, r_Merge.������, r_Merge.Ӥ������id, r_Merge.Ӥ������id,
                 r_Merge.ĸӤת�Ʊ�־, r_Merge.ҽ������ʱ��);
            End If;
          Else
            n_Cur��ҳid := v_Count;
            Insert Into ������ҳ
              (����id, ��ҳid, ��������, ҽ�Ƹ��ʽ, �ѱ�, ����Ժ, ��Ժ����id, ��Ժ����id, ��Ժ����, ��Ժ����, ��Ժ��ʽ, ����Ժת��, סԺĿ��, ��Ժ����, �Ƿ����, ��ǰ����,
               ��ǰ����id, ����ȼ�id, ��Ժ����id, ��Ժ����, ��Ժ����, סԺ����, ��Ժ��ʽ, �Ƿ�ȷ��, ȷ������, �·�����, Ѫ��, ���ȴ���, �ɹ�����, �����־, ��������, ʬ���־, ����ҽʦ,
               ���λ�ʿ, סԺҽʦ, ��ĿԱ���, ��ĿԱ����, ��Ŀ����, ״̬, ���ú�, ����, �Ա�, ����, ����״��, ְҵ, ����, ѧ��, ��λ�绰, ��λ�ʱ�, ��λ��ַ, ����, ��ͥ��ַ, ��ͥ�绰,
               ��ͥ��ַ�ʱ�, ���ڵ�ַ, ���ڵ�ַ�ʱ�, ��ϵ������, ��ϵ�˹�ϵ, ��ϵ�˵�ַ, ��ϵ�˵绰, ��ҽ�������, �Ǽ���, �Ǽ�ʱ��, ����, ��˱�־, �Ƿ��ϴ�, ��ע, ����ת��, ������, סԺ��,
               ���ۺ�, ��������, ���ʱ��, ·��״̬, ������, Ӥ������id, Ӥ������id, ĸӤת�Ʊ�־, ҽ������ʱ��)
            Values
              (v_����id, n_Cur��ҳid, r_Merge.��������, r_Merge.ҽ�Ƹ��ʽ, r_Merge.�ѱ�, r_Merge.����Ժ, r_Merge.��Ժ����id, r_Merge.��Ժ����id,
               r_Merge.��Ժ����, r_Merge.��Ժ����, r_Merge.��Ժ��ʽ, r_Merge.����Ժת��, r_Merge.סԺĿ��, r_Merge.��Ժ����, r_Merge.�Ƿ����,
               r_Merge.��ǰ����, r_Merge.��ǰ����id, r_Merge.����ȼ�id, r_Merge.��Ժ����id, r_Merge.��Ժ����, r_Merge.��Ժ����, r_Merge.סԺ����,
               r_Merge.��Ժ��ʽ, r_Merge.�Ƿ�ȷ��, r_Merge.ȷ������, r_Merge.�·�����, r_Merge.Ѫ��, r_Merge.���ȴ���, r_Merge.�ɹ�����,
               r_Merge.�����־, r_Merge.��������, r_Merge.ʬ���־, r_Merge.����ҽʦ, r_Merge.���λ�ʿ, r_Merge.סԺҽʦ, r_Merge.��ĿԱ���,
               r_Merge.��ĿԱ����, r_Merge.��Ŀ����, r_Merge.״̬, r_Merge.���ú�, r_Merge.����, r_Merge.�Ա�, r_Merge.����, r_Merge.����״��,
               r_Merge.ְҵ, r_Merge.����, r_Merge.ѧ��, r_Merge.��λ�绰, r_Merge.��λ�ʱ�, r_Merge.��λ��ַ, r_Merge.����, r_Merge.��ͥ��ַ,
               r_Merge.��ͥ�绰, r_Merge.��ͥ��ַ�ʱ�, r_Merge.���ڵ�ַ, r_Merge.���ڵ�ַ�ʱ�, r_Merge.��ϵ������, r_Merge.��ϵ�˹�ϵ, r_Merge.��ϵ�˵�ַ,
               r_Merge.��ϵ�˵绰, r_Merge.��ҽ�������, r_Merge.�Ǽ���, r_Merge.�Ǽ�ʱ��, r_Merge.����, r_Merge.��˱�־, r_Merge.�Ƿ��ϴ�,
               r_Merge.��ע, r_Merge.����ת��, r_Merge.������, Decode(n_ÿ����סԺ��, 1, r_Merge.סԺ��, v_סԺ��), r_Merge.���ۺ�, r_Merge.��������,
               r_Merge.���ʱ��, r_Merge.·��״̬, r_Merge.������, r_Merge.Ӥ������id, r_Merge.Ӥ������id, r_Merge.ĸӤת�Ʊ�־, r_Merge.ҽ������ʱ��);
          End If;
        Else
          Exit;
        End If;

        ---- v_����id,n_Cur��ҳid:r_Merge.����id, r_Merge.��ҳid
        v_Chgs := v_Chgs || ';' || v_����id || ',' || n_Cur��ҳid || ':' || r_Merge.����id || ',' || r_Merge.��ҳid;
      
        --���²�����ر�Ĳ���ָ��
        ---------------------------------------------------------------
        --���˱䶯��¼
        Update ���˱䶯��¼
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
        
		--�����Զ�����
        Update �����Զ�����
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;

        --������ҳ�ӱ�
        Update ������ҳ�ӱ�
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
      
        --סԺ���ü�¼
        Update סԺ���ü�¼
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid,
            ��ʶ�� = Nvl(Decode(�����־, 1, v_�����, Decode(n_ÿ����סԺ��, 1, r_Merge.סԺ��, v_סԺ��)), ��ʶ��)
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
        Update HסԺ���ü�¼
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid,
            ��ʶ�� = Nvl(Decode(�����־, 1, v_�����, Decode(n_ÿ����סԺ��, 1, r_Merge.סԺ��, v_סԺ��)), ��ʶ��)
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
        --������ü�¼
        --Update ������ü�¼
        --Set ����id = v_����id,
        --    ��ʶ�� = Nvl(Decode(�����־, 1, v_�����, Decode(n_ÿ����סԺ��, 1, r_Merge.סԺ��, v_סԺ��)), ��ʶ��)
        --Where ����id = r_Merge.����id;
        --Update H������ü�¼
        --Set ����id = v_����id,
        --    ��ʶ�� = Nvl(Decode(�����־, 1, v_�����, Decode(n_ÿ����סԺ��, 1, r_Merge.סԺ��, v_סԺ��)), ��ʶ��)
        --Where ����id = r_Merge.����id;
      
        --����Ԥ����¼
        Update ����Ԥ����¼
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
        Update H����Ԥ����¼
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
      
        --����δ�����
        Update ����δ�����
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
      
        --δ��ҩƷ��¼
        Update δ��ҩƷ��¼
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
      
        --������
        Update ������
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
      
        --���ս����¼(����ID�ͷ�סԺ����һ���ں��洦��)
        Update ���ս����¼ Set ��ҳid = n_Cur��ҳid Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
      
        --����ģ�����
        Update ����ģ�����
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
      
        --����ҽ����¼(ZLHIS+)
        Update ����ҽ����¼
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
        Update H����ҽ����¼
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
      
        --���˹�����¼(ZLHIS+)
        Update ���˹�����¼
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
        Update H���˹�����¼
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
      
        --������ϼ�¼(ZLHIS+)
        Update ������ϼ�¼
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
        Update H������ϼ�¼
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
      
        --���������¼(ZLHIS+)
        Update ���������¼
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
        Update H���������¼
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
      
        --���˵�����¼(zlhis+)
        Update ���˵�����¼
        Set ����id = v_����id, ��ҳid = n_Cur��ҳid
        Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
      
        --����ϵͳ�ı�
        Begin
          v_Sql := 'Update ���˷��� Set ����ID=:1,��ҳID=:2 Where ����ID=:3 And ��ҳID=:4';
          Execute Immediate v_Sql
            Using v_����id, n_Cur��ҳid, r_Merge.����id, r_Merge.��ҳid;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Update �����¼ Set ����ID=:1,��ҳID=:2 Where ����ID=:3 And ��ҳID=:4';
          Execute Immediate v_Sql
            Using v_����id, n_Cur��ҳid, r_Merge.����id, r_Merge.��ҳid;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Update ��Ϸ������ Set ����ID=:1,��ҳID=:2 Where ����ID=:3 And ��ҳID=:4';
          Execute Immediate v_Sql
            Using v_����id, n_Cur��ҳid, r_Merge.����id, r_Merge.��ҳid;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Update �������ֽ�� Set ����ID=:1,��ҳID=:2 Where ����ID=:3 And ��ҳID=:4';
          Execute Immediate v_Sql
            Using v_����id, n_Cur��ҳid, r_Merge.����id, r_Merge.��ҳid;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Insert Into ���˷�����Ϣ(����ID,��ҳID,̥������,���䷽ʽ,����̥λ,�������,����ȱ��,Ӥ���Ա�,Ӥ������,Apgar����) ' ||
                   'Select :1,:2,̥������,���䷽ʽ,����̥λ,�������,����ȱ��,Ӥ���Ա�,Ӥ������,Apgar���� From ���˷�����Ϣ Where ����ID=:3 And ��ҳID=:4';
          Execute Immediate v_Sql
            Using v_����id, n_Cur��ҳid, r_Merge.����id, r_Merge.��ҳid;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Delete From ���˷�����Ϣ Where ����ID=:1 And ��ҳID=:2';
          Execute Immediate v_Sql
            Using r_Merge.����id, r_Merge.��ҳid;
        Exception
          When Others Then
            Null;
        End;
      
        Begin
          v_Sql := 'Update ���ļ�¼ Set ����ID=:1,��ҳID=:2 Where ����ID=:3 And ��ҳID=:4';
          Execute Immediate v_Sql
            Using v_����id, n_Cur��ҳid, r_Merge.����id, r_Merge.��ҳid;
        Exception
          When Others Then
            Null;
        End;
      
        --����������ҳ��ر�
        For v_Loop In 1 .. Arronpage.Count Loop
          v_Sql := 'Update ' || Arronpage(v_Loop) || ' Set ����ID=:1,��ҳID=:2 Where ����ID=:3 And ��ҳID=:4';
          Execute Immediate v_Sql
            Using v_����id, n_Cur��ҳid, r_Merge.����id, r_Merge.��ҳid;
        End Loop;
      
        --ɾ���ѵ�����Ĳ�����ҳ
        Delete From ������ҳ Where ����id = r_Merge.����id And ��ҳid = r_Merge.��ҳid;
      End If;
      If r_Merge.��ҳid <> 0 Then
        v_Count := v_Count - 1;
      End If;
    End Loop;
  End If;

  --���漰��ҳID���ݵĸ���(����ҳID����ҳID����Ϊ��)
  ---------------------------------------------------------------
  --סԺ���ü�¼
  Update סԺ���ü�¼
  Set ����id = v_����id, ��ʶ�� = Nvl(Decode(�����־, 2, v_סԺ��, v_�����), ��ʶ��)
  Where ����id = v_�ϲ�id;
  Update HסԺ���ü�¼
  Set ����id = v_����id, ��ʶ�� = Nvl(Decode(�����־, 2, v_סԺ��, v_�����), ��ʶ��)
  Where ����id = v_�ϲ�id;
  --������ü�¼
  Update ������ü�¼
  Set ����id = v_����id, ��ʶ�� = Nvl(Decode(�����־, 2, v_סԺ��, v_�����), ��ʶ��)
  Where ����id = v_�ϲ�id;
  Update H������ü�¼
  Set ����id = v_����id, ��ʶ�� = Nvl(Decode(�����־, 2, v_סԺ��, v_�����), ��ʶ��)
  Where ����id = v_�ϲ�id;

  --����Ԥ����¼
  Update ����Ԥ����¼ Set ����id = v_����id Where ����id = v_�ϲ�id And ��ҳid Is Null;
  Update H����Ԥ����¼ Set ����id = v_����id Where ����id = v_�ϲ�id And ��ҳid Is Null;

  --δ��ҩƷ��¼
  Update δ��ҩƷ��¼ Set ����id = v_����id Where ����id = v_�ϲ�id And ��ҳid Is Null;

  --������
  Update ������ Set ����id = v_����id Where ����id = v_�ϲ�id And ��ҳid Is Null;

  --����ҽ����¼(ZLHIS+)
  Update ����ҽ����¼ Set ����id = v_����id Where ����id = v_�ϲ�id And ��ҳid Is Null;
  Update H����ҽ����¼ Set ����id = v_����id Where ����id = v_�ϲ�id And ��ҳid Is Null;

  --���˹�����¼(ZLHIS+):��ҳID�����ǹҺ�ID
  Update ���˹�����¼ Set ����id = v_����id Where ����id = v_�ϲ�id;
  Update H���˹�����¼ Set ����id = v_����id Where ����id = v_�ϲ�id;

  --������ϼ�¼(ZLHIS+):��ҳID�����ǹҺ�ID
  Update ������ϼ�¼ Set ����id = v_����id Where ����id = v_�ϲ�id;
  Update H������ϼ�¼ Set ����id = v_����id Where ����id = v_�ϲ�id;

  --���������¼(ZLHIS+)
  Update ���������¼ Set ����id = v_����id Where ����id = v_�ϲ�id And ��ҳid Is Null;
  Update H���������¼ Set ����id = v_����id Where ����id = v_�ϲ�id And ��ҳid Is Null;

  --���˹Һż�¼(ZLHIS+)
  Update ���˹Һż�¼ Set ����id = v_����id, ����� = Nvl(v_�����, �����) Where ����id = v_�ϲ�id;

  --���˽��ʼ�¼
  Update ���˽��ʼ�¼ Set ����id = v_����id Where ����id = v_�ϲ�id;
  Update H���˽��ʼ�¼ Set ����id = v_����id Where ����id = v_�ϲ�id;

  --��λ״����¼
  Update ��λ״����¼ Set ����id = v_����id Where ����id = v_�ϲ�id;

  --���˵�����¼
  Update ���˵�����¼ Set ����id = v_����id Where ����id = v_�ϲ�id;
  --���ⲡ��
  Select Count(*) Into v_Count From ���ⲡ�� Where ����id = v_����id;
  If v_Count = 0 Then
    Update ���ⲡ�� Set ����id = v_����id Where ����id = v_�ϲ�id;
  Else
    Delete From ���ⲡ�� Where ����id = v_�ϲ�id;
  End If;

  --����δ�����
  For r_Owe In c_Owe(v_�ϲ�id) Loop
    Update ����δ�����
    Set ��� = Nvl(���, 0) + Nvl(r_Owe.���, 0)
    Where ��ҳid Is Null And ����id = v_����id And Nvl(���˲���id, 0) = Nvl(r_Owe.���˲���id, 0) And
          Nvl(���˿���id, 0) = Nvl(r_Owe.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Owe.��������id, 0) And
          Nvl(ִ�в���id, 0) = Nvl(r_Owe.ִ�в���id, 0) And Nvl(������Ŀid, 0) = Nvl(r_Owe.������Ŀid, 0) And
          Nvl(��Դ;��, 0) = Nvl(r_Owe.��Դ;��, 0);
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (v_����id, Null, r_Owe.���˲���id, r_Owe.���˿���id, r_Owe.��������id, r_Owe.ִ�в���id, r_Owe.������Ŀid, r_Owe.��Դ;��, r_Owe.���);
    End If;
  End Loop;
  Delete From ����δ����� Where ����id = v_�ϲ�id;
  Delete From ����δ����� Where ����id = v_����id And Nvl(���, 0) = 0;

  --�������
  For r_Spare In c_Spare(v_�ϲ�id) Loop
    Update �������
    Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(r_Spare.Ԥ�����, 0), ������� = Nvl(�������, 0) + Nvl(r_Spare.�������, 0)
    Where Nvl(����, 0) = Nvl(r_Spare.����, 0) And ����id = v_����id And ���� = Nvl(r_Spare.����, 2);
    If Sql%RowCount = 0 Then
      Insert Into �������
        (����id, ����, ����, Ԥ�����, �������)
      Values
        (v_����id, r_Spare.����, Nvl(r_Spare.����, 2), r_Spare.Ԥ�����, r_Spare.�������);
    End If;
  End Loop;
  Delete From ������� Where ����id = v_�ϲ�id;
  Delete From ������� Where ����id = v_����id And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0 And ���� = 1;

  --���˹���ҩ��
  Insert Into ���˹���ҩ��
    (����id, ����ҩ��id, ����ҩ��)
    Select v_����id, ����ҩ��id, ����ҩ��
    From ���˹���ҩ��
    Where ����id = v_�ϲ�id And ����ҩ��id Not In (Select ����ҩ��id From ���˹���ҩ�� Where ����id = v_����id);
  Delete From ���˹���ҩ�� Where ����id = v_�ϲ�id;

  --����������Ϣ
  Insert Into ����������Ϣ
    (����id, ����, ������, ��־, ��������, ����ʱ��)
    Select v_����id, ����, ������, ��־, ��������, ����ʱ��
    From ����������Ϣ
    Where ����id = v_�ϲ�id And ���� Not In (Select ���� From ����������Ϣ Where ����id = v_����id);
  Delete From ����������Ϣ Where ����id = v_�ϲ�id;

  --�������߼�¼
  Insert Into �������߼�¼
    (����id, ����ʱ��, ��������)
    Select v_����id, a.����ʱ��, a.��������
    From �������߼�¼ A
    Where a.����id = v_�ϲ�id And Not Exists (Select 1 From �������߼�¼ Where ����id = v_����id And ����ʱ�� = a.����ʱ��);
  Delete From �������߼�¼ Where ����id = v_�ϲ�id;

  --������Ϣ�ӱ�
  Insert Into ������Ϣ�ӱ�
    (����id, ��Ϣ��, ��Ϣֵ, ����id)
    Select v_����id, a.��Ϣ��, a.��Ϣֵ, a.����id
    From ������Ϣ�ӱ� A
    Where a.����id = v_�ϲ�id And Not Exists (Select 1
           From ������Ϣ�ӱ�
           Where ����id = v_����id And ��Ϣ�� = a.��Ϣ�� And Nvl(����id, 0) = Nvl(a.����id, 0));
  Delete From ������Ϣ�ӱ� Where ����id = v_�ϲ�id;

  --����ҽ�ƿ�����
  Insert Into ����ҽ�ƿ�����
    (����id, �����id, ����, ��Ϣ��, ��Ϣֵ)
    Select v_����id, a.�����id, a.����, a.��Ϣ��, a.��Ϣֵ
    From ����ҽ�ƿ����� A
    Where a.����id = v_�ϲ�id And Not Exists (Select 1
           From ����ҽ�ƿ�����
           Where ����id = v_����id And �����id = a.�����id And ���� = a.���� And ��Ϣ�� = a.��Ϣ��);
  Delete From ����ҽ�ƿ����� Where ����id = v_�ϲ�id;

  --���ﲡ����¼
  Select Count(*) Into v_Count From ���ﲡ����¼ Where ����id = v_����id;
  If v_Count = 0 Then
    Select Count(*) Into v_Count From ���ﲡ����¼ Where ����id = v_�ϲ�id;
    If v_Count > 0 Then
      Update ���ﲡ����¼ Set ����id = v_����id Where ����id = v_�ϲ�id;
    End If;
  Else
    Delete From ���ﲡ����¼ Where ����id = v_�ϲ�id;
  End If;

  --סԺ������¼
  Select Count(*) Into v_Count From סԺ������¼ Where ����id = v_����id;

  If v_Count = 0 Then
    Select Count(*) Into v_Count From סԺ������¼ Where ����id = v_�ϲ�id;
    If v_Count > 0 Then
      Update סԺ������¼ Set ����id = v_����id Where ����id = v_�ϲ�id;
    End If;
  Else
    Begin
      v_Sql := 'Delete From ���ļ�¼ Where ����ID=:1';
      Execute Immediate v_Sql
        Using v_�ϲ�id;
    Exception
      When Others Then
        Null;
    End;
  
    Delete From סԺ������¼ Where ����id = v_�ϲ�id;
  End If;

  --ҽ��������ش���
  --��ʹ�ϲ������Ĳ��˵�ǰ����ҽ���ʻ�,ֻҪ����ҽ���ʻ�,���಻ͬҲ���ܺϲ�
  Select Count(Distinct ����) Into v_Count From ҽ�����˹����� Where ����id In (v_�ϲ�id, v_����id);
  If v_Count = 2 Then
    Close c_Infoa;
    Close c_Infob;
    v_Error := '�������˷ֱ����ڲ�ͬ�ı�����𣬲�����ϲ���';
    Raise Err_Custom;
  End If;

  Select Count(*) Into v_Count From ҽ�����˹����� Where ����id = v_�ϲ�id And ��־ = 0;
  --a.�ϲ��Ĳ�����ǰ��ҽ���ʻ�,���ڲ���
  If v_Count > 0 Then
    Select Count(*) Into v_Count From ҽ�����˹����� Where ����id = v_����id;
    --a.1�����Ĳ���������ҽ���ʻ�
    --a.2.1�����Ĳ������ڲ���ҽ���ʻ�,��ǰ��,��a.1��ͬ����
    If v_Count > 0 Then
      Delete From �ʻ������Ϣ Where ����id = v_�ϲ�id;
    
      Select Count(Distinct ҽ����) Into v_Count From ҽ�����˹����� Where ����id In (v_�ϲ�id, v_����id);
      If v_Count <> 2 Then
        --��������ҽ������ͬʱ,���ô���ҽ�����˵���
        For r_Insure In c_Insure(v_�ϲ�id) Loop
          --���ϲ��Ĳ��˿��ܹ����˶��ҽ������,��Ϊ�����������Ĳ�����
          --����27445 ����ָ�����˵�����š�סԺ�š�ҽ����
          If v_�ϲ�id = B����id_In Then
            Update ҽ�����˹�����
            Set ҽ���� =
                 (Select ҽ���� From ҽ�����˹����� Where ����id = v_�ϲ�id), ��־ = 0
            Where ���� = r_Insure.���� And ���� = r_Insure.���� And ҽ���� = r_Insure.ҽ���� And ����id <> v_�ϲ�id;
          Else
            Update ҽ�����˹�����
            Set ҽ���� =
                 (Select ҽ���� From ҽ�����˹����� Where ����id = v_����id), ��־ = 0
            Where ���� = r_Insure.���� And ���� = r_Insure.���� And ҽ���� = r_Insure.ҽ���� And ����id <> v_�ϲ�id;
          End If;
          --�ϲ��Ĳ������ڲ���ҽ��,��ʹ���û�ָ��Ҫ�����ò���,Ҳ�����������ʻ���Ϣ
          Delete From ҽ�����˵��� Where ���� = r_Insure.���� And ҽ���� = r_Insure.ҽ����;
        End Loop;
      End If;
      Delete From ҽ�����˹����� Where ����id = v_�ϲ�id;
    Else
      --a.2.2�����Ĳ������ں���ǰ������ҽ���ʻ�
      Update �ʻ������Ϣ Set ����id = v_����id Where ����id = v_�ϲ�id;
      Update ҽ�����˹����� Set ����id = v_����id Where ����id = v_�ϲ�id;
      --ҽ�����˵������ô���,��Ϊͨ��ҽ���Ź���<ҽ�����˹�����>
    End If;
  Else
    Select Count(*) Into v_Count From ҽ�����˹����� Where ����id = v_�ϲ�id And ��־ = 1;
    --b.�ϲ��Ĳ���������ҽ���ʻ�
    If v_Count > 0 Then
      Select Count(*) Into v_Count From ҽ�����˹����� Where ����id = v_����id;
      --b.1�����Ĳ�������Ҳ��ҽ���ʻ�
      --b.2.1�����Ĳ������ڲ���ҽ���ʻ�,��ǰ��,��b.1��ͬ����
      If v_Count > 0 Then
        For r_Insure In c_Insure(v_�ϲ�id) Loop
          --ת���ʻ������Ϣ
          For r_Year In c_Year(v_�ϲ�id, r_Insure.����) Loop
            Update �ʻ������Ϣ
            Set �ʻ������ۼ� = Nvl(�ʻ������ۼ�, 0) + Nvl(r_Year.�ʻ������ۼ�, 0), �ʻ�֧���ۼ� = Nvl(�ʻ�֧���ۼ�, 0) + Nvl(r_Year.�ʻ�֧���ۼ�, 0),
                ����ͳ���ۼ� = Nvl(����ͳ���ۼ�, 0) + Nvl(r_Year.����ͳ���ۼ�, 0), ͳ�ﱨ���ۼ� = Nvl(ͳ�ﱨ���ۼ�, 0) + Nvl(r_Year.ͳ�ﱨ���ۼ�, 0),
                סԺ�����ۼ� = Nvl(סԺ�����ۼ�, 0) + Nvl(r_Year.סԺ�����ۼ�, 0), ���ͳ���ۼ� = Nvl(���ͳ���ۼ�, 0) + Nvl(r_Year.���ͳ���ۼ�, 0),
                �����ۼ� = Nvl(�����ۼ�, 0) + Nvl(r_Year.�����ۼ�, 0), �������� = Nvl(��������, r_Year.��������),
                ����ͳ���޶� = Nvl(����ͳ���޶�, r_Year.����ͳ���޶�), ���ͳ���޶� = Nvl(���ͳ���޶�, r_Year.���ͳ���޶�), ������Ϣ = Nvl(������Ϣ, r_Year.������Ϣ)
            Where ����id = v_����id And ���� = r_Insure.���� And ��� = r_Year.���;
            If Sql%RowCount = 0 Then
              Insert Into �ʻ������Ϣ
                (����id, ����, ���, �ʻ������ۼ�, �ʻ�֧���ۼ�, ����ͳ���ۼ�, ͳ�ﱨ���ۼ�, סԺ�����ۼ�, ��������, ����ͳ���޶�, ���ͳ���޶�, �����ۼ�, ���ͳ���ۼ�, ������Ϣ)
              Values
                (v_����id, r_Insure.����, r_Year.���, r_Year.�ʻ������ۼ�, r_Year.�ʻ�֧���ۼ�, r_Year.����ͳ���ۼ�, r_Year.ͳ�ﱨ���ۼ�,
                 r_Year.סԺ�����ۼ�, r_Year.��������, r_Year.����ͳ���޶�, r_Year.���ͳ���޶�, r_Year.�����ۼ�, r_Year.���ͳ���ۼ�, r_Year.������Ϣ);
            End If;
          End Loop;
          Delete From �ʻ������Ϣ Where ����id = v_�ϲ�id;
        
          Select Count(Distinct ҽ����) Into v_Count From ҽ�����˹����� Where ����id In (v_�ϲ�id, v_����id);
          If v_Count <> 2 Then
            --��������ҽ������ͬʱ,���ô���ҽ�����˵���
            If v_�ϲ�id = B����id_In Then
              Update ҽ�����˹�����
              Set ��־ = 0
              Where (����, ����, ҽ����) In (Select ����, ����, ҽ���� From ҽ�����˹����� Where ����id = v_����id);
              Update ҽ�����˹����� Set ��־ = 1 Where ����id = v_����id;
            End If;
            Delete From ҽ�����˹����� Where ����id = v_�ϲ�id;
          Else
            --���ϲ��Ĳ��˿��ܹ����˶��ҽ������,��Ϊ�����������Ĳ�����
            --����27445 ����ָ�����˵�����š�סԺ�š�ҽ����
            If v_�ϲ�id = B����id_In Then
              Update ҽ�����˹�����
              Set ҽ���� =
                   (Select ҽ���� From ҽ�����˹����� Where ����id = v_�ϲ�id), ��־ = 0
              Where ���� = r_Insure.���� And ���� = r_Insure.���� And ҽ���� = r_Insure.ҽ���� And ����id <> v_�ϲ�id;
            Else
              Update ҽ�����˹�����
              Set ҽ���� =
                   (Select ҽ���� From ҽ�����˹����� Where ����id = v_����id), ��־ = 0
              Where ���� = r_Insure.���� And ���� = r_Insure.���� And ҽ���� = r_Insure.ҽ���� And ����id <> v_�ϲ�id;
            End If;
            --�ݴ��û�ָ��Ҫ�������˵��ʻ���Ϣ
            If v_�ϲ�id = B����id_In Then
              Open c_Keepinsure(B����id_In, r_Insure.����);
              Fetch c_Keepinsure
                Into r_Keepinsure;
            End If;
          
            Delete From ҽ�����˹����� Where ����id = v_�ϲ�id;
            Delete From ҽ�����˵��� Where ���� = r_Insure.���� And ҽ���� = r_Insure.ҽ����;
          
            --�����û�ָ��Ҫ�������˵��ʻ���Ϣ
            If v_�ϲ�id = B����id_In Then
              If c_Keepinsure%RowCount > 0 Then
                Update ҽ�����˵���
                Set ���� = r_Keepinsure.����, ҽ���� = r_Keepinsure.ҽ����, ���� = r_Keepinsure.����, ��Ա��� = r_Keepinsure.��Ա���,
                    ��λ���� = r_Keepinsure.��λ����, ˳��� = r_Keepinsure.˳���, ����֤�� = r_Keepinsure.����֤��, �ʻ���� = r_Keepinsure.�ʻ����,
                    ��ǰ״̬ = r_Keepinsure.��ǰ״̬, ����id = r_Keepinsure.����id, ��ְ = r_Keepinsure.��ְ, ����� = r_Keepinsure.�����,
                    �Ҷȼ� = r_Keepinsure.�Ҷȼ�, ����ʱ�� = r_Keepinsure.����ʱ��
                Where (����, ����, ҽ����) In (Select ����, ����, ҽ���� From ҽ�����˹����� Where ����id = v_����id);
                --�������˿��ܹ����˶��ҽ������,��Ҫ����ҽ����
                Update ҽ�����˹�����
                Set ҽ���� = r_Keepinsure.ҽ����, ��־ = 0
                Where (����, ����, ҽ����) In (Select ����, ����, ҽ���� From ҽ�����˹����� Where ����id = v_����id);
                Update ҽ�����˹����� Set ��־ = 1 Where ����id = v_����id;
              End If;
              Close c_Keepinsure;
            End If;
          End If;
        End Loop;
      Else
        --b.2.2�����Ĳ������ں���ǰ������ҽ���ʻ�
        Update �ʻ������Ϣ Set ����id = v_����id Where ����id = v_�ϲ�id;
        Update ҽ�����˹����� Set ����id = v_����id Where ����id = v_�ϲ�id;
        --ҽ�����˵������ô���,��Ϊͨ��ҽ���Ź���<ҽ�����˹�����>
      End If;
    Else
      --c.�ϲ��Ĳ�����ǰ�����ڶ�����ҽ���ʻ�,�����κδ���
      Null;
    End If;
  End If;

  --���������ϵͳ�Ĳ��˺ϲ�
  n_Have := 0;
  Begin
    Select 1 Into n_Have From zlSystems Where Floor(��� / 100) = 21;
  Exception
    When Others Then
      Null;
  End;
  If n_Have = 1 Then
    v_Sql := 'Begin zl21_������Ϣ_Merge(:1,:2); End;';
    Execute Immediate v_Sql
      Using v_�ϲ�id, v_����id;
  End If;

  --��������,������ҳ��ر�
  For v_Loop In 1 .. Arronpage.Count Loop
    --Executesql('Update ' || Arronpage(v_Loop) || ' Set ����ID=' || v_����id || ' Where ����ID=' || v_�ϲ�id || ' And Nvl(��ҳID,0) = 0');
    --"��ҳ=0����ҳID is NULL����ҳID=�Һ�ID"���п��ܣ�ǰ�沿������ҳID������û��������˲�������
    v_Sql := 'Update ' || Arronpage(v_Loop) || ' Set ����ID=:1 Where ����ID=:2';
    Execute Immediate v_Sql
      Using v_����id, v_�ϲ�id;
  End Loop;
  For v_Loop In 1 .. Arronbase.Count Loop
    If Arronbase(v_Loop) = '������Ƭ' Then
      Select Count(1) Into n_Have From ������Ƭ Where ����id = v_����id;
      If n_Have = 1 Then
        Delete From ������Ƭ Where ����id = v_�ϲ�id;
      End If;
    End If;
    v_Sql := 'Update ' || Arronbase(v_Loop) || ' Set ����ID=:1 Where ����ID=:2';
    Execute Immediate v_Sql
      Using v_����id, v_�ϲ�id;
  End Loop;

  --ɾ��ʵ�ʲ������Ĳ�����Ϣ
  Delete From ������Ϣ Where ����id = v_�ϲ�id;

  --���ݽ���ѡ����������Ϣ
  Update ������Ϣ
  Set ���� = Nvl(r_Infob.����, r_Infoa.����), �Ա� = Nvl(r_Infob.�Ա�, r_Infoa.�Ա�), ���� = Nvl(r_Infob.����, r_Infoa.����), ����� = v_�����,
      סԺ�� = v_סԺ��, ���￨�� = Nvl(r_Infob.���￨��, r_Infoa.���￨��), ����֤�� = Decode(r_Infob.���￨��, Null, r_Infoa.����֤��, r_Infob.����֤��),
      �ѱ� = Nvl(r_Infob.�ѱ�, r_Infoa.�ѱ�), ҽ�Ƹ��ʽ = Nvl(r_Infob.ҽ�Ƹ��ʽ, r_Infoa.ҽ�Ƹ��ʽ),
      �������� = Nvl(r_Infob.��������, r_Infoa.��������), �����ص� = Nvl(r_Infob.�����ص�, r_Infoa.�����ص�),
      ���֤�� = Nvl(r_Infob.���֤��, r_Infoa.���֤��), ��� = Nvl(r_Infob.���, r_Infoa.���), ְҵ = Nvl(r_Infob.ְҵ, r_Infoa.ְҵ),
      ���� = Nvl(r_Infob.����, r_Infoa.����), ���� = Nvl(r_Infob.����, r_Infoa.����), ѧ�� = Nvl(r_Infob.ѧ��, r_Infoa.ѧ��),
      ���� = Nvl(r_Infob.����, r_Infoa.����), ���� = Nvl(r_Infob.����, r_Infoa.����), ����״�� = Nvl(r_Infob.����״��, r_Infoa.����״��),
      ��ͥ��ַ = Nvl(r_Infob.��ͥ��ַ, r_Infoa.��ͥ��ַ), ��ͥ�绰 = Nvl(r_Infob.��ͥ�绰, r_Infoa.��ͥ�绰),
      ��ͥ��ַ�ʱ� = Nvl(r_Infob.��ͥ��ַ�ʱ�, r_Infoa.��ͥ��ַ�ʱ�), ���ڵ�ַ = Nvl(r_Infob.���ڵ�ַ, r_Infoa.���ڵ�ַ),
      ���ڵ�ַ�ʱ� = Nvl(r_Infob.���ڵ�ַ�ʱ�, r_Infoa.���ڵ�ַ�ʱ�), ��ϵ������ = Nvl(r_Infob.��ϵ������, r_Infoa.��ϵ������),
      ��ϵ�˹�ϵ = Nvl(r_Infob.��ϵ�˹�ϵ, r_Infoa.��ϵ�˹�ϵ), ��ϵ�˵�ַ = Nvl(r_Infob.��ϵ�˵�ַ, r_Infoa.��ϵ�˵�ַ),
      ��ϵ�˵绰 = Nvl(r_Infob.��ϵ�˵绰, r_Infoa.��ϵ�˵绰), ��ͬ��λid = Nvl(r_Infob.��ͬ��λid, r_Infoa.��ͬ��λid),
      ������λ = Nvl(r_Infob.������λ, r_Infoa.������λ), ��λ�绰 = Nvl(r_Infob.��λ�绰, r_Infoa.��λ�绰),
      ��λ�ʱ� = Nvl(r_Infob.��λ�ʱ�, r_Infoa.��λ�ʱ�), ��λ������ = Nvl(r_Infob.��λ������, r_Infoa.��λ������),
      ��λ�ʺ� = Nvl(r_Infob.��λ�ʺ�, r_Infoa.��λ�ʺ�), ����ʱ�� = Nvl(r_Infob.����ʱ��, r_Infoa.����ʱ��),
      ����״̬ = Nvl(r_Infob.����״̬, r_Infoa.����״̬), �������� = Nvl(r_Infob.��������, r_Infoa.��������), ���� = Nvl(r_Infob.����, r_Infoa.����),
      �Ǽ�ʱ�� = Nvl(r_Infob.�Ǽ�ʱ��, r_Infoa.�Ǽ�ʱ��), סԺ���� = Null, ��ҳid = Null, ��ǰ���� = Null, ��ǰ����id = Null, ��ǰ����id = Null,
      ��Ժʱ�� = Null, ��Ժʱ�� = Null, ��Ժ = Decode(Nvl(r_Infob.��Ժ, 0), 1, 1, Null), ������ = Nvl(r_Infob.������, r_Infoa.������)
  Where ����id = v_����id;

  Open c_Info(v_����id);
  Fetch c_Info
    Into r_Info;
  If c_Info%RowCount > 0 Then
    --���һ��ΪԤԼ����,ֻ��Ҫ����סԺ��������Ժʱ��
    If r_Info.��ҳid = 0 Then
      Update ������Ϣ
      Set ��ҳid = Decode(r_Info.�����ҳid, 0, Null, r_Info.�����ҳid), סԺ���� = Decode(n_MaxסԺ����, 0, Null, n_MaxסԺ����)
      Where ����id = v_����id;
    Else
      Update ������Ϣ
      Set ��ҳid = Decode(r_Info.�����ҳid, 0, Null, r_Info.�����ҳid), סԺ���� = Decode(n_MaxסԺ����, 0, Null, n_MaxסԺ����),
          ��ǰ���� = Decode(r_Info.��Ժ����, Null, r_Info.��Ժ����, Null), ��ǰ����id = Decode(r_Info.��Ժ����, Null, r_Info.��ǰ����id, Null),
          ��ǰ����id = Decode(r_Info.��Ժ����, Null, r_Info.��Ժ����id, Null), ��Ժʱ�� = r_Info.��Ժ����, ��Ժʱ�� = r_Info.��Ժ����









      
      Where ����id = v_����id;
    End If;
    --��������Ϣ
    Select Nvl(��ҳid, -1) Into n_��ҳid From ������Ϣ Where ����id = v_����id;
    --��ȡ��ǰ��Ч������������¼,ȷ��������������ʱ����������
    Select Nvl(Sum(������), 0), Count(����id)
    Into n_������, n_Row
    From ���˵�����¼
    Where ����id = v_����id And Nvl(��ҳid, -1) = n_��ҳid And (����ʱ�� Is Null Or ����ʱ�� > Sysdate) And �������� = 0 And ɾ����־ = 1;
    If n_Row = 0 Then
      --�������һ����ʱ������¼,���ൽ��
      Update ���˵�����¼
      Set ����ʱ�� = Sysdate - 1 / 24 / 60 / 60
      Where ����id = v_����id And Nvl(��ҳid, -1) = n_��ҳid And �������� = 1 And (����ʱ�� Is Null Or ����ʱ�� > Sysdate) And ɾ����־ = 1 And
            �Ǽ�ʱ�� <> (Select Max(�Ǽ�ʱ��)
                     From ���˵�����¼
                     Where ����id = v_����id And Nvl(��ҳid, -1) = n_��ҳid And �������� = 1 And (����ʱ�� Is Null Or ����ʱ�� > Sysdate) And
                           ɾ����־ = 1);
    Else
      --����������������ʱ����ʧЧ
      Update ���˵�����¼
      Set ����ʱ�� = Sysdate - 1 / 24 / 60 / 60
      Where ����id = v_����id And Nvl(��ҳid, -1) = n_��ҳid And �������� = 1 And (����ʱ�� Is Null Or ����ʱ�� > Sysdate) And ɾ����־ = 1;
    End If;
  
    --��ȡ��ǰ��Ч�������Ч������¼��
    n_Row    := 0;
    n_������ := 0;
    v_������ := '';
    For r_�ᱣ��Ϣ In (Select ������, ������
                   From ���˵�����¼
                   Where ����id = v_����id And Nvl(��ҳid, -1) = n_��ҳid And (����ʱ�� Is Null Or ����ʱ�� > Sysdate) And ɾ����־ = 1) Loop
      n_Row     := n_Row + 1;
      n_������  := n_������ + r_�ᱣ��Ϣ.������;
      v_������  := v_������ || ',' || r_�ᱣ��Ϣ.������;
      n_Lengthb := Lengthb(v_������);
      If n_Lengthb >= 101 Then
        v_Error := '���ܺϲ�������¼���ڲ�����Ϣ����ʱ�����������ֶγ��ȣ�';
        Raise Err_Custom;
      End If;
    End Loop;
    v_������ := Substr(v_������, 2, 100);
  
    If n_Row = 0 Then
      Update ������Ϣ Set ������ = Null, ������ = Null, �������� = Null Where ����id = v_����id;
    Else
      --��ȡ���һ����Ч�����˺͵�������
      Select ��������
      Into n_��������
      From ���˵�����¼
      Where ����id = v_����id And Nvl(��ҳid, -1) = n_��ҳid And ɾ����־ = 1 And
            �Ǽ�ʱ�� =
            (Select Max(�Ǽ�ʱ��)
             From ���˵�����¼
             Where ����id = v_����id And Nvl(��ҳid, -1) = n_��ҳid And (����ʱ�� Is Null Or ����ʱ�� > Sysdate) And ɾ����־ = 1);
    
      Update ������Ϣ Set ������ = v_������, ������ = n_������, �������� = n_�������� Where ����id = v_����id;
    End If;
  End If;

  Close c_Info;
  Close c_Infoa;
  Close c_Infob;

  --�Բ��˽��н���
  Update ������Ϣ Set ���� = 0 Where ����id In (A����id_In, B����id_In);  
  v_Chgs := Substr(v_Chgs, 2);
  b_Message.Zlhis_Patient_017(v_����id, v_�ϲ�id, v_Chgs);
Exception
  When Err_Custom Then
    Begin
      Rollback; --��Ȼ������
      Zl_������Ϣ_����(A����id_In, 0);
      Zl_������Ϣ_����(B����id_In, 0);
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    End;
  When Others Then
    Begin
      Rollback; --��Ȼ������
      Zl_������Ϣ_����(A����id_In, 0);
      Zl_������Ϣ_����(B����id_In, 0);
      zl_ErrorCenter(SQLCode, SQLErrM);
    End;
End Zl_������Ϣ_Merge;
/



------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0049' Where ���=&n_System;
Commit;