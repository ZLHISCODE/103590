----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--123451:�ƽ�,2018-03-28,RIS�ӿ�ԤԼ���Ӵ�ӡ�˺ʹ�ӡʱ��
alter table RIS���ԤԼ add ��ӡʱ�� date;
alter table RIS���ԤԼ add ��ӡ�� VARCHAR2(100);

--122954:��ΰ��,2018-03-26,����������ҩ
Create Global Temporary Table ����������ҩ����(�������� clob) On Commit Delete Rows;



------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--112136:������,2018-03-26,�������ʱ仯
Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
Select zlParameters_ID.Nextval,&n_System,1254,A.* From (
Select ˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� From zlParameters Where 1 = 0 Union All 
Select 0,0,0,0,0,0,84,'סԺ����ִ���Զ���ɷ���',NULL,NULL,'סԺ ����ִ���Զ����ҽ����� �������ڰ���ͬ�Ŀ������ò�ͬ����ִ���Զ����ҽ�������շ�����','����1,����2;����3,����4������ÿ���ֺ�Ϊһ������','���ұ���ִ���Զ����ҽ�������ղ����ķ������գ�',NULL,NULL From Dual Union All    
Select ˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� From zlParameters Where 1 = 0) A;

--112136:������,2018-03-26,�������ʱ仯
Declare
  v_����ids Varchar2(4000);
Begin
  For P In (Select ID, ����ֵ, ����
            From zlParameters
            Where ������ = '����ִ���Զ����ҽ�����' And ģ�� = 1254 And ϵͳ = &n_System) Loop
    If Nvl(p.����, 0) = 0 Then
      Update zlParameters Set ����ֵ=null,ȱʡֵ=null,���� = 1 Where ID = p.Id;
      If p.����ֵ Is Not Null Then
        For R In (Select Distinct a.Id
                  From ���ű� A, ��������˵�� B
                  Where b.����id = a.Id And
                        (b.�������� = '�ٴ�' And ((b.������� In (2, 3)) Or
                        (b.������� = 1 And Exists (Select 1 From ��λ״����¼ C Where b.����id = c.����id))) Or
                        b.������� In (1, 2, 3) And b.�������� = '����') And
                        (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)) Loop
          v_����ids := v_����ids || ',' || r.Id;
          Insert Into Zldeptparas (����id, ����id, ����ֵ) Values (p.Id, r.Id, p.����ֵ);
        End Loop;
        v_����ids := Substr(v_����ids, 2);
        Update zlParameters
        Set ����ֵ = v_����ids
        Where ������ = 'סԺ����ִ���Զ���ɷ���' And ģ�� = 1254 And ϵͳ = &n_System;
      End If;
    End If;
  End Loop;
End;
/

--123386:��˶,2018-03-23,�շѼ�Ŀ���շѶ���ê��
Update Zlmsg_Lists Set Key_Define='<root><�շ���ĿID></�շ���ĿID></root>' Where Code='ZLHIS_DICT_053';
Update Zlmsg_Lists Set Key_Define='<root><������ĿID></������ĿID></root>' Where Code='ZLHIS_DICT_054';

--122954:��ΰ��,2018-03-26,����������ҩ
Insert into zlTables(ϵͳ,����,��ռ�,����) Values(100,'����������ҩ����','','B2');

Insert Into zlParameters(ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
Select Zlparameters_Id.Nextval, &n_System, Null, a.* From (Select ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵�� From zlParameters Where 1 = 0 Union All
Select 1, 0, 0, 0, 0, 0, 299, 'ҩƷ˵����Ҫ����ʾ', NULL, NULL, '���ڿ����´�ҩƷҽ��ʱ�Ƿ񵯳�ҩƷ��ʾ�Լ�����չʾ��ҩƷ��ʾ��Ŀ','����ֵΪ0|1����;0-����ر�,1��������;��һλ�����Ƿ���Ҫ����ʾ;�ӵڶ�λ��ʼ��Ӧÿһ��Ҫ����ʾ��Ŀ�Ŀ���״̬,����ֵλ������Ҫ����ʾ��Ŀ������1','�����á�������ҩ���ӿڡ�Ϊ������Ϣ��ǰ������Ч', '���ݸ�����Ҫ�ر���ʾ,���������ʾ��Ŀ', NULL From Dual Union All
Select ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵�� From zlParameters Where 1 = 0) A;

-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--122954:��ΰ��,2018-03-26,����������ҩ
--1252 ����ҽ���´�
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1252,'������ҩ���',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_Lob_Append','EXECUTE' From Dual Union All
Select 'Zl_����������ҩ����_Update','EXECUTE' From Dual Union All
Select 'Zl_Read_����������ҩ����','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

--1253:סԺҽ���´�
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1253,'������ҩ���',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_Lob_Append','EXECUTE' From Dual Union All
Select 'Zl_����������ҩ����_Update','EXECUTE' From Dual Union All
Select 'Zl_Read_����������ҩ����','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;
--1254:סԺҽ������
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1254,'������ҩ���',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_Lob_Append','EXECUTE' From Dual Union All
Select 'Zl_����������ҩ����_Update','EXECUTE' From Dual Union All
Select 'Zl_Read_����������ҩ����','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;
--1341:ҩƷ������ҩ
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1341,'������ҩ���',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_Lob_Append','EXECUTE' From Dual Union All
Select 'Zl_����������ҩ����_Update','EXECUTE' From Dual Union All
Select 'Zl_Read_����������ҩ����','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;
--1342:ҩƷ���ŷ�ҩ
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1342,'������ҩ���',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_Lob_Append','EXECUTE' From Dual Union All
Select 'Zl_����������ҩ����_Update','EXECUTE' From Dual Union All
Select 'Zl_Read_����������ҩ����','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;
--1345:��Һ�������Ĺ���
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1345,'������ҩ���',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_Lob_Append','EXECUTE' From Dual Union All
Select 'Zl_����������ҩ����_Update','EXECUTE' From Dual Union All
Select 'Zl_Read_����������ҩ����','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;





-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--123451:�ƽ�,2018-03-28,RIS�ӿ�ԤԼ���Ӵ�ӡ�˺ʹ�ӡʱ��
Create Or Replace Package b_Zlxwinterface Is
  Type t_Refcur Is Ref Cursor;

  --1������RIS״̬�ı�
  Procedure Receiverisstate
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ״̬_In     Number,
    ������Ա_In ����ҽ������.�����%Type,
    ִ��ʱ��_In ����ҽ������.���ʱ��%Type := Null,
    ִ��˵��_In ����ҽ������.ִ��˵��%Type := Null,
    ����ִ��_In Number := 0
  );

  --2������ȷ��
  Procedure Ӱ�����ִ��
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  );

  --3��ȡ������ȷ��
  Procedure Ӱ�����ִ��_Cancel
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  );

  --4������RIS�ı���
  Procedure Receivereport
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ��������_In ���Ӳ�������.�����ı�%Type,
    �������_In ���Ӳ�������.�����ı�%Type,
    ���潨��_In ���Ӳ�������.�����ı�%Type,
    ����ҽ��_In ���Ӳ�����¼.������%Type
  );

  --5���޸����뵥��Ϣ
  Procedure Ӱ������Ϣ_�޸�
  (
    ҽ��id_In       ����ҽ����¼.Id%Type,
    ����_In         ������Ϣ.����%Type,
    �Ա�_In         ������Ϣ.�Ա�%Type,
    ����_In         ������Ϣ.����%Type,
    �ѱ�_In         ������Ϣ.�ѱ�%Type,
    ҽ�Ƹ��ʽ_In ������Ϣ.ҽ�Ƹ��ʽ%Type,
    ����_In         ������Ϣ.����%Type,
    ����_In         ������Ϣ.����״��%Type,
    ְҵ_In         ������Ϣ.ְҵ%Type,
    ���֤��_In     ������Ϣ.���֤��%Type,
    ��ͥ��ַ_In     ������Ϣ.��ͥ��ַ%Type,
    ��ͥ�绰_In     ������Ϣ.��ͥ�绰%Type,
    ��ͥ��ַ�ʱ�_In ������Ϣ.��ͥ��ַ�ʱ�%Type,
    ��������_In     ������Ϣ.��������%Type := Null
  );

  --6��ȡ�����뵥��Ϣ
  Procedure ȡ��������뵥
  (
    ҽ��id_In     ����ҽ��ִ��.ҽ��id%Type,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := 0,
    �ܾ�ԭ��_In   ����ҽ������.ִ��˵��%Type := Null
  );

  --7������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_Insert
  (
    ������Դ_In   In Risҽ��ʧ�ܼ�¼.������Դ%Type,
    ����id_In     In Risҽ��ʧ�ܼ�¼.����id%Type,
    ��ҳid_In     In Risҽ��ʧ�ܼ�¼.��ҳid%Type,
    �Һŵ���_In   In Risҽ��ʧ�ܼ�¼.�Һŵ���%Type,
    ���ͺ�_In     In Risҽ��ʧ�ܼ�¼.���ͺ�%Type,
    �������id_In In Risҽ��ʧ�ܼ�¼.�������id%Type,
    ��챨����_In In Risҽ��ʧ�ܼ�¼.��챨����%Type,
    ��������_In   In Risҽ��ʧ�ܼ�¼.��������%Type
  );

  --8������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_�ط�
  (
    Id_In       In Risҽ��ʧ�ܼ�¼.Id%Type,
    ��������_In In Number
  );

  --9�����˺��½�סԺ���˵���
  Procedure ����ҽ��_�ؽ�����
  (
    ҽ��id_In In ����ҽ������.ҽ��id%Type,
    No_In     In ����ҽ������.No%Type,
    Action_In In Number
  );

  --10����ӡRIS���ԤԼ֪ͨ��
  Procedure Ris���ԤԼ_��ӡ(ҽ��id_In In Ris���ԤԼ.ҽ��id%Type);

  --11������RIS�ֿ������ò���
  Procedure Ris���ÿ���_Update
  (
    �������_In Ris���ÿ���.�������%Type,
    ����_In     Ris���ÿ���.����%Type,
    ����ids_In  Varchar2,
    ��������_In Number
  );

  --12��ɾ��RIS�ֿ������ò���
  Procedure Ris���ÿ���_Delete;

  --13������Ԫ������ȡ��Ϣ
  Function Ris_Replace_Element_Value
  (
    Ԫ����_In   In ����������Ŀ.������%Type,
    ����id_In   In ���Ӳ�����¼.����id%Type,
    ����id_In   In ���Ӳ�����¼.��ҳid%Type,
    ������Դ_In In ���Ӳ�����¼.������Դ%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type
  ) Return Varchar2;

  --14��ɾ��RIS��Ժ���ò���
  Procedure Ris��Ժ����_Delete;

  --15������RISRis��Ժ���ò���
  Procedure Ris��Ժ����_Update
  (
    Id_In           Ris��Ժ����.Id%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    �û���_In       Ris��Ժ����.�û���%Type,
    ����_In         Ris��Ժ����.����%Type,
    ���ݿ������_In Ris��Ժ����.���ݿ������%Type
  );
End b_Zlxwinterface;
/

Create Or Replace Package Body b_Zlxwinterface Is

  --1������RIS״̬�ı�
  Procedure Receiverisstate
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ״̬_In     Number,
    ������Ա_In ����ҽ������.�����%Type,
    ִ��ʱ��_In ����ҽ������.���ʱ��%Type := Null,
    ִ��˵��_In ����ҽ������.ִ��˵��%Type := Null,
    ����ִ��_In Number := 0
  ) Is
  
    --������ҽ��ID_IN - ����ִ�е�ҽ��ID��
    --      ״̬_IN - -1-ɾ����0-ԤԼ��1-�Ǽǣ�3-�����ɣ�4-�����ֹ��9-�������棻12-������ˣ�15-����
    --     ����ִ��_In -0-ȫ��ִ�У�1-����ִ�У����ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ
  
    Cursor c_Adviceinfo Is
      Select a.Id, a.���id, Nvl(a.���id, a.Id) As ��id, a.�������, a.������Դ, a.ִ�п���id, b.ִ�й���
      From ����ҽ����¼ A, ����ҽ������ B
      Where a.Id = b.ҽ��id And ID = ҽ��id_In;
    r_Adviceinfo c_Adviceinfo%RowType;
  
    v_ִ��״̬ ����ҽ������.ִ��״̬%Type;
    v_ִ�й��� ����ҽ������.ִ�й���%Type;
    n_ִ��     Number; --����Ƿ���Ҫ����״̬��1����Ҫ���£���������Ҫ����
    v_Count    Number;
    v_�����   ����ҽ������.�����%Type;
    v_���ʱ�� ����ҽ������.���ʱ��%Type;
    v_Error    Varchar2(255);
    Err_Custom Exception;
  
  Begin
  
    v_ִ��״̬ := 0;
    v_ִ�й��� := 0;
  
    --��ȡҽ������ҽ��ID������ID
    Open c_Adviceinfo;
    Fetch c_Adviceinfo
      Into r_Adviceinfo;
    Close c_Adviceinfo;
  
    --����״̬_INִ��ҽ��
    ---1-ɾ����0-ԤԼ��1-�Ǽǣ�3-�����ɣ�4-�����ֹ��9-�������棻12-������ˣ�13-ȡ����ˣ�14-����ɾ����15-����
  
    If ״̬_In = -1 Or ״̬_In = 0 Then
      v_ִ��״̬ := 0; --δִ��
      v_ִ�й��� := 0;
    Elsif ״̬_In = 1 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 2; --�ѱ���
    Elsif ״̬_In = 3 Or ״̬_In = 14 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 3; --�Ѽ��
    Elsif ״̬_In = 4 Then
      --���ı�
      v_ִ��״̬ := v_ִ��״̬;
    Elsif ״̬_In = 9 Or ״̬_In = 13 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 4; --�ѱ���
    Elsif ״̬_In = 12 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 5; --�����
    Elsif ״̬_In = 15 Then
      v_ִ��״̬ := 1; --��ȫִ��
      v_ִ�й��� := 6; --�����
      v_�����   := ������Ա_In;
      v_���ʱ�� := ִ��ʱ��_In;
    End If;
  
    n_ִ�� := 1; --Ĭ�϶�Ҫ����״̬
  
    If ״̬_In = 13 Or ״̬_In = 14 Then
      --ɾ����Ӧ��������
      Delete From ���Ӳ�����¼
      Where ID = (Select ����id From ����ҽ������ Where ҽ��id = ҽ��id_In And Risid = Risid_In);
      Delete From ����ҽ������ Where ҽ��id = ҽ��id_In And Risid = Risid_In;
    
      --ɾ�����ж��Ƿ񻹴��ڱ��棬��������ҽ��״̬���ֲ��䣬������ȫ��ɾ�������ҽ��״̬
      Select Count(1) Into v_Count From ����ҽ������ Where ҽ��id = ҽ��id_In;
    
      If v_Count > 0 Then
        n_ִ�� := 0; --��������ҽ��״̬���ֲ���
      End If;
    End If;
  
    --����ǵǼǣ����жϴ˼���Ƿ�δִ��
    If ״̬_In = 1 Then
      If r_Adviceinfo.ִ�й��� >= 3 Then
        v_Error := '�����Ѿ���������ˣ������ظ��Ǽǡ�';
        Raise Err_Custom;
      End If;
    End If;
  
    --��ʼִ��ҽ��
    If n_ִ�� = 1 Then
      If Nvl(����ִ��_In, 0) = 1 Then
        -- ������λҽ������ִ��
        Update ����ҽ������
        Set ִ��״̬ = v_ִ��״̬, ִ�й��� = v_ִ�й���, ִ��˵�� = ִ��˵��_In, ����� = v_�����, ���ʱ�� = v_���ʱ��
        Where ҽ��id = ҽ��id_In;
      Else
        Update ����ҽ������
        Set ִ��״̬ = v_ִ��״̬, ִ�й��� = v_ִ�й���, ִ��˵�� = ִ��˵��_In, ����� = v_�����, ���ʱ�� = v_���ʱ��
        Where ҽ��id In (Select ID From ����ҽ����¼ Where (ID = r_Adviceinfo.��id Or ���id = r_Adviceinfo.��id));
      End If;
    End If;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Receiverisstate;

  --2������ȷ��
  Procedure Ӱ�����ִ��
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  ) Is
    --������ҽ��ID_IN=����ִ�е�ҽ��ID��
    --      ����ִ��_In=���ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ,0-������ִ��
    Cursor c_Advice Is
      Select ID, ���id, Nvl(���id, ID) As ��id, �������, ������Դ From ����ҽ����¼ Where ID = ҽ��id_In;
    r_Advice c_Advice%RowType;
  
    v_Temp     Varchar2(255);
    v_��Ա��� ��Ա��.���%Type;
    v_��Ա���� ��Ա��.����%Type;
    v_����id   ���ű�.Id%Type;
    v_�������� ����ҽ������.��¼����%Type;
    v_���ͺ�   ����ҽ������.���ͺ�%Type;
    v_ִ�й��� ����ҽ������.ִ�й���%Type;
    v_Count    Number;
    v_Error    Varchar2(255);
    Err_Custom Exception;
  Begin
  
    --ȡ��ҽ��ID
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
    Close c_Advice;
  
    Select ���ͺ�, ִ�й��� Into v_���ͺ�, v_ִ�й��� From ����ҽ������ Where ҽ��id = r_Advice.��id;
  
    --�ǼǺ���ɲ�ִ�з���  2-�Ǽǣ�3-��飬4-���棬5-��ˣ�6-���
    If v_ִ�й��� >= 2 Or v_ִ�й��� <= 6 Then
    
      --�ȼ���Ƿ��Ѿ���Ժ��סԺ���ˣ��Ѿ�Ԥ��Ժ���߳�Ժ�ļ�����룬����ִ�з���
      Select Count(*)
      Into v_Count
      From ����ҽ����¼ A, ������ҳ B
      Where a.����id = b.����id And a.��ҳid = b.��ҳid And (b.��Ժ���� Is Not Null Or b.״̬ = 3) And a.Id = r_Advice.��id;
    
      If v_Count > 0 Then
        --�Ѿ���Ժ��Ԥ��Ժ��תԺ����Ҫ�ж����Ƿ�����
        Select Count(*)
        Into v_Count
        From ����ҽ����¼ A, ������ĿĿ¼ B, ����ҽ������ C
        Where a.Id = c.ҽ��id And a.������Ŀid = b.Id And b.��� = 'Z' And b.�������� = 11 And
              a.����id = (Select d.����id From ����ҽ����¼ D Where d.Id = r_Advice.��id);
        If v_Count > 0 Then
          v_Error := '�Ѿ��Ի����´�����ҽ��������ִ�з��á�';
          Raise Err_Custom;
        End If;
        --���ж��Ƿ��Ѿ�ԤԼ���Ѿ�ԤԼ��ִ��
        Select Count(*) Into v_Count From Ris���ԤԼ Where ҽ��id = r_Advice.��id;
        If v_Count = 0 Then
          --�Ѿ���Ժ����Ԥ��Ժ��δԤԼ������ھɰ�PACS�Ѿ�������Ҳ����ִ��
          Select Count(*) Into v_Count From Ӱ�����¼ Where ҽ��id = r_Advice.��id;
          If v_Count = 0 Then
            v_Error := 'סԺ�����Ѿ���Ժ����Ԥ��Ժ������ִ�з��á�';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      --ȡ��ǰ������Ա
      If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null And ִ�в���id_In Is Not Null Then
        v_��Ա��� := ����Ա���_In;
        v_��Ա���� := ����Ա����_In;
        v_����id   := ִ�в���id_In;
      Else
        v_Temp     := Zl_Identity;
        v_����id   := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      End If;
    
      If r_Advice.������Դ = 2 Then
        Select Decode(��¼����, 1, 1, Decode(�������, 1, 1, 2))
        Into v_��������
        From ����ҽ������
        Where ���ͺ� = v_���ͺ� And ҽ��id = ҽ��id_In;
      Else
        v_�������� := 1;
      End If;
    
      --ִ�з��ú��Զ�����
      If v_�������� = 1 Then
        Zl_����ҽ��ִ��_Finish(ҽ��id_In, v_���ͺ�, ����ִ��_In, v_��Ա���, v_��Ա����, r_Advice.��id, r_Advice.�������, v_����id);
      Else
        Zl_סԺҽ��ִ��_Finish(ҽ��id_In, v_���ͺ�, ����ִ��_In, v_��Ա���, v_��Ա����, r_Advice.��id, r_Advice.�������, v_����id);
      End If;
    End If;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ӱ�����ִ��;

  --3��ȡ������ȷ��
  Procedure Ӱ�����ִ��_Cancel
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  ) Is
    --������
    --      ҽ��ID_IN=����ִ�е�ҽ��ID��
    --      ����ִ��_In=���ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ,0-������ִ��
  
    Cursor c_Advice Is
      Select ID, ���id, Nvl(���id, ID) As ��id From ����ҽ����¼ Where ID = ҽ��id_In;
    r_Advice c_Advice%RowType;
  
    v_���ͺ� ����ҽ������.���ͺ�%Type;
    v_Count  Number;
    v_Error  Varchar2(255);
    Err_Custom Exception;
  
  Begin
  
    --ȡ��ҽ��ID
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
    Close c_Advice;
  
    --�ȼ���Ƿ��Ѿ���Ժ��סԺ���ˣ��Ѿ�Ԥ��Ժ���߳�Ժ�ļ�����룬����ִ�з���
    Select Count(*)
    Into v_Count
    From ����ҽ����¼ A, ������ҳ B
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And (b.��Ժ���� Is Not Null Or b.״̬ = 3) And a.Id = r_Advice.��id;
  
    If v_Count > 0 Then
      v_Error := 'סԺ�����Ѿ���Ժ����Ԥ��Ժ������ȡ�����á�';
      Raise Err_Custom;
    End If;
  
    Select ���ͺ� Into v_���ͺ� From ����ҽ������ Where ҽ��id = r_Advice.��id;
  
    --����ͳһ��ҽ��ִ��Cancel����
    Zl_����ҽ��ִ��_Cancel(ҽ��id_In, v_���ͺ�, Null, ����ִ��_In, ִ�в���id_In, ����Ա���_In, ����Ա����_In);
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ӱ�����ִ��_Cancel;

  --4������RIS�ı���
  Procedure Receivereport
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ��������_In ���Ӳ�������.�����ı�%Type,
    �������_In ���Ӳ�������.�����ı�%Type,
    ���潨��_In ���Ӳ�������.�����ı�%Type,
    ����ҽ��_In ���Ӳ�����¼.������%Type
  ) Is
    --��ȡ����ҽ��������������Ϣ
    Cursor c_Advice
    (
      v_��id  Number,
      v_Risid Number
    ) Is
      Select e.Id, e.������Դ, e.����id, e.��ҳid, e.Ӥ��, e.���˿���id, e.�ļ�id, e.��������, e.��������, f.����id, e.ִ�п���id
      From (Select c.Id, c.������Դ, c.����id, c.��ҳid, c.Ӥ��, c.���˿���id, c.�ļ�id, d.���� ��������, d.���� ��������, c.ִ�п���id
             From (Select a.Id, a.������Դ, a.����id, a.��ҳid, a.Ӥ��, a.���˿���id, b.�����ļ�id �ļ�id, a.ִ�п���id
                    From ����ҽ����¼ A, ��������Ӧ�� B
                    Where a.Id = v_��id And a.������Ŀid = b.������Ŀid(+) And b.Ӧ�ó���(+) = Decode(a.������Դ, 2, 2, 4, 4, 1)) C,
                  �����ļ��б� D
             Where c.�ļ�id = d.Id(+)) E, ����ҽ������ F
      Where e.Id = f.ҽ��id(+) And f.Risid(+) = v_Risid;
  
    --�����ļ������Ԫ��
    Cursor c_File(v_File Number) Is
      Select a.Id, a.�ļ�id, a.��id, a.�������, a.��������, a.������, a.��������, a.��������, a.�����д�, a.�����ı�, a.�Ƿ���, a.Ԥ�����id, a.�������,
             a.ʹ��ʱ��, a.����Ҫ��id, a.�滻��, a.Ҫ������, a.Ҫ������, a.Ҫ�س���, a.Ҫ��С��, a.Ҫ�ص�λ, a.Ҫ�ر�ʾ, a.������̬, a.Ҫ��ֵ��
      From �����ļ��ṹ A
      Where a.�ļ�id = v_File
      Order By a.�������;
  
    Cursor c_Report(v_���Ӳ�����¼id Number) Is
      Select b.Id, a.�����ı�
      From ���Ӳ������� A, ���Ӳ������� B
      Where a.�������� = 3 And a.Id = b.��id And b.�������� = 2 And b.��ֹ�� = 0 And a.�ļ�id = v_���Ӳ�����¼id;
  
    Cursor c_Content
    (
      v_�ļ�id Number,
      v_���id Number
    ) Is
      Select a.Id, a.�ļ�id, a.��id, a.�������, a.��������, a.������, a.��������, a.��������, a.�����д�, a.�����ı�, a.�Ƿ���, a.Ԥ�����id, a.�������,
             a.ʹ��ʱ��, a.����Ҫ��id, a.�滻��, a.Ҫ������, a.Ҫ������, a.Ҫ�س���, a.Ҫ��С��, a.Ҫ�ص�λ, a.Ҫ�ر�ʾ, a.������̬, a.Ҫ��ֵ��
      From �����ļ��ṹ A
      Where �ļ�id = v_�ļ�id And ��id = v_���id;
  
    r_Advice        c_Advice%RowType;
    v_����id        ���Ӳ�������.�ļ�id%Type;
    v_��������id    ���Ӳ�������.Id%Type;
    v_��������idnew ���Ӳ�������.Id%Type;
    v_�������      ���Ӳ�������.�������%Type;
    v_��id          ���Ӳ�������.��id%Type;
    v_�����ı�      ���Ӳ�������.�����ı�%Type;
    v_�������id    ���Ӳ�������.�������id%Type;
    --v_��ʽ����    ���Ӳ�����ʽ.����%Type;
    v_Error Varchar2(255);
    Err_Custom Exception;
    v_��ҽ��id ����ҽ������.ҽ��id%Type;
    v_���     Varchar2(300);
    n_����     Number;
    n_Rptcount Number;
    v_�������� ���Ӳ�����¼.��������%Type;
    v_�Һŵ�id ���˹Һż�¼.Id%Type;
  
    Function Getrptno
    (
      v_ҽ��idin   ����ҽ������.ҽ��id%Type,
      v_��������in ���Ӳ�����¼.��������%Type
    ) Return Varchar As
      v_Return Number;
      v_No     Number;
      v_Count  Number;
    Begin
      Select Count(ҽ��id) + 1 Into v_No From ����ҽ������ Where ҽ��id = v_ҽ��idin;
      v_Count := 1;
      While v_Count = 1 Loop
        Select Count(ID)
        Into v_Count
        From ����ҽ������ A, ���Ӳ�����¼ B
        Where a.ҽ��id = v_ҽ��idin And a.����id = b.Id And b.�������� = v_��������in || v_No;
        If v_Count = 1 Then
          v_No := v_No + 1;
        End If;
      End Loop;
      v_Return := v_No;
      Return v_Return;
    End Getrptno;
  
  Begin
  
    -- ��ȡ��ҽ��ID ����ֹ��Ϊ���벿λҽ�������±��汣�����
    Select Nvl(���id, ID) As ��id Into v_��ҽ��id From ����ҽ����¼ Where ID = ҽ��id_In;
  
    Open c_Advice(v_��ҽ��id, Nvl(Risid_In, 0));
    Fetch c_Advice
      Into r_Advice;
  
    If Nvl(r_Advice.�ļ�id, 0) = 0 Then
      v_Error := '���μ����Ŀû�ж�Ӧ��صļ�鱨�棬�������Ա��ϵ��';
      Raise Err_Custom;
    Else
      If Nvl(r_Advice.����id, 0) > 0 Then
        ----����������
        --�ҳ��������д�ı�������к���"%����%","%����%","%����%","%���%",���ô���Ĳ�������
        For r_Report In c_Report(r_Advice.����id) Loop
          If r_Report.�����ı� Like '%����%' Then
            Update ���Ӳ������� Set �����ı� = ��������_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          Elsif r_Report.�����ı� Like '%���%' Then
            Update ���Ӳ������� Set �����ı� = �������_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          Elsif r_Report.�����ı� Like '%����%' Then
            Update ���Ӳ������� Set �����ı� = ���潨��_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          End If;
        End Loop;
        --���±���ʱ��
        Update ���Ӳ�����¼
        Set ���ʱ�� = Sysdate, ������ = ����ҽ��_In, ����ʱ�� = Sysdate
        Where ID = r_Advice.����id;
      Else
        --���жϵ������Ƿ��ж�Ӧ����ٺͱ��
        If Nvl(��������_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_����
          From �����ļ��ṹ A, �����ļ��ṹ B
          Where a.��id = b.Id And a.�������� = 3 And b.�������� = 1 And a.�����ı� Like '%����%' And a.�ļ�id = r_Advice.�ļ�id;
        
          If n_���� <= 0 Then
            v_Error := '�����Ƶ�����û���ҵ�����������Ӧ����ٻ�������ϵ����Ա���ã�';
            Raise Err_Custom;
          End If;
        End If;
        If Nvl(�������_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_����
          From �����ļ��ṹ A, �����ļ��ṹ B
          Where a.��id = b.Id And a.�������� = 3 And b.�������� = 1 And a.�����ı� Like '%���%' And a.�ļ�id = r_Advice.�ļ�id;
        
          If n_���� <= 0 Then
            v_Error := '�����Ƶ�����û���ҵ����������Ӧ����ٻ�������ϵ����Ա���ã�';
            Raise Err_Custom;
          End If;
        End If;
        If Nvl(���潨��_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_����
          From �����ļ��ṹ A, �����ļ��ṹ B
          Where a.��id = b.Id And a.�������� = 3 And b.�������� = 1 And a.�����ı� Like '%����%' And a.�ļ�id = r_Advice.�ļ�id;
        
          If n_���� <= 0 Then
            v_Error := '�����Ƶ�����û���ҵ������顿��Ӧ����ٻ�������ϵ����Ա���ã�';
            Raise Err_Custom;
          End If;
        End If;
      
        If r_Advice.������Դ = 1 Then
          --�����ȡ�Һŵ�ID
          Select Nvl(c.Id, 0)
          Into v_�Һŵ�id
          From ����ҽ����¼ B, ���˹Һż�¼ C
          Where b.�Һŵ� = c.No(+) And c.��¼״̬ In (1, 3) And b.Id = v_��ҽ��id;
        Else
          --����������޹Һŵ�ID��ֱ������Ϊ0
          v_�Һŵ�id := 0;
        End If;
      
        --�������Ӳ�����¼
        Select ���Ӳ�����¼_Id.Nextval Into v_����id From Dual;
        n_Rptcount := Getrptno(ҽ��id_In, r_Advice.��������);
        If n_Rptcount > 1 Then
          v_�������� := r_Advice.�������� || n_Rptcount;
        Else
          v_�������� := r_Advice.��������;
        End If;
        Insert Into ���Ӳ�����¼
          (ID, ������Դ, ����id, ��ҳid, Ӥ��, ����id, ��������, �ļ�id, ��������, ������, ����ʱ��, ���ʱ��, ������, ����ʱ��, ���汾, ǩ������)
        Values
          (v_����id, r_Advice.������Դ, r_Advice.����id, Decode(r_Advice.������Դ, 2, r_Advice.��ҳid, v_�Һŵ�id), r_Advice.Ӥ��,
           r_Advice.���˿���id, r_Advice.��������, r_Advice.�ļ�id, v_��������, ����ҽ��_In, Sysdate, Sysdate, ����ҽ��_In, Sysdate, 1, 2);
      
        --����ҽ�������¼
        Insert Into ����ҽ������ (ҽ��id, ����id, Risid) Values (v_��ҽ��id, v_����id, Risid_In);
      
        v_������� := 0;
      
        --�²�����������
        For r_File In c_File(r_Advice.�ļ�id) Loop
          Select ���Ӳ�������_Id.Nextval Into v_��������id From Dual;
          v_�����ı�   := r_File.�����ı�;
          v_�������id := 0;
        
          If Nvl(r_File.��������, 0) = 1 And Nvl(r_File.��id, 0) = 0 Then
            --���
            v_�������id := r_File.Id;
            v_��id       := v_��������id;
          End If;
        
          If Nvl(r_File.��������, 0) = 4 And r_File.Ҫ������ Is Not Null Then
            --Ԫ��
            v_�����ı� := Zl_Replace_Element_Value(r_File.Ҫ������, r_Advice.����id, r_Advice.��ҳid, r_Advice.������Դ, r_Advice.Id);
          End If;
        
          If Nvl(r_File.��id, 0) <> 0 Then
            v_�������id := 0;
          End If;
        
          v_������� := v_������� + 1;
        
          If Instr(v_���, '|' || r_File.��id || '|') > 0 Then
            Null;
          Else
            Insert Into ���Ӳ�������
              (ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id, �滻��,
               Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��, �������id)
            Values
              (v_��������id, v_����id, 1, 0, Decode(v_�������id, 0, v_��id, Null), v_�������, r_File.��������, r_File.������, r_File.��������,
               r_File.��������, Null, v_�����ı�, r_File.�Ƿ���, r_File.Ԥ�����id, r_File.�������, r_File.ʹ��ʱ��, r_File.����Ҫ��id,
               r_File.�滻��, r_File.Ҫ������, r_File.Ҫ������, r_File.Ҫ�س���, r_File.Ҫ��С��, r_File.Ҫ�ص�λ, r_File.Ҫ�ر�ʾ, r_File.������̬,
               r_File.Ҫ��ֵ��, Decode(v_�������id, 0, Null, v_�������id));
          End If;
        
          --Ϊ���ʱ�������ı�����
          If Nvl(r_File.��������, 0) = 3 And Nvl(r_File.��id, 0) <> 0 Then
            v_��� := v_��� || ',|' || r_File.Id || '|';
          
            If r_File.�����ı� Like '%����%' Then
              v_�����ı� := ��������_In || Chr(13) || Chr(13);
            Elsif r_File.�����ı� Like '%���%' Then
              v_�����ı� := �������_In || Chr(13) || Chr(13);
            Else
              v_�����ı� := ���潨��_In || Chr(13) || Chr(13);
            End If;
          
            For r_Con In c_Content(r_Advice.�ļ�id, r_File.Id) Loop
              Select ���Ӳ�������_Id.Nextval Into v_��������idnew From Dual;
              v_������� := v_������� + 1;
            
              Insert Into ���Ӳ�������
                (ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id,
                 �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��, �������id)
              Values
                (v_��������idnew, v_����id, 1, 0, v_��������id, v_�������, 2, r_Con.������, r_Con.��������, r_Con.��������, Null, v_�����ı�,
                 r_Con.�Ƿ���, r_Con.Ԥ�����id, r_Con.�������, r_Con.ʹ��ʱ��, r_Con.����Ҫ��id, r_Con.�滻��, r_Con.Ҫ������, r_Con.Ҫ������,
                 r_Con.Ҫ�س���, r_Con.Ҫ��С��, r_Con.Ҫ�ص�λ, r_Con.Ҫ�ر�ʾ, r_Con.������̬, r_Con.Ҫ��ֵ��,
                 Decode(v_�������id, 0, Null, v_�������id));
            End Loop;
          End If;
        End Loop;
      
        --����Ӳ�����ʽ�к����������ָ�ʽ�����ַ�������֮���������ֽ����ɼ�
        --Select ���� Into v_��ʽ���� From �����ļ���ʽ Where �ļ�ID=r_Advice.�ļ�ID;
        --Insert Into ���Ӳ�����ʽ (�ļ�ID,����) Values (v_����id,v_��ʽ����);
      
      End If;
    End If;
    Close c_Advice;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Receivereport;

  --5���޸����뵥��Ϣ
  Procedure Ӱ������Ϣ_�޸�
  (
    ҽ��id_In       ����ҽ����¼.Id%Type,
    ����_In         ������Ϣ.����%Type,
    �Ա�_In         ������Ϣ.�Ա�%Type,
    ����_In         ������Ϣ.����%Type,
    �ѱ�_In         ������Ϣ.�ѱ�%Type,
    ҽ�Ƹ��ʽ_In ������Ϣ.ҽ�Ƹ��ʽ%Type,
    ����_In         ������Ϣ.����%Type,
    ����_In         ������Ϣ.����״��%Type,
    ְҵ_In         ������Ϣ.ְҵ%Type,
    ���֤��_In     ������Ϣ.���֤��%Type,
    ��ͥ��ַ_In     ������Ϣ.��ͥ��ַ%Type,
    ��ͥ�绰_In     ������Ϣ.��ͥ�绰%Type,
    ��ͥ��ַ�ʱ�_In ������Ϣ.��ͥ��ַ�ʱ�%Type,
    ��������_In     ������Ϣ.��������%Type := Null
  ) As
  
    v_����     Varchar2(20);
    v_���䵥λ Varchar2(20);
    v_�������� Date;
    v_������Դ ����ҽ����¼.������Դ%Type;
    v_����id   ����ҽ����¼.����id%Type;
  Begin
    Begin
      Select ������Դ, ����id Into v_������Դ, v_����id From ����ҽ����¼ Where ID = ҽ��id_In;
    Exception
      When Others Then
        Return;
    End;
  
    If ��������_In Is Null And ����_In Is Not Null Then
      --�����������������
      v_���䵥λ := Substr(����_In, Length(����_In), 1);
      If Instr('��,��,��', v_���䵥λ) <= 0 Then
        v_���䵥λ := Null;
      Else
        v_���� := Replace(����_In, v_���䵥λ, '');
      End If;
      Begin
        v_���� := To_Number(v_����);
      Exception
        When Others Then
          v_���� := Null;
      End;
      If v_���� Is Not Null And v_���䵥λ Is Not Null Then
        Select Decode(v_���䵥λ, '��', Add_Months(Sysdate, -12 * v_����), '��', Add_Months(Sysdate, -1 * v_����), '��',
                       Sysdate - v_����)
        Into v_��������
        From Dual;
      End If;
    Else
      v_�������� := ��������_In;
    End If;
  
    If v_������Դ = 3 Then
      Update ������Ϣ
      Set ���� = ����_In, �Ա� = Nvl(�Ա�_In, �Ա�), ���� = ����_In, �������� = v_��������, �ѱ� = Nvl(�ѱ�_In, �ѱ�),
          ҽ�Ƹ��ʽ = Nvl(ҽ�Ƹ��ʽ_In, ҽ�Ƹ��ʽ), ���� = Nvl(����_In, ����), ����״�� = Nvl(����_In, ����״��), ְҵ = Nvl(ְҵ_In, ְҵ),
          ���֤�� = ���֤��_In, ��ͥ��ַ = ��ͥ��ַ_In, ��ͥ�绰 = ��ͥ�绰_In, ��ͥ��ַ�ʱ� = ��ͥ��ַ�ʱ�_In
      Where ����id = v_����id;
    
      --�޸Ķ�Ӧ��ҽ����¼
      Update ����ҽ����¼
      Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In
      Where ID = ҽ��id_In Or ���id = ҽ��id_In;
    Else
      Update ������Ϣ
      Set ���� = Nvl(����_In, ����), ����״�� = Nvl(����_In, ����״��), ְҵ = Nvl(ְҵ_In, ְҵ), ��ͥ��ַ = ��ͥ��ַ_In, ��ͥ�绰 = ��ͥ�绰_In,
          ��ͥ��ַ�ʱ� = ��ͥ��ַ�ʱ�_In
      Where ����id = v_����id;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ӱ������Ϣ_�޸�;

  --6��ȡ�����뵥��Ϣ
  Procedure ȡ��������뵥
  (
    ҽ��id_In     ����ҽ��ִ��.ҽ��id%Type,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := 0,
    �ܾ�ԭ��_In   ����ҽ������.ִ��˵��%Type := Null
  ) As
    --������ҽ��ID_IN=����ִ�е�ҽ��ID
  
    v_���ͺ� ����ҽ��ִ��.���ͺ�%Type;
  
  Begin
  
    Begin
      Select ���ͺ� Into v_���ͺ� From ����ҽ������ Where ҽ��id = ҽ��id_In;
    Exception
      When Others Then
        Return;
    End;
  
    Zl_����ҽ��ִ��_�ܾ�ִ��(ҽ��id_In, v_���ͺ�, ����Ա���_In, ����Ա����_In, ִ�в���id_In, �ܾ�ԭ��_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End ȡ��������뵥;

  --7������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_Insert
  (
    ������Դ_In   In Risҽ��ʧ�ܼ�¼.������Դ%Type,
    ����id_In     In Risҽ��ʧ�ܼ�¼.����id%Type,
    ��ҳid_In     In Risҽ��ʧ�ܼ�¼.��ҳid%Type,
    �Һŵ���_In   In Risҽ��ʧ�ܼ�¼.�Һŵ���%Type,
    ���ͺ�_In     In Risҽ��ʧ�ܼ�¼.���ͺ�%Type,
    �������id_In In Risҽ��ʧ�ܼ�¼.�������id%Type,
    ��챨����_In In Risҽ��ʧ�ܼ�¼.��챨����%Type,
    ��������_In   In Risҽ��ʧ�ܼ�¼.��������%Type
  ) Is
  Begin
    Insert Into Risҽ��ʧ�ܼ�¼
      (ID, ������Դ, ����id, ��ҳid, �Һŵ���, ���ͺ�, �������id, ��챨����, ��������, ����ʱ��, �ط�����)
    Values
      (Risҽ��ʧ�ܼ�¼_Id.Nextval, ������Դ_In, ����id_In, ��ҳid_In, �Һŵ���_In, ���ͺ�_In, �������id_In, ��챨����_In, ��������_In, Sysdate, 0);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Risҽ��ʧ�ܼ�¼_Insert;

  --8������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_�ط�
  (
    Id_In       In Risҽ��ʧ�ܼ�¼.Id%Type,
    ��������_In In Number
  ) Is
    v_�ط����� Risҽ��ʧ�ܼ�¼.�ط�����%Type;
  Begin
    --��������_In -- 1 �ط��ɹ���ɾ����¼��2--�ط�ʧ��
  
    If ��������_In = 1 Then
      Delete From Risҽ��ʧ�ܼ�¼ Where ID = Id_In;
    Else
      Select �ط����� Into v_�ط����� From Risҽ��ʧ�ܼ�¼ Where ID = Id_In;
      If v_�ط����� >= 99 Then
        v_�ط����� := 99;
      Else
        v_�ط����� := v_�ط����� + 1;
      End If;
      Update Risҽ��ʧ�ܼ�¼ Set ����ʱ�� = Sysdate, �ط����� = v_�ط����� Where ID = Id_In;
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Risҽ��ʧ�ܼ�¼_�ط�;

  --9�����˺��½�סԺ���˵���
  Procedure ����ҽ��_�ؽ�����
  (
    ҽ��id_In In ����ҽ������.ҽ��id%Type,
    No_In     In ����ҽ������.No%Type,
    Action_In In Number
  ) Is
    -- Action_In: 1 �ؽ����ݣ�2 ȡ���ؽ�����
    v_No ����ҽ������.No%Type;
  Begin
    If Action_In = 1 Then
      Select Nextno(14) Into v_No From Dual;
    
      Update ����ҽ������
      Set NO = v_No, �Ʒ�״̬ = 0
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In);
      Update סԺ���ü�¼ Set ҽ����� = Null Where NO = No_In;
    Elsif Action_In = 2 Then
      Update סԺ���ü�¼ Set ҽ����� = ҽ��id_In Where NO = No_In;
      Update ����ҽ������
      Set NO = No_In, �Ʒ�״̬ = 4
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In);
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End ����ҽ��_�ؽ�����;

  --10����ӡRIS���ԤԼ֪ͨ��
  Procedure Ris���ԤԼ_��ӡ(ҽ��id_In In Ris���ԤԼ.ҽ��id%Type) Is
    v_Temp     Varchar2(255);
    v_��Ա���� ��Ա��.����%Type;
  Begin
    --ȡ��ǰ������Ա  
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  
    Update Ris���ԤԼ Set �Ƿ��ӡ = 1, ��ӡ�� = v_��Ա����, ��ӡʱ�� = Sysdate Where ҽ��id = ҽ��id_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris���ԤԼ_��ӡ;

  --11������RIS�ֿ������ò���
  Procedure Ris���ÿ���_Update
  (
    �������_In Ris���ÿ���.�������%Type,
    ����_In     Ris���ÿ���.����%Type,
    ����ids_In  Varchar2,
    ��������_In Number
  ) Is
  
    l_����id   t_Numlist := t_Numlist();
    v_����ris  Ris���ÿ���.�Ƿ�����ris%Type;
    v_����ԤԼ Ris���ÿ���.�Ƿ�����ԤԼ%Type;
  
    Cursor c_Dept(Dept_In Varchar2) Is
      Select Column_Value From Table(f_Num2list(Dept_In));
  Begin
  
    If ��������_In = 1 Then
      v_����ris  := 1;
      v_����ԤԼ := Null;
      Delete From Ris���ÿ��� Where ������� = �������_In And ���� = ����_In And �Ƿ�����ris = 1;
    Else
      v_����ris  := Null;
      v_����ԤԼ := 1;
      Delete From Ris���ÿ��� Where ������� = �������_In And ���� = ����_In And �Ƿ�����ԤԼ = 1;
    End If;
  
    If ����ids_In Is Null Then
      Insert Into Ris���ÿ���
        (ID, �������, ����, ����id, �Ƿ�����ris, �Ƿ�����ԤԼ)
      Values
        (Ris���ÿ���_Id.Nextval, �������_In, ����_In, Null, v_����ris, v_����ԤԼ);
    Else
      Open c_Dept(����ids_In);
      Fetch c_Dept Bulk Collect
        Into l_����id;
      Close c_Dept;
    
      Forall I In 1 .. l_����id.Count
        Insert Into Ris���ÿ���
          (ID, �������, ����, ����id, �Ƿ�����ris, �Ƿ�����ԤԼ)
        Values
          (Ris���ÿ���_Id.Nextval, �������_In, ����_In, l_����id(I), v_����ris, v_����ԤԼ);
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris���ÿ���_Update;

  --12��ɾ��RIS�ֿ������ò���
  Procedure Ris���ÿ���_Delete Is
  
  Begin
    Delete From Ris���ÿ���;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris���ÿ���_Delete;

  --13������Ԫ������ȡ��Ϣ
  Function Ris_Replace_Element_Value
  (
    Ԫ����_In   In ����������Ŀ.������%Type,
    ����id_In   In ���Ӳ�����¼.����id%Type,
    ����id_In   In ���Ӳ�����¼.��ҳid%Type,
    ������Դ_In In ���Ӳ�����¼.������Դ%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type
  ) Return Varchar2 Is
    v_Return Varchar2(4000) := Null;
    Cursor c_Patient Is
      Select ����, �Ա�, Decode(�Ա�, '��', 'M', 'Ů', 'F', 'O') As �Ա����, ��������, ����id, ��ϵ�˵�ַ, ��ͥ�绰, ��ϵ�˵绰, ����״��, ���֤��, ��ǰ����id,
             ��ǰ����id, ��ǰ���� As ����, ���￨��, ��Ժʱ��, ��Ժʱ��
      From ������Ϣ
      Where ����id = ����id_In;
    r_Patient c_Patient%RowType;
  
    Cursor c_Order Is
      Select ��ҳid, Ӥ��, Decode(������Դ, 1, 'OUTPAT', 2, 'INPAT', 'UNK') As ������Դ, ����ҽ��, ����ʱ��, У�Ի�ʿ, ҽ������, ������־, ִ�п���id
      From ����ҽ����¼
      Where ID = ҽ��id_In;
    r_Order c_Order%RowType;
  
    Cursor c_Diagnose Is
      Select ������� || Decode(Nvl(�Ƿ�����, 0), 0, '', ' (��)') As �ٴ����
      From �������ҽ�� A, ������ϼ�¼ B
      Where a.ҽ��id = ҽ��id_In And a.���id = b.Id;
    r_Diagnose c_Diagnose%RowType;
  
    --��ȡָ�����������
    Procedure p_Get_Rowtype(Table_In In Varchar2) Is
    Begin
      If Table_In = '������Ϣ' Then
        Open c_Patient;
        Fetch c_Patient
          Into r_Patient;
      Elsif Table_In = '����ҽ����¼' Then
        Open c_Order;
        Fetch c_Order
          Into r_Order;
      Elsif Table_In = '������ϼ�¼' Then
        Open c_Diagnose;
        Fetch c_Diagnose
          Into r_Diagnose;
      End If;
    Exception
      When Others Then
        Null;
    End p_Get_Rowtype;
  
  Begin
    Case
    --ֱ�ӷ��ص�����Ԫ��
      When Ԫ����_In = 'ҽ��ID' Then
        v_Return := ҽ��id_In;
      When Ԫ����_In = '����ID' Then
        v_Return := ����id_In;
      
    --�������Ա𵥶�����������Ӥ��
      When Instr(',����,�Ա�,�Ա����,��������,', ',' || Ԫ����_In || ',') > 0 Then
        p_Get_Rowtype('����ҽ����¼');
        p_Get_Rowtype('������Ϣ');
        If Nvl(r_Order.Ӥ��, 0) = 0 Then
          If Ԫ����_In = '����' Then
            v_Return := r_Patient.����;
          Elsif Ԫ����_In = '�Ա�' Then
            v_Return := r_Patient.�Ա�;
          Elsif Ԫ����_In = '�Ա����' Then
            v_Return := r_Patient.�Ա����;
          Elsif Ԫ����_In = '��������' Then
            v_Return := To_Char(r_Patient.��������, 'YYYYMMDDMISS');
          End If;
        Else
          If Ԫ����_In = '����' Then
            Select Decode(Ӥ������, Null, r_Patient.���� || '֮Ӥ' || Trim(To_Char(���, '9')), Ӥ������) As Ӥ������
            Into v_Return
            From ������������¼
            Where ����id = ����id_In And ��ҳid = r_Order.��ҳid And ��� = Nvl(r_Order.Ӥ��, 0);
          Elsif Instr('�Ա�', Ԫ����_In) > 0 Then
            Select Ӥ���Ա�
            Into v_Return
            From ������������¼
            Where ����id = ����id_In And ��ҳid = r_Order.��ҳid And ��� = Nvl(r_Order.Ӥ��, 0);
            If Ԫ����_In = '�Ա����' Then
              Select Decode(v_Return, '��', 'M', 'Ů', 'F', 'O') Into v_Return From Dual;
            End If;
          Elsif Ԫ����_In = '��������' Then
            Select ����ʱ��
            Into v_Return
            From ������������¼
            Where ����id = ����id_In And ��ҳid = r_Order.��ҳid And ��� = Nvl(r_Order.Ӥ��, 0);
            v_Return := To_Char(v_Return, 'YYYYMMDDMISS');
          End If;
        End If;
      
    --��ѯ������Ϣ���ص�Ԫ��
      When Instr(',��ϵ�˵�ַ,��ͥ�绰,��ϵ�˵绰,����״��,���֤��,����,���￨��,��Ժʱ��,��Ժʱ��,', ',' || Ԫ����_In || ',') > 0 Then
        p_Get_Rowtype('������Ϣ');
        Case Ԫ����_In
          When '��ϵ�˵�ַ' Then
            v_Return := r_Patient.��ϵ�˵�ַ;
          When '��ͥ�绰' Then
            v_Return := r_Patient.��ͥ�绰;
          When '��ϵ�˵绰' Then
            v_Return := r_Patient.��ϵ�˵绰;
          When '����״��' Then
            v_Return := r_Patient.����״��;
          When '���֤��' Then
            v_Return := r_Patient.���֤��;
          When '����' Then
            v_Return := r_Patient.����;
          When '���￨��' Then
            v_Return := r_Patient.���￨��;
          When '��Ժʱ��' Then
            v_Return := To_Char(r_Patient.��Ժʱ��, 'YYYYMMDDMISS');
          When '��Ժʱ��' Then
            v_Return := To_Char(r_Patient.��Ժʱ��, 'YYYYMMDDMISS');
          Else
            v_Return := '';
        End Case;
        --��ѯҽ�����ص�Ԫ��
      When Instr(',������Դ,����ҽ��,����ʱ��,У�Ի�ʿ,ҽ������,������־,������־����,', ',' || Ԫ����_In || ',') > 0 Then
        p_Get_Rowtype('����ҽ����¼');
        Case Ԫ����_In
          When '������Դ' Then
            v_Return := r_Order.������Դ;
          When '����ҽ��' Then
            v_Return := r_Order.����ҽ��;
          When '����ʱ��' Then
            v_Return := To_Char(r_Order.����ʱ��, 'YYYYMMDDMISS');
          When 'У�Ի�ʿ' Then
            v_Return := r_Order.У�Ի�ʿ;
          When 'ҽ������' Then
            v_Return := r_Order.ҽ������;
          When '������־' Then
            v_Return := r_Order.������־;
        End Case;
        --��ѯ��ϼ�¼���ص�Ԫ��
      When Ԫ����_In = '�ٴ����' Then
        p_Get_Rowtype('������ϼ�¼');
        v_Return := r_Diagnose.�ٴ����;
      
      Else
        --���в�ѯSQL����ֵ��Ԫ��
        If Ԫ����_In = 'ִ��վ��' Then
          p_Get_Rowtype('����ҽ����¼');
          Select Decode(վ��, 1, 'SITE0002', 2, 'SITE0001', 3, 'SITE0003', 'SITE0001')
          Into v_Return
          From ���ű�
          Where ID = r_Order.ִ�п���id;
        End If;
        If Ԫ����_In = '��ǰ��������' Then
          p_Get_Rowtype('������Ϣ');
          Select ���� Into v_Return From ���ű� Where ID = r_Patient.��ǰ����id;
        End If;
        If Ԫ����_In = '��������' Then
          p_Get_Rowtype('������Ϣ');
          Select ���� Into v_Return From ���ű� Where ID = r_Patient.��ǰ����id;
        End If;
        If Ԫ����_In = '��ʶ��' Then
          Select Decode(a.������Դ, 1, c.�����, 2, Decode(c.סԺ��, Null, c.�����, c.סԺ��), 4, c.������, c.�����)
          Into v_Return
          From ����ҽ����¼ A, ������Ϣ C
          Where a.����id = c.����id And a.Id = ҽ��id_In;
        End If;
    End Case;
  
    Return Trim(v_Return);
  Exception
    When Others Then
      Return Null;
  End Ris_Replace_Element_Value;

  --14��ɾ��RIS��Ժ���ò���
  Procedure Ris��Ժ����_Delete Is
  Begin
    Delete From Ris��Ժ����;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris��Ժ����_Delete;

  --15������RISRis��Ժ���ò���
  Procedure Ris��Ժ����_Update
  (
    Id_In           Ris��Ժ����.Id%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    �û���_In       Ris��Ժ����.�û���%Type,
    ����_In         Ris��Ժ����.����%Type,
    ���ݿ������_In Ris��Ժ����.���ݿ������%Type
  ) Is
  
  Begin
  
    Insert Into Ris��Ժ����
      (ID, ҽԺ����, ҽԺ����, �û���, ����, ���ݿ������)
    Values
      (Id_In, ҽԺ����_In, ҽԺ����_In, �û���_In, ����_In, ���ݿ������_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris��Ժ����_Update;

End b_Zlxwinterface;
/

--123112:��¶¶,2018-03-27,���鴦����������Ϣ¼����ҩ����
Create Or Replace Procedure Zl_��������Ϣ_Insert
(
  ����id_In         In ������Ϣ�ӱ�.����id%Type,
  ���֤��_In       In ������Ϣ.���֤��%Type,
  ����������_In     In ������Ϣ�ӱ�.��Ϣֵ%Type,
  ���������֤��_In In ������Ϣ�ӱ�.��Ϣֵ%Type,
  ����id_In         In ������Ϣ�ӱ�.����id%Type,
  �������Ա�_In     In ������Ϣ�ӱ�.��Ϣֵ%Type := Null,
  ����������_In     In ������Ϣ�ӱ�.��Ϣֵ%Type := Null,
  �����˵绰_In     In ������Ϣ�ӱ�.��Ϣֵ%Type := Null,
  ��ҩ����_In       In ������Ϣ�ӱ�.��Ϣֵ%Type := Null
) As
Begin
  --�޸Ĳ������֤�� 
  Update ������Ϣ
  Set ���֤�� = ���֤��_In
  Where ����id = ����id_In And (���֤�� Is Null Or ���֤�� <> ���֤��_In);

  Update ������Ϣ�ӱ�
  Set ��Ϣֵ = ���֤��_In
  Where ����id = ����id_In And Nvl(����id, 0) = Nvl(����id_In, 0) And ��Ϣ�� = '�������֤��';
  If Sql%RowCount = 0 Then
    Insert Into ������Ϣ�ӱ�
      (����id, ����id, ��Ϣ��, ��Ϣֵ)
    Values
      (����id_In, ����id_In, '�������֤��', ���֤��_In);
  End If;
  --�޸Ĳ�����Ϣ�ӱ������������������������֤�� 
  If ����������_In Is Null Then
    Delete From ������Ϣ�ӱ� Where ����id = ����id_In And Nvl(����id, 0) = Nvl(����id_In, 0) And ��Ϣ�� = '����������';
    Delete From ������Ϣ�ӱ�
    Where ����id = ����id_In And Nvl(����id, 0) = Nvl(����id_In, 0) And ��Ϣ�� = '���������֤��';
    Delete From ������Ϣ�ӱ� Where ����id = ����id_In And Nvl(����id, 0) = Nvl(����id_In, 0) And ��Ϣ�� = '�������Ա�';
    Delete From ������Ϣ�ӱ� Where ����id = ����id_In And Nvl(����id, 0) = Nvl(����id_In, 0) And ��Ϣ�� = '����������';
    Delete From ������Ϣ�ӱ� Where ����id = ����id_In And Nvl(����id, 0) = Nvl(����id_In, 0) And ��Ϣ�� = '�����˵绰';
    Delete From ������Ϣ�ӱ� Where ����id = ����id_In And Nvl(����id, 0) = Nvl(����id_In, 0) And ��Ϣ�� = '��ҩ����';
  Else
    Update ������Ϣ�ӱ�
    Set ��Ϣֵ = ����������_In
    Where ����id = ����id_In And Nvl(����id, 0) = Nvl(����id_In, 0) And ��Ϣ�� = '����������';
    If Sql%RowCount = 0 Then
      Insert Into ������Ϣ�ӱ�
        (����id, ����id, ��Ϣ��, ��Ϣֵ)
      Values
        (����id_In, ����id_In, '����������', ����������_In);
    End If;
  
    Update ������Ϣ�ӱ�
    Set ��Ϣֵ = ���������֤��_In
    Where ����id = ����id_In And Nvl(����id, 0) = Nvl(����id_In, 0) And ��Ϣ�� = '���������֤��';
    If Sql%RowCount = 0 Then
      Insert Into ������Ϣ�ӱ�
        (����id, ����id, ��Ϣ��, ��Ϣֵ)
      Values
        (����id_In, ����id_In, '���������֤��', ���������֤��_In);
    End If;
  
    Update ������Ϣ�ӱ�
    Set ��Ϣֵ = �������Ա�_In
    Where ����id = ����id_In And Nvl(����id, 0) = Nvl(����id_In, 0) And ��Ϣ�� = '�������Ա�';
    If Sql%RowCount = 0 Then
      Insert Into ������Ϣ�ӱ�
        (����id, ����id, ��Ϣ��, ��Ϣֵ)
      Values
        (����id_In, ����id_In, '�������Ա�', �������Ա�_In);
    End If;
  
    Update ������Ϣ�ӱ�
    Set ��Ϣֵ = ����������_In
    Where ����id = ����id_In And Nvl(����id, 0) = Nvl(����id_In, 0) And ��Ϣ�� = '����������';
    If Sql%RowCount = 0 Then
      Insert Into ������Ϣ�ӱ�
        (����id, ����id, ��Ϣ��, ��Ϣֵ)
      Values
        (����id_In, ����id_In, '����������', ����������_In);
    End If;
  
    Update ������Ϣ�ӱ�
    Set ��Ϣֵ = �����˵绰_In
    Where ����id = ����id_In And Nvl(����id, 0) = Nvl(����id_In, 0) And ��Ϣ�� = '�����˵绰';
    If Sql%RowCount = 0 Then
      Insert Into ������Ϣ�ӱ�
        (����id, ����id, ��Ϣ��, ��Ϣֵ)
      Values
        (����id_In, ����id_In, '�����˵绰', �����˵绰_In);
    End If;

    Update ������Ϣ�ӱ�
    Set ��Ϣֵ = ��ҩ����_In
    Where ����id = ����id_In And Nvl(����id, 0) = Nvl(����id_In, 0) And ��Ϣ�� = '��ҩ����';
    If Sql%RowCount = 0 Then
      Insert Into ������Ϣ�ӱ� (����id, ����id, ��Ϣ��, ��Ϣֵ) Values (����id_In, ����id_In, '��ҩ����', ��ҩ����_In);
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��������Ϣ_Insert;
/
--122954:��ΰ��,2018-03-26,����������ҩ
Create Or Replace Procedure Zl_Lob_Append
(
  Tab_In     In Number,
  Key_In     In Varchar2,
  Txt_In     In Varchar2, --16���Ƶ��ļ�Ƭ�λ�����Ƭ��
  Cls_In     In Number := 0, --�Ƿ����ԭ�������ݣ���һƬ�δ���ʱΪ1���Ժ�Ϊ0
  Lobtype_In In Number := 0 --0-BLOB;1-CLOB
  --����˵����
  --Tab_In������LOB�����ݱ�
  --        0-�������ͼ��;1-�����ļ���ʽ;2-�����ļ�ͼ��;3-�������ĸ�ʽ;4-��������ͼ��;
  --        5/21-���Ӳ�����ʽ;6-���Ӳ���ͼ��;7-����ҳ���ʽ��8-���Ӳ�������;9-�����ص����
  --        10-�ٴ�·���ļ�,11-�ٴ�·��ͼ��;14-��Ա֤���¼;15-��Ա��;16-��Ա��Ƭ;
  --        19-������չ��Ϣ;20-��Ա��չ��Ϣ;22-ҽ����������;
  --        23-��Ӧ����Ƭ;24-�Զ������뵥�ļ�;25-ҽ�����뵥�ļ�
  --        26-����·���ļ�,27-������Ƭ,28-��ѯͼƬԪ��,29-��ѯ����Ŀ¼,30-����������ҩ����
  --Key_In�����ݼ�¼�Ĺؼ���
  --Txt_In��16���Ƶ��ļ�Ƭ�λ�����Ƭ��
  --Cls_In���Ƿ����ԭ�������ݣ���һƬ�δ���ʱΪ1���Ժ�Ϊ0
  --Lobtype_In:--0-BLOB;1-CLOB
) Is
  l_Blob Blob;
  l_Clob Clob;
  t_Key  t_Strlist;
Begin
  If Tab_In = 0 Then
    If Cls_In = 1 Then
      Update �������ͼ�� Set ͼ�� = Empty_Blob() Where ���� = Key_In;
    End If;
    Select ͼ�� Into l_Blob From �������ͼ�� Where ���� = Key_In For Update;
  Elsif Tab_In = 1 Then
    If Cls_In = 1 Then
      Update �����ļ���ʽ Set ���� = Empty_Blob() Where �ļ�id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into �����ļ���ʽ (�ļ�id, ����) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ���� Into l_Blob From �����ļ���ʽ Where �ļ�id = To_Number(Key_In) For Update;
  Elsif Tab_In = 2 Then
    If Cls_In = 1 Then
      Update �����ļ�ͼ�� Set ͼ�� = Empty_Blob() Where ����id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into �����ļ�ͼ�� (����id, ͼ��) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ͼ�� Into l_Blob From �����ļ�ͼ�� Where ����id = To_Number(Key_In) For Update;
  Elsif Tab_In = 3 Then
    If Cls_In = 1 Then
      Update �������ĸ�ʽ Set ���� = Empty_Blob() Where �ļ�id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into �������ĸ�ʽ (�ļ�id, ����) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ���� Into l_Blob From �������ĸ�ʽ Where �ļ�id = To_Number(Key_In) For Update;
  Elsif Tab_In = 4 Then
    If Cls_In = 1 Then
      Update ��������ͼ�� Set ͼ�� = Empty_Blob() Where ����id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into ��������ͼ�� (����id, ͼ��) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ͼ�� Into l_Blob From ��������ͼ�� Where ����id = To_Number(Key_In) For Update;
  Elsif Tab_In = 5 Then
    If Cls_In = 1 Then
      Update ���Ӳ�����ʽ Set ���� = Empty_Blob() Where �ļ�id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into ���Ӳ�����ʽ (�ļ�id, ����) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ���� Into l_Blob From ���Ӳ�����ʽ Where �ļ�id = To_Number(Key_In) For Update;
  Elsif Tab_In = 6 Then
    If Cls_In = 1 Then
      Update ���Ӳ���ͼ�� Set ͼ�� = Empty_Blob() Where ����id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into ���Ӳ���ͼ�� (����id, ͼ��) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ͼ�� Into l_Blob From ���Ӳ���ͼ�� Where ����id = To_Number(Key_In) For Update;
  Elsif Tab_In = 7 Then
    If Cls_In = 1 Then
      Update ����ҳ���ʽ
      Set ͼ�� = Empty_Blob()
      Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3);
    End If;
    Select ͼ��
    Into l_Blob
    From ����ҳ���ʽ
    Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3)
    For Update;
  Elsif Tab_In = 8 Then
    If Cls_In = 1 Then
      Update ���Ӳ�������
      Set ���� = Empty_Blob()
      Where ����id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And ��� = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
    Select ����
    Into l_Blob
    From ���Ӳ�������
    Where ����id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And ��� = Substr(Key_In, Instr(Key_In, ',') + 1)
    For Update;
  Elsif Tab_In = 9 Then
    If Cls_In = 1 Then
      Update �����ص���� Set ���ͼ�� = Empty_Blob() Where ��� = To_Number(Key_In);
    End If;
    Select ���ͼ�� Into l_Blob From �����ص���� Where ��� = To_Number(Key_In) For Update;
  Elsif Tab_In = 10 Then
    If Cls_In = 1 Then
      Update �ٴ�·���ļ�
      Set ���� = Empty_Blob()
      Where ·��id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
            �ļ��� = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
    Select ����
    Into l_Blob
    From �ٴ�·���ļ�
    Where ·��id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And �ļ��� = Substr(Key_In, Instr(Key_In, ',') + 1)
    For Update;
  Elsif Tab_In = 11 Then
    If Cls_In = 1 Then
      Update �ٴ�·��ͼ�� Set ͼ�� = Empty_Blob() Where ID = To_Number(Key_In);
    End If;
    Select ͼ�� Into l_Blob From �ٴ�·��ͼ�� Where ID = To_Number(Key_In) For Update;
  Elsif Tab_In = 12 Then
    If Cls_In = 1 Then
      Update ����ҳ���ʽ
      Set ҳü�ļ� = Empty_Blob()
      Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3);
    End If;
    Select ҳü�ļ�
    Into l_Blob
    From ����ҳ���ʽ
    Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3)
    For Update;
  Elsif Tab_In = 13 Then
    If Cls_In = 1 Then
      Update ����ҳ���ʽ
      Set ҳ���ļ� = Empty_Blob()
      Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3);
    End If;
    Select ҳ���ļ�
    Into l_Blob
    From ����ҳ���ʽ
    Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3)
    For Update;
  Elsif Tab_In = 14 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update ��Ա֤���¼ Set ǩ����Ϣ = Empty_Clob() Where ��Աid = To_Number(t_Key(1)) And Certsn = t_Key(2);
    End If;
    Select ǩ����Ϣ Into l_Clob From ��Ա֤���¼ Where ��Աid = To_Number(t_Key(1)) And Certsn = t_Key(2) For Update;
  Elsif Tab_In = 15 Then
    If Cls_In = 1 Then
      Update ��Ա�� Set ǩ��ͼƬ = Empty_Blob() Where ID = To_Number(Key_In);
    End If;
    Select ǩ��ͼƬ Into l_Blob From ��Ա�� Where ID = To_Number(Key_In) For Update;
    Update ��Ա�� Set ����޸�ʱ�� = Sysdate Where ID = To_Number(Key_In);
  Elsif Tab_In = 16 Then
    If Cls_In = 1 Then
      Update ��Ա��Ƭ Set ��Ƭ = Empty_Blob() Where ��Աid = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into ��Ա��Ƭ (��Աid, ��Ƭ) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ��Ƭ Into l_Blob From ��Ա��Ƭ Where ��Աid = To_Number(Key_In) For Update;
    Update ��Ա�� Set ����޸�ʱ�� = Sysdate Where ID = To_Number(Key_In);
  Elsif Tab_In = 19 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update ������չ��Ϣ Set ͼƬ = Empty_Blob() Where ����id = To_Number(t_Key(1)) And ��Ŀ = t_Key(2);
    End If;
    Select ͼƬ Into l_Blob From ������չ��Ϣ Where ����id = To_Number(t_Key(1)) And ��Ŀ = t_Key(2) For Update;
    Update ���ű� Set ����޸�ʱ�� = Sysdate Where ID = To_Number(t_Key(1));
  Elsif Tab_In = 20 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update ��Ա��չ��Ϣ Set ͼƬ = Empty_Blob() Where ��Աid = To_Number(t_Key(1)) And ��Ŀ = t_Key(2);
    End If;
    Select ͼƬ Into l_Blob From ��Ա��չ��Ϣ Where ��Աid = To_Number(t_Key(1)) And ��Ŀ = t_Key(2) For Update;
    Update ��Ա�� Set ����޸�ʱ�� = Sysdate Where ID = To_Number(t_Key(1));
  Elsif Tab_In = 21 Then
    If Cls_In = 1 Then
      Update ���Ӳ�����ʽ Set �ı����� = Empty_Clob() Where �ļ�id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into ���Ӳ�����ʽ (�ļ�id, �ı�����) Values (To_Number(Key_In), Empty_Clob());
      End If;
    End If;
    Select �ı����� Into l_Clob From ���Ӳ�����ʽ Where �ļ�id = To_Number(Key_In) For Update;
  Elsif Tab_In = 22 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update ҽ���������� Set ���� = Empty_Blob() Where ID = To_Number(t_Key(1));
    End If;
    Select ���� Into l_Blob From ҽ���������� Where ID = To_Number(t_Key(1)) For Update;
  Elsif Tab_In = 23 Then
    If To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=0 Then
      If Cls_In = 1 Then
        Update ��Ӧ����Ƭ Set ���֤����Ƭ = Empty_Blob() Where ��Ӧ��ID = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
        If Sql%RowCount = 0 Then
          Insert Into ��Ӧ����Ƭ (��Ӧ��ID, ���֤����Ƭ,ִ�պ���Ƭ,��Ȩ����Ƭ) Values (To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)), Empty_Blob(), Empty_Blob(), Empty_Blob());
        End If;
      End If;
      Select ���֤����Ƭ Into l_Blob From ��Ӧ����Ƭ Where ��Ӧ��ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) For Update;
    Elsif  To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=1 Then
      If Cls_In = 1 Then
        Update ��Ӧ����Ƭ Set ִ�պ���Ƭ = Empty_Blob() Where ��Ӧ��ID = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
        If Sql%RowCount = 0 Then
          Insert Into ��Ӧ����Ƭ (��Ӧ��ID, ���֤����Ƭ,ִ�պ���Ƭ,��Ȩ����Ƭ) Values (To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)), Empty_Blob(), Empty_Blob(), Empty_Blob());
        End If;
      End If;
      Select ִ�պ���Ƭ Into l_Blob From ��Ӧ����Ƭ Where ��Ӧ��ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) For Update;
    Elsif To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=2 Then
     If Cls_In = 1 Then
        Update ��Ӧ����Ƭ Set ��Ȩ����Ƭ = Empty_Blob() Where ��Ӧ��ID = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
        If Sql%RowCount = 0 Then
          Insert Into ��Ӧ����Ƭ (��Ӧ��ID, ���֤����Ƭ,ִ�պ���Ƭ,��Ȩ����Ƭ) Values (To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)), Empty_Blob(), Empty_Blob(), Empty_Blob());
        End If;
      End If;
      Select ��Ȩ����Ƭ Into l_Blob From ��Ӧ����Ƭ Where ��Ӧ��ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) For Update;
    End If;
  Elsif Tab_In = 24 Then
    If Cls_In = 1 Then
      Update �Զ������뵥�ļ�
      Set ���� = Empty_Clob()
      Where �ļ�id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
            ��� = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) ;
    End If;
    Select ����
    Into l_Clob
    From �Զ������뵥�ļ�
    Where �ļ�id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And ��� = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) 
    For Update;
  ElsIf Tab_In = 25 Then
    If Cls_In = 1 Then
      Update ҽ�����뵥�ļ�
      Set ���� = Empty_Clob()
      Where ҽ��id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 
            ��� = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))  ;
    End If;
    Select ����
    Into l_Clob
    From ҽ�����뵥�ļ�
    Where ҽ��id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And ��� = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) 
    For Update;
  Elsif Tab_In = 26 Then
    If Cls_In = 1 Then
      Update ����·���ļ�
      Set ���� = Empty_Blob()
      Where ·��id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
            �ļ��� = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
    Select ����
    Into l_Blob
    From ����·���ļ�
    Where ·��id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And �ļ��� = Substr(Key_In, Instr(Key_In, ',') + 1)
    For Update;
  Elsif Tab_In = 27 Then
    If Cls_In = 1 Then
      Update ������Ƭ Set ��Ƭ = Empty_Blob() Where ����id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into ������Ƭ (����id, ��Ƭ) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ��Ƭ Into l_Blob From ������Ƭ Where ����id = To_Number(Key_In) For Update;
  Elsif Tab_In = 28 Then
    If Cls_In = 1 Then
      Update ��ѯͼƬԪ�� Set ͼ�� = Empty_Blob() Where ��� = To_Number(Key_In);
    End If;
    Select ͼ�� Into l_Blob From ��ѯͼƬԪ�� Where ��� = To_Number(Key_In) For Update;
  Elsif Tab_In = 29 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update ��ѯ����Ŀ¼
      Set �����ı� = Empty_Clob()
      Where ҳ����� = To_Number(t_Key(1)) And ������� = To_Number(t_Key(2));
    End If;
    Select �����ı�
    Into l_Clob
    From ��ѯ����Ŀ¼
    Where ҳ����� = To_Number(t_Key(1)) And ������� = To_Number(t_Key(2))
    For Update;
  Elsif Tab_In = 30 Then
    If Cls_In = 1 Then
      Insert Into ����������ҩ���� (��������) Values (Empty_Clob());
    End If;
    Select �������� Into l_Clob From ����������ҩ���� For Update;
  End If;

  If Not Txt_In Is Null Then
    If Lobtype_In = 1 Then
      Dbms_Lob.Writeappend(l_Clob, Length(Txt_In), Txt_In);
    Else
      Dbms_Lob.Writeappend(l_Blob, Length(Hextoraw(Txt_In)) / 2, Hextoraw(Txt_In));
    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Lob_Append;
/

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
Begin
  Update ������ĿĿ¼ Set �Ƽ����� = �Ƽ�����_In Where ID = ������Ŀid_In;
  If �Ƿ�ɾ��_In = 1 Then
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
    b_Message.Zlhis_Dict_054(������Ŀid_In);
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
  End If;

  If �۸�ȼ�_In Is Null Then
    b_Message.Zlhis_Dict_053(�շ�ϸĿid_In);
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
    b_Message.Zlhis_Dict_053(�շ�ϸĿid_In);
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
    b_Message.Zlhis_Dict_053(�շ�ϸĿid_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�շѼ�Ŀ_Insert;
/

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


--122954:��ΰ��,2018-03-26,����������ҩ

Create Or Replace Procedure Zl_����������ҩ����_Update
(
  ����id_In In ������Ϣ.����id%Type,
  ��ҳid_In In ������ҳ.��ҳid%Type,
  �Һ�id_In In ���˹Һż�¼.Id%Type := Null
) As

  --------------------------------------------------------------------------------------------------
  --����:������ҩ��⴫��ֵ�����غ�����ҩ����ֵ
  --����:Xml_Return  ���ز����XML��
  -- <details_xml>
  --    <patient_info>
  --      <info name="��������" value="28114.45"/>
  --     <info name="��������" value="����"/>
  --      <info name="�Ա�" value="Ů"/>
  --      <info name="ְҵ" value="�˶�Ա"/>
  --      <info name="����" value="1"/>
  --      <info name="����" value="1"/>
  --      <info name="�ι��ܲ�ȫ" value="1">
  --      <info name="���ظι��ܲ�ȫ" value="1">
  --      <info name="�����ܲ�ȫ" value="1">
  --      <info name="���������ܲ�ȫ" value="1">
  --      <info name="���" value="J18.000"/> --��ϴ����룬�������Զ��ŷָ�
  --    </patient_info>
  --    <medicine_info>
  --      <medicine>
  --        <info name="ҽ��ID" value="1"/>
  --        <info name="��λ��" value="86900967000160" main="46d64420-8319-4768-9a11-f4b0f5e4ce7a"/> --mainֵ�ǹ̶���
  --        <info name="������ĿID" value="67232" main="4e19df1c-c1b9-4a43-a83d-0741a19961ab"/>
  --        <info name="��Һ���" value="1"/>
  --        <info name="������λ" value="ml"/>
  --        <info name="������" value="250"/>
  --        <info name="������-������" value="5.21"/>--������-������= trunc(������/��������,2)
  --        <info name="������-�����" value="170.3"/>--������-�����= trunc(������/(0.0061*�������+0.0128*��������-0.1529),2)
  --        <info name="ÿ����" value="250"/>
  --        1.ÿ����=������*��Ƶ��
  --        2.��Ƶ�μ��㣺
  --            a.���÷�Χ=-1����Ƶ��=1
  --            b.�����λ=�� and Ƶ�ʼ��=1����Ƶ��=Ƶ�ʴ���
  --            c.�����λ=�� and Ƶ�ʼ��>1 and Ƶ�ʴ���=1����Ƶ��=1
  --            d.�����λ=Сʱ and Ƶ�ʼ��<=24,��Ƶ��=24/Ƶ�ʼ��*Ƶ�ʴ���
  --            e.�����λ=Сʱ and Ƶ�ʼ��>24 and Ƶ�ʴ���=1����Ƶ��=1
  --            f.�����λ=�� and Ƶ�ʴ���=1����Ƶ��=1
  --        <info name="ÿ����-������" value="5.21"/>  --trunc(ÿ����/��������,2)
  --        <info name="ÿ����-�����" value="170.3"/>  --ÿ����-�����= trunc(ÿ����/(0.0061*�������+0.0128*��������-0.1529),2)
  --        <info name="��ҩƵ��" value="ÿ��һ��"/>
  --        <info name="��ҩ;��" value="001"/>
  --      </medicine>
  --    </medicine_info>
  --  </details_xml>
  --------------------------------------------------------------------------------------------------
  Xml_Ret             Xmltype;
  Xml_Document        Xmldom.Domdocument;
  Xml_Nodelist        Xmldom.Domnodelist;
  Xml_Domelement      Xmldom.Domelement;
  Xml_Domnamednodemap Xmldom.Domnamednodemap;
  Xml_Node_Med        Xmldom.Domnode;
  Xml_Node            Xmldom.Domnode;
  Xml_Node_New        Xmldom.Domnode;
  ----------------------------------
  n_��� Number(10, 2); --��λ:cm
  n_���� Number(10, 2); --����:KG

  l_Clob    Clob;
  v_Err_Msg Varchar2(2000);
  v_Temp    Varchar2(200);
  v_Value   Varchar2(200);
  n_Nodenum Number(5);
  Err_Item Exception;
Begin
  --��
  --��CLOB������ȡ��v_XML��
  Select �������� Into l_Clob From ����������ҩ����;
  Xml_Ret        := Xmltype(l_Clob); --���溯������ֵ
  Xml_Document   := Xmldom.Newdomdocument(Xml_Ret);
  Xml_Domelement := Xmldom.Getdocumentelement(Xml_Document);
  Xml_Nodelist   := Xmldom.Getelementsbytagname(Xml_Domelement, 'patient_info');
  --��ȡpatient_info/INfo�ڵ�
  Xml_Nodelist := Xmldom.Getchildnodes(Xmldom.Item(Xml_Nodelist, 0));
  n_Nodenum    := Xmldom.Getlength(Xml_Nodelist);
  For I In 0 .. n_Nodenum - 1 Loop
    Xml_Domnamednodemap := Xmldom.Getattributes(Xmldom.Item(Xml_Nodelist, I));
    v_Temp              := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'name'));
    If v_Temp = '���' Then
      n_��� := Nvl(To_Number(Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'value'))), 0);
    End If;
    If v_Temp = '����' Then
      n_���� := Nvl(To_Number(Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'value'))), 0);
    End If;
  End Loop;
  --��ȡmedicine/INfo�ڵ�

  Xml_Nodelist := Xmldom.Getelementsbytagname(Xml_Domelement, 'medicine');
  n_Nodenum    := Xmldom.Getlength(Xml_Nodelist);
  For I In 0 .. n_Nodenum - 1 Loop
    Xml_Node_Med := Xmldom.Item(Xml_Nodelist, I); --ȡ��һ�����ӽڵ�medicine
    Xml_Nodelist := Xmldom.Getchildnodes(Xml_Node_Med); --infos
    Xml_Node     := Xmldom.Getfirstchild(Xml_Node_Med); --ȡ��һ�����ӽڵ�
    While Not Xmldom.Isnull(Xml_Node) Loop
      Xml_Domnamednodemap := Xmldom.Getattributes(Xml_Node);
      v_Temp              := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'name'));
      v_Value             := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'value'));
      If v_Temp = '������' Then
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', '������-������');
        If n_���� > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', Trunc(To_Number(v_Value) / n_����, 2));
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
        --������-�����trunc(ÿ����/(0.0061*�������+0.0128*��������-0.1529),2)
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', '������-�����');
        If n_���� > 0 And n_��� > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value',
                              Trunc(To_Number(v_Value) / (0.0061 * n_��� + 0.0128 * n_���� - 0.1529), 2));
        
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
      End If;
    
      If v_Temp = 'ÿ����' Then
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', 'ÿ����-������');
        If n_���� > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', Trunc(To_Number(v_Value) / n_����, 2));
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
      
        --ÿ����-�����
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', 'ÿ����-�����');
        If n_���� > 0 And n_��� > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value',
                              Trunc(To_Number(v_Value) / (0.0061 * n_��� + 0.0128 * n_���� - 0.1529), 2));
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
      End If;
      --ȡ��һ���ֵܽڵ�
      Xml_Node := Xmldom.Getnextsibling(Xml_Node);
    End Loop;
  End Loop;

  --����������ֵ������ʱ��,ZLHIS���������ǰ��ȡ����Ϊ���������Ʒ���ֵ���ܳ���4000���ƣ�
  Update ����������ҩ���� Set �������� = Xml_Ret.Getclobval();

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����������ҩ����_Update;
/

--122954:��ΰ��,2018-03-26,����������ҩ

Create Or Replace Function Zl_Read_����������ҩ����(Pos_In In Number
                                            --����˵����
                                            --Pos_In����0��ʼ���϶�ȡ��ֱ������Ϊ��
                                            ) Return Varchar2 Is
  l_Clob    Clob;
  v_Buffer  Varchar2(32767);
  n_Amount  Number := 2000;
  n_Offset  Number := 1;
  v_Err_Msg Varchar2(2000);
  Err_Item Exception;
Begin
  Select �������� Into l_Clob From ����������ҩ����;
  n_Offset := n_Offset + Pos_In * n_Amount;

  If l_Clob Is Null Then
    v_Buffer := Null;
  Else
    Begin
      Dbms_Lob.Read(l_Clob, n_Amount, n_Offset, v_Buffer);
    Exception
      When No_Data_Found Then
        v_Buffer := Null;
    End;
  End If;
  Return v_Buffer;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Read_����������ҩ����;
/

------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0004' Where ���=&n_System;
Commit;