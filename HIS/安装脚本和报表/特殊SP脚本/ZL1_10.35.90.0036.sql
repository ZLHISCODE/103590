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
--133498:Ƚ����,2018-11-06,֧�ֲ���������������תΪסԺ����
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_system,1131,'����Ǽ�',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'Zl_����תסԺ_������ת��','EXECUTE' From Dual) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_system,1137,'�������תסԺ',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'Zl_����תסԺ_������ת��','EXECUTE' From Dual) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_system,1604,'������Ժ',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'Zl_����תסԺ_������ת��','EXECUTE' From Dual) A;




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--133498:Ƚ����,2018-11-07,֧�ֲ���������������תΪסԺ����
Create Or Replace Procedure Zl_�������תסԺ_Insert
(
  No_In         סԺ���ü�¼.No%Type,
  סԺ��_In     סԺ���ü�¼.��ʶ��%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
  ��ҳid_In     סԺ���ü�¼.��ҳid%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
  ��Ժʱ��_In   סԺ���ü�¼.����ʱ��%Type,
  ��Ժ����id_In ����Ԥ����¼.����id%Type,
  �˷�ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type, --���ŵ����˷�ʱ,ÿ�ŵ��ݵ��˷�ʱ����ͬ,����ϵͳ��ǰʱ��
  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
  ��Ժ����id_In סԺ���ü�¼.���˲���id%Type := Null,
  ����_In       Number := 1,
  ����id_In     סԺ���ü�¼.����id%Type := Null,
  ԭ����id_In   סԺ���ü�¼.����id%Type := Null,
  ��������_In   Number := 1
) As
  --����_In:1-�����շѵ�;2-���ʵ�
  v_Billno   סԺ���ü�¼.No%Type;
  n_ʵ�պϼ� סԺ���ü�¼.ʵ�ս��%Type;
  n_����ֵ   �������.Ԥ�����%Type;

  n_����id     סԺ���ü�¼.���˲���id%Type;
  v_����       סԺ���ü�¼.����%Type;
  n_ҽ��С��id סԺ���ü�¼.ҽ��С��id%Type;

  n_��������id     ���ű�.Id%Type;
  n_����Ա���     ������ü�¼.����Ա���%Type;
  v_����Ա����     ������ü�¼.����Ա����%Type;
  v_������         ��Ա��.����%Type;
  n_����id         ������Ϣ.����id%Type;
  n_��˱�־       ������ҳ.��˱�־%Type;
  n_סԺ״̬       ������ҳ.״̬%Type;
  n_������˷�ʽ   Number(2);
  n_δ��ƽ�ֹ���� Number(2);

  n_����id  ������ü�¼.����id%Type;
  v_Err_Msg Varchar2(255);
  n_��id    ����ɿ����.Id%Type;
  Err_Item Exception;
  n_Count Number(18);
Begin
  n_��id := Zl_Get��id(����Ա����_In);
  If Nvl(��������_In, 0) = 1 Then
    If ����id_In Is Null Then
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    Else
      n_����id := ����id_In;
    End If;
  End If;

  If Nvl(��ҳid_In, 0) <> 0 Then
    n_������˷�ʽ   := Nvl(zl_GetSysParameter(185), 0);
    n_δ��ƽ�ֹ���� := Nvl(zl_GetSysParameter(215), 0);
    If n_������˷�ʽ = 1 Or n_δ��ƽ�ֹ���� = 1 Then
      Begin
        Select ����id
        Into n_����id
        From ������ü�¼
        Where NO = No_In And Mod(��¼����, 10) = Mod(����_In, 10) And Rownum = 1;
      Exception
        When Others Then
          n_����id := 0;
      End;
      Begin
        Select ��˱�־, ״̬
        Into n_��˱�־, n_סԺ״̬
        From ������ҳ
        Where ����id = Nvl(n_����id, 0) And ��ҳid = Nvl(��ҳid_In, 0);
      Exception
        When Others Then
          n_��˱�־ := 0;
          n_סԺ״̬ := 0;
      End;
      If n_δ��ƽ�ֹ���� = 1 And n_סԺ״̬ = 1 Then
        v_Err_Msg := '����δ���,��ֹ�Բ�����ط��õĲ���!';
        Raise Err_Item;
      End If;
    
      If n_������˷�ʽ = 1 Then
        If Nvl(n_��˱�־, 0) = 1 Then
          v_Err_Msg := '�ò���Ŀǰ������˷���,���ܽ��з�����ص���!';
          Raise Err_Item;
        End If;
        If Nvl(n_��˱�־, 0) = 2 Then
          v_Err_Msg := '�ò���Ŀǰ�Ѿ�����˷������,���ܽ��з�����ص���!';
          Raise Err_Item;
        End If;
      End If;
    End If;
  End If;

  If ԭ����id_In Is Null Then
    If Nvl(��������_In, 0) = 1 Then
      If Mod(����_In, 10) = 1 Then
        --ת�շѵ�
        --No_In;����Ա���_In,����Ա����_In,�˷�ʱ��_In,�����˷�_In(0-����תסԺ��������;1-�����˷�ģʽ),��Ժ����id_In,��ҳid_In
        Zl_����תסԺ_�շ�ת��(No_In, ����Ա���_In, ����Ա����_In, �˷�ʱ��_In, 0, ��Ժ����id_In, ��ҳid_In, Null, n_����id);
      Else
        --ת���ʵ�
        --No_In;����Ա���_In,����Ա����_In,�˷�ʱ��_In
        Zl_����תסԺ_����ת��(No_In, ����Ա���_In, ����Ա����_In, �˷�ʱ��_In);
      End If;
    End If;
    --����
    -- 1.��Ժ����ID_IN<>NULL ��Ϊ:��Ժ����ID_IN
    -- 2.��ҳid_In<>0 :
    If Nvl(��Ժ����id_In, 0) <> 0 Then
      n_����id := ��Ժ����id_In;
    Elsif Nvl(��ҳid_In, 0) <> 0 Then
      Begin
        Select Nvl(b.��ǰ����id, a.��ǰ����id), Nvl(b.��Ժ����, a.��ǰ����)
        Into n_����id, v_����
        From ������Ϣ A, ������ҳ B
        Where a.����id = b.����id(+) And a.����id = n_����id And b.��ҳid(+) = ��ҳid_In;
      Exception
        When Others Then
          n_����id := Null;
      End;
    End If;
  
    If Nvl(n_����id, 0) = 0 Then
      --����Ժ����Ϊ׼
      n_����id := Nvl(��Ժ����id_In, 0);
    End If;
  
    --��Ժ�Ǽ�֮ǰ,��ҳID��û�в���,����Ԥ����¼,����δ�����,δ��ҩƷ��¼,������ü�¼
    --���в���ID,��ҳID�����,ֻ����Ժ�ǼǺ��ٵ���Zl_�������תסԺ_Update��д
    Select ����id, ��������id, ������
    Into n_����id, n_��������id, v_������
    From ������ü�¼
    Where NO = No_In And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ In (1, 3) And Rownum = 1;
  
    --5.�������ʵ�
    --��Ҫ����Ƿ��Ѿ�ת��
    Select Count(*)
    Into n_Count
    From ������ü�¼ A
    Where NO = No_In And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ In (1, 3) And Exists
     (Select 1 From ������˼�¼ Where a.Id = ����id And ת��id Is Not Null) And Rownum <= 1;
    If n_Count >= 1 Then
      v_Err_Msg := '�����򲢷�ԭ��,�÷����Ѿ�������ת��,���ܼ�������!';
      Raise Err_Item;
    End If;
    If Mod(����_In, 10) = 1 Then
      --�շѰ��ս�����Ų������NO�Ž��д���
      n_ҽ��С��id := Zl_ҽ��С��_Get(n_��������id, v_������, n_����id, ��ҳid_In, ��Ժʱ��_In);
      v_Billno     := Nextno(14);
    
      Insert Into סԺ���ü�¼
        (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ����, ���˲���id, ���˿���id, �ѱ�, �շ����,
         �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����,
         ���ʷ���, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �Ƿ���, ҽ�����, �ɿ���id, ҽ��С��id)
        Select ���˷��ü�¼_Id.Nextval, 2, v_Billno, 1, ���, ��������, �۸񸸺�, 0, 2, ����id, ��ҳid_In, סԺ��_In, ����, �Ա�, ����, v_����,
               Decode(n_����id, Null, Null, 0, Null, n_����id), ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������,
               ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, 1, ��������id, ������, ��Ժʱ��_In, �˷�ʱ��_In,
               ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ������, ����Ա���, ����Ա����, ���ʵ�id, '�������ת��', �Ƿ���, ҽ�����, n_��id, n_ҽ��С��id
        From ������ü�¼
        Where NO = No_In And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ = 1 And Nvl(���ӱ�־, 0) Not In (8, 9);
    
      If Nvl(��������_In, 0) = 1 Then
        Update ������ü�¼ Set ��¼״̬ = 3 Where NO = No_In And Mod(��¼����, 10) In (1, 2) And ��¼״̬ = 1;
      End If;
    
      For r_Clinic In (Select Min(��¼����) As ��¼����, ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��,
                              ���մ���id, ���ձ���, ��������, ��ҩ����, ����, Sum(����) As ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����,
                              Sum(Ӧ�ս��) As Ӧ�ս��, Sum(ʵ�ս��) As ʵ�ս��, Sum(ͳ����) As ͳ����, ��������id, ������, ִ�в���id,
                              Max(ִ����) As ִ����, ������, Max(���ʵ�id) As ���ʵ�id, Max(�Ƿ���) As �Ƿ���, ����ʱ��, Min(ʵ��Ʊ��) As ʵ��Ʊ��,
                              Max(ִ��״̬) As ִ��״̬, Max(ִ��ʱ��) As ִ��ʱ��
                       From ������ü�¼
                       Where NO = No_In And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ In (2, 3) And
                             Nvl(���ӱ�־, 0) Not In (8, 9)
                       Group By ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���,
                                ��������, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, ��������id, ������, ִ�в���id, ������, ����ʱ��
                       Having Sum(����) <> 0) Loop
        Select ����Ա���, ����Ա����
        Into n_����Ա���, v_����Ա����
        From ������ü�¼
        Where NO = No_In And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ = 3 And Rownum < 2;
        Insert Into סԺ���ü�¼
          (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ����, ���˲���id, ���˿���id, �ѱ�, �շ����,
           �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���,
           ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �Ƿ���, �ɿ���id, ҽ��С��id, ִ��״̬, ִ��ʱ��)
        Values
          (���˷��ü�¼_Id.Nextval, 2, v_Billno, 1, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�, 0, 2, r_Clinic.����id, ��ҳid_In,
           סԺ��_In, r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, v_����, Decode(n_����id, Null, Null, 0, Null, n_����id),
           r_Clinic.���˿���id, r_Clinic.�ѱ�, r_Clinic.�շ����, r_Clinic.�շ�ϸĿid, r_Clinic.���㵥λ, r_Clinic.������Ŀ��, r_Clinic.���մ���id,
           r_Clinic.���ձ���, r_Clinic.��������, r_Clinic.��ҩ����, r_Clinic.����, r_Clinic.����, r_Clinic.�Ӱ��־, r_Clinic.���ӱ�־,
           r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ, r_Clinic.��׼����, r_Clinic.Ӧ�ս��, r_Clinic.ʵ�ս��, r_Clinic.ͳ����, 1,
           r_Clinic.��������id, r_Clinic.������, ��Ժʱ��_In, �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.ִ����, r_Clinic.������, n_����Ա���,
           v_����Ա����, r_Clinic.���ʵ�id, '�������ת��', r_Clinic.�Ƿ���, n_��id, n_ҽ��С��id, r_Clinic.ִ��״̬, r_Clinic.ִ��ʱ��);
      
        If Nvl(��������_In, 0) = 1 Then
          Insert Into ������ü�¼
            (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ,
             ������Ŀ��, ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ��������id,
             ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �Ƿ���, �ɿ���id, ����id, ���ʽ��, ִ��״̬, ����״̬, ִ��ʱ��)
          Values
            (���˷��ü�¼_Id.Nextval, r_Clinic.��¼����, No_In, r_Clinic.ʵ��Ʊ��, 2, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�, 1,
             r_Clinic.����id, סԺ��_In, r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, r_Clinic.���˿���id, r_Clinic.�ѱ�, r_Clinic.�շ����,
             r_Clinic.�շ�ϸĿid, r_Clinic.���㵥λ, r_Clinic.������Ŀ��, r_Clinic.���մ���id, r_Clinic.���ձ���, r_Clinic.��������,
             r_Clinic.��ҩ����, r_Clinic.����, -1 * r_Clinic.����, r_Clinic.�Ӱ��־, r_Clinic.���ӱ�־, r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ,
             r_Clinic.��׼����, -1 * r_Clinic.Ӧ�ս��, -1 * r_Clinic.ʵ�ս��, -1 * r_Clinic.ͳ����, 0, r_Clinic.��������id, r_Clinic.������,
             r_Clinic.����ʱ��, �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.������, ����Ա���_In, ����Ա����_In, r_Clinic.���ʵ�id, '',
             r_Clinic.�Ƿ���, n_��id, n_����id, -1 * r_Clinic.ʵ�ս��, -1, 1, r_Clinic.ִ��ʱ��);
        End If;
      End Loop;
    
      --8-�����ѣ�9-����
      --�������
      Select Nvl(Sum(ʵ�ս��), 0) Into n_ʵ�պϼ� From סԺ���ü�¼ Where NO = v_Billno And ��¼���� = 2;
      Update �������
      Set ������� = Nvl(�������, 0) + n_ʵ�պϼ�
      Where ����id = n_����id And ���� = 1 And ���� = 2
      Returning ������� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (n_����id, 1, 2, n_ʵ�պϼ�, 0);
        n_����ֵ := n_ʵ�պϼ�;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete ������� Where ���� = 1 And ����id = n_����id And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --����δ�����
      For r_Fee In (Select ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Nvl(Sum(ʵ�ս��), 0) ʵ�պϼ�
                    From סԺ���ü�¼
                    Where NO = v_Billno And ��¼���� = 2
                    Group By ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid) Loop
        Update ����δ�����
        Set ��� = Nvl(���, 0) + r_Fee.ʵ�պϼ�
        Where ����id = n_����id And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0) And Nvl(���˲���id, 0) = Nvl(r_Fee.���˲���id, 0) And
              Nvl(���˿���id, 0) = Nvl(r_Fee.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Fee.��������id, 0) And
              Nvl(ִ�в���id, 0) = Nvl(r_Fee.ִ�в���id, 0) And ������Ŀid + 0 = r_Fee.������Ŀid And ��Դ;�� + 0 = 2;
        If Sql%RowCount = 0 Then
          Insert Into ����δ�����
            (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
          Values
            (n_����id, ��ҳid_In, r_Fee.���˲���id, r_Fee.���˿���id, r_Fee.��������id, r_Fee.ִ�в���id, r_Fee.������Ŀid, 2, r_Fee.ʵ�պϼ�);
        End If;
      End Loop;
    
      --6.ҩƷ������ݴ���
      For r_Fee In (Select a.Id, a.���, b.����id, b.��������
                    From סԺ���ü�¼ A, �������� B
                    Where a.�շ�ϸĿid = b.����id(+) And a.No = v_Billno And a.��¼���� = 2 And a.��¼״̬ In (1, 3)) Loop
        Update ҩƷ�շ���¼
        Set ���� = Decode(����, 8, 9, 24, 25, ����), ����id = r_Fee.Id, NO = v_Billno
        Where NO = No_In And ���� In (8, 9, 24, 25) And
              ����id In
              (Select ID
               From ������ü�¼
               Where NO = No_In And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ In (1, 3) And ��� = r_Fee.���);
        If Nvl(r_Fee.����id, 0) <> 0 And Nvl(r_Fee.��������, 0) = 1 Then
          --���±�������
          Update ҩƷ�շ���¼
          Set ����id = r_Fee.Id
          Where ���� = 21 And
                ����id In
                (Select ID
                 From ������ü�¼
                 Where NO = No_In And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ In (1, 3) And ��� = r_Fee.���);
        End If;
        --���·�����˼�¼
        Update ������˼�¼
        Set ת��id = r_Fee.Id, ��¼״̬ = Decode(��������_In, 1, 2, 1), ��ҳid = ��ҳid_In, ת���� = ����Ա����_In, ת��ʱ�� = �˷�ʱ��_In
        Where ����id In
              (Select ID
               From ������ü�¼
               Where NO = No_In And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ In (1, 3) And ��� = r_Fee.���) And ���� = 1;
        If Sql%NotFound Then
          --δ�ҵ�����ʱ��Ҫǿ�ƽ��ж�Ӧ.
          Insert Into ������˼�¼
            (����, ����id, ����id, ��ҳid, �����, �������, ��¼״̬, ת��id, ת����, ת��ʱ��)
            Select 1, ID, n_����id, ��ҳid_In, ����Ա����_In, �˷�ʱ��_In, Decode(��������_In, 1, 2, 1), r_Fee.Id, ����Ա����_In, �˷�ʱ��_In
            From ������ü�¼
            Where NO = No_In And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ In (1, 3) And ��� = r_Fee.���;
        End If;
      End Loop;
      Update δ��ҩƷ��¼
      Set ���� = Decode(����, 8, 9, 24, 25, ����), ��ҳid = ��ҳid_In, NO = v_Billno
      Where NO = No_In And ���� In (8, 24) And ����id = n_����id;
    Else
      --���˰��յ���NO���д���
      v_Billno     := Nextno(14);
      n_ҽ��С��id := Zl_ҽ��С��_Get(n_��������id, v_������, n_����id, ��ҳid_In, ��Ժʱ��_In);
    
      Insert Into סԺ���ü�¼
        (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ����, ���˲���id, ���˿���id, �ѱ�, �շ����,
         �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����,
         ���ʷ���, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �Ƿ���, ҽ�����, �ɿ���id, ҽ��С��id)
        Select ���˷��ü�¼_Id.Nextval, 2, v_Billno, 1, ���, ��������, �۸񸸺�, 0, 2, ����id, ��ҳid_In, סԺ��_In, ����, �Ա�, ����, v_����,
               Decode(n_����id, Null, Null, 0, Null, n_����id), ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������,
               ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, 1, ��������id, ������, ��Ժʱ��_In, �˷�ʱ��_In,
               ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ������, ����Ա���, ����Ա����, ���ʵ�id, '�������ת��', �Ƿ���, ҽ�����, n_��id, n_ҽ��С��id
        From ������ü�¼
        Where NO = No_In And ��¼���� = ����_In And ��¼״̬ = 1 And Nvl(���ӱ�־, 0) Not In (8, 9);
    
      If Nvl(��������_In, 0) = 1 Then
        Update ������ü�¼ Set ��¼״̬ = 3 Where NO = No_In And ��¼���� In (1, 2) And ��¼״̬ = 1;
      End If;
    
      For r_Clinic In (Select ��¼����, ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���,
                              ��������, ��ҩ����, ����, Sum(����) As ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Sum(Ӧ�ս��) As Ӧ�ս��,
                              Sum(ʵ�ս��) As ʵ�ս��, Sum(ͳ����) As ͳ����, ��������id, ������, ִ�в���id, Max(ִ����) As ִ����, ������,
                              Max(���ʵ�id) As ���ʵ�id, ����ʱ��, ʵ��Ʊ��, Max(ִ��״̬) As ִ��״̬, Max(ִ��ʱ��) As ִ��ʱ��

                       
                       From ������ü�¼
                       Where NO = No_In And ��¼���� = ����_In And ��¼״̬ In (2, 3) And Nvl(���ӱ�־, 0) Not In (8, 9)
                       Group By ��¼����, ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id,
                                ���ձ���, ��������, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, ��������id, ������, ִ�в���id, ������, ����ʱ��,
                                ʵ��Ʊ��
                       Having Sum(����) <> 0) Loop
        Select ����Ա���, ����Ա����
        Into n_����Ա���, v_����Ա����
        From ������ü�¼
        Where NO = No_In And ��¼���� = ����_In And ��¼״̬ = 3 And Rownum < 2;
        Insert Into סԺ���ü�¼
          (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ����, ���˲���id, ���˿���id, �ѱ�, �շ����,
           �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���,
           ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �ɿ���id, ҽ��С��id, ִ��״̬, ִ��ʱ��)
        Values
          (���˷��ü�¼_Id.Nextval, 2, v_Billno, 1, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�, 0, 2, r_Clinic.����id, ��ҳid_In,
           סԺ��_In, r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, v_����, Decode(n_����id, Null, Null, 0, Null, n_����id),
           r_Clinic.���˿���id, r_Clinic.�ѱ�, r_Clinic.�շ����, r_Clinic.�շ�ϸĿid, r_Clinic.���㵥λ, r_Clinic.������Ŀ��, r_Clinic.���մ���id,
           r_Clinic.���ձ���, r_Clinic.��������, r_Clinic.��ҩ����, r_Clinic.����, r_Clinic.����, r_Clinic.�Ӱ��־, r_Clinic.���ӱ�־,
           r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ, r_Clinic.��׼����, r_Clinic.Ӧ�ս��, r_Clinic.ʵ�ս��, r_Clinic.ͳ����, 1,
           r_Clinic.��������id, r_Clinic.������, ��Ժʱ��_In, �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.ִ����, r_Clinic.������, n_����Ա���,
           v_����Ա����, r_Clinic.���ʵ�id, '�������ת��', n_��id, n_ҽ��С��id, r_Clinic.ִ��״̬, r_Clinic.ִ��ʱ��);
        If Nvl(��������_In, 0) = 1 Then
          Insert Into ������ü�¼
            (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ,
             ������Ŀ��, ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ��������id,
             ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �ɿ���id, ����id, ���ʽ��, ����״̬, ִ��ʱ��)
          Values
            (���˷��ü�¼_Id.Nextval, r_Clinic.��¼����, No_In, r_Clinic.ʵ��Ʊ��, 2, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�, 1,
             r_Clinic.����id, סԺ��_In, r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, r_Clinic.���˿���id, r_Clinic.�ѱ�, r_Clinic.�շ����,
             r_Clinic.�շ�ϸĿid, r_Clinic.���㵥λ, r_Clinic.������Ŀ��, r_Clinic.���մ���id, r_Clinic.���ձ���, r_Clinic.��������,
             r_Clinic.��ҩ����, r_Clinic.����, -1 * r_Clinic.����, r_Clinic.�Ӱ��־, r_Clinic.���ӱ�־, r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ,
             r_Clinic.��׼����, -1 * r_Clinic.Ӧ�ս��, -1 * r_Clinic.ʵ�ս��, -1 * r_Clinic.ͳ����, 1, r_Clinic.��������id, r_Clinic.������,
             r_Clinic.����ʱ��, �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.������, ����Ա���_In, ����Ա����_In, r_Clinic.���ʵ�id, '', n_��id,
             n_����id, -1 * r_Clinic.ʵ�ս��, 1, r_Clinic.ִ��ʱ��);
        End If;
      End Loop;
    
      --8-�����ѣ�9-����
      --�������
      Select Nvl(Sum(ʵ�ս��), 0) Into n_ʵ�պϼ� From סԺ���ü�¼ Where NO = v_Billno And ��¼���� = 2;
      Update �������
      Set ������� = Nvl(�������, 0) + n_ʵ�պϼ�
      Where ����id = n_����id And ���� = 1 And ���� = 2
      Returning ������� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (n_����id, 1, 2, n_ʵ�պϼ�, 0);
        n_����ֵ := n_ʵ�պϼ�;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete ������� Where ���� = 1 And ����id = n_����id And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --����δ�����
      For r_Fee In (Select ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Nvl(Sum(ʵ�ս��), 0) ʵ�պϼ�
                    From סԺ���ü�¼
                    Where NO = v_Billno And ��¼���� = 2
                    Group By ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid) Loop
        Update ����δ�����
        Set ��� = Nvl(���, 0) + r_Fee.ʵ�պϼ�
        Where ����id = n_����id And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0) And Nvl(���˲���id, 0) = Nvl(r_Fee.���˲���id, 0) And
              Nvl(���˿���id, 0) = Nvl(r_Fee.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Fee.��������id, 0) And
              Nvl(ִ�в���id, 0) = Nvl(r_Fee.ִ�в���id, 0) And ������Ŀid + 0 = r_Fee.������Ŀid And ��Դ;�� + 0 = 2;
        If Sql%RowCount = 0 Then
          Insert Into ����δ�����
            (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
          Values
            (n_����id, ��ҳid_In, r_Fee.���˲���id, r_Fee.���˿���id, r_Fee.��������id, r_Fee.ִ�в���id, r_Fee.������Ŀid, 2, r_Fee.ʵ�պϼ�);
        End If;
      End Loop;
    
      --6.ҩƷ������ݴ���
      For r_Fee In (Select a.Id, a.���, b.����id, b.��������
                    From סԺ���ü�¼ A, �������� B
                    Where a.�շ�ϸĿid = b.����id(+) And a.No = v_Billno And a.��¼���� = 2 And a.��¼״̬ In (1, 3)) Loop
        Update ҩƷ�շ���¼
        Set ���� = Decode(����, 8, 9, 24, 25, ����), ����id = r_Fee.Id, NO = v_Billno
        Where NO = No_In And ���� In (8, 9, 24, 25) And
              ����id In (Select ID
                       From ������ü�¼
                       Where NO = No_In And ��¼���� = ����_In And ��¼״̬ In (1, 3) And ��� = r_Fee.���);
        If Nvl(r_Fee.����id, 0) <> 0 And Nvl(r_Fee.��������, 0) = 1 Then
          --���±�������
          Update ҩƷ�շ���¼
          Set ����id = r_Fee.Id
          Where ���� = 21 And
                ����id In (Select ID
                         From ������ü�¼
                         Where NO = No_In And ��¼���� = ����_In And ��¼״̬ In (1, 3) And ��� = r_Fee.���);
        End If;
        --���·�����˼�¼
        Update ������˼�¼
        Set ת��id = r_Fee.Id, ��¼״̬ = Decode(��������_In, 1, 2, 1), ��ҳid = ��ҳid_In, ת���� = ����Ա����_In, ת��ʱ�� = �˷�ʱ��_In
        Where ����id In (Select ID
                       From ������ü�¼
                       Where NO = No_In And ��¼���� = ����_In And ��¼״̬ In (1, 3) And ��� = r_Fee.���) And ���� = 1;
        If Sql%NotFound Then
          --δ�ҵ�����ʱ��Ҫǿ�ƽ��ж�Ӧ.
          Insert Into ������˼�¼
            (����, ����id, ����id, ��ҳid, �����, �������, ��¼״̬, ת��id, ת����, ת��ʱ��)
            Select 1, ID, n_����id, ��ҳid_In, ����Ա����_In, �˷�ʱ��_In, Decode(��������_In, 1, 2, 1), r_Fee.Id, ����Ա����_In, �˷�ʱ��_In
            From ������ü�¼
            Where NO = No_In And ��¼���� = ����_In And ��¼״̬ In (1, 3) And ��� = r_Fee.���;
        End If;
      End Loop;
      Update δ��ҩƷ��¼
      Set ��ҳid = ��ҳid_In, NO = v_Billno
      Where NO = No_In And ���� In (9, 25) And ����id = n_����id;
    End If;
  Else
    If Nvl(��������_In, 0) = 1 Then
      If Mod(����_In, 10) = 1 Then
        Zl_����תסԺ_�շ�ת��(No_In, ����Ա���_In, ����Ա����_In, �˷�ʱ��_In, 0, ��Ժ����id_In, ��ҳid_In, Null, n_����id, ԭ����id_In);
      Else
        --ת���ʵ�
        --No_In;����Ա���_In,����Ա����_In,�˷�ʱ��_In
        Zl_����תסԺ_����ת��(No_In, ����Ա���_In, ����Ա����_In, �˷�ʱ��_In);
      End If;
    End If;
    --����
    -- 1.��Ժ����ID_IN<>NULL ��Ϊ:��Ժ����ID_IN
    -- 2.��ҳid_In<>0 :
    If Nvl(��Ժ����id_In, 0) <> 0 Then
      n_����id := ��Ժ����id_In;
    Elsif Nvl(��ҳid_In, 0) <> 0 Then
      Begin
        Select Nvl(b.��ǰ����id, a.��ǰ����id), Nvl(b.��Ժ����, a.��ǰ����)
        Into n_����id, v_����
        From ������Ϣ A, ������ҳ B
        Where a.����id = b.����id(+) And a.����id = n_����id And b.��ҳid(+) = ��ҳid_In;
      Exception
        When Others Then
          n_����id := Null;
      End;
    End If;
  
    If Nvl(n_����id, 0) = 0 Then
      --����Ժ����Ϊ׼
      n_����id := Nvl(��Ժ����id_In, 0);
    End If;
  
    --��Ժ�Ǽ�֮ǰ,��ҳID��û�в���,����Ԥ����¼,����δ�����,δ��ҩƷ��¼,������ü�¼
    --���в���ID,��ҳID�����,ֻ����Ժ�ǼǺ��ٵ���Zl_�������תסԺ_Update��д
    Select ����id, ��������id, ������
    Into n_����id, n_��������id, v_������
    From ������ü�¼
    Where NO = No_In And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ In (1, 3) And Rownum = 1;
  
    --5.�������ʵ�
    --��Ҫ����Ƿ��Ѿ�ת��
    Select Count(*)
    Into n_Count
    From ������ü�¼ A
    Where NO = No_In And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ In (1, 3) And Exists
     (Select 1 From ������˼�¼ Where a.Id = ����id And ת��id Is Not Null) And Rownum <= 1;
    If n_Count >= 1 Then
      v_Err_Msg := '�����򲢷�ԭ��,�÷����Ѿ�������ת��,���ܼ�������!';
      Raise Err_Item;
    End If;
    If Mod(����_In, 10) = 1 Then
      --�շѰ��ս�����Ų������NO�Ž��д���
      n_ҽ��С��id := Zl_ҽ��С��_Get(n_��������id, v_������, n_����id, ��ҳid_In, ��Ժʱ��_In);
      For r_Nos In (Select Distinct c.No
                    From ������ü�¼ A, ������ü�¼ C
                    Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And a.����id = ԭ����id_In And c.No = a.No And
                          Mod(c.��¼����, 10) = 1 And c.��¼״̬ In (1, 3)) Loop
        v_Billno := Nextno(14);
      
        Insert Into סԺ���ü�¼
          (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ����, ���˲���id, ���˿���id, �ѱ�, �շ����,
           �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����,
           ���ʷ���, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �Ƿ���, ҽ�����, �ɿ���id,
           ҽ��С��id)
          Select ���˷��ü�¼_Id.Nextval, 2, v_Billno, 1, ���, ��������, �۸񸸺�, 0, 2, ����id, ��ҳid_In, סԺ��_In, ����, �Ա�, ����, v_����,
                 Decode(n_����id, Null, Null, 0, Null, n_����id), ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������,
                 ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, 1, ��������id, ������, ��Ժʱ��_In, �˷�ʱ��_In,
                 ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ������, ����Ա���, ����Ա����, ���ʵ�id, '�������ת��', �Ƿ���, ҽ�����, n_��id, n_ҽ��С��id
          From ������ü�¼
          Where NO = r_Nos.No And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ = 1 And Nvl(���ӱ�־, 0) Not In (8, 9);
      
        If Nvl(��������_In, 0) = 1 Then
          Update ������ü�¼ Set ��¼״̬ = 3 Where NO = r_Nos.No And Mod(��¼����, 10) In (1, 2) And ��¼״̬ = 1;
        End If;
      
        For r_Clinic In (Select Min(��¼����) As ��¼����, ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ,
                                ������Ŀ��, ���մ���id, ���ձ���, ��������, ��ҩ����, ����, Sum(����) As ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����,
                                Sum(Ӧ�ս��) As Ӧ�ս��, Sum(ʵ�ս��) As ʵ�ս��, Sum(ͳ����) As ͳ����, ��������id, ������, ִ�в���id,
                                Max(ִ����) As ִ����, ������, Max(���ʵ�id) As ���ʵ�id, Max(�Ƿ���) As �Ƿ���, ����ʱ��, Min(ʵ��Ʊ��) As ʵ��Ʊ��,
                                Max(ִ��״̬) As ִ��״̬, Max(ִ��ʱ��) As ִ��ʱ��
                         From ������ü�¼
                         Where NO = r_Nos.No And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ In (2, 3) And
                               Nvl(���ӱ�־, 0) Not In (8, 9)
                         Group By ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���,
                                  ��������, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, ��������id, ������, ִ�в���id, ������, ����ʱ��
                         Having Sum(����) <> 0) Loop
          Select ����Ա���, ����Ա����
          Into n_����Ա���, v_����Ա����
          From ������ü�¼
          Where NO = r_Nos.No And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ = 3 And Rownum < 2;
          Insert Into סԺ���ü�¼
            (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ����, ���˲���id, ���˿���id, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����,
             ���ʷ���, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �Ƿ���, �ɿ���id, ҽ��С��id, ִ��״̬, ִ��ʱ��)
          Values
            (���˷��ü�¼_Id.Nextval, 2, v_Billno, 1, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�, 0, 2, r_Clinic.����id, ��ҳid_In,
             סԺ��_In, r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, v_����, Decode(n_����id, Null, Null, 0, Null, n_����id),
             r_Clinic.���˿���id, r_Clinic.�ѱ�, r_Clinic.�շ����, r_Clinic.�շ�ϸĿid, r_Clinic.���㵥λ, r_Clinic.������Ŀ��,
             r_Clinic.���մ���id, r_Clinic.���ձ���, r_Clinic.��������, r_Clinic.��ҩ����, r_Clinic.����, r_Clinic.����, r_Clinic.�Ӱ��־,
             r_Clinic.���ӱ�־, r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ, r_Clinic.��׼����, r_Clinic.Ӧ�ս��, r_Clinic.ʵ�ս��, r_Clinic.ͳ����,
             1, r_Clinic.��������id, r_Clinic.������, ��Ժʱ��_In, �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.ִ����, r_Clinic.������, n_����Ա���,
             v_����Ա����, r_Clinic.���ʵ�id, '�������ת��', r_Clinic.�Ƿ���, n_��id, n_ҽ��С��id, r_Clinic.ִ��״̬, r_Clinic.ִ��ʱ��);
        
          If Nvl(��������_In, 0) = 1 Then
            Insert Into ������ü�¼
              (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ,
               ������Ŀ��, ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ��������id,
               ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �Ƿ���, �ɿ���id, ����id, ���ʽ��, ִ��״̬, ����״̬, ִ��ʱ��)
            Values
              (���˷��ü�¼_Id.Nextval, r_Clinic.��¼����, r_Nos.No, r_Clinic.ʵ��Ʊ��, 2, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�,
               1, r_Clinic.����id, סԺ��_In, r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, r_Clinic.���˿���id, r_Clinic.�ѱ�,
               r_Clinic.�շ����, r_Clinic.�շ�ϸĿid, r_Clinic.���㵥λ, r_Clinic.������Ŀ��, r_Clinic.���մ���id, r_Clinic.���ձ���,
               r_Clinic.��������, r_Clinic.��ҩ����, r_Clinic.����, -1 * r_Clinic.����, r_Clinic.�Ӱ��־, r_Clinic.���ӱ�־,
               r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ, r_Clinic.��׼����, -1 * r_Clinic.Ӧ�ս��, -1 * r_Clinic.ʵ�ս��, -1 * r_Clinic.ͳ����,
               0, r_Clinic.��������id, r_Clinic.������, r_Clinic.����ʱ��, �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.������, ����Ա���_In,
               ����Ա����_In, r_Clinic.���ʵ�id, '', r_Clinic.�Ƿ���, n_��id, n_����id, -1 * r_Clinic.ʵ�ս��, -1, 1, r_Clinic.ִ��ʱ��);
          End If;
        End Loop;
      
        --8-�����ѣ�9-����
        --�������
        Select Nvl(Sum(ʵ�ս��), 0) Into n_ʵ�պϼ� From סԺ���ü�¼ Where NO = v_Billno And ��¼���� = 2;
        Update �������
        Set ������� = Nvl(�������, 0) + n_ʵ�պϼ�
        Where ����id = n_����id And ���� = 1 And ���� = 2
        Returning ������� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (n_����id, 1, 2, n_ʵ�պϼ�, 0);
          n_����ֵ := n_ʵ�պϼ�;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete ������� Where ���� = 1 And ����id = n_����id And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
        End If;
      
        --����δ�����
        For r_Fee In (Select ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Nvl(Sum(ʵ�ս��), 0) ʵ�պϼ�
                      From סԺ���ü�¼
                      Where NO = v_Billno And ��¼���� = 2
                      Group By ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid) Loop
          Update ����δ�����
          Set ��� = Nvl(���, 0) + r_Fee.ʵ�պϼ�
          Where ����id = n_����id And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0) And Nvl(���˲���id, 0) = Nvl(r_Fee.���˲���id, 0) And
                Nvl(���˿���id, 0) = Nvl(r_Fee.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Fee.��������id, 0) And
                Nvl(ִ�в���id, 0) = Nvl(r_Fee.ִ�в���id, 0) And ������Ŀid + 0 = r_Fee.������Ŀid And ��Դ;�� + 0 = 2;
          If Sql%RowCount = 0 Then
            Insert Into ����δ�����
              (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
            Values
              (n_����id, ��ҳid_In, r_Fee.���˲���id, r_Fee.���˿���id, r_Fee.��������id, r_Fee.ִ�в���id, r_Fee.������Ŀid, 2, r_Fee.ʵ�պϼ�);
          End If;
        End Loop;
      
        --6.ҩƷ������ݴ���
        For r_Fee In (Select a.Id, a.���, b.����id, b.��������
                      From סԺ���ü�¼ A, �������� B
                      Where a.�շ�ϸĿid = b.����id(+) And a.No = v_Billno And a.��¼���� = 2 And a.��¼״̬ In (1, 3)) Loop
          Update ҩƷ�շ���¼
          Set ���� = Decode(����, 8, 9, 24, 25, ����), ����id = r_Fee.Id, NO = v_Billno
          Where NO = r_Nos.No And ���� In (8, 9, 24, 25) And
                ����id In
                (Select ID
                 From ������ü�¼
                 Where NO = r_Nos.No And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ In (1, 3) And ��� = r_Fee.���);
          If Nvl(r_Fee.����id, 0) <> 0 And Nvl(r_Fee.��������, 0) = 1 Then
            --���±�������
            Update ҩƷ�շ���¼
            Set ����id = r_Fee.Id
            Where ���� = 21 And ����id In (Select ID
                                       From ������ü�¼
                                       Where NO = r_Nos.No And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ In (1, 3) And
                                             ��� = r_Fee.���);
          End If;
          --���·�����˼�¼
          Update ������˼�¼
          Set ת��id = r_Fee.Id, ��¼״̬ = Decode(��������_In, 1, 2, 1), ��ҳid = ��ҳid_In, ת���� = ����Ա����_In, ת��ʱ�� = �˷�ʱ��_In
          Where ����id In
                (Select ID
                 From ������ü�¼
                 Where NO = r_Nos.No And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ In (1, 3) And ��� = r_Fee.���) And
                ���� = 1;
          If Sql%NotFound Then
            --δ�ҵ�����ʱ��Ҫǿ�ƽ��ж�Ӧ.
            Insert Into ������˼�¼
              (����, ����id, ����id, ��ҳid, �����, �������, ��¼״̬, ת��id, ת����, ת��ʱ��)
              Select 1, ID, n_����id, ��ҳid_In, ����Ա����_In, �˷�ʱ��_In, Decode(��������_In, 1, 2, 1), r_Fee.Id, ����Ա����_In, �˷�ʱ��_In
              From ������ü�¼
              Where NO = r_Nos.No And Mod(��¼����, 10) = Mod(����_In, 10) And ��¼״̬ In (1, 3) And ��� = r_Fee.���;
          End If;
        End Loop;
        Update δ��ҩƷ��¼
        Set ���� = Decode(����, 8, 9, 24, 25, ����), ��ҳid = ��ҳid_In, NO = v_Billno
        Where NO = No_In And ���� In (8, 24) And ����id = n_����id;
      End Loop;
    Else
      --���˰��յ���NO���д���
      v_Billno     := Nextno(14);
      n_ҽ��С��id := Zl_ҽ��С��_Get(n_��������id, v_������, n_����id, ��ҳid_In, ��Ժʱ��_In);
    
      Insert Into סԺ���ü�¼
        (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ����, ���˲���id, ���˿���id, �ѱ�, �շ����,
         �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����,
         ���ʷ���, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �Ƿ���, ҽ�����, �ɿ���id, ҽ��С��id)
        Select ���˷��ü�¼_Id.Nextval, 2, v_Billno, 1, ���, ��������, �۸񸸺�, 0, 2, ����id, ��ҳid_In, סԺ��_In, ����, �Ա�, ����, v_����,
               Decode(n_����id, Null, Null, 0, Null, n_����id), ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������,
               ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, 1, ��������id, ������, ��Ժʱ��_In, �˷�ʱ��_In,
               ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ������, ����Ա���, ����Ա����, ���ʵ�id, '�������ת��', �Ƿ���, ҽ�����, n_��id, n_ҽ��С��id
        From ������ü�¼
        Where NO = No_In And ��¼���� = ����_In And ��¼״̬ = 1 And Nvl(���ӱ�־, 0) Not In (8, 9);
    
      If Nvl(��������_In, 0) = 1 Then
        Update ������ü�¼ Set ��¼״̬ = 3 Where NO = No_In And ��¼���� In (1, 2) And ��¼״̬ = 1;
      End If;
    
      For r_Clinic In (Select ��¼����, ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���,
                              ��������, ��ҩ����, ����, Sum(����) As ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Sum(Ӧ�ս��) As Ӧ�ս��,
                              Sum(ʵ�ս��) As ʵ�ս��, Sum(ͳ����) As ͳ����, ��������id, ������, ִ�в���id, Max(ִ����) As ִ����, ������,
                              Max(���ʵ�id) As ���ʵ�id, ����ʱ��, ʵ��Ʊ��, Max(ִ��״̬) As ִ��״̬, Max(ִ��ʱ��) As ִ��ʱ��

                       
                       From ������ü�¼
                       Where NO = No_In And ��¼���� = ����_In And ��¼״̬ In (2, 3) And Nvl(���ӱ�־, 0) Not In (8, 9)
                       Group By ��¼����, ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id,
                                ���ձ���, ��������, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, ��������id, ������, ִ�в���id, ������, ����ʱ��,
                                ʵ��Ʊ��
                       Having Sum(����) <> 0) Loop
        Select ����Ա���, ����Ա����
        Into n_����Ա���, v_����Ա����
        From ������ü�¼
        Where NO = No_In And ��¼���� = ����_In And ��¼״̬ = 3 And Rownum < 2;
        Insert Into סԺ���ü�¼
          (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ����, ���˲���id, ���˿���id, �ѱ�, �շ����,
           �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���,
           ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �ɿ���id, ҽ��С��id, ִ��״̬, ִ��ʱ��)
        Values
          (���˷��ü�¼_Id.Nextval, 2, v_Billno, 1, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�, 0, 2, r_Clinic.����id, ��ҳid_In,
           סԺ��_In, r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, v_����, Decode(n_����id, Null, Null, 0, Null, n_����id),
           r_Clinic.���˿���id, r_Clinic.�ѱ�, r_Clinic.�շ����, r_Clinic.�շ�ϸĿid, r_Clinic.���㵥λ, r_Clinic.������Ŀ��, r_Clinic.���մ���id,
           r_Clinic.���ձ���, r_Clinic.��������, r_Clinic.��ҩ����, r_Clinic.����, r_Clinic.����, r_Clinic.�Ӱ��־, r_Clinic.���ӱ�־,
           r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ, r_Clinic.��׼����, r_Clinic.Ӧ�ս��, r_Clinic.ʵ�ս��, r_Clinic.ͳ����, 1,
           r_Clinic.��������id, r_Clinic.������, ��Ժʱ��_In, �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.ִ����, r_Clinic.������, n_����Ա���,
           v_����Ա����, r_Clinic.���ʵ�id, '�������ת��', n_��id, n_ҽ��С��id, r_Clinic.ִ��״̬, r_Clinic.ִ��ʱ��);
        If Nvl(��������_In, 0) = 1 Then
          Insert Into ������ü�¼
            (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ,
             ������Ŀ��, ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ��������id,
             ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �ɿ���id, ����id, ���ʽ��, ����״̬, ִ��ʱ��)
          Values
            (���˷��ü�¼_Id.Nextval, r_Clinic.��¼����, No_In, r_Clinic.ʵ��Ʊ��, 2, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�, 1,
             r_Clinic.����id, סԺ��_In, r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, r_Clinic.���˿���id, r_Clinic.�ѱ�, r_Clinic.�շ����,
             r_Clinic.�շ�ϸĿid, r_Clinic.���㵥λ, r_Clinic.������Ŀ��, r_Clinic.���մ���id, r_Clinic.���ձ���, r_Clinic.��������,
             r_Clinic.��ҩ����, r_Clinic.����, -1 * r_Clinic.����, r_Clinic.�Ӱ��־, r_Clinic.���ӱ�־, r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ,
             r_Clinic.��׼����, -1 * r_Clinic.Ӧ�ս��, -1 * r_Clinic.ʵ�ս��, -1 * r_Clinic.ͳ����, 1, r_Clinic.��������id, r_Clinic.������,
             r_Clinic.����ʱ��, �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.������, ����Ա���_In, ����Ա����_In, r_Clinic.���ʵ�id, '', n_��id,
             n_����id, -1 * r_Clinic.ʵ�ս��, 1, r_Clinic.ִ��ʱ��);
        End If;
      End Loop;
    
      --8-�����ѣ�9-����
      --�������
      Select Nvl(Sum(ʵ�ս��), 0) Into n_ʵ�պϼ� From סԺ���ü�¼ Where NO = v_Billno And ��¼���� = 2;
      Update �������
      Set ������� = Nvl(�������, 0) + n_ʵ�պϼ�
      Where ����id = n_����id And ���� = 1 And ���� = 2
      Returning ������� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (n_����id, 1, 2, n_ʵ�պϼ�, 0);
        n_����ֵ := n_ʵ�պϼ�;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete ������� Where ���� = 1 And ����id = n_����id And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --����δ�����
      For r_Fee In (Select ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Nvl(Sum(ʵ�ս��), 0) ʵ�պϼ�
                    From סԺ���ü�¼
                    Where NO = v_Billno And ��¼���� = 2
                    Group By ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid) Loop
        Update ����δ�����
        Set ��� = Nvl(���, 0) + r_Fee.ʵ�պϼ�
        Where ����id = n_����id And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0) And Nvl(���˲���id, 0) = Nvl(r_Fee.���˲���id, 0) And
              Nvl(���˿���id, 0) = Nvl(r_Fee.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Fee.��������id, 0) And
              Nvl(ִ�в���id, 0) = Nvl(r_Fee.ִ�в���id, 0) And ������Ŀid + 0 = r_Fee.������Ŀid And ��Դ;�� + 0 = 2;
        If Sql%RowCount = 0 Then
          Insert Into ����δ�����
            (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
          Values
            (n_����id, ��ҳid_In, r_Fee.���˲���id, r_Fee.���˿���id, r_Fee.��������id, r_Fee.ִ�в���id, r_Fee.������Ŀid, 2, r_Fee.ʵ�պϼ�);
        End If;
      End Loop;
    
      --6.ҩƷ������ݴ���
      For r_Fee In (Select a.Id, a.���, b.����id, b.��������
                    From סԺ���ü�¼ A, �������� B
                    Where a.�շ�ϸĿid = b.����id(+) And a.No = v_Billno And a.��¼���� = 2 And a.��¼״̬ In (1, 3)) Loop
        Update ҩƷ�շ���¼
        Set ���� = Decode(����, 8, 9, 24, 25, ����), ����id = r_Fee.Id, NO = v_Billno
        Where NO = No_In And ���� In (8, 9, 24, 25) And
              ����id In (Select ID
                       From ������ü�¼
                       Where NO = No_In And ��¼���� = ����_In And ��¼״̬ In (1, 3) And ��� = r_Fee.���);
        If Nvl(r_Fee.����id, 0) <> 0 And Nvl(r_Fee.��������, 0) = 1 Then
          --���±�������
          Update ҩƷ�շ���¼
          Set ����id = r_Fee.Id
          Where ���� = 21 And
                ����id In (Select ID
                         From ������ü�¼
                         Where NO = No_In And ��¼���� = ����_In And ��¼״̬ In (1, 3) And ��� = r_Fee.���);
        End If;
        --���·�����˼�¼
        Update ������˼�¼
        Set ת��id = r_Fee.Id, ��¼״̬ = Decode(��������_In, 1, 2, 1), ��ҳid = ��ҳid_In, ת���� = ����Ա����_In, ת��ʱ�� = �˷�ʱ��_In
        Where ����id In (Select ID
                       From ������ü�¼
                       Where NO = No_In And ��¼���� = ����_In And ��¼״̬ In (1, 3) And ��� = r_Fee.���) And ���� = 1;
        If Sql%NotFound Then
          --δ�ҵ�����ʱ��Ҫǿ�ƽ��ж�Ӧ.
          Insert Into ������˼�¼
            (����, ����id, ����id, ��ҳid, �����, �������, ��¼״̬, ת��id, ת����, ת��ʱ��)
            Select 1, ID, n_����id, ��ҳid_In, ����Ա����_In, �˷�ʱ��_In, Decode(��������_In, 1, 2, 1), r_Fee.Id, ����Ա����_In, �˷�ʱ��_In
            From ������ü�¼
            Where NO = No_In And ��¼���� = ����_In And ��¼״̬ In (1, 3) And ��� = r_Fee.���;
        End If;
      End Loop;
      Update δ��ҩƷ��¼
      Set ��ҳid = ��ҳid_In, NO = v_Billno
      Where NO = No_In And ���� In (9, 25) And ����id = n_����id;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������תסԺ_Insert;
/

--133498:Ƚ����,2018-11-07,֧�ֲ���������������תΪסԺ����
Create Or Replace Procedure Zl_����תסԺ_������ת��
(
  No_In         ���ò����¼.No%Type,
  ���ó���id_In ����Ԥ����¼.����id%Type,
  �������id_In ����Ԥ����¼.����id%Type,
  �������_In   ����Ԥ����¼.�������%Type,
  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
  ��Ժ����id_In ����Ԥ����¼.����id%Type,
  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
  ����_In     ����Ԥ����¼.��Ԥ��%Type := Null
) As
  --���ܣ��Է��ò�������������ý���תסԺ���ô���
  --��Σ�
  --  ���㷽ʽ_In ��Ϊ�գ���ʾ���г�Ԥ����ķ�ҽ�����ȫ����Ϊָ���Ľ��㷽ʽ��
  --              Ϊ�գ���ʾ���г�Ԥ����ķ�ҽ�����ȫ��תΪסԺԤ����
  Err_Item Exception;
  v_Err_Msg Varchar2(200);
  n_����ֵ  ����Ԥ����¼.��Ԥ��%Type;

  n_��id   ����ɿ����.Id%Type;
  v_���� ���㷽ʽ.����%Type;
  n_���� ����Ԥ����¼.��Ԥ��%Type;
  n_Dec    Number; --���С��λ�� 

  v_Nos    Varchar2(4000);
  n_����id ����Ԥ����¼.����id%Type;

  n_���˽�� ����Ԥ����¼.��Ԥ��%Type;
  n_δ�˽�� ����Ԥ����¼.��Ԥ��%Type;
  n_��Ԥ��   ����Ԥ����¼.��Ԥ��%Type;
  v_���㷽ʽ Varchar2(4000);
  v_Ԥ��no   ����Ԥ����¼.No%Type;

  --����Ԥ�����
  Procedure ����Ԥ����¼_Insert
  (
    ����id_In     ����Ԥ����¼.����id%Type,
    ���_In       ����Ԥ����¼.���%Type,
    ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
    �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
    �������_In   ����Ԥ����¼.�������%Type,
    �����id_In   ����Ԥ����¼.�����id%Type := Null,
    ����_In       ����Ԥ����¼.����%Type := Null,
    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    ����˵��_In   ����Ԥ����¼.����˵��%Type := Null
  ) As
    v_Ԥ��no ����Ԥ����¼.No%Type;
    n_����ֵ ����Ԥ����¼.���%Type;
  Begin
    If Nvl(���_In, 0) = 0 Or ���㷽ʽ_In Is Null Then
      Return;
    End If;
  
    --һ��ͨ��ÿһ�ʶ�����һ��Ԥ�����¼
    --������ͬһ�ֽ��㷽ʽֻ����һ��Ԥ�����¼
    Update ����Ԥ����¼
    Set ��� = Nvl(���, 0) + ���_In
    Where ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� = �տ�ʱ��_In And ����id + 0 = ����id_In And ���㷽ʽ = ���㷽ʽ_In And Nvl(�����id, 0) = 0;
    If Sql%RowCount = 0 Or Nvl(�����id_In, 0) <> 0 Then
      v_Ԥ��no := Nextno(11);
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id, Ԥ�����,
         �����id, ����, ����˵��, ������ˮ��, �������)
      Values
        (����Ԥ����¼_Id.Nextval, v_Ԥ��no, Null, 1, 1, ����id_In, ��ҳid_In, ��Ժ����id_In, ���_In, ���㷽ʽ_In, �տ�ʱ��_In, Null, Null, Null,
         ����Ա���_In, ����Ա����_In, '����תסԺԤ��', n_��id, 2, �����id_In, ����_In, ����˵��_In, ������ˮ��_In, �������_In);
    End If;
  
    Update �������
    Set Ԥ����� = Nvl(Ԥ�����, 0) + ���_In
    Where ���� = 1 And ����id = ����id_In And ���� = 2
    Returning Ԥ����� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, ����, Ԥ�����, �������) Values (����id_In, 1, 2, ���_In, 0);
      n_����ֵ := ���_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
    End If;
  End;
Begin
  n_��id := Zl_Get��id(����Ա����_In);
  --����
  Begin
    Select ���� Into v_���� From ���㷽ʽ Where ���� = 9 And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := 'û�з��������㷽ʽ�������Ƿ���ȷ���ã�';
      Raise Err_Item;
  End;
  n_���� := Nvl(����_In, 0);

  --���С��λ�� 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  Select f_List2str(Cast(Collect(a.No) As t_Strlist), ',', 1), Max(a.����id)
  Into v_Nos, n_����id
  From ������ü�¼ A, ���ò����¼ B
  Where a.����id = b.�շѽ���id And b.��¼���� = 1 And b.���ӱ�־ = 0 And b.No = No_In;
  If v_Nos Is Null Then
    v_Err_Msg := 'δ�ҵ�ԭҽ�����������ݣ�����ת��ʧ��!';
    Raise Err_Item;
  End If;

  --1.���·�����˼�¼ 
  Update ������˼�¼
  Set ��¼״̬ = 2
  Where ���� = 1 And ����id In (Select /*+cardinality(b,10)*/
                             a.Id
                            From ������ü�¼ A, (Select Column_Value As NO From Table(f_Str2list(v_Nos))) B
                            Where a.No = b.No And Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3));

  --2.����������ü�¼ 
  Update ������ü�¼
  Set ��¼״̬ = 3
  Where Mod(��¼����, 10) = 1 And ��¼״̬ = 1 And NO In (Select Column_Value As NO From Table(f_Str2list(v_Nos)));

  For c_���� In (Select /*+cardinality(b,10)*/
                a.No, a.���, a.��������, a.�۸񸸺�, a.����id, a.ҽ�����, a.�����־, a.����, a.�Ա�, a.����, a.��ʶ��, a.���ʽ, a.���˿���id, a.�ѱ�,
                a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.��ҩ����, Sum(Nvl(a.����, 1) * a.����) As ����, a.�Ӱ��־, a.���ӱ�־, a.Ӥ����, a.������Ŀid, a.�վݷ�Ŀ,
                a.��׼����, Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��, a.������, a.��������id, a.������, a.����ʱ��, a.ִ�в���id, a.ִ����,
                Min(Decode(a.��¼״̬, 2, a.ִ��״̬, 0)) - 1 As ִ��״̬, a.����, Sum(a.���ʽ��) As ���ʽ��, Max(���մ���id) As ���մ���id,
                Max(������Ŀ��) As ������Ŀ��, Max(���ձ���) As ���ձ���, Max(��������) As ��������, Sum(a.ͳ����) As ͳ����, Max(�Ƿ��ϴ�) As �Ƿ��ϴ�, �Ƿ���,
                a.�Һ�id, a.��ҳid
               From ������ü�¼ A, (Select Column_Value As NO From Table(f_Str2list(v_Nos))) B
               Where a.No = b.No And a.��¼���� In (1, 11)
               Group By a.No, a.���, a.��������, a.�۸񸸺�, a.����id, a.ҽ�����, a.�����־, a.����, a.�Ա�, a.����, a.��ʶ��, a.���ʽ, a.���˿���id,
                        a.�ѱ�, a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.��ҩ����, a.�Ӱ��־, a.���ӱ�־, a.Ӥ����, a.������Ŀid, a.�վݷ�Ŀ, a.��׼����, a.������,
                        a.��������id, a.������, a.����ʱ��, a.ִ�в���id, a.ִ����, a.����, �Ƿ���, a.�Һ�id, a.��ҳid
               Having Nvl(Sum(Nvl(a.����, 1) * a.����), 0) <> 0) Loop
  
    Insert Into ������ü�¼
      (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��,
       ����, ����Ա���, ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, �ɿ���id, ����״̬, �Һ�id, ��ҳid)
    Values
      (���˷��ü�¼_Id.Nextval, 1, c_����.No, 2, c_����.���, c_����.��������, c_����.�۸񸸺�, c_����.����id, c_����.ҽ�����, c_����.�����־, c_����.����,
       c_����.�Ա�, c_����.����, c_����.��ʶ��, c_����.���ʽ, c_����.���˿���id, c_����.�ѱ�, c_����.�շ����, c_����.�շ�ϸĿid, c_����.���㵥λ, 1, c_����.��ҩ����,
       -1 * c_����.����, c_����.�Ӱ��־, c_����.���ӱ�־, c_����.Ӥ����, c_����.������Ŀid, c_����.�վݷ�Ŀ, c_����.��׼����, -1 * c_����.Ӧ�ս��, -1 * c_����.ʵ�ս��,
       c_����.������, c_����.��������id, c_����.������, c_����.����ʱ��, �˷�ʱ��_In, c_����.ִ�в���id, c_����.ִ����, c_����.ִ��״̬, Null, c_����.����, ����Ա���_In,
       ����Ա����_In, ���ó���id_In, -1 * c_����.���ʽ��, c_����.���մ���id, c_����.������Ŀ��, c_����.���ձ���, c_����.��������, -1 * c_����.ͳ����, c_����.�Ƿ��ϴ�, '',
       c_����.�Ƿ���, n_��id, 0, c_����.�Һ�id, c_����.��ҳid);
  End Loop;
  Zl_�����˷ѽ���_Modify(1, n_����id, ���ó���id_In, Null);

  --3.���ϲ�������¼��ͬʱ�ѽ�����Ʊ�ݻ��պ�ҽ��ԭ���ˣ�
  Zl_���ò����¼_Delete(No_In, �������id_In, Null, �������_In, ���ó���id_In, ����Ա���_In, ����Ա����_In, �˷�ʱ��_In);
  Update ���ò����¼ Set ����״̬ = 0 Where ������� = �������_In;
  --����Ϊҽ���ӿ��ѵ��óɹ�
  Update ����Ԥ����¼
  Set У�Ա�־ = 2
  Where ��¼���� = 6 And ����id = �������id_In And ���㷽ʽ In (Select ���� From ���㷽ʽ Where ���� In (3, 4));

  --4.�������ݴ���
  Select -1 * Nvl(Sum(a.��Ԥ��), 0)
  Into n_δ�˽��
  From ����Ԥ����¼ A
  Where a.������� = �������_In And a.���㷽ʽ Is Null;
  If Nvl(n_����, 0) = 0 Then
    n_���� := Round(n_δ�˽��, n_Dec) - n_δ�˽��;
  End If;
  n_δ�˽�� := n_δ�˽�� - n_����;

  For r_Ԥ�� In (Select Case
                        When Mod(a.��¼����, 10) = 1 Then
                         1
                        When Nvl(a.�����id, 0) <> 0 Then
                         2
                        Else
                         0
                      End As ����, a.����id, Nvl(a.��Ԥ��, 0) As ��Ԥ��, a.No, a.����id, a.���㷽ʽ, a.�����id, a.����, a.������ˮ��, a.����˵��,
                      a.�������
               From ����Ԥ����¼ A, ���㷽ʽ B
               Where a.���㷽ʽ = b.���� And a.��¼״̬ In (1, 3) And b.���� Not In (3, 4, 9) And
                     a.����id In (Select �շѽ���id From ���ò����¼ Where ��¼���� = 1 And ���ӱ�־ = 0 And NO = No_In)) Loop
  
    --���ǵ��ֽ��㷽ʽ
    If r_Ԥ��.���� = 1 Then
      --Ԥ����
      Zl_���ò������_����˷�(�������id_In, Null, Null, Null, Null, Null, n_����, 0, 0, -1 * n_δ�˽��);
      Exit;
    Elsif r_Ԥ��.���� = 2 Then
      --һ��ͨ
      Select Nvl(Sum(���), 0) Into n_���˽�� From �����˿���Ϣ Where ��¼id = r_Ԥ��.����id;
      If r_Ԥ��.��Ԥ�� - n_���˽�� > 0 Then
        If r_Ԥ��.��Ԥ�� - n_���˽�� > n_δ�˽�� Then
          n_��Ԥ�� := n_δ�˽��;
        Else
          n_��Ԥ�� := r_Ԥ��.��Ԥ�� - n_���˽��;
        End If;
      
        v_���㷽ʽ := r_Ԥ��.���㷽ʽ || '|' || -1 * n_��Ԥ�� || '| | ';
        Zl_���ò������_����˷�(�������id_In, v_���㷽ʽ, r_Ԥ��.�����id, r_Ԥ��.����, r_Ԥ��.������ˮ��, r_Ԥ��.����˵��, n_����, 0, 1);
        Zl_�����˿���Ϣ_Insert(�������_In, r_Ԥ��.����id, n_��Ԥ��, r_Ԥ��.����, r_Ԥ��.������ˮ��, r_Ԥ��.����˵��);
      
        --תΪסԺԤ����
        ����Ԥ����¼_Insert(r_Ԥ��.����id, n_��Ԥ��, r_Ԥ��.���㷽ʽ, �˷�ʱ��_In, r_Ԥ��.�������, r_Ԥ��.�����id, r_Ԥ��.����, r_Ԥ��.������ˮ��, r_Ԥ��.����˵��);
      
        n_δ�˽�� := n_δ�˽�� - n_��Ԥ��;
        n_����   := 0;
      End If;
      If n_δ�˽�� = 0 Then
        Exit;
      End If;
    Else
      --������ҽ�����㷽ʽ
      --���㷽ʽ|������|�������|����ժҪ
      v_���㷽ʽ := r_Ԥ��.���㷽ʽ || '|' || n_δ�˽�� || '| | ';
      Zl_���ò������_����˷�(�������id_In, v_���㷽ʽ, Null, Null, Null, Null, n_����, 0);
    
      --תΪסԺԤ����
      ����Ԥ����¼_Insert(r_Ԥ��.����id, n_δ�˽��, r_Ԥ��.���㷽ʽ, �˷�ʱ��_In, r_Ԥ��.�������);
      Exit;
    End If;
  End Loop;

  --5.ת����ɴ���   
  Delete From ����Ԥ����¼ Where ����id = �������id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '������δ�ɿ�����ݣ�������ɽ��㣡';
    Raise Err_Item;
  End If;
  Delete From ����Ԥ����¼ Where ����id = ���ó���id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '������δ�ɿ�����ݣ�������ɽ��㣡';
    Raise Err_Item;
  End If;
  Update ����Ԥ����¼ Set У�Ա�־ = 0, �Ự�� = Null Where ������� = �������_In;

  --��Ա�ɿ�����Ҫ��ҽ����
  For c_Ԥ�� In (Select a.���㷽ʽ, a.����Ա����, Nvl(Sum(a.��Ԥ��), 0) As ��Ԥ��
               From ����Ԥ����¼ A, ���㷽ʽ B
               Where a.���㷽ʽ = b.���� And b.���� In (3, 4) And a.������� = �������_In
               Group By a.���㷽ʽ, a.����Ա����) Loop
  
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + c_Ԥ��.��Ԥ��
    Where �տ�Ա = c_Ԥ��.����Ա���� And ���� = 1 And ���㷽ʽ = c_Ԥ��.���㷽ʽ
    Returning ��� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (c_Ԥ��.����Ա����, c_Ԥ��.���㷽ʽ, 1, c_Ԥ��.��Ԥ��);
      n_����ֵ := c_Ԥ��.��Ԥ��;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = c_Ԥ��.����Ա���� And ���� = 1 And ���㷽ʽ = c_Ԥ��.���㷽ʽ And Nvl(���, 0) = 0;
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����תסԺ_������ת��;
/





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0036' Where ���=&n_System;
Commit;
