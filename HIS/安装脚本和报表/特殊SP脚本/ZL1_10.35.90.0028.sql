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
--129954:��͢��,2018-09-07,������������ҽ������վΣ��ֵ��������
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1260, 1, 1, 0, 0, 0, 0, 39, '����Σ��ֵ��������', '1', '1', '��������Σ��ֵ�����Ƿ񵯴�',
         '0-��������Σ��ֵ���Ѳ�������1-��������Σ��ֵ��������', '', '�������û���Ҫ��������Σ��ֵ��������', Null
  From Dual;

Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1261, 1, 1, 0, 0, 0, 0, 55, 'סԺΣ��ֵ��������', '1', '1', '����סԺΣ��ֵ�����Ƿ񵯴�',
         '0-����סԺΣ��ֵ���Ѳ�������1-����סԺΣ��ֵ��������', '', '�������û���Ҫ����סԺΣ��ֵ��������', Null
  From Dual;


--126863:����,2018-09-06,�������ģ����������������Դ
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1122, 1, 0, 0, 0, 0, 0, 84, '������Դ', Null, '1',
         ' ���Ƶ�ǰ������ȱʡ�����ﲡ�˻�סԺ���������Ҳ�����Ϣ�Ϳ��ƿ������Ҽ�ִ�п��ҵ�:' || Chr(13) || '1) ��Ҫ�������ط����ò���:һ�ǲ���������,�����ڼ��ʴ��ڵ�״̬����.' || Chr(13) ||
          '2)������ͨ������"1. ����ID,2.סԺ��,3. ���￨��,4.�����,5.ҽ����,6.���֤��,7.IC����"ʱ,�����Զ��л����ò���������Դ,����:��ǰ��������Ժ����,�����ǰ���õ������ﲡ��,�����Զ��л���סԺ����״̬,�������סԺ���˺�,���Զ��л����˲������õ�״̬' ||
          Chr(13) || '3)���ݲ�����Դ��ȷ��"��������"�������շ���Ŀ��ִ�п���,��������ԴΪ�����,�򿪵����һ�ִ�п���ֻ���Ƿ�����������ܷ����������סԺ�Ŀ���',
         '' || Chr(13) || '1-���ﲡ��,2-סԺ����', Null, '��������Ҫ���������ʱ���ݲ�����Դ�����п��Ƶ��û�.', Null
  From Dual;

-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--126863:����,2018-09-06,�������ģ����������������Դ
Create Or Replace Procedure Zl_������ʼ�¼_Insert
(
  No_In         ������ü�¼.No%Type,
  ���_In       ������ü�¼.���%Type,
  ����id_In     ������ü�¼.����id%Type,
  ��ʶ��_In     ������ü�¼.��ʶ��%Type,
  ����_In       ������ü�¼.����%Type,
  �Ա�_In       ������ü�¼.�Ա�%Type,
  ����_In       ������ü�¼.����%Type,
  �ѱ�_In       ������ü�¼.�ѱ�%Type,
  �Ӱ��־_In   ������ü�¼.�Ӱ��־%Type,
  Ӥ����_In     ������ü�¼.Ӥ����%Type,
  ���˿���id_In ������ü�¼.���˿���id%Type,
  ��������id_In ������ü�¼.��������id%Type,
  ������_In     ������ü�¼.������%Type,
  ��������_In   ������ü�¼.��������%Type,
  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
  �շ����_In   ������ü�¼.�շ����%Type,
  ���㵥λ_In   ������ü�¼.���㵥λ%Type,
  ����_In       ������ü�¼.����%Type,
  ����_In       ������ü�¼.����%Type,
  ���ӱ�־_In   ������ü�¼.���ӱ�־%Type,
  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
  �۸񸸺�_In   ������ü�¼.�۸񸸺�%Type,
  ������Ŀid_In ������ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In   ������ü�¼.�վݷ�Ŀ%Type,
  ��׼����_In   ������ü�¼.��׼����%Type,
  Ӧ�ս��_In   ������ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In   ������ü�¼.ʵ�ս��%Type,
  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type,
  ҩƷժҪ_In   ҩƷ�շ���¼.ժҪ%Type,
  ����_In       Number,
  ����Ա���_In ������ü�¼.����Ա���%Type,
  ����Ա����_In ������ü�¼.����Ա����%Type,
  ���ʵ�id_In   ������ü�¼.���ʵ�id%Type := Null,
  ����ժҪ_In   ������ü�¼.ժҪ%Type := Null,
  ҽ�����_In   ������ü�¼.ҽ�����%Type := Null,
  Ƶ��_In       ҩƷ�շ���¼.Ƶ��%Type := Null,
  ����_In       ҩƷ�շ���¼.����%Type := Null,
  �÷�_In       ҩƷ�շ���¼.�÷�%Type := Null, --�÷�[|�巨]
  ��Ч_In       ҩƷ�շ���¼.����%Type := Null,
  �Ƽ�����_In   ҩƷ�շ���¼.����%Type := Null,
  �����־_In   ������ü�¼.�����־%Type := 1,
  ��ҩ��̬_In   ������ü�¼.����%Type := Null,
  ��������_In   Number := 0,
  ����_In       ҩƷ�շ���¼.����%Type := Null
) As
  --���ܣ�����һ��������ʵ���
  --������
  --   ҩƷժҪ_IN:�޸ı����µ���ʱ�á�Ŀǰ�����ڴ����ҩƷ�շ���¼��ժҪ�С�
  --         ԭ����(��¼״̬=2)��¼�޸Ĳ������µ��ݺš�
  --         �µ���(��¼״̬=1)��¼���޸ĵ�ԭ���ݺš�
  v_����id ������ü�¼.Id%Type;
  n_����   ���˹Һż�¼.����%Type;

  --��ʱ����
  v_�÷�     ҩƷ�շ���¼.�÷�%Type;
  v_�巨     ҩƷ�շ���¼.���%Type;
  n_����С�� Number;
  n_�Һ�id   ���˹Һż�¼.Id%Type;
  n_���۴��� ������ҳ.��ҳid%Type;

  n_Dec     Number;
  n_Count   Number;
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_��ҩ���� ҩƷ�շ���¼.��ҩ����%Type;
  n_�������� ��������.��������%Type;

Begin
  n_�������� := 0;
  If �շ����_In = '4' Then
    --�������õ����ĲŴ���
    Select Nvl(��������, 0) Into n_�������� From �������� Where ����id = �շ�ϸĿid_In;
  End If;

  --���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))
  Into n_Dec, n_����С��
  From Dual;

  Select Max(��ҳid) Into n_���۴��� From ������ҳ Where ����id = ����id_In And �������� = 1 And ��Ժ���� Is Null;

  If (�շ����_In In ('5', '6', '7') Or �շ����_In = '4' And n_�������� = 1) And Nvl(����_In, 0) = 0 Then
    --ͬһ�ŵ���,����ͬһҩ��ͬһ����
    Begin
      Select ��ҩ����
      Into v_��ҩ����
      From ������ü�¼
      Where �շ���� In ('5', '6', '7', '4') And NO = No_In And ��¼���� = 2 And ִ�в���id = ִ�в���id_In And ��ҩ���� Is Not Null And
            Rownum <= 1;
    Exception
      When Others Then
        v_��ҩ���� := Null;
    End;
    If v_��ҩ���� Is Null Then
      --ͬһ��������ͨ�ŹҺ���Ч�Һ���������δ��ҩ�����ϰ��,�����һ�μ��˴���Ϊ׼
      n_Count := To_Number(Substr(Nvl(zl_GetSysParameter(21), '11') || '11', 1, 1));
      If n_Count = 0 Then
        n_Count := 1;
      End If;
    
      Begin
        Select ��ҩ����
        Into v_��ҩ����
        From (Select �Ǽ�ʱ��, ��ҩ����
               From ������ü�¼ A
               Where �շ���� In ('5', '6', '7', '4') And ����id = ����id_In And �Ǽ�ʱ�� Between Sysdate - n_Count And Sysdate And
                     ��¼���� = 2 And ִ�в���id = ִ�в���id_In And ��ҩ���� Is Not Null And Exists
                (Select 1
                      From δ��ҩƷ��¼
                      Where a.No = NO And ���� In (9, 26) And �ⷿid + 0 = ִ�в���id_In And ����id + 0 = ����id_In) And Exists
                (Select 1
                      From ��ҩ����
                      Where Nvl(�ϰ��, 0) = 1 And ���� = a.��ҩ���� And Nvl(ר��, 0) = 0 And ҩ��id = ִ�в���id_In)
               Order By �Ǽ�ʱ�� Desc)
        Where Rownum <= 1;
      
      Exception
        When Others Then
          v_��ҩ���� := Null;
      End;
      If v_��ҩ���� Is Null Then
        v_��ҩ���� := Zl_Get��ҩ����(ִ�в���id_In);
      End If;
    End If;
  End If;
  --������ü�¼
  Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;

  --�Ƿ��Ǽ���Һŵ�
  If Nvl(ҽ�����_In, 0) <> 0 Then
    Begin
      Select Nvl(Max(����), 0), Max(ID)
      Into n_����, n_�Һ�id
      From ���˹Һż�¼
      Where NO In (Select �Һŵ� From ����ҽ����¼ Where ID = Nvl(ҽ�����_In, 0)) And ����id = ����id_In;
    Exception
      When Others Then
        n_����   := Null;
        n_�Һ�id := Null;
    End;
  End If;

  Insert Into ������ü�¼
    (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ����, �Ӱ��־,
     ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʷ���, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ����Ա���, ����Ա����, Ӥ����, ���ʵ�id,
     ժҪ, ҽ�����, ����, ��ҩ����, �Ƿ���, ��ҳid, �Һ�id)
  Values
    (v_����id, 2, No_In, Decode(����_In, 1, 0, 1), ���_In, Decode(��������_In, 0, Null, ��������_In),
     Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), �����־_In, ����id_In, Decode(��ʶ��_In, 0, Null, ��ʶ��_In), ����_In, �Ա�_In, ����_In,
     ���˿���id_In, �ѱ�_In, �շ����_In, �շ�ϸĿid_In, ���㵥λ_In, ����_In, ����_In, �Ӱ��־_In, ���ӱ�־_In, ������Ŀid_In, �վݷ�Ŀ_In, ��׼����_In, Ӧ�ս��_In,
     ʵ�ս��_In, 1, ����Ա����_In, ��������id_In, ������_In, ����ʱ��_In, �Ǽ�ʱ��_In, ִ�в���id_In, 0, Decode(����_In, 1, Null, ����Ա���_In),
     Decode(����_In, 1, Null, ����Ա����_In), Ӥ����_In, ���ʵ�id_In, ����ժҪ_In, ҽ�����_In, ��ҩ��̬_In, v_��ҩ����, Nvl(n_����, 0), n_���۴���, n_�Һ�id);

  --��ػ��ܱ�Ĵ���
  If Nvl(����_In, 0) = 0 Then
    --�������
    If Nvl(�����־_In, 0) <> 4 Then
      Update �������
      Set ������� = Nvl(�������, 0) + ʵ�ս��_In
      Where ����id = ����id_In And ���� = 1 And ���� = Decode(�����־_In, 2, 2, 1);
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (����id_In, 1, Decode(�����־_In, 2, 2, 1), ʵ�ս��_In, 0);
      End If;
    End If;
  
    --����δ�����
    Update ����δ�����
    Set ��� = Nvl(���, 0) + ʵ�ս��_In
    Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(���˿���id_In, 0) And
          Nvl(��������id, 0) = Nvl(��������id_In, 0) And Nvl(ִ�в���id, 0) = Nvl(ִ�в���id_In, 0) And ������Ŀid + 0 = ������Ŀid_In And
          ��Դ;�� + 0 = �����־_In;
  
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (����id_In, Null, Null, ���˿���id_In, ��������id_In, ִ�в���id_In, ������Ŀid_In, �����־_In, ʵ�ս��_In);
    End If;
  
  End If;

  --ҩƷ���������ϲ���
  If �շ����_In In ('4', '5', '6', '7') Then
    --ҩƷ�÷��巨�ֽ�
    If �÷�_In Is Not Null Then
      If Instr(�÷�_In, '|') > 0 Then
        v_�÷� := Substr(�÷�_In, 1, Instr(�÷�_In, '|') - 1);
        v_�巨 := Substr(�÷�_In, Instr(�÷�_In, '|') + 1);
      Else
        v_�÷� := �÷�_In;
      End If;
    End If;
    Zl_ҩƷ�շ���¼_���۳���(v_����id, ҩƷժҪ_In, Ƶ��_In, ����_In, v_�÷�, v_�巨, ��Ч_In, �Ƽ�����_In, Null, ��������_In, ����_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������ʼ�¼_Insert;
/

--130471:��ΰ��,2018-09-04,����ԤԼ

Create Or Replace Procedure Zl_Third_Outpatireg
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
    Select Count(1) Into n_Count From ������ҳ B Where b.�Һ�id = n_�Һ�id;
    If n_Count > 0 Then
      Select b.����id Into n_����id From ������ҳ B Where b.�Һ�id = n_�Һ�id;
      Zl_��Ժ������ҳ_Delete(n_����id, 0);
    End If;
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



------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0028' Where ���=&n_System;
Commit;