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
--129571:����,2018-08-08,�޸�Oracle����Zl_���������Һ�_Delete,��ȡ��ȷ����ĿID������
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
  
    Select a.����ʱ��, a.�Ǽ�ʱ��, b.�շ�ϸĿid As ��Ŀid, a.ִ�в���id As ����id, a.ִ���� As ҽ������, c.Id As ҽ��id, a.�ű� As ����
    From ���˹Һż�¼ A, ������ü�¼ B, ��Ա�� C
    Where a.��¼���� = Decode(v_��Ч����, 0, v_����, a.��¼����) And b.��¼���� = 4 And a.��¼״̬ = v_״̬ And a.No = ���ݺ�_In And a.No = b.No And
          a.ִ���� = c.����(+) And b.��� = 1 And Rownum = 1;

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
        --Ҫɾ������
        For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
          n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, Nvl(v_�Һ�.ִ�в���id, 0)));
          If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
            Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
          End If;
        End Loop;
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
        --Ҫɾ������
        For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
          n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(v_�Һ�.ִ�в���id, 0)));
          If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
            Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
          End If;
        End Loop;
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

--128435:����,2018-08-08,�޸�Oracle����Zl_Third_Saveexes,������νڵ�ҽ��ID�ͳ��νڵ㵥�ݺ�
Create Or Replace Procedure Zl_Third_Saveexes
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:������ü�¼
  --���:Xml_In:
  --<IN>
  --  <PATIID></PATIID>  //����ID
  --  <PAGEID></PAGEID>  //��ҳID
  --  <MZBZ></MZBZ>   //�����־��1-���2-סԺ
  --  <JZBZ></JZBZ>   //���ʱ�־��0-�շѣ�1-����
  --  <CZY></CZY>   //����Ա
  --  <CZSJ></CZSJ>   //����ʱ��
  --  <KDR></KDR>  //������
  --  <KDKSID></KDKSID>  //��������ID
  --  <YQBH></YQBH>  //Ժ�����
  --  <MXLIST>
  --    <MX>
  --      <YZID><YZID> //ҽ��ID
  --      <SFXMID></SFXMID>  //�շ�ϸĿID
  --      <SL></SL>   //����
  --      <ZXR></ZXR>  //ִ����,��ʾ��ȫִ��
  --      <ZXKSID></ZXKSID>  //ִ�п���ID
  --    </MX>
  --    ...
  --  </MXLIST>
  --</IN>
  --����:Xml_Out
  --<OUTPUT>
  --   <RESULT></RESULT> //true-��ʾ����ɹ�;false-��ʾ����ʧ��
  --   <DJH></DJH>  //���ݺ�
  --   <ERROR>      //ʧ��ʱ����
  --     <MSG></MSG>   //��ϸ������ʾ
  --   </ERROR>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  n_����id     ������Ϣ.����id%Type;
  n_��ҳid     ������ҳ.��ҳid%Type;
  n_�����־   ������ü�¼.�����־%Type; --1-���2-סԺ
  n_���ʱ�־   Number(2); --0-�շѣ�1-����
  v_����Ա���� ������ü�¼.����Ա����%Type;
  d_����ʱ��   ������ü�¼.�Ǽ�ʱ��%Type;
  v_������     ������ü�¼.������%Type;
  n_��������id ������ü�¼.��������id%Type;
  v_վ��       ���ű�.վ��%Type;
  Xml_��ϸ�б� Xmltype;

  v_No           ������ü�¼.No%Type;
  n_��ʶ��       ������ü�¼.��ʶ��%Type;
  v_����         ������ü�¼.����%Type;
  v_�Ա�         ������ü�¼.�Ա�%Type;
  v_����         ������ü�¼.����%Type;
  v_�ѱ�         ������ü�¼.�ѱ�%Type;
  v_���ʽ���� ������ü�¼.���ʽ%Type;
  v_���ʽ���� ������Ϣ.ҽ�Ƹ��ʽ%Type;
  n_����id       סԺ���ü�¼.���˲���id%Type;
  n_����id       ������ü�¼.���˿���id%Type;
  v_����         סԺ���ü�¼.����%Type;
  v_����Ա���   ������ü�¼.����Ա���%Type;
  d_��Ժ����     ������ҳ.��Ժ����%Type;
  d_��Ժ����     ������ҳ.��Ժ����%Type;

  Type Ty_Rec_Bill Is Record(
    ���       ������ü�¼.���%Type,
    �۸񸸺�   ������ü�¼.�۸񸸺�%Type,
    �շ�ϸĿid ������ü�¼.�շ�ϸĿid%Type,
    �շ����   ������ü�¼.�շ����%Type,
    ���㵥λ   ������ü�¼.���㵥λ%Type,
    ������Ŀid ������ü�¼.������Ŀid%Type,
    �վݷ�Ŀ   ������ü�¼.�վݷ�Ŀ%Type,
    ����       ������ü�¼.����%Type,
    ��׼����   ������ü�¼.��׼����%Type,
    Ӧ�ս��   ������ü�¼.Ӧ�ս��%Type,
    ʵ�ս��   ������ü�¼.ʵ�ս��%Type,
    ִ����     ������ü�¼.ִ����%Type,
    ִ�в���id ������ü�¼.ִ�в���id%Type,
    ����ժҪ   ������ü�¼.ժҪ%Type,
    ҽ��id     ������ü�¼.ҽ�����%Type);

  Type Ty_Tb_Bill Is Table Of Ty_Rec_Bill;
  c_Bill Ty_Tb_Bill := Ty_Tb_Bill();

  n_����С�� Number;
  n_���С�� Number;

  v_Temp     Varchar2(4000);
  v_�۸�ȼ� �շѼ�Ŀ.�۸�ȼ�%Type;
  v_��ͨ�ȼ� �շѼ�Ŀ.�۸�ȼ�%Type;
  v_ҩƷ�ȼ� �շѼ�Ŀ.�۸�ȼ�%Type;
  v_���ĵȼ� �շѼ�Ŀ.�۸�ȼ�%Type;

  n_���         ������ü�¼.���%Type;
  n_�۸񸸺�     ������ü�¼.�۸񸸺�%Type;
  n_��ǰ�۸񸸺� ������ü�¼.�۸񸸺�%Type;
  n_�۸�         ������ü�¼.��׼����%Type;
  n_ʣ����       ������ü�¼.����%Type;
  n_ʵ�ս��     ������ü�¼.ʵ�ս��%Type;
  d_�Ǽ�ʱ��     ������ü�¼.�Ǽ�ʱ��%Type;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  Procedure Zl_Third_Checkdata
  (
    ������Դ_In   In Number,
    ����id_In     ������ü�¼.����id%Type,
    ��ҳid_In     ������ҳ.��ҳid%Type,
    ����id_In     ������ҳ.��ǰ����id%Type,
    �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
    ����_In       ������ü�¼.����%Type,
    ʵ�ս��_In   ������ü�¼.ʵ�ս��%Type,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type,
    �Ƿ����_In   Number := 0
  ) Is
  
    --��Σ�
    --        ������Դ_In  1-����/2-סԺ
    --        �Ƿ����_In �Ƿ���˷���:0-�շ�/1-����
    n_�������� ��������.��������%Type;
    n_���÷��� ��������.���÷���%Type;
    n_�Ƿ��� �շ���ĿĿ¼.�Ƿ���%Type;
    n_���     ҩƷ���.��������%Type;
    n_��Ŀ���� �շ���ĿĿ¼.����%Type;
  
    n_��鷽ʽ  ���ϳ�����.��鷽ʽ%Type;
    v_�շ����  ������ü�¼.�շ����%Type;
    n_��������  ���ʱ�����.��������%Type;
    n_����ֵ    ���ʱ�����.����ֵ%Type;
    v_������־2 ���ʱ�����.������־2%Type;
    v_������־3 ���ʱ�����.������־3%Type;
    n_��־      Number;
    v_�������  �շ���Ŀ���.����%Type;
    n_�������  ������ü�¼.ʵ�ս��%Type;
    v_����      Varchar2(100);
    n_ʣ����  ������ü�¼.ʵ�ս��%Type;
    n_���ս��  ������ü�¼.ʵ�ս��%Type;
  
    v_Err_Msg Varchar2(255);
    Err_Item Exception;
  Begin
  
    --1.ҩƷ/���Ŀ���飬Ҫ�ֿ�
    Begin
      Select b.�Ƿ���, b.����, b.���, a.��������, Decode(b.���, '4', a.���÷���, c.ҩ������)
      Into n_�Ƿ���, n_��Ŀ����, v_�շ����, n_��������, n_���÷���
      From �������� A, �շ���ĿĿ¼ B, ҩƷ��� C
      Where b.Id = a.����id(+) And b.Id = c.ҩƷid(+) And b.Id = �շ�ϸĿid_In;
    Exception
      When Others Then
        v_Err_Msg := 'δ�����շ���Ŀ��';
        Raise Err_Item;
    End;
  
    If Instr('5,6,7', v_�շ����) > 0 Or v_�շ���� = '4' And Nvl(n_��������, 0) = 1 Then
      Select Nvl(Sum(a.��������), 0)
      Into n_���
      From ҩƷ��� A
      Where a.���� = 1 And a.�ⷿid = ִ�в���id_In And (Nvl(a.����, 0) = 0 Or a.Ч�� Is Null Or a.Ч�� > Trunc(Sysdate)) And
            a.ҩƷid = �շ�ϸĿid_In;
      If n_��� < ����_In Then
        If Nvl(n_���÷���, 0) = 1 Or Nvl(n_�Ƿ���, 0) = 1 Then
          v_Err_Msg := '[' || n_��Ŀ���� || ']�ĵ�ǰ���ÿ�治������������';
          Raise Err_Item;
        Else
          Begin
            If Instr('5,6,7', v_�շ����) > 0 Then
              Select a.��鷽ʽ Into n_��鷽ʽ From ҩƷ������ A Where a.�ⷿid = ִ�в���id_In;
            Else
              Select a.��鷽ʽ Into n_��鷽ʽ From ���ϳ����� A Where a.�ⷿid = ִ�в���id_In;
            End If;
          Exception
            When Others Then
              n_��鷽ʽ := 0;
          End;
          If Nvl(n_��鷽ʽ, 0) = 2 Then
            v_Err_Msg := '[' || n_��Ŀ���� || ']�ĵ�ǰ���ÿ�治������������';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End If;
  
    --2.סԺ���˺��������Ҫ���м��˷��౨��
    If Nvl(�Ƿ����_In, 0) = 1 Then
      --���ʷ��౨��
      Begin
        Select Nvl(��������, 1) As ��������, ����ֵ, ������־2, ������־3
        Into n_��������, n_����ֵ, v_������־2, v_������־3
        From ���ʱ�����
        Where ���ò��� = Zl_Patiwarnscheme(����id_In, ��ҳid_In) And ����id = ����id_In;
      Exception
        When Others Then
          n_�������� := 0;
      End;
    
      If n_����ֵ Is Not Null And Nvl(n_��������, 0) > 0 Then
        If v_������־2 Is Not Null Then
          If v_������־2 = '-' Or Instr(v_������־2, v_�շ����) > 0 Then
            n_��־ := 2;
          End If;
          If v_������־2 = '-' Then
            v_������� := ''; --�������ʱ,������ʾ��������
          End If;
        End If;
        If Nvl(n_��־, 0) = 0 And v_������־3 Is Not Null Then
          If v_������־3 = '-' Or Instr(v_������־3, v_�շ����) > 0 Then
            n_��־ := 3;
          End If;
          If v_������־3 = '-' Then
            v_������� := ''; --�������ʱ,������ʾ��������
          End If;
        End If;
      
        If n_�������� = 1 Then
          --�ۼƷ��ñ���(����)\
          n_������� := Zl_Patientsurety(����id_In, ��ҳid_In);
          If n_������� > 0 Then
            v_���� := '(�������' || n_������� || ')';
          End If;
        
          Select Nvl(Sum(Ԥ����� - �������), 0)
          Into n_ʣ����
          From �������
          Where ���� = 1 And ���� = Decode(������Դ_In, 1, 2, 1) And ����id = ����id_In;
        
          n_ʣ���� := n_ʣ���� + n_������� - Nvl(ʵ�ս��_In, 0);
          If n_��־ = 2 Then
            --Ԥ����ľ�ʱ��ֹ����
            If n_ʣ���� < 0 Then
              v_Err_Msg := 'ʣ���' || v_���� || '�Ѿ��ľ���' || v_������� || '��ֹ���ʡ�';
              Raise Err_Item;
            End If;
          Elsif n_��־ = 3 Then
            --���ڱ���ֵ��ֹ����
            If n_ʣ���� < n_����ֵ Then
              v_Err_Msg := 'ʣ���' || v_���� || '����' || v_������� || '����ֵ��' || n_����ֵ || '����ֹ���ʡ�';
              Raise Err_Item;
            End If;
          End If;
        Elsif n_�������� = 2 Then
          --ÿ�շ��ñ���(����)
          If n_��־ = 3 Then
            --���ڱ���ֵ��ֹ����
            n_���ս�� := Zl_Patidaycharge(����id_In);
            n_���ս�� := n_���ս�� + Nvl(ʵ�ս��_In, 0);
            If n_���ս�� > n_����ֵ Then
              v_Err_Msg := '���շ��ã�' || n_���ս�� || '������' || v_������� || '����ֵ��' || n_����ֵ || '����ֹ���ʡ�';
              Raise Err_Item;
            End If;
          End If;
        End If;
      End If;
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Zl_Third_Checkdata;

Begin
  --��ȡ���
  Select Extractvalue(Value(A), 'IN/PATIID'),
         Decode(Extractvalue(Value(A), 'IN/PAGEID'), 0, Null, Extractvalue(Value(A), 'IN/PAGEID')),
         Nvl(Extractvalue(Value(A), 'IN/MZBZ'), 0), Nvl(Extractvalue(Value(A), 'IN/JZBZ'), 0),
         Extractvalue(Value(A), 'IN/CZY'), To_Date(Extractvalue(Value(A), 'IN/CZSJ'), 'yyyy-mm-dd hh24:mi:ss'),
         Extractvalue(Value(A), 'IN/KDR'),
         Decode(Extractvalue(Value(A), 'IN/KDKSID'), 0, Null, Extractvalue(Value(A), 'IN/KDKSID')),
         Extractvalue(Value(A), 'IN/YQBH'), Extract(Value(A), 'IN/MXLIST')
  Into n_����id, n_��ҳid, n_�����־, n_���ʱ�־, v_����Ա����, d_����ʱ��, v_������, n_��������id, v_վ��, Xml_��ϸ�б�
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Begin
    Select a.��� Into v_����Ա��� From ��Ա�� A Where a.���� = v_����Ա����;
  Exception
    When Others Then
      v_Err_Msg := 'δ���ҵ�����Ա��Ϣ�����ݱ���ʧ�ܣ�';
      Raise Err_Item;
  End;

  Begin
    If n_�����־ = 2 Then
      Select a.����, a.�Ա�, a.����, a.�ѱ�, a.סԺ��, a.��ǰ����id, Nvl(a.��Ժ����id, n_��������id), a.��Ժ����, c.����, c.����, a.��Ժ����, a.��Ժ����,
             a.��ǰ����id
      Into v_����, v_�Ա�, v_����, v_�ѱ�, n_��ʶ��, n_����id, n_����id, v_����, v_���ʽ����, v_���ʽ����, d_��Ժ����, d_��Ժ����, n_����id
      From ������ҳ A, ������Ϣ B, ҽ�Ƹ��ʽ C
      Where a.����id = b.����id And b.����id = n_����id And a.��ҳid = n_��ҳid And a.ҽ�Ƹ��ʽ = c.����(+);
    Else
      Select a.����, a.�Ա�, a.����, a.�ѱ�, a.�����, n_��������id, b.����, b.����
      Into v_����, v_�Ա�, v_����, v_�ѱ�, n_��ʶ��, n_����id, v_���ʽ����, v_���ʽ����
      From ������Ϣ A, ҽ�Ƹ��ʽ B
      Where a.����id = n_����id And a.ҽ�Ƹ��ʽ = b.����(+);
    End If;
  Exception
    When Others Then
      v_Err_Msg := 'δ���ҵ�������Ϣ�����ݱ���ʧ�ܣ�';
      Raise Err_Item;
  End;

  --סԺ���˷���ʱ��ļ��
  If Nvl(n_�����־, 0) = 2 Then
    If d_��Ժ���� Is Not Null Then
      If d_����ʱ�� > d_��Ժ���� Then
        v_Err_Msg := 'ǿ�ƶԳ�Ժ���˼���ʱ������ʱ�䲻�ܴ��ڲ��˳�Ժʱ��:' || To_Char(d_��Ժ����, 'yyyy-mm-dd hh24:mi:ss') || '��';
        Raise Err_Item;
      End If;
    End If;
    If d_��Ժ���� Is Not Null Then
      If d_����ʱ�� < d_��Ժ���� Then
        v_Err_Msg := '���õķ���ʱ�䲻��С�ڲ��˵���Ժʱ��:' || To_Char(d_��Ժ����, 'yyyy-mm-dd hh24:mi:ss') || '��';
        Raise Err_Item;
      End If;
    End If;
  End If;

  If v_�ѱ� Is Null Then
    Select Max(����) Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
    If v_�ѱ� Is Null Then
      v_Err_Msg := '�޷�ȷ�����˷ѱ�����ȱʡ�ѱ��Ƿ���ȷ���ã�';
      Raise Err_Item;
    End If;
  End If;

  --������С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))
  Into n_���С��, n_����С��
  From Dual;

  --�۸�ȼ�
  v_Temp := Zl_Get_Pricegrade(v_վ��, n_����id, n_��ҳid, v_���ʽ����);
  For c_�۸�ȼ� In (Select Rownum As ���, Column_Value As �۸�ȼ� From Table(f_Str2list(v_Temp, '|'))) Loop
    If c_�۸�ȼ�.��� = 1 Then
      v_��ͨ�ȼ� := c_�۸�ȼ�.�۸�ȼ�;
    End If;
    If c_�۸�ȼ�.��� = 2 Then
      v_ҩƷ�ȼ� := c_�۸�ȼ�.�۸�ȼ�;
    End If;
    If c_�۸�ȼ�.��� = 3 Then
      v_���ĵȼ� := c_�۸�ȼ�.�۸�ȼ�;
    End If;
  End Loop;

  n_��� := 1;
  For c_��ϸ In (Select a.�շ�ϸĿid, a.����, a.ִ����, a.ִ�п���id, a.ժҪ, a.ҽ��id, b.���, b.����, b.���㵥λ, b.�Ƿ���, b.���ηѱ�, b.����ʱ��
               From (Select Extractvalue(Value(J), '/MX/SFXMID') As �շ�ϸĿid, Extractvalue(Value(J), '/MX/SL') As ����,
                             Extractvalue(Value(J), '/MX/ZXR') As ִ����, Extractvalue(Value(J), '/MX/ZXKSID') As ִ�п���id,
                             Extractvalue(Value(J), '/MX/FYZY') As ժҪ, Extractvalue(Value(J), '/MX/YZID') As ҽ��id
                      From Table(Xmlsequence(Extract(Xml_��ϸ�б�, '/MXLIST/MX'))) J) A, �շ���ĿĿ¼ B
               Where a.�շ�ϸĿid = b.Id) Loop
  
    If Nvl(c_��ϸ.����ʱ��, Sysdate + 1) < Sysdate Then
      v_Err_Msg := '��' || c_��ϸ.���� || '����ͣ�ã����ݱ���ʧ�ܣ�';
      Raise Err_Item;
    End If;
  
    n_�۸񸸺� := n_���;
    If c_��ϸ.��� = '4' Then
      v_�۸�ȼ� := v_���ĵȼ�;
    Elsif Instr(',5,6,7,', ',' || c_��ϸ.��� || ',') > 0 Then
      v_�۸�ȼ� := v_ҩƷ�ȼ�;
    Else
      v_�۸�ȼ� := v_��ͨ�ȼ�;
    End If;
    For c_�շѼ�Ŀ In (Select a.������Ŀid, b.�վݷ�Ŀ, a.�ּ�, a.ȱʡ�۸�
                   From �շѼ�Ŀ A, ������Ŀ B
                   Where a.������Ŀid = b.Id And Sysdate Between a.ִ������ And Nvl(a.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) And
                         a.�շ�ϸĿid = c_��ϸ.�շ�ϸĿid And
                         (a.�۸�ȼ� = v_�۸�ȼ� Or
                         (a.�۸�ȼ� Is Null And Not Exists
                          (Select 1
                            From �շѼ�Ŀ
                            Where a.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_�۸�ȼ� And Sysdate Between ִ������ And
                                  Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))) Loop
    
      If n_�۸񸸺� = n_��� Then
        n_��ǰ�۸񸸺� := Null;
      Else
        n_��ǰ�۸񸸺� := n_�۸񸸺�;
      End If;
    
      If Instr(',4,5,6,7,', ',' || c_��ϸ.��� || ',') = 0 Then
        --��ͨ�շ���Ŀ
        If Nvl(c_��ϸ.�Ƿ���, 0) = 0 Then
          n_�۸� := Nvl(c_�շѼ�Ŀ.�ּ�, 0);
        Else
          n_�۸� := Nvl(c_�շѼ�Ŀ.ȱʡ�۸�, 0);
        End If;
      Else
        --ҩƷ����
        v_Temp   := Zl_Get_Retailprice(c_��ϸ.�շ�ϸĿid, v_�۸�ȼ�, c_��ϸ.ִ�п���id, c_��ϸ.����) || '||';
        n_�۸�   := Nvl(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1), 0);
        v_Temp   := Substr(v_Temp, Instr(v_Temp, '|') + 1);
        n_ʣ���� := Substr(v_Temp, 1, Instr(v_Temp, '|') - 1);
      
        If Nvl(n_ʣ����, 0) <> 0 And Nvl(c_��ϸ.�Ƿ���, 0) = 1 Then
          --����δ�ֽ����
          If Instr(',5,6,7,', ',' || c_��ϸ.��� || ',') > 0 Then
            v_Err_Msg := 'ʱ��ҩƷ"' || c_��ϸ.���� || '"��治�㣬�޷�����۸�';
          Else
            v_Err_Msg := 'ʱ����������"' || c_��ϸ.���� || '"��治�㣬�޷�����۸�';
          End If;
          Raise Err_Item;
        End If;
      End If;
    
      c_Bill.Extend;
      c_Bill(c_Bill.Count).��� := n_���;
      c_Bill(c_Bill.Count).�۸񸸺� := n_��ǰ�۸񸸺�;
      c_Bill(c_Bill.Count).�շ�ϸĿid := c_��ϸ.�շ�ϸĿid;
      c_Bill(c_Bill.Count).�շ���� := c_��ϸ.���;
      c_Bill(c_Bill.Count).���㵥λ := c_��ϸ.���㵥λ;
      c_Bill(c_Bill.Count).������Ŀid := c_�շѼ�Ŀ.������Ŀid;
      c_Bill(c_Bill.Count).�վݷ�Ŀ := c_�շѼ�Ŀ.�վݷ�Ŀ;
      c_Bill(c_Bill.Count).���� := c_��ϸ.����;
      c_Bill(c_Bill.Count).��׼���� := Round(n_�۸�, n_����С��);
      c_Bill(c_Bill.Count).Ӧ�ս�� := Round(c_Bill(c_Bill.Count).��׼���� * c_Bill(c_Bill.Count).����, n_���С��);
      If Nvl(c_��ϸ.���ηѱ�, 0) = 1 Or c_Bill(c_Bill.Count).Ӧ�ս�� = 0 Then
        c_Bill(c_Bill.Count).ʵ�ս�� := c_Bill(c_Bill.Count).Ӧ�ս��;
      Else
        v_Temp := Zl_Actualmoney(v_�ѱ�, c_��ϸ.�շ�ϸĿid, c_�շѼ�Ŀ.������Ŀid, c_Bill(c_Bill.Count).Ӧ�ս��, c_��ϸ.����, c_��ϸ.ִ�п���id) || '::';
        v_Temp := Substr(v_Temp, Instr(v_Temp, ':') + 1);
        n_ʵ�ս�� := Substr(v_Temp, 1, Instr(v_Temp, ':') - 1);
        c_Bill(c_Bill.Count).ʵ�ս�� := Round(Nvl(n_ʵ�ս��, 0), n_���С��);
      End If;
      c_Bill(c_Bill.Count).ִ���� := c_��ϸ.ִ����;
      c_Bill(c_Bill.Count).ִ�в���id := c_��ϸ.ִ�п���id;
      c_Bill(c_Bill.Count).����ժҪ := c_��ϸ.ժҪ;
      c_Bill(c_Bill.Count).ҽ��id := c_��ϸ.ҽ��id;
    
      n_��� := n_��� + 1;
      Zl_Third_Checkdata(n_�����־, n_����id, n_��ҳid, n_����id, c_��ϸ.�շ�ϸĿid, c_��ϸ.����, c_Bill(c_Bill.Count).ʵ�ս��, c_��ϸ.ִ�п���id,
                         n_���ʱ�־);
    End Loop;
  End Loop;

  --���ݺ�
  If (n_�����־ = 1 And n_���ʱ�־ = 1) Or n_�����־ = 2 Then
    v_No := Nextno(14);
  Else
    v_No := Nextno(13);
  End If;

  --���浥��
  d_�Ǽ�ʱ�� := Sysdate;
  For I In 1 .. c_Bill.Count Loop
    If n_�����־ = 1 Then
      If n_���ʱ�־ = 0 Then
        --���ﻮ��
        Zl_���ﻮ�ۼ�¼_Insert(v_No, c_Bill(I).���, n_����id, n_��ҳid, n_��ʶ��, v_���ʽ����, v_����, v_�Ա�, v_����, v_�ѱ�, 0, n_����id,
                         n_��������id, v_������, Null, c_Bill(I).�շ�ϸĿid, c_Bill(I).�շ����, c_Bill(I).���㵥λ, Null, 1, c_Bill(I).����,
                         0, c_Bill(I).ִ�в���id, c_Bill(I).�۸񸸺�, c_Bill(I).������Ŀid, c_Bill(I).�վݷ�Ŀ, c_Bill(I).��׼����,
                         c_Bill(I).Ӧ�ս��, c_Bill(I).ʵ�ս��, d_����ʱ��, d_�Ǽ�ʱ��, Null, v_����Ա����, c_Bill(I).����ժҪ, c_Bill(I).ҽ��id);
      
        If c_Bill(I).ִ���� Is Not Null Then
          --���Ϊ��ȫִ��
          Update ������ü�¼
          Set ִ��״̬ = 1, ִ���� = c_Bill(I).ִ����, ִ��ʱ�� = Sysdate
          Where ��¼���� = 1 And NO = v_No And ��� = c_Bill(I).���;
        End If;
      Else
        --�������
        Zl_������ʼ�¼_Insert(v_No, c_Bill(I).���, n_����id, n_��ʶ��, v_����, v_�Ա�, v_����, v_�ѱ�, 0, 0, n_����id, n_��������id, v_������, Null,
                         c_Bill(I).�շ�ϸĿid, c_Bill(I).�շ����, c_Bill(I).���㵥λ, 1, c_Bill(I).����, 0, c_Bill(I).ִ�в���id,
                         c_Bill(I).�۸񸸺�, c_Bill(I).������Ŀid, c_Bill(I).�վݷ�Ŀ, c_Bill(I).��׼����, c_Bill(I).Ӧ�ս��,
                         c_Bill(I).ʵ�ս��, d_����ʱ��, d_�Ǽ�ʱ��, Null, 0, v_����Ա���, v_����Ա����, Null, c_Bill(I).����ժҪ, c_Bill(I).ҽ��id);
      
        If c_Bill(I).ִ���� Is Not Null Then
          --���Ϊ��ȫִ��
          Update ������ü�¼
          Set ִ��״̬ = 1, ִ���� = c_Bill(I).ִ����, ִ��ʱ�� = Sysdate
          Where ��¼���� = 2 And NO = v_No And ��� = c_Bill(I).���;
        End If;
      End If;
    Elsif n_�����־ = 2 Then
      --סԺ����
      Zl_סԺ���ʼ�¼_Insert(v_No, c_Bill(I).���, n_����id, n_��ҳid, n_��ʶ��, v_����, v_�Ա�, v_����, v_����, v_�ѱ�, n_����id, n_����id, 0, 0,
                       n_��������id, v_������, Null, c_Bill(I).�շ�ϸĿid, c_Bill(I).�շ����, c_Bill(I).���㵥λ, 0, Null, Null, 1,
                       c_Bill(I).����, 0, c_Bill(I).ִ�в���id, c_Bill(I).�۸񸸺�, c_Bill(I).������Ŀid, c_Bill(I).�վݷ�Ŀ,
                       c_Bill(I).��׼����, c_Bill(I).Ӧ�ս��, c_Bill(I).ʵ�ս��, Null, d_����ʱ��, d_�Ǽ�ʱ��, Null, 0, v_����Ա���, v_����Ա����,
                       0, Null, Null, c_Bill(I).����ժҪ, 0, c_Bill(I).ҽ��id);
    
      If c_Bill(I).ִ���� Is Not Null Then
        --���Ϊ��ȫִ��
        Update סԺ���ü�¼
        Set ִ��״̬ = 1, ִ���� = c_Bill(I).ִ����, ִ��ʱ�� = Sysdate
        Where ��¼���� = 2 And NO = v_No And ��� = c_Bill(I).���;
      End If;
    End If;
  End Loop;

  Xml_Out := Xmltype('<OUTPUT><RESULT>true</RESULT><DJH>' || v_No || '</DJH></OUTPUT>');
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Saveexes;
/

--129714:��͢��,2018-08-06,����ִ��ʱ���ڿ�ʼʱ��֮ǰ
Create Or Replace Procedure Zl_����ҽ��ִ��_Update
(
  ԭִ��ʱ��_In ����ҽ��ִ��.ִ��ʱ��%Type,
  ҽ��id_In     ����ҽ��ִ��.ҽ��id%Type,
  ���ͺ�_In     ����ҽ��ִ��.���ͺ�%Type,
  Ҫ��ʱ��_In   ����ҽ��ִ��.Ҫ��ʱ��%Type,
  ��������_In   ����ҽ��ִ��.��������%Type,
  ִ��ժҪ_In   ����ҽ��ִ��.ִ��ժҪ%Type,
  ִ����_In     ����ҽ��ִ��.ִ����%Type,
  ִ��ʱ��_In   ����ҽ��ִ��.ִ��ʱ��%Type,
  ִ�н��_In   ����ҽ��ִ��.ִ�н��%Type := 1,
  δִ��ԭ��_In ����ҽ��ִ��.˵��%Type := Null,
  ����ִ��_In   Number := 0,
  ����Ա���_In ��Ա��.���%Type := Null,
  ����Ա����_In ��Ա��.����%Type := Null,
  ִ�в���id_In ������ü�¼.ִ�в���id%Type := 0
  --������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID��
  --ִ�в���id_In=������ָ��ִ�в��ŵķ��ã���������0ʱ������ִ�в���
) Is
  --����Ҫִ�е�����¼,�������˸�������,��鲿λ�ļ�¼
  --��������,��ҩ�巨,�ɼ�������������,�������ֻ��д�ڵ�һ����Ŀ�ϣ���ִ��״̬��ͬ
  v_Temp     Varchar2(255); 
  v_��Ա���� ��Ա��.����%Type;

  v_��id        ����ҽ����¼.Id%Type;
  v_�������    ����ҽ����¼.�������%Type;
  v_ִ�н��old ����ҽ��ִ��.ִ�н��%Type;
  n_��������old ����ҽ��ִ��.��������%Type;

  v_������Դ ����ҽ����¼.������Դ%Type;
  v_�������� ����ҽ������.��¼����%Type;

  n_ִ�д��� Number;
  n_ʣ����� Number;
  n_ִ��״̬ Number;
  n_�������� Number;
  n_�������� Number;
  v_Count    Number;
  n_�Ǽ����� Number;
  d_Ҫ��ʱ�� Date;
  d_ִ��ʱ�� Date;
  d_��ʼʱ�� Date;

  d_�Ǽ�ʱ��   ����ҽ��ִ��.�Ǽ�ʱ��%Type;
  n_ȡ��ִ��   Number;
  n_Diffday    Number(18, 3);
  n_ִ�п���id Number;

  v_Date  Date;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --��ǰ������Ա
  If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null Then
    v_��Ա���� := ����Ա����_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  Select Sysdate Into v_Date From Dual;
  Select Nvl(ִ�н��, 1), Nvl(��������, 0), �Ǽ�ʱ��
  Into v_ִ�н��old, n_��������old, d_�Ǽ�ʱ��
  From ����ҽ��ִ��
  Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And ִ��ʱ�� = ԭִ��ʱ��_In;
  -----ȡ��ִ����Ч��������
  Select Zl_To_Number(Nvl(zl_GetSysParameter(220), '999')) Into n_ȡ��ִ�� From Dual;
  Select v_Date - d_�Ǽ�ʱ�� Into n_Diffday From Dual;
  --�Ǽ�ʱ�䳬��ȡ��ִ�������ļ�¼���������޸�ҽ��ִ�����
  If n_Diffday > n_ȡ��ִ�� Then
    v_Error := 'ҽ��ִ�еǼ�ʱ�䳬����ȡ��ִ����Ч�����������޸�ҽ��ִ�������';
    Raise Err_Custom;
  End If;

  If ��������_In = 1 Then
    --��ҽ����ʼʱ����м�� 
    Select a.��ʼִ��ʱ�� Into d_��ʼʱ�� From ����ҽ����¼ A Where a.Id = ҽ��id_In;
    If Not d_��ʼʱ�� Is Null Then
      If ִ��ʱ��_In < d_��ʼʱ�� Then
        v_Error := 'ִ��ʱ��������ҽ���Ŀ�ʼִ��ʱ��''' || To_Char(d_��ʼʱ��, 'yyyy-mm-dd HH24:mi:ss') || '''��';
        Raise Err_Custom;
      End If;
    End If;
  End If;

  Select ִ�в���id Into n_ִ�п���id From ����ҽ������ Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In;
  --����ҽ��ִ��
  Update ����ҽ��ִ��
  Set Ҫ��ʱ�� = Ҫ��ʱ��_In, �������� = ��������_In, ִ��ժҪ = ִ��ժҪ_In, ִ���� = ִ����_In, ִ��ʱ�� = ִ��ʱ��_In, �Ǽ�ʱ�� = v_Date, �Ǽ��� = v_��Ա����,
      ִ�н�� = ִ�н��_In, ˵�� = δִ��ԭ��_In, ִ�п���id = n_ִ�п���id
  Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And ִ��ʱ�� = ԭִ��ʱ��_In;
  --����ִ�д�������ִ�н���޸ĺ���Ҫ���µ��ݵ�ִ��״̬
  If v_ִ�н��old <> ִ�н��_In Or n_��������old <> ��������_In Then
    Select ������Դ, Nvl(���id, ID), �������
    Into v_������Դ, v_��id, v_�������
    From ����ҽ����¼
    Where ID = ҽ��id_In;
  
    If v_������Դ = 2 Then
      Select Decode(��¼����, 1, 1, Decode(�������, 1, 1, 2))
      Into v_��������
      From ����ҽ������
      Where ���ͺ� = ���ͺ�_In And ҽ��id = ҽ��id_In;
    Else
      v_�������� := 1;
    End If;
  
    Select Decode(a.ִ��״̬, 1, a.��������, c.�ǼǴ���), Decode(a.ִ��״̬, 1, 0, a.�������� - c.�ǼǴ���), a.��������, c.�ǼǴ���



    Into n_ִ�д���, n_ʣ�����, n_��������, n_�Ǽ�����
    From ����ҽ������ A,
         (Select ҽ��id_In ҽ��id, ���ͺ�_In ���ͺ�, Nvl(Sum(b.��������), 0) As �ǼǴ���
           From ����ҽ��ִ�� B
           Where b.ҽ��id = ҽ��id_In And b.���ͺ� = ���ͺ�_In And Nvl(b.ִ�н��, 1) = 1) C
    Where a.ҽ��id = c.ҽ��id And a.���ͺ� = c.���ͺ� And a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In;
  
    --���ȫ��ִ����״̬Ϊ1��δִ��״̬Ϊ0������ִ��״̬Ϊ2
    Select Decode(n_ʣ�����, 0, 1, Decode(n_ִ�д���, 0, 0, 2)) Into n_ִ��״̬ From Dual;
  
    --����ҽ��ִ�мƼ�.ִ��״̬
    If n_�������� > 0 Then
      Select Count(Distinct Ҫ��ʱ��) Into v_Count From ҽ��ִ�мƼ� Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In;
      If v_Count > 0 Then
        n_�������� := n_�������� / v_Count;
        --��ִ������+�������� �ܹ��ܹ�ִ�ж��ٸ�ʱ���,ȡ�������
        v_Count := Ceil((n_�Ǽ�����) / n_��������);
        If n_�Ǽ����� = 0 Then
          Update ҽ��ִ�мƼ�
          Set ִ��״̬ = 0
          Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And Nvl(ִ��״̬, 0) <> 2;
        Else
          --��ȡִ�н���Ҫ��ʱ��
          Select Ҫ��ʱ��
          Into d_Ҫ��ʱ��
          From (Select Ҫ��ʱ��, Rownum As ����
                 From (Select Distinct Ҫ��ʱ��
                        From ҽ��ִ�мƼ�
                        Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In
                        Order By Ҫ��ʱ��))
          Where ���� = v_Count;
        
          If Not d_Ҫ��ʱ�� Is Null Then
            --�ȼ���Ƿ��Ѿ��˷�
            Select Max(Nvl(ִ��״̬, 0))
            Into v_Count
            From ҽ��ִ�мƼ�
            Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And Ҫ��ʱ�� <= d_Ҫ��ʱ��;
            If v_Count = 2 Then
              v_Error := '��ָ����ִ��ʱ��ε�ҽ�������Ѿ����˷ѣ���������ִ�С�';
              Raise Err_Custom;
            End If;
            --���½���Ҫ��ʱ��֮ǰ(��)�ļ�¼ִ��״̬��
            Update ҽ��ִ�мƼ�
            Set ִ��״̬ = 1
            Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And Ҫ��ʱ�� <= d_Ҫ��ʱ�� And Nvl(ִ��״̬, 0) <> 2;
            Update ҽ��ִ�мƼ�
            Set ִ��״̬ = 0
            Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And Ҫ��ʱ�� > d_Ҫ��ʱ�� And Nvl(ִ��״̬, 0) <> 2;
          End If;
        End If;
      End If;
    End If;
  
    --ִ�д�����Ϊ0�ͱ��Ϊ����ִ��
    If Nvl(����ִ��_In, 0) = 1 Then
      Update ����ҽ������
      Set ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3), ����� = Null, ���ʱ�� = Null
      Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In;
    Else
      Update ����ҽ������
      Set ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3), ����� = Null, ���ʱ�� = Null
      Where ִ��״̬ In (0, 3) And ���ͺ� + 0 = ���ͺ�_In And
            ҽ��id In (Select ID From ����ҽ����¼ Where (ID = v_��id Or ���id = v_��id) And ������� = v_�������);
    End If;
  
    If v_�������� = 2 Then
      If Nvl(����ִ��_In, 0) = 1 Then
        Update סԺ���ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, d_ִ��ʱ��, ִ��ʱ��_In)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In
              (Select ҽ��id, NO, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In);
      Else
        Update סԺ���ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, d_ִ��ʱ��, ִ��ʱ��_In)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In
              (Select ҽ��id, NO, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ���ͺ� + 0 = ���ͺ�_In And
                     ҽ��id In
                     (Select ID From ����ҽ����¼ Where (ID = v_��id Or ���id = v_��id) And ������� = v_�������));
      End If;
    Else
      If Nvl(����ִ��_In, 0) = 1 Then
        Update ������ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, d_ִ��ʱ��, ִ��ʱ��_In)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In
              (Select ҽ��id, NO, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In);
      Else
        Update ������ü�¼ A
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, d_ִ��ʱ��, ִ��ʱ��_In)
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And
              (ҽ�����, NO, ��¼����) In
              (Select ҽ��id, NO, ��¼����
               From ����ҽ������
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ���ͺ� + 0 = ���ͺ�_In And
                     ҽ��id In
                     (Select ID From ����ҽ����¼ Where (ID = v_��id Or ���id = v_��id) And ������� = v_�������));
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ��ִ��_Update;
/





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0024' Where ���=&n_System;
Commit;
