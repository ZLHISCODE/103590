----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--126898:����,2018-07-09,�������ԡ���׼�ĺš�
Alter Table ҩƷ�ƻ����� Add ��׼�ĺ� varchar2(40);




------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--127075:����,2018-07-12,�޸ķ���ģ���������̨ǩ���Ŷ�
Update zlParameters
Set ���� = 1, ����˵�� = '1.���á��Ŷӽк�ģʽ������������Ч;' || Chr(13) || '2.���÷���̨ǩ���ŶӵĿ���ʱ���������������Դ�Ŀ��ң�Ĭ��Ϊ����̨ǩ���Ŷ�',
    ����˵�� = '����������������Ӱ���ŶӶ��еĲ���ʱ����',
    Ӱ�����˵�� = '1. �����˴˲�����,���ڷ���̨������"ǩ��"����,�ڲ���ǩ����,�Ž����ŶӶ���,����Һź�ͽ��ж���.' || Chr(13) ||
              '2. Ҳ���ڷ���̨ǩ���Ŷӵ�ģʽ�£�������Щ�����������ǩ���Ŷ�,���������;' || Chr(13) || '1��������������õĿ��ң���Ϊ���п��ҷ���̨ǩ���Ŷ�;' || Chr(13) ||
              '2)������������õĿ���,���δ���õĿ��ң��򲻰�����̨ǩ���ŶӵĹ�����,���ҺŻ�ԤԼ��ز��������Ŷ�;' || Chr(13) ||
              '3)������������õĿ���,��Ժ���������Դ�Ŀ��ң��򰴷���̨ǩ���Ŷӹ�����;' || Chr(13) || '4��������õĿ��ң���ֻ���ڷ���̨ǩ��ʱ�����Ŷӡ�'
Where ϵͳ = &n_System And ģ�� = 1113 And ������ = '����̨ǩ���Ŷ�';

--127487:����,2018-07-11,������������Ƿ������Ų���
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, Null, 0, 0, 0, 0, 0, 0, 305, '�����������Ų��ؿ���',  '1', '1',
         '���øò�����,���漰���������ĵط������������Ƿ�¼�������źͲ���', '0-�������������Ƿ�¼�������źͲ��أ�1-�����������Ƿ�¼�������źͲ��ء�', Null,
         '�������û���Ҫ�������������ʱ�Ƿ�¼�������źͲ���', Null
  From Dual;

--119905:Ƚ����,2018-07-09,�����շѹ������˲���Ʊ�ݺ��Һŷ�
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1121, 1, 0, 0, 0, 0, 0, 115, '�����˲���Ʊ�ݺ��Һŷ�', Null, '0', '�����˲���Ʊ��ʱ���Ƿ�ȱʡ��ȡ�Һŷ�',
         '0-��ȱʡ��ȡ�Һŷѣ�1-ȱʡ��ȡ�Һŷ�', Null, '���ڸ��Ի�����', Null
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
--125421:������,2018-07-12,����ʱ֧���������˷��û���
Create Or Replace Procedure Zl_����ҽ����¼_����
(
  Id_In         In ����ҽ����¼.Id%Type,
  ����Ա���_In In ��Ա��.���%Type := Null,
  ����Ա����_In In ��Ա��.����%Type := Null,
  ����ҽ��id_In In ����ҽ����¼.Id%Type := Null,
  ����ʱ��_In   In ����ҽ��״̬.����ʱ��%Type := Null
) Is
  --���ܣ�����ָ����ҽ��(δ���͵ĳ���������)
  --˵����һ����ҩ��ֻ�ܵ���һ��(������ʾ�ж���)
  --������ID_IN=��ҽ��ID
  --      ����ҽ��id_In ȡ�����������ϵĻ���ȼ�ҽ�����������Զ�ֹͣ�Ļ���ȼ�ҽ��id
  v_���ͺ�       ����ҽ������.���ͺ�%Type;
  v_����no       ������ü�¼.No%Type;
  v_��¼����     ������ü�¼.��¼����%Type;
  v_�������     Varchar2(255);
  n_�Զ�ȡ��ִ�� Number(1) := 0;
  n_�����Ϻ���ҩ Number(1) := 0;

  v_Date     Date;
  v_Count    Number;
  v_Temp     Varchar2(255);
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;
  v_No       ����ҽ������.No%Type;

  --����ҽ�������Ϣ
  Cursor c_Advice Is
    Select Nvl(a.���id, a.Id) As ��id, a.����id, a.�Һŵ�, a.��ҳid, a.Ӥ��, a.ҽ��״̬, a.�ϴ�ִ��ʱ��, a.ҽ������, a.�������, b.��������, a.������Դ,
           a.ִ�п���id, b.ִ��Ƶ��, a.������Ŀid, a.��ʼִ��ʱ��
    From ����ҽ����¼ a, ������ĿĿ¼ b
    Where a.������Ŀid = b.Id And a.Id = Id_In;
  r_Advice c_Advice%Rowtype;

  --����ҽ������ʱ��ȡ��Ӧ�ķ������ʻ�����(�շѻ��۵�)��
  --����ҽ��������NO������λ���Ҫ���ʻ��˷ѵļ�¼
  --һ��ҽ�������Ƕ���д�˷��ͼ�¼,Ҳ��һ�����Ʒ���,�ҿ���NO��ͬ
  --ֻ�ܼ�¼״̬Ϊ1�ļ�¼,����Ѿ����ʻ򲿷����ʵļ�¼,���ٴ���
  --����ֻ��۸񸸺�Ϊ�յ�,�Ա�ȡ�������
  --���������ҩ�������Ϻ���ҩ��,�򲻶���Ӧ����(������ҩ;����)���м��ʹ���,�����ǻ�û��ִ�еļ��ʵ�,��δִ�С��շѵĻ��۵���������ɾ�ˡ�

  Cursor c_Rollmoney(v_���ͺ� ����ҽ������.���ͺ�%Type) Is
    Select Decode(a.��¼����, 11, 1, a.��¼����) As ��¼����, a.��¼״̬, a.No, a.���, a.ִ��״̬ As ����ִ��, c.ִ��״̬ As ҽ��ִ��, c.ִ�в���id, b.���˿���id,
           b.�������, i.��������
    From ������ü�¼ a, ����ҽ����¼ b, ����ҽ������ c, ������ĿĿ¼ i
    Where c.ҽ��id = b.Id And c.���ͺ� = v_���ͺ� And (b.Id = Id_In Or b.���id = Id_In) And a.ҽ����� = b.Id And a.��¼״̬ In (0, 1) And
          a.No = c.No And (a.��¼���� = c.��¼���� Or a.��¼���� = 11 And c.��¼���� = 1) And b.������Ŀid = i.Id And a.�۸񸸺� Is Null And
          (n_�����Ϻ���ҩ = 0 Or
          n_�����Ϻ���ҩ = 1 And
          Not (Exists (Select 1
                        From ������ü�¼ d
                        Where d.ҽ����� = b.Id And d.��¼״̬ In (0, 1) And d.No = c.No And
                              (d.��¼���� = c.��¼���� Or d.��¼���� = 11 And c.��¼���� = 1) And d.�շ���� In ('5', '6', '7'))) Or
          Nvl(a.ִ��״̬, 0) = 0 And Not (a.��¼���� = 1 And a.��¼״̬ <> 0))
    Order By a.��¼����, a.No, a.���, a.�շ�ϸĿid;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --���ҽ��״̬�Ƿ���ȷ:��������
  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;

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

  --����Ƿ��Ѿ����˱��浥���Ѿ������浥��ҽ�����ܹ�����
  Select Count(1) Into v_Count From ����ҽ������ Where ҽ��id = Id_In;
  If v_Count > 0 Then
    If Not (r_Advice.�������� = '7' And r_Advice.������� = 'Z') Then
      v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"�ѳ����棬�������ϡ�';
      Raise Err_Custom;
    End If;
  End If;

  --����Ƿ�����Һ��Һ��¼�����Ƿ��Ѿ�����
  Select Count(1) Into v_Count From ��Һ��ҩ��¼ Where �Ƿ����� = 1 And ҽ��id = Id_In;
  If v_Count > 0 Then
    v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"����ҺҩƷ���Ѿ�����Һ���������������������ϡ�';
    Raise Err_Custom;
  End If;

  If r_Advice.�Һŵ� Is Null And r_Advice.������Դ <> 3 Then
    If r_Advice.ҽ��״̬ In (4, 8, 9) Then
      v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"�Ѿ������ϻ�ֹͣ�����������ϡ�';
      Raise Err_Custom;
    Elsif r_Advice.�ϴ�ִ��ʱ�� Is Not Null Then
      v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"�Ѿ����ͣ����ܱ����ϡ�';
      Raise Err_Custom;
    End If;
  
    --�����Ի���ȼ����뷢�ͣ�У�Ժ�Ϳ������Զ��Ʒѣ����ϼ��������϶�Ӧ��ֹͣ���̴���
    If r_Advice.������� = 'H' And r_Advice.�������� = '1' And r_Advice.ִ��Ƶ�� = '2' And Nvl(r_Advice.Ӥ��, 0) = 0 Then
      --(��ȡ�������ڴ����޷���Ժ�����������ţ�45977)a.��ʼʱ���ǵ���֮ǰ�ģ�˵������Ч���Զ����ü��㣩�����������ϡ�
      --ҽ����ʱ��ֻ��ȷ���˷��ӣ����Ա䶯��¼�Ŀ�ʼʱ��Ҫȥ�������Ƚϡ�
      v_Count := 0;
      Begin
        Select b.��ֹʱ��
        Into v_Date
        From ���˱䶯��¼ b, ����ҽ���Ƽ� c
        Where b.����id = r_Advice.����id And b.��ҳid = r_Advice.��ҳid And c.ҽ��id = Id_In And c.�շ�ϸĿid = b.����ȼ�id And
              b.��ʼԭ�� = 6 And b.���Ӵ�λ = 0 And
              To_Char(b.��ʼʱ��, 'yyyy-mm-dd hh24:mi') = To_Char(r_Advice.��ʼִ��ʱ��, 'yyyy-mm-dd hh24:mi');
      Exception
        When Others Then
          v_Count := 1;
      End;
      If v_Count = 0 Then
        --d.�����������䶯����
        If v_Date Is Not Null Then
          v_Error := '���ڻ���ȼ�ҽ����Ч���Ѿ������������䶯��¼,�������ϸ�ҽ����';
          Raise Err_Custom;
        Else
          --������Ҫ�Զ����õĻ���ȼ��������ԭ������ȼ���ͬ���ó�������䶯��¼
          If Nvl(����ҽ��id_In, 0) <> 0 Then
            Delete ����ҽ��״̬ Where ҽ��id = ����ҽ��id_In And �������� In (8, 9);
            Select ��������
            Into v_Count
            From (Select �������� From ����ҽ��״̬ Where ҽ��id = ����ҽ��id_In Order By ����ʱ�� Desc)
            Where Rownum < 2;
            Update ����ҽ����¼
            Set ҽ��״̬ = v_Count, ִ����ֹʱ�� = Null, ͣ��ҽ�� = Null, ͣ��ʱ�� = Null, ȷ��ͣ��ʱ�� = Null, ȷ��ͣ����ʿ = Null
            Where Id = ����ҽ��id_In;
            --�ų�����Ƶ���Ĳ���
            Select Count(a.Id)
            Into v_Count
            From ����ҽ����¼ a, �����շѹ�ϵ b, ������ҳ c
            Where a.������Ŀid = b.������Ŀid And c.����ȼ�id = b.�շ���Ŀid And c.����id = a.����id And c.��ҳid = a.��ҳid And
                  a.Id = ����ҽ��id_In;
          End If;
          If v_Count = 0 Then
            --c.����ȼ������һ���䶯
            Zl_���˱䶯��¼_Undo(r_Advice.����id, r_Advice.��ҳid, v_��Ա���, v_��Ա����, '1', Null, Null, '����ȼ��䶯');
          End If;
        End If;
      Else
        --�ָ����һ�α��Զ�ֹͣ�Ļ���ȼ�
        If Nvl(����ҽ��id_In, 0) <> 0 Then
          Delete ����ҽ��״̬ Where ҽ��id = ����ҽ��id_In And �������� In (8, 9);
          Select ��������
          Into v_Count
          From (Select �������� From ����ҽ��״̬ Where ҽ��id = ����ҽ��id_In Order By ����ʱ�� Desc)
          Where Rownum < 2;
          Update ����ҽ����¼
          Set ҽ��״̬ = v_Count, ִ����ֹʱ�� = Null, ͣ��ҽ�� = Null, ͣ��ʱ�� = Null, ȷ��ͣ��ʱ�� = Null, ȷ��ͣ����ʿ = Null
          Where Id = ����ҽ��id_In;
        Else
          --������Ժʱָ���Ļ��������ı䶯��¼��ҽ���¿������ı䶯��¼��ͬ������Ҫ���ж�
          Select Count(a.Id)
          Into v_Count
          From ���˱䶯��¼ a
          Where a.����id = r_Advice.����id And a.��ҳid = r_Advice.��ҳid And a.��ʼԭ�� = 6;
          If v_Count <> 0 Then
            --b.�������ǰ�Ļ���ȼ���ͬ����У��ʱû�в�������ȼ��䶯,��������ȼ�ֹͣ�䶯
            Zl_���˱䶯��¼_Nurse(r_Advice.����id, r_Advice.��ҳid, Null, Sysdate, v_��Ա���, v_��Ա����);
          End If;
        End If;
      End If;
    End If;
  Else
    If r_Advice.ҽ��״̬ <> 8 Then
      v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"��δ���ͻ��Ѿ����ϡ�';
      Raise Err_Custom;
    End If;
    --ҽ�������ж�
    Select Count(1)
    Into v_Count
    From ����ҽ������ a, ����ҽ����¼ b
    Where a.ҽ��id = b.Id And (b.Id = Id_In Or b.���id = Id_In);
    If v_Count <> 0 Then
      v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"���ڸ��ӷ��ã��������ϡ�';
      Raise Err_Custom;
    End If;
  
    Begin
      --ҽ��IDΪ����ֵ������ҽ����һ�������˵�,�����޷��͡�
      Select Distinct ���ͺ�
      Into v_���ͺ�
      From ����ҽ������
      Where ҽ��id In (Select Id From ����ҽ����¼ Where Id = Id_In Or ���id = Id_In);
    Exception
      When Others Then
        v_���ͺ� := Null;
    End;
  
    Select Zl_To_Number(Nvl(Zl_Getsysparameter(68), 0)) Into n_�����Ϻ���ҩ From Dual;
    Select Zl_To_Number(Nvl(Zl_Getsysparameter('���ﱾ���Զ�ִ��', '1252'), 0)) Into n_�Զ�ȡ��ִ�� From Dual;
    If n_�Զ�ȡ��ִ�� = 1 And v_���ͺ� Is Not Null Then
      --�ȸ���ҽ���ͷ��õ�ִ��״̬����Ϊ�������жϣ��Լ�����Zl_������ʼ�¼_Delete���м��
      For Rc In (Select a.ҽ��id, a.ִ�в���id
                 From ����ҽ������ a, ����ҽ����¼ b
                 Where a.ҽ��id = b.Id And (b.Id = Id_In Or b.���id = Id_In) And a.ִ�в���id = b.���˿���id) Loop
        Zl_����ҽ��ִ��_Cancel(Rc.ҽ��id, v_���ͺ�, Null, 1, Rc.ִ�в���id);
      End Loop;
    End If;
  
    --����ҽ��ֻ���ܷ���һ��
    --�����˷�ʱ���м�飬��Ϊ����ҽ��û�з��ã�����Ҫ���һ��ִ��״̬
    Select Count(*)
    Into v_Count
    From ����ҽ������ a, ����ҽ����¼ b, ������ĿĿ¼ i
    Where a.ҽ��id = b.Id And b.������Ŀid = i.Id And a.ִ��״̬ In (1, 3) And (b.Id = Id_In Or b.���id = Id_In) And
          (n_�����Ϻ���ҩ = 0 Or
          n_�����Ϻ���ҩ = 1 And Not (b.������� In ('5', '6', '7') Or b.������� = 'E' And i.�������� In ('2', '3', '4')));
    If v_Count > 0 Then
      v_Error := '��ҽ���Ѿ�ִ�л�����ִ�У��������ϡ�';
      Raise Err_Custom;
    End If;
  End If;

  If ����ʱ��_In Is Null Then
    Select Sysdate Into v_Date From Dual;
  Else
    v_Date := ����ʱ��_In;
  End If;

  Update ����ҽ����¼ Set ҽ��״̬ = 4 Where Id = Id_In Or ���id = Id_In;

  Insert Into ����ҽ��״̬
    (ҽ��id, ��������, ������Ա, ����ʱ��)
    Select Id, 4, v_��Ա����, v_Date From ����ҽ����¼ Where Id = Id_In Or ���id = Id_In;

  --סԺҽ������ʱ,δ��ӡ�������,ȱʡ����Ϊ���δ�ӡ
  If r_Advice.�Һŵ� Is Null And r_Advice.������Դ <> 3 Then
    Select Count(*)
    Into v_Count
    From ����ҽ����ӡ
    Where ҽ��id In (Select Id From ����ҽ����¼ Where Id = Id_In Or ���id = Id_In);
    If Nvl(v_Count, 0) = 0 Then
      Zl_����ҽ����¼_���δ�ӡ(Id_In, 1);
    End If;
    If Nvl(r_Advice.Ӥ��, 0) > 0 And r_Advice.�������� = '11' Then
      Update ������������¼
      Set ����ʱ�� = Null
      Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = Nvl(r_Advice.��ҳid, 0) And ��� = Nvl(r_Advice.Ӥ��, 0);
    End If;
  Else
    --����ҽ��(����)����ʱ����Ҫ�����������:ֻ��һ�η���
    --���˻��ۻ���ʷ���
    If v_���ͺ� Is Not Null Then
      --������ҽ���ķ���ɾ��������(��һ��ҽ�������в�ͬNO����)
      --������ʣ����ԭʼ�����ѱ�����(�򲿷�����),���ù��������ж�
      --���ﻮ�ۣ�������շѣ�������ɾ��
      v_����no   := Null;
      v_������� := Null;
      For r_Rollmoney In c_Rollmoney(v_���ͺ�) Loop
        If Nvl(r_Rollmoney.ҽ��ִ��, 0) In (1, 3) Then
          --1-��ȫִ��;3-����ִ��
          v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"�Ѿ�ִ�л�����ִ�У��������ϡ�';
          Raise Err_Custom;
        End If;
        If Nvl(r_Rollmoney.����ִ��, 0) In (1, 2) Then
          --1-��ȫִ��;2-����ִ��
          v_Error := 'ҽ�����õ���"' || r_Rollmoney.No || '"�е������Ѿ�ȫ���򲿷�ִ�У��������ϡ�';
          Raise Err_Custom;
        End If;
        If r_Rollmoney.����ִ�� = 9 Then
          v_Error := 'ҽ�����õ���"' || r_Rollmoney.No || '"�е��շѽ�������쳣���������ϡ�';
          Raise Err_Custom;
        End If;
        v_Count := 1;
        If r_Rollmoney.��¼���� = 1 And r_Rollmoney.��¼״̬ <> 0 Then
          If 1 = n_�����Ϻ���ҩ And r_Rollmoney.������� = 'E' And r_Rollmoney.�������� In ('2', '3', '4') Then
            v_Count := 0;
          Else
            v_Error := 'ҽ�����õ���"' || r_Rollmoney.No || '"�Ѿ��շѣ��������ϡ�';
            Raise Err_Custom;
          End If;
        End If;
        If 1 = v_Count Then
          If Nvl(v_����no, '��') <> r_Rollmoney.No Then
            If v_������� Is Not Null And v_����no Is Not Null Then
              v_������� := Substr(v_�������, 2);
              If v_��¼���� = 1 Then
                Zl_���ﻮ�ۼ�¼_Delete(v_����no, v_�������);
              Elsif v_��¼���� = 2 Then
                Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����);
              End If;
            End If;
            v_������� := Null;
          End If;
          v_��¼���� := r_Rollmoney.��¼����;
          v_����no   := r_Rollmoney.No;
          v_������� := v_������� || ',' || r_Rollmoney.���;
        End If;
      End Loop;
      If v_������� Is Not Null And v_����no Is Not Null Then
        v_������� := Substr(v_�������, 2);
        If v_��¼���� = 1 Then
          Zl_���ﻮ�ۼ�¼_Delete(v_����no, v_�������);
        Elsif v_��¼���� = 2 Then
          Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����);
        End If;
      End If;
    
      --���"����ҩ�������Ϻ���ҩ"�����Ӧ�ĸ�ҩ;����������Ϊδִ�У��Ա��˷�
      If n_�����Ϻ���ҩ = 1 Then
        Update ������ü�¼
        Set ִ��״̬ = 0
        Where ִ��״̬ = 1 And ҽ����� = Id_In And Exists
         (Select 1
               From ����ҽ����¼ a, ������ĿĿ¼ b
               Where a.������Ŀid = b.Id And b.��� = 'E' And b.�������� In ('2', '3', '4') And a.Id = Id_In);
      End If;
    
      --����ҽ�����ͼ�¼(��ִ�м�¼)
      Delete From ����ҽ��ִ�� Where ҽ��id In (Select Id From ����ҽ����¼ Where Id = Id_In Or ���id = Id_In);
      Delete From ����ҽ������ Where ҽ��id In (Select Id From ����ҽ����¼ Where Id = Id_In Or ���id = Id_In);
    
      --��������ҽ���Ĵ���
      If r_Advice.������� = 'Z' And Nvl(r_Advice.��������, '0') <> '0' And Nvl(r_Advice.Ӥ��, 0) = 0 Then
      
        If r_Advice.�������� = '1' And r_Advice.ִ�п���id Is Not Null Then
          --����ҽ��
          Select Count(*)
          Into v_Count
          From ������ҳ
          Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = 0 And ��Ժ����id = r_Advice.ִ�п���id And �������� In (1, 2);
          If v_Count = 1 Then
            Zl_��Ժ������ҳ_Delete(r_Advice.����id, 0);
          End If;
        Elsif r_Advice.�������� = '2' And r_Advice.ִ�п���id Is Not Null Then
          --סԺҽ��
          Select Count(*)
          Into v_Count
          From ������ҳ
          Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = 0 And ��Ժ����id = r_Advice.ִ�п���id And Nvl(��������, 0) = 0;
          If v_Count = 1 Then
            Zl_��Ժ������ҳ_Delete(r_Advice.����id, 0);
          End If;
        End If;
      End If;
    End If;
  End If;

  --ɾ�������ǼǼ�¼
  If r_Advice.������� = 'E' And r_Advice.�������� = '1' Then
    --Update ����ҽ����¼ Set Ƥ�Խ��=Null Where ID=ID_IN; --��������Ƥ�Խ��
    --ɾ���������ļ�¼��������¼��������Ϊ����ҽ���Ƿ����ϣ����˶Ը�ҩ����
    For r_Test In (Select ����ʱ�� From ����ҽ��״̬ Where ҽ��id = Id_In And �������� = 10) Loop
      Delete From ���˹�����¼
      Where ����id = r_Advice.����id And ��¼��Դ = 2 And Nvl(��ҳid, 0) = Nvl(r_Advice.��ҳid, 0) And ��¼ʱ�� = r_Test.����ʱ�� And
            Nvl(���, 0) = 0;
    End Loop;
  End If;
  If r_Advice.������Ŀid Is Not Null Then
    --������¼��ҽ����סԺ����ִ��ҽ�����ϣ�Zlhis_Cis_003
    If r_Advice.��ҳid Is Not Null Then
      For r In (Select a.Id
                From ����ҽ����¼ a
                Where (a.Id = Id_In Or a.���id = Id_In) And Exists
                 (Select 1 From ��������˵�� b Where b.����id = a.ִ�п���id And b.�������� = '����')) Loop
        b_Message.Zlhis_Cis_003(r_Advice.����id, r_Advice.��ҳid, Null, r.Id);
      End Loop;
    
      If r_Advice.������� = 'Z' And r_Advice.�������� = '4' Then
        --����ҽ������
        b_Message.Zlhis_Cis_005(r_Advice.����id, r_Advice.��ҳid, r_Advice.��id);
      End If;
    End If;
  
    If r_Advice.�Һŵ� Is Not Null Then
      If r_Advice.������� = 'E' And r_Advice.�������� = '6' Or Instr(',D,F,K,', r_Advice.�������) > 0 Then
        Select Max(a.No) Into v_No From ����ҽ������ a Where a.ҽ��id = r_Advice.��id;
        If r_Advice.������� = 'E' And r_Advice.�������� = '6' Then
          --����
          b_Message.Zlhis_Cis_036(r_Advice.����id, Null, r_Advice.�Һŵ�, v_���ͺ�, r_Advice.��id, v_No, 1);
        Elsif r_Advice.������� = 'D' Then
          --���
          b_Message.Zlhis_Cis_037(r_Advice.����id, Null, r_Advice.�Һŵ�, v_���ͺ�, r_Advice.��id, v_No, 1);
        Elsif r_Advice.������� = 'F' Then
          --����
          b_Message.Zlhis_Cis_038(r_Advice.����id, Null, r_Advice.�Һŵ�, v_���ͺ�, r_Advice.��id, v_No);
        Elsif r_Advice.������� = 'K' Then
          --��Ѫ
          b_Message.Zlhis_Cis_039(r_Advice.����id, Null, r_Advice.�Һŵ�, v_���ͺ�, r_Advice.��id, v_No);
        End If;
      End If;
    End If;
  End If;
  Close c_Advice;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ҽ����¼_����;
/
--125595:���ϴ�,2018-07-11,�Һ���Ŵ�������
Create Or Replace Procedure Zl_����ԤԼ�Һż�¼_Update
(
  ���ݺ�_In     ������ü�¼.No%Type,
  ���_In       ������ü�¼.���%Type,
  �۸񸸺�_In   ������ü�¼.�۸񸸺�%Type,
  ��������_In   ������ü�¼.��������%Type,
  �շ����_In   ������ü�¼.�շ����%Type,
  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
  ����_In       ������ü�¼.����%Type,
  ��׼����_In   ������ü�¼.��׼����%Type,
  ������Ŀid_In ������ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In   ������ü�¼.�վݷ�Ŀ%Type,
  Ӧ�ս��_In   ������ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In   ������ü�¼.ʵ�ս��%Type,
  ������_In     Number, --������¼�Ƿ���������
  ���մ���id_In ������ü�¼.���մ���id%Type,
  ������Ŀ��_In ������ü�¼.������Ŀ��%Type,
  ͳ����_In   ������ü�¼.ͳ����%Type,
  ���ձ���_In   ������ü�¼.���ձ���%Type,
  ���˿���id_In ������ü�¼.���˿���id%Type,
  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
  ժҪ_In       ������ü�¼.ժҪ%Type := Null,
  �Ƿ�Һ���_In Number := 0
) As
  v_����id ������ü�¼.Id%Type;
  v_Error  Varchar2(255);
  Err_Custom Exception;
  Cursor c_���� Is
    Select ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�,
           �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��,
           �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���
    From ������ü�¼
    Where NO = ���ݺ�_In And ��¼���� = 4 And ��� = 1 And ��¼״̬ = 0;
Begin

  If Nvl(���_In, 1) = 1 Then
    --��һ����¼,ֻ��������
    Update ������ü�¼
    Set �۸񸸺� = Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), �������� = Decode(��������_In, 0, Null, ��������_In), ���ӱ�־ = ������_In,
        �շ���� = �շ����_In, �շ�ϸĿid = �շ�ϸĿid_In, ������Ŀid = ������Ŀid_In, �վݷ�Ŀ = �վݷ�Ŀ_In, ���� = 1, ���� = ����_In, ��׼���� = ��׼����_In,
        Ӧ�ս�� = Ӧ�ս��_In, ʵ�ս�� = ʵ�ս��_In, ���մ���id = ���մ���id_In, ������Ŀ�� = ������Ŀ��_In, ���ձ��� = ���ձ���_In, ͳ���� = ͳ����_In,
        ���˿���id =  Decode(�Ƿ�Һ���_In, 1, ���˿���id, ���˿���id_In), ִ�в���id = Decode(�Ƿ�Һ���_In, 1, ִ�в���id, ִ�в���id_In), ժҪ = Nvl(ժҪ_In, ժҪ)
    Where NO = ���ݺ�_In And ��� = 1 And ��¼״̬ = 0 And ��¼���� = 4;
    --ɾ����Ŵ���1������;
    Delete ������ü�¼ Where NO = ���ݺ�_In And ��� > 1 And ��¼���� = 4;
  Else
    --��������
    If ������_In <> 3 Then
      Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual; --Ӧ��ͨ������õ�
      For r_���� In c_���� Loop
        Insert Into ������ü�¼
          (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id,
           �շ����, ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����,
           ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����)
        Values
          (v_����id, 4, 0, ���_In, Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), ��������_In, ���ݺ�_In, r_����.ʵ��Ʊ��, 1, r_����.�Ӱ��־, ������_In,
           r_����.��ҩ����, r_����.����id, r_����.��ʶ��, r_����.���ʽ, r_����.����, r_����.�Ա�, r_����.����, r_����.�ѱ�, ���˿���id_In, �շ����_In, r_����.���㵥λ,
           �շ�ϸĿid_In, ������Ŀid_In, �վݷ�Ŀ_In, 1, ����_In, ��׼����_In, Ӧ�ս��_In, ʵ�ս��_In, Null, Null, 0, r_����.��������id, r_����.����Ա����,
           r_����.����Ա����, ִ�в���id_In, r_����.ִ����, r_����.����Ա���, r_����.����Ա����, r_����.����ʱ��, r_����.�Ǽ�ʱ��, ���մ���id_In, ������Ŀ��_In, ���ձ���_In,
           ͳ����_In, Nvl(ժҪ_In, r_����.ժҪ), r_����.����);
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ԤԼ�Һż�¼_Update;
/

--125075:����,2018-07-10,����ʹ�õ��˲��Ų�������̨ǩ���Ŷӵ�Oracle����
--127075:����,2018-07-10,����ʹ�õ��˲��Ų�������̨ǩ���Ŷӵ�Oracle����
Create Or Replace Procedure Zl_���˹Һż�¼_ǩ��
(
  Id_In       ���˹Һż�¼.Id%Type,
  ��������_In Integer := 0,
  ԤԼ��ʽ_In ԤԼ��ʽ.����%Type := Null,
  ����_In     ���˹Һż�¼.����%Type := Null,
  ҽ��_In     ���˹Һż�¼.ִ����%Type := Null
  
) As
  --:��������_In:0-��Ҫ������ص��ŶӼ�¼��������ǩ��,�޸��Ŷ���Ϣ;1-���ٴ����ŶӼ�¼,ֻ��ǩ���ı�ʶ��д�� 
  Err_Item Exception;
  v_Err_Msg          Varchar2(200);
  n_�Һ����ɶ���     Number;
  n_����̨ǩ���Ŷ�   Number;
  n_�ٴ�ǩ�������Ŷ� Number;
  n_���Ŷ�           Number;
  v_ԭ��������       �ŶӽкŶ���.��������%Type;
  v_�ֶ�������       �ŶӽкŶ���.��������%Type;
  v_��������         ���˹Һż�¼.����%Type;
  n_�Һ�id           ���˹Һż�¼.Id%Type;
  n_ִ�в���id       ���˹Һż�¼.ִ�в���id%Type;
  n_�ŶӲ���id       ���˹Һż�¼.Id%Type;
  v_ԭʼ����         �ŶӽкŶ���.�ŶӺ���%Type;
  v_�ŶӺ���         �ŶӽкŶ���.�ŶӺ���%Type;
  d_�Ŷ�ʱ��         �ŶӽкŶ���.�Ŷ�ʱ��%Type;
  v_����             ���˹Һż�¼.����%Type;
  v_ҽ��             ���˹Һż�¼.ִ����%Type;
  n_����id           ���˹Һż�¼.����id%Type;
  v_�ű�             ���˹Һż�¼.�ű�%Type;
  n_����             ���˹Һż�¼.����%Type;
  v_�Ŷ����         �ŶӽкŶ���.�Ŷ����%Type;
  n_��¼����         ���˹Һż�¼.��¼����%Type;
  d_����ʱ��         ���˹Һż�¼.����ʱ��%Type;
  n_ת��             Number;
  n_����             Number;
  v_�Һŵ�           ���˹Һż�¼.No%Type;
  v_ԤԼ��ʽ         ���˹Һż�¼.ԤԼ��ʽ%Type;
  n_��Һ�ģʽ       Number;
  v_No               ���˹Һż�¼.No%Type;
Begin
  --:��¼��־:0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���; 
  Update ���˹Һż�¼
  Set ��¼��־ = 1, ���� = Nvl(����_In, ����), ִ���� = Nvl(ҽ��_In, ִ����)
  Where ID = Id_In
  Returning ԤԼ��ʽ, NO Into v_ԤԼ��ʽ, v_No;

  If Sql%NotFound Then
    v_Err_Msg := '�Һ���Ŀδ�ҵ�,���ܼ�������!';
    Raise Err_Item;
  End If;
  Update ������ü�¼
  Set ִ���� = Nvl(ҽ��_In, ִ����), ��ҩ���� = Nvl(����_In, ��ҩ����)
  Where NO = v_No And ��¼���� = 4 And ��¼״̬ In (0, 1, 3);

  If Nvl(��������_In, 0) = 1 Then
    Return;
  End If;

  If ԤԼ��ʽ_In Is Not Null Then
    v_ԤԼ��ʽ := ԤԼ��ʽ_In;
  End If;

  Begin
    Select ID, Nvl(ת�����id, ִ�в���id), ����, Decode(ת�����id, Null, ����, ת������), Decode(ת�����id, Null, ִ����, ת��ҽ��), ����id,
           Nvl(ת��ű�, �ű�), ����, Nvl(��¼����, 0), Decode(ת�����id, Null, 0, 1), NO, ����ʱ��
    Into n_�Һ�id, n_ִ�в���id, v_��������, v_����, v_ҽ��, n_����id, v_�ű�, n_����, n_��¼����, n_ת��, v_�Һŵ�, d_����ʱ��
    From ���˹Һż�¼
    Where ID = Id_In And Rownum = 1;
  Exception
    When Others Then
      n_�Һ�id := -1;
  End;

  n_�Һ����ɶ���     := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
  n_����̨ǩ���Ŷ�   := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(n_ִ�в���id, 0)));
  n_�ٴ�ǩ�������Ŷ� := Zl_To_Number(zl_GetSysParameter('�ٴ�ǩ���������Ŷ�', 1113));

  n_��Һ�ģʽ := Zl_To_Number(zl_GetSysParameter('��Һ�ģʽ', 0));
  If Nvl(n_��Һ�ģʽ, 0) = 1 And Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
    n_����̨ǩ���Ŷ� := 1;
  End If;

  If n_�Һ����ɶ��� = 0 Then
    Return;
  End If;

  Begin
    Select 1, ����id, �ŶӺ���, ��������
    Into n_���Ŷ�, n_�ŶӲ���id, v_ԭʼ����, v_ԭ��������
    From �ŶӽкŶ���
    Where ҵ��id = Id_In And ҵ������ = 0;
  Exception
    When Others Then
      n_���Ŷ� := 0;
  End;
  --������ŶӼ�¼��˵��������ǩ������Ӧ���˳� 
  If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 And Nvl(n_���Ŷ�, 0) = 0 Then
    Return;
  End If;

  Begin
    Select 1 Into n_���� From ����䶯��¼ Where �Һŵ� = v_�Һŵ� And Rownum = 1;
  Exception
    When Others Then
      n_���� := 0;
  End;

  v_�ֶ������� := n_ִ�в���id;
  If n_�Һ�id > 0 Then
    If n_���Ŷ� > 0 And (Nvl(n_�ٴ�ǩ�������Ŷ�, 0) = 0 Or Nvl(n_��¼����, 0) = 2) Then
      v_�ŶӺ��� := Zl_Get_Requeue(0, n_�Һ�id, n_ִ�в���id, v_ҽ��, v_����);
      --���»�ȡ�Ŷ�ʱ�� 
      d_�Ŷ�ʱ�� := Zl_Get_Requeuedate(0, n_�Һ�id, n_ִ�в���id, v_ҽ��, v_����);
      If v_ԭʼ���� <> v_�ŶӺ��� Then
        --���»�ȡ��� 
        v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 1);
        --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In , ҽ������_In 
        Zl_�ŶӽкŶ���_Update(v_�ֶ�������, 0, n_�Һ�id, n_ִ�в���id, v_��������, v_����, v_ҽ��, v_�ŶӺ���, v_�Ŷ����, d_�Ŷ�ʱ��);
      Else
        --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In , ҽ������_In 
        Zl_�ŶӽкŶ���_Update(v_�ֶ�������, 0, n_�Һ�id, n_ִ�в���id, v_��������, v_����, v_ҽ��);
      End If;
    Elsif n_���Ŷ� > 0 Then
      --�����Ŷ� 
      v_�ŶӺ��� := Zlgetnextqueue(n_ִ�в���id, n_�Һ�id, v_�ű� || '|' || n_����);
      v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
      --���»�ȡ�Ŷ�ʱ�� 
      d_�Ŷ�ʱ�� := Zl_Get_Requeuedate(0, n_�Һ�id, n_ִ�в���id, v_ҽ��, v_����);
      --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In , ҽ������_In 
      Zl_�ŶӽкŶ���_Update(v_�ֶ�������, 0, n_�Һ�id, n_ִ�в���id, v_��������, v_����, v_ҽ��, v_�ŶӺ���, v_�Ŷ����, d_�Ŷ�ʱ��);
    Else
      --�����Ŷ� 
      For v_���� In (Select ��������, ҵ��id
                   From �ŶӽкŶ���
                   Where ����id = n_����id And ҵ������ = 0 And Trunc(�Ŷ�ʱ��) < Sysdate) Loop
        --ɾ��ԭ���ŶӼ�¼�����Ŷӣ���������_IN��ҵ��ID_IN 
        Zl_�ŶӽкŶ���_Delete(v_����.��������, v_����.ҵ��id);
      End Loop;
    
      v_�ŶӺ��� := Zlgetnextqueue(n_ִ�в���id, n_�Һ�id, v_�ű� || '|' || n_����);
      v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
      --���»�ȡ�Ŷ�ʱ�� 
      If n_ת�� = 1 Then
        d_�Ŷ�ʱ�� := Zl_Get_Requeuedate(2, n_�Һ�id, n_ִ�в���id, v_ҽ��, v_����);
      Elsif n_���� = 1 Then
        d_�Ŷ�ʱ�� := Zl_Get_Requeuedate(3, n_�Һ�id, n_ִ�в���id, v_ҽ��, v_����);
      Else
        d_�Ŷ�ʱ�� := Zl_Get_Requeuedate(-1, n_�Һ�id, n_ִ�в���id, v_ҽ��, v_����);
      End If;
      --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In 
      Zl_�ŶӽкŶ���_Insert(v_�ֶ�������, 0, n_�Һ�id, n_ִ�в���id, v_�ŶӺ���, Null, v_��������, n_����id, v_����, v_ҽ��, d_�Ŷ�ʱ��, v_ԤԼ��ʽ, Null,
                       v_�Ŷ����);
    End If;
    Update �ŶӽкŶ��� Set �Ŷ�״̬ = 0 Where ҵ��id = n_�Һ�id And ҵ������ = 0;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_ǩ��;
/

--126898:����,2018-07-09,���Ӵ��Ρ���׼�ĺš�
Create Or Replace Procedure Zl_ҩƷ�ƻ�����α�_Insert
(
  �ƻ�id_In     In ҩƷ�ƻ�����.�ƻ�id%Type,
  ҩƷid_In     In ҩƷ�ƻ�����.ҩƷid%Type,
  ���_In       In ҩƷ�ƻ�����.���%Type,
  �ƻ�����_In   In ҩƷ�ƻ�����.�ƻ�����%Type,
  ����_In       In ҩƷ�ƻ�����.����%Type := Null,
  ���_In       In ҩƷ�ƻ�����.���%Type := Null,
  ǰ������_In   In ҩƷ�ƻ�����.ǰ������%Type := Null,
  ��������_In   In ҩƷ�ƻ�����.��������%Type := Null,
  �������_In   In ҩƷ�ƻ�����.�������%Type := Null,
  �ϴι�Ӧ��_In In ҩƷ�ƻ�����.�ϴι�Ӧ��%Type := Null,
  �ϴ�������_In In ҩƷ�ƻ�����.�ϴ�������%Type := Null,
  ˵��_In       In ҩƷ�ƻ�����.˵��%Type := Null,
  �ۼ�_In       In ҩƷ�ƻ�����.�ۼ�%Type := Null,
  �ۼ۽��_In   In ҩƷ�ƻ�����.�ۼ۽��%Type := Null,
  ��������_In   In ҩƷ�ƻ�����.��������%Type := Null,
  ��������_In   In ҩƷ�ƻ�����.��������%Type := Null,
  �ͻ�����_In   In ҩƷ�ƻ�����.�ͻ�����%Type := Null,
  ��׼�ĺ�_In   In ҩƷ�ƻ�����.��׼�ĺ�%Type := Null
) Is
Begin
  Insert Into ҩƷ�ƻ�����
    (�ƻ�id, ҩƷid, ���, ǰ������, ��������, �������, �ƻ�����, ����, ���, �ϴι�Ӧ��, �ϴ�������, ˵��, �ۼ�, �ۼ۽��, ��������, ��������, �ͻ�����, ��׼�ĺ�)
  Values
    (�ƻ�id_In, ҩƷid_In, ���_In, ǰ������_In, ��������_In, �������_In, �ƻ�����_In, ����_In, ���_In, �ϴι�Ӧ��_In, �ϴ�������_In, ˵��_In, �ۼ�_In,
     �ۼ۽��_In, ��������_In, ��������_In, �ͻ�����_In, ��׼�ĺ�_In);
End Zl_ҩƷ�ƻ�����α�_Insert;
/

--125595:���ϴ�,2018-07-11,�Һ���Ŵ�������
--127075:����,2018-07-06,����ʹ�õ��˲��Ų�������̨ǩ���Ŷӵ�Oracle����
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
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id, Min(Decode(��¼����, 1, �տ�ʱ��, Null)) As �տ�ʱ��
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1
     Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By NO, ����id
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), �տ�ʱ��;
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

  n_��� := ����_In;

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
      Where ��¼id = �����¼id_In And ��ʼʱ�� = ����ʱ��_In And Rownum < 2;
    Exception
      When Others Then
        n_��� := Null;
    End;
  End If;

  --��ʱ�ιҺ�ʱ�ж��Ƿ��ǹ��ڹҺ� Ҳ����׷�Ӻŵ���� Ŀǰֻ���ר�Һŷ�ʱ�ν��д���
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 And Nvl(n_��ſ���, 0) = 1 And ����_In Is Null And Nvl(��������_In, 0) = 0 Then
    Begin
      Select Max(To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_������ʱ��
      From �ٴ�������ſ���
      Where ��¼id = �����¼id_In And Nvl(����, 0) <> 0;
    
      n_׷�Ӻ� := Case Sign(����ʱ��_In - d_������ʱ��)
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
  d_ʱ��ʱ�� := ����ʱ��_In;

  If ���_In = 1 And n_��ʱ�� > 0 Then
    If Nvl(n_��ſ���, 0) = 1 Then
      Begin
        Select Nvl(���, 0),
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               ����, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ����)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �ٴ�������ſ���
        Where ��¼id = �����¼id_In And ��� = n_���;
      Exception
        When Others Then
          n_ʱ����� := -1;
          n_��ʱ��   := 0;
          d_ʱ��ʱ�� := ����ʱ��_In;
          n_ʱ���޺� := 0;
          n_ʱ����Լ := 0;
      End;
    Else
      --�Һ�ʱ��� �Ƿ����ʱ��,����ʱ��,��ʱ�ε�����������ȡ����
      Begin
        Select Nvl(���, 0),
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               ����, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ����)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �ٴ�������ſ���
        Where ��¼id = �����¼id_In And ��� = n_��� And ԤԼ˳��� Is Null;
      Exception
        When Others Then
          n_ʱ����� := -1;
          n_��ʱ��   := 0;
          d_ʱ��ʱ�� := ����ʱ��_In;
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
        Select Min(���) Into n_������� From �ٴ�������ſ��� Where ��¼id = �����¼id_In And Nvl(�Һ�״̬, 0) = 0;
        If n_��� Is Null Then
          n_��� := Nvl(n_�������, 0);
        End If;
        IF nvl(n_���,0)=0 THEN 
          Select Nvl(Max(���), 0) + 1 Into n_��� From �ٴ�������ſ��� Where ��¼id = �����¼id_In;
	      END IF;
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
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd ') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        
        Else
          --ԤԼʱ���
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
        
          If n_��Լ�� <= n_��Լ�� And n_��Լ�� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd') || '�Ѵﵽ�����Լ����';
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
          v_Err_Msg := '�ű�' || �ű�_In || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
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
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
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
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
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
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        Else
          --ԤԼ
          If (Nvl(n_��Լ��, 0) > 0 And Nvl(n_��Լ��, 0) >= Nvl(n_��Լ��, 0)) Or
             (Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0) Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
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
      d_���ʱ�� := ����ʱ��_In;
    Else
      d_���ʱ�� := Trunc(����ʱ��_In);
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
                    Select �����¼id_In, n_���, ����ʱ��_In, ����ʱ��_In, 1, Decode(ԤԼ�Һ�_In, 1, 1, 0), Decode(ԤԼ�Һ�_In, 1, 2, 1),
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
     ����Ա����_In, Decode(ԤԼ�Һ�_In, 1, ����Ա����_In, Null), ִ�в���id_In, ҽ������_In, ����Ա���_In, ����Ա����_In, ����ʱ��_In, �Ǽ�ʱ��_In, ���մ���id_In,
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
      Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = ����_In Where ����id = ����id_In;
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
       Null, ִ�в���id_In, ҽ������_In, 0, Null, �Ǽ�ʱ��_In, ����ʱ��_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����ʱ��_In, Null), ����Ա���_In,
       ����Ա����_In, ����_In, n_���, ����_In, Decode(ԤԼ����_In, 1, 1, 0), ԤԼ��ʽ_In, ժҪ_In, ������ˮ��_In, ����˵��_In, ������λ_In,
       Decode(Nvl(ԤԼ�Һ�_In, 0), 0, �Ǽ�ʱ��_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 0, ����Ա����_In, Null),
       Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա����_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա���_In, Null),
       Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, �����¼id_In, �շѵ�_In);
  
    If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
      Update ���˹Һż�¼
      Set ԤԼ = 1, ԤԼʱ�� = ����ʱ��_In, ԤԼ����Ա = ����Ա����_In, ԤԼ����Ա��� = ����Ա���_In
      Where ID = n_�Һ�id;
    End If;
  
    n_ԤԼ���ɶ��� := 0;
    If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
      n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    End If;
  
    --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
    If Nvl(���ɶ���_In, 0) <> 0 And Nvl(ԤԼ�Һ�_In, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(ִ�в���id_In, 0)));
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
        d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, �ű�_In, n_���, ����ʱ��_In);
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

--127075:����,2018-07-06,����ʹ�õ��˲��Ų�������̨ǩ���Ŷӵ�Oracle����
Create Or Replace Procedure Zl_���˹Һż�¼_����_Delete
(
  ���ݺ�_In       ������ü�¼.No%Type,
  ����Ա���_In   ������ü�¼.����Ա���%Type,
  ����Ա����_In   ������ü�¼.����Ա����%Type,
  ժҪ_In         ������ü�¼.ժҪ%Type := Null, --ԤԼȡ��ʱ ��д ���ԤԼȡ��ԭ��
  ɾ�������_In   Number := 0,
  ��ԭ���˽���_In Varchar2 := Null,
  �˷�����_In     In Number := 0, --0-ȫ�� 1-�˹Һŷ� 2-�˲�����
  ��ָ������_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
  �˺�����_In     Number := 1,
  ���㷽ʽ_In     Varchar2 := Null,
  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null,
  �ջ�Ʊ�ݺ�_In   Varchar2 := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null
) As
  --�˷�����_In,��һ�¼�������²�׼���в����˷�
  --    2.�����ӿ�,��ʱ��֧��
  -- �ҺŷѲ����ѷֿ���,����
  --    ��ͨ���㷽ʽ:ԭ���㷽ʽ�˲��ַ���
  --    Ԥ����:Ԥ����,�˲���
  --    Ԥ��������ͨ���㷽ʽ���:�˿����ͨ���㷽ʽ������
  --    ���ѿ�:ԭ�������ò����������ѿ�
  --��ԭ���˽���_In:ָ�����˻���ԭ�����㷽ʽ(��ҽ���ĸ����˻�,�����˻������ֵ�),����ö�����
  --��ָ������_IN:ָ��ԭ���˽��㲿��,Ӧ���˸����ֽ��㷽ʽ,Ϊ��ʱȱʡ�˸��ֽ�,�����˸�ָ���Ľ��㷽ʽ

  --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ���
  Cursor c_Registinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select a.����ʱ��, a.�Ǽ�ʱ��, c.����ʱ��, a.�շ�ϸĿid As ��Ŀid, c.ִ�в���id As ����id, c.ִ���� As ҽ������, d.Id As ҽ��id, c.�ű� As ����
    From ������ü�¼ A, ���˹Һż�¼ C, ��Ա�� D
    Where a.��¼���� = 4 And a.No = ���ݺ�_In And a.No = c.No And a.��¼״̬ = v_״̬ And c.ִ���� = d.����(+) And Rownum < 2;
  r_Registrow c_Registinfo%RowType;

  --���α������жϼ�¼�Ƿ����,�����û��ܱ���
  Cursor c_Moneyinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select ���˿���id, ��������id, ִ�в���id, ������Ŀid, Nvl(Sum(Ӧ�ս��), 0) As Ӧ��, Nvl(Sum(ʵ�ս��), 0) As ʵ��, Nvl(Sum(���ʽ��), 0) As ����
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = v_״̬ And NO = ���ݺ�_In
    Group By ���˿���id, ��������id, ִ�в���id, ������Ŀid;
  r_Moneyrow c_Moneyinfo%RowType;

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
  Cursor c_Opermoney(n_Id ����Ԥ����¼.����id%Type) Is
    Select Distinct b.���㷽ʽ, -1 * Nvl(b.��Ԥ��, 0) As ��Ԥ��
    From ����Ԥ����¼ B
    Where b.����id = n_Id And b.��¼���� = 4 And b.��¼״̬ = 2 And Nvl(b.��Ԥ��, 0) <> 0;

  v_Err_Msg Varchar(255);
  Err_Item Exception;

  n_����id ����Ԥ����¼.����id%Type;
  n_����id ������ü�¼.����id%Type;

  v_��ָ�����㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_�˿���       ����Ԥ����¼.��Ԥ��%Type;
  n_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;
  n_����id         ������Ϣ.����id%Type;
  n_�˷ѽ��       ����Ԥ����¼.��Ԥ��%Type;
  n_Ԥ�����       ����Ԥ����¼.��Ԥ��%Type; --ԭ��¼ Ԥ���ɿ���
  n_����ֵ         �������.Ԥ�����%Type;
  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_��id           ����ɿ����.Id%Type;

  n_�����˷�       Number; --��¼�Ƿ��Ǵ˵��ݵĵڶ����˷�
  n_����̨ǩ���Ŷ� Number;
  n_ԤԼ���ɶ���   Number;
  n_ԤԼ�Һ�       Number;
  n_�Һ����ɶ���   Number;
  d_Date           Date;
  n_����           ������ü�¼.���ʷ���%Type;
  n_����id1        ������Ϣ.����id%Type;
  n_���ض�         ������ü�¼.ʵ�ս��%Type;
  n_�ѽ���         Number;
  n_���           ���˹Һż�¼.����%Type;
  n_���ﲡ��id     ������Ϣ.����id%Type;
  d_����ʱ��       ����ǼǼ�¼.����ʱ��%Type;
  n_�����¼id     �ٴ������¼.Id%Type;
  v_��������       Varchar2(5000);
  v_��ǰ����       Varchar2(1000);
  v_����ids        Varchar2(500);
  v_Temp           Varchar2(500);
  v_���㷽ʽ       ����Ԥ����¼.���㷽ʽ%Type;
  n_��������־     Number;
  n_������       ����Ԥ����¼.��Ԥ��%Type;
  n_�����         Number;
Begin
  n_��id           := Zl_Get��id(����Ա����_In);
  v_��ָ�����㷽ʽ := ��ָ������_In;

  Select �����¼id, ���� Into n_�����¼id, n_��� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum < 2;

  --�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ����
  Open c_Moneyinfo(1);
  Fetch c_Moneyinfo
    Into r_Moneyrow;
  If c_Moneyinfo%RowCount = 0 Then
    Close c_Moneyinfo;
    Open c_Moneyinfo(0);
    Fetch c_Moneyinfo
      Into r_Moneyrow;
    If c_Moneyinfo%RowCount = 0 Then
      v_Err_Msg := 'Ҫ����ĵ��ݲ����ڡ�';
      Raise Err_Item;
    End If;
    n_ԤԼ�Һ� := 1;
  End If;
  Close c_Moneyinfo;

  Begin
    Select Zl_Fun_Regcustomname Into v_Temp From Dual;
    If v_Temp Is Not Null Then
      v_����ids := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    End If;
  Exception
    When Others Then
      v_����ids := Null;
  End;

  --1.ԤԼ����
  If Nvl(n_ԤԼ�Һ�, 0) = 1 Then
    --������Լ��
    Open c_Registinfo(0);
    Fetch c_Registinfo
      Into r_Registrow;
  
    n_����� := Null;
    Update �ٴ������¼ Set ��Լ�� = Nvl(��Լ��, 0) - 1 Where ID = n_�����¼id Returning ��Լ�� Into n_�����;
    If Nvl(n_�����, 0) < 0 Then
      Update �ٴ������¼ Set ��Լ�� = 0 Where ID = n_�����¼id;
    End If;
  
    Update ���˹ҺŻ���
    Set ��Լ�� = Nvl(��Լ��, 0) - 1
    Where ���� = Trunc(r_Registrow.����ʱ��) And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
          Nvl(ҽ������, '-') = Nvl(r_Registrow.ҽ������, '-') And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
          (���� = r_Registrow.���� Or ���� Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
      Values
        (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
         Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1);
    End If;
  
    Close c_Registinfo;
  
    --���¹Һ����״̬
    Update �ٴ�������ſ���
    Set �Һ�״̬ = 0, ����Ա���� = Null
    Where �Һ�״̬ = 2 And ��¼id = n_�����¼id And ��� = n_���;
  
    Update �ٴ�������ſ���
    Set �Һ�״̬ = 4, ����Ա���� = Null
    Where �Һ�״̬ = 2 And ��¼id = n_�����¼id And ��ע = To_Char(n_���);
  
    --��Ӳ��˹Һż�¼�� ������¼
    Select ���˹Һż�¼_Id.Nextval, Sysdate Into n_�Һ�id, d_Date From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼״̬ = 1 And ��¼���� = 2;
    If Sql%NotFound Then
      v_Err_Msg := 'ԤԼ����' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ���ȡ��ԤԼ';
      Raise Err_Item;
    End If;
  
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ, �����¼id, ԤԼ����Ա, ԤԼ����Ա���)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ,
             n_�����¼id, ԤԼ����Ա, ԤԼ����Ա���
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
  
    Update ������ü�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 0;
    Insert Into ������ü�¼
      (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����,
       �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id,
       ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, �ɿ���id, ����״̬, ��ת��,
       �Һ�id, ��ҳid)
      Select ���˷��ü�¼_Id.Nextval, ��¼����, NO, ʵ��Ʊ��, 2, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���ʽ,
             ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, -1 * ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, -1 * Ӧ�ս��,
             -1 * ʵ�ս��, ������, ��������id, ������, ����ʱ��, d_Date, ִ�в���id, ִ����, -1, ִ��ʱ��, ����, ����Ա���_In, ����Ա����_In, Null, Null,
             ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, �ɿ���id, ����״̬, ��ת��, �Һ�id, ��ҳid
      From ������ü�¼
      Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3;
  
    --���ԤԼ���ɶ���ʱ��Ҫ�������
  
    n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    If Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
      --Ҫɾ������
      For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
        Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
      End Loop;
    End If;
    Return;
  End If;

  Select Nvl(���ʷ���, 0), ����id, Decode(Sign(Nvl(����id, 0)), 0, 0, 1)
  Into n_����, n_����id, n_�ѽ���
  From ������ü�¼
  Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ In (1, 3) And Rownum < 2;

  --2.�ҺŴ���
  n_�ѽ��� := Nvl(n_�ѽ���, 0);

  If n_�ѽ��� = 1 And n_���� = 1 Then
    Select Sysdate, Null Into d_Date, n_����id From Dual;
  Else
    Select Sysdate, ���˽��ʼ�¼_Id.Nextval Into d_Date, n_����id From Dual;
  End If;

  ----0-ȫ�� 1-�˹Һŷ� 2-�˲�����
  If Nvl(�˷�����_In, 0) <> 2 Then
    --���ǹ��˲�����ʱ����
    --���¹Һ����״̬
    If �˺�����_In = 1 Then
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 0, ����Ա���� = Null
      Where �Һ�״̬ = 1 And ��¼id = n_�����¼id And ��� = n_���;
    
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 4, ����Ա���� = Null
      Where �Һ�״̬ = 1 And ��¼id = n_�����¼id And ��ע = To_Char(n_���);
    Else
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 4, ����Ա���� = ����Ա����_In
      Where �Һ�״̬ = 1 And ��¼id = n_�����¼id And (��� = n_��� Or ��ע = To_Char(n_���));
    End If;
  
    --���˾���״̬
    If n_����id Is Not Null Then
      Update ������Ϣ Set ����״̬ = 0, �������� = Null Where ����id = n_����id;
    
      --ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ��
      If ɾ�������_In = 1 Then
        Delete ���ﲡ����¼ Where ����id = n_����id;
        Update ������Ϣ Set ����� = Null Where ����id = n_����id;
        --���ü�¼�����Һż����������￨����,�Լ����˽��Ѻ��˷ѻ����ʵķ���,�Һż�¼�������
        Update ������ü�¼ Set ��ʶ�� = Null Where �����־ = 1 And ����id = n_����id;
      End If;
    End If;
  
    --�����ʱ���˾��￨��,�˷�ʱ������￨��,�ڷǹ��˲�����ʱ
    n_����id1 := Null;
    Begin
      Select ����id
      Into n_����id1
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And ���ӱ�־ = 2 And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If n_����id1 Is Not Null And Nvl(�˷�����_In, 0) <> 2 Then
      Update ������Ϣ
      Set ���￨�� = Null, ����֤�� = Null, Ic���� = Decode(Ic����, ���￨��, Null, Ic����)
      Where ����id = n_����id1;
    End If;
  
  End If;

  --���ǰ���Ƿ��Ѿ������˹�����
  Begin
    Select 1 Into n_�����˷� From ������ü�¼ Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ = 3 And Rownum < 2;
  Exception
    When Others Then
      n_�����˷� := 0;
  End;

  --������ü�¼
  --������¼
  If Nvl(�˷�����_In, 0) = 0 Or Nvl(�˷�����_In, 0) = 2 Then
    --ȫ��,�˲�����
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 1 Or Nvl(�˷�����_In, 0) = 4 Then
    --�˹Һŷ�,�˹Һ��벡����
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 3 Then
    --�˸��ӷ�
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 5 Then
    --�˹Һ��븽�ӷ�
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
    End If;
  End If;

  n_����id := 0;
  If n_���� = 0 Then
    --��ȡ����ID
    Select Nvl(����id, 0)
    Into n_����id
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And
          Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0)) And
          Rownum = 1;
  End If;

  If n_���� = 1 Then
    --����
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And
                       Nvl(���ӱ�־, 0) =
                       Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0)) And
                       Nvl(���ʷ���, 0) = 1) Loop
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
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
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
    --1.�˷�
    --���˹ҺŽ���:�ֽ�͸����ʻ�����
    If ���㷽ʽ_In Is Null Then
      If ��ԭ���˽���_In Is Not Null Then
        --�˿����ȡ
        If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
          --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ��
          Begin
            --��ȡ�����˿���
            Select Sum(Nvl(ʵ�ս��, 0)) As �տ���
            Into n_�˿���
            From ������ü�¼
            Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3 And
                  Nvl(���ӱ�־, 0) =
                  Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
          
          Exception
            When Others Then
              v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                             When 1 Then
                              '�Һŷ���'
                             When 2 Then
                              '������'
                           End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
              Raise Err_Item;
          End;
          Begin
            Select ��Ԥ��
            Into n_�˷ѽ��
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
          Exception
            When Others Then
              n_�˷ѽ�� := 0;
          End;
        
          --a.����Ľ��㷽ʽ
        
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -n_�˿���,
                   n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
        
          If n_�˷ѽ�� = 0 Then
            --b.����������ֽ�
            If n_�˿��� <> 0 Then
              If v_��ָ�����㷽ʽ Is Null Then
                --�˸��ֽ�
                Begin
                  Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
                Exception
                  When Others Then
                    v_��ָ�����㷽ʽ := '�ֽ�';
                End;
              End If;
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˿���), ����˵�� = Nvl(����˵��_In, ����˵��)
              Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
                   �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
                  Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                         ����Ա���_In, ����Ա����_In, -1 * n_�˿���, n_����id, n_��id, Ԥ�����, Decode(����˵��_In, Null, �����id, Null), ���㿨���,
                         Decode(����˵��_In, Null, ����, Null), Decode(����˵��_In, Null, ������ˮ��, Null), Nvl(����˵��_In, ����˵��), ������λ,
                         4
                  From ����Ԥ����¼ A
                  Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
              End If;
            End If;
          End If;
        Else
          --a.����Ľ��㷽ʽԭ����
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -��Ԥ��,
                   n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
        
          --b.����������ֽ�
          Begin
            Select Sum(��Ԥ��)
            Into n_�˷ѽ��
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') > 0;
          Exception
            When Others Then
              n_�˷ѽ�� := 0;
          End;
          If n_�˷ѽ�� <> 0 Then
            If v_��ָ�����㷽ʽ Is Null Then
              --�˸��ֽ�
              Begin
                Select ���㷽ʽ
                Into v_��ָ�����㷽ʽ
                From ����Ԥ����¼
                Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And
                      Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
              
              Exception
                When Others Then
                  Begin
                    Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
                  Exception
                    When Others Then
                      v_��ָ�����㷽ʽ := '�ֽ�';
                  End;
              End;
            End If;
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˷ѽ��), ����˵�� = Nvl(����˵��_In, ����˵��)
            Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
                 �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
                Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                       ����Ա���_In, ����Ա����_In, -1 * n_�˷ѽ��, n_����id, n_��id, Ԥ�����, Decode(����˵��_In, Null, �����id, Null), ���㿨���,
                       Decode(����˵��_In, Null, ����, Null), Decode(����˵��_In, Null, ������ˮ��, Null), Nvl(����˵��_In, ����˵��), ������λ, 4
                From ����Ԥ����¼ A
                Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
            End If;
          End If;
        End If;
      Else
        --�˿����ȡ
        If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
          --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ��
          Begin
            --��ȡ�����˿���
            Select Sum(Nvl(ʵ�ս��, 0)) As �տ���
            Into n_�˿���
            From ������ü�¼
            Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3 And
                  Nvl(���ӱ�־, 0) =
                  Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
          Exception
            When Others Then
              v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                             When 1 Then
                              '�Һŷ���'
                             When 2 Then
                              '������'
                           End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
              Raise Err_Item;
          End;
        End If;
        If Nvl(n_�����˷�, 0) = 0 And Nvl(�˷�����_In, 0) = 0 Then
          --�״�ȫ��
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id;
        Else
          --�����˷�,���߱��ε���һ����
          --�����˷�ʱ,��¼״̬=3 ,�״β�����,��¼״̬Ϊ1
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ժҪ = 'ҽ���Һ�' And
                  ��Ԥ�� = n_�˿��� And Rownum < 2;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ��Ԥ�� = n_�˿��� And
                    Rownum < 2;
          End If;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And Rownum < 2;
            If Sql%RowCount = 0 Then
              --�����˷�,����ȫ��ʹ��Ԥ����ɷ�ʱ�Ŵ��ڴ������
              n_Ԥ����� := n_�˿���;
            End If;
          End If;
        
        End If;
      End If;
    Else
      --�����㷽ʽ��
      v_�������� := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
      While v_�������� Is Not Null Loop
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
      
        v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_��������־ := To_Number(v_��ǰ����);
      
        If n_��������־ = 0 Then
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������, �������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, v_���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * n_������, n_����id, n_��id, Ԥ�����, Null, Null, Null, Null, ����˵��_In, ������λ, 4, �������
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And Rownum < 2;
        Else
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������, �������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, v_���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * n_������, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, Nvl(����˵��_In, ����˵��), ������λ, 4, �������
            
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And
                  (�����id Is Not Null Or ���㿨��� Is Not Null) And Rownum < 2;
        End If;
        v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
      End Loop;
      n_Ԥ����� := Nvl(��Ԥ��_In, 0);
    End If;
    --�״��˷�ʱ,��¼״̬�����Ϊ��3
    Update ����Ԥ����¼
    Set ��¼״̬ = 3
    Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id;
  
    --��Ԥ�� 1-ȫ�� 2-������,������ʱ��ȫ��ʹ��Ԥ�����нɿ�
    If Nvl(�˷�����_In, 0) = 0 Or (Nvl(�˷�����_In, 0) <> 0 And n_Ԥ����� <> 0) Then
      --���˹ҺŽ���:��Ԥ�����
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_Date,
               ����Ա����_In, ����Ա���_In, -1 * Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, ��Ԥ��, n_Ԥ�����), n_����id, n_��id, Ԥ�����,
               �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
        From ����Ԥ����¼
        Where ��¼���� In (1, 11) And ����id = n_����id And Nvl(��Ԥ��, 0) <> 0 And
              Rownum = Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, Rownum, 1);
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
  
    If �ջ�Ʊ�ݺ�_In Is Not Null Then
      --���˹Һŷ�,������Ʊ��
      --�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�)
      Begin
        --�����һ�δ�ӡ��������ȡ
        Select ID
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 4 And b.No = ���ݺ�_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          n_��ӡid := Null;
      End;
    
      --���ջ�ԭƱ��
      If n_��ӡid Is Not Null Then
        Begin
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
        Exception
          When Others Then
            Delete From Ʊ��ʹ����ϸ Where ��ӡid = n_��ӡid And ���� = 2 And ԭ�� = 2;
            Insert Into Ʊ��ʹ����ϸ
              (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
              Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
              From Ʊ��ʹ����ϸ
              Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
        End;
      End If;
    End If;
  End If;

  --�����˲�������,��������ܼ�¼
  --��ػ��ܱ�Ĵ���

  --���˹ҺŻ���
  Open c_Registinfo(3);
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
      Select Decode(ԤԼ, Null, 0, 0, 0, 1) Into n_ԤԼ�Һ� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1;
    Exception
      When Others Then
        n_ԤԼ�Һ� := 0;
    End;
    n_����� := Null;
    Update �ٴ������¼
    Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
    Where ID = n_�����¼id
    Returning �ѹ��� Into n_�����;
  
    If Nvl(n_�����, 0) < 0 Then
      Update �ٴ������¼ Set �ѹ��� = 0 Where ID = n_�����¼id;
    End If;
  
    Update ���˹ҺŻ���
    Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
    Where ���� = Trunc(r_Registrow.����ʱ��) And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
          Nvl(ҽ������, '-') = Nvl(r_Registrow.ҽ������, '-') And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
          (���� = r_Registrow.���� Or ���� Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���, ��Լ��)
      Values
        (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
         Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1 * n_ԤԼ�Һ�, -1 * n_ԤԼ�Һ�);
    End If;
  
    Close c_Registinfo;
  End If;

  If n_���� = 0 Then
    --��Ա�ɿ����(���������ʻ��ȵĽ�����,�����˳�Ԥ����)
    For r_Opermoney In c_Opermoney(n_����id) Loop
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + (-1 * r_Opermoney.��Ԥ��)
      Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1
      Returning ��� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_Opermoney.���㷽ʽ, 1, -1 * r_Opermoney.��Ԥ��);
        n_����ֵ := r_Opermoney.��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
      End If;
    End Loop;
  End If;
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
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
    Begin
      Select ����id, ����ʱ�� Into n_���ﲡ��id, d_����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In;
      Delete From ����ǼǼ�¼ Where ����id = n_���ﲡ��id And ����ʱ�� = d_����ʱ�� And ��ҳid Is Null;
    Exception
      When Others Then
        Null;
    End;
  
    --���˹Һż�¼
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼���� = 1 And ��¼״̬ = 1;
    If Sql%NotFound Then
      v_Err_Msg := '�Һŵ���' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ����˺�';
      Raise Err_Item;
    End If;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ, �����¼id)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ,
             n_�����¼id
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
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
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_����_Delete;
/

--127075:����,2018-07-06,����ʹ�õ��˲��Ų�������̨ǩ���Ŷӵ�Oracle����
Create Or Replace Procedure Zl_ԤԼ�ҺŽ���_����_Insert
(
  No_In            ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,

  ����id_In        ������ü�¼.����id%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���㷽ʽ_In      Varchar2, --�ֽ�Ľ�������
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱʹ�õ�Ԥ�����
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  ����_In          �Һ����״̬.���%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ���ɶ���_In      Number := 0,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  ��Ԥ������ids_In Varchar2 := Null,
  ��������_In      Number := 0,
  ���½������_In  Number := 1, --�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�����������
  ժҪ_In          ���˹Һż�¼.ժҪ%Type := Null,
  �շѵ�_In        ���˹Һż�¼.�շѵ�%Type := Null
) As
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

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  Err_Special Exception;
  v_����Ա���� ���˹Һż�¼.������%Type;
  v_�ֽ�       ���㷽ʽ.����%Type;
  v_�����ʻ�   ���㷽ʽ.����%Type;
  v_��������   �ŶӽкŶ���.��������%Type;
  v_�ű�       ������ü�¼.���㵥λ%Type;
  v_����       ������ü�¼.��ҩ����%Type;
  v_�ŶӺ���   �ŶӽкŶ���.�ŶӺ��� %Type;
  v_ԤԼ��ʽ   ���˹Һż�¼.ԤԼ��ʽ %Type;

  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_����̨ǩ���Ŷ� Number;
  n_��id           ����ɿ����.Id%Type;
  n_Count          Number(18);
  n_�Ŷ�           Number;
  n_�����Ŷ�       Number;
  n_��ǰ���       ����Ԥ����¼.���%Type;
  n_Ԥ��id         ����Ԥ����¼.Id%Type;

  d_Date         Date;
  d_ԤԼʱ��     ������ü�¼.����ʱ��%Type;
  d_����ʱ��     Date;
  d_�Ŷ�ʱ��     Date;
  n_ʱ��         Number := 0;
  n_����         Number := 0;
  v_��������     Varchar2(2000);
  v_��ǰ����     Varchar2(500);
  n_������     ����Ԥ����¼.��Ԥ��%Type;
  v_�������     ����Ԥ����¼.�������%Type;
  v_���㷽ʽ     ����Ԥ����¼.���㷽ʽ%Type;
  n_��������־   Number(3);
  v_�Ŷ����     �ŶӽкŶ���.�Ŷ����%Type;
  n_����ģʽ     ������Ϣ.����ģʽ%Type;
  v_���ʽ     ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  n_����ģʽ     Number := 0;
  n_�����¼id   ���˹Һż�¼.�����¼id%Type;
  n_�³����¼id ���˹Һż�¼.�����¼id%Type;
  n_��Դid       �ٴ������¼.��Դid%Type;
  n_ԤԼ˳���   �ٴ�������ſ���.ԤԼ˳���%Type;
  n_�ɷ�ʱ��     �ٴ������¼.�Ƿ��ʱ��%Type;
  n_����ſ���   �ٴ������¼.�Ƿ���ſ���%Type;
  n_�ɿ���id     �ٴ������¼.����id%Type;
  n_����Ŀid     �ٴ������¼.��Ŀid%Type;
  n_��ҽ��id     �ٴ������¼.ҽ��id%Type;
  n_�Һ�ģʽ     Number(3);
  d_����ʱ��     Date;
  v_Paratemp     Varchar2(500);
  v_Registtemp   Varchar2(500);
  n_���         Number(3);
  n_��ſ���     �ٴ������¼.�Ƿ���ſ���%Type;
  v_���ϰ�ʱ��   �ٴ������¼.�ϰ�ʱ��%Type;
Begin
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
  v_Paratemp      := Nvl(zl_GetSysParameter('�Һ��Ű�ģʽ'), 0);
  n_����ģʽ      := Nvl(zl_GetSysParameter('ԤԼ����ģʽ', 1111), 0);
  n_�Һ�ģʽ      := To_Number(Substr(v_Paratemp, 1, 1));
  If n_�Һ�ģʽ = 1 Then
    Begin
      d_����ʱ�� := To_Date(Substr(v_Paratemp, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
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
  If �Ǽ�ʱ��_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := �Ǽ�ʱ��_In;
  End If;

  --���¹Һ����״̬
  Begin
    Select �ű�, ����, Trunc(����ʱ��), ����ʱ��, ԤԼ��ʽ, �����¼id
    Into v_�ű�, v_����, d_ԤԼʱ��, d_����ʱ��, v_ԤԼ��ʽ, n_�����¼id
    From ���˹Һż�¼
    Where ��¼���� = 2 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      Select Max(������) Into v_����Ա���� From ���˹Һż�¼ Where ��¼���� = 2 And ��¼״̬ In (1, 3) And NO = No_In;
      If v_����Ա���� Is Null Then
        v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�ȡ��';
        Raise Err_Item;
      Else
        If v_����Ա���� = ����Ա����_In Then
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�����';
          Raise Err_Special;
        Else
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ������˽���';
          Raise Err_Item;
        End If;
      End If;
  End;

  --�ж��Ƿ��ʱ��
  Select Nvl(�Ƿ��ʱ��, 0), ��Դid, Nvl(�Ƿ���ſ���, 0)
  Into n_ʱ��, n_��Դid, n_��ſ���
  From �ٴ������¼
  Where ID = n_�����¼id;

  If n_ʱ�� = 1 And ��������_In = 0 And n_����ģʽ = 0 Then
    If Trunc(����ʱ��_In) <> Trunc(Sysdate) Then
      v_Err_Msg := '��ʱ�ε�ԤԼ�Һŵ�ֻ�ܵ�����գ�';
      Raise Err_Item;
    End If;
  End If;

  If n_ʱ�� = 0 And ��������_In = 0 Then
    If n_����ģʽ = 0 Then
      If Trunc(����ʱ��_In) = Trunc(Sysdate) Then
        d_����ʱ�� := ����ʱ��_In;
      Else
        d_����ʱ�� := Sysdate;
      End If;
    Else
      d_����ʱ�� := ����ʱ��_In;
    End If;
  Else
    If Not ����ʱ��_In Is Null Then
      d_����ʱ�� := ����ʱ��_In;
    End If;
  End If;

  If d_����ʱ�� Is Not Null Then
    If d_����ʱ�� < d_����ʱ�� Then
      v_Err_Msg := '��ǰԤԼ�Һŵ����ڳ�����Ű�ģʽ���ţ�������' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '֮ǰ����!';
      Raise Err_Item;
    End If;
  End If;

  If Not v_���� Is Null Then
    If ����_In Is Null Then
      Update �ٴ�������ſ��� Set �Һ�״̬ = 0 Where (��� = v_���� Or ��ע = v_����) And ��¼id = n_�����¼id;
    Else
      If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
        If n_ʱ�� = 0 And ��������_In = 0 Then
          --��ǰ���ջ��ӳٽ���
          Update �ٴ�������ſ��� Set �Һ�״̬ = 0 Where ��� = v_���� And ��¼id = n_�����¼id;
        
          Select �Ƿ��ʱ��, �Ƿ���ſ���, ����id, ҽ��id, ��Ŀid, �ϰ�ʱ��
          Into n_�ɷ�ʱ��, n_����ſ���, n_�ɿ���id, n_��ҽ��id, n_����Ŀid, v_���ϰ�ʱ��
          From �ٴ������¼
          Where ID = n_�����¼id;
          Begin
            Select ID
            Into n_�³����¼id
            From �ٴ������¼
            Where ��Դid = n_��Դid And �Ƿ��ʱ�� = n_�ɷ�ʱ�� And �Ƿ���ſ��� = n_����ſ��� And ����id = n_�ɿ���id And
                  Nvl(ҽ��id, 0) = Nvl(n_��ҽ��id, 0) And �ϰ�ʱ�� = v_���ϰ�ʱ�� And Nvl(�Ƿ񷢲�, 0) = 1 And �������� = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '���յ���û�ж�Ӧ�ĳ��ﰲ��,�޷�����!';
              Raise Err_Item;
          End;
        
          Begin
            Select 1
            Into n_����
            From �ٴ�������ſ���
            Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
          Exception
            When Others Then
              n_���� := 0;
          End;
        
          If n_���� = 1 Then
            Update �ٴ�������ſ���
            Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
            Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
          Else
            --�����ѱ�ʹ�õ����
            Select Min(���) Into v_���� From �ٴ�������ſ��� Where ��¼id = n_�³����¼id And Nvl(�Һ�״̬, 0) = 0;
            If v_���� Is Null Then
              v_Err_Msg := '���յ���û�п������,�޷�����!';
              Raise Err_Item;
            End If;
            Update �ٴ�������ſ���
            Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
            Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
          End If;
        Else
          Select �Ƿ��ʱ��, �Ƿ���ſ���, ����id, ҽ��id, ��Ŀid, �ϰ�ʱ��
          Into n_�ɷ�ʱ��, n_����ſ���, n_�ɿ���id, n_��ҽ��id, n_����Ŀid, v_���ϰ�ʱ��
          From �ٴ������¼
          Where ID = n_�����¼id;
          Begin
            Select ID
            Into n_�³����¼id
            From �ٴ������¼
            Where ��Դid = n_��Դid And �Ƿ��ʱ�� = n_�ɷ�ʱ�� And �Ƿ���ſ��� = n_����ſ��� And ����id = n_�ɿ���id And
                  Nvl(ҽ��id, 0) = Nvl(n_��ҽ��id, 0) And �ϰ�ʱ�� = v_���ϰ�ʱ�� And Nvl(�Ƿ񷢲�, 0) = 1 And �������� = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '���յ���û�ж�Ӧ�ĳ��ﰲ��,�޷�����!';
              Raise Err_Item;
          End;
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 0, ����Ա���� = ����Ա����_In
          Where (��� = v_���� Or ��ע = v_����) And ��¼id = n_�����¼id And Nvl(�Һ�״̬, 0) = 2
          Returning ԤԼ˳��� Into n_ԤԼ˳���;
        
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In, ԤԼ˳��� = n_ԤԼ˳���
          Where ��� = v_���� And ��¼id = n_�³����¼id And Nvl(�Һ�״̬, 0) = 0;
          If Sql% RowCount = 0 Then
            v_Err_Msg := '���յ������' || v_���� || '�ѱ�������ʹ��,�޷�����.';
            Raise Err_Item;
          End If;
        End If;
      Else
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
        Where (��� = v_���� Or ��ע = v_����) And ��¼id = n_�����¼id;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
          Raise Err_Item;
        End If;
      End If;
    End If;
  Else
    If Not ����_In Is Null Then
      If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
        Select �Ƿ��ʱ��, �Ƿ���ſ���, ����id, ҽ��id, ��Ŀid, �ϰ�ʱ��
        Into n_�ɷ�ʱ��, n_����ſ���, n_�ɿ���id, n_��ҽ��id, n_����Ŀid, v_���ϰ�ʱ��
        From �ٴ������¼
        Where ID = n_�����¼id;
        Begin
          Select ID
          Into n_�³����¼id
          From �ٴ������¼
          Where ��Դid = n_��Դid And �Ƿ��ʱ�� = n_�ɷ�ʱ�� And �Ƿ���ſ��� = n_����ſ��� And ����id = n_�ɿ���id And
                Nvl(ҽ��id, 0) = Nvl(n_��ҽ��id, 0) And �ϰ�ʱ�� = v_���ϰ�ʱ�� And Nvl(�Ƿ񷢲�, 0) = 1 And �������� = Trunc(Sysdate) And
                Rownum < 2;
        Exception
          When Others Then
            v_Err_Msg := '���յ���û�ж�Ӧ�ĳ��ﰲ��,�޷�����!';
            Raise Err_Item;
        End;
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 0, ����Ա���� = ����Ա����_In
        Where (��� = ����_In Or ��ע = ����_In) And ��¼id = n_�����¼id And Nvl(�Һ�״̬, 0) = 2
        Returning ԤԼ˳��� Into n_ԤԼ˳���;
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In, ԤԼ˳��� = n_ԤԼ˳���
        Where ��� = ����_In And ��¼id = n_�³����¼id And Nvl(�Һ�״̬, 0) = 0;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '���յ������' || ����_In || '�ѱ�������ʹ��,�޷�����.';
          Raise Err_Item;
        End If;
      Else
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
        Where (��� = ����_In Or ��ע = ����_In) And ��¼id = n_�����¼id;
      
      End If;
      v_���� := ����_In;
    Else
      v_���� := Null;
    End If;
  End If;

  --����������ü�¼
  Update ������ü�¼
  Set ��¼״̬ = 1, ʵ��Ʊ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, Ʊ�ݺ�_In), ����id = Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In),
      ���ʽ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��), ��ҩ���� = ����_In, ����id = ����id_In, ��ʶ�� = �����_In, ���� = ����_In, ���� = ����_In,
      �Ա� = �Ա�_In, ���ʽ = ���ʽ_In, �ѱ� = �ѱ�_In, ����ʱ�� = d_����ʱ��, �Ǽ�ʱ�� = d_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In,
      �ɿ���id = n_��id, ���ʷ��� = Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ժҪ = Decode(�շѵ�_In, Null, Nvl(ժҪ_In, ժҪ), '����:' || �շѵ�_In)
  Where ��¼���� = 4 And ��¼״̬ = 0 And NO = No_In;

  v_Registtemp := zl_GetSysParameter('�Һ��Ű�ģʽ');
  If Substr(v_Registtemp, 1, 1) = 1 Then
    Begin
      If To_Date(Substr(v_Registtemp, 3), 'yyyy-mm-dd hh24:mi:ss') > d_����ʱ�� Then
        v_Err_Msg := '����ʱ��' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || 'δ���ó�����Ű�ģʽ,Ŀǰ�޷�����!';
        Raise Err_Item;
      End If;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select 1
      Into n_���
      From �ٴ������¼
      Where ID = Nvl(n_�³����¼id, n_�����¼id) And d_����ʱ�� Between ͣ�￪ʼʱ�� And ͣ����ֹʱ��;
    Exception
      When Others Then
        n_��� := 0;
    End;
    If n_��� = 1 And Not (n_ʱ�� = 1 And n_��ſ��� = 1) Then
      v_Err_Msg := '����ʱ��' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '�İ����Ѿ���ͣ��,�޷�����!';
      Raise Err_Item;
    End If;
  End If;

  --���˹Һż�¼
  Update ���˹Һż�¼
  Set ������ = ����Ա����_In, ����ʱ�� = d_Date, ��¼���� = 1, ����id = ����id_In, ����� = �����_In, ����ʱ�� = d_����ʱ��, ���� = ����_In, �Ա� = �Ա�_In,
      ���� = ����_In, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, ���� = Decode(Nvl(����_In, 0), 0, Null, ����_In), ���� = v_����, ���� = ����_In,
      �����¼id = Nvl(n_�³����¼id, n_�����¼id), ժҪ = Nvl(ժҪ_In, ժҪ), �շѵ� = �շѵ�_In
  Where ��¼״̬ = 1 And NO = No_In And ��¼���� = 2
  Returning ID Into n_�Һ�id;
  If Sql%NotFound Then
    Begin
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ժҪ, ����, ԤԼ, ԤԼ��ʽ, ������, ����ʱ��, ԤԼʱ��, ����, ҽ�Ƹ��ʽ, �����¼id, �շѵ�)
        Select n_�Һ�id, No_In, 1, 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, ���㵥λ, �Ӱ��־, ����_In, Null, ִ�в���id, ִ����, 0, Null,
               �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����, Nvl(ժҪ_In, ժҪ), v_����, 1, Substr(����, 1, 10) As ԤԼ��ʽ, ����Ա����_In,
               Nvl(�Ǽ�ʱ��_In, Sysdate), ����ʱ��, Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, Nvl(n_�³����¼id, n_�����¼id),
               �շѵ�_In
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '���ڲ���ԭ��,���ݺ�Ϊ��' || No_In || '���Ĳ���' || ����_In || '�Ѿ�������';
        Raise Err_Item;
    End;
  End If;

  --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
  If Nvl(���ɶ���_In, 0) <> 0 Then
    For v_�Һ� In (Select ID, ����, ����, ִ����, ִ�в���id, ����ʱ��, �ű�, ���� From ���˹Һż�¼ Where NO = No_In) Loop
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(v_�Һ�.ִ�в���id, 0)));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
        Begin
          Select 1,
                 Case
                   When �Ŷ�ʱ�� < Trunc(Sysdate) Then
                    1
                   Else
                    0
                 End
          Into n_�Ŷ�, n_�����Ŷ�
          From �ŶӽкŶ���
          Where ҵ������ = 0 And ҵ��id = v_�Һ�.Id And Rownum <= 1;
        Exception
          When Others Then
            n_�Ŷ� := 0;
        End;
        If n_�Ŷ� = 0 Then
          --��������
          --����ִ�в��š���������
          n_�Һ�id   := v_�Һ�.Id;
          v_�������� := v_�Һ�.ִ�в���id;
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, n_�Һ�id, v_�Һ�.�ű� || '|' || v_�Һ�.����);
          v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        
          --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
          d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, v_�Һ�.�ű�, v_�Һ�.����, v_�Һ�.����ʱ��);
          --   ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In,
          Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, v_�Һ�.ִ�в���id, v_�ŶӺ���, Null, ����_In, ����id_In, v_�Һ�.����, v_�Һ�.ִ����, d_�Ŷ�ʱ��,
                           v_ԤԼ��ʽ, Null, v_�Ŷ����);
        Elsif Nvl(n_�����Ŷ�, 0) = 1 Then
          --���¶��к�
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, v_�Һ�.Id, v_�Һ�.�ű� || '|' || Nvl(v_�Һ�.����, 0));
          v_�Ŷ���� := Zlgetsequencenum(0, v_�Һ�.Id, 1);
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����, v_�ŶӺ���, v_�Ŷ����);
        
        Else
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����);
        End If;
      End If;
    End Loop;
  End If;

  --���ܽ��㵽����Ԥ����¼
  If Nvl(���ʷ���_In, 0) = 0 Then
    If Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
         ��������)
      Values
        (n_Ԥ��id, 4, 1, No_In, Decode(����id_In, 0, Null, ����id_In), v_�ֽ�, 0, �Ǽ�ʱ��_In, ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�',
         n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, Null, 4);
    End If;
    If Nvl(�ֽ�֧��_In, 0) <> 0 Then
      v_�������� := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
      While v_�������� Is Not Null Loop
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        v_������� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_��������־ := To_Number(v_��ǰ����);
      
        If n_��������־ = 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, No_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, Null, Null, Null, Null, Null, Null, 4, v_�������);
        Else
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, No_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, Null, 4, v_�������);
        
          If Nvl(���㿨���_In, 0) <> 0 And Nvl(n_������, 0) <> 0 Then
            Zl_���˿������¼_֧��(���㿨���_In, ����_In, 0, n_������, n_Ԥ��id, ����Ա���_In, ����Ա����_In, �Ǽ�ʱ��_In);
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
  End If;

  --���ھ��￨ͨ��Ԥ����Һ�
  If Nvl(Ԥ��֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
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
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �Ǽ�ʱ��_In,
               ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, ����id_In, 4
        From ����Ԥ����¼
        Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
      Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (r_Deposit.����id, Nvl(1, 2), -1 * n_��ǰ���, 1);
        n_����ֵ := -1 * n_��ǰ���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
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
  End If;

  --����ҽ���Һ�
  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
       Ԥ�����, �������, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 4, 1, No_In, ����id_In, v_�����ʻ�, ����֧��_In, d_Date, ����Ա���_In, ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id,
       Null, Null, Null, Null, Null, Null, Null, ����id_In, 4);
  End If;

  --��ػ��ܱ�Ĵ���
  --��Ա�ɿ����
  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 And Nvl(���½������_In, 1) = 1 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ����֧��_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
      n_����ֵ := ����֧��_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ���� Where �տ�Ա = ����Ա����_In And ���� = 1 And Nvl(���, 0) = 0;
    End If;
  End If;

  If Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = No_In And Nvl(���ʷ���, 0) = 1) Loop
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (����id_In, 1, 1, Nvl(c_����.ʵ�ս��, 0), 0);
      End If;
    
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (����id_In, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
  End If;
  If Nvl(����id_In, 0) <> 0 Then
    n_����ģʽ := 0;
    Update ������Ϣ
    Set ����ʱ�� = d_����ʱ��, ����״̬ = 1, �������� = ����_In
    Where ����id = ����id_In
    Returning Nvl(����ģʽ, 0) Into n_����ģʽ;
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
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
  End If;

  --���˵�����Ϣ
  If ����id_In Is Not Null Then
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Nvl(��Ժ, 0) = 0 And Exists
     (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = d_Date
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, d_Date) >= d_Date;
    End If;
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 1, n_�Һ�id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ԤԼ�ҺŽ���_����_Insert;
/

--127075:����,2018-07-06,����ʹ�õ��˲��Ų�������̨ǩ���Ŷӵ�Oracle����
Create Or Replace Procedure Zl_���˹ҺŲ�����_Delete
(
  ���ݺ�_In     ������ü�¼.No%Type,
  ����Ա���_In ������ü�¼.����Ա���%Type,
  ����Ա����_In ������ü�¼.����Ա����%Type,
  ����id_In     ������ü�¼.����id%Type := Null,
  �������_In   ����Ԥ����¼.�������%Type := Null,
  �˺�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
  ɾ�������_In Number := 0
) As
  --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ���

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_����id   ����Ԥ����¼.����id%Type;
  n_����id   ������ü�¼.����id%Type;
  n_������� ����Ԥ����¼.�������%Type;

  n_����id   ������Ϣ.����id%Type;
  n_�˷ѽ�� ����Ԥ����¼.��Ԥ��%Type;
  n_�Һ�id   ���˹Һż�¼.Id%Type;
  n_��id     ����ɿ����.Id%Type;

  n_����̨ǩ���Ŷ� Number;
  n_�Һ����ɶ���   Number;
  d_Date           Date;
  n_����id1        ������Ϣ.����id%Type;
  d_Temp           Date;
  n_���ﲡ��id     ������Ϣ.����id%Type;
  d_����ʱ��       ����ǼǼ�¼.����ʱ��%Type;
Begin
  n_��id := Zl_Get��id(����Ա����_In);
  --�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ����

  Begin
    Select a.����id, a.����id
    Into n_����id, n_����id
    From ������ü�¼ A
    Where a.��¼���� = 4 And a.No = ���ݺ�_In And a.��¼״̬ = 1 And Rownum < 2;
  Exception
  
    When Others Then
      n_����id := -1;
  End;
  If Nvl(n_����id, 0) = -1 Then
    v_Err_Msg := 'δ�ҵ�ָ���ĹҺŵ�:' || ���ݺ�_In || ',�����Ѿ������˺�,�������ٴ��˺š�';
    Raise Err_Item;
  End If;

  --2.�ҺŴ���

  d_Date     := �˺�ʱ��_In;
  n_����id   := ����id_In;
  n_������� := �������_In;

  If d_Date Is Null Then
    d_Date := Sysdate;
  End If;
  If n_����id Is Null Then
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  End If;

  If n_������� Is Null Then
    n_������� := -1 * n_����id;
  End If;
  --���¹Һ����״̬
  If Zl_To_Number(zl_GetSysParameter('�����������Һ�', 1111)) = 1 Then
    Delete �Һ����״̬
    Where ״̬ = 1 And
          (����, ���, ����) = (Select �ű�, ����, Trunc(����ʱ��) From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1) Or
          (����, ���, ����) = (Select �ű�, ����, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1);
  Else
    Update �Һ����״̬
    Set ״̬ = 4
    Where ״̬ = 1 And
          (����, ���, ����) = (Select �ű�, ����, Trunc(����ʱ��) From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1) Or
          (����, ���, ����) = (Select �ű�, ����, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1);
  End If;

  --���˾���״̬
  If n_����id Is Not Null Then
    Update ������Ϣ Set ����״̬ = 0, �������� = Null Where ����id = n_����id;
    --ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ��(����Ҫ������ʾ��ɾ��)
    If ɾ�������_In = 1 Then
      Delete ���ﲡ����¼ Where ����id = n_����id;
      Update ������Ϣ Set ����� = Null Where ����id = n_����id;
      --���ü�¼�����Һż����������￨����,�Լ����˽��Ѻ��˷ѻ����ʵķ���,�Һż�¼�������
      Update ������ü�¼ Set ��ʶ�� = Null Where �����־ = 1 And ����id = n_����id;
    End If;
  End If;

  --�����ʱ���˾��￨��,�˷�ʱ������￨��,�ڷǹ��˲�����ʱ
  n_����id1 := Null;
  Begin
    Select ����id
    Into n_����id1
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And ���ӱ�־ = 2 And Rownum < 2;
  Exception
    When Others Then
      Null;
  End;
  If n_����id1 Is Not Null Then
    Update ������Ϣ
    Set ���￨�� = Null, ����֤�� = Null, Ic���� = Decode(Ic����, ���￨��, Null, Ic����)
    Where ����id = n_����id1;
  End If;

  --������ü�¼
  --������¼
  Insert Into ������ü�¼
    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
     ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
     ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id, ����״̬, ִ��״̬)
    Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
           �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
           ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id, -1 * ���ʽ��, ������Ŀ��, ���մ���id, -1 * ͳ����, ժҪ As ժҪ, ���ӱ�־, ���ձ���, ��������,
           n_��id, 1, 1
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;

  Update ������ü�¼ Set ��¼״̬ = 3 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
  Select Sum(ʵ�ս��) Into n_�˷ѽ�� From ������ü�¼ Where ����id = n_����id;
  Insert Into ����Ԥ����¼
    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �������, �ɿ���id, Ԥ�����, �����id,
     ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, ��������)
    Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, 4, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, Null, d_Date, ����Ա���_In, ����Ա����_In, n_�˷ѽ��,
           n_����id, n_�������, n_��id, Ԥ�����, Null, Null, Null, Null, Null, Null, 1, 4
    From ����Ԥ����¼ A
    Where a.����id = n_����id And Rownum < 2;
  If Sql%NotFound Then
    v_Err_Msg := 'δ�ҵ��Һŵ�Ϊ��' || ���ݺ�_In || '����ԭʼ�Һż�¼!';
    Raise Err_Item;
  End If;
  Update ����Ԥ����¼ Set ��¼״̬ = 3 Where Mod(��¼����, 10) <> 1 And ����id = n_����id;

  --���˹ҺŻ���
  For c_�Һ� In (Select a.�շ�ϸĿid, a.����ʱ��, a.�Ǽ�ʱ��, c.����ʱ��, c.ִ�в���id, c.ִ����, m.Id As ҽ��id, Nvl(c.�ű�, b.����) As ����, a.����id,
                      a.����id, Decode(c.ԤԼ, Null, 0, 0, 0, 1) As ԤԼ
               From ������ü�¼ A, ���˹Һż�¼ C, �ҺŰ��� B, ��Ա�� M
               Where a.��¼���� = 4 And a.����id = n_����id And a.�������� Is Null And c.ִ���� = m.����(+) And a.No = c.No And
                     Nvl(c.�ű�, Nvl(a.���㵥λ, '-')) = b.���� And Nvl(a.���ӱ�־, 0) = 0 And Rownum < 2) Loop
    --�˷ǹҺŷ���,�򲻴�����ܱ�����
  
    If Nvl(c_�Һ�.ԤԼ, 0) <> 0 Then
      d_Temp := Trunc(c_�Һ�.����ʱ��);
    Else
      d_Temp := Trunc(c_�Һ�.����ʱ��);
    End If;
  
    Update ���˹ҺŻ���
    Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - Nvl(c_�Һ�.ԤԼ, 0), ��Լ�� = Nvl(��Լ��, 0) - Nvl(c_�Һ�.ԤԼ, 0)
    Where ���� = d_Temp And ����id = c_�Һ�.ִ�в���id And ��Ŀid = c_�Һ�.�շ�ϸĿid And Nvl(ҽ������, 'ҽ��') = Nvl(c_�Һ�.ִ����, 'ҽ��') And
          Nvl(ҽ��id, 0) = Nvl(c_�Һ�.ҽ��id, 0) And (���� = c_�Һ�.���� Or ���� Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���, ��Լ��)
      Values
        (d_Temp, c_�Һ�.ִ�в���id, c_�Һ�.�շ�ϸĿid, c_�Һ�.ִ����, Decode(c_�Һ�.ҽ��id, 0, Null, c_�Һ�.ҽ��id), c_�Һ�.����, -1,
         -1 * Nvl(c_�Һ�.ԤԼ, 0), -1 * Nvl(c_�Һ�.ԤԼ, 0));
    End If;
  End Loop;

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
  Begin
    Select ����id, ����ʱ�� Into n_���ﲡ��id, d_����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In;
    Delete From ����ǼǼ�¼ Where ����id = n_���ﲡ��id And ����ʱ�� = d_����ʱ�� And ��ҳid Is Null;
  Exception
    When Others Then
      Null;
  End;
  --���˹Һż�¼
  Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;

  Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼���� = 1 And ��¼״̬ = 1;
  If Sql%NotFound Then
    v_Err_Msg := '�Һŵ���' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ����˺�';
    Raise Err_Item;
  End If;

  Insert Into ���˹Һż�¼
    (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����, ����,
     ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ)
    Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��, ����Ա���_In,
           ����Ա����_In, ����, ����, ����, ԤԼ, ժҪ As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ
    From ���˹Һż�¼
    Where NO = ���ݺ�_In And ��¼״̬ = 3;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹ҺŲ�����_Delete;
/

--127075:����,2018-07-06,����ʹ�õ��˲��Ų�������̨ǩ���Ŷӵ�Oracle����
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

--127075:����,2018-07-06,����ʹ�õ��˲��Ų�������̨ǩ���Ŷӵ�Oracle����
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
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id, Min(Decode(��¼����, 1, �տ�ʱ��, Null)) As �տ�ʱ��
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1
     Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By NO, ����id
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), �տ�ʱ��;
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

  n_��� := ����_In;
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

  --��ʱ�ιҺ�ʱ�ж��Ƿ��ǹ��ڹҺ� Ҳ����׷�Ӻŵ���� Ŀǰֻ���ר�Һŷ�ʱ�ν��д���
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 And Nvl(n_��ſ���, 0) = 1 And ����_In Is Null And Nvl(��������_In, 0) = 0 Then
    --����ʱ��_in>Sysdate ����ʱ��>����ʱ��ʱ��--����_in is null
    Begin
      Select Max(To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_������ʱ��
      From �ҺŰ���ʱ��
      Where ����id = n_����id And ���� = v_���� And Nvl(��������, 0) <> 0;
      n_׷�Ӻ� := Case Sign(����ʱ��_In - d_������ʱ��)
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
  d_ʱ��ʱ�� := ����ʱ��_In;

  If ���_In = 1 And Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 Then
    --�Һ�ʱ��� �Ƿ����ʱ��,����ʱ��,��ʱ�ε�����������ȡ����
    Begin
      Select Nvl(���, 0),
             To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������)
      Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
      From �ҺŰ���ʱ��
      Where ����id = n_����id And ���� = v_���� And
            (���, ����id, ����) In (Select Nvl(Max(���), -1), ����id, ����
                               From �ҺŰ���ʱ��
                               Where ����id = n_����id And ���� = v_���� And
                                     Decode(��������_In + n_׷�Ӻ�, 0, To_Char(����ʱ��_In, 'hh24:mi'),
                                            To_Char(Nvl(��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')) =
                                     To_Char(Nvl(��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
                               Group By ����id, ����);
    Exception
      When Others Then
        n_ʱ����� := -1;
        n_��ʱ��   := 0;
        d_ʱ��ʱ�� := ����ʱ��_In;
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
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(c.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ʱ��ʱ��,
               ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �ҺŰ���ʱ�� C
        Where ����id = n_����id And ���� = v_���� And
              (���, ����id, ����) In
              (Select Nvl(Max(c.���), -1), ����id, ����
               From �ҺŰ���ʱ�� C
               Where ����id = n_����id And c.���� = v_���� And
                     Decode(��������_In, 1, To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(����ʱ��_In, 'hh24:mi')) =
                     To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By ����id, ����);
      Else
        --�мƻ���Чȡ�ƻ�
        --û��Ч�������ǴӹҺżƻ�ʱ�β�ѯ
        Select Nvl(���, -1),
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(c.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ʱ��ʱ��,
               ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �Һżƻ�ʱ�� C
        Where �ƻ�id = n_�ƻ�id And ���� = v_���� And
              (���, �ƻ�id, ����) In
              (Select Nvl(Max(c.���), -1), �ƻ�id, ����
               From �Һżƻ�ʱ�� C
               Where �ƻ�id = n_�ƻ�id And c.���� = v_���� And
                     Decode(��������_In, 1, To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(����ʱ��_In, 'hh24:mi')) =
                     To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By �ƻ�id, ����);
      End If;
    Exception
      When Others Then
        n_ʱ����� := -1;
        n_��ʱ��   := 0;
        d_ʱ��ʱ�� := ����ʱ��_In;
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
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ Not In (4, 5);
        Else
          Select Max(Nvl(���, 0)), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Nvl(ԤԼ, 0))
          Into n_�������, n_��������, n_��Լ��
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ <> 5;
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
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd ') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        Else
          --ԤԼʱ���
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
        
          If n_��Լ�� <= n_��Լ�� And n_��Լ�� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd') || '�Ѵﵽ�����Լ����';
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
          Where a.���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And
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
          v_Err_Msg := '�ű�' || �ű�_In || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
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
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ Not In (4, 5);
        Else
          Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(���� - d_ʱ��ʱ��), 0, 1, 0))
          Into n_�������, n_�ѹ���, n_��������
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ <> 5;
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
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
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
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
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
        Where ���� = Trunc(����ʱ��_In) And ���� = �ű�_In;
      Exception
        When Others Then
          n_�������� := 0;
          n_��Լ��   := 0;
      End;
      If Nvl(��������_In, 0) = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�
          If Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        Else
          --ԤԼ
          If (Nvl(n_��Լ��, 0) > 0 And Nvl(n_��Լ��, 0) >= Nvl(n_��Լ��, 0)) Or
             (Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0) Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
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
      d_���ʱ�� := ����ʱ��_In;
    Else
      d_���ʱ�� := Trunc(����ʱ��_In);
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
     ����Ա����_In, Decode(ԤԼ�Һ�_In, 1, ����Ա����_In, Null), ִ�в���id_In, ҽ������_In, ����Ա���_In, ����Ա����_In, ����ʱ��_In, �Ǽ�ʱ��_In, ���մ���id_In,
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
      Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = ����_In Where ����id = ����id_In;
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
       Null, ִ�в���id_In, ҽ������_In, 0, Null, �Ǽ�ʱ��_In, ����ʱ��_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����ʱ��_In, Null), ����Ա���_In,
       ����Ա����_In, ����_In, n_���, ����_In, Decode(ԤԼ����_In, 1, 1, 0), ԤԼ��ʽ_In, ժҪ_In, ������ˮ��_In, ����˵��_In, ������λ_In,
       Decode(Nvl(ԤԼ�Һ�_In, 0), 0, �Ǽ�ʱ��_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 0, ����Ա����_In, Null),
       Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա����_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա���_In, Null),
       Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, �շѵ�_In);
    If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
      Update ���˹Һż�¼
      Set ԤԼ = 1, ԤԼʱ�� = ����ʱ��_In, ԤԼ����Ա = ����Ա����_In, ԤԼ����Ա��� = ����Ա���_In
      Where ID = n_�Һ�id;
    End If;
    n_ԤԼ���ɶ��� := 0;
    If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
      n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    End If;
  
    --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
    If Nvl(���ɶ���_In, 0) <> 0 And Nvl(ԤԼ�Һ�_In, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(ִ�в���id_In, 0)));
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
        d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, �ű�_In, n_���, ����ʱ��_In);
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

--127075:����,2018-07-06,����ʹ�õ��˲��Ų�������̨ǩ���Ŷӵ�Oracle����
Create Or Replace Procedure Zl_���˹Һż�¼_Delete
(
  ���ݺ�_In       ������ü�¼.No%Type,
  ����Ա���_In   ������ü�¼.����Ա���%Type,
  ����Ա����_In   ������ü�¼.����Ա����%Type,
  ժҪ_In         ������ü�¼.ժҪ%Type := Null, --ԤԼȡ��ʱ ��д ���ԤԼȡ��ԭ��
  ɾ�������_In   Number := 0,
  ��ԭ���˽���_In Varchar2 := Null,
  �˷�����_In     In Number := 0, --0-ȫ�� 1-�˹Һŷ� 2-�˲����� 3-�˸��ӷ� 4-�˹Һ��벡�� 5-�˹Һ��븽��
  ��ָ������_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
  �˺�����_In     Number := 1,
  �ջ�Ʊ�ݺ�_In   Varchar2 := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null
) As
  --�˷�����_In,��һ�¼�������²�׼���в����˷�
  --    2.�����ӿ�,��ʱ��֧��
  -- �ҺŷѲ����ѷֿ���,����
  --    ��ͨ���㷽ʽ:ԭ���㷽ʽ�˲��ַ���
  --    Ԥ����:Ԥ����,�˲���
  --    Ԥ��������ͨ���㷽ʽ���:�˿����ͨ���㷽ʽ������
  --    ���ѿ�:ԭ�������ò����������ѿ�
  --��ԭ���˽���_In:ָ�����˻���ԭ�����㷽ʽ(��ҽ���ĸ����˻�,�����˻������ֵ�),����ö�����
  --��ָ������_IN:ָ��ԭ���˽��㲿��,Ӧ���˸����ֽ��㷽ʽ,Ϊ��ʱȱʡ�˸��ֽ�,�����˸�ָ���Ľ��㷽ʽ

  --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ���
  Cursor c_Registinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select a.����ʱ��, a.�Ǽ�ʱ��, c.����ʱ��, a.�շ�ϸĿid As ��Ŀid, c.ִ�в���id As ����id, c.ִ���� As ҽ������, d.Id As ҽ��id, c.�ű� As ����
    From ������ü�¼ A, �ҺŰ��� B, ���˹Һż�¼ C, ��Ա�� D
    Where a.��¼���� = 4 And a.��� = 1 And a.��¼״̬ = v_״̬ And c.No = a.No And c.ִ���� = d.����(+) And a.No = ���ݺ�_In And
          Nvl(a.���㵥λ, '�ű�') = c.�ű� And Rownum < 2;
  r_Registrow c_Registinfo%RowType;

  --���α������жϼ�¼�Ƿ����,�����û��ܱ���
  Cursor c_Moneyinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select ���˿���id, ��������id, ִ�в���id, ������Ŀid, Nvl(Sum(Ӧ�ս��), 0) As Ӧ��, Nvl(Sum(ʵ�ս��), 0) As ʵ��, Nvl(Sum(���ʽ��), 0) As ����
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = v_״̬ And NO = ���ݺ�_In
    Group By ���˿���id, ��������id, ִ�в���id, ������Ŀid;
  r_Moneyrow c_Moneyinfo%RowType;

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
  Cursor c_Opermoney(n_Id ����Ԥ����¼.����id%Type) Is
    Select Distinct b.���㷽ʽ, -1 * Nvl(b.��Ԥ��, 0) As ��Ԥ��
    From ����Ԥ����¼ B
    Where b.����id = n_Id And b.��¼���� = 4 And b.��¼״̬ = 2 And Nvl(b.��Ԥ��, 0) <> 0;

  v_Err_Msg Varchar(255);
  Err_Item Exception;

  n_����id ����Ԥ����¼.����id%Type;
  n_����id ������ü�¼.����id%Type;

  v_��ָ�����㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_�˿���       ����Ԥ����¼.��Ԥ��%Type;
  n_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;
  n_����id         ������Ϣ.����id%Type;
  n_�˷ѽ��       ����Ԥ����¼.��Ԥ��%Type;
  n_Ԥ�����       ����Ԥ����¼.��Ԥ��%Type; --ԭ��¼ Ԥ���ɿ���
  n_����ֵ         �������.Ԥ�����%Type;
  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_��id           ����ɿ����.Id%Type;

  n_�����˷�       Number; --��¼�Ƿ��Ǵ˵��ݵĵڶ����˷�
  n_����̨ǩ���Ŷ� Number;
  n_ԤԼ���ɶ���   Number;
  n_ԤԼ�Һ�       Number;
  n_�Һ����ɶ���   Number;
  d_Date           Date;
  n_����           ������ü�¼.���ʷ���%Type;
  n_����id1        ������Ϣ.����id%Type;
  n_���ض�         ������ü�¼.ʵ�ս��%Type;
  n_�ѽ���         Number;
  n_����id         �ҺŰ���.Id%Type;
  n_�ƻ�id         �ҺŰ��żƻ�.Id%Type;
  d_����ʱ��       Date;
  d_����ʱ��       ���˹Һż�¼.����ʱ��%Type;
  n_�����¼id     �ٴ������¼.Id%Type;
  v_����           �ҺŰ���.����%Type;
  n_���           ���˹Һż�¼.����%Type;
  v_ʱ���         ʱ���.ʱ���%Type;
  d_��鿪ʼʱ��   Date;
  d_������ʱ��   Date;
  v_Temp           Varchar2(500);
  v_����ids        Varchar2(500);
  n_���ﲡ��id     ������Ϣ.����id%Type;
  d_����ʱ��       ����ǼǼ�¼.����ʱ��%Type;
Begin
  n_��id           := Zl_Get��id(����Ա����_In);
  v_��ָ�����㷽ʽ := ��ָ������_In;

  --�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ����
  Open c_Moneyinfo(1);
  Fetch c_Moneyinfo
    Into r_Moneyrow;
  If c_Moneyinfo%RowCount = 0 Then
    Close c_Moneyinfo;
    Open c_Moneyinfo(0);
    Fetch c_Moneyinfo
      Into r_Moneyrow;
    If c_Moneyinfo%RowCount = 0 Then
      v_Err_Msg := 'Ҫ����ĵ��ݲ����ڡ�';
      Raise Err_Item;
    End If;
    n_ԤԼ�Һ� := 1;
  End If;
  Close c_Moneyinfo;

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

  Select �ű�, ����, ����ʱ�� Into v_����, n_���, d_����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum < 2;

  Begin
    Select a.Id Into n_����id From �ҺŰ��� A Where a.���� = v_����;
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
           Where a.���ʱ�� Is Not Null And d_����ʱ�� Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                 Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.����id = n_����id) And
          d_����ʱ�� Between Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
          Nvl(ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd'));
  Exception
    When Others Then
      n_�ƻ�id := 0;
  End;

  Begin
    If Nvl(n_�ƻ�id, 0) = 0 Then
      Select Decode(To_Char(d_����ʱ��, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ���
      Where ID = n_����id;
    Else
      Select Decode(To_Char(d_����ʱ��, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ��żƻ�
      Where ID = n_�ƻ�id;
    End If;
  Exception
    When Others Then
      v_ʱ��� := Null;
  End;

  If v_ʱ��� Is Not Null And d_����ʱ�� Is Not Null Then
    --����Ƿ��ģʽ�ҺŰ���
    Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
           To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
    Into d_��鿪ʼʱ��, d_������ʱ��
    From ʱ���
    Where ʱ��� = v_ʱ��� And վ�� Is Null And ���� Is Null;
    If d_��鿪ʼʱ�� > d_������ʱ�� Then
      d_������ʱ�� := d_������ʱ�� + 1;
    End If;
    If d_������ʱ�� > d_����ʱ�� Then
      --��ȡ�����¼id
      Begin
        Select a.Id
        Into n_�����¼id
        From �ٴ������¼ A, �ٴ������Դ B
        Where a.��Դid = b.Id And b.���� = v_���� And �ϰ�ʱ�� = v_ʱ��� And d_����ʱ�� Between ��ʼʱ�� And ��ֹʱ��;
      Exception
        When Others Then
          n_�����¼id := Null;
      End;
    End If;
  End If;

  Begin
    Select Zl_Fun_Regcustomname Into v_Temp From Dual;
    If v_Temp Is Not Null Then
      v_����ids := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    End If;
  Exception
    When Others Then
      v_����ids := Null;
  End;

  --1.ԤԼ����
  If Nvl(n_ԤԼ�Һ�, 0) = 1 Then
    --������Լ��
    Open c_Registinfo(0);
    Fetch c_Registinfo
      Into r_Registrow;
  
    Update ���˹ҺŻ���
    Set ��Լ�� = Nvl(��Լ��, 0) - 1
    Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
          (���� = r_Registrow.���� Or ���� Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
      Values
        (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
         Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1);
    End If;
    Close c_Registinfo;
  
    --���¹Һ����״̬
    Delete �Һ����״̬
    Where ״̬ = 2 And
          (����, ���, ����) = (Select ���㵥λ, ��ҩ����, Trunc(����ʱ��)
                          From ������ü�¼
                          Where ��¼���� = 4 And ��¼״̬ = 0 And ��� = 1 And Rownum = 1 And NO = ���ݺ�_In) Or
          (����, ���, ����) = (Select ���㵥λ, ��ҩ����, ����ʱ��
                          From ������ü�¼
                          Where ��¼���� = 4 And ��¼״̬ = 0 And ��� = 1 And Rownum = 1 And NO = ���ݺ�_In);
  
    --��Ӳ��˹Һż�¼�� ������¼
    Select ���˹Һż�¼_Id.Nextval, Sysdate Into n_�Һ�id, d_Date From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼״̬ = 1 And ��¼���� = 2;
    If Sql%NotFound Then
      v_Err_Msg := 'ԤԼ����' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ���ȡ��ԤԼ';
      Raise Err_Item;
    End If;
  
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
  
    If n_�����¼id Is Not Null Then
      Update �ٴ������¼ Set ��Լ�� = ��Լ�� - 1 Where ID = n_�����¼id And Nvl(��Լ��, 0) > 0;
      Update �ٴ�������ſ��� Set �Һ�״̬ = Null, ����Ա���� = Null Where ��¼id = n_�����¼id And ��� = n_���;
    End If;
  
    --Update ���˹Һż�¼ set ժҪ=nvl(ժҪ_IN,ժҪ) where NO=���ݺ�_IN;
    --ɾ��������ü�¼
    Delete From ������ü�¼ Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 0;
    --���ԤԼ���ɶ���ʱ��Ҫ�������
  
    n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    If Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
      --Ҫɾ������
      For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
        Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
      End Loop;
    End If;
    Return;
  End If;

  Select Nvl(���ʷ���, 0), ����id, Decode(Sign(Nvl(����id, 0)), 0, 0, 1)
  Into n_����, n_����id, n_�ѽ���
  From ������ü�¼
  Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ In (1, 3) And Rownum < 2;

  --2.�ҺŴ���
  n_�ѽ��� := Nvl(n_�ѽ���, 0);

  If n_�ѽ��� = 1 And n_���� = 1 Then
    Select Sysdate, Null Into d_Date, n_����id From Dual;
  Else
    Select Sysdate, ���˽��ʼ�¼_Id.Nextval Into d_Date, n_����id From Dual;
  End If;

  ----0-ȫ�� 1-�˹Һŷ� 2-�˲�����
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
    --���ǹ��˲�����ʱ����
    --���¹Һ����״̬
    If �˺�����_In = 1 Then
      Delete �Һ����״̬
      Where ״̬ = 1 And
            (����, ���, ����) = (Select �ű�, ����, Trunc(����ʱ��) From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1) Or
            (����, ���, ����) = (Select �ű�, ����, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1);
    Else
      Update �Һ����״̬
      Set ״̬ = 4
      Where ״̬ = 1 And
            (����, ���, ����) = (Select �ű�, ����, Trunc(����ʱ��) From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1) Or
            (����, ���, ����) = (Select �ű�, ����, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1);
    End If;
  
    --���˾���״̬
    If n_����id Is Not Null Then
      Update ������Ϣ Set ����״̬ = 0, �������� = Null Where ����id = n_����id;
    
      --ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ��
      If ɾ�������_In = 1 Then
        Delete ���ﲡ����¼ Where ����id = n_����id;
        Update ������Ϣ Set ����� = Null Where ����id = n_����id;
        --���ü�¼�����Һż����������￨����,�Լ����˽��Ѻ��˷ѻ����ʵķ���,�Һż�¼�������
        Update ������ü�¼ Set ��ʶ�� = Null Where �����־ = 1 And ����id = n_����id;
      End If;
    End If;
  
    --�����ʱ���˾��￨��,�˷�ʱ������￨��,�ڷǹ��˲�����ʱ
    n_����id1 := Null;
    Begin
      Select ����id
      Into n_����id1
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And ���ӱ�־ = 2 And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If n_����id1 Is Not Null And Nvl(�˷�����_In, 0) Not In (2, 3) Then
      Update ������Ϣ
      Set ���￨�� = Null, ����֤�� = Null, Ic���� = Decode(Ic����, ���￨��, Null, Ic����)
      Where ����id = n_����id1;
    End If;
  
  End If;

  --���ǰ���Ƿ��Ѿ������˹�����
  Begin
    Select 1 Into n_�����˷� From ������ü�¼ Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ = 3 And Rownum < 2;
  Exception
    When Others Then
      n_�����˷� := 0;
  End;

  If Nvl(�˷�����_In, 0) = 0 Or Nvl(�˷�����_In, 0) = 2 Then
    --ȫ��,�˲�����
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 1 Or Nvl(�˷�����_In, 0) = 4 Then
    --�˹Һŷ�,�˹Һ��벡����
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 3 Then
    --�˸��ӷ�
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 5 Then
    --�˹Һ��븽�ӷ�
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
    End If;
  End If;

  n_����id := 0;
  If n_���� = 0 Then
    --��ȡ����ID
    Select Nvl(����id, 0)
    Into n_����id
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And Rownum < 2;
  End If;

  If n_���� = 1 Then
    --����
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And
                       Nvl(���ӱ�־, 0) =
                       Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0)) And
                       Nvl(���ʷ���, 0) = 1) Loop
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
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
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
    --1.�˷�
    --���˹ҺŽ���:�ֽ�͸����ʻ�����
    If ��ԭ���˽���_In Is Not Null Then
      --�˿����ȡ
      If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
        --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ��
        Begin
          --��ȡ�����˿���
          Select -1 * Sum(Nvl(ʵ�ս��, 0)) As �տ���
          Into n_�˿���
          From ������ü�¼
          Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id;
        
        Exception
          When Others Then
            v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                           When 1 Then
                            '�Һŷ���'
                           When 2 Then
                            '������'
                         End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
            Raise Err_Item;
        End;
        Begin
          Select ��Ԥ��
          Into n_�˷ѽ��
          From ����Ԥ����¼
          Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
        Exception
          When Others Then
            n_�˷ѽ�� := 0;
        End;
      
        --a.����Ľ��㷽ʽ
      
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
           ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -n_�˿���,
                 n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
      
        If n_�˷ѽ�� = 0 Then
          --b.����������ֽ�
          If n_�˿��� <> 0 Then
            If v_��ָ�����㷽ʽ Is Null Then
              --�˸��ֽ�
              Begin
                Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
              Exception
                When Others Then
                  v_��ָ�����㷽ʽ := '�ֽ�';
              End;
            End If;
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˿���), ����˵�� = Nvl(����˵��_In, ����˵��)
            Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
                 �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
                Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                       ����Ա���_In, ����Ա����_In, -1 * n_�˿���, n_����id, n_��id, Ԥ�����, Decode(����˵��_In, Null, �����id, Null), ���㿨���,
                       Decode(����˵��_In, Null, ����, Null), Decode(����˵��_In, Null, ������ˮ��, Null), Nvl(����˵��_In, ����˵��), ������λ, 4
                From ����Ԥ����¼ A
                Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
            End If;
          End If;
        End If;
      Else
        --a.����Ľ��㷽ʽԭ����
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
           ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -��Ԥ��,
                 n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
      
        --b.����������ֽ�
        Begin
          Select Sum(��Ԥ��)
          Into n_�˷ѽ��
          From ����Ԥ����¼
          Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') > 0;
        Exception
          When Others Then
            n_�˷ѽ�� := 0;
        End;
        If n_�˷ѽ�� <> 0 Then
          If v_��ָ�����㷽ʽ Is Null Then
            --�˸��ֽ�
            Begin
              Select ���㷽ʽ
              Into v_��ָ�����㷽ʽ
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
            
            Exception
              When Others Then
                Begin
                  Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
                Exception
                  When Others Then
                    v_��ָ�����㷽ʽ := '�ֽ�';
                End;
            End;
          End If;
          Update ����Ԥ����¼
          Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˷ѽ��), ����˵�� = Nvl(����˵��_In, ����˵��)
          Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                     ����Ա���_In, ����Ա����_In, -1 * n_�˷ѽ��, n_����id, n_��id, Ԥ�����, Decode(����˵��_In, Null, �����id, Null), ���㿨���,
                     Decode(����˵��_In, Null, ����, Null), Decode(����˵��_In, Null, ������ˮ��, Null), Nvl(����˵��_In, ����˵��), ������λ, 4
              From ����Ԥ����¼ A
              Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
          End If;
        End If;
      End If;
    Else
      --�˿����ȡ
      If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
        --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ��
        Begin
          --��ȡ�����˿���
          Select -1 * Sum(Nvl(ʵ�ս��, 0)) As �տ���
          Into n_�˿���
          From ������ü�¼
          Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id;
        Exception
          When Others Then
            v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                           When 1 Then
                            '�Һŷ���'
                           When 2 Then
                            '������'
                         End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
            Raise Err_Item;
        End;
      End If;
      If Nvl(n_�����˷�, 0) = 0 And Nvl(�˷�����_In, 0) = 0 Then
        --�״�ȫ��
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
           ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -1 * ��Ԥ��,
                 n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id;
      Else
        --�����˷�,���߱��ε���һ����
        --�����˷�ʱ,��¼״̬=3 ,�״β�����,��¼״̬Ϊ1
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
           ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                 -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ժҪ = 'ҽ���Һ�' And ��Ԥ�� = n_�˿��� And
                Rownum < 2;
        If Sql%RowCount = 0 Then
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ��Ԥ�� = n_�˿��� And Rownum < 2;
        End If;
        If Sql%RowCount = 0 Then
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And Rownum < 2;
          If Sql%RowCount = 0 Then
            --�����˷�,����ȫ��ʹ��Ԥ����ɷ�ʱ�Ŵ��ڴ������
            n_Ԥ����� := n_�˿���;
          End If;
        End If;
      
      End If;
    End If;
    --�״��˷�ʱ,��¼״̬�����Ϊ��3
    Update ����Ԥ����¼
    Set ��¼״̬ = 3
    Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id;
  
    --��Ԥ�� 1-ȫ�� 2-������,������ʱ��ȫ��ʹ��Ԥ�����нɿ�
    If Nvl(�˷�����_In, 0) = 0 Or (Nvl(�˷�����_In, 0) <> 0 And n_Ԥ����� <> 0) Then
      --���˹ҺŽ���:��Ԥ�����
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_Date,
               ����Ա����_In, ����Ա���_In, -1 * Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, ��Ԥ��, n_Ԥ�����), n_����id, n_��id, Ԥ�����,
               �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
        From ����Ԥ����¼
        Where ��¼���� In (1, 11) And ����id = n_����id And Nvl(��Ԥ��, 0) <> 0 And
              Rownum = Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, Rownum, 1);
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
  
    If �ջ�Ʊ�ݺ�_In Is Not Null Then
      --���˹Һŷ�,������Ʊ��
      --�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�)
      Begin
        --�����һ�δ�ӡ��������ȡ
        Select ID
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 4 And b.No = ���ݺ�_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          n_��ӡid := Null;
      End;
    
      --���ջ�ԭƱ��
      If n_��ӡid Is Not Null Then
        Begin
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
        Exception
          When Others Then
            Delete From Ʊ��ʹ����ϸ Where ��ӡid = n_��ӡid And ���� = 2 And ԭ�� = 2;
            Insert Into Ʊ��ʹ����ϸ
              (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
              Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
              From Ʊ��ʹ����ϸ
              Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
        End;
      End If;
    End If;
  End If;

  --�����˲�������,��������ܼ�¼
  --��ػ��ܱ�Ĵ���

  --���˹ҺŻ���
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
    Open c_Registinfo(3);
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
        Select Decode(ԤԼ, Null, 0, 0, 0, 1) Into n_ԤԼ�Һ� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1;
      Exception
        When Others Then
          n_ԤԼ�Һ� := 0;
      End;
    
      Update ���˹ҺŻ���
      Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
      Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
            (���� = r_Registrow.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���, ��Լ��)
        Values
          (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
           Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1 * n_ԤԼ�Һ�, -1 * n_ԤԼ�Һ�);
      End If;
    
      If n_�����¼id Is Not Null Then
        Update �ٴ������¼
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
        Where ID = n_�����¼id And Nvl(��Լ��, 0) > 0;
        Update �ٴ�������ſ��� Set �Һ�״̬ = Null, ����Ա���� = Null Where ��¼id = n_�����¼id And ��� = n_���;
      End If;
    
      Close c_Registinfo;
    End If;
  End If;

  If n_���� = 0 Then
    --��Ա�ɿ����(���������ʻ��ȵĽ�����,�����˳�Ԥ����)
    For r_Opermoney In c_Opermoney(n_����id) Loop
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + (-1 * r_Opermoney.��Ԥ��)
      Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1
      Returning ��� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_Opermoney.���㷽ʽ, 1, -1 * r_Opermoney.��Ԥ��);
        n_����ֵ := r_Opermoney.��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
      End If;
    End Loop;
  End If;
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
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
    Begin
      Select ����id, ����ʱ�� Into n_���ﲡ��id, d_����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In;
      Delete From ����ǼǼ�¼ Where ����id = n_���ﲡ��id And ����ʱ�� = d_����ʱ�� And ��ҳid Is Null;
    Exception
      When Others Then
        Null;
    End;
  
    --���˹Һż�¼
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼���� = 1 And ��¼״̬ = 1;
    If Sql%NotFound Then
      v_Err_Msg := '�Һŵ���' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ����˺�';
      Raise Err_Item;
    End If;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
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
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_Delete;
/

--127075:����,2018-07-06,����ʹ�õ��˲��Ų�������̨ǩ���Ŷӵ�Oracle����
Create Or Replace Procedure Zl_ԤԼ�ҺŽ���_Insert
(
  No_In            ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���㷽ʽ_In      ����Ԥ����¼.���㷽ʽ%Type, --�ֽ�Ľ�������
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱʹ�õ�Ԥ�����
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  ����_In          �Һ����״̬.���%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ���ɶ���_In      Number := 0,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  ��Ԥ������ids_In Varchar2 := Null,
  ��������_In      Number := 0,
  ���½������_In  Number := 1,
  ժҪ_In          ���˹Һż�¼.ժҪ%Type := Null,
  �շѵ�_In        ���˹Һż�¼.�շѵ�%Type := Null
) As
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  --���½������_In:0-��zl_��Ա�ɿ����_Update �и��� 1-�ڱ������и���
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

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  Err_Special Exception;

  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
  v_�������� �ŶӽкŶ���.��������%Type;
  v_�ű�     ������ü�¼.���㵥λ%Type;
  v_����     ������ü�¼.��ҩ����%Type;
  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ��� %Type;
  v_ԤԼ��ʽ ���˹Һż�¼.ԤԼ��ʽ %Type;

  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_����̨ǩ���Ŷ� Number;
  n_��id           ����ɿ����.Id%Type;
  n_Count          Number(18);
  n_�Ŷ�           Number;
  n_�����Ŷ�       Number;
  n_��ǰ���       ����Ԥ����¼.���%Type;
  n_Ԥ��id         ����Ԥ����¼.Id%Type;

  d_Date     Date;
  d_ԤԼʱ�� ������ü�¼.����ʱ��%Type;
  d_����ʱ�� Date;
  d_�Ŷ�ʱ�� Date;
  n_ʱ��     Number := 0;
  n_����     Number := 0;
  v_�Ŷ���� �ŶӽкŶ���.�Ŷ����%Type;
  n_����ģʽ ������Ϣ.����ģʽ%Type;

  v_���ʽ   ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  v_����Ա���� ���˹Һż�¼.������%Type;
  n_����ģʽ   Number := 0;
Begin
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
  n_����ģʽ      := Nvl(zl_GetSysParameter('ԤԼ����ģʽ', 1111), 0);

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
  If �Ǽ�ʱ��_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := �Ǽ�ʱ��_In;
  End If;

  --���¹Һ����״̬
  Begin
    Select �ű�, ����, Trunc(����ʱ��), ����ʱ��, ԤԼ��ʽ
    Into v_�ű�, v_����, d_ԤԼʱ��, d_����ʱ��, v_ԤԼ��ʽ
    From ���˹Һż�¼
    Where ��¼���� = 2 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      Select Max(������) Into v_����Ա���� From ���˹Һż�¼ Where ��¼���� = 2 And ��¼״̬ In (1, 3) And NO = No_In;
      If v_����Ա���� Is Null Then
        v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�ȡ��';
        Raise Err_Item;
      Else
        If v_����Ա���� = ����Ա����_In Then
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�����';
          Raise Err_Special;
        Else
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ������˽���';
          Raise Err_Item;
        End If;
      End If;
  End;

  --�ж��Ƿ��ʱ��
  Begin
    Select 1
    Into n_ʱ��
    From Dual
    Where Exists (Select 1
           From �ҺŰ���ʱ�� A, �ҺŰ��� B
           Where a.����id = b.Id And b.���� = v_�ű� And Rownum < 2
           Union All
           Select 1
           From �Һżƻ�ʱ�� C, �ҺŰ��żƻ� D ��
           Where c.�ƻ�id = d.Id And d.���� = v_�ű� And d.��Чʱ�� > Sysdate And Rownum < 2);
  Exception
    When Others Then
      n_ʱ�� := 0;
  End;
  --��ʱ�εĺű�ֻ�ܵ������
  If n_ʱ�� = 1 And ��������_In = 0 And n_����ģʽ = 0 Then
    If Trunc(����ʱ��_In) <> Trunc(Sysdate) Then
      v_Err_Msg := '��ʱ�ε�ԤԼ�Һŵ�ֻ�ܵ�����գ�';
      Raise Err_Item;
    End If;
  End If;

  If n_ʱ�� = 0 And ��������_In = 0 Then
    If n_����ģʽ = 0 Then
      If Trunc(����ʱ��_In) = Trunc(Sysdate) Then
        d_����ʱ�� := ����ʱ��_In;
      Else
        d_����ʱ�� := Sysdate;
      End If;
    Else
      d_����ʱ�� := ����ʱ��_In;
    End If;
  Else
    If Not ����ʱ��_In Is Null Then
      d_����ʱ�� := ����ʱ��_In;
    End If;
  End If;
  If Not v_���� Is Null Then
    If ����_In Is Null Then
      Delete �Һ����״̬ Where ���� = v_�ű� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_����;
    Else
      If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
      
        If n_ʱ�� = 0 And ��������_In = 0 Then
          --��ǰ���ջ��ӳٽ���
          Delete �Һ����״̬ Where ���� = v_�ű� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_����;
          Begin
            Select 1 Into n_���� From �Һ����״̬ Where ���� = v_�ű� And ���� = Trunc(Sysdate) And ��� = v_����;
          Exception
            When Others Then
              n_���� := 0;
          End;
          If n_���� = 0 Then
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
            Values
              (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
          Else
            --�����ѱ�ʹ�õ����
            Begin
              v_���� := 1;
              Insert Into �Һ����״̬
                (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
              Values
                (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
            Exception
              When Others Then
                Select Min(��� + 1)
                Into v_����
                From �Һ����״̬ A
                Where ���� = v_�ű� And ���� = Trunc(Sysdate) And Not Exists
                 (Select 1 From �Һ����״̬ Where ���� = a.���� And ���� = a.���� And ��� = a.��� + 1);
                Insert Into �Һ����״̬
                  (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
                Values
                  (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
            End;
          End If;
        Else
          Update �Һ����״̬
          Set ״̬ = 1, �Ǽ�ʱ�� = Sysdate
          Where Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_���� And ���� = v_�ű� And ״̬ = 2;
          If Sql% NotFound Then
            Begin
              Insert Into �Һ����״̬
                (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
              Values
                (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
            Exception
              When Others Then
                v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
                Raise Err_Item;
            End;
          End If;
        
        End If;
      
      Else
        Update �Һ����״̬
        Set ��� = ����_In, ״̬ = 1, �Ǽ�ʱ�� = Sysdate
        Where ���� = v_�ű� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_����;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
            Values
              (v_�ű�, Trunc(d_����ʱ��), v_����, 1, ����Ա����_In, Sysdate);
          Exception
            When Others Then
              v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
              Raise Err_Item;
          End;
        End If;
      End If;
    End If;
  Else
    If Not ����_In Is Null Then
      Begin
        Insert Into �Һ����״̬
          (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
        Values
          (v_�ű�, Trunc(Sysdate), ����_In, 1, ����Ա����_In, Sysdate);
      Exception
        When Others Then
          v_Err_Msg := '���' || ����_In || '�ѱ�������ʹ��,������ѡ��һ�����.';
          Raise Err_Item;
      End;
      v_���� := ����_In;
    Else
      v_���� := Null;
    End If;
  End If;

  --����������ü�¼
  Update ������ü�¼
  Set ��¼״̬ = 1, ʵ��Ʊ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, Ʊ�ݺ�_In), ����id = Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In),
      ���ʽ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��), ��ҩ���� = ����_In, ����id = ����id_In, ��ʶ�� = �����_In, ���� = ����_In, ���� = ����_In,
      �Ա� = �Ա�_In, ���ʽ = ���ʽ_In, �ѱ� = �ѱ�_In, ����ʱ�� = d_����ʱ��, �Ǽ�ʱ�� = d_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In,
      �ɿ���id = n_��id, ���ʷ��� = Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ժҪ = Decode(�շѵ�_In, Null, Nvl(ժҪ_In, ժҪ), '����:' || �շѵ�_In)
  Where ��¼���� = 4 And ��¼״̬ = 0 And NO = No_In;

  --���˹Һż�¼
  Update ���˹Һż�¼
  Set ������ = ����Ա����_In, ����ʱ�� = d_Date, ��¼���� = 1, ����id = ����id_In, ����� = �����_In, ����ʱ�� = d_����ʱ��, ���� = ����_In, �Ա� = �Ա�_In,
      ���� = ����_In, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, ���� = Decode(Nvl(����_In, 0), 0, Null, ����_In), ���� = v_����, ���� = ����_In,
      ժҪ = Nvl(ժҪ_In, ժҪ), �շѵ� = �շѵ�_In
  Where ��¼״̬ = 1 And NO = No_In And ��¼���� = 2
  Returning ID Into n_�Һ�id;
  If Sql%NotFound Then
    Begin
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ժҪ, ����, ԤԼ, ԤԼ��ʽ, ������, ����ʱ��, ԤԼʱ��, ����, ҽ�Ƹ��ʽ, �շѵ�)
        Select n_�Һ�id, No_In, 1, 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, ���㵥λ, �Ӱ��־, ����_In, Null, ִ�в���id, ִ����, 0, Null,
               �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����, Nvl(ժҪ_In, ժҪ), v_����, 1, Substr(����, 1, 10) As ԤԼ��ʽ, ����Ա����_In,
               Nvl(�Ǽ�ʱ��_In, Sysdate), ����ʱ��, Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, �շѵ�_In
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '���ڲ���ԭ��,���ݺ�Ϊ��' || No_In || '���Ĳ���' || ����_In || '�Ѿ�������';
        Raise Err_Item;
    End;
  End If;

  --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
  If Nvl(���ɶ���_In, 0) <> 0 Then
    For v_�Һ� In (Select ID, ����, ����, ִ����, ִ�в���id, ����ʱ��, �ű�, ���� From ���˹Һż�¼ Where NO = No_In) Loop
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(v_�Һ�.ִ�в���id, 0)));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
        Begin
          Select 1,
                 Case
                   When �Ŷ�ʱ�� < Trunc(Sysdate) Then
                    1
                   Else
                    0
                 End
          Into n_�Ŷ�, n_�����Ŷ�
          From �ŶӽкŶ���
          Where ҵ������ = 0 And ҵ��id = v_�Һ�.Id And Rownum <= 1;
        Exception
          When Others Then
            n_�Ŷ� := 0;
        End;
        If n_�Ŷ� = 0 Then
          --��������
          --����ִ�в��š���������
          n_�Һ�id   := v_�Һ�.Id;
          v_�������� := v_�Һ�.ִ�в���id;
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, n_�Һ�id, v_�Һ�.�ű� || '|' || v_�Һ�.����);
          v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        
          --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
          d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, v_�Һ�.�ű�, v_�Һ�.����, v_�Һ�.����ʱ��);
          --   ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In,
          Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, v_�Һ�.ִ�в���id, v_�ŶӺ���, Null, ����_In, ����id_In, v_�Һ�.����, v_�Һ�.ִ����, d_�Ŷ�ʱ��,
                           v_ԤԼ��ʽ, Null, v_�Ŷ����);
        Elsif Nvl(n_�����Ŷ�, 0) = 1 Then
          --���¶��к�
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, v_�Һ�.Id, v_�Һ�.�ű� || '|' || Nvl(v_�Һ�.����, 0));
          v_�Ŷ���� := Zlgetsequencenum(0, v_�Һ�.Id, 1);
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����, v_�ŶӺ���, v_�Ŷ����);
        
        Else
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����);
        End If;
        --ԤԼ����ʱ���ı��¼��־
        Update ���˹Һż�¼ Set ��¼��־ = 1 Where ID = n_�Һ�id;
      End If;
    End Loop;
  End If;

  --���ܽ��㵽����Ԥ����¼
  If (Nvl(�ֽ�֧��_In, 0) <> 0 Or (Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0)) And
     Nvl(���ʷ���_In, 0) = 0 Then
    Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������,
       ��������)
    Values
      (n_Ԥ��id, 4, 1, No_In, ����id_In, Nvl(���㷽ʽ_In, v_�ֽ�), Nvl(�ֽ�֧��_In, 0), d_Date, ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�',
       n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ����id_In, 4);
  
    If Nvl(���㿨���_In, 0) <> 0 And Nvl(�ֽ�֧��_In, 0) <> 0 Then
      Zl_���˿������¼_֧��(���㿨���_In, ����_In, 0, �ֽ�֧��_In, n_Ԥ��id, ����Ա���_In, ����Ա����_In, d_Date);
    End If;
  
  End If;

  --���ھ��￨ͨ��Ԥ����Һ�
  If Nvl(Ԥ��֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
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
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �Ǽ�ʱ��_In,
               ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, ����id_In, 4
        From ����Ԥ����¼
        Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
      Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (r_Deposit.����id, Nvl(1, 2), -1 * n_��ǰ���, 1);
        n_����ֵ := -1 * n_��ǰ���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
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
  End If;

  --����ҽ���Һ�
  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
       Ԥ�����, �������, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 4, 1, No_In, ����id_In, v_�����ʻ�, ����֧��_In, d_Date, ����Ա���_In, ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id,
       Null, Null, Null, Null, Null, Null, Null, ����id_In, 4);
  End If;

  --��ػ��ܱ�Ĵ���
  --��Ա�ɿ����
  If Nvl(�ֽ�֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 And Nvl(���½������_In, 1) = 1 Then
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
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�) And Nvl(���, 0) = 0;
    End If;
  End If;

  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 And Nvl(���½������_In, 1) = 1 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ����֧��_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
      n_����ֵ := ����֧��_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ���� Where �տ�Ա = ����Ա����_In And ���� = 1 And Nvl(���, 0) = 0;
    End If;
  End If;

  If Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = No_In And Nvl(���ʷ���, 0) = 1) Loop
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (����id_In, 1, 1, Nvl(c_����.ʵ�ս��, 0), 0);
      End If;
    
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (����id_In, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
  End If;
  If Nvl(����id_In, 0) <> 0 Then
    n_����ģʽ := 0;
    Update ������Ϣ
    Set ����ʱ�� = d_����ʱ��, ����״̬ = 1, �������� = ����_In
    Where ����id = ����id_In
    Returning Nvl(����ģʽ, 0) Into n_����ģʽ;
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
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
  End If;

  --���˵�����Ϣ
  If ����id_In Is Not Null Then
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Nvl(��Ժ, 0) = 0 And Exists
     (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = d_Date
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, d_Date) >= d_Date;
    End If;
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 1, n_�Һ�id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ԤԼ�ҺŽ���_Insert;
/

--127075:����,2018-07-06,����ʹ�õ��˲��Ų�������̨ǩ���Ŷӵ�Oracle����
Create Or Replace Procedure Zl_���˹Һż�¼_ȡ��ǩ��
(
  Id_In           ���˹Һż�¼.Id%Type,
  ����ǩ����־_In Integer := 0
) As
  --:����ǩ����־_In:0-��Ҫ������ص��ŶӼ�¼;1-���ٴ����ŶӼ�¼,ֻ��ǩ���ı�ʶ��д 
  Err_Item Exception;
  v_Err_Msg      Varchar2(200);
  n_�Һ����ɶ��� Number;
Begin
  --:��¼��־:0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���; 
  Update ���˹Һż�¼ Set ��¼��־ = 0 Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '�Һ���Ŀδ�ҵ�,���ܼ�������!';
    Raise Err_Item;
  End If;
  If Nvl(����ǩ����־_In, 0) = 1 Then
    Return;
  End If;
  n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
  --�Һ�����ǩ��ҲӦ��֧��ȡ��ǩ�� 
  If n_�Һ����ɶ��� = 0 Then
    Return;
  End If;
  For v_�Һ� In (Select ִ�в���id, ID From ���˹Һż�¼ Where ID = Id_In) Loop
    Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_ȡ��ǩ��;
/





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0020' Where ���=&n_System;
Commit;
