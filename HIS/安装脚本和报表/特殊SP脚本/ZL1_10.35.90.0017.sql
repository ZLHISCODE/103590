----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--122504:������,2018-06-21,������ϵͳZLHIS����޸�
Alter Table ������ҳ Add �Һ�ID number(18); 
Create Table ������������Ŀ¼(
    ϵͳ��ʶ    varchar2(100),  
    ��������    varchar2(100),  
    �����ַ  varchar2(300))    
    TABLESPACE zl9BaseItem;
Alter Table ������������Ŀ¼ Add Constraint ������������Ŀ¼_PK Primary Key (ϵͳ��ʶ,��������) Using Index Tablespace zl9IndexHis;    

Create Index ������ҳ_IX_�Һ�ID On ������ҳ(�Һ�ID)  Tablespace zl9Indexcis;

------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--127144:��ΰ��,2018-06-21,��������
Insert Into ������������Ŀ¼ (ϵͳ��ʶ, ��������) Values ('֪ʶ��', '��������');

--122504:������,2018-06-21,������ϵͳZLHIS����޸�
Insert into zlTables ( ϵͳ,����,��ռ�,���� ) Values( &n_System,'������������Ŀ¼','ZL9BASEITEM','A1');

Insert Into ������������Ŀ¼(ϵͳ��ʶ,��������) 
Select '������ϵͳ','�ж�ҽ���Ƿ��շ�' From Dual Union All
Select '������ϵͳ','�������תסԺ����ȷ��' From Dual Union All
Select '������ϵͳ','�������תסԺ����' From Dual;


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--122504:������,2018-06-21,������ϵͳZLHIS����޸�
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1011,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
Select 'Zl_������������Ŀ¼_Update','EXECUTE' From Dual Union All 
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;  




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------

--122504:������,2018-06-21,������ϵͳZLHIS����޸�
CREATE OR REPLACE Procedure Zl_����ҽ������_Insert
(
  ҽ��id_In     In ����ҽ������.ҽ��id%Type,
  ���ͺ�_In     In ����ҽ������.���ͺ�%Type,
  ��¼����_In   In ����ҽ������.��¼����%Type,
  No_In         In ����ҽ������.No%Type,
  ��¼���_In   In ����ҽ������.��¼���%Type,
  ��������_In   In ����ҽ������.��������%Type,
  �״�ʱ��_In   In ����ҽ������.�״�ʱ��%Type,
  ĩ��ʱ��_In   In ����ҽ������.ĩ��ʱ��%Type,
  ����ʱ��_In   In ����ҽ������.����ʱ��%Type,
  ִ��״̬_In   In ����ҽ������.ִ��״̬%Type,
  ִ�в���id_In In ����ҽ������.ִ�в���id%Type,
  �Ʒ�״̬_In   In ����ҽ������.�Ʒ�״̬%Type,
  First_In      In Number := 0,
  ��������_In   In ����ҽ������.��������%Type := Null,
  ����Ա���_In In ��Ա��.���%Type := Null,
  ����Ա����_In In ��Ա��.����%Type := Null,
  ԭҺƤ��_In   In Varchar2 := Null
  --���ܣ���д����ҽ�����ͼ�¼
  --������First_IN=��ʾ�Ƿ�һ��ҽ���ĵ�һҽ����,�Ա㴦��ҽ���������(���ҩ,�䷽�ĵ�һ��,��Ϊ��ҩ;��,�䷽�巨,�÷�����Ϊ����������)
  --      ԴҺƤ��_In ԭҺƤ��ҽ��ID�������7107/bug115972���ڹ���ҩƷҽ���к�Ƥ��ҽ���С������ֶ�Ϊ ����ҽ������.�걾�������� ����ҩƷ�е�ҽ��IDֵ
  --      ��ʽ��1ҽ��ID,2ҽ��ID ǰ��һ��ΪƤ��ҽ����ҽ��ID���ڶ���ΪҩƷ��ҽ����ҽ��ID
) Is
  --�������˼�ҽ��(һ��ҽ���е�һ��)�����Ϣ���α�
  Cursor c_Advice Is
    Select Nvl(a.���id, a.Id) As ��id, a.���id, a.���, a.����id, a.�Һŵ�, a.Ӥ��, b.����, c.��������, a.�������, a.ҽ��״̬, a.ҽ������, a.����ҽ��,
           a.��ʼִ��ʱ��, a.ִ��ʱ�䷽��, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, Nvl(a.������־, 0) As ������־, a.������Ŀid, a.�շ�ϸĿid
    From ����ҽ����¼ A, ������Ϣ B, ������ĿĿ¼ C
    Where a.����id = b.����id And a.������Ŀid = c.Id And a.Id = ҽ��id_In
    Group By a.���id, a.Id, a.���, a.����id, a.�Һŵ�, a.Ӥ��, b.����, c.��������, a.�������, a.ҽ��״̬, a.ҽ������, a.����ҽ��, a.��ʼִ��ʱ��, a.ִ��ʱ�䷽��,
             a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, a.������־, a.������Ŀid, a.�շ�ϸĿid;
  r_Advice c_Advice%RowType;

  Cursor c_Pati(v_����id ������Ϣ.����id%Type) Is
    Select * From ������Ϣ Where ����id = v_����id;
  r_Pati c_Pati%RowType;

  --������ʱ����
  v_Temp       Varchar2(255);
  v_Count      Number;
  v_��������   ������ҳ.��������%Type;
  v_��Ա���   ��Ա��.���%Type;
  v_��Ա����   ��Ա��.����%Type;
  v_��Ժ��ʽ   ��Ժ��ʽ.����%Type;
  n_�Һ�id     ���˹Һż�¼.Id%Type;
  d_��ʼʱ��   ����ҽ����¼.��ʼִ��ʱ��%Type;
  n_ҽ��״̬   ����ҽ����¼.ҽ��״̬%Type;
  n_Ƥ�Ա��   ����ҽ������.ҽ��id%Type;
  n_Ƥ��ҽ��id ����ҽ������.ҽ��id%Type;
  v_Error      Varchar2(255);
  Err_Custom Exception;
Begin
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
  --����״�ʱ��Ϊ�������뿪ʼִ��ʱ��
  Select ��ʼִ��ʱ��, ҽ��״̬ Into d_��ʼʱ��, n_ҽ��״̬ From ����ҽ����¼ Where ID = ҽ��id_In;

  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;
  --��һ��ҽ���ĵ�һ��ʱ����ҽ������
  If Nvl(First_In, 0) = 1 Or n_ҽ��״̬ = 1 Then
    --�����������
    ---------------------------------------------------------------------------------------
    If Nvl(r_Advice.ҽ��״̬, 0) <> 1 Then
      v_Error := '"' || r_Advice.���� || '"��ҽ��"' || r_Advice.ҽ������ || '"�Ѿ��������˷��͡�' || Chr(13) || Chr(10) ||
                 '�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
      Raise Err_Custom;
    End If;
  
    --���ͺ��ҽ������:�������ͺ��Զ�ֹͣ
    ---------------------------------------------------------------------------------------
    Update ����ҽ����¼
    Set ҽ��״̬ = 8, ִ����ֹʱ�� = ĩ��ʱ��_In,
        --����û��
        ͣ��ʱ�� = ����ʱ��_In,
        --Ҫ��Ϊ����ʱ����ʾ
        ͣ��ҽ�� = v_��Ա���� --Ҫ��Ϊ��������ʾ,��ͬ��סԺ,����ҽ���޻�ʿ����
    Where ID = r_Advice.��id Or ���id = r_Advice.��id;
  
    Insert Into ����ҽ��״̬
      (ҽ��id, ��������, ������Ա, ����ʱ��)
      Select ID, 8, v_��Ա����, ����ʱ��_In From ����ҽ����¼ Where ID = r_Advice.��id Or ���id = r_Advice.��id;
  
    --����ҽ���Ĵ���
    ---------------------------------------------------------------------------------------
    If r_Advice.������� = 'Z' And Nvl(r_Advice.��������, '0') <> '0' And Nvl(r_Advice.Ӥ��, 0) = 0 Then
      --1-����;2-סԺ;
      If Instr(',1,2,', r_Advice.��������) > 0 And ִ�в���id_In Is Not Null Then
        --��������µ�ԤԼ�Ǽǵ�������1.��ǰ��ԤԼ,2.��ǰ����Ժ,3-��Ҫ��ԤԼʱ���ڵ�סԺ��¼
      
        --ɾ�������Һ���Ч������ԤԼ�Ǽ�
        Begin
          Select Count(*) Into v_Count From ������ҳ Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = 0;
        Exception
          When Others Then
            v_Count := 0;
        End;
        If Nvl(v_Count, 0) > 0 Then
          Zl_��Ժ������ҳ_Delete(r_Advice.����id, 0, 0, 0);
          v_Count := 0;
        End If;
      
        If v_Count = 0 Then
          Select Count(*) Into v_Count From ������ҳ Where ����id = r_Advice.����id And ��Ժ���� Is Null;
        End If;
        If v_Count = 0 Then
          Select Count(*)
          Into v_Count
          From ������ҳ
          Where ����id = r_Advice.����id And (��Ժ���� >= r_Advice.��ʼִ��ʱ�� Or ��Ժ���� >= r_Advice.��ʼִ��ʱ��);
        End If;
        If v_Count = 0 Then
          If r_Advice.�������� = '1' Then
            --����ҽ��,��������"��ʼʱ��"���۵��ٴ�ִ�п���
            Begin
              v_�������� := 2;
              Select Decode(�������, 1, 1, 2)
              Into v_��������
              From ��������˵��
              Where �������� = '�ٴ�' And ����id = ִ�в���id_In;
            Exception
              When Others Then
                Null;
            End;
          Elsif r_Advice.�������� = '2' Then
            --סԺҽ��,��������"��ʼʱ��"�Ǽǵ��ٴ�ִ�п���
            v_�������� := 0;
          End If;
        
          Open c_Pati(r_Advice.����id);
          Fetch c_Pati
            Into r_Pati;
        
          v_��Ժ��ʽ := Null;
          If r_Advice.������־ = 1 Then
            v_��Ժ��ʽ := '����';
            Select Max(ID)
            Into n_�Һ�id
            From ���˹Һż�¼
            Where NO = r_Advice.�Һŵ� And ��¼���� = 1 And ��¼״̬ = 1;
          Else
            Select Decode(����, 1, '����', Null), ID
            Into v_��Ժ��ʽ, n_�Һ�id
            From ���˹Һż�¼
            Where NO = r_Advice.�Һŵ� And ��¼���� = 1 And ��¼״̬ = 1;
          End If;
        
          If v_�������� = 1 Then
            Zl_��Ժ������ҳ_Insert(1, v_��������, r_Pati.����id, r_Pati.�����, Null, r_Pati.����, r_Pati.�Ա�, r_Pati.����, r_Pati.�ѱ�,
                             r_Pati.��������, r_Pati.����, r_Pati.����, r_Pati.ѧ��, r_Pati.����״��, r_Pati.ְҵ, r_Pati.���,
                             r_Pati.���֤��, r_Pati.�����ص�, r_Pati.��ͥ��ַ, r_Pati.��ͥ��ַ�ʱ�, r_Pati.��ͥ�绰, r_Pati.���ڵ�ַ,
                             r_Pati.���ڵ�ַ�ʱ�, r_Pati.��ϵ������, r_Pati.��ϵ�˹�ϵ, r_Pati.��ϵ�˵�ַ, r_Pati.��ϵ�˵绰, r_Pati.������λ,
                             r_Pati.��ͬ��λid, r_Pati.��λ�绰, r_Pati.��λ�ʱ�, r_Pati.��λ������, r_Pati.��λ�ʺ�, r_Pati.������, r_Pati.������,
                             r_Pati.��������, ִ�в���id_In, Null, Null, v_��Ժ��ʽ, Null, Null, r_Advice.����ҽ��, r_Pati.����, r_Pati.����,
                             r_Advice.��ʼִ��ʱ��, Null, Null, r_Pati.ҽ�Ƹ��ʽ, Null, Null, Null, Null, Null, Null, r_Pati.����,
                             v_��Ա���, v_��Ա����, 0, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, n_�Һ�id);
          Else
            Zl_��Ժ������ҳ_Insert(1, v_��������, r_Pati.����id, r_Pati.סԺ��, Null, r_Pati.����, r_Pati.�Ա�, r_Pati.����, r_Pati.�ѱ�,
                             r_Pati.��������, r_Pati.����, r_Pati.����, r_Pati.ѧ��, r_Pati.����״��, r_Pati.ְҵ, r_Pati.���,
                             r_Pati.���֤��, r_Pati.�����ص�, r_Pati.��ͥ��ַ, r_Pati.��ͥ��ַ�ʱ�, r_Pati.��ͥ�绰, r_Pati.���ڵ�ַ,
                             r_Pati.���ڵ�ַ�ʱ�, r_Pati.��ϵ������, r_Pati.��ϵ�˹�ϵ, r_Pati.��ϵ�˵�ַ, r_Pati.��ϵ�˵绰, r_Pati.������λ,
                             r_Pati.��ͬ��λid, r_Pati.��λ�绰, r_Pati.��λ�ʱ�, r_Pati.��λ������, r_Pati.��λ�ʺ�, r_Pati.������, r_Pati.������,
                             r_Pati.��������, ִ�в���id_In, Null, Null, v_��Ժ��ʽ, Null, Null, r_Advice.����ҽ��, r_Pati.����, r_Pati.����,
                             r_Advice.��ʼִ��ʱ��, Null, Null, r_Pati.ҽ�Ƹ��ʽ, Null, Null, Null, Null, Null, Null, r_Pati.����,
                             v_��Ա���, v_��Ա����, 0, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, n_�Һ�id);
          End If;
          Close c_Pati;
        End If;
      End If;
    End If;
  End If;
  Close c_Advice;

  If ԭҺƤ��_In Is Not Null Then
    v_Count      := Instr(ԭҺƤ��_In, ',');
    n_Ƥ��ҽ��id := Substr(ԭҺƤ��_In, 1, v_Count - 1);
    n_Ƥ�Ա��   := Substr(ԭҺƤ��_In, v_Count + 1);
    Update ����ҽ������ Set �걾�������� = n_Ƥ�Ա�� Where ҽ��id = n_Ƥ��ҽ��id;
  End If;
  --��д���ͼ�¼
  ---------------------------------------------------------------------------------------
  Insert Into ����ҽ������
    (ҽ��id, ���ͺ�, ��¼����, NO, ��¼���, ��������, ������, ����ʱ��, ִ��״̬, ִ�в���id, �Ʒ�״̬, �״�ʱ��, ĩ��ʱ��, ��������, �������, �걾��������)
  Values
    (ҽ��id_In, ���ͺ�_In, ��¼����_In, No_In, ��¼���_In, ��������_In, v_��Ա����, ����ʱ��_In, ִ��״̬_In, ִ�в���id_In, �Ʒ�״̬_In,
     Nvl(�״�ʱ��_In, d_��ʼʱ��), Nvl(ĩ��ʱ��_In, d_��ʼʱ��), ��������_In, Decode(��¼����_In, 2, 1, Null), n_Ƥ�Ա��);

  --�����ͼ��ҽ��ͬ��������ҽ���ļƷ�״̬
  If �Ʒ�״̬_In = 1 And r_Advice.��id <> ҽ��id_In And (r_Advice.������� = 'D' Or r_Advice.������� = 'F') Then
    Update ����ҽ������ Set �Ʒ�״̬ = 1 Where ҽ��id = r_Advice.��id And ���ͺ� = ���ͺ�_In;
  End If;

  --�Զ���Ϊ��ִ��ʱ����Ҫͬ���������ִ��״̬����˻���״̬
  If ִ��״̬_In = 1 Then
    Zl_����ҽ��ִ��_Finish(ҽ��id_In, ���ͺ�_In, Null, Null, v_��Ա���, v_��Ա����, ִ�в���id_In);
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin zl_������Ϣ_����(:1,:2); End;'
      Using 3, ���ͺ�_In;
  Exception
    When Others Then
      Null;
  End;

  If r_Advice.������� = 'E' And r_Advice.�������� = '6' Then
    --������Ŀ
    b_Message.Zlhis_Cis_016(r_Advice.����id, Null, r_Advice.�Һŵ�, ���ͺ�_In, r_Advice.��id, 1);
  Elsif r_Advice.������� = 'D' And r_Advice.���id Is Null Then
    b_Message.Zlhis_Cis_017(r_Advice.����id, Null, r_Advice.�Һŵ�, ���ͺ�_In, r_Advice.��id, 1);
  Elsif r_Advice.������� = 'F' And r_Advice.���id Is Null Then
    b_Message.Zlhis_Cis_018(r_Advice.����id, Null, r_Advice.�Һŵ�, ���ͺ�_In, r_Advice.��id);
  Elsif r_Advice.������� = 'K' Then
    b_Message.Zlhis_Cis_019(r_Advice.����id, Null, r_Advice.�Һŵ�, ���ͺ�_In, r_Advice.��id);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ������_Insert;
/

--122504:������,2018-06-21,������ϵͳZLHIS����޸�
Create Or Replace Procedure Zl_��Ժ������ҳ_Insert
(
  �Ǽ�ģʽ_In       Number,
  ��������_In       ������ҳ.��������%Type,
  ����id_In         ������Ϣ.����id%Type,
  סԺ��_In         ������Ϣ.סԺ��%Type,
  ҽ����_In         �����ʻ�.ҽ����%Type,
  ����_In           ������Ϣ.����%Type,
  �Ա�_In           ������Ϣ.�Ա�%Type,
  ����_In           ������Ϣ.����%Type,
  �ѱ�_In           ������Ϣ.�ѱ�%Type,
  ��������_In       ������Ϣ.��������%Type,
  ����_In           ������Ϣ.����%Type,
  ����_In           ������Ϣ.����%Type,
  ѧ��_In           ������Ϣ.ѧ��%Type,
  ����״��_In       ������Ϣ.����״��%Type,
  ְҵ_In           ������Ϣ.ְҵ%Type,
  ���_In           ������Ϣ.���%Type,
  ���֤��_In       ������Ϣ.���֤��%Type,
  �����ص�_In       ������Ϣ.�����ص�%Type,
  ��ͥ��ַ_In       ������Ϣ.��ͥ��ַ%Type,
  ��ͥ��ַ�ʱ�_In   ������Ϣ.��ͥ��ַ�ʱ�%Type,
  ��ͥ�绰_In       ������Ϣ.��ͥ�绰%Type,
  ���ڵ�ַ_In       ������Ϣ.���ڵ�ַ%Type,
  ���ڵ�ַ�ʱ�_In   ������Ϣ.���ڵ�ַ�ʱ�%Type,
  ��ϵ������_In     ������Ϣ.��ϵ������%Type,
  ��ϵ�˹�ϵ_In     ������Ϣ.��ϵ�˹�ϵ%Type,
  ��ϵ�˵�ַ_In     ������Ϣ.��ϵ�˵�ַ%Type,
  ��ϵ�˵绰_In     ������Ϣ.��ϵ�˵绰%Type,
  ������λ_In       ������Ϣ.������λ%Type,
  ��ͬ��λid_In     ������Ϣ.��ͬ��λid%Type,
  ��λ�绰_In       ������Ϣ.��λ�绰%Type,
  ��λ�ʱ�_In       ������Ϣ.��λ�ʱ�%Type,
  ��λ������_In     ������Ϣ.��λ������%Type,
  ��λ�ʺ�_In       ������Ϣ.��λ�ʺ�%Type,
  ������_In         ������Ϣ.������%Type,
  ������_In         ������Ϣ.������%Type,
  ��������_In       ������Ϣ.��������%Type,
  ��Ժ����id_In     ������ҳ.��Ժ����id%Type,
  ����ȼ�id_In     ������ҳ.����ȼ�id%Type,
  ��Ժ����_In       ������ҳ.��Ժ����%Type,
  ��Ժ��ʽ_In       ������ҳ.��Ժ��ʽ%Type,
  סԺĿ��_In       ������ҳ.סԺĿ��%Type,
  ����Ժת��_In     ������ҳ.����Ժת��%Type,
  ����ҽʦ_In       ������ҳ.����ҽʦ%Type,
  ����_In           ������Ϣ.����%Type,
  ����_In           ������ҳ.����%Type,
  ��Ժʱ��_In       ������ҳ.��Ժ����%Type,
  �Ƿ����_In       ������ҳ.�Ƿ����%Type,
  ����_In           ������ҳ.��Ժ����%Type,
  ���ʽ_In       ������ҳ.ҽ�Ƹ��ʽ%Type,
  ����id_In         ������ϼ�¼.����id%Type,
  ���id_In         ������ϼ�¼.���id%Type,
  �������_In       ������ϼ�¼.�������%Type,
  ��ҽ����id_In     ������ϼ�¼.����id%Type,
  ��ҽ���id_In     ������ϼ�¼.���id%Type,
  ��ҽ���_In       ������ϼ�¼.�������%Type,
  ����_In           ������ҳ.����%Type,
  ����Ա���_In     ������ҳ.��ĿԱ���%Type,
  ����Ա����_In     ������ҳ.��ĿԱ����%Type,
  �²���_In         Number := 1,
  ��ע_In           ������ҳ.��ע%Type,
  ��Ժ����id_In     ������ҳ.��Ժ����id%Type,
  ����Ժ_In         ������ҳ.����Ժ%Type,
  ��Ժ����_In       ������ҳ.��Ժ����%Type := Null,
  ��ҳid_In         ������ҳ.��ҳid%Type := Null,
  סԺ����_In       ������Ϣ.סԺ����%Type := Null,
  ����֤��_In       ������Ϣ.����֤��%Type := Null,
  ��������_In       ������ҳ.��������%Type := Null,
  ��ϵ�����֤��_In ������Ϣ.��ϵ�����֤��%Type := Null,
  �ֻ���_In         ������Ϣ.�ֻ���%Type := Null,
  �Һ�id_In         ������ҳ.�Һ�id%Type := Null
) As
  -----------------------------------------------------------
  --���ܣ�����Ժ��������һ�Ų�����ҳ��ͬʱ���ܴ�����ơ�
  --������
  --      �Ǽ�ģʽ_IN=0-�����Ǽ�,1-ԤԼ�Ǽ�,2-����ԤԼ(�²���_IN=0)
  --      ��������_IN=��Ӧ"������ҳ.��������"
  --      ����_IN=Null:��ͬʱ���;'��ͥ����':�����ͥ����,��Ϊ��;����:������崲λ��
  --      �²���_IN=��������е����Ĳ�����Ժ,��ò���Ϊ0��ȱʡΪ�²���
  --      ��Ժ����ID_IN=ֻ�е�ʹ��[����������]ģʽ(������99)ʱ,������Ժͬʱ��Ʒִ�ʱ,����ֵ
  --      סԺ��_In = �Ǽ��������۲���ʱ סԺ��_In Ϊ���������
  -----------------------------------------------------------
  v_��ҳid   ������ҳ.��ҳid%Type;
  v_�ȼ�id   ��λ״����¼.�ȼ�id%Type;
  n_סԺ���� ������Ϣ.סԺ����%Type;

  v_�ѱ�      ������ҳ.�ѱ�%Type;
  v_Count     Number;
  n_Uniqueid  Number;
  v_Date      Date;
  d_Indeptime Date;
  v_Error     Varchar2(255);
  Err_Custom Exception;
Begin
  --�жϲ����Ƿ�����
  Select Count(����id) Into v_Count From ������Ϣ Where ����id = ����id_In;
  If v_Count <> 0 Then
    Zl_������Ϣ_�������(����id_In);
  End If;

  Select Sysdate Into v_Date From Dual;
  Zl_������Ǽ�¼_Clear(����id_In);
  
  --���֤�Ų����ڿ�,����ϵͳ�����ж��Ƿ�Ψһ��������
  If ���֤��_In Is Not Null Then
    n_Uniqueid := Nvl(zl_GetSysParameter(279), 0);
    If n_Uniqueid = 1 Then
      Select Count(1) Into v_Count From ������Ϣ Where ���֤�� = ���֤��_In And ����id <> Nvl(����id_In, 0);
      If v_Count <> 0 Then
        v_Error := '�Ѿ��������֤��Ϊ' || ���֤��_In || '�Ĳ���,������¼����ͬ�����֤��!';
        Raise Err_Custom;
      End If;
    End If;
  End If;

  --���˻�����Ϣ
  If ��������_In = 1 Then
    If �²���_In = 1 Then
      Insert Into ������Ϣ
        (����id, �����, סԺ��, ����, �Ա�, ����, �ѱ�, ҽ�Ƹ��ʽ, ��������, ����, ����, ����, ����, ѧ��, ����״��, ְҵ, ���, ���֤��, �����ص�, ��ͥ��ַ, ��ͥ��ַ�ʱ�, ��ͥ�绰,
         ���ڵ�ַ, ���ڵ�ַ�ʱ�, ��ϵ������, ��ϵ�˹�ϵ, ��ϵ�˵�ַ, ��ϵ�˵绰, ������λ, ��ͬ��λid, ��λ�绰, ��λ�ʱ�, ��λ������, ��λ�ʺ�, ������, ������, ��������, ����, �Ǽ�ʱ��, ����֤��, ��������,
         ��ϵ�����֤��, �ֻ���)
      Values
        (����id_In, סԺ��_In, Null, ����_In, �Ա�_In, ����_In, �ѱ�_In, ���ʽ_In, ��������_In, ����_In, ����_In, ����_In, ����_In, ѧ��_In,
         ����״��_In, ְҵ_In, ���_In, ���֤��_In, �����ص�_In, ��ͥ��ַ_In, ��ͥ��ַ�ʱ�_In, ��ͥ�绰_In, ���ڵ�ַ_In, ���ڵ�ַ�ʱ�_In, ��ϵ������_In, ��ϵ�˹�ϵ_In,
         ��ϵ�˵�ַ_In, ��ϵ�˵绰_In, ������λ_In, Decode(��ͬ��λid_In, 0, Null, ��ͬ��λid_In), ��λ�绰_In, ��λ�ʱ�_In, ��λ������_In, ��λ�ʺ�_In, ������_In,
         Decode(������_In, 0, Null, ������_In), ��������_In, ����_In, v_Date, ����֤��_In, ��������_In, ��ϵ�����֤��_In, �ֻ���_In);
    Else
      --�ϲ��˵�����ѱ𲻱�,�������������۲���
      Update ������Ϣ
      Set ����� = סԺ��_In, ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In, �ѱ� = Decode(��������_In, 1, �ѱ�_In, �ѱ�), ҽ�Ƹ��ʽ = ���ʽ_In,
          �������� = ��������_In, ���� = ����_In, ���� = ����_In, ���� = ����_In, ���� = ����_In, ѧ�� = ѧ��_In, ����״�� = ����״��_In, ְҵ = ְҵ_In,
          ��� = ���_In, ���֤�� = ���֤��_In, �����ص� = �����ص�_In, ��ͥ��ַ = ��ͥ��ַ_In, ��ͥ��ַ�ʱ� = ��ͥ��ַ�ʱ�_In, ��ͥ�绰 = ��ͥ�绰_In, ���ڵ�ַ = ���ڵ�ַ_In,
          ���ڵ�ַ�ʱ� = ���ڵ�ַ�ʱ�_In, ��ϵ������ = ��ϵ������_In, ��ϵ�˹�ϵ = ��ϵ�˹�ϵ_In, ��ϵ�˵�ַ = ��ϵ�˵�ַ_In, ��ϵ�˵绰 = ��ϵ�˵绰_In, ������λ = ������λ_In,
          ��ͬ��λid = Decode(��ͬ��λid_In, 0, Null, ��ͬ��λid_In), ��λ�绰 = ��λ�绰_In, ��λ�ʱ� = ��λ�ʱ�_In, ��λ������ = ��λ������_In,
          ��λ�ʺ� = ��λ�ʺ�_In, ������ = ������_In, ������ = Decode(������_In, 0, Null, ������_In), �������� = ��������_In, ���� = ����_In,
          ����֤�� = ����֤��_In, ��������=��������_In, ��ϵ�����֤�� = ��ϵ�����֤��_In, �ֻ��� = Nvl(�ֻ���_In, �ֻ���)
      Where ����id = ����id_In;
    End If;
  Else
    If �²���_In = 1 Then
      Insert Into ������Ϣ
        (����id, סԺ��, ����, �Ա�, ����, �ѱ�, ҽ�Ƹ��ʽ, ��������, ����, ����, ����, ����, ѧ��, ����״��, ְҵ, ���, ���֤��, �����ص�, ��ͥ��ַ, ��ͥ��ַ�ʱ�, ��ͥ�绰,
         ���ڵ�ַ, ���ڵ�ַ�ʱ�, ��ϵ������, ��ϵ�˹�ϵ, ��ϵ�˵�ַ, ��ϵ�˵绰, ������λ, ��ͬ��λid, ��λ�绰, ��λ�ʱ�, ��λ������, ��λ�ʺ�, ������, ������, ��������, ����, �Ǽ�ʱ��, ����֤��, ��������,
         ��ϵ�����֤��, �ֻ���)
      Values
        (����id_In, Decode(��������_In, 2, Null, סԺ��_In), ����_In, �Ա�_In, ����_In, �ѱ�_In, ���ʽ_In, ��������_In, ����_In, ����_In, ����_In,
         ����_In, ѧ��_In, ����״��_In, ְҵ_In, ���_In, ���֤��_In, �����ص�_In, ��ͥ��ַ_In, ��ͥ��ַ�ʱ�_In, ��ͥ�绰_In, ���ڵ�ַ_In, ���ڵ�ַ�ʱ�_In,
         ��ϵ������_In, ��ϵ�˹�ϵ_In, ��ϵ�˵�ַ_In, ��ϵ�˵绰_In, ������λ_In, Decode(��ͬ��λid_In, 0, Null, ��ͬ��λid_In), ��λ�绰_In, ��λ�ʱ�_In,
         ��λ������_In, ��λ�ʺ�_In, ������_In, Decode(������_In, 0, Null, ������_In), ��������_In, ����_In, v_Date, ����֤��_In, ��������_In, ��ϵ�����֤��_In, �ֻ���_In);
    Else
      --�ϲ��˵�����ѱ𲻱�,�������������۲���
      Update ������Ϣ
      Set סԺ�� = Decode(��������_In, 2, סԺ��, Decode(סԺ��_In, Null, סԺ��, סԺ��_In)), ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In,
          �ѱ� = Decode(��������_In, 1, �ѱ�_In, �ѱ�), ҽ�Ƹ��ʽ = ���ʽ_In, �������� = ��������_In, ���� = ����_In, ���� = ����_In, ���� = ����_In,
          ���� = ����_In, ѧ�� = ѧ��_In, ����״�� = ����״��_In, ְҵ = ְҵ_In, ��� = ���_In, ���֤�� = ���֤��_In, �����ص� = �����ص�_In, ��ͥ��ַ = ��ͥ��ַ_In,
          ��ͥ��ַ�ʱ� = ��ͥ��ַ�ʱ�_In, ��ͥ�绰 = ��ͥ�绰_In, ���ڵ�ַ = ���ڵ�ַ_In, ���ڵ�ַ�ʱ� = ���ڵ�ַ�ʱ�_In, ��ϵ������ = ��ϵ������_In, ��ϵ�˹�ϵ = ��ϵ�˹�ϵ_In,
          ��ϵ�˵�ַ = ��ϵ�˵�ַ_In, ��ϵ�˵绰 = ��ϵ�˵绰_In, ������λ = ������λ_In, ��ͬ��λid = Decode(��ͬ��λid_In, 0, Null, ��ͬ��λid_In),
          ��λ�绰 = ��λ�绰_In, ��λ�ʱ� = ��λ�ʱ�_In, ��λ������ = ��λ������_In, ��λ�ʺ� = ��λ�ʺ�_In, ������ = ������_In,
          ������ = Decode(������_In, 0, Null, ������_In), �������� = ��������_In, ���� = ����_In, ����֤�� = ����֤��_In, ��������=��������_In, ��ϵ�����֤�� = ��ϵ�����֤��_In,
          �ֻ��� = Nvl(�ֻ���_In, �ֻ���)
      Where ����id = ����id_In;
    End If;
  End If;

  --������Ϣ
  Begin
    If �Ǽ�ģʽ_In = 1 Then
      v_��ҳid := 0; --ԤԼ�ǼǼ�¼����ҳID=0
    Else
      If ��ҳid_In Is Null Then
        Select Nvl(Max(��ҳid), 0) + 1 Into v_��ҳid From ������ҳ Where ����id = ����id_In And Nvl(��ҳid, 0) <> 0;
      Else
        v_��ҳid := ��ҳid_In;
      End If;
    End If;
  Exception
    When Others Then
      Null;
  End;

  If �Ǽ�ģʽ_In <> 1 Then
    Update ������Ϣ
    Set ��ҳid = v_��ҳid, ��ǰ����id = ��Ժ����id_In, ��ǰ����id = ��Ժ����id_In, ��ǰ���� = Decode(����_In, '��ͥ����', Null, ����_In), ��Ժʱ�� = ��Ժʱ��_In,
        ��Ժʱ�� = Null, ��Ժ = 1
    Where ����id = ����id_In;
  End If;

  --����סԺ����
  If �Ǽ�ģʽ_In <> 1 And ��������_In = 0 Then
    If Nvl(סԺ����_In, 0) = 0 Then
      Select Nvl(סԺ����, 0) + 1 Into n_סԺ���� From ������Ϣ Where ����id = ����id_In;
    Else
      n_סԺ���� := סԺ����_In;
    End If;
    Update ������Ϣ Set סԺ���� = n_סԺ���� Where ����id = ����id_In;
  End If;

  --ȡ���ʱ��
  If ����_In Is Null Then
    d_Indeptime := Null;
  Else
    d_Indeptime := ��Ժʱ��_In;
  End If;

  --״̬��0-������Ժ,1-�ȴ����,2-�ȴ�ת��
  If �Ǽ�ģʽ_In = 2 Then
    --��������ҳ�ӱ�
    Delete From ������ҳ�ӱ� Where ����id = ����id_In And Nvl(��ҳid, 0) = 0;
    --����ԤԼ
    Update ������ҳ
    Set ��ҳid = v_��ҳid, �������� = ��������_In, סԺ�� = Decode(��������_In, 1, Null, 2, Null, סԺ��_In),
        ���ۺ� = Decode(��������_In, 2, סԺ��_In, Null),
        --��ҳID���,�������ʿ��ܱ��
        �ѱ� = �ѱ�_In, ��Ժ����id = ��Ժ����id_In, ��Ժ����id = ��Ժ����id_In, ��Ժ���� = ��Ժʱ��_In, ���ʱ�� = d_Indeptime, ��Ժ���� = ��Ժ����_In,
        ��Ժ��ʽ = ��Ժ��ʽ_In, ��Ժ���� = ��Ժ����_In, ����Ժת�� = ����Ժת��_In, סԺĿ�� = סԺĿ��_In, ��Ժ���� = Decode(����_In, '��ͥ����', Null, ����_In),
        �Ƿ���� = �Ƿ����_In, ��ǰ���� = ��Ժ����_In, ��ǰ����id = ��Ժ����id_In, ����ȼ�id = Decode(����ȼ�id_In, 0, Null, ����ȼ�id_In),
        ��Ժ����id = ��Ժ����id_In, ��Ժ���� = Decode(����_In, '��ͥ����', Null, ����_In), ����ҽʦ = ����ҽʦ_In, ��ĿԱ��� = ����Ա���_In,
        ��ĿԱ���� = ����Ա����_In, ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In, ����״�� = ����״��_In, ְҵ = ְҵ_In, ���� = ����_In, ѧ�� = ѧ��_In,
        ��λ�绰 = ��λ�绰_In, ��λ�ʱ� = ��λ�ʱ�_In, ��λ��ַ = ������λ_In, ���� = ����_In, ��ͥ��ַ = ��ͥ��ַ_In, ��ͥ�绰 = ��ͥ�绰_In, ��ͥ��ַ�ʱ� = ��ͥ��ַ�ʱ�_In,
        ���ڵ�ַ = ���ڵ�ַ_In, ���ڵ�ַ�ʱ� = ���ڵ�ַ�ʱ�_In, ��ϵ������ = ��ϵ������_In, ��ϵ�˹�ϵ = ��ϵ�˹�ϵ_In, ��ϵ�˵�ַ = ��ϵ�˵�ַ_In, ��ϵ�����֤�� = ��ϵ�����֤��_In,
        ��ϵ�˵绰 = ��ϵ�˵绰_In, ҽ�Ƹ��ʽ = ���ʽ_In, ��ע = ��ע_In, ���� = ����_In, ״̬ = Decode(����_In, Null, 1, 0), �Ǽ��� = ����Ա����_In,
        �Ǽ�ʱ�� = v_Date, ����Ժ = ����Ժ_In, �������� = ��������_In
    Where ����id = ����id_In And Nvl(��ҳid, 0) = 0;
    Update ����Ԥ����¼
    Set ��ҳid = ��ҳid_In
    Where ����id = ����id_In And ��ҳid Is Null And ����id = ��Ժ����id_In And Ԥ����� = 2 And ��Ԥ�� Is Null And
          Trunc(�տ�ʱ��) = Trunc(Sysdate);
  Else
    --��Ժ�Ǽǻ�ԤԼ�Ǽ�
    Insert Into ������ҳ
      (��������, ����id, ��ҳid, סԺ��, ���ۺ�, �ѱ�, ��Ժ����id, ��Ժ����id, ��Ժ����, ���ʱ��, ��Ժ����, ��Ժ��ʽ, ��Ժ����, ����Ժת��, סԺĿ��, ��Ժ����, �Ƿ����, ��ǰ����,
       ��ǰ����id, ����ȼ�id, ��Ժ����id, ��Ժ����, ����ҽʦ, ��ĿԱ���, ��ĿԱ����, ״̬, ����, �Ա�, ����, ����״��, ְҵ, ����, ѧ��, ��λ�绰, ��λ�ʱ�, ��λ��ַ, ����, ��ͥ��ַ,
       ��ͥ�绰, ��ͥ��ַ�ʱ�, ���ڵ�ַ, ���ڵ�ַ�ʱ�, ��ϵ������, ��ϵ�˹�ϵ, ��ϵ�˵�ַ, ��ϵ�˵绰, ��ϵ�����֤��, ҽ�Ƹ��ʽ, ����, ��ע, �Ǽ���, �Ǽ�ʱ��, ����Ժ, ��������,�Һ�id)
    Values
      (��������_In, ����id_In, v_��ҳid, Decode(��������_In, 1, Null, 2, Null, סԺ��_In), Decode(��������_In, 2, סԺ��_In, Null), �ѱ�_In,
       ��Ժ����id_In, ��Ժ����id_In, ��Ժʱ��_In, d_Indeptime, ��Ժ����_In, ��Ժ��ʽ_In, ��Ժ����_In, ����Ժת��_In, סԺĿ��_In,
       Decode(����_In, '��ͥ����', Null, ����_In), �Ƿ����_In, ��Ժ����_In, ��Ժ����id_In, Decode(����ȼ�id_In, 0, Null, ����ȼ�id_In), ��Ժ����id_In,
       Decode(����_In, '��ͥ����', Null, ����_In), ����ҽʦ_In, ����Ա���_In, ����Ա����_In, Decode(����_In, Null, 1, 0), ����_In, �Ա�_In, ����_In,
       ����״��_In, ְҵ_In, ����_In, ѧ��_In, ��λ�绰_In, ��λ�ʱ�_In, ������λ_In, ����_In, ��ͥ��ַ_In, ��ͥ�绰_In, ��ͥ��ַ�ʱ�_In, ���ڵ�ַ_In, ���ڵ�ַ�ʱ�_In,
       ��ϵ������_In, ��ϵ�˹�ϵ_In, ��ϵ�˵�ַ_In, ��ϵ�˵绰_In, ��ϵ�����֤��_In, ���ʽ_In, ����_In, ��ע_In, ����Ա����_In, v_Date, ����Ժ_In, ��������_In,�Һ�id_In);
  End If;

  Begin
    If �Ǽ�ģʽ_In <> 1 Then
      Update ��Ժ���� Set ����id = Nvl(��Ժ����id_In, 0), ����id = ��Ժ����id_In Where ����id = ����id_In;
      If Sql%RowCount = 0 Then
        Insert Into ��Ժ����
          (����id, ����id, ����id, ��ҳid)
        Values
          (����id_In, ��Ժ����id_In, Nvl(��Ժ����id_In, 0), Nvl(v_��ҳid, 0));
      End If;
    End If;
  Exception
    When Others Then
      Null;
  End;

  Select �ѱ� Into v_�ѱ� From ������Ϣ Where ����id = ����id_In;
  If v_�ѱ� Is Null Then
    Update ������Ϣ
    Set �ѱ� =
         (Select �ѱ� From ������ҳ Where ����id = ����id_In And ��ҳid = v_��ҳid)
    Where ����id = ����id_In;
  End If;

  --ҽ����
  If �Ǽ�ģʽ_In <> 1 Then
    Select Zl_סԺ�ձ�_Count(��Ժ����id_In, Trunc(��Ժʱ��_In)) Into v_Count From Dual;
    If v_Count > 0 Then
      v_Error := '�Ѳ���ҵ��ʱ���ڵ�סԺ�ձ�,���ܰ����ҵ��!';
      Raise Err_Custom;
    End If;
  
    If ҽ����_In Is Not Null Then
      Insert Into ������ҳ�ӱ� (����id, ��ҳid, ��Ϣ��, ��Ϣֵ) Values (����id_In, v_��ҳid, 'ҽ����', ҽ����_In);
    End If;
  
    --���˱䶯��¼
    --ͬʱ����ҷǼ�ͥ����ʱ�еȼ�
    If ����_In Is Not Null And ����_In <> '��ͥ����' Then
      Select �ȼ�id Into v_�ȼ�id From ��λ״����¼ Where ����id = ��Ժ����id_In And ���� = ����_In;
    End If;
  
    --���ͬʱ���,����Ժ�������д��һ����Ժ�䶯
    Insert Into ���˱䶯��¼
      (ID, ����id, ��ҳid, ��ʼʱ��, ��ʼԭ��, ���Ӵ�λ, ����id, ����id, ����ȼ�id, ��λ�ȼ�id, ����, ����, ����Ա���, ����Ա����)
    Values
      (���˱䶯��¼_Id.Nextval, ����id_In, v_��ҳid, ��Ժʱ��_In, 1, 0, ��Ժ����id_In, ��Ժ����id_In, Decode(����ȼ�id_In, 0, Null, ����ȼ�id_In),
       v_�ȼ�id, Decode(����_In, '��ͥ����', Null, ����_In), ��Ժ����_In, ����Ա���_In, ����Ա����_In);
  
    Insert Into �����Զ�����
      (ID, ����id, ��ҳid, ��ʼʱ��, ��ʼԭ��, ����, ����id, ����id, ����ȼ�id, ����Ա���, ����Ա����)
    Values
      (�����Զ�����_Id.Nextval, ����id_In, v_��ҳid, ��Ժʱ��_In, 1, 1, ��Ժ����id_In, ��Ժ����id_In, Decode(����ȼ�id_In, 0, Null, ����ȼ�id_In),
       ����Ա���_In, ����Ա����_In);
    Insert Into �����Զ�����
      (ID, ����id, ��ҳid, ��ʼʱ��, ��ʼԭ��, ����, ���Ӵ�λ, ����id, ����id, ��λ�ȼ�id, ����, ����Ա���, ����Ա����)
    Values
      (�����Զ�����_Id.Nextval, ����id_In, v_��ҳid, ��Ժʱ��_In, 1, 2, 0, ��Ժ����id_In, ��Ժ����id_In, v_�ȼ�id,
       Decode(����_In, '��ͥ����', Null, ����_In), ����Ա���_In, ����Ա����_In);
    Insert Into �����Զ�����
      (ID, ����id, ��ҳid, ��ʼʱ��, ��ʼԭ��, ����, ���Ӵ�λ, ����id, ����id, ��λ�ȼ�id, ����, ����Ա���, ����Ա����)
    Values
      (�����Զ�����_Id.Nextval, ����id_In, v_��ҳid, ��Ժʱ��_In, 1, 3, 0, ��Ժ����id_In, ��Ժ����id_In, v_�ȼ�id,
       Decode(����_In, '��ͥ����', Null, ����_In), ����Ա���_In, ����Ա����_In);
  
    --ͬʱ����ҷǼ�ͥ����ʱ��λ��ռ��
    If ����_In Is Not Null And ����_In <> '��ͥ����' Then
      Select Count(*) Into v_Count From ��λ״����¼ Where ����id = ��Ժ����id_In And ���� = ����_In And ״̬ = '�մ�';
    
      If v_Count = 0 Then
        v_Error := '����ʧ��,��λ ' || ����_In || ' ���ǿմ���';
        Raise Err_Custom;
      End If;
    
      Update ��λ״����¼
      Set ״̬ = 'ռ��', ����id = ����id_In, ����id = Decode(����, 1, ��Ժ����id_In, ����id)
      Where ����id = ��Ժ����id_In And ���� = ����_In;
    End If;
  
    --������ϼ�¼
    If �������_In Is Not Null Or ����id_In Is Not Null Then
      Insert Into ������ϼ�¼
        (ID, ����id, ��ҳid, ��¼��Դ, �������, ��ϴ���, ����id, ���id, �������, ��¼����, ��¼��)
      Values
        (������ϼ�¼_Id.Nextval, ����id_In, v_��ҳid, 2, 1, 1, ����id_In, ���id_In, �������_In, Sysdate, ����Ա����_In);
    End If;
    If ��ҽ���_In Is Not Null Or ��ҽ����id_In Is Not Null Then
      Insert Into ������ϼ�¼
        (ID, ����id, ��ҳid, ��¼��Դ, �������, ��ϴ���, ����id, ���id, �������, ��¼����, ��¼��)
      Values
        (������ϼ�¼_Id.Nextval, ����id_In, v_��ҳid, 2, 11, 1, ��ҽ����id_In, ��ҽ���id_In, ��ҽ���_In, Sysdate, ����Ա����_In);
    End If;
    --���˵�����¼
    Update ���˵�����¼
    Set ����ʱ�� = Sysdate
    Where ����id = ����id_In And ����ʱ�� Is Not Null And ����ʱ�� > Sysdate;
  
    --���˷���������Ŀ
    If �Ǽ�ģʽ_In <> 1 Then
      Delete From ����������Ŀ Where ����id = ����id_In;
      b_Message.Zlhis_Patient_001(����id_In, v_��ҳid);
    End If;
  
    If �Ǽ�ģʽ_In = 0 And ((�������_In Is Not Null Or ����id_In Is Not Null) Or (��ҽ���_In Is Not Null Or ��ҽ����id_In Is Not Null)) Then
      --����������дʱ��
      Zl_���Ӳ���ʱ��_Insert(����id_In, ��ҳid_In, 2, '���', ��Ժ����id_In, Null, Sysdate, Sysdate);
    End If;
  
    If �Ǽ�ģʽ_In = 0 And ����_In Is Not Null Then
      If ����Ժ_In = 0 Then
        Zl_���Ӳ���ʱ��_Insert(����id_In, ��ҳid_In, 2, '��Ժ', ��Ժ����id_In, Null, ��Ժʱ��_In, ��Ժʱ��_In);
      Else
        Zl_���Ӳ���ʱ��_Insert(����id_In, ��ҳid_In, 2, '�ٴ���Ժ', ��Ժ����id_In, Null, ��Ժʱ��_In, ��Ժʱ��_In);
      End If;
    End If;
  
    If ����_In Is Not Null Then
      --����׷����µ�
      Zl_�������µ�_Newfirst(����id_In, ��ҳid_In, ��Ժ����id_In);
    End If;
  
    --�����������
    Select Count(*) Into v_Count From ������ҳ Where ����id = ����id_In And ��Ժ���� Is Null;
    If v_Count > 1 Then
      v_Error := '���ֲ��˴��ڷǷ��Ĳ�����¼,��ǰ�������ܼ�����' || Chr(13) || Chr(10) || '��������������粢�����������,��ˢ�²���״̬�����ԣ�';
      Raise Err_Custom;
    End If;
  
    Select Count(*)
    Into v_Count
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = v_��ҳid And Nvl(���Ӵ�λ, 0) = 0 And ��ʼʱ�� Is Not Null And ��ֹʱ�� Is Null;
    If v_Count > 1 Then
      v_Error := '���ֲ��˴��ڷǷ��ı䶯��¼,��ǰ�������ܼ�����' || Chr(13) || Chr(10) || '��������������粢�����������,��ˢ�²���״̬�����ԣ�';
      Raise Err_Custom;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Ժ������ҳ_Insert;
/

--122504:������,2018-06-21,������ϵͳZLHIS����޸�
Create Or Replace Procedure Zl_������������Ŀ¼_Update
(
  ϵͳ��ʶ_In In ������������Ŀ¼.ϵͳ��ʶ%Type,
  ��������_In In ������������Ŀ¼.��������%Type,
  �����ַ_In In ������������Ŀ¼.�����ַ%Type
) Is
Begin
  Update ������������Ŀ¼ Set �����ַ = �����ַ_In Where ϵͳ��ʶ = ϵͳ��ʶ_In And �������� = ��������_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������������Ŀ¼_Update;
/

--127450:���ϴ�,2018-06-19,�ҺŰ��Ƚ��ȳ�ԭ��ʹ��Ԥ����
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
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id, Min(Decode(��¼����, 1, �տ�ʱ��, NULL)) as �տ�ʱ��
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

--127450:���ϴ�,2018-06-20,����˿�ʱ�����˿��¼�ĳ�Ԥ����Ϣ�����ⱻ��Ԥ���ٴ�ʹ��
CREATE OR REPLACE Procedure Zl_����Ԥ����¼_Insert
(
  Id_In           ����Ԥ����¼.Id%Type,
  ���ݺ�_In       ����Ԥ����¼.No%Type,
  Ʊ�ݺ�_In       Ʊ��ʹ����ϸ.����%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ��ҳid_In       ����Ԥ����¼.��ҳid%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ���_In         ����Ԥ����¼.���%Type,
  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type,
  �������_In     ����Ԥ����¼.�������%Type,
  �ɿλ_In     ����Ԥ����¼.�ɿλ%Type,
  ��λ������_In   ����Ԥ����¼.��λ������%Type,
  ��λ�ʺ�_In     ����Ԥ����¼.��λ�ʺ�%Type,
  ժҪ_In         ����Ԥ����¼.ժҪ%Type,
  ����Ա���_In   ����Ԥ����¼.����Ա���%Type,
  ����Ա����_In   ����Ԥ����¼.����Ա����%Type,
  ����id_In       Ʊ��ʹ����ϸ.����id%Type,
  Ԥ�����_In     ����Ԥ����¼.Ԥ�����%Type := Null,
  �����id_In     ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In   ����Ԥ����¼.���㿨���%Type := Null,
  ����_In         ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In     ����Ԥ����¼.������λ%Type := Null,
  �տ�ʱ��_In     ����Ԥ����¼.�տ�ʱ��%Type := Null,
  ��������_In     Integer := 0,
  ����id_In       ����Ԥ����¼.����id%Type := Null,
  ��������_In     ����Ԥ����¼.��������%Type := Null,
  �˿���_In     Number := 0,
  ǿ������_In     Number := 0,
  ���½������_In Number := 1,
  �Ƿ�ת��_In     Number := 0
) As
  ----------------------------------------------
  --��������_In:0-������Ԥ��;1-��Ϊ���۵�;3-����˿�
  --����ID_IN:>0ʱ,��ʾĳ�ν���ʱ,ͬ��������Ԥ����¼
  --�˿���_In;0-�����˿����Ƿ�����˲�����1-����˿���
  --���½������_In:0-�� zl_��Ա�ɿ����_Update �и��£�1-�ڱ������и���
  --ǿ������_In:0-��ǿ�ƣ�1-�����������ѿ����������ֵ�ǿ�����ֽ������
  --�Ƿ�ת��_In:0-ԭ���˻����֣�1-ת�˵�֧�ֵ���������

  v_Err_Msg         Varchar2(200);
  Err_Item          Exception;

  v_����            ���㷽ʽ.����%Type;
  v_��ӡid          Ʊ�ݴ�ӡ����.Id%Type;
  v_����            ������Ϣ.��������%Type;
  v_Date            Date;
  n_����ֵ          �������.Ԥ�����%Type;
  n_��id            ����ɿ����.Id%Type;
  n_�������        �������.Ԥ�����%Type;
  n_����Ԥ��        �������.Ԥ�����%Type;
  n_�˿���        ����Ԥ����¼.���%Type;
  n_ʣ���          ����Ԥ����¼.���%Type;
  n_����id          ���˽��ʼ�¼.ID%Type;
  
  Cursor C_��Ԥ�� is
    Select A.NO, A.����id, A.Ԥ�����, A.�����id, A.����, A.������ˮ��, A.����˵��, 0 as ���, A.�տ�ʱ��, A.��� AS Ԥ����
    From ����Ԥ����¼ A Where RowNum < 2;
  r_��Ԥ�� C_��Ԥ��%Rowtype;
  
  Type Ty_ʣ��� Is Ref Cursor;
  C_ʣ��� Ty_ʣ���; --��̬�α���� 
Begin
  v_Date := �տ�ʱ��_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  n_��id := Zl_Get��id(����Ա����_In);

  --����Ԥ���ɿ��¼
  Insert Into ����Ԥ����¼
    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id, Ԥ�����,
     �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, ��Ԥ��, ��������)
  Values
    (Id_In, ���ݺ�_In, Ʊ�ݺ�_In, 1, Decode(��������_In, 1, 0, 1), ����id_In, Decode(��ҳid_In, 0, Null, ��ҳid_In),
     Decode(����id_In, 0, Null, ����id_In), ���_In, ���㷽ʽ_In, �������_In, v_Date, �ɿλ_In, ��λ������_In, ��λ�ʺ�_In, ����Ա���_In, ����Ա����_In,
     ժҪ_In, n_��id, Ԥ�����_In, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, ����id_In,
     Decode(����id_In, Null, Null, 0), ��������_In);
     
  If ��������_In = 1 Then
    --�ݲ�������ܱ�
    Return;
  Elsif ��������_In = 3 Then
    --����һ��ԭԤ��ID�ĳ�����¼��ͬʱҲ����һ������˿�ĳ�����¼
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    IF Nvl(�����id_In, 0) = 0 And Nvl(���㿨���_In, 0) =0 then
      --���֣�������ͨ���㷽ʽ���֡�ǿ�����֡���������������
      Open C_ʣ��� For
           Select A.NO, A.����id, A.Ԥ�����, A.�����id, A.����, A.������ˮ��, A.����˵��, 
                   Min(decode(sign(A.���),-1,0,1)) AS ���, Min(decode(A.��¼����,1,A.�տ�ʱ��,null)) AS �տ�ʱ��,  
                   Nvl(Sum(A.���), 0) - Nvl(Sum(A.��Ԥ��), 0) as Ԥ����
              From ����Ԥ����¼ A, ҽ�ƿ���� B, ���ѿ����Ŀ¼ C
             Where A.����ID = ����id_In And A.��¼���� In (1,11) And A.Ԥ����� = Nvl(Ԥ�����_In, 2)
               And A.�����ID = B.ID(+) And Decode(ǿ������_In, 1, 1, Nvl(B.�Ƿ�����, 1)) = 1
               And A.�����ID = C.���(+) And Decode(ǿ������_In, 1, 1, Nvl(C.�Ƿ�����, 1)) = 1
             Group By A.NO, A.����id, A.Ԥ�����, A.�����id, A.����, A.������ˮ��, A.����˵��
            Having Nvl(Sum(A.���), 0) - Nvl(Sum(A.��Ԥ��), 0) <> 0
             Order By ���,�տ�ʱ��;
    ElsIF Nvl(�Ƿ�ת��_In, 0) = 1 Then
      --ת�ˣ��������������ֻ���ǿ�����֣�����Ŀ��ſ��ܲ���ԭ����,�����ͬ�ֿ�����Ԥ���ɿ��̯
      --Ŀǰֻ֧��ͬһ�ֿ�ת��
      Open C_ʣ��� For
           Select A.NO, A.����id, A.Ԥ�����, A.�����id, A.����, A.������ˮ��, A.����˵��, 
                   Min(decode(sign(A.���),-1,0,1)) AS ���, Min(decode(A.��¼����,1,A.�տ�ʱ��,null)) AS �տ�ʱ��,  
                   Nvl(Sum(A.���), 0) - Nvl(Sum(A.��Ԥ��), 0) as Ԥ����
              From ����Ԥ����¼ A, ҽ�ƿ���� B
             Where A.����ID = ����id_In And A.��¼���� In (1,11) And A.Ԥ����� = Nvl(Ԥ�����_In, 2)
               And A.�����ID = B.ID(+)
               And Nvl(�����id, 0) = Nvl(�����id_In, 0) And Nvl(������ˮ��, '-') = Nvl(������ˮ��_In, '-')
             Group By A.NO, A.����id, A.Ԥ�����, A.�����id, A.����, A.������ˮ��, A.����˵��
            Having Nvl(Sum(A.���), 0) - Nvl(Sum(A.��Ԥ��), 0) <> 0
             Order By ���,�տ�ʱ��;
    Else
      --�����������������ѿ������ݿ����ID�����㿨��š����š�������ˮ��ȱʡԭԤ����¼���������ȷ��Ψһ����з�̯
      Open C_ʣ��� For
           Select A.NO, A.����id, A.Ԥ�����, A.�����id, A.����, A.������ˮ��, A.����˵��, 
                   Min(decode(sign(A.���),-1,0,1)) AS ���, Min(decode(A.��¼����,1,A.�տ�ʱ��,null)) AS �տ�ʱ��,  
                   Nvl(Sum(A.���), 0) - Nvl(Sum(A.��Ԥ��), 0) as Ԥ����
              From ����Ԥ����¼ A
             Where A.����ID = ����id_In And A.��¼���� In (1,11) And A.Ԥ����� = Nvl(Ԥ�����_In, 2)
               And Nvl(A.�����id, 0) = Nvl(�����id_In, 0) And Nvl(A.���㿨���, 0) = Nvl(���㿨���_In, 0) 
               And Nvl(A.����, '-') = Nvl(����_In, '-') And Nvl(������ˮ��, '-') = Nvl(������ˮ��_In, '-')
             Group By A.NO, A.����id, A.Ԥ�����, A.�����id, A.����, A.������ˮ��, A.����˵��
            Having Nvl(Sum(A.���), 0) - Nvl(Sum(A.��Ԥ��), 0) <> 0
             Order By ���,�տ�ʱ��;
    End IF;
    
    n_ʣ��� := -1 * ���_In;
    n_�˿��� := 0;
    Loop
      Fetch C_ʣ���
        Into r_��Ԥ��;
      Exit When C_ʣ���%NotFound;
      IF r_��Ԥ��.NO <> ���ݺ�_In Then
        IF n_ʣ��� > r_��Ԥ��.Ԥ���� then
           n_�˿��� := r_��Ԥ��.Ԥ����;
           n_ʣ��� := n_ʣ��� - n_�˿���;
        Else
           n_�˿��� := n_ʣ���;
           n_ʣ��� := 0;
        End IF;
          	  
        IF nvl(n_�˿���, 0) <> 0 THEN 
          UPDATE ����Ԥ����¼  SET ����ID = n_����id WHERE NO = r_��Ԥ��.NO AND ��¼���� = 1 AND ����ID IS NULL;
          Insert Into ����Ԥ����¼
             (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���,
             �տ�ʱ��, ����Ա����, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, ��Ԥ��, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, 1, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In,
             v_Date, ����Ա����_In, ժҪ, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_����id, n_�˿���, NULL
          From ����Ԥ����¼
          Where NO = r_��Ԥ��.NO And ��¼���� In (1, 11) And RowNum < 2;
        END IF;

        IF n_ʣ��� = 0 Then 
          Exit;
        End IF;
      End IF;
    END LOOP;

    IF n_ʣ��� <> 0 And Nvl(�˿���_In, 0) = 1 THEN 
      v_Err_Msg := '�˿�����ڲ���ʣ��Ԥ����';
      Raise Err_Item;
    END IF;
    
    n_�˿��� := -1 * (-1 * ���_In - n_ʣ���);
    IF n_�˿��� <> 0 Then
      Update ����Ԥ����¼ Set ����id = n_����id Where ID = Id_In;
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id, Ԥ�����,
         �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, ��Ԥ��, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, ���ݺ�_In, Ʊ�ݺ�_In, 11, 1, ����id_In, Decode(��ҳid_In, 0, Null, ��ҳid_In),
         Decode(����id_In, 0, Null, ����id_In), NULL, ���㷽ʽ_In, �������_In, v_Date, �ɿλ_In, ��λ������_In, ��λ�ʺ�_In, ����Ա���_In, ����Ա����_In,
         ժҪ_In, n_��id, Ԥ�����_In, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, n_����id, n_�˿���, NULL);
    End IF;
  End If;

  --����Ʊ��
  If Ʊ�ݺ�_In Is Not Null Then
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into v_��ӡid From Dual;

    --����Ʊ��
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (v_��ӡid, 2, ���ݺ�_In);

    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, 2, Ʊ�ݺ�_In, 1, 1, ����id_In, v_��ӡid, v_Date, ����Ա����_In, ���_In);

    --״̬�Ķ�
    Update Ʊ�����ü�¼
    Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
    Where ID = Nvl(����id_In, 0);
  End If;

  --��ػ��ܱ���

  --�������(Ԥ���������)
  Begin
    Select ���� Into v_���� From ���㷽ʽ Where ���� = ���㷽ʽ_In;
  Exception
    When Others Then
      Null;
  End;
  If Nvl(v_����, 1) <> 5 Then
    Update �������
    Set Ԥ����� = Nvl(Ԥ�����, 0) + ���_In
    Where ���� = 1 And ����id = ����id_In And Nvl(����, 0) = Nvl(Ԥ�����_In, 0)
    Returning Ԥ����� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into �������
        (����id, ����, ����, Ԥ�����, �������)
      Values
        (����id_In, 1, Nvl(Ԥ�����_In, 0), ���_In, 0);
      n_����ֵ := ���_In;
    End If;
    If Nvl(���_In, 0) = 0 Then
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
    End If;
  End If;

  If ���_In < 0 Then
    Begin
      Select Nvl(Ԥ�����, 0) - Nvl(�������, 0)
      Into n_�������
      From �������
      Where ���� = 1 And ����id = ����id_In And Nvl(����, 0) = Nvl(Ԥ�����_In, 0);
    Exception
      When Others Then
        Null;
    End;
    --����˿�Ҫ��������Ԥ���Ƿ�֧������
    If ��������_In = 3 And Nvl(ǿ������_In, 0) = 0 Then
      For c_����Ԥ�� In (Select a.Ԥ��id, a.Ԥ�����, a.�����id, a.���㿨��� As ���ѽӿ�id, Nvl(b.����, c.���) As ����, Nvl(b.����, c.����) As ����,
                            Decode(b.����, Null, c.�Ƿ�ȫ��, b.�Ƿ�ȫ��) As �Ƿ�ȫ��, Decode(b.����, Null, c.�Ƿ�����, b.�Ƿ�����) As �Ƿ�����, a.����,
                            a.������ˮ��, a.����˵��, a.Ԥ�����
                     From (Select a.Ԥ�����, Nvl(a.�����id, 0) As �����id, Nvl(a.���㿨���, 0) As ���㿨���, a.����, a.������ˮ��, a.����˵��,
                                   Max(Decode(Sign(���), -1, Decode(a.��¼״̬, 1, 0, 2, 0, ID), ID)) As Ԥ��id,
                                   Nvl(Sum(���), 0) - Nvl(Sum(Nvl(��Ԥ��, 0)), 0) As Ԥ�����
                            From ����Ԥ����¼ A
                            Where a.����id = ����id_In And (Nvl(a.���㿨���, 0) <> 0 Or Nvl(�����id, 0) <> 0)
                            Group By a.Ԥ�����, Nvl(a.�����id, 0), Nvl(a.���㿨���, 0), a.����, a.������ˮ��, a.����˵��
                            Having Nvl(Sum(���), 0) - Nvl(Sum(Nvl(��Ԥ��, 0)), 0) <> 0) A, ҽ�ƿ���� B, ���ѿ����Ŀ¼ C
                     Where a.Ԥ����� = Nvl(Ԥ�����_In, 0) And a.�����id = b.Id(+) And a.���㿨��� = c.���(+) And Nvl(a.Ԥ�����, 0) <> 0
                     Order By ����, a.����, a.������ˮ��, a.����˵��) Loop

        If Instr(',7,8,', ',' || v_���� || ',') = 0 And Nvl(c_����Ԥ��.�Ƿ�����, 0) = 0 And Nvl(c_����Ԥ��.Ԥ�����, 0) > 0 Then
          n_����Ԥ�� := Nvl(n_����Ԥ��, 0) + Nvl(c_����Ԥ��.Ԥ�����, 0);
        Elsif Instr(',7,8,', ',' || v_���� || ',') > 0 Then
          If Nvl(c_����Ԥ��.����, '0') = Nvl(����_In, '0') And Nvl(c_����Ԥ��.������ˮ��, '0') = Nvl(������ˮ��_In, '0') And
             Nvl(c_����Ԥ��.����˵��, '0') = Nvl(����˵��_In, '0') Then
            n_����Ԥ�� := Nvl(n_����Ԥ��, 0) + Nvl(c_����Ԥ��.Ԥ�����, 0);
          End If;
        End If;
      End Loop;
    End If;

    If Instr(',7,8,', ',' || v_���� || ',') > 0 And Nvl(n_����Ԥ��, 0) < 0 And ��������_In = 3 Then
      v_Err_Msg := '�˿�����ڲ�������Ԥ����';
      Raise Err_Item;
    Elsif Nvl(n_�������, 0) < 0 And �˿���_In = 1 Then
      v_Err_Msg := '�˿�����ڲ���ʣ��Ԥ����';
      Raise Err_Item;
    Elsif Instr(',7,8,', ',' || v_���� || ',') = 0 And Nvl(n_�������, 0) - Nvl(n_����Ԥ��, 0) < 0 And ��������_In = 3 And
          �˿���_In = 1 Then
      v_Err_Msg := '�˿�����ڲ���ʣ��Ԥ����';
      Raise Err_Item;
    End If;
  End If;

  --��Ա�ɿ����(����)
  If Nvl(���½������_In, 1) = 1 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ���_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = ���㷽ʽ_In
    Returning ��� Into n_����ֵ;

    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, ���㷽ʽ_In, 1, ���_In);
      n_����ֵ := ���_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = ���㷽ʽ_In And Nvl(���, 0) = 0;
    End If;
  End If;
  --����ʱ�����Ĵ���
  Select Nvl(��������, 0) Into v_���� From ������Ϣ Where ����id = ����id_In;
  If v_���� = 1 And Nvl(���_In, 0) > 0 Then
    Update ������Ϣ
    Set ������ = Decode(Sign(Nvl(������, 0) - Nvl(���_In, 0)), 1, Nvl(������, 0) - Nvl(���_In, 0), Null),
        ������ = Decode(Sign(Nvl(������, 0) - Nvl(���_In, 0)), 1, ������, Null),
        �������� = Decode(Sign(Nvl(������, 0) - Nvl(���_In, 0)), 1, ��������, Null)
    Where ����id = ����id_In;
  End If;
  If ��������_In <> 1 And ����id_In Is Null Then
    If ���_In < 0 Then
      b_Message.Zlhis_Charge_006(Id_In, ���ݺ�_In);
    Else
      b_Message.Zlhis_Charge_005(Id_In, ���ݺ�_In);
    End If;
    --��Ϣ����;
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 11, Id_In;
    Exception
      When Others Then
        Null;
    End;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ԥ����¼_Insert;
/

--127450:���ϴ�,2018-06-20,�ҺŰ��Ƚ��ȳ�ԭ��ʹ��Ԥ����
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
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id, Min(Decode(��¼����, 1, �տ�ʱ��, NULL)) as �տ�ʱ��
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

--127271:����,2018-06-19,��������ʱ,��ղ��˽��ʼ�¼�е�ʵ��Ʊ��
CREATE OR REPLACE Procedure Zl_���˽����쳣_Update
(
  �Ǽ�ʱ��_In ������ü�¼.�Ǽ�ʱ��%Type,
  ����id_In   ������ü�¼.����id%Type := Null
) As
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  d_Date Date;
  v_No   ������ü�¼.No%Type;
Begin
  --���ܣ������쳣���ݵĵǼ�ʱ�估�տ�ʱ��
  --����ID_IN: ����ʱ,�Խ���ID���и��²���;������NO_IN���в���

  d_Date := �Ǽ�ʱ��_In;
  If d_Date Is Null Then
    d_Date := Sysdate;
  End If;

  --����ָ�����ʵ�������ü�Ԥ���ѵĵǼ�ʱ��
  Update ���˽��ʼ�¼ Set �շ�ʱ�� = d_Date Where ID = ����id_In Returning NO Into v_No;
  Update ���˽��ʼ�¼ Set ʵ��Ʊ�� = Null Where NO = v_No;
  Update ����Ԥ����¼
  Set �տ�ʱ�� = d_Date
  Where ����id = ����id_In And ((��¼���� = 1 And �������� = 12) Or (��¼���� <> 1));

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˽����쳣_Update;
/







------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0017' Where ���=&n_System;
Commit;
