----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;

------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--138776:�ƽ�,2019-03-26,����Ӱ��������ö�Ӧ��ͼ������ʽ
alter table Ӱ����Ļ���� add ͼ������ NUMBER(1);

--138725:����,2019-03-25,�����ɹ�����
Create Table ��ѡҩƷ����(
    ��� NUMBER(4),
    ��ѡҩƷid NUMBER(18),
    ����ҩƷid NUMBER(18))
    TABLESPACE zl9BaseItem;

Alter Table ҩƷ��� Add �Ƿ�����ɹ� number(1);

Alter Table ��ѡҩƷ���� Add Constraint ��ѡҩƷ����_UQ_��ѡҩƷid Unique (��ѡҩƷid,����ҩƷid) Using Index Tablespace zl9Indexhis;

Alter Table ��ѡҩƷ���� Modify ��ѡҩƷid Constraint ��ѡҩƷ����_NN_��ѡҩƷid Not Null;

Alter Table ��ѡҩƷ���� Add Constraint ��ѡҩƷ����_FK_��ѡҩƷid Foreign Key (��ѡҩƷid) References �շ���ĿĿ¼(ID) On Delete Cascade;
Alter Table ��ѡҩƷ���� Add Constraint ��ѡҩƷ����_FK_����ҩƷid Foreign Key (����ҩƷid) References �շ���ĿĿ¼(ID) On Delete Cascade;


------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--138725:����,2019-03-25,�����ɹ�����
Insert into zlTables(ϵͳ,����,��ռ�,����) Values(100,'��ѡҩƷ����','ZL9BASEITEM','A2');

--129946:������,2019-03-22,��Ⱦ�����濨һ��һ��
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1277, 0, 0, 0, 0, 0, 0, 3, '��Ⱦ�����濨һ��һ��', '0', '0',
         '�л����񹲺͹���Ⱦ�����濨����ѡ����ʱ�Ƿ�����ֻѡһ�����֡�', '0-����ѡ������֣�1-ֻ��ѡһ�ֲ���', Null, '������һ�ű��濨��дһ�ִ�Ⱦ�������', Null
  From Dual;

-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--138725:����,2019-03-25,�����ɹ�����
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1023,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select '��ѡҩƷ����','SELECT' From Dual
Union All Select 'Zl_��ѡҩƷ����_Update','EXECUTE' From Dual) A;




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--138776:�ƽ�,2019-03-26,����Ӱ��������ö�Ӧ��ͼ������ʽ
Create Or Replace Procedure Zl_Ӱ����Ļ����_Update
(
  ��Աid_In       In Ӱ����Ļ����.��Աid%Type,
  Ӱ������_In     In Ӱ����Ļ����.Ӱ������%Type,
  �Զ����в���_In In Ӱ����Ļ����.�Զ����в���%Type,
  �Զ�ͼ�񲼾�_In In Ӱ����Ļ����.�Զ�ͼ�񲼾�%Type,
  ��������_In     In Ӱ����Ļ����.��������%Type,
  ��������_In     In Ӱ����Ļ����.��������%Type,
  ͼ������_In     In Ӱ����Ļ����.ͼ������%Type,
  ͼ������_In     In Ӱ����Ļ����.ͼ������%Type,
  �Զ�����_In     In Ӱ����Ļ����.�Զ�����%Type,
  ��ʾ������Ϣ_In In Ӱ����Ļ����.��ʾ������Ϣ%Type,
  ѡ��λ��_In   In Ӱ����Ļ����.ѡ��λ��%Type,
  ѡ������ͬ��_In In Ӱ����Ļ����.ѡ������ͬ��%Type,
  ��ֵģʽ_In     In Ӱ����Ļ����.��ֵģʽ%Type,
  ͼ������_In     In Ӱ����Ļ����.ͼ������%Type
) Is
Begin
  Update Ӱ����Ļ����
  Set �Զ����в��� = �Զ����в���_In, �Զ�ͼ�񲼾� = �Զ�ͼ�񲼾�_In, �������� = ��������_In, �������� = ��������_In, ͼ������ = ͼ������_In, ͼ������ = ͼ������_In,
      �Զ����� = �Զ�����_In, ��ʾ������Ϣ = ��ʾ������Ϣ_In, ѡ��λ�� = ѡ��λ��_In, ѡ������ͬ�� = ѡ������ͬ��_In, ��ֵģʽ = ��ֵģʽ_In, ͼ������ = ͼ������_In
  Where ��Աid = ��Աid_In And Ӱ������ = Ӱ������_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Ӱ����Ļ����_Update;
/

--138960:���ϴ�,2019-03-26,�Һ�ȡ���ų�������
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
    Select ����id, No, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, Id, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id, Min(Decode(��¼����, 1, �տ�ʱ��, Null)) As �տ�ʱ��
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select /*+cardinality(d,10)*/
                                        d.Column_Value
                                       From Table(f_Num2list(v_��Ԥ������ids)) d) And Nvl(Ԥ�����, 2) = 1 Having
     Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By No, ����id
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
  n_�������      ����Ԥ����¼.���%Type;
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
    Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + (1 - 1/24/60/60) And ��� = ����_In 
    And ((״̬ = 3 And ����Ա���� = ����Ա����_In) Or (״̬ = 5 And ����Ա����_In = ����Ա���� And v_������ = ������));
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
                   a.ʧЧʱ�� And a.����id = n_����id) And
            ����ʱ��_In Between Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And ʧЧʱ��;
    
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
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ <> 4;
        Else
          Select Max(Nvl(���, 0)), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Nvl(ԤԼ, 0))
          Into n_�������, n_��������, n_��Լ��
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60;
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
                ״̬ <> 4;
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
        Where a.���� = d_ʱ��ʱ�� And ���� = �ű�_In And ״̬ <> 4;
      Else
        Select Nvl(Max(���), 0)
        Into n_�ҳ���������
        From �Һ����״̬ A
        Where a.���� = d_ʱ��ʱ�� And ���� = �ű�_In;
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
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ <> 4;
        Else
          Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(���� - d_ʱ��ʱ��), 0, 1, 0))
          Into n_�������, n_�ѹ���, n_��������
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60;
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
      Select Nvl(Sum(Nvl(Ԥ�����, 0) - Nvl(�������, 0)), 0)
      Into n_�������
      From �������
      Where ����id In (Select /*+cardinality(d,10)*/
                      d.Column_Value
                     From Table(f_Num2list(v_��Ԥ������ids)) d) And Nvl(����, 0) = 1 And Nvl(����, 0) = 1;
      if n_������� < Ԥ��֧��_In Then
        v_Err_Msg := '���˵ĵ�ǰԤ�����Ϊ ' || Ltrim(To_Char(n_�������, '9999999990.00')) || '��С�ڱ���֧����� ' ||
                     Ltrim(To_Char(Ԥ��֧��_In, '9999999990.00')) || '��֧��ʧ�ܣ�';
        Raise Err_Item;
      End if;
      
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

--138960:���ϴ�,2019-03-26,�Һ�ȡ���ų�������
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
    Select ����id, No, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, Id, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id, Min(Decode(��¼����, 1, �տ�ʱ��, Null)) As �տ�ʱ��
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select /*+cardinality(d,10)*/
                                        d.Column_Value
                                       From Table(f_Num2list(v_��Ԥ������ids)) d) And Nvl(Ԥ�����, 2) = 1 Having
     Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By No, ����id
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
  n_�������      ����Ԥ����¼.���%Type;
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
  If Nvl(���_In, 0) = 1 Then
    If �����¼id_In Is Not Null Then
      Begin
        Select 1
        Into n_Exists
        From �ٴ������¼ a, �ٴ������Դ b
        Where a.Id = �����¼id_In And a.��Դid = b.Id And b.���� = �ű�_In And a.����id = ִ�в���id_In And Nvl(a.�Ƿ񷢲�, 0) = 1 And
              Nvl(a.�Ƿ�����, 0) = 0;
      Exception
        When Others Then
          v_Err_Msg := '�޷�ȷ�������¼����������¼�Ƿ���ڻ�������';
          Raise Err_Item;
      End;
    End If;
  End if;

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
    Where ��¼id = �����¼id_In And ��� = ����_In And 
          ((Nvl(�Һ�״̬, 0) = 3 And ����Ա���� = ����Ա����_In) Or (Nvl(�Һ�״̬, 0) = 5 And ����Ա����_In = ����Ա���� And v_������ = ����վ����));
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
          Where ��¼id = �����¼id_In And �Һ�״̬ Not In (0, 4);
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
      Where ��¼id = �����¼id_In And ԤԼ˳��� Is Null And �Һ�״̬ <> 0;
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
        Where ��¼id = �����¼id_In And �Һ�״̬ Not In (0, 4);
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
      Select Nvl(Sum(Nvl(Ԥ�����, 0) - Nvl(�������, 0)), 0)
      Into n_�������
      From �������
      Where ����id In (Select /*+cardinality(d,10)*/
                      d.Column_Value
                     From Table(f_Num2list(v_��Ԥ������ids)) d) And Nvl(����, 0) = 1 And Nvl(����, 0) = 1;
      if n_������� < Ԥ��֧��_In Then
        v_Err_Msg := '���˵ĵ�ǰԤ�����Ϊ ' || Ltrim(To_Char(n_�������, '9999999990.00')) || '��С�ڱ���֧����� ' ||
                     Ltrim(To_Char(Ԥ��֧��_In, '9999999990.00')) || '��֧��ʧ�ܣ�';
        Raise Err_Item;
      End if;
      
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

--138725:����,2019-03-25,�����ɹ�ҩƷ
Create Or Replace Procedure Zl_��ѡҩƷ����_Update(��������_In In Varchar2 --��"|"�ָ��Ķ����������ݣ�ÿ����¼�� ���^ ��ѡҩƷID^ ����ҩƷID
                                             ) Is
  v_Records    Varchar2(4000);
  v_Currrec    Varchar2(1000);
  v_Fields     Varchar2(1000);
  n_���       ��ѡҩƷ����.���%Type;
  n_��ѡҩƷid ��ѡҩƷ����.��ѡҩƷid%Type;
  n_����ҩƷid ��ѡҩƷ����.����ҩƷid%Type;
Begin

  Delete From ��ѡҩƷ����;

  If ��������_In Is Null Then
    v_Records := Null;
  Else
    v_Records := ��������_In;
  End If;
  While v_Records Is Not Null Loop
    v_Currrec    := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
    v_Fields     := v_Currrec;
    n_���       := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    n_��ѡҩƷid := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    n_����ҩƷid := To_Number(v_Fields);
  
    Insert Into ��ѡҩƷ���� (���, ��ѡҩƷid, ����ҩƷid) Values (n_���, n_��ѡҩƷid, n_����ҩƷid);
    v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��ѡҩƷ����_Update;
/

Create Or Replace Procedure Zl_��ѡҩƷ����_Delete
(
  ��ѡҩƷid_In In ��ѡҩƷ����.��ѡҩƷid%Type
) Is
Begin
  Delete From ��ѡҩƷ���� Where ��ѡҩƷid = ��ѡҩƷid_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��ѡҩƷ����_Delete;
/

--138725:����,2019-03-25,���Ӵ��Ρ��Ƿ�����ɹ���
Create Or Replace Procedure Zl_��ҩ���_Update
(
  ҩƷid_In         In ҩƷ���.ҩƷid%Type,
  ����_In           In �շ���ĿĿ¼.����%Type,
  ���_In           In �շ���ĿĿ¼.���%Type,
  ����_In           In �շ���ĿĿ¼.����%Type := Null,
  Ʒ��_In           In �շ���Ŀ����.����%Type := Null,
  ƴ��_In           In �շ���Ŀ����.����%Type := Null,
  ���_In           In �շ���Ŀ����.����%Type := Null,
  ������_In         In �շ���Ŀ����.����%Type := Null,
  ��ʶ��_In         In ҩƷ���.��ʶ��%Type := Null,
  ҩƷ��Դ_In       In ҩƷ���.ҩƷ��Դ%Type := Null,
  ��׼�ĺ�_In       In ҩƷ���.��׼�ĺ�%Type := Null,
  ע���̱�_In       In ҩƷ���.ע���̱�%Type := Null,
  �ۼ۵�λ_In       In �շ���ĿĿ¼.���㵥λ%Type := Null,
  ����ϵ��_In       In ҩƷ���.����ϵ��%Type := Null,
  ���ﵥλ_In       In ҩƷ���.���ﵥλ%Type := Null,
  �����װ_In       In ҩƷ���.�����װ%Type := Null,
  סԺ��λ_In       In ҩƷ���.סԺ��λ%Type := Null,
  סԺ��װ_In       In ҩƷ���.סԺ��װ%Type := Null,
  ҩ�ⵥλ_In       In ҩƷ���.ҩ�ⵥλ%Type := Null,
  ҩ���װ_In       In ҩƷ���.ҩ���װ%Type := Null,
  ���쵥λ_In       In ҩƷ���.���쵥λ%Type := 1,
  ���췧ֵ_In       In ҩƷ���.���췧ֵ%Type := Null,
  �Ƿ���_In       In �շ���ĿĿ¼.�Ƿ���%Type := Null,
  ָ��������_In     In ҩƷ���.ָ��������%Type := Null,
  ����_In           In ҩƷ���.����%Type := 95,
  ָ�����ۼ�_In     In ҩƷ���.ָ�����ۼ�%Type := Null,
  �ӳ���_In         In ҩƷ���.�ӳ���%Type := Null,
  ����ѱ���_In     In ҩƷ���.����ѱ���%Type := Null,
  ҩ�ۼ���_In       In ҩƷ���.ҩ�ۼ���%Type := Null,
  ��������_In       In �շ���ĿĿ¼.��������%Type := Null,
  �������_In       In �շ���ĿĿ¼.�������%Type := Null,
  Gmp��֤_In        In ҩƷ���.Gmp��֤%Type := 0,
  �б�ҩƷ_In       In ҩƷ���.�б�ҩƷ%Type := 0,
  ���ηѱ�_In       In �շ���ĿĿ¼.���ηѱ�%Type := 0,
  סԺ�ɷ����_In   In ҩƷ���.סԺ�ɷ����%Type := 0,
  ҩ�����_In       In ҩƷ���.ҩ�����%Type := Null,
  ҩ������_In       In ҩƷ���.ҩ������%Type := Null,
  ���Ч��_In       In ҩƷ���.���Ч��%Type := Null,
  ���������_In     In ҩƷ���.���������%Type := 0,
  �ɱ���_In         In ҩƷ���.�ɱ���%Type := 0,
  ��ǰ�ۼ�_In       In �շѼ�Ŀ.�ּ�%Type := 0,
  ����id_In         In �շѼ�Ŀ.������Ŀid%Type := Null,
  ��ͬ��λid_In     In ҩƷ���.��ͬ��λid%Type := Null,
  ˵��_In           In �շ���ĿĿ¼.˵��%Type := Null,
  ��̬����_In       In ҩƷ���.��̬����%Type := 0,
  ��ҩ����_In       In ҩƷ���.��ҩ����%Type := Null,
  ��ѡ��_In         In �շ���ĿĿ¼.��ѡ��%Type := Null,
  ��ֵ˰��_In       In ҩƷ���.��ֵ˰��%Type := Null,
  ����ҩ��_In       In ҩƷ���.����ҩ��%Type := Null,
  վ��_In           In �շ���ĿĿ¼.վ��%Type := Null,
  �Ƿ񳣱�_In       In ҩƷ���.�Ƿ񳣱�%Type := Null,
  �洢�¶�_In       In ��ҺҩƷ����.�洢�¶�%Type := Null,
  �洢����_In       In ��ҺҩƷ����.�洢����%Type := Null,
  ��ҩ����_In       In ��ҺҩƷ����.��ҩ����%Type := Null,
  �Ƿ�������_In   In ��ҺҩƷ����.�Ƿ�������%Type := Null,
  ����_In           In ҩƷ���.����%Type := Null,
  ������Ŀ_In       In �շ���ĿĿ¼.������Ŀ%Type := Null,
  ����ɷ����_In   In ҩƷ���.����ɷ����%Type := 0,
  Dddֵ_In          In ҩƷ���.Dddֵ%Type := 0,
  ��ΣҩƷ_In       ҩƷ���.��ΣҩƷ%Type := Null,
  �ͻ���λ_In       In ҩƷ���.�ͻ���λ%Type := Null,
  �ͻ���װ_In       In ҩƷ���.�ͻ���װ%Type := Null,
  ��Һע������_In   In ��ҺҩƷ����.��Һע������%Type := Null,
  �Ƿ��ҩ_In       In ҩƷ���.�Ƿ��ҩ%Type := Null,
  �Ƿ����۹���_In In ҩƷ���.�Ƿ����۹���%Type := Null,
  ��λ��_In         In ҩƷ���.��λ��%Type := Null,
  �Ƿ���������_In   In ҩƷ���.�Ƿ���������%Type := Null,
  �����ɹ�_In       In ҩƷ���.�Ƿ�����ɹ�%Type := Null
) Is
  v_ҩ��id   ������ĿĿ¼.Id%Type;
  v_����     ������ĿĿ¼.����%Type;
  v_�Ƿ��� �շ���ĿĿ¼.�Ƿ���%Type; --������ҩƷ��ʱ��Ϊʱ�ۣ�ʱ��ҩƷֻ����δ������������޸�Ϊ���ۣ���������������޸Ķ������� 
  v_����     Number(2);
  Err_Notfind Exception;
  v_No           �շѼ�Ŀ.No%Type;
  v_Temp         �շ���ĿĿ¼.������Ŀ%Type;
  v_������Ŀ     �շ���ĿĿ¼.������Ŀ%Type;
  n_ָ�������   ҩƷ���.ָ�������%Type;
  n_ҩƷ�ϴ��ۼ� ҩƷ���.�ϴ��ۼ�%Type;
  n_���۽��     ҩƷ���.ʵ�ʽ��%Type;
  n_�շ�id       ҩƷ�շ���¼.Id%Type;
  n_��ͨ���С�� Number;
  n_���         Number(8);
  Classid        Number(18); --������
  v_Billno       ҩƷ�շ���¼.No%Type; --���۵���
  n_�۸�id       �շѼ�Ŀ.Id%Type;
  n_�շѼ�Ŀ�ּ� �շѼ�Ŀ.�ּ�%Type;
  n_ԭ��         ҩƷ�۸��¼.ԭ��%Type;
  n_ҩƷ�۸��¼ Number(1);
  v_���         �շ���ĿĿ¼.���%Type;
  --����->ʱ�ۺ����ҩƷ�۸��¼��ֵ

  Cursor c_Priceadjust Is
    Select s.ҩƷid, s.�ⷿid, Nvl(s.����, 0) As ����, s.�ϴι�Ӧ��id As ��Ӧ��id, s.�ϴ����� As ����, s.Ч��, s.�ϴβ��� As ����,
           Nvl(s.ʵ������, 0) As ʵ������, s.�ϴο��� As ����, Nvl(s.ʵ�ʽ��, 0) As ʵ�ʽ��, Nvl(s.ʵ�ʲ��, 0) As ʵ�ʲ��, Nvl(s.���ۼ�, 0) As ���ۼ�,
           s.ƽ���ɱ���, s.���Ч��, s.��׼�ĺ�, s.�ϴ��������� As ��������
    From ҩƷ��� S
    Where s.ҩƷid = ҩƷid_In And s.���� = 1 
    Order By s.ҩƷid, s.����, s.�ⷿid;

  r_Priceadjust c_Priceadjust%RowType;

Begin
  v_������Ŀ := ������Ŀ_In;
  --�жϲ�����Ŀ 
  If v_������Ŀ Is Null Then
    If ����id_In Is Not Null Then
      Begin
        Select ������Ŀ Into v_Temp From ������Ŀ Where ID = ����id_In;
      Exception
        When Others Then
          v_Temp := Null;
      End;
      If v_Temp Is Not Null Then
        v_������Ŀ := v_Temp;
      End If;
    End If;
  End If;
  --ͨ������ 
  Select ID, ����
  Into v_ҩ��id, v_����
  From ������ĿĿ¼
  Where ID = (Select ҩ��id From ҩƷ��� Where ҩƷid = ҩƷid_In);
  --ȡԭʼ�Ķ������� 
  Select �Ƿ��� Into v_�Ƿ��� From �շ���ĿĿ¼ Where ID = ҩƷid_In;
  --�����Ϣ 
  Update �շ���ĿĿ¼
  Set ���� = ����_In, ���� = v_����, ��� = ���_In, ���� = ����_In, ���㵥λ = �ۼ۵�λ_In, �������� = ��������_In, ������� = �������_In, ���ηѱ� = ���ηѱ�_In,
      ������Ŀ = v_������Ŀ, ˵�� = ˵��_In, ��ѡ�� = ��ѡ��_In, վ�� = վ��_In
  Where ID = ҩƷid_In
  Returning ��� Into v_���;
  If Sql%RowCount = 0 Then
    Raise Err_Notfind;
  End If;
  n_ָ������� := (1 - 1 / (1 + �ӳ���_In / 100)) * 100;
  Update ҩƷ���
  Set ��ʶ�� = ��ʶ��_In, ҩƷ��Դ = ҩƷ��Դ_In, ��׼�ĺ� = ��׼�ĺ�_In, ע���̱� = ע���̱�_In, ����ϵ�� = ����ϵ��_In, ���ﵥλ = ���ﵥλ_In, �����װ = �����װ_In,
      סԺ��λ = סԺ��λ_In, סԺ��װ = סԺ��װ_In, ҩ�ⵥλ = ҩ�ⵥλ_In, ҩ���װ = ҩ���װ_In, ���쵥λ = ���쵥λ_In, ���췧ֵ = ���췧ֵ_In, ָ�������� = ָ��������_In,
      ���� = ����_In, ָ�����ۼ� = ָ�����ۼ�_In, ָ������� = n_ָ�������, ����ѱ��� = ����ѱ���_In, ҩ�ۼ��� = ҩ�ۼ���_In, סԺ�ɷ���� = סԺ�ɷ����_In,
      ҩ����� = ҩ�����_In, ҩ������ = ҩ������_In, ���Ч�� = ���Ч��_In, �б�ҩƷ = �б�ҩƷ_In, Gmp��֤ = Gmp��֤_In, ��������� = ���������_In,
      ��ͬ��λid = ��ͬ��λid_In, ��̬���� = ��̬����_In, ��ҩ���� = ��ҩ����_In, ��ֵ˰�� = ��ֵ˰��_In, ����ҩ�� = ����ҩ��_In, �Ƿ񳣱� = �Ƿ񳣱�_In, ���� = ����_In,
      ����ɷ���� = ����ɷ����_In, Dddֵ = Dddֵ_In, ��ΣҩƷ = ��ΣҩƷ_In, �ͻ���λ = �ͻ���λ_In, �ͻ���װ = �ͻ���װ_In, �ӳ��� = �ӳ���_In, �Ƿ��ҩ = �Ƿ��ҩ_In,
      �Ƿ����۹��� = �Ƿ����۹���_In, ��λ�� = ��λ��_In, �Ƿ��������� = �Ƿ���������_In, �Ƿ�����ɹ� = �����ɹ�_In
  Where ҩƷid = ҩƷid_In;

  --���û�д����ɹ����ԣ���ɾ����ѡҩƷ����(���ܶ��ձ���Ҳû�����ݣ���ִ�иù���û��Ӱ��)
  If Nvl(�����ɹ�_In, 0) = 0 Then
    Zl_��ѡҩƷ����_Delete(ҩƷid_In);
  End If;

  --�����Ĵ��� 
  If ������_In Is Null Then
    Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 1 And ���� = 3;
  Else
    Update �շ���Ŀ���� Set ���� = v_����, ���� = ������_In Where �շ�ϸĿid = ҩƷid_In And ���� = 1 And ���� = 3;
    If Sql%RowCount = 0 Then
      Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, v_����, 1, ������_In, 3);
    End If;
  End If;
  If Ʒ��_In Is Null Then
    Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 3;
  Else
    If ƴ��_In Is Null Then
      Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 1;
    Else
      Update �շ���Ŀ���� Set ���� = Ʒ��_In, ���� = ƴ��_In Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 1;
      If Sql%RowCount = 0 Then
        Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, Ʒ��_In, 3, ƴ��_In, 1);
      End If;
    End If;
    If ���_In Is Null Then
      Delete �շ���Ŀ���� Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 2;
    Else
      Update �շ���Ŀ���� Set ���� = Ʒ��_In, ���� = ���_In Where �շ�ϸĿid = ҩƷid_In And ���� = 3 And ���� = 2;
      If Sql%RowCount = 0 Then
        Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, Ʒ��_In, 3, ���_In, 2);
      End If;
    End If;
  End If;

  --������Ϣ������Ѿ��з�����������ֱ�Ӹ�����Щ��Ϣ 
  Select Nvl(Count(*), 0) Into v_���� From ҩƷ�շ���¼ Where ҩƷid = ҩƷid_In And Rownum < 2;
  If v_���� = 0 Then
    Update ҩƷ��� Set �ɱ��� = �ɱ���_In Where ҩƷid = ҩƷid_In;
    If ����id_In Is Not Null Then
      Update �շѼ�Ŀ
      Set �ּ� = ��ǰ�ۼ�_In, ������Ŀid = ����id_In, �䶯ԭ�� = 1, ����˵�� = '�޸Ķ���', ������ = User
      Where �շ�ϸĿid = ҩƷid_In And (Sysdate Between ִ������ And ��ֹ���� Or Sysdate >= ִ������ And ��ֹ���� Is Null) And �䶯ԭ�� = 1;
      If Sql%RowCount = 0 Then
        v_No := Nextno(9);
        Insert Into �շѼ�Ŀ
          (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
        Values
          (�շѼ�Ŀ_Id.Nextval, Null, ҩƷid_In, 0, ��ǰ�ۼ�_In, ����id_In, 1, '��������', User, Sysdate,
           To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
      End If;
    End If;
  Else
    --����ҵ������ʱ�������޸ļ۸��ǿ����޸�������Ŀ 
    Update �շѼ�Ŀ
    Set ������Ŀid = ����id_In
    Where �շ�ϸĿid = ҩƷid_In And (Sysdate Between ִ������ And ��ֹ���� Or Sysdate >= ִ������ And ��ֹ���� Is Null) And �䶯ԭ�� = 1;
  End If;

  --ʱ��->����
  If v_�Ƿ��� = 1 And �Ƿ���_In = 0 Then
    Update �շ���ĿĿ¼ Set �Ƿ��� = �Ƿ���_In Where ID = ҩƷid_In;
  
    Begin
      Select �ּ�, ID As �۸�id
      Into n_�շѼ�Ŀ�ּ�, n_�۸�id
      From �շѼ�Ŀ
      Where �շ�ϸĿid = ҩƷid_In And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And �䶯ԭ�� = 1;
    Exception
      When Others Then
        n_�շѼ�Ŀ�ּ� := Null;
        n_�۸�id       := Null;
    End;
  
    Begin
      Select �ϴ��ۼ� Into n_ҩƷ�ϴ��ۼ� From ҩƷ��� Where ҩƷid = ҩƷid_In;
    Exception
      When Others Then
        n_ҩƷ�ϴ��ۼ� := Null;
    End;
  
    If n_ҩƷ�ϴ��ۼ� Is Null Then
      n_ҩƷ�ϴ��ۼ� := n_�շѼ�Ŀ�ּ�;
    End If;
  
    Zl_�շѼ�Ŀ_Stop(ҩƷid_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into �շѼ�Ŀ
      (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
    Values
      (�շѼ�Ŀ_Id.Nextval, n_�۸�id, ҩƷid_In, n_�շѼ�Ŀ�ּ�, n_ҩƷ�ϴ��ۼ�, ����id_In, 1, 'ʱ��ת����', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    Select ���� Into n_��ͨ���С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 4 And ��λ = 5;
  
    --ȡ������ID
    Select ���id Into Classid From ҩƷ�������� Where ���� = 13;
  
    n_���   := 0;
    v_Billno := Null;
  
    For r_Priceadjust In c_Priceadjust Loop
      If n_ҩƷ�ϴ��ۼ� <> r_Priceadjust.���ۼ� Then
        If v_Billno Is Null Then
          Select Nextno(147) Into v_Billno From Dual;
        End If;
        n_��� := n_��� + 1;
        Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
        n_���۽�� := Round(n_ҩƷ�ϴ��ۼ� * r_Priceadjust.ʵ������, n_��ͨ���С��) -
                  Round(r_Priceadjust.���ۼ� * r_Priceadjust.ʵ������, n_��ͨ���С��);
        --��������Ӱ���¼
        Insert Into ҩƷ�շ���¼
          (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ���۽��, ���, ժҪ, ������,
           ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ����, Ƶ��, ��ҩ��λid)
        Values
          (n_�շ�id, 1, 13, v_Billno, n_���, Classid, r_Priceadjust.ҩƷid, r_Priceadjust.����, r_Priceadjust.����,
           r_Priceadjust.Ч��, r_Priceadjust.����, 1, r_Priceadjust.ʵ������, 0, r_Priceadjust.���ۼ�, 0, n_ҩƷ�ϴ��ۼ�,
           r_Priceadjust.����, n_���۽��, n_���۽��, 'ʱ��ת����', Zl_Username, Sysdate, r_Priceadjust.�ⷿid, 1, n_�۸�id, Zl_Username,
           Sysdate, r_Priceadjust.ʵ�ʽ��, r_Priceadjust.ʵ�ʲ��, r_Priceadjust.��Ӧ��id);
      
        Zl_ҩƷ���_Update(n_�շ�id, 2, 0);
      End If;
    End Loop;
  
    --����->ʱ��
  Elsif v_�Ƿ��� = 0 And �Ƿ���_In = 1 Then
    Update �շ���ĿĿ¼ Set �Ƿ��� = �Ƿ���_In Where ID = ҩƷid_In;
    Begin
      Select �ּ�, ID As �۸�id
      Into n_�շѼ�Ŀ�ּ�, n_�۸�id
      From �շѼ�Ŀ
      Where �շ�ϸĿid = ҩƷid_In And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And �䶯ԭ�� = 1;
    Exception
      When Others Then
        n_�շѼ�Ŀ�ּ� := Null;
        n_�۸�id       := Null;
    End;
  
    Zl_�շѼ�Ŀ_Stop(ҩƷid_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into �շѼ�Ŀ
      (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
    Values
      (�շѼ�Ŀ_Id.Nextval, n_�۸�id, ҩƷid_In, n_�շѼ�Ŀ�ּ�, n_�շѼ�Ŀ�ּ�, ����id_In, 1, '����תʱ��', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    For r_Priceadjust In c_Priceadjust Loop
      n_ҩƷ�۸��¼ := 0;
      Begin
        Select 1, �ּ�
        Into n_ҩƷ�۸��¼, n_ԭ��
        From ҩƷ�۸��¼
        Where ҩƷid = r_Priceadjust.ҩƷid And �ⷿid = r_Priceadjust.�ⷿid And Nvl(����, 0) = r_Priceadjust.���� And ��¼״̬ = 1 And
              �۸����� = 1;
      Exception
        When Others Then
          n_ҩƷ�۸��¼ := 0;
          n_ԭ��         := n_�շѼ�Ŀ�ּ�;
      End;
    
      If n_ҩƷ�۸��¼ = 1 Then
        Zl_ҩƷ�۸��¼_Stop(1, r_Priceadjust.�ⷿid, r_Priceadjust.ҩƷid, r_Priceadjust.����, Sysdate - 1 / 24 / 60 / 60, 2);
      End If;
      Zl_ҩƷ�۸��¼_Insert(0, 1, r_Priceadjust.�ⷿid, r_Priceadjust.ҩƷid, r_Priceadjust.����, n_ԭ��, n_�շѼ�Ŀ�ּ�, Sysdate, '����תʱ��',
                       Zl_Username, Null, r_Priceadjust.��Ӧ��id, r_Priceadjust.����, r_Priceadjust.Ч��, r_Priceadjust.����,
                       r_Priceadjust.���Ч��, Null, Null, Null, Null, 1);
    
      Update ҩƷ���
      Set ���ۼ� = n_�շѼ�Ŀ�ּ�
      Where ���� = 1 And �ⷿid = r_Priceadjust.�ⷿid And ҩƷid = r_Priceadjust.ҩƷid And Nvl(����, 0) = r_Priceadjust.����;
    
    End Loop;
  End If;

  --ҩƷ�����̱Ƚ����� 
  If ����_In Is Not Null Then
    Update ҩƷ������ Set ���� = ����_In Where ���� = ����_In;
    If Sql%RowCount = 0 Then
      Insert Into ҩƷ������
        (����, ����, ����)
        Select Nvl(Max(To_Number(����)), 0) + 1, ����_In, zlSpellCode(����_In, 10) From ҩƷ������;
    End If;
  End If;

  --�޸���ҺҩƷ���� 
  Update ��ҺҩƷ����
  Set �洢�¶� = �洢�¶�_In, �洢���� = �洢����_In, ��ҩ���� = ��ҩ����_In, �Ƿ������� = �Ƿ�������_In, ��Һע������ = ��Һע������_In
  Where ҩƷid = ҩƷid_In;

  If Sql%NotFound Then
    Insert Into ��ҺҩƷ����
      (ҩƷid, �洢�¶�, �洢����, ��ҩ����, �Ƿ�������, ��Һע������)
    Values
      (ҩƷid_In, �洢�¶�_In, �洢����_In, ��ҩ����_In, �Ƿ�������_In, ��Һע������_In);
  End If;

  --ҩƷ���ȵ���(����ģʽʱ)
  Zl_ҩƷ���ľ���_���۵���;

  b_Message.Zlhis_Dict_036(v_���, ҩƷid_In);
Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]�ù�񲻴��ڣ������ѱ������û�ɾ����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��ҩ���_Update;
/

--138725:����,2019-03-25,���Ӵ��Ρ�,�Ƿ�����ɹ���
Create Or Replace Procedure Zl_��ҩ���_Insert
(
  ҩ��id_In         In ҩƷ���.ҩ��id%Type,
  ҩƷid_In         In ҩƷ���.ҩƷid%Type,
  ����_In           In �շ���ĿĿ¼.����%Type,
  ���_In           In �շ���ĿĿ¼.���%Type,
  ����_In           In �շ���ĿĿ¼.����%Type := Null,
  Ʒ��_In           In �շ���Ŀ����.����%Type := Null,
  ƴ��_In           In �շ���Ŀ����.����%Type := Null,
  ���_In           In �շ���Ŀ����.����%Type := Null,
  ������_In         In �շ���Ŀ����.����%Type := Null,
  ��ʶ��_In         In ҩƷ���.��ʶ��%Type := Null,
  ҩƷ��Դ_In       In ҩƷ���.ҩƷ��Դ%Type := Null,
  ��׼�ĺ�_In       In ҩƷ���.��׼�ĺ�%Type := Null,
  ע���̱�_In       In ҩƷ���.ע���̱�%Type := Null,
  �ۼ۵�λ_In       In �շ���ĿĿ¼.���㵥λ%Type := Null,
  ����ϵ��_In       In ҩƷ���.����ϵ��%Type := Null,
  ���ﵥλ_In       In ҩƷ���.���ﵥλ%Type := Null,
  �����װ_In       In ҩƷ���.�����װ%Type := Null,
  סԺ��λ_In       In ҩƷ���.סԺ��λ%Type := Null,
  סԺ��װ_In       In ҩƷ���.סԺ��װ%Type := Null,
  ҩ�ⵥλ_In       In ҩƷ���.ҩ�ⵥλ%Type := Null,
  ҩ���װ_In       In ҩƷ���.ҩ���װ%Type := Null,
  ���쵥λ_In       In ҩƷ���.���쵥λ%Type := 1,
  ���췧ֵ_In       In ҩƷ���.���췧ֵ%Type := Null,
  �Ƿ���_In       In �շ���ĿĿ¼.�Ƿ���%Type := Null,
  ָ��������_In     In ҩƷ���.ָ��������%Type := Null,
  ����_In           In ҩƷ���.����%Type := 95,
  ָ�����ۼ�_In     In ҩƷ���.ָ�����ۼ�%Type := Null,
  �ӳ���_In         In ҩƷ���.�ӳ���%Type := Null,
  ����ѱ���_In     In ҩƷ���.����ѱ���%Type := Null,
  ҩ�ۼ���_In       In ҩƷ���.ҩ�ۼ���%Type := Null,
  ��������_In       In �շ���ĿĿ¼.��������%Type := Null,
  �������_In       In �շ���ĿĿ¼.�������%Type := Null,
  Gmp��֤_In        In ҩƷ���.Gmp��֤%Type := 0,
  �б�ҩƷ_In       In ҩƷ���.�б�ҩƷ%Type := 0,
  ���ηѱ�_In       In �շ���ĿĿ¼.���ηѱ�%Type := 0,
  סԺ�ɷ����_In   In ҩƷ���.סԺ�ɷ����%Type := 0,
  ҩ�����_In       In ҩƷ���.ҩ�����%Type := Null,
  ҩ������_In       In ҩƷ���.ҩ������%Type := Null,
  ���Ч��_In       In ҩƷ���.���Ч��%Type := Null,
  ���������_In     In ҩƷ���.���������%Type := 0,
  �ɱ���_In         In ҩƷ���.�ɱ���%Type := 0,
  ��ǰ�ۼ�_In       In �շѼ�Ŀ.�ּ�%Type := 0,
  ����id_In         In �շѼ�Ŀ.������Ŀid%Type := Null,
  ��ͬ��λid_In     In ҩƷ���.��ͬ��λid%Type := Null,
  ˵��_In           In �շ���ĿĿ¼.˵��%Type := Null,
  ��̬����_In       In ҩƷ���.��̬����%Type := 0,
  ��ҩ����_In       In ҩƷ���.��ҩ����%Type := Null,
  ��ѡ��_In         In �շ���ĿĿ¼.��ѡ��%Type := Null,
  ��ֵ˰��_In       In ҩƷ���.��ֵ˰��%Type := Null,
  ����ҩ��_In       In ҩƷ���.����ҩ��%Type := Null,
  վ��_In           In �շ���ĿĿ¼.վ��%Type := Null,
  �Ƿ񳣱�_In       In ҩƷ���.�Ƿ񳣱�%Type := Null,
  �洢�¶�_In       In ��ҺҩƷ����.�洢�¶�%Type := Null,
  �洢����_In       In ��ҺҩƷ����.�洢����%Type := Null,
  ��ҩ����_In       In ��ҺҩƷ����.��ҩ����%Type := Null,
  �Ƿ�������_In   In ��ҺҩƷ����.�Ƿ�������%Type := Null,
  ����_In           In ҩƷ���.����%Type := Null,
  ������Ŀ_In       In �շ���ĿĿ¼.������Ŀ%Type := Null,
  ����ɷ����_In   In ҩƷ���.����ɷ����%Type := 0,
  Dddֵ_In          In ҩƷ���.Dddֵ%Type := 0,
  ��ΣҩƷ_In       In ҩƷ���.��ΣҩƷ%Type := Null,
  �ͻ���λ_In       In ҩƷ���.�ͻ���λ%Type := Null,
  �ͻ���װ_In       In ҩƷ���.�ͻ���װ%Type := Null,
  ��Һע������_In   In ��ҺҩƷ����.��Һע������%Type := Null,
  �Ƿ��ҩ_In       In ҩƷ���.�Ƿ��ҩ%Type := Null,
  �Ƿ����۹���_In In ҩƷ���.�Ƿ����۹���%Type := Null,
  ��λ��_In         In ҩƷ���.��λ��%Type := Null,
  �Ƿ���������_In   In ҩƷ���.�Ƿ���������%Type := Null,
  �����ɹ�_In       In ҩƷ���.�Ƿ�����ɹ�%Type := Null
) Is

  v_���       ������ĿĿ¼.���%Type;
  v_����       ������ĿĿ¼.����%Type;
  v_Kind       Varchar2(20);
  v_No         �շѼ�Ŀ.No%Type;
  v_Temp       �շ���ĿĿ¼.������Ŀ%Type;
  v_������Ŀ   �շ���ĿĿ¼.������Ŀ%Type;
  n_ָ������� ҩƷ���.ָ�������%Type;

  --�̵�ⷿ�Ĺ������� 
  Cursor c_Storageid Is
    Select Distinct ����id From ��������˵�� Where �������� Like v_Kind Or �������� = '�Ƽ���';
  r_Storageid c_Storageid%RowType;
Begin
  v_������Ŀ := ������Ŀ_In;
  --�жϲ�����Ŀ 
  If v_������Ŀ Is Null Then
    If ����id_In Is Not Null Then
      Begin
        Select ������Ŀ Into v_Temp From ������Ŀ Where ID = ����id_In;
      Exception
        When Others Then
          v_Temp := Null;
      End;
      If v_Temp Is Not Null Then
        v_������Ŀ := v_Temp;
      End If;
    End If;
  End If;
  --�������� 
  Select ���, ���� Into v_���, v_���� From ������ĿĿ¼ Where ID = ҩ��id_In;
  n_ָ������� := (1 - 1 / (1 + �ӳ���_In / 100)) * 100;
  --�����Ϣ 
  Insert Into �շ���ĿĿ¼
    (���, ID, ����, ����, ���, ����, ���㵥λ, ��������, �������, ���ηѱ�, �Ƿ���, ����ʱ��, ����ʱ��, ˵��, ��ѡ��, վ��, ������Ŀ)
  Values
    (v_���, ҩƷid_In, ����_In, v_����, ���_In, ����_In, �ۼ۵�λ_In, ��������_In, �������_In, ���ηѱ�_In, �Ƿ���_In, Sysdate,
     To_Date('3000-01-01', 'YYYY-MM-DD'), ˵��_In, ��ѡ��_In, վ��_In, v_������Ŀ);
  Insert Into ҩƷ���
    (ҩ��id, ҩƷid, ��ʶ��, ҩƷ��Դ, ��׼�ĺ�, ע���̱�, ����ϵ��, ���ﵥλ, �����װ, סԺ��λ, סԺ��װ, ҩ�ⵥλ, ҩ���װ, ���쵥λ, ���췧ֵ, ָ��������, ����, ָ�����ۼ�, ָ�������,
     ����ѱ���, ҩ�ۼ���, �ɱ���, Gmp��֤, �б�ҩƷ, ���������, סԺ�ɷ����, ҩ�����, ҩ������, ���Ч��, ��ͬ��λid, ��̬����, ��ҩ����, ��ֵ˰��, ����ҩ��, �Ƿ񳣱�, ����, ����ɷ����,
     Dddֵ, ��ΣҩƷ, �ͻ���λ, �ͻ���װ, �ӳ���, �Ƿ��ҩ, �Ƿ����۹���, ��λ��, �Ƿ���������, �Ƿ�����ɹ�)
  Values
    (ҩ��id_In, ҩƷid_In, ��ʶ��_In, ҩƷ��Դ_In, ��׼�ĺ�_In, ע���̱�_In, ����ϵ��_In, ���ﵥλ_In, �����װ_In, סԺ��λ_In, סԺ��װ_In, ҩ�ⵥλ_In, ҩ���װ_In,
     ���쵥λ_In, ���췧ֵ_In, ָ��������_In, ����_In, ָ�����ۼ�_In, n_ָ�������, ����ѱ���_In, ҩ�ۼ���_In, �ɱ���_In, Gmp��֤_In, �б�ҩƷ_In, ���������_In,
     סԺ�ɷ����_In, ҩ�����_In, ҩ������_In, ���Ч��_In, ��ͬ��λid_In, ��̬����_In, ��ҩ����_In, ��ֵ˰��_In, ����ҩ��_In, �Ƿ񳣱�_In, ����_In, ����ɷ����_In,
     Dddֵ_In, ��ΣҩƷ_In, �ͻ���λ_In, �ͻ���װ_In, �ӳ���_In, �Ƿ��ҩ_In, �Ƿ����۹���_In, ��λ��_In, �Ƿ���������_In, �����ɹ�_In);

  --�����޸ģ�����ҩƷ������ҩ���г�ҩ��ʱ��ȱʡ�������Ϊ�����סԺ����˽������ҩƷʱ�����ٸ��ݹ��ҩƷ�ķ���������ҩƷ�ķ������ 
  --������Ŀ�������ĸ��� 
  --select nvl(sum(distinct I.�������),0) into v_���� 
  --from �շ���ĿĿ¼ I,ҩƷ��� S 
  --where I.ID=S.ҩƷID and S.ҩ��ID=ҩ��ID_IN; 
  --update ������ĿĿ¼ 
  --set �������=decode(v_����,0,0,1,1,2,2,3) 
  --where ID=ҩ��ID_IN; 

  --�����Ĵ��� 
  Insert Into �շ���Ŀ����
    (�շ�ϸĿid, ����, ����, ����, ����)
    Select ҩƷid_In, ����, ����, ����, ���� From ������Ŀ���� Where ������Ŀid = ҩ��id_In;
  If ������_In Is Not Null Then
    Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, v_����, 1, ������_In, 3);
  End If;
  If (Ʒ��_In Is Not Null) And (ƴ��_In Is Not Null) Then
    Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, Ʒ��_In, 3, ƴ��_In, 1);
  End If;
  If (Ʒ��_In Is Not Null) And (���_In Is Not Null) Then
    Insert Into �շ���Ŀ���� (�շ�ϸĿid, ����, ����, ����, ����) Values (ҩƷid_In, Ʒ��_In, 3, ���_In, 2);
  End If;

  --������Ϣ 
  If ����id_In Is Not Null Then
    v_No := Nextno(9);
    Insert Into �շѼ�Ŀ
      (ID, ԭ��id, �շ�ϸĿid, ԭ��, �ּ�, ������Ŀid, �䶯ԭ��, ����˵��, ������, ִ������, ��ֹ����, NO, ���)
    Values
      (�շѼ�Ŀ_Id.Nextval, Null, ҩƷid_In, 0, ��ǰ�ۼ�_In, ����id_In, 1, '��������', User, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  End If;

  --ҩƷ�����̱Ƚ����� 
  If ����_In Is Not Null Then
    Update ҩƷ������ Set ���� = ����_In Where ���� = ����_In;
    If Sql%RowCount = 0 Then
      Insert Into ҩƷ������
        (����, ����, ����)
        Select Nvl(Max(To_Number(����)), 0) + 1, ����_In, zlSpellCode(����_In, 10) From ҩƷ������;
    End If;
  End If;

  --����ù��ķ������ 
  Insert Into �շ�ִ�п���
    (�շ�ϸĿid, ������Դ, ��������id, ִ�п���id)
    Select ҩƷid_In, ������Դ, ��������id, ִ�п���id From ����ִ�п��� Where ������Ŀid = ҩ��id_In;

  --�����̵����� 

  If v_��� = 5 Then
    v_Kind := '��ҩ%';
  Else
    v_Kind := '��ҩ%';
  End If;

  For r_Storageid In c_Storageid Loop
    Insert Into ҩƷ�����޶�
      (�ⷿid, ҩƷid, ����, ����, �̵�����, �ⷿ��λ)
    Values
      (r_Storageid.����id, ҩƷid_In, 0, 0, '1111', Null);
  End Loop;

  --������ҺҩƷ���� 
  Insert Into ��ҺҩƷ����
    (ҩƷid, �洢�¶�, �洢����, ��ҩ����, �Ƿ�������, ��Һע������)
  Values
    (ҩƷid_In, �洢�¶�_In, �洢����_In, ��ҩ����_In, �Ƿ�������_In, ��Һע������_In);

  --ҩƷ���ȵ���(����ģʽʱ)
  Zl_ҩƷ���ľ���_���۵���;

  b_Message.Zlhis_Dict_035(v_���, ҩƷid_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��ҩ���_Insert;
/





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0055' Where ���=&n_System;
Commit;