----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--140325:����,2019-05-06,Ӧ����¼���Ӷ��ⵥ�ݺŵ�����
Create Index Ӧ����¼_IX_��ⵥ�ݺ� On Ӧ����¼(��ⵥ�ݺ�) Tablespace zl9Indexhis;




------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--140678:��˶,2019-05-07,�ϻ���Ա�䶯��Ϣ֪ͨ
Insert Into Zlmsg_Lists(Bz_Type, Code, Name, Key_Define, Note)
Select '������', 'ZLTOOLS_USERS_001', '�ϻ���Աɾ��', '<root><�û���></�û���><��ԱID></��ԱID></root>', '�û���Ȩ����:�޸��û���ɾ���û�ʱ'  From Dual Union All 
Select '������', 'ZLTOOLS_USERS_002', '�ϻ���Ա����', '<root><�û���></�û���><��ԱID></��ԱID></root>', '�û���Ȩ����:�����û����޸��û�ʱ;���������û�'  From Dual;


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--139097:������,2019-04-28,�������۲��˴���
Create Or Replace Procedure Zl_����ҽ����¼_�ջ�
(
  --���ܣ���ָ��ҽ�����ڷ��Ͳ����ջء�����ϴη���û�в������ã�����ջ�ҽ�����ϴ�ִ��ʱ�䡣
  --������
  --      �ջ���_IN=����ҩ���г�ҩΪ��סԺ��λ���ջ���,����ҩΪ�ջظ���,������ҽ��Ϊ�ջ������������
  --      ҽ��ID_IN=ÿ��Ҫ�ջص�ҽ����¼��ID(��ϸ�洢��ID),�Գ�ҩ���䷽,��һ��������ҩ;�����÷��巨(����Ϊ������δ��ȡ)
  --      �ϴ�ʱ��_IN=ҽ�����ڷ��Ͳ����ջغ�Ӧ�û�ԭ���ϴ�ִ��ʱ��(�ϸ�Ƶ�ʼ������),Ϊ��ʱ��ʾ��ȫ���ջ��ˡ�
  --      NO_IN=���ջ�Ҫ�����������ü�¼ʱ��Ϊ�����ɼ�¼�ĵ��ݺ�(�����ü�ҩƷʹ��),��ǰ�����ֻ����NO��һ���ݡ�
  --            ��ΪҩƷ���ܷ���,��������ڴ���ʱȡ��
  --            ���ȫ�ǻ��۵�������ֵΪ���������۵������򲻲����������ݣ�ֱ���޸Ļ�ɾ�����۵�
  �ջ���_In     In ����ҽ������.��������%Type,
  ҽ��id_In     In ����ҽ����¼.Id%Type,
  �ϴ�ʱ��_In   In ����ҽ����¼.�ϴ�ִ��ʱ��%Type,
  �ջ�ʱ��_In   In ����ҽ����¼.�ϴ�ִ��ʱ��%Type,
  No_In         In סԺ���ü�¼.No%Type := Null,
  ����Ա���_In In ��Ա��.���%Type := Null,
  ����Ա����_In In ��Ա��.����%Type := Null
) Is
  --�ջ�ҽ����Ӧ�ķ��ͷ�����ϸ��ʣ������,��������ķ������ջ�
  --ʣ������û���ſ���������������ݣ��ڲ���������ʱ����ԭ��������
  --��ҩƷ�����ģ���һ�����������ܴ���δִ�к���ִ�в��֣���ֱ���д�����¼������δִ������
  --ִ�б�־=0-δִ��,1-��ִ�У�ҩƷ���в���ִ�У����շ���¼�е���ϸ������Ϊ׼����ҩƷ��ֻ���ȴ���δִ�е�
  Cursor c_Detail Is
    Select a.����id, a.No, a.���, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.��������, a.�������, a.ҽ������, a.��������, a.ʣ������, a.��ִ����, a.δִ����,
           a.ִ�б�־, a.��¼״̬, a.�Ǽ�ʱ��, a.�շѷ�ʽ
    From (With ҽ�����ü�¼ As (Select Max(Decode(b.��¼״̬, 2, 0, b.Id)) As ����id, b.No, Nvl(b.�۸񸸺�, b.���) As ���, b.�շ�ϸĿid,
                                 b.���˲���id, Sum(Nvl(b.����, 1) * b.����) As ʣ������, b.�շ����, Max(Nvl(b.ִ��״̬, 0)) As ִ��״̬, d.��������,
                                 c.�������, c.ҽ������, c.��������, Max(b.��¼״̬) As ��¼״̬, Max(b.�Ǽ�ʱ��) As �Ǽ�ʱ��, Nvl(e.�շѷ�ʽ, 0) As �շѷ�ʽ
                          From ����ҽ������ A, סԺ���ü�¼ B, ����ҽ����¼ C, �������� D, ����ҽ���Ƽ� E
                          Where a.ҽ��id = ҽ��id_In And a.No = b.No And a.��¼���� = b.��¼���� And a.ҽ��id = b.ҽ����� And
                                b.�۸񸸺� Is Null And b.�շ�ϸĿid = d.����id(+) And c.Id = ҽ��id_In And e.ҽ��id(+) = b.ҽ����� And
                                e.�շ�ϸĿid(+) = b.�շ�ϸĿid And Not Exists
                           (Select 1 From ��Һ��ҩ��¼ F Where f.ҽ��id = c.���id And a.���ͺ� = f.���ͺ�)
                          Group By b.No, b.��¼����, Nvl(b.�۸񸸺�, b.���), b.�շ�ϸĿid, b.���˲���id, b.�շ����, d.��������, c.�������, c.ҽ������,
                                   c.��������, e.�շѷ�ʽ
                          Having Sum(Nvl(b.����, 1) * b.����) > 0)
           Select ����id, NO, ���, �շ�ϸĿid, ���˲���id, �շ����, ��������, �������, ҽ������, ��������, ʣ������, Null As ��ִ����, Null As δִ����,
                  ִ��״̬ As ִ�б�־, ��¼״̬, �Ǽ�ʱ��, �շѷ�ʽ
           From ҽ�����ü�¼
           Where �շ���� Not In ('5', '6', '7') And Not (�շ���� = '4' And Nvl(��������, 0) = 1)
           Union All
           Select a.����id, a.No, a.���, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.��������, a.�������, a.ҽ������, a.��������, a.ʣ������, 0 As ��ִ����,
                  Sum(Nvl(b.����, 1) * b.ʵ������) As δִ����, 0 As ִ�б�־, a.��¼״̬, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��, a.�շѷ�ʽ
           From ҽ�����ü�¼ A, ҩƷ�շ���¼ B
           Where (a.�շ���� In ('5', '6', '7') Or (a.�շ���� = '4' And Nvl(a.��������, 0) = 1)) And a.����id = b.����id And
                 a.No = b.No And b.���� In (9, 10, 25, 26) And Mod(b.��¼״̬, 3) = 1 And b.����� Is Null
           Group By a.����id, a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.��������, a.�������, a.ҽ������, a.��������, a.ʣ������,
                    a.�շѷ�ʽ
           Having Sum(Nvl(b.����, 1) * b.ʵ������) > 0
           Union All
           Select a.����id, a.No, a.���, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.��������, a.�������, a.ҽ������, a.��������, a.ʣ������,
                  Sum(Nvl(b.����, 1) * b.ʵ������) As ��ִ����, 0 As δִ����, 1 As ִ�б�־, a.��¼״̬, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��, a.�շѷ�ʽ
           From ҽ�����ü�¼ A, ҩƷ�շ���¼ B
           Where (a.�շ���� In ('5', '6', '7') Or (a.�շ���� = '4' And Nvl(a.��������, 0) = 1)) And a.����id = b.����id And
                 a.No = b.No And b.���� In (9, 10, 25, 26) And Not (Mod(b.��¼״̬, 3) = 1 And b.����� Is Null)
           Group By a.����id, a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.��������, a.�������, a.ҽ������, a.��������, a.ʣ������,
                    a.�շѷ�ʽ
           Having Sum(Nvl(b.����, 1) * b.ʵ������) > 0) A
           Order By Decode(a.�������, '5', 0, '6', 0, '7', 0, a.�շ�ϸĿid), a.ִ�б�־, a.�Ǽ�ʱ�� Desc;


  Cursor c_Applay(v_����ids Varchar2) Is
    Select a.����id, b.No, b.���, a.����, a.����ʱ��, a.�������
    From ���˷������� A, סԺ���ü�¼ B
    Where a.����id = b.Id And a.���벿��id = a.��˲���id And a.����ʱ�� = �ջ�ʱ��_In And
          a.����id In (Select * From Table(Cast(f_Num2list(v_����ids) As Zltools.t_Numlist)))
    Order By NO, ���;

  --����ָ��ҩƷ��������ʱ��������ط��ü�ҩƷ/���ļ�¼��Ϣ(���η����ж�����¼,���������ڽ����ֹ)
  --ҩƷҽ����д��"����ҽ������"��¼,��Ӧ�ĸ�ҩ;����һ����д�˵�(����Ϊ����),��NO��ͬ��
  --��ΪҪ�ջصĴ������ܰ����˶�η��͵�����,����Ҫ����η��͵��շ���¼��ȡ��������η���ʱ�����۵����ջأ��޸Ļ�ɾ����
  Cursor c_Drug Is
    Select a.����id, a.��ҳid, d.����, Nvl(Nvl(x.����ϵ��, y.����ϵ��), 1) As ����ϵ��, Nvl(x.סԺ��װ, 1) As סԺ��װ,
           Nvl(x.���Ч��, y.���Ч��) As ���Ч��, Nvl(b.����, 1) * b.ʵ������ As ����, b.Id As �շ�id, b.����, b.ҩƷid, b.�Է�����id, b.�ⷿid,
           b.����id, Nvl(Nvl(x.ҩ������, y.���÷���), 0) As ����, b.����, b.����, b.Ч��, a.��¼״̬, a.No, a.���, a.�շ�ϸĿid, a.ִ��״̬ As ִ�б�־
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ����ҽ������ C, ������Ϣ D, ҩƷ��� X, �������� Y
    Where c.ҽ��id = ҽ��id_In And a.No = c.No And a.��¼���� = c.��¼���� And a.��¼״̬ In (0, 1, 3) And a.ҽ����� + 0 = ҽ��id_In And
          a.No = b.No And a.Id = b.����id + 0 And b.���� In (9, 10, 25, 26) And (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0) And
          a.����id = d.����id And b.ҩƷid = x.ҩƷid(+) And b.ҩƷid = y.����id(+)
    Order By a.��¼״̬, b.No Desc, b.Id Desc;

  --������ҩ����(����ҩ;��)����ʱ�������ķ���(����������ж�����¼)
  --�Է�ҩҽ��,ֱ���ջ�ָ����,���ܶ�η���(�����η��ͼ۸�ͬ,���ջصļ۸��������εģ���Ȼ��Ҫ���ݶ���������μ��ջ���)��
  --���ı������ۼ۵�λ������סԺ��λת��
  --��ҩ��������д�˷��ͼ�¼(�����˶���������ȼ�)
  --һ��ֻ��һ�λ�һ�η���ֻ��һ�ε���Ŀ��ʱ��֧�ָ�������
  Cursor c_Other(n_���ͺ� ����ҽ������.���ͺ�%Type) Is
    With ҽ�����ü�¼ As
     (Select a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.Id As ����id, a.���� As ʣ������, Nvl(a.ִ��״̬, 0) As ִ��״̬, a.ҽ�����, b.���ͺ�,
             c.���� As ��������, Nvl(c.�շѷ�ʽ, 0) As �շѷ�ʽ, a.�շ����
      From סԺ���ü�¼ A, ����ҽ������ B, ����ҽ���Ƽ� C
      Where a.No = b.No And a.��¼���� = b.��¼���� And a.ҽ����� + 0 = b.ҽ��id And b.ҽ��id = ҽ��id_In And a.ҽ����� = c.ҽ��id(+) And
            a.�շ�ϸĿid = c.�շ�ϸĿid(+))
    Select a.No, a.���, a.����id, a.ʣ������, a.�շ�ϸĿid, a.��¼״̬, a.ִ��״̬, a.��������, a.�շѷ�ʽ, a.�շ����
    From (Select a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.����id, a.ʣ������, a.��������, a.ִ��״̬, a.ҽ�����, a.�շѷ�ʽ, a.�շ����
           From ҽ�����ü�¼ A
           Where a.��¼״̬ In (1, 3) And a.���ͺ� = n_���ͺ�
           Union All
           Select a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.����id, a.ʣ������, a.��������, a.ִ��״̬, a.ҽ�����, a.�շѷ�ʽ, a.�շ����
           From ҽ�����ü�¼ A
           Where a.��¼״̬ = 0) A
    Order By a.�շ�ϸĿid, a.���, a.��¼״̬;

  --�����������Ϊ�˲����¼�¼ʱ,��дͬһ�շ�ϸĿ�Ĳ�ͬ������Ŀ�ļ۸񸸺�

  --���α����ڴ��������ػ��ܱ�
  Cursor c_Money
  (
    v_Start סԺ���ü�¼.���%Type,
    v_End   סԺ���ü�¼.���%Type
  ) Is
    Select ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Sum(Nvl(Ӧ�ս��, 0)) As Ӧ�ս��, Sum(Nvl(ʵ�ս��, 0)) As ʵ�ս��
    From סԺ���ü�¼
    Where ��¼���� = 2 And ��¼״̬ = 1 And NO = No_In And ��� Between v_Start And v_End
    Group By ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid;

  --ϵͳ����ָ��ִ�к���Ҫ�Զ���˵Ļ��۷��ã����ڷ�ҩҽ����������Ӧ��ҩƷ�����ķ���
  Cursor c_Verify
  (
    v_Start סԺ���ü�¼.���%Type,
    v_End   סԺ���ü�¼.���%Type
  ) Is
    Select NO, ���
    From סԺ���ü�¼
    Where ��¼���� = 2 And ��¼״̬ = 0 And NO = No_In And �۸񸸺� Is Null And ��� Between v_Start And v_End;

  Cursor c_Compound
  (
    ���id_In       ����ҽ����¼.���id%Type,
    ִ����ֹʱ��_In ����ҽ����¼.ִ����ֹʱ��%Type,
    ��ҩid_In       ��Һ��ҩ��¼.Id%Type,
    ҽ�����_In     ����ҽ����¼.Id%Type
  ) Is
    Select b.����id, b.ҩƷid As �շ�ϸĿid, Sum(a.����) As ����, c.סԺ��װ, c.סԺ��λ, d.����, e.���˲���id, e.����״̬, e.Id As ��ҩid, f.No,
           Nvl(f.�۸񸸺�, f.���) As ���, f.��¼״̬ As ��¼״̬, f.ִ��״̬ As ִ�б�־
    From ��Һ��ҩ���� A, ҩƷ�շ���¼ B, ҩƷ��� C, �շ���ĿĿ¼ D, ��Һ��ҩ��¼ E, סԺ���ü�¼ F
    Where a.�շ�id = b.Id And b.ҩƷid = c.ҩƷid And c.ҩƷid = d.Id And e.Id = a.��¼id And f.No = b.No And f.Id = b.����id And
          e.ҽ��id = ���id_In And e.ִ��ʱ�� > ִ����ֹʱ��_In And e.Id = ��ҩid_In And f.ҽ����� + 0 = ҽ�����_In
    Group By b.����id, b.ҩƷid, c.סԺ��װ, c.סԺ��λ, d.����, e.���˲���id, e.����״̬, e.Id, f.No, f.�۸񸸺�, f.���, f.��¼״̬, f.ִ��״̬;

  --����а����������õ�δ������ʱ�����ݲ��������Ƿ��Զ�����
  Cursor c_Stuff Is
    Select m.Id, m.�ⷿid, Decode(b.���÷���, 1, m.����, 0) As ����
    From ҩƷ�շ���¼ M, סԺ���ü�¼ A, �������� B
    Where m.No = No_In And m.���� In (25, 26) And m.�ⷿid Is Not Null And m.��¼״̬ = 1 And m.����� Is Null And m.No = a.No And
          a.Id = m.����id + 0 And a.��¼���� = 2 And a.��¼״̬ = 1 And a.�շ�ϸĿid = b.����id And b.�������� = 1
    Order By m.�ⷿid, m.ҩƷid;

  v_Dec      Number;
  v_First    Number;
  v_������� Varchar2(255);

  v_������� ����ҽ����¼.�������%Type;
  v_�������� ����ҽ����¼.��������%Type;
  v_�������� ��������.��������%Type;

  v_������� סԺ���ü�¼.���%Type;
  v_�շ���� ҩƷ�շ���¼.���%Type;
  v_����id   סԺ���ü�¼.Id%Type;
  v_ʵ�ս�� סԺ���ü�¼.ʵ�ս��%Type;

  v_��ʼ��� סԺ���ü�¼.���%Type;
  v_������� סԺ���ü�¼.���%Type;

  v_ҽ��ִ�� ����ҽ������.ִ��״̬%Type;

  v_����ϵ�� ҩƷ���.����ϵ��%Type;
  v_סԺ��װ ҩƷ���.סԺ��װ%Type;
  v_ҽ������ ����ҽ����¼.ҽ������%Type;

  v_���ʲ���       Zlparameters.����ֵ%Type;
  v_��Һҩ�������� Zlparameters.����ֵ%Type;
  v_���ʽ��       סԺ���ü�¼.���ʽ��%Type;

  v_�շ�ϸĿid   סԺ���ü�¼.�շ�ϸĿid%Type;
  v_ʣ������     סԺ���ü�¼.����%Type;
  v_�ջ�����     סԺ���ü�¼.����%Type;
  v_��ǰ����     סԺ���ü�¼.����%Type;
  v_��ǰ����     סԺ���ü�¼.����%Type;
  v_����ids      Varchar2(4000);
  v_��id         ����ҽ����¼.Id%Type;
  v_��������     ����ҽ���Ƽ�.����%Type;
  v_�ջ���       סԺ���ü�¼.����%Type;
  v_�ջ�ʣ��     סԺ���ü�¼.����%Type;
  v_��Һ�ջ�ʣ�� סԺ���ü�¼.����%Type;
  n_����         ҩƷ�շ���¼.��д����%Type;

  v_Delno    Varchar2(4000);
  v_Temp     Varchar2(4000);
  v_�շ����� Varchar2(4000);
  v_No       סԺ���ü�¼.No%Type;
  v_��Ա��� סԺ���ü�¼.����Ա���%Type;
  v_��Ա���� סԺ���ü�¼.����Ա����%Type;

  n_���id       ����ҽ����¼.���id%Type;
  d_ִ����ֹʱ�� ����ҽ����¼.ִ����ֹʱ��%Type;
  b_��Һ��ҩ��¼ Boolean;
  d_�ջ�ʱ��     ����ҽ����¼.ִ����ֹʱ��%Type;
  n_�������     ���˷�������.�������%Type;
  v_����ԭ��     ���˷�������.����ԭ��%Type;
  n_Count        Number;
  v_Lngid        ҩƷ�շ���¼.Id%Type; --�շ�ID
  n_Tmp���      ����ҽ����¼.���%Type;
  n_��Һ����     Number; ----�Ƿ������Һ��ҩ��¼��״̬
  n_���Ϻ�       ҩƷ�շ���¼.���ܷ�ҩ��%Type;
  n_�ⷿid       ҩƷ�շ���¼.�ⷿid%Type;
  v_�շ�ids      Varchar2(4000);
  v_Error        Varchar2(255);
  Err_Custom Exception;

  Procedure �����շ���¼_Insert
  (
    ����id_In     Number,
    ����_In       ҩƷ�շ���¼.����%Type,
    ����_In       ҩƷ���.ҩ������%Type,
    ����_In       ҩƷ�շ���¼.����%Type,
    Ч��_In       ҩƷ�շ���¼.Ч��%Type,
    ���Ч��_In   ҩƷ���.���Ч��%Type,
    �շ�id_In     ҩƷ�շ���¼.Id%Type,
    ����id_In     סԺ���ü�¼.����id%Type,
    ��ҳid_In     סԺ���ü�¼.��ҳid%Type,
    �ⷿid_In     ҩƷ�շ���¼.�ⷿid%Type,
    ����_In       ҩƷ�շ���¼.����%Type,
    ����_In       ������Ϣ.����%Type,
    �Է�����id_In ҩƷ�շ���¼.�Է�����id%Type,
    �շ����_In   סԺ���ü�¼.�շ����%Type,
    �������_In   Varchar,
    P����         ҩƷ�շ���¼.����%Type,
    P����         ҩƷ�շ���¼.��д����%Type
  ) Is
    v_����   ҩƷ�շ���¼.����%Type;
    v_Ч��   ҩƷ�շ���¼.Ч��%Type;
    v_����   ҩƷ�շ���¼.����%Type;
    v_���ȼ� ���.���ȼ�%Type;
  Begin
    --ȷ������
    If Nvl(����_In, 0) <> 0 And ����_In = 0 Then
      --ԭ����,�ֲ�����
      v_���� := Null;
      v_���� := ����_In;
      v_Ч�� := Ч��_In;
    Elsif Nvl(����_In, 0) = 0 And ����_In = 1 Then
      --ԭ������,�ַ���
      Select ҩƷ�շ���¼_Id.Nextval Into v_���� From Dual;
      Select To_Char(Sysdate, 'YYYYMMDD') Into v_���� From Dual;
      If ���Ч��_In Is Not Null Then
        v_Ч�� := Trunc(Sysdate + ���Ч��_In * 30);
      Else
        v_Ч�� := Null;
      End If;
    Else
      v_���� := ����_In;
      v_���� := ����_In;
      v_Ч�� := Ч��_In;
    End If;
  
    Select ҩƷ�շ���¼_Id.Nextval Into v_Lngid From Dual;
    Insert Into ҩƷ�շ���¼
      (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ����, ��д����, ʵ������, ���ۼ�, ���۽��, ժҪ, ������, ��������,
       ����id, ����, Ƶ��, �÷�, ��ҩ��λid, ��������, ��׼�ĺ�, ���Ч��)
      Select v_Lngid, 1, ����, No_In, v_�շ����, �ⷿid, �Է�����id, ������id, -1, ҩƷid, Nvl(v_����, 0), ����, v_����, v_Ч��, P����, -1 * P����,
             -1 * P����, ���ۼ�, Round(-1 * P���� * P���� * ���ۼ�, v_Dec), '���ڷ����ջ�', v_��Ա����, �ջ�ʱ��_In, ����id_In, ����, Ƶ��, �÷�, ��ҩ��λid,
             ��������, ��׼�ĺ�, ���Ч��
      From ҩƷ�շ���¼
      Where ID = �շ�id_In;
  
    Zl_δ��ҩƷ��¼_Insert(v_Lngid);
  
    Zl_ҩƷ���_Update(v_Lngid, 0, 1);
  
    --δ��ҩƷ��¼
    Update δ��ҩƷ��¼
    Set ����id = ����id_In, ��ҳid = ��ҳid_In, ���� = ����_In
    Where ���� = ����_In And NO = No_In And �ⷿid + 0 = �ⷿid_In;
  
    If Sql%RowCount = 0 Then
      --ȡ������ȼ�
      Begin
        Select b.���ȼ� Into v_���ȼ� From ������Ϣ A, ��� B Where a.��� = b.����(+) And a.����id = ����id_In;
      Exception
        When Others Then
          Null;
      End;
    
      Insert Into δ��ҩƷ��¼
        (����, NO, ����id, ��ҳid, ����, ���ȼ�, �Է�����id, �ⷿid, ��������, ���շ�, ��ӡ״̬)
      Values
        (����_In, No_In, ����id_In, ��ҳid_In, ����_In, v_���ȼ�, �Է�����id_In, �ⷿid_In, �ջ�ʱ��_In,
         Decode(Nvl(Instr(�������_In, Decode(�շ����_In, '4', '4', '5')), 0), 0, 1, 0), 0);
    End If;
    v_�շ���� := v_�շ���� + 1;
  End;
Begin
  --ȡ����Ա��Ϣ(����ID,��������;��ԱID,��Ա���,��Ա����)
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
  --����Ƿ�����Һ��Һ��¼�����Ƿ��Ѿ�����  
  Select Count(1)
  Into n_Count
  From ��Һ��ҩ��¼ A, ����ҽ����¼ B
  Where a.ҽ��id = b.Id And ҽ��id = ҽ��id_In And a.ִ��ʱ�� > b.ִ����ֹʱ�� And a.�Ƿ����� = 1;

  If n_Count > 0 Then
    v_Error := 'ҽ��"' || v_ҽ������ || '"����ҺҩƷ���Ѿ�����Һ�����������������ܳ����ջء�';
    Raise Err_Custom;
  End If;

  Select a.ҽ������, Nvl(a.���id, a.Id), b.��������
  Into v_ҽ������, n_���id, n_Count
  From ����ҽ����¼ A, ������ҳ B
  Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.Id = ҽ��id_In;
  --�ж��������۲���ִ�е����Ĺ���ά��ʱ��ͬ���޸�
  If n_Count = 1 Then
    Zl_����ҽ����¼_�ջ�_��������(�ջ���_In, ҽ��id_In, �ϴ�ʱ��_In, �ջ�ʱ��_In, No_In, ����Ա���_In, ����Ա����_In);
    Return;
  End If;

  Select Max(����˵��) Into v_����ԭ�� From ����ҽ��״̬ Where ҽ��id = ҽ��id_In And �������� = 8;
  If Nvl(�ջ���_In, 0) > 0 Then
    --�ж��Ƿ�����Һ��ҩҩƷ(��Һ��������ҩƷͳһ����������)
    b_��Һ��ҩ��¼ := False;
    v_��Һ�ջ�ʣ�� := �ջ���_In;
  
    Select Max(a.ִ����ֹʱ��)
    Into d_ִ����ֹʱ��
    From ��Һ��ҩ��¼ E, ����ҽ����¼ A
    Where a.Id = n_���id And e.ҽ��id = a.Id And e.ִ��ʱ�� > a.ִ����ֹʱ��;
  
    If d_ִ����ֹʱ�� Is Not Null Then
      d_�ջ�ʱ��       := �ջ�ʱ��_In;
      v_��Һҩ�������� := zl_GetSysParameter('��Һ��Һ����ҩ��������������', 1345);
      b_��Һ��ҩ��¼   := True;
    
      If n_���id = ҽ��id_In Then
        --��ҩ;���У�����״̬�������ı�����
        n_��Һ���� := 1;
      Else
        n_��Һ���� := 0;
        n_Tmp���  := ҽ��id_In;
      End If;
    
      For X In (Select e.Id As ��ҩid, e.����״̬, e.�Ƿ���
                From ��Һ��ҩ��¼ E
                Where e.ҽ��id = n_���id And e.ִ��ʱ�� > d_ִ����ֹʱ�� And Nvl(e.����״̬, 0) In (1, 2, 3, 4, 5, 6, 7, 8)) Loop
        If Not (x.����״̬ In (4, 5, 6, 7, 8) And Nvl(x.�Ƿ���, 0) = 0 And Nvl(v_��Һҩ��������, '0') = '0') Then
          If n_��Һ���� = 0 Then
            --����ҩƷ����ϸ��������
            For r_Compound In c_Compound(n_���id, d_ִ����ֹʱ��, x.��ҩid, n_Tmp���) Loop
            
              v_��Һ�ջ�ʣ�� := v_��Һ�ջ�ʣ�� - r_Compound.����;
              If x.����״̬ = 1 Then
                n_������� := 0;
              Else
                n_������� := 1;
              End If;
              Zl_���˷�������_Insert(r_Compound.����id, r_Compound.�շ�ϸĿid, r_Compound.���˲���id, r_Compound.����, v_��Ա����, d_�ջ�ʱ��,
                               n_�������, Null, r_Compound.��ҩid, v_����ԭ��, 0);
              If x.����״̬ = 1 Then
                --δ��ҩ�ģ��Զ���ˡ�
                Zl_���˷�������_Audit(r_Compound.����id, d_�ջ�ʱ��, v_��Ա����, d_�ջ�ʱ��, 1, 1, n_�������);
                Zl_סԺ���ʼ�¼_Delete(r_Compound.No, r_Compound.��� || ':' || r_Compound.���� || ':' || r_Compound.��ҩid, v_��Ա���,
                                 v_��Ա����, 2, Null, Null, d_�ջ�ʱ��);
              End If;
            End Loop;
          End If;
        
          --����״̬
          If n_��Һ���� = 1 Then
            Select Count(1)
            Into n_Count
            From ��Һ��ҩ״̬
            Where ��ҩid = x.��ҩid And �������� = 9 And ����ʱ�� = d_�ջ�ʱ��;
            If n_Count = 0 Then
              Insert Into ��Һ��ҩ״̬
                (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
              Values
                (x.��ҩid, 9, v_��Ա����, d_�ջ�ʱ��, v_����ԭ��);
            End If;
            Update ��Һ��ҩ��¼ Set ������Ա = v_��Ա����, ����ʱ�� = d_�ջ�ʱ��, ����״̬ = 9 Where ID = x.��ҩid;
          
            If x.����״̬ = 1 Then
              Insert Into ��Һ��ҩ״̬
                (��ҩid, ��������, ������Ա, ����ʱ��)
              Values
                (x.��ҩid, 10, v_��Ա����, d_�ջ�ʱ��);
              Update ��Һ��ҩ��¼ Set ������Ա = v_��Ա����, ����ʱ�� = d_�ջ�ʱ��, ����״̬ = 10 Where ID = x.��ҩid;
            End If;
          End If;
        
          --���ڲ�ͬ���Σ�ִ��ʱ�䣩����ʱ������ʱ��ͷ���ID��ΨһԼ��������ͬʱ���ʶ������ʱ�����μ�һ��
          d_�ջ�ʱ�� := d_�ջ�ʱ�� + 1 / 24 / 60 / 60;
        End If;
      End Loop;
    End If;
  
    --a.���������ջ�ģʽ
    --��Һ��ҩ��¼������������
    If b_��Һ��ҩ��¼ = False Or v_��Һ�ջ�ʣ�� > 0 Then
      If No_In Is Null Then
        v_���ʲ��� := zl_GetSysParameter(23);
        --�����ջ���������ԭʼ���ý��з�̯����
        For r_Detail In c_Detail Loop
          --ȷ�����շ�ϸĿID���ջ�������
          If Nvl(v_�շ�ϸĿid, 0) <> r_Detail.�շ�ϸĿid And (r_Detail.������� Not In ('5', '6', '7') Or Nvl(v_�շ�ϸĿid, 0) = 0) Then
            --����δ��̯���
            If v_�շ�ϸĿid Is Not Null And v_�ջ����� > 0 Then
              v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ����������ܴ����ֹ����ʻ��ѽ��ʵķ��ã�����Ӧ�Ļ��۵��ѱ�ɾ����';
              Raise Err_Custom;
            End If;
            --ҩƷ�ջ�������������͹��Ϊ׼����ģ��Դ˼�����ջ��ۼ�����
            Begin
              Select ����ϵ��, סԺ��װ Into v_����ϵ��, v_סԺ��װ From ҩƷ��� Where ҩƷid = r_Detail.�շ�ϸĿid;
            Exception
              When Others Then
                Null;
            End;
            --�����ջ�����������շѷ�ʽ����0������ȡ����ʹ������ķ������м���
            If r_Detail.�շѷ�ʽ = 0 Then
              If r_Detail.������� = '7' Then
                --��ҩ�䷽ҩƷ������*����
                v_�ջ����� := Round(v_��Һ�ջ�ʣ�� * r_Detail.�������� / Nvl(v_����ϵ��, 1), 5);
              Else
                If r_Detail.������� Not In ('5', '6') Then
                  Select Nvl(Max(����), 1)
                  Into v_��������
                  From ����ҽ���Ƽ�
                  Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Detail.�շ�ϸĿid;
                Else
                  v_�������� := 1;
                End If;
                v_�ջ����� := Round(v_��Һ�ջ�ʣ�� * Nvl(v_סԺ��װ, 1), 5) * v_��������;
              End If;
            Else
              Select Nvl(Sum(����), 0)
              Into v_�ջ�����
              From ҽ��ִ�мƼ�
              Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Detail.�շ�ϸĿid And
                    Ҫ��ʱ�� > Nvl(�ϴ�ʱ��_In, To_Date('1900-01-01', 'yyyy-MM-dd'));
            
              v_�ջ����� := Round(v_�ջ�����, 5);
            
            End If;
            v_ҽ������ := r_Detail.ҽ������;
          End If;
        
          --���շ�ϸĿ��ÿ��������ϸ��̯�ջ�
          If v_�ջ����� > 0 Then
            --����Ӧ�����Ƿ��ѽ��ʣ�����ֹʱ
            v_���ʽ�� := 0;
            If v_���ʲ��� = '2' And r_Detail.��¼״̬ <> 0 Then
              Select Sum(���ʽ��)
              Into v_���ʽ��
              From סԺ���ü�¼
              Where NO = r_Detail.No And ��¼���� In (2, 12) And Nvl(�۸񸸺�, ���) = r_Detail.���;
            End If;
          
            If Nvl(v_���ʽ��, 0) = 0 Then
              If r_Detail.�շ���� In ('5', '6', '7') Or r_Detail.�շ���� = '4' And r_Detail.�������� = 1 Then
                --ҩƷ�͸������õ�����
                If r_Detail.ִ�б�־ = 0 Then
                  v_ʣ������ := r_Detail.δִ����;
                Elsif r_Detail.ִ�б�־ = 1 Then
                  v_ʣ������ := r_Detail.��ִ����;
                End If;
              Else
                --��ͨ����
                v_ʣ������ := r_Detail.ʣ������;
              End If;
              If v_�ջ����� > v_ʣ������ Then
                v_��ǰ���� := v_ʣ������;
              Else
                v_��ǰ���� := v_�ջ�����;
              End If;
              v_�ջ����� := v_�ջ����� - v_��ǰ����;
              --ϵͳ��������ִ�к��Ƿ���˻��۵������ԣ���ִ�е���Ȼ�����ǻ��۵�
              If r_Detail.ִ�б�־ = 0 And r_Detail.��¼״̬ = 0 Then
                v_Delno := v_Delno || '|' || r_Detail.No || ',' || r_Detail.��� || ':' || v_��ǰ����;
              Else
                If Not (r_Detail.�շ���� = '7' And r_Detail.ִ�б�־ <> 0) Then
                  Zl_���˷�������_Insert(r_Detail.����id, r_Detail.�շ�ϸĿid, r_Detail.���˲���id, v_��ǰ����, v_��Ա����, �ջ�ʱ��_In,
                                   r_Detail.ִ�б�־, Null, Null, v_����ԭ��);
                End If;
              End If;
              v_����ids := v_����ids || ',' || r_Detail.����id;
            End If;
          End If;
          v_�շ�ϸĿid := r_Detail.�շ�ϸĿid;
        End Loop;
      
        --����δ��̯���
        If v_�ջ����� > 0 Then
          v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ����������ܴ����ֹ����ʻ��ѽ��ʵķ��ã�����Ӧ�Ļ��۵��ѱ�ɾ����';
          Raise Err_Custom;
        End If;
        --���Ƶ����������Զ����
        If zl_GetSysParameter('�����ջط��ñ����Զ����', 1254) = '1' And v_����ids Is Not Null Then
          For r_Applay In c_Applay(Substr(v_����ids, 2)) Loop
            Zl_���˷�������_Audit(r_Applay.����id, r_Applay.����ʱ��, v_��Ա����, �ջ�ʱ��_In, 1, 1, r_Applay.�������);
            v_Delno := v_Delno || '|' || r_Applay.No || ',' || r_Applay.��� || ':' || r_Applay.����;
          End Loop;
        End If;
      Else
        ---b.�����ջ�ģʽ-------------------------------------------------------------------------------------------------------
        --���ȫ�ǻ��۵����Ͳ��ò���������������
        If No_In = '�������۵�' Then
          --δ��˵Ļ��۵����Ƚ����޸Ļ�ɾ�������ܶ�η���Ϊ��ͬ��NO,Ϊ�˼���ÿ�ε��ջ�������Ҫ���շ�ϸĿID����
          For r_Price In (Select c.�������, b.No, b.���, b.�շ�ϸĿid, Nvl(b.����, 1) * b.���� As ʣ������, c.��������, d.����ϵ��, d.סԺ��װ,
                                 c.ҽ������, Nvl(e.�շѷ�ʽ, 0) As �շѷ�ʽ
                          From ����ҽ������ A, סԺ���ü�¼ B, ����ҽ����¼ C, ҩƷ��� D, ����ҽ���Ƽ� E
                          Where a.ҽ��id = ҽ��id_In And a.No = b.No And a.��¼���� = b.��¼���� And a.ҽ��id = b.ҽ����� And
                                b.�۸񸸺� Is Null And b.�շ�ϸĿid = d.ҩƷid(+) And b.��¼״̬ = 0 And c.Id = a.ҽ��id And
                                b.ҽ����� = e.ҽ��id(+) And b.�շ�ϸĿid = e.�շ�ϸĿid(+)
                          Order By �շ�ϸĿid, NO Desc) Loop
            If Nvl(v_�շ�ϸĿid, 0) <> r_Price.�շ�ϸĿid Then
              --����δ��̯���
              If v_�շ�ϸĿid Is Not Null And v_�ջ����� > 0 Then
                v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ�������������ػ��۵��ѱ�ɾ������ˡ�';
                Raise Err_Custom;
              End If;
              --�����ջ�����������շѷ�ʽ����0������ȡ����ʹ������ķ������м���
              If r_Price.�շѷ�ʽ = 0 Then
                If r_Price.������� = '7' Then
                  --��ҩ�䷽ҩƷ������*����
                  v_�ջ����� := Round(v_��Һ�ջ�ʣ�� * r_Price.�������� / Nvl(r_Price.����ϵ��, 1), 5);
                Else
                  If r_Price.������� Not In ('5', '6') Then
                    Select Nvl(Max(����), 1)
                    Into v_��������
                    From ����ҽ���Ƽ�
                    Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Price.�շ�ϸĿid;
                  Else
                    v_�������� := 1;
                  End If;
                  v_�ջ����� := Round(v_��Һ�ջ�ʣ�� * Nvl(r_Price.סԺ��װ, 1), 5) * v_��������;
                End If;
              Else
                Select Nvl(Sum(����), 0)
                Into v_�ջ�����
                From ҽ��ִ�мƼ�
                Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Price.�շ�ϸĿid And
                      Ҫ��ʱ�� > Nvl(�ϴ�ʱ��_In, To_Date('1900-01-01', 'yyyy-MM-dd'));
              
                v_�ջ����� := Round(v_�ջ�����, 5);
              End If;
              v_ҽ������ := r_Price.ҽ������;
            End If;
            If v_�ջ����� > 0 Then
              If v_�ջ����� > r_Price.ʣ������ Then
                v_��ǰ���� := r_Price.ʣ������;
              Else
                v_��ǰ���� := v_�ջ�����;
              End If;
              v_�ջ����� := v_�ջ����� - v_��ǰ����;
              v_Delno    := v_Delno || '|' || r_Price.No || ',' || r_Price.��� || ':' || v_��ǰ����;
            End If;
            v_�շ�ϸĿid := r_Price.�շ�ϸĿid;
          End Loop;
          --����δ��̯���
          If v_�ջ����� > 0 Then
            v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ�������������ػ��۵��ѱ�ɾ������ˡ�';
            Raise Err_Custom;
          End If;
        Else
          --�������������ܴ��ڻ��۵�����ʵ���ϵ����
          --���С��λ��
          Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;
          --���ɻ��۵�ϵͳ����
          Select zl_GetSysParameter(80) Into v_������� From Dual;
          v_��ʼ��� := Null;
          v_������� := Null;
        
          Select a.�������, a.��������, b.��������
          Into v_�������, v_��������, v_��������
          From ����ҽ����¼ A, �������� B
          Where ID = ҽ��id_In And a.�շ�ϸĿid = b.����id(+);
        
          If v_������� In ('5', '6', '7') Or (v_������� = '4' And Nvl(v_��������, 0) = 1) Then
            --ҩƷ������
            -----------------------------------------------------------------------------------------------------
            v_�ջ����� := Null;
            Select Nvl(Max(���), 0) + 1
            Into v_�շ����
            From ҩƷ�շ���¼
            Where ���� In (9, 10, 25, 26) And ��¼״̬ = 1 And NO = No_In;
            Select Nvl(Max(���), 0) + 1
            Into v_�������
            From סԺ���ü�¼
            Where ��¼���� = 2 And ��¼״̬ In (0, 1) And NO = No_In;
          
            --һ��ҽ����ҩƷֻ��һ�У������ѭ����Ϊ�˴����η��͵����������ҩƷ�ڽ����ѽ��ø����ջ�
            For r_Drug In c_Drug Loop
              --��ʼ��Ҫ�ջص�������(��������)
              v_First := 0;
              If v_�ջ����� Is Null Then
                If v_������� = '7' Then
                  v_�ջ����� := Round(v_��Һ�ջ�ʣ�� * v_�������� / r_Drug.����ϵ��, 5);
                Else
                  If v_������� Not In ('5', '6') Then
                    Select Nvl(Max(����), 1)
                    Into v_��������
                    From ����ҽ���Ƽ�
                    Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Drug.�շ�ϸĿid;
                  Else
                    v_�������� := 1;
                  End If;
                  v_�ջ����� := Round(v_��Һ�ջ�ʣ�� * r_Drug.סԺ��װ, 5) * v_��������;
                End If;
                v_First := 1;
              End If;
            
              --�����һ���������㹻���򰴸����������������ô���
              If v_�ջ����� > r_Drug.���� Then
                v_��ǰ���� := 1;
                v_��ǰ���� := r_Drug.����;
                v_�ջ����� := v_�ջ����� - r_Drug.����;
              Else
                If v_First = 1 And v_������� = '7' Then
                  v_��ǰ���� := v_��Һ�ջ�ʣ��;
                  v_��ǰ���� := Round(v_�������� / r_Drug.����ϵ��, 5);
                Else
                  v_��ǰ���� := 1;
                  v_��ǰ���� := v_�ջ�����;
                End If;
                v_�ջ����� := 0;
              End If;
            
              If r_Drug.��¼״̬ = 0 Then
                v_Delno := v_Delno || '|' || r_Drug.No || ',' || r_Drug.��� || ':' || v_��ǰ���� * v_��ǰ����;
              Else
                If Not (v_������� = '7' And r_Drug.ִ�б�־ <> 0) Then
                
                  Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
                  �����շ���¼_Insert(v_����id, r_Drug.����, r_Drug.����, r_Drug.����, r_Drug.Ч��, r_Drug.���Ч��, r_Drug.�շ�id,
                                r_Drug.����id, r_Drug.��ҳid, r_Drug.�ⷿid, r_Drug.����, r_Drug.����, r_Drug.�Է�����id, v_�������,
                                v_�������, v_��ǰ����, v_��ǰ����);
                
                  --סԺ���ü�¼
                  -------------------------------------------------------------------------------------
                  --��¼��ŷ�Χ�Դ�����ܱ�
                  If v_��ʼ��� Is Null Then
                    v_��ʼ��� := v_�������;
                  End If;
                  v_������� := v_�������;
                
                  Insert Into סԺ���ü�¼
                    (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ����, ���˲���id, ���˿���id,
                     �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��,
                     ͳ����, ���ʷ���, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ҽ�����, ������, ����Ա���, ����Ա����)
                    Select v_����id, 2, No_In, Decode(Nvl(Instr(v_�������, Decode(v_�������, '4', '4', '5')), 0), 0, 1, 0),
                           v_�������, Null, Null, �ಡ�˵�, 2, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ����, ���˲���id, ���˿���id, �ѱ�, �շ����,
                           �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, v_��ǰ����, -1 * v_��ǰ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����,
                           Round(-1 * v_��ǰ���� * v_��ǰ���� * ��׼����, v_Dec), Round(-1 * v_��ǰ���� * v_��ǰ���� * ��׼����, v_Dec), Null, 1,
                           ��������id, ������, �ջ�ʱ��_In, �ջ�ʱ��_In, ִ�в���id, 0, ҽ�����, v_��Ա����,
                           Decode(Nvl(Instr(v_�������, Decode(v_�������, '4', '4', '5')), 0), 0, v_��Ա���, Null),
                           Decode(Nvl(Instr(v_�������, Decode(v_�������, '4', '4', '5')), 0), 0, v_��Ա����, Null)
                    From סԺ���ü�¼
                    Where ID = r_Drug.����id;
                
                  Select Zl_Actualmoney(�ѱ�, �շ�ϸĿid, ������Ŀid, Ӧ�ս��, ����, ִ�в���id)
                  Into v_Temp
                  From סԺ���ü�¼
                  Where ID = v_����id;
                  v_ʵ�ս�� := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
                  Update סԺ���ü�¼ A Set ʵ�ս�� = v_ʵ�ս�� Where ID = v_����id;
                
                  v_������� := v_������� + 1;
                End If;
                If v_�ջ����� <= 0 Then
                  Exit;
                End If;
              End If;
            End Loop;
          
            If v_�ջ����� <> 0 Then
              --û���ջ���������,�շ���¼����������(���¼��ȫ������Ϊ��)
              Null;
            End If;
          Else
            --������ҩҽ��(������ҩ;�������󶨵����ĵ�)
            -----------------------------------------------------------------------------------------------------
            Select Nvl(Max(���), 0) + 1
            Into v_�շ����
            From ҩƷ�շ���¼
            Where ���� In (9, 10, 25, 26) And ��¼״̬ = 1 And NO = No_In;
            --ȡ�������
            Select Nvl(Max(���), 0) + 1
            Into v_�������
            From סԺ���ü�¼
            Where ��¼���� = 2 And ��¼״̬ In (0, 1) And NO = No_In;
          
            v_�ջ�ʣ�� := v_��Һ�ջ�ʣ��;
          
            For r_Othersend In (Select ���ͺ�, ��������
                                From ����ҽ������
                                Where ҽ��id = ҽ��id_In And ĩ��ʱ�� > Nvl(�ϴ�ʱ��_In, To_Date('1900-01-01', 'yyyy-MM-dd'))
                                Order By ���ͺ� Desc) Loop
              If r_Othersend.�������� < v_�ջ�ʣ�� Then
                --һ���ջض�η��ͣ�����ÿ�η��ͷ��������䶯���Ƽۣ�
                v_�ջ�ʣ�� := v_�ջ�ʣ�� - r_Othersend.��������;
                v_�ջ���   := r_Othersend.��������;
              Else
                --һ�η������ջ�ʣ�ࣻ
                v_�ջ���   := v_�ջ�ʣ��;
                v_�ջ�ʣ�� := 0;
              End If;
              v_�շ����� := '';
              For r_Other In c_Other(r_Othersend.���ͺ�) Loop
                If Nvl(v_�շ�����, '0') <> r_Other.�շ�ϸĿid || ',' || r_Other.��� Then
                  --�������һ�η��͵ķ��ü�¼������Ҫ�ջص�����ȫ���ջ�
                  --�����ջ�����������շѷ�ʽ����0������ȡ����ʹ������ķ������м���
                  If r_Other.�շѷ�ʽ = 0 Then
                    v_�ջ����� := v_�ջ��� * Nvl(r_Other.��������, 1);
                  Else
                    Select Nvl(Sum(����), 0)
                    Into v_�ջ�����
                    From ҽ��ִ�мƼ�
                    Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Other.�շ�ϸĿid And
                          Ҫ��ʱ�� > Nvl(�ϴ�ʱ��_In, To_Date('1900-01-01', 'yyyy-MM-dd'));
                  End If;
                End If;
              
                If v_�ջ����� > 0 Then
                  If r_Other.��¼״̬ = 0 Then
                    If v_�ջ����� > r_Other.ʣ������ Then
                      v_��ǰ���� := r_Other.ʣ������;
                    Else
                      v_��ǰ���� := v_�ջ�����;
                    End If;
                  Else
                    v_��ǰ���� := v_�ջ�����;
                  End If;
                  v_�ջ����� := v_�ջ����� - v_��ǰ����;
                  v_��ǰ���� := 1;
                
                  If r_Other.��¼״̬ = 0 Then
                    v_Delno := v_Delno || '|' || r_Other.No || ',' || r_Other.��� || ':' || v_��ǰ����;
                  Else
                    --��¼��ŷ�Χ�Դ�����ܱ�
                    If v_��ʼ��� Is Null Then
                      v_��ʼ��� := v_�������;
                    End If;
                    v_������� := v_�������;
                  
                    --סԺ���ü�¼:��������ջ����������ϴη�����,����ȷ�������ε����Ŀ����ж����շ���¼
                    Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
                    If r_Other.�շ���� In ('4', '5', '6', '7') Then
                      n_���� := v_��ǰ����;
                      For r_Otherdrug In (Select a.����id, a.��ҳid, d.����, Nvl(Nvl(x.����ϵ��, y.����ϵ��), 1) As ����ϵ��,
                                                 Nvl(x.סԺ��װ, 1) As סԺ��װ, Nvl(x.���Ч��, y.���Ч��) As ���Ч��,
                                                 Nvl(b.����, 1) * b.ʵ������ As ����, b.Id As �շ�id, b.����, b.ҩƷid, b.�Է�����id,
                                                 b.�ⷿid, b.����id, Nvl(Nvl(x.ҩ������, y.���÷���), 0) As ����, b.����, b.����, b.Ч��,
                                                 a.��¼״̬, a.No, a.���, a.�շ�ϸĿid
                                          From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ������Ϣ D, ҩƷ��� X, �������� Y
                                          Where a.Id = r_Other.����id And a.��¼״̬ In (0, 1, 3) And a.No = b.No And
                                                a.Id = b.����id + 0 And b.���� In (9, 10, 25, 26) And
                                                (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0) And a.����id = d.����id And
                                                b.ҩƷid = x.ҩƷid(+) And b.ҩƷid = y.����id(+)
                                          Order By a.��¼״̬, b.No Desc, b.Id Desc) Loop
                        If n_���� > 0 Then
                          n_Count := r_Otherdrug.����;
                          If n_���� < n_Count Then
                            n_Count := n_����;
                          End If;
                          �����շ���¼_Insert(v_����id, r_Otherdrug.����, r_Otherdrug.����, r_Otherdrug.����, r_Otherdrug.Ч��,
                                        r_Otherdrug.���Ч��, r_Otherdrug.�շ�id, r_Otherdrug.����id, r_Otherdrug.��ҳid,
                                        r_Otherdrug.�ⷿid, r_Otherdrug.����, r_Otherdrug.����, r_Otherdrug.�Է�����id,
                                        r_Other.�շ����, v_�������, 1, n_Count);
                          n_���� := n_���� - r_Otherdrug.����;
                        End If;
                      End Loop;
                    End If;
                    --ҽ����ִ�У��ջصķ���Ҳ��Ϊ��ִ�У�������ҩƷ�͸������õ����ģ���Ϊʵ�ʷ��ű�ʾִ��
                    Insert Into סԺ���ü�¼
                      (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ����, ���˲���id, ���˿���id,
                       �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��,
                       ͳ����, ���ʷ���, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ��ʱ��, ִ����, ҽ�����, ������, ����Ա���, ����Ա����)
                      Select v_����id, 2, No_In, Decode(Nvl(Instr(v_�������, r_Other.�շ����), 0), 0, 1, 0), v_�������, Null,
                             Decode(a.�۸񸸺�, Null, Null, v_������� + a.�۸񸸺� - a.���), a.�ಡ�˵�, 2, a.����id, a.��ҳid, a.��ʶ��, a.����,
                             a.�Ա�, a.����, a.����, a.���˲���id, a.���˿���id, a.�ѱ�, a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.������Ŀ��, a.���մ���id, 1,
                             -1 * v_��ǰ����, a.�Ӱ��־, a.���ӱ�־, a.Ӥ����, a.������Ŀid, a.�վݷ�Ŀ, a.��׼����,
                             Round(-1 * v_��ǰ���� * a.��׼����, v_Dec), Round(-1 * v_��ǰ���� * a.��׼����, v_Dec), Null, 1, a.��������id,
                             a.������, �ջ�ʱ��_In, �ջ�ʱ��_In, a.ִ�в���id,
                             Decode(r_Other.ִ��״̬, 1,
                                     Decode(a.�շ����, '4', Decode(b.��������, 1, 0, 1),
                                             Decode(Instr(',5,6,7,', a.�շ����), 0, 1, 0)), 0),
                             Decode(r_Other.ִ��״̬, 1,
                                     Decode(a.�շ����, '4', Decode(b.��������, 1, Null, �ջ�ʱ��_In),
                                             Decode(Instr(',5,6,7,', a.�շ����), 0, �ջ�ʱ��_In, Null)), Null),
                             Decode(r_Other.ִ��״̬, 1,
                                     Decode(a.�շ����, '4', Decode(b.��������, 1, Null, v_��Ա����),
                                             Decode(Instr(',5,6,7,', a.�շ����), 0, v_��Ա����, Null)), Null), a.ҽ�����, v_��Ա����,
                             Decode(Nvl(Instr(v_�������, r_Other.�շ����), 0), 0, v_��Ա���, Null),
                             Decode(Nvl(Instr(v_�������, r_Other.�շ����), 0), 0, v_��Ա����, Null)
                      From סԺ���ü�¼ A, �������� B
                      Where a.Id = r_Other.����id And a.�շ�ϸĿid = b.����id(+);
                  
                    Select Zl_Actualmoney(�ѱ�, �շ�ϸĿid, ������Ŀid, Ӧ�ս��, ����, ִ�в���id)
                    Into v_Temp
                    From סԺ���ü�¼
                    Where ID = v_����id;
                    v_ʵ�ս�� := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
                    Update סԺ���ü�¼ A Set ʵ�ս�� = v_ʵ�ս�� Where ID = v_����id;
                  
                    v_������� := v_������� + 1;
                    v_ҽ��ִ�� := r_Other.ִ��״̬; --����շ���Ŀ��ִ��״̬��һ����
                  End If;
                
                  v_�շ����� := r_Other.�շ�ϸĿid || ',' || r_Other.���;
                End If;
              End Loop;
              If v_�ջ�ʣ�� = 0 Then
                Exit;
              End If;
            End Loop;
          
            --���ҽ����ִ�У���ϵͳ����ִ�к��Զ���˷��ã�������ִ��ҽ����Ӧ��ҩƷ�����ķ��á�
            -----------------------------------------------------------------------------------------------------
            If Nvl(v_ҽ��ִ��, 0) = 1 And v_��ʼ��� Is Not Null And v_������� Is Not Null Then
              For r_Verify In c_Verify(v_��ʼ���, v_�������) Loop
                Zl_סԺ���ʼ�¼_Verify(r_Verify.No, v_��Ա���, v_��Ա����, r_Verify.���, Null, �ջ�ʱ��_In);
              End Loop;
            End If;
          End If;
        
          --������û��ܱ�
          -----------------------------------------------------------------------------------------------------
          If v_��ʼ��� Is Not Null And v_������� Is Not Null Then
            --���ͳһ���������ػ��ܱ�
            For r_Money In c_Money(v_��ʼ���, v_�������) Loop
              --�������
              Update �������
              Set ������� = Nvl(�������, 0) + r_Money.ʵ�ս��
              Where ����id = r_Money.����id And ���� = 1 And ���� = 2;
            
              If Sql%RowCount = 0 Then
                Insert Into �������
                  (����id, ����, ����, �������, Ԥ�����)
                Values
                  (r_Money.����id, 1, 2, r_Money.ʵ�ս��, 0);
              End If;
            
              --����δ�����
              Update ����δ�����
              Set ��� = Nvl(���, 0) + r_Money.ʵ�ս��
              Where ����id = r_Money.����id And ��ҳid = r_Money.��ҳid And Nvl(���˲���id, 0) = Nvl(r_Money.���˲���id, 0) And
                    Nvl(���˿���id, 0) = Nvl(r_Money.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Money.��������id, 0) And
                    Nvl(ִ�в���id, 0) = Nvl(r_Money.ִ�в���id, 0) And ������Ŀid + 0 = r_Money.������Ŀid And ��Դ;�� + 0 = 2;
            
              If Sql%RowCount = 0 Then
                Insert Into ����δ�����
                  (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
                Values
                  (r_Money.����id, r_Money.��ҳid, r_Money.���˲���id, r_Money.���˿���id, r_Money.��������id, r_Money.ִ�в���id,
                   r_Money.������Ŀid, 2, r_Money.ʵ�ս��);
              End If;
            End Loop;
          End If;
        End If;
      End If;
    End If;
  End If;

  --����Zl_סԺ���ʼ�¼_Delete����֧��ÿ��ɾ��һ�е�ѭ����������������������һ������Ҫɾ�������һ���Դ���
  If Not v_Delno Is Null Then
    v_Temp := '';
    v_No   := '';
    For r_Price In (Select /*+ rule*/
                     C1 As NO, C2 As �������
                    From Table(f_Str2list2(Substr(v_Delno, 2), '|', ','))
                    Order By NO) Loop
      If v_No Is Not Null And v_No <> r_Price.No Then
        Zl_סԺ���ʼ�¼_Delete(v_No, v_Temp, v_��Ա���, v_��Ա����, 2);
        v_No := '';
      End If;
      If v_No Is Null Then
        v_No   := r_Price.No;
        v_Temp := r_Price.�������;
      Else
        v_Temp := v_Temp || ',' || r_Price.�������;
      End If;
    End Loop;
    If Not v_No Is Null Then
      Zl_סԺ���ʼ�¼_Delete(v_No, v_Temp, v_��Ա���, v_��Ա����, 2);
    End If;
  End If;

  --����ҽ�����ϴ�ִ��ʱ��:��ҩ;���ȿ�����Ϊδ���Ͷ�û�����ջع��̡�
  -----------------------------------------------------------------------------------------------------
  Select Nvl(���id, ID) Into v_��id From ����ҽ����¼ Where ID = ҽ��id_In;
  Update ����ҽ����¼ Set �ϴ�ִ��ʱ�� = �ϴ�ʱ��_In Where ID = v_��id Or ���id = v_��id;

  --ɾ��ҽ��ִ��ʱ��
  If �ϴ�ʱ��_In Is Null Then
    --ȫ���ջ�
    Delete From ҽ��ִ��ʱ�� Where ҽ��id = v_��id;
    Delete From ҽ��ִ�мƼ� Where ҽ��id = ҽ��id_In;
  Else
    --�����ջض�η��͵�����
    Delete From ҽ��ִ��ʱ�� Where ҽ��id = v_��id And Ҫ��ʱ�� > �ϴ�ʱ��_In;
    Delete From ҽ��ִ�мƼ� Where ҽ��id = ҽ��id_In And Ҫ��ʱ�� > �ϴ�ʱ��_In;
  End If;
  --������Һ��Һ��¼���������⣬ÿ��ҽ�������е��ã��ڹ�������ֻ��������Һ��Һ��ҽ��
  Zl_��Һ��ҩ��¼_���ε���(ҽ��id_In);

  If zl_GetSysParameter(63) = '1' And Nvl(No_In, '�������۵�') <> '�������۵�' Then
    --����������������Զ�����
    For r_Stuff In c_Stuff Loop
      If n_���Ϻ� Is Null Then
        n_���Ϻ� := Nextno(20);
      End If;
      If r_Stuff.�ⷿid <> Nvl(n_�ⷿid, 0) Then
        If Nvl(n_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
          v_�շ�ids := Substr(v_�շ�ids, 2);
          Zl_ҩƷ�շ���¼_��������(v_�շ�ids, n_�ⷿid, v_��Ա����, Sysdate, 1, v_��Ա����, n_���Ϻ�, v_��Ա����);
        End If;
      
        n_�ⷿid  := r_Stuff.�ⷿid;
        v_�շ�ids := Null;
      End If;
    
      v_�շ�ids := v_�շ�ids || '|' || r_Stuff.Id || ',' || r_Stuff.����;
    End Loop;
    If Nvl(n_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
      v_�շ�ids := Substr(v_�շ�ids, 2);
      Zl_ҩƷ�շ���¼_��������(v_�շ�ids, n_�ⷿid, v_��Ա����, Sysdate, 1, v_��Ա����, n_���Ϻ�, v_��Ա����);
    End If;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_�ջ�;
/

--139097:������,2019-04-28,�������۲��˴���
CREATE OR REPLACE Procedure Zl_����ҽ����¼_�ջ�_��������
(
  --���ܣ��������۲��˳����ջ��ڲ��߼��� Zl_����ҽ����¼_�ջ� ��ͬ,���˵�����������
  --˵��:Zl_����ҽ����¼_�ջ�/Zl_����ҽ����¼_�ջ�_�������� ��ѯ�ı�һ��,�������� ͳһ����ѯ  ������ü�¼
  
  �ջ���_In     In ����ҽ������.��������%Type,
  ҽ��id_In     In ����ҽ����¼.Id%Type,
  �ϴ�ʱ��_In   In ����ҽ����¼.�ϴ�ִ��ʱ��%Type,
  �ջ�ʱ��_In   In ����ҽ����¼.�ϴ�ִ��ʱ��%Type,
  No_In         In ������ü�¼.No%Type := Null,
  ����Ա���_In In ��Ա��.���%Type := Null,
  ����Ա����_In In ��Ա��.����%Type := Null
) Is
  --�ջ�ҽ����Ӧ�ķ��ͷ�����ϸ��ʣ������,��������ķ������ջ�
  --ʣ������û���ſ���������������ݣ��ڲ���������ʱ����ԭ��������
  --��ҩƷ�����ģ���һ�����������ܴ���δִ�к���ִ�в��֣���ֱ���д�����¼������δִ������
  --ִ�б�־=0-δִ��,1-��ִ�У�ҩƷ���в���ִ�У����շ���¼�е���ϸ������Ϊ׼����ҩƷ��ֻ���ȴ���δִ�е�
  Cursor c_Detail Is
    Select a.����id, a.No, a.���, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.��������, a.�������, a.ҽ������, a.��������, a.ʣ������, a.��ִ����, a.δִ����,
           a.ִ�б�־, a.��¼״̬, a.�Ǽ�ʱ��, a.�շѷ�ʽ
    From (With ҽ�����ü�¼ As (Select Max(Decode(b.��¼״̬, 2, 0, b.Id)) As ����id, b.No, Nvl(b.�۸񸸺�, b.���) As ���, b.�շ�ϸĿid,
                                 b.���˲���id, Sum(Nvl(b.����, 1) * b.����) As ʣ������, b.�շ����, Max(Nvl(b.ִ��״̬, 0)) As ִ��״̬, d.��������,
                                 c.�������, c.ҽ������, c.��������, Max(b.��¼״̬) As ��¼״̬, Max(b.�Ǽ�ʱ��) As �Ǽ�ʱ��, Nvl(e.�շѷ�ʽ, 0) As �շѷ�ʽ
                          From ����ҽ������ A, ������ü�¼ B, ����ҽ����¼ C, �������� D, ����ҽ���Ƽ� E
                          Where a.ҽ��id = ҽ��id_In And a.No = b.No And a.��¼���� = b.��¼���� And a.ҽ��id = b.ҽ����� And
                                b.�۸񸸺� Is Null And b.�շ�ϸĿid = d.����id(+) And c.Id = ҽ��id_In And e.ҽ��id(+) = b.ҽ����� And
                                e.�շ�ϸĿid(+) = b.�շ�ϸĿid And Not Exists
                           (Select 1 From ��Һ��ҩ��¼ F Where f.ҽ��id = c.���id And a.���ͺ� = f.���ͺ�)
                          Group By b.No, b.��¼����, Nvl(b.�۸񸸺�, b.���), b.�շ�ϸĿid, b.���˲���id, b.�շ����, d.��������, c.�������, c.ҽ������,
                                   c.��������, e.�շѷ�ʽ
                          Having Sum(Nvl(b.����, 1) * b.����) > 0)
           Select ����id, NO, ���, �շ�ϸĿid, ���˲���id, �շ����, ��������, �������, ҽ������, ��������, ʣ������, Null As ��ִ����, Null As δִ����,
                  ִ��״̬ As ִ�б�־, ��¼״̬, �Ǽ�ʱ��, �շѷ�ʽ
           From ҽ�����ü�¼
           Where �շ���� Not In ('5', '6', '7') And Not (�շ���� = '4' And Nvl(��������, 0) = 1)
           Union All
           Select a.����id, a.No, a.���, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.��������, a.�������, a.ҽ������, a.��������, a.ʣ������, 0 As ��ִ����,
                  Sum(Nvl(b.����, 1) * b.ʵ������) As δִ����, 0 As ִ�б�־, a.��¼״̬, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��, a.�շѷ�ʽ
           From ҽ�����ü�¼ A, ҩƷ�շ���¼ B
           Where (a.�շ���� In ('5', '6', '7') Or (a.�շ���� = '4' And Nvl(a.��������, 0) = 1)) And a.����id = b.����id And
                 a.No = b.No And b.���� In (9, 10, 25, 26) And Mod(b.��¼״̬, 3) = 1 And b.����� Is Null
           Group By a.����id, a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.��������, a.�������, a.ҽ������, a.��������, a.ʣ������,
                    a.�շѷ�ʽ
           Having Sum(Nvl(b.����, 1) * b.ʵ������) > 0
           Union All
           Select a.����id, a.No, a.���, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.��������, a.�������, a.ҽ������, a.��������, a.ʣ������,
                  Sum(Nvl(b.����, 1) * b.ʵ������) As ��ִ����, 0 As δִ����, 1 As ִ�б�־, a.��¼״̬, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��, a.�շѷ�ʽ
           From ҽ�����ü�¼ A, ҩƷ�շ���¼ B
           Where (a.�շ���� In ('5', '6', '7') Or (a.�շ���� = '4' And Nvl(a.��������, 0) = 1)) And a.����id = b.����id And
                 a.No = b.No And b.���� In (9, 10, 25, 26) And Not (Mod(b.��¼״̬, 3) = 1 And b.����� Is Null)
           Group By a.����id, a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.��������, a.�������, a.ҽ������, a.��������, a.ʣ������,
                    a.�շѷ�ʽ
           Having Sum(Nvl(b.����, 1) * b.ʵ������) > 0) A
           Order By Decode(a.�������, '5', 0, '6', 0, '7', 0, a.�շ�ϸĿid), a.ִ�б�־, a.�Ǽ�ʱ�� Desc;


  Cursor c_Applay(v_����ids Varchar2) Is
    Select a.����id, b.No, b.���, a.����, a.����ʱ��, a.�������
    From ���˷������� A, ������ü�¼ B
    Where a.����id = b.Id And a.���벿��id = a.��˲���id And a.����ʱ�� = �ջ�ʱ��_In And
          a.����id In (Select * From Table(Cast(f_Num2list(v_����ids) As Zltools.t_Numlist)))
    Order By NO, ���;

  --����ָ��ҩƷ��������ʱ��������ط��ü�ҩƷ/���ļ�¼��Ϣ(���η����ж�����¼,���������ڽ����ֹ)
  --ҩƷҽ����д��"����ҽ������"��¼,��Ӧ�ĸ�ҩ;����һ����д�˵�(����Ϊ����),��NO��ͬ��
  --��ΪҪ�ջصĴ������ܰ����˶�η��͵�����,����Ҫ����η��͵��շ���¼��ȡ��������η���ʱ�����۵����ջأ��޸Ļ�ɾ����
  Cursor c_Drug Is
    Select a.����id, a.��ҳid, d.����, Nvl(Nvl(x.����ϵ��, y.����ϵ��), 1) As ����ϵ��, Nvl(x.סԺ��װ, 1) As סԺ��װ,
           Nvl(x.���Ч��, y.���Ч��) As ���Ч��, Nvl(b.����, 1) * b.ʵ������ As ����, b.Id As �շ�id, b.����, b.ҩƷid, b.�Է�����id, b.�ⷿid,
           b.����id, Nvl(Nvl(x.ҩ������, y.���÷���), 0) As ����, b.����, b.����, b.Ч��, a.��¼״̬, a.No, a.���, a.�շ�ϸĿid, a.ִ��״̬ As ִ�б�־
    From ������ü�¼ A, ҩƷ�շ���¼ B, ����ҽ������ C, ������Ϣ D, ҩƷ��� X, �������� Y
    Where c.ҽ��id = ҽ��id_In And a.No = c.No And a.��¼���� = c.��¼���� And a.��¼״̬ In (0, 1, 3) And a.ҽ����� + 0 = ҽ��id_In And
          a.No = b.No And a.Id = b.����id + 0 And b.���� In (9, 10, 25, 26) And (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0) And
          a.����id = d.����id And b.ҩƷid = x.ҩƷid(+) And b.ҩƷid = y.����id(+)
    Order By a.��¼״̬, b.No Desc, b.Id Desc;

  --������ҩ����(����ҩ;��)����ʱ�������ķ���(����������ж�����¼)
  --�Է�ҩҽ��,ֱ���ջ�ָ����,���ܶ�η���(�����η��ͼ۸�ͬ,���ջصļ۸��������εģ���Ȼ��Ҫ���ݶ���������μ��ջ���)��
  --���ı������ۼ۵�λ������סԺ��λת��
  --��ҩ��������д�˷��ͼ�¼(�����˶���������ȼ�)
  --һ��ֻ��һ�λ�һ�η���ֻ��һ�ε���Ŀ��ʱ��֧�ָ�������
  Cursor c_Other(n_���ͺ� ����ҽ������.���ͺ�%Type) Is
    With ҽ�����ü�¼ As
     (Select a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.Id As ����id, a.���� As ʣ������, Nvl(a.ִ��״̬, 0) As ִ��״̬, a.ҽ�����, b.���ͺ�,
             c.���� As ��������, Nvl(c.�շѷ�ʽ, 0) As �շѷ�ʽ, a.�շ����
      From ������ü�¼ A, ����ҽ������ B, ����ҽ���Ƽ� C
      Where a.No = b.No And a.��¼���� = b.��¼���� And a.ҽ����� + 0 = b.ҽ��id And b.ҽ��id = ҽ��id_In And a.ҽ����� = c.ҽ��id(+) And
            a.�շ�ϸĿid = c.�շ�ϸĿid(+))
    Select a.No, a.���, a.����id, a.ʣ������, a.�շ�ϸĿid, a.��¼״̬, a.ִ��״̬, a.��������, a.�շѷ�ʽ, a.�շ����
    From (Select a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.����id, a.ʣ������, a.��������, a.ִ��״̬, a.ҽ�����, a.�շѷ�ʽ, a.�շ����
           From ҽ�����ü�¼ A
           Where a.��¼״̬ In (1, 3) And a.���ͺ� = n_���ͺ�
           Union All
           Select a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.����id, a.ʣ������, a.��������, a.ִ��״̬, a.ҽ�����, a.�շѷ�ʽ, a.�շ����
           From ҽ�����ü�¼ A
           Where a.��¼״̬ = 0) A
    Order By a.�շ�ϸĿid, a.���, a.��¼״̬;

  --�����������Ϊ�˲����¼�¼ʱ,��дͬһ�շ�ϸĿ�Ĳ�ͬ������Ŀ�ļ۸񸸺�

  --���α����ڴ��������ػ��ܱ�
  Cursor c_Money
  (
    v_Start ������ü�¼.���%Type,
    v_End   ������ü�¼.���%Type
  ) Is
    Select ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Sum(Nvl(Ӧ�ս��, 0)) As Ӧ�ս��, Sum(Nvl(ʵ�ս��, 0)) As ʵ�ս��
    From ������ü�¼
    Where ��¼���� = 2 And ��¼״̬ = 1 And NO = No_In And ��� Between v_Start And v_End
    Group By ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid;

  --ϵͳ����ָ��ִ�к���Ҫ�Զ���˵Ļ��۷��ã����ڷ�ҩҽ����������Ӧ��ҩƷ�����ķ���
  Cursor c_Verify
  (
    v_Start ������ü�¼.���%Type,
    v_End   ������ü�¼.���%Type
  ) Is
    Select NO, ���
    From ������ü�¼
    Where ��¼���� = 2 And ��¼״̬ = 0 And NO = No_In And �۸񸸺� Is Null And ��� Between v_Start And v_End;

  Cursor c_Compound
  (
    ���id_In       ����ҽ����¼.���id%Type,
    ִ����ֹʱ��_In ����ҽ����¼.ִ����ֹʱ��%Type,
    ��ҩid_In       ��Һ��ҩ��¼.Id%Type,
    ҽ�����_In     ����ҽ����¼.Id%Type
  ) Is
    Select b.����id, b.ҩƷid As �շ�ϸĿid, Sum(a.����) As ����, c.סԺ��װ, c.סԺ��λ, d.����, e.���˲���id, e.����״̬, e.Id As ��ҩid, f.No,
           Nvl(f.�۸񸸺�, f.���) As ���, f.��¼״̬ As ��¼״̬, f.ִ��״̬ As ִ�б�־
    From ��Һ��ҩ���� A, ҩƷ�շ���¼ B, ҩƷ��� C, �շ���ĿĿ¼ D, ��Һ��ҩ��¼ E, ������ü�¼ F
    Where a.�շ�id = b.Id And b.ҩƷid = c.ҩƷid And c.ҩƷid = d.Id And e.Id = a.��¼id And f.No = b.No And f.Id = b.����id And
          e.ҽ��id = ���id_In And e.ִ��ʱ�� > ִ����ֹʱ��_In And e.Id = ��ҩid_In And f.ҽ����� + 0 = ҽ�����_In
    Group By b.����id, b.ҩƷid, c.סԺ��װ, c.סԺ��λ, d.����, e.���˲���id, e.����״̬, e.Id, f.No, f.�۸񸸺�, f.���, f.��¼״̬, f.ִ��״̬;

  --����а����������õ�δ������ʱ�����ݲ��������Ƿ��Զ�����
  Cursor c_Stuff Is
    Select m.Id, m.�ⷿid, Decode(b.���÷���, 1, m.����, 0) As ����
    From ҩƷ�շ���¼ M, ������ü�¼ A, �������� B
    Where m.No = No_In And m.���� In (25, 26) And m.�ⷿid Is Not Null And m.��¼״̬ = 1 And m.����� Is Null And m.No = a.No And
          a.Id = m.����id + 0 And a.��¼���� = 2 And a.��¼״̬ = 1 And a.�շ�ϸĿid = b.����id And b.�������� = 1
    Order By m.�ⷿid, m.ҩƷid;

  v_Dec      Number;
  v_First    Number;
  v_������� Varchar2(255);

  v_������� ����ҽ����¼.�������%Type;
  v_�������� ����ҽ����¼.��������%Type;
  v_�������� ��������.��������%Type;

  v_������� ������ü�¼.���%Type;
  v_�շ���� ҩƷ�շ���¼.���%Type;
  v_����id   ������ü�¼.Id%Type;
  v_ʵ�ս�� ������ü�¼.ʵ�ս��%Type;

  v_��ʼ��� ������ü�¼.���%Type;
  v_������� ������ü�¼.���%Type;

  v_ҽ��ִ�� ����ҽ������.ִ��״̬%Type;

  v_����ϵ�� ҩƷ���.����ϵ��%Type;
  v_סԺ��װ ҩƷ���.סԺ��װ%Type;
  v_ҽ������ ����ҽ����¼.ҽ������%Type;

  v_���ʲ���       Zlparameters.����ֵ%Type;
  v_��Һҩ�������� Zlparameters.����ֵ%Type;
  v_���ʽ��       ������ü�¼.���ʽ��%Type;

  v_�շ�ϸĿid   ������ü�¼.�շ�ϸĿid%Type;
  v_ʣ������     ������ü�¼.����%Type;
  v_�ջ�����     ������ü�¼.����%Type;
  v_��ǰ����     ������ü�¼.����%Type;
  v_��ǰ����     ������ü�¼.����%Type;
  v_����ids      Varchar2(4000);
  v_��id         ����ҽ����¼.Id%Type;
  v_��������     ����ҽ���Ƽ�.����%Type;
  v_�ջ���       ������ü�¼.����%Type;
  v_�ջ�ʣ��     ������ü�¼.����%Type;
  v_��Һ�ջ�ʣ�� ������ü�¼.����%Type;
  n_����         ҩƷ�շ���¼.��д����%Type;

  v_Delno    Varchar2(4000);
  v_Temp     Varchar2(4000);
  v_�շ����� Varchar2(4000);
  v_No       ������ü�¼.No%Type;
  v_��Ա��� ������ü�¼.����Ա���%Type;
  v_��Ա���� ������ü�¼.����Ա����%Type;

  n_���id       ����ҽ����¼.���id%Type;
  d_ִ����ֹʱ�� ����ҽ����¼.ִ����ֹʱ��%Type;
  b_��Һ��ҩ��¼ Boolean;
  d_�ջ�ʱ��     ����ҽ����¼.ִ����ֹʱ��%Type;
  n_�������     ���˷�������.�������%Type;
  v_����ԭ��     ���˷�������.����ԭ��%Type;
  n_Count        Number;
  v_Lngid        ҩƷ�շ���¼.Id%Type; --�շ�ID
  n_Tmp���      ����ҽ����¼.���%Type;
  n_��Һ����     Number; ----�Ƿ������Һ��ҩ��¼��״̬
  n_���Ϻ�       ҩƷ�շ���¼.���ܷ�ҩ��%Type;
  n_�ⷿid       ҩƷ�շ���¼.�ⷿid%Type;
  v_�շ�ids      Varchar2(4000);
  v_Error        Varchar2(255);
  Err_Custom Exception;

  Procedure �����շ���¼_Insert
  (
    ����id_In     Number,
    ����_In       ҩƷ�շ���¼.����%Type,
    ����_In       ҩƷ���.ҩ������%Type,
    ����_In       ҩƷ�շ���¼.����%Type,
    Ч��_In       ҩƷ�շ���¼.Ч��%Type,
    ���Ч��_In   ҩƷ���.���Ч��%Type,
    �շ�id_In     ҩƷ�շ���¼.Id%Type,
    ����id_In     ������ü�¼.����id%Type,
    ��ҳid_In     ������ü�¼.��ҳid%Type,
    �ⷿid_In     ҩƷ�շ���¼.�ⷿid%Type,
    ����_In       ҩƷ�շ���¼.����%Type,
    ����_In       ������Ϣ.����%Type,
    �Է�����id_In ҩƷ�շ���¼.�Է�����id%Type,
    �շ����_In   ������ü�¼.�շ����%Type,
    �������_In   Varchar,
    P����         ҩƷ�շ���¼.����%Type,
    P����         ҩƷ�շ���¼.��д����%Type
  ) Is
    v_����   ҩƷ�շ���¼.����%Type;
    v_Ч��   ҩƷ�շ���¼.Ч��%Type;
    v_����   ҩƷ�շ���¼.����%Type;
    v_���ȼ� ���.���ȼ�%Type;
  Begin
    --ȷ������
    If Nvl(����_In, 0) <> 0 And ����_In = 0 Then
      --ԭ����,�ֲ�����
      v_���� := Null;
      v_���� := ����_In;
      v_Ч�� := Ч��_In;
    Elsif Nvl(����_In, 0) = 0 And ����_In = 1 Then
      --ԭ������,�ַ���
      Select ҩƷ�շ���¼_Id.Nextval Into v_���� From Dual;
      Select To_Char(Sysdate, 'YYYYMMDD') Into v_���� From Dual;
      If ���Ч��_In Is Not Null Then
        v_Ч�� := Trunc(Sysdate + ���Ч��_In * 30);
      Else
        v_Ч�� := Null;
      End If;
    Else
      v_���� := ����_In;
      v_���� := ����_In;
      v_Ч�� := Ч��_In;
    End If;
  
    Select ҩƷ�շ���¼_Id.Nextval Into v_Lngid From Dual;
    Insert Into ҩƷ�շ���¼
      (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ����, ��д����, ʵ������, ���ۼ�, ���۽��, ժҪ, ������, ��������,
       ����id, ����, Ƶ��, �÷�, ��ҩ��λid, ��������, ��׼�ĺ�, ���Ч��)
      Select v_Lngid, 1, ����, No_In, v_�շ����, �ⷿid, �Է�����id, ������id, -1, ҩƷid, Nvl(v_����, 0), ����, v_����, v_Ч��, P����, -1 * P����,
             -1 * P����, ���ۼ�, Round(-1 * P���� * P���� * ���ۼ�, v_Dec), '���ڷ����ջ�', v_��Ա����, �ջ�ʱ��_In, ����id_In, ����, Ƶ��, �÷�, ��ҩ��λid,
             ��������, ��׼�ĺ�, ���Ч��
      From ҩƷ�շ���¼
      Where ID = �շ�id_In;
  
    Zl_δ��ҩƷ��¼_Insert(v_Lngid);
  
    Zl_ҩƷ���_Update(v_Lngid, 0, 1);
  
    --δ��ҩƷ��¼
    Update δ��ҩƷ��¼
    Set ����id = ����id_In, ��ҳid = ��ҳid_In, ���� = ����_In
    Where ���� = ����_In And NO = No_In And �ⷿid + 0 = �ⷿid_In;
  
    If Sql%RowCount = 0 Then
      --ȡ������ȼ�
      Begin
        Select b.���ȼ� Into v_���ȼ� From ������Ϣ A, ��� B Where a.��� = b.����(+) And a.����id = ����id_In;
      Exception
        When Others Then
          Null;
      End;
    
      Insert Into δ��ҩƷ��¼
        (����, NO, ����id, ��ҳid, ����, ���ȼ�, �Է�����id, �ⷿid, ��������, ���շ�, ��ӡ״̬)
      Values
        (����_In, No_In, ����id_In, ��ҳid_In, ����_In, v_���ȼ�, �Է�����id_In, �ⷿid_In, �ջ�ʱ��_In,
         Decode(Nvl(Instr(�������_In, Decode(�շ����_In, '4', '4', '5')), 0), 0, 1, 0), 0);
    End If;
  
    v_�շ���� := v_�շ���� + 1;
  End;
Begin
  --ȡ����Ա��Ϣ(����ID,��������;��ԱID,��Ա���,��Ա����)
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
  --����Ƿ�����Һ��Һ��¼�����Ƿ��Ѿ�����
  Select a.ҽ������, Nvl(a.���id, a.Id) Into v_ҽ������, n_���id From ����ҽ����¼ A Where a.Id = ҽ��id_In;
  Select Count(1)
  Into n_Count
  From ��Һ��ҩ��¼ A, ����ҽ����¼ B
  Where a.ҽ��id = b.Id And ҽ��id = ҽ��id_In And a.ִ��ʱ�� > b.ִ����ֹʱ�� And a.�Ƿ����� = 1;

  If n_Count > 0 Then
    v_Error := 'ҽ��"' || v_ҽ������ || '"����ҺҩƷ���Ѿ�����Һ�����������������ܳ����ջء�';
    Raise Err_Custom;
  End If;
  Select Max(����˵��) Into v_����ԭ�� From ����ҽ��״̬ Where ҽ��id = ҽ��id_In And �������� = 8;
  If Nvl(�ջ���_In, 0) > 0 Then
    --�ж��Ƿ�����Һ��ҩҩƷ(��Һ��������ҩƷͳһ����������)
    b_��Һ��ҩ��¼ := False;
    v_��Һ�ջ�ʣ�� := �ջ���_In;
  
    Select Max(a.ִ����ֹʱ��)
    Into d_ִ����ֹʱ��
    From ��Һ��ҩ��¼ E, ����ҽ����¼ A
    Where a.Id = n_���id And e.ҽ��id = a.Id And e.ִ��ʱ�� > a.ִ����ֹʱ��;
  
    If d_ִ����ֹʱ�� Is Not Null Then
      d_�ջ�ʱ��       := �ջ�ʱ��_In;
      v_��Һҩ�������� := zl_GetSysParameter('��Һ��Һ����ҩ��������������', 1345);
      b_��Һ��ҩ��¼   := True;
    
      If n_���id = ҽ��id_In Then
        --��ҩ;���У�����״̬�������ı�����
        n_��Һ���� := 1;
      Else
        n_��Һ���� := 0;
        n_Tmp���  := ҽ��id_In;
      End If;
    
      For X In (Select e.Id As ��ҩid, e.����״̬, e.�Ƿ���
                From ��Һ��ҩ��¼ E
                Where e.ҽ��id = n_���id And e.ִ��ʱ�� > d_ִ����ֹʱ�� And Nvl(e.����״̬, 0) In (1, 2, 3, 4, 5, 6, 7, 8)) Loop
        If Not (x.����״̬ In (4, 5, 6, 7, 8) And Nvl(x.�Ƿ���, 0) = 0 And Nvl(v_��Һҩ��������, '0') = '0') Then
          If n_��Һ���� = 0 Then
            --����ҩƷ����ϸ��������
            For r_Compound In c_Compound(n_���id, d_ִ����ֹʱ��, x.��ҩid, n_Tmp���) Loop
            
              v_��Һ�ջ�ʣ�� := v_��Һ�ջ�ʣ�� - r_Compound.����;
              If x.����״̬ = 1 Then
                n_������� := 0;
              Else
                n_������� := 1;
              End If;
              Zl_���˷�������_Insert(r_Compound.����id, r_Compound.�շ�ϸĿid, r_Compound.���˲���id, r_Compound.����, v_��Ա����, d_�ջ�ʱ��,
                               n_�������, Null, r_Compound.��ҩid, v_����ԭ��, 0);
              If x.����״̬ = 1 Then
                --δ��ҩ�ģ��Զ���ˡ�
                Zl_���˷�������_Audit(r_Compound.����id, d_�ջ�ʱ��, v_��Ա����, d_�ջ�ʱ��, 1, 1, n_�������);
                Zl_������ʼ�¼_Delete(r_Compound.No, r_Compound.��� || ':' || r_Compound.���� || ':' || r_Compound.��ҩid, v_��Ա���,
                                 v_��Ա����, 0, d_�ջ�ʱ��);
              End If;
            End Loop;
          End If;
        
          --����״̬
          If n_��Һ���� = 1 Then
            Select Count(1)
            Into n_Count
            From ��Һ��ҩ״̬
            Where ��ҩid = x.��ҩid And �������� = 9 And ����ʱ�� = d_�ջ�ʱ��;
            If n_Count = 0 Then
              Insert Into ��Һ��ҩ״̬
                (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
              Values
                (x.��ҩid, 9, v_��Ա����, d_�ջ�ʱ��, v_����ԭ��);
            End If;
            Update ��Һ��ҩ��¼ Set ������Ա = v_��Ա����, ����ʱ�� = d_�ջ�ʱ��, ����״̬ = 9 Where ID = x.��ҩid;
          
            If x.����״̬ = 1 Then
              Insert Into ��Һ��ҩ״̬
                (��ҩid, ��������, ������Ա, ����ʱ��)
              Values
                (x.��ҩid, 10, v_��Ա����, d_�ջ�ʱ��);
              Update ��Һ��ҩ��¼ Set ������Ա = v_��Ա����, ����ʱ�� = d_�ջ�ʱ��, ����״̬ = 10 Where ID = x.��ҩid;
            End If;
          End If;
        
          --���ڲ�ͬ���Σ�ִ��ʱ�䣩����ʱ������ʱ��ͷ���ID��ΨһԼ��������ͬʱ���ʶ������ʱ�����μ�һ��
          d_�ջ�ʱ�� := d_�ջ�ʱ�� + 1 / 24 / 60 / 60;
        End If;
      End Loop;
    End If;
  
    --a.���������ջ�ģʽ
    --��Һ��ҩ��¼������������
    If b_��Һ��ҩ��¼ = False Or v_��Һ�ջ�ʣ�� > 0 Then
      If No_In Is Null Then
        v_���ʲ��� := zl_GetSysParameter(23);
        --�����ջ���������ԭʼ���ý��з�̯����
        For r_Detail In c_Detail Loop
          --ȷ�����շ�ϸĿID���ջ�������
          If Nvl(v_�շ�ϸĿid, 0) <> r_Detail.�շ�ϸĿid And (r_Detail.������� Not In ('5', '6', '7') Or Nvl(v_�շ�ϸĿid, 0) = 0) Then
            --����δ��̯���
            If v_�շ�ϸĿid Is Not Null And v_�ջ����� > 0 Then
              v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ����������ܴ����ֹ����ʻ��ѽ��ʵķ��ã�����Ӧ�Ļ��۵��ѱ�ɾ����';
              Raise Err_Custom;
            End If;
            --ҩƷ�ջ�������������͹��Ϊ׼����ģ��Դ˼�����ջ��ۼ�����
            Begin
              Select ����ϵ��, סԺ��װ Into v_����ϵ��, v_סԺ��װ From ҩƷ��� Where ҩƷid = r_Detail.�շ�ϸĿid;
            Exception
              When Others Then
                Null;
            End;
            --�����ջ�����������շѷ�ʽ����0������ȡ����ʹ������ķ������м���
            If r_Detail.�շѷ�ʽ = 0 Then
              If r_Detail.������� = '7' Then
                --��ҩ�䷽ҩƷ������*����
                v_�ջ����� := Round(v_��Һ�ջ�ʣ�� * r_Detail.�������� / Nvl(v_����ϵ��, 1), 5);
              Else
                If r_Detail.������� Not In ('5', '6') Then
                  Select Nvl(Max(����), 1)
                  Into v_��������
                  From ����ҽ���Ƽ�
                  Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Detail.�շ�ϸĿid;
                Else
                  v_�������� := 1;
                End If;
                v_�ջ����� := Round(v_��Һ�ջ�ʣ�� * Nvl(v_סԺ��װ, 1), 5) * v_��������;
              End If;
            Else
              Select Nvl(Sum(����), 0)
              Into v_�ջ�����
              From ҽ��ִ�мƼ�
              Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Detail.�շ�ϸĿid And
                    Ҫ��ʱ�� > Nvl(�ϴ�ʱ��_In, To_Date('1900-01-01', 'yyyy-MM-dd'));
            
              v_�ջ����� := Round(v_�ջ�����, 5);
            
            End If;
            v_ҽ������ := r_Detail.ҽ������;
          End If;
        
          --���շ�ϸĿ��ÿ��������ϸ��̯�ջ�
          If v_�ջ����� > 0 Then
            --����Ӧ�����Ƿ��ѽ��ʣ�����ֹʱ
            v_���ʽ�� := 0;
            If v_���ʲ��� = '2' And r_Detail.��¼״̬ <> 0 Then
              Select Sum(���ʽ��)
              Into v_���ʽ��
              From ������ü�¼
              Where NO = r_Detail.No And ��¼���� In (2, 12) And Nvl(�۸񸸺�, ���) = r_Detail.���;
            End If;
          
            If Nvl(v_���ʽ��, 0) = 0 Then
              If r_Detail.�շ���� In ('5', '6', '7') Or r_Detail.�շ���� = '4' And r_Detail.�������� = 1 Then
                --ҩƷ�͸������õ�����
                If r_Detail.ִ�б�־ = 0 Then
                  v_ʣ������ := r_Detail.δִ����;
                Elsif r_Detail.ִ�б�־ = 1 Then
                  v_ʣ������ := r_Detail.��ִ����;
                End If;
              Else
                --��ͨ����
                v_ʣ������ := r_Detail.ʣ������;
              End If;
              If v_�ջ����� > v_ʣ������ Then
                v_��ǰ���� := v_ʣ������;
              Else
                v_��ǰ���� := v_�ջ�����;
              End If;
              v_�ջ����� := v_�ջ����� - v_��ǰ����;
              --ϵͳ��������ִ�к��Ƿ���˻��۵������ԣ���ִ�е���Ȼ�����ǻ��۵�
              If r_Detail.ִ�б�־ = 0 And r_Detail.��¼״̬ = 0 Then
                v_Delno := v_Delno || '|' || r_Detail.No || ',' || r_Detail.��� || ':' || v_��ǰ����;
              Else
                If Not (r_Detail.�շ���� = '7' And r_Detail.ִ�б�־ <> 0) Then
                  Zl_���˷�������_Insert(r_Detail.����id, r_Detail.�շ�ϸĿid, r_Detail.���˲���id, v_��ǰ����, v_��Ա����, �ջ�ʱ��_In,
                                   r_Detail.ִ�б�־, Null, Null, v_����ԭ��);
                End If;
              End If;
              v_����ids := v_����ids || ',' || r_Detail.����id;
            End If;
          End If;
          v_�շ�ϸĿid := r_Detail.�շ�ϸĿid;
        End Loop;
      
        --����δ��̯���
        If v_�ջ����� > 0 Then
          v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ����������ܴ����ֹ����ʻ��ѽ��ʵķ��ã�����Ӧ�Ļ��۵��ѱ�ɾ����';
          Raise Err_Custom;
        End If;
        --���Ƶ����������Զ����
        If zl_GetSysParameter('�����ջط��ñ����Զ����', 1254) = '1' And v_����ids Is Not Null Then
          For r_Applay In c_Applay(Substr(v_����ids, 2)) Loop
            Zl_���˷�������_Audit(r_Applay.����id, r_Applay.����ʱ��, v_��Ա����, �ջ�ʱ��_In, 1, 1, r_Applay.�������);
            v_Delno := v_Delno || '|' || r_Applay.No || ',' || r_Applay.��� || ':' || r_Applay.����;
          End Loop;
        End If;
      Else
        ---b.�����ջ�ģʽ-------------------------------------------------------------------------------------------------------
        --���ȫ�ǻ��۵����Ͳ��ò���������������
        If No_In = '�������۵�' Then
          --δ��˵Ļ��۵����Ƚ����޸Ļ�ɾ�������ܶ�η���Ϊ��ͬ��NO,Ϊ�˼���ÿ�ε��ջ�������Ҫ���շ�ϸĿID����
        
          For r_Price In (Select c.�������, b.No, b.���, b.�շ�ϸĿid, Nvl(b.����, 1) * b.���� As ʣ������, c.��������, d.����ϵ��, d.סԺ��װ,
                                 c.ҽ������, Nvl(e.�շѷ�ʽ, 0) As �շѷ�ʽ
                          From ����ҽ������ A, ������ü�¼ B, ����ҽ����¼ C, ҩƷ��� D, ����ҽ���Ƽ� E
                          Where a.ҽ��id = ҽ��id_In And a.No = b.No And a.��¼���� = b.��¼���� And a.ҽ��id = b.ҽ����� And
                                b.�۸񸸺� Is Null And b.�շ�ϸĿid = d.ҩƷid(+) And b.��¼״̬ = 0 And c.Id = a.ҽ��id And
                                b.ҽ����� = e.ҽ��id(+) And b.�շ�ϸĿid = e.�շ�ϸĿid(+)
                          Order By �շ�ϸĿid, NO Desc) Loop
            If Nvl(v_�շ�ϸĿid, 0) <> r_Price.�շ�ϸĿid Then
              --����δ��̯���
              If v_�շ�ϸĿid Is Not Null And v_�ջ����� > 0 Then
                v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ�������������ػ��۵��ѱ�ɾ������ˡ�';
                Raise Err_Custom;
              End If;
              --�����ջ�����������շѷ�ʽ����0������ȡ����ʹ������ķ������м���
              If r_Price.�շѷ�ʽ = 0 Then
                If r_Price.������� = '7' Then
                  --��ҩ�䷽ҩƷ������*����
                  v_�ջ����� := Round(v_��Һ�ջ�ʣ�� * r_Price.�������� / Nvl(r_Price.����ϵ��, 1), 5);
                Else
                  If r_Price.������� Not In ('5', '6') Then
                    Select Nvl(Max(����), 1)
                    Into v_��������
                    From ����ҽ���Ƽ�
                    Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Price.�շ�ϸĿid;
                  Else
                    v_�������� := 1;
                  End If;
                  v_�ջ����� := Round(v_��Һ�ջ�ʣ�� * Nvl(r_Price.סԺ��װ, 1), 5) * v_��������;
                End If;
              Else
                Select Nvl(Sum(����), 0)
                Into v_�ջ�����
                From ҽ��ִ�мƼ�
                Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Price.�շ�ϸĿid And
                      Ҫ��ʱ�� > Nvl(�ϴ�ʱ��_In, To_Date('1900-01-01', 'yyyy-MM-dd'));
              
                v_�ջ����� := Round(v_�ջ�����, 5);
              End If;
              v_ҽ������ := r_Price.ҽ������;
            End If;
            If v_�ջ����� > 0 Then
              If v_�ջ����� > r_Price.ʣ������ Then
                v_��ǰ���� := r_Price.ʣ������;
              Else
                v_��ǰ���� := v_�ջ�����;
              End If;
              v_�ջ����� := v_�ջ����� - v_��ǰ����;
              v_Delno    := v_Delno || '|' || r_Price.No || ',' || r_Price.��� || ':' || v_��ǰ����;
            End If;
            v_�շ�ϸĿid := r_Price.�շ�ϸĿid;
          End Loop;
        
          --����δ��̯���
          If v_�ջ����� > 0 Then
            v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ�������������ػ��۵��ѱ�ɾ������ˡ�';
            Raise Err_Custom;
          End If;
        Else
          --�������������ܴ��ڻ��۵�����ʵ���ϵ����
          --���С��λ��
          Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;
          --���ɻ��۵�ϵͳ����
          Select zl_GetSysParameter(80) Into v_������� From Dual;
          v_��ʼ��� := Null;
          v_������� := Null;
        
          Select a.�������, a.��������, b.��������
          Into v_�������, v_��������, v_��������
          From ����ҽ����¼ A, �������� B
          Where ID = ҽ��id_In And a.�շ�ϸĿid = b.����id(+);
        
          If v_������� In ('5', '6', '7') Or (v_������� = '4' And Nvl(v_��������, 0) = 1) Then
            --ҩƷ������
            -----------------------------------------------------------------------------------------------------
            v_�ջ����� := Null;
            Select Nvl(Max(���), 0) + 1
            Into v_�շ����
            From ҩƷ�շ���¼
            Where ���� In (9, 10, 25, 26) And ��¼״̬ = 1 And NO = No_In;
          
            Select Nvl(Max(���), 0) + 1
            Into v_�������
            From ������ü�¼
            Where ��¼���� = 2 And ��¼״̬ In (0, 1) And NO = No_In;
          
            --һ��ҽ����ҩƷֻ��һ�У������ѭ����Ϊ�˴����η��͵����������ҩƷ�ڽ����ѽ��ø����ջ�
            For r_Drug In c_Drug Loop
              --��ʼ��Ҫ�ջص�������(��������)
              v_First := 0;
              If v_�ջ����� Is Null Then
                If v_������� = '7' Then
                  v_�ջ����� := Round(v_��Һ�ջ�ʣ�� * v_�������� / r_Drug.����ϵ��, 5);
                Else
                  If v_������� Not In ('5', '6') Then
                    Select Nvl(Max(����), 1)
                    Into v_��������
                    From ����ҽ���Ƽ�
                    Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Drug.�շ�ϸĿid;
                  Else
                    v_�������� := 1;
                  End If;
                  v_�ջ����� := Round(v_��Һ�ջ�ʣ�� * r_Drug.סԺ��װ, 5) * v_��������;
                End If;
                v_First := 1;
              End If;
            
              --�����һ���������㹻���򰴸����������������ô���
              If v_�ջ����� > r_Drug.���� Then
                v_��ǰ���� := 1;
                v_��ǰ���� := r_Drug.����;
                v_�ջ����� := v_�ջ����� - r_Drug.����;
              Else
                If v_First = 1 And v_������� = '7' Then
                  v_��ǰ���� := v_��Һ�ջ�ʣ��;
                  v_��ǰ���� := Round(v_�������� / r_Drug.����ϵ��, 5);
                Else
                  v_��ǰ���� := 1;
                  v_��ǰ���� := v_�ջ�����;
                End If;
                v_�ջ����� := 0;
              End If;
            
              If r_Drug.��¼״̬ = 0 Then
                v_Delno := v_Delno || '|' || r_Drug.No || ',' || r_Drug.��� || ':' || v_��ǰ���� * v_��ǰ����;
              Else
                If Not (v_������� = '7' And r_Drug.ִ�б�־ <> 0) Then
                
                  Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
                  �����շ���¼_Insert(v_����id, r_Drug.����, r_Drug.����, r_Drug.����, r_Drug.Ч��, r_Drug.���Ч��, r_Drug.�շ�id,
                                r_Drug.����id, r_Drug.��ҳid, r_Drug.�ⷿid, r_Drug.����, r_Drug.����, r_Drug.�Է�����id, v_�������,
                                v_�������, v_��ǰ����, v_��ǰ����);
                
                  --������ü�¼
                  -------------------------------------------------------------------------------------
                  --��¼��ŷ�Χ�Դ�����ܱ�
                  If v_��ʼ��� Is Null Then
                    v_��ʼ��� := v_�������;
                  End If;
                  v_������� := v_�������;
                
                  Insert Into ������ü�¼
                    (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ���˲���id, ���˿���id, �ѱ�, �շ����,
                     �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���,
                     ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ҽ�����, ������, ����Ա���, ����Ա����)
                    Select v_����id, 2, No_In, Decode(Nvl(Instr(v_�������, Decode(v_�������, '4', '4', '5')), 0), 0, 1, 0),
                           v_�������, Null, Null, 1, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ,
                           ������Ŀ��, ���մ���id, v_��ǰ����, -1 * v_��ǰ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����,
                           Round(-1 * v_��ǰ���� * v_��ǰ���� * ��׼����, v_Dec), Round(-1 * v_��ǰ���� * v_��ǰ���� * ��׼����, v_Dec), Null, 1,
                           ��������id, ������, �ջ�ʱ��_In, �ջ�ʱ��_In, ִ�в���id, 0, ҽ�����, v_��Ա����,
                           Decode(Nvl(Instr(v_�������, Decode(v_�������, '4', '4', '5')), 0), 0, v_��Ա���, Null),
                           Decode(Nvl(Instr(v_�������, Decode(v_�������, '4', '4', '5')), 0), 0, v_��Ա����, Null)
                    From ������ü�¼
                    Where ID = r_Drug.����id;
                
                  Select Zl_Actualmoney(�ѱ�, �շ�ϸĿid, ������Ŀid, Ӧ�ս��, ����, ִ�в���id)
                  Into v_Temp
                  From ������ü�¼
                  Where ID = v_����id;
                  v_ʵ�ս�� := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
                  Update ������ü�¼ A Set ʵ�ս�� = v_ʵ�ս�� Where ID = v_����id;
                
                  v_������� := v_������� + 1;
                End If;
                If v_�ջ����� <= 0 Then
                  Exit;
                End If;
              End If;
            End Loop;
          
            If v_�ջ����� <> 0 Then
              --û���ջ���������,�շ���¼����������(���¼��ȫ������Ϊ��)
              Null;
            End If;
          Else
            --������ҩҽ��(������ҩ;�������󶨵����ĵ�)
            -----------------------------------------------------------------------------------------------------
            Select Nvl(Max(���), 0) + 1
            Into v_�շ����
            From ҩƷ�շ���¼
            Where ���� In (9, 10, 25, 26) And ��¼״̬ = 1 And NO = No_In;
            --ȡ�������
            Select Nvl(Max(���), 0) + 1
            Into v_�������
            From ������ü�¼
            Where ��¼���� = 2 And ��¼״̬ In (0, 1) And NO = No_In;
          
            v_�ջ�ʣ�� := v_��Һ�ջ�ʣ��;
          
            For r_Othersend In (Select ���ͺ�, ��������
                                From ����ҽ������
                                Where ҽ��id = ҽ��id_In And ĩ��ʱ�� > Nvl(�ϴ�ʱ��_In, To_Date('1900-01-01', 'yyyy-MM-dd'))
                                Order By ���ͺ� Desc) Loop
              If r_Othersend.�������� < v_�ջ�ʣ�� Then
                --һ���ջض�η��ͣ�����ÿ�η��ͷ��������䶯���Ƽۣ�
                v_�ջ�ʣ�� := v_�ջ�ʣ�� - r_Othersend.��������;
                v_�ջ���   := r_Othersend.��������;
              Else
                --һ�η������ջ�ʣ�ࣻ
                v_�ջ���   := v_�ջ�ʣ��;
                v_�ջ�ʣ�� := 0;
              End If;
              v_�շ����� := '';
              For r_Other In c_Other(r_Othersend.���ͺ�) Loop
                If Nvl(v_�շ�����, '0') <> r_Other.�շ�ϸĿid || ',' || r_Other.��� Then
                  --�������һ�η��͵ķ��ü�¼������Ҫ�ջص�����ȫ���ջ�
                  --�����ջ�����������շѷ�ʽ����0������ȡ����ʹ������ķ������м���
                  If r_Other.�շѷ�ʽ = 0 Then
                    v_�ջ����� := v_�ջ��� * Nvl(r_Other.��������, 1);
                  Else
                    Select Nvl(Sum(����), 0)
                    Into v_�ջ�����
                    From ҽ��ִ�мƼ�
                    Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Other.�շ�ϸĿid And
                          Ҫ��ʱ�� > Nvl(�ϴ�ʱ��_In, To_Date('1900-01-01', 'yyyy-MM-dd'));
                  End If;
                End If;
              
                If v_�ջ����� > 0 Then
                  If r_Other.��¼״̬ = 0 Then
                    If v_�ջ����� > r_Other.ʣ������ Then
                      v_��ǰ���� := r_Other.ʣ������;
                    Else
                      v_��ǰ���� := v_�ջ�����;
                    End If;
                  Else
                    v_��ǰ���� := v_�ջ�����;
                  End If;
                  v_�ջ����� := v_�ջ����� - v_��ǰ����;
                  v_��ǰ���� := 1;
                
                  If r_Other.��¼״̬ = 0 Then
                    v_Delno := v_Delno || '|' || r_Other.No || ',' || r_Other.��� || ':' || v_��ǰ����;
                  Else
                    --��¼��ŷ�Χ�Դ�����ܱ�
                    If v_��ʼ��� Is Null Then
                      v_��ʼ��� := v_�������;
                    End If;
                    v_������� := v_�������;
                  
                    --������ü�¼:��������ջ����������ϴη�����,����ȷ�������ε����Ŀ����ж����շ���¼
                    Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
                    If r_Other.�շ���� In ('4', '5', '6', '7') Then
                      n_���� := v_��ǰ����;
                      For r_Otherdrug In (Select a.����id, a.��ҳid, d.����, Nvl(Nvl(x.����ϵ��, y.����ϵ��), 1) As ����ϵ��,
                                                 Nvl(x.סԺ��װ, 1) As סԺ��װ, Nvl(x.���Ч��, y.���Ч��) As ���Ч��,
                                                 Nvl(b.����, 1) * b.ʵ������ As ����, b.Id As �շ�id, b.����, b.ҩƷid, b.�Է�����id,
                                                 b.�ⷿid, b.����id, Nvl(Nvl(x.ҩ������, y.���÷���), 0) As ����, b.����, b.����, b.Ч��,
                                                 a.��¼״̬, a.No, a.���, a.�շ�ϸĿid
                                          From ������ü�¼ A, ҩƷ�շ���¼ B, ������Ϣ D, ҩƷ��� X, �������� Y
                                          Where a.Id = r_Other.����id And a.��¼״̬ In (0, 1, 3) And a.No = b.No And
                                                a.Id = b.����id + 0 And b.���� In (9, 10, 25, 26) And
                                                (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0) And a.����id = d.����id And
                                                b.ҩƷid = x.ҩƷid(+) And b.ҩƷid = y.����id(+)
                                          Order By a.��¼״̬, b.No Desc, b.Id Desc) Loop
                        If n_���� > 0 Then
                          n_Count := r_Otherdrug.����;
                          If n_���� < n_Count Then
                            n_Count := n_����;
                          End If;
                          �����շ���¼_Insert(v_����id, r_Otherdrug.����, r_Otherdrug.����, r_Otherdrug.����, r_Otherdrug.Ч��,
                                        r_Otherdrug.���Ч��, r_Otherdrug.�շ�id, r_Otherdrug.����id, r_Otherdrug.��ҳid,
                                        r_Otherdrug.�ⷿid, r_Otherdrug.����, r_Otherdrug.����, r_Otherdrug.�Է�����id,
                                        r_Other.�շ����, v_�������, 1, n_Count);
                          n_���� := n_���� - r_Otherdrug.����;
                        End If;
                      End Loop;
                    End If;
                    --ҽ����ִ�У��ջصķ���Ҳ��Ϊ��ִ�У�������ҩƷ�͸������õ����ģ���Ϊʵ�ʷ��ű�ʾִ��
                    Insert Into ������ü�¼
                      (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ���˲���id, ���˿���id, �ѱ�, �շ����,
                       �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���,
                       ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ��ʱ��, ִ����, ҽ�����, ������, ����Ա���, ����Ա����)
                      Select v_����id, 2, No_In, Decode(Nvl(Instr(v_�������, r_Other.�շ����), 0), 0, 1, 0), v_�������, Null,
                             Decode(a.�۸񸸺�, Null, Null, v_������� + a.�۸񸸺� - a.���), 1, a.����id, a.��ҳid, a.��ʶ��, a.����, a.�Ա�,
                             a.����, a.���˲���id, a.���˿���id, a.�ѱ�, a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.������Ŀ��, a.���մ���id, 1, -1 * v_��ǰ����,
                             a.�Ӱ��־, a.���ӱ�־, a.Ӥ����, a.������Ŀid, a.�վݷ�Ŀ, a.��׼����, Round(-1 * v_��ǰ���� * a.��׼����, v_Dec),
                             Round(-1 * v_��ǰ���� * a.��׼����, v_Dec), Null, 1, a.��������id, a.������, �ջ�ʱ��_In, �ջ�ʱ��_In, a.ִ�в���id,
                             Decode(r_Other.ִ��״̬, 1,
                                     Decode(a.�շ����, '4', Decode(b.��������, 1, 0, 1),
                                             Decode(Instr(',5,6,7,', a.�շ����), 0, 1, 0)), 0),
                             Decode(r_Other.ִ��״̬, 1,
                                     Decode(a.�շ����, '4', Decode(b.��������, 1, Null, �ջ�ʱ��_In),
                                             Decode(Instr(',5,6,7,', a.�շ����), 0, �ջ�ʱ��_In, Null)), Null),
                             Decode(r_Other.ִ��״̬, 1,
                                     Decode(a.�շ����, '4', Decode(b.��������, 1, Null, v_��Ա����),
                                             Decode(Instr(',5,6,7,', a.�շ����), 0, v_��Ա����, Null)), Null), a.ҽ�����, v_��Ա����,
                             Decode(Nvl(Instr(v_�������, r_Other.�շ����), 0), 0, v_��Ա���, Null),
                             Decode(Nvl(Instr(v_�������, r_Other.�շ����), 0), 0, v_��Ա����, Null)
                      From ������ü�¼ A, �������� B
                      Where a.Id = r_Other.����id And a.�շ�ϸĿid = b.����id(+);
                  
                    Select Zl_Actualmoney(�ѱ�, �շ�ϸĿid, ������Ŀid, Ӧ�ս��, ����, ִ�в���id)
                    Into v_Temp
                    From ������ü�¼
                    Where ID = v_����id;
                    v_ʵ�ս�� := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
                    Update ������ü�¼ A Set ʵ�ս�� = v_ʵ�ս�� Where ID = v_����id;
                  
                    v_������� := v_������� + 1;
                    v_ҽ��ִ�� := r_Other.ִ��״̬; --����շ���Ŀ��ִ��״̬��һ����
                  End If;
                
                  v_�շ����� := r_Other.�շ�ϸĿid || ',' || r_Other.���;
                End If;
              End Loop;
              If v_�ջ�ʣ�� = 0 Then
                Exit;
              End If;
            End Loop;
          
            --���ҽ����ִ�У���ϵͳ����ִ�к��Զ���˷��ã�������ִ��ҽ����Ӧ��ҩƷ�����ķ��á�
            -----------------------------------------------------------------------------------------------------
            --If Nvl(v_ҽ��ִ��, 0) = 1 And v_��ʼ��� Is Not Null And v_������� Is Not Null Then
            -- For r_Verify In c_Verify(v_��ʼ���, v_�������) Loop
            --   Zl_סԺ���ʼ�¼_Verify(r_Verify.No, v_��Ա���, v_��Ա����, r_Verify.���, Null, �ջ�ʱ��_In);
            -- End Loop;
            --End If;
          End If;
        
          --������û��ܱ�
          -----------------------------------------------------------------------------------------------------
          If v_��ʼ��� Is Not Null And v_������� Is Not Null Then
            --���ͳһ���������ػ��ܱ�
            For r_Money In c_Money(v_��ʼ���, v_�������) Loop
              --�������
              Update �������
              Set ������� = Nvl(�������, 0) + r_Money.ʵ�ս��
              Where ����id = r_Money.����id And ���� = 1 And ���� = 2;
            
              If Sql%RowCount = 0 Then
                Insert Into �������
                  (����id, ����, ����, �������, Ԥ�����)
                Values
                  (r_Money.����id, 1, 2, r_Money.ʵ�ս��, 0);
              End If;
            
              --����δ�����
              Update ����δ�����
              Set ��� = Nvl(���, 0) + r_Money.ʵ�ս��
              Where ����id = r_Money.����id And ��ҳid = r_Money.��ҳid And Nvl(���˲���id, 0) = Nvl(r_Money.���˲���id, 0) And
                    Nvl(���˿���id, 0) = Nvl(r_Money.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Money.��������id, 0) And
                    Nvl(ִ�в���id, 0) = Nvl(r_Money.ִ�в���id, 0) And ������Ŀid + 0 = r_Money.������Ŀid And ��Դ;�� + 0 = 2;
            
              If Sql%RowCount = 0 Then
                Insert Into ����δ�����
                  (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
                Values
                  (r_Money.����id, r_Money.��ҳid, r_Money.���˲���id, r_Money.���˿���id, r_Money.��������id, r_Money.ִ�в���id,
                   r_Money.������Ŀid, 2, r_Money.ʵ�ս��);
              End If;
            End Loop;
          End If;
        End If;
      End If;
    End If;
  End If;

  --����Zl_������ʼ�¼_Delete����֧��ÿ��ɾ��һ�е�ѭ����������������������һ������Ҫɾ�������һ���Դ���
  If Not v_Delno Is Null Then
    v_Temp := '';
    v_No   := '';
    For r_Price In (Select /*+ rule*/
                     C1 As NO, C2 As �������
                    From Table(f_Str2list2(Substr(v_Delno, 2), '|', ','))
                    Order By NO) Loop
      If v_No Is Not Null And v_No <> r_Price.No Then
        Zl_������ʼ�¼_Delete(v_No, v_Temp, v_��Ա���, v_��Ա����, 0);
        v_No := '';
      End If;
      If v_No Is Null Then
        v_No   := r_Price.No;
        v_Temp := r_Price.�������;
      Else
        v_Temp := v_Temp || ',' || r_Price.�������;
      End If;
    End Loop;
    If Not v_No Is Null Then
      Zl_������ʼ�¼_Delete(v_No, v_Temp, v_��Ա���, v_��Ա����, 0);
    End If;
  End If;

  --����ҽ�����ϴ�ִ��ʱ��:��ҩ;���ȿ�����Ϊδ���Ͷ�û�����ջع��̡�
  -----------------------------------------------------------------------------------------------------
  Select Nvl(���id, ID) Into v_��id From ����ҽ����¼ Where ID = ҽ��id_In;
  Update ����ҽ����¼ Set �ϴ�ִ��ʱ�� = �ϴ�ʱ��_In Where ID = v_��id Or ���id = v_��id;

  --ɾ��ҽ��ִ��ʱ��
  If �ϴ�ʱ��_In Is Null Then
    --ȫ���ջ�
    Delete From ҽ��ִ��ʱ�� Where ҽ��id = v_��id;
    Delete From ҽ��ִ�мƼ� Where ҽ��id = ҽ��id_In;
  Else
    --�����ջض�η��͵�����
    Delete From ҽ��ִ��ʱ�� Where ҽ��id = v_��id And Ҫ��ʱ�� > �ϴ�ʱ��_In;
    Delete From ҽ��ִ�мƼ� Where ҽ��id = ҽ��id_In And Ҫ��ʱ�� > �ϴ�ʱ��_In;
  End If;
  --������Һ��Һ��¼���������⣬ÿ��ҽ�������е��ã��ڹ�������ֻ��������Һ��Һ��ҽ��
  Zl_��Һ��ҩ��¼_���ε���(ҽ��id_In);

  If zl_GetSysParameter(63) = '1' And Nvl(No_In, '�������۵�') <> '�������۵�' Then
    --����������������Զ�����
    For r_Stuff In c_Stuff Loop
      If n_���Ϻ� Is Null Then
        n_���Ϻ� := Nextno(20);
      End If;
      If r_Stuff.�ⷿid <> Nvl(n_�ⷿid, 0) Then
        If Nvl(n_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
          v_�շ�ids := Substr(v_�շ�ids, 2);
          Zl_ҩƷ�շ���¼_��������(v_�շ�ids, n_�ⷿid, v_��Ա����, Sysdate, 1, v_��Ա����, n_���Ϻ�, v_��Ա����);
        End If;
      
        n_�ⷿid  := r_Stuff.�ⷿid;
        v_�շ�ids := Null;
      End If;
    
      v_�շ�ids := v_�շ�ids || '|' || r_Stuff.Id || ',' || r_Stuff.����;
    End Loop;
    If Nvl(n_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
      v_�շ�ids := Substr(v_�շ�ids, 2);
      Zl_ҩƷ�շ���¼_��������(v_�շ�ids, n_�ⷿid, v_��Ա����, Sysdate, 1, v_��Ա����, n_���Ϻ�, v_��Ա����);
    End If;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_�ջ�_��������;
/
--140678:��˶,2019-05-07,�ϻ���Ա�䶯��Ϣ֪ͨ
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


  --79.���۲���תסԺ����
  Procedure Zlhis_Patient_029
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );


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
  --�������ϻ���Ա�䶯��Ϣ
  Procedure Zltools_Users_001
  (
    �û���_In In �ϻ���Ա��.�û���%Type,
    ��Աid_In In �ϻ���Ա��.��Աid%Type
  );
  Procedure Zltools_Users_002
  (
    �û���_In In �ϻ���Ա��.�û���%Type,
    ��Աid_In In �ϻ���Ա��.��Աid%Type
  );
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

  --79.���۲���תסԺ����
  Procedure Zlhis_Patient_029
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    n_�䶯id Number(18);
  Begin
    Select max(ID)
    Into n_�䶯id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0 And ��ʼԭ�� = 9;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_029',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  
  End Zlhis_Patient_029;

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
  --�������ϻ���Ա�䶯��Ϣ
  Procedure Zltools_Users_001
  (
    �û���_In In �ϻ���Ա��.�û���%Type,
    ��Աid_In In �ϻ���Ա��.��Աid%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�û���>' || �û���_In || '</�û���><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLTOOLS_USERS_001', v_Value);
  End Zltools_Users_001;
  Procedure Zltools_Users_002
  (
    �û���_In In �ϻ���Ա��.�û���%Type,
    ��Աid_In In �ϻ���Ա��.��Աid%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�û���>' || �û���_In || '</�û���><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLTOOLS_USERS_002', v_Value);
  End Zltools_Users_002;
End b_Message;
/
------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0060' Where ���=&n_System;
Commit;
