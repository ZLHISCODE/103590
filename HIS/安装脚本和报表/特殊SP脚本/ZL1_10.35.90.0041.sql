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
--135585:������,2018-12-12,���Ŀ����γ����ջ�
CREATE OR REPLACE Procedure Zl_����ҽ����¼_�ջ�
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
  Select ҽ������, Nvl(���id, ID) Into v_ҽ������, n_���id From ����ҽ����¼ Where ID = ҽ��id_In;
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

--124609:��¶¶,2018-12-12,���ͬ��¼�������
--124609:��¶¶,2018-12-13,���ͬ��¼��
CREATE OR REPLACE Procedure Zl_������ϼ�¼_Insert
(
  ����id_In   ������ϼ�¼.����id%Type,
  ��ҳid_In   ������ϼ�¼.��ҳid%Type,
  ��¼��Դ_In ������ϼ�¼.��¼��Դ%Type,
  ����id_In   ������ϼ�¼.����id%Type,
  �������_In ������ϼ�¼.�������%Type,
  ����id_In   ������ϼ�¼.����id%Type,
  ���id_In   ������ϼ�¼.���id%Type,
  ֤��id_In   ������ϼ�¼.֤��id%Type,
  �������_In ������ϼ�¼.�������%Type,
  ��Ժ���_In ������ϼ�¼.��Ժ���%Type,
  �Ƿ�δ��_In ������ϼ�¼.�Ƿ�δ��%Type,
  �Ƿ�����_In ������ϼ�¼.�Ƿ�����%Type,
  ��¼����_In ������ϼ�¼.��¼����%Type,
  ҽ��id_In   varchar2 := Null,
  ��ϴ���_In ������ϼ�¼.��ϴ���%Type := 1,
  ��ע_In     ������ϼ�¼.��ע%Type := Null,
  ��Ժ����_In ������ϼ�¼.��Ժ����%Type := Null,
  ����ʱ��_In ������ϼ�¼.����ʱ��%Type := Null,
  ��¼��_In   ������ϼ�¼.��¼��%Type := Null,
  Id_In       ������ϼ�¼.Id%Type := Null,
  ����id_In   ������ϼ�¼.����id%Type := Null
) Is
  --���ܣ����벡����ϼ�¼
  --ҽ��id_In=�뵱ǰ���������ģ���","�����ҽ��ID��
  v_���id ������ϼ�¼.Id%Type;
  v_ҽ��id ����ҽ����¼.Id%Type;

  v_���˿���id ������Ϣ.��ǰ����id%Type;
  v_����ҽʦ   ��Ա��.����%Type;
  v_����       ��������Ŀ¼.����%Type;
  n_Count      Number;
  n_Mz         Number;
  v_����ʱ��   ������ϼ�¼.����ʱ��%Type;

  v_Temp     varchar2(255);
  v_��Ա���� ��Ա��.����%Type;
  v_Error Varchar2(255); 
  Err_Custom Exception; 
Begin
  --��ǰ������Ա
  If ��¼��_In Is Null Then
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  Else
    v_��Ա���� := ��¼��_In;
  End If;

  If Id_In Is Null Then
    Select ������ϼ�¼_Id.Nextval Into v_���id From Dual;
  Else
    v_���id := Id_In;
  End If;

  v_ҽ��id := Zl_To_Number(ҽ��id_In);
  If v_ҽ��id = 0 Then
    v_ҽ��id := Null;
  End If;

  Select count(*) into n_Count from ������ϼ�¼ where ����id=����id_In And ��ҳid=��ҳid_In And nvl(���id,0)=nvl(���id_In,0) And nvl(����id,0)=nvl(����id_In,0) And nvl(֤��id,0)=nvl(֤��id_In,0) And �������=�������_In And ��¼��Դ=��¼��Դ_In And �������=�������_In And ��ϴ���=��ϴ���_In; 
  Select Count(1) Into n_Mz From ���˹Һż�¼ Where ����id = ����id_In And ID = ��ҳid_In; 
  If n_Count=0 Or (n_Mz > 0 And Instr(',1,11,', �������_In) > 0) then 
	  Insert Into ������ϼ�¼
	    (ID, ����id, ��ҳid, ��¼��Դ, ����id, �������, ��ϴ���, ����id, ���id, ֤��id, �������, ��Ժ����, ��Ժ���, �Ƿ�δ��, �Ƿ�����, 
	         ��¼����, ��¼��, ҽ��id, ��ע, ����ʱ��)
	  Values
	    (v_���id, ����id_In, ��ҳid_In, ��¼��Դ_In, ����id_In, �������_In, ��ϴ���_In, ����id_In, ���id_In, ֤��id_In, �������_In, ��Ժ����_In, ��Ժ���_In,
	     �Ƿ�δ��_In, �Ƿ�����_In, ��¼����_In, v_��Ա����, v_ҽ��id, ��ע_In, ����ʱ��_In);
  Else 
	  v_Error:='�ò����Ѿ�������ͬ��������ݣ����ܱ������ϣ�'; 
	  Raise Err_Custom; 
  End If; 

 If ����id_In Is Not Null Then
    Insert Into ������ϼ�¼
      (Id, ����id, ��ҳid, ��¼��Դ, ����id, �������, ��ϴ���,����id, ���id, ֤��id, �������, ��Ժ����, ��Ժ���, �Ƿ�δ��, �Ƿ�����, 
          ��¼����, ��¼��, ҽ��id, ��ע, ����ʱ��,�������)
      Select ������ϼ�¼_Id.Nextval, ����id_In, ��ҳid_In, ��¼��Դ_In, ����id_In, �������_In, ��ϴ���_In, ����id_In, ���id_In, ֤��id_In, �������_In,
           ��Ժ����_In, ��Ժ���_In, �Ƿ�δ��_In, �Ƿ�����_In, ��¼����_In, v_��Ա����, v_ҽ��id, ��ע_In, ����ʱ��_In,2
      From Dual;
  End If;

  --����������һ�������²��˹Һż�¼.����ʱ��
  v_����ʱ�� := ����ʱ��_In;
  If �������_In = 1 And ��ϴ���_In = 1 Then
    If ����ʱ��_In Is Null Then
      --�����ҽ�ķ���ʱ�䣬����ȡ��ҽ�ģ��������
      Select Max(����ʱ��)
      Into v_����ʱ��
      From ������ϼ�¼
      Where ����id = ����id_In And ��ҳid = ��ҳid_In And ������� = 11 And ��ϴ��� = 1;
    End If;
    If v_����ʱ�� Is Null Then
      --�����ΪNULL����ȡ�Һż�¼�е�
      Select Max(����ʱ��) Into v_����ʱ�� From ���˹Һż�¼ Where ����id = ����id_In And ID = ��ҳid_In;
    End If;
    Update ���˹Һż�¼ Set ����ʱ�� = v_����ʱ�� Where ����id = ����id_In And ID = ��ҳid_In;
  End If;
  If �������_In = 11 And ��ϴ���_In = 1 Then
    --�������ҽ�����ж��Ƿ���д����ҽ�ķ���ʱ�䣬û����д�����޸ģ���������ҽ����ʱ��Ϊ׼
    Select Count(*)
    Into n_Count
    From ������ϼ�¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ������� = 1 And ��ϴ��� = 1 And ����ʱ�� Is Not Null;
    If n_Count = 0 Then
      If v_����ʱ�� Is Null Then
        --�����ΪNULL����ȡ�Һż�¼�е�
        Select Max(����ʱ��) Into v_����ʱ�� From ���˹Һż�¼ Where ����id = ����id_In And ID = ��ҳid_In;
      End If;
      Update ���˹Һż�¼ Set ����ʱ�� = v_����ʱ�� Where ����id = ����id_In And ID = ��ҳid_In;
    End If;
  End If;

  If ҽ��id_In Is Not Null Then
    For r_Advice In (Select Column_Value As ҽ��id
                    From Table(Cast(f_Num2list(ҽ��id_In) As Zltools.t_Numlist)) A, ����ҽ����¼ B
                    Where a.Column_Value = b.Id) Loop 
      Insert Into �������ҽ�� (���id, ҽ��id) Values (v_���id, r_Advice.ҽ��id);
    End Loop;
  End If;

  --�������Ժ��һ��ϣ����ж��Ƿ��ǵ�����
  If �������_In = 2 And ��ϴ���_In = 1 And ��¼��Դ_In = 3 Then
    If ����id_In Is Not Null Then
      Select ���� Into v_���� From ��������Ŀ¼ Where ID = ����id_In;
      Select Max(Upper(����))
      Into v_����
      From ������Ŀ¼
      Where Instr('/' || Replace(Upper(Icd����), ' ', '') || '/', '/' || Upper(v_����) || '/') > 0 And Rownum < 2;
    Else
      v_���� := '';
    End If;
    Update ������ҳ Set ������ = v_���� Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  End If;

  --���ݴ������ҳid_In��ѯ�Һż�¼��������������ҳ����סԺ��ҳ����
  Begin
    Select ִ����, ִ�в���id Into v_���˿���id, v_����ҽʦ From ���˹Һż�¼ Where ID = ��ҳid_In;
    Zl_���Ӳ���ʱ��_Insert(����id_In, ��ҳid_In, 1, '���', v_���˿���id, v_����ҽʦ, ��¼����_In, ��¼����_In);
  Exception
    When Others Then
      Null;
  End;
  If v_���˿���id Is Null And (�������_In <> 1 Or �������_In <> 11) Then
    Begin
      Select ��Ժ����id, סԺҽʦ
      Into v_���˿���id, v_����ҽʦ
      From ������ҳ
      Where ����id = ����id_In And ��ҳid = ��ҳid_In;
      Zl_���Ӳ���ʱ��_Insert(����id_In, ��ҳid_In, 2, '���', v_���˿���id, v_����ҽʦ, ��¼����_In, ��¼����_In);
    Exception
      When Others Then
        Null;
    End;
  End If;
  b_Message.Zlhis_Cis_010(����id_In, ��ҳid_In, v_���id);
Exception
  When Err_Custom Then 
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������ϼ�¼_Insert;
/





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0041' Where ���=&n_System;
Commit;
