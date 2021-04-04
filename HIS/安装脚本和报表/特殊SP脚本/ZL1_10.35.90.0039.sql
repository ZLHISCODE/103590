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
--133839:����,2018-11-26,����ģ�鹫������ҽ��վ�Һ��������,���ڿ���ҽ��վ�Һ�ʱ��ʾ�Һ��Ű��˳��
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 9000, 0, 0, 0, 0, 0, 0, 18, 'ҽ��վ�Һ��������', Null, 'ҽ��,1|ִ��ʱ��,1|����,1|�ű�,1|��Ŀ,1',
         '��Ҫ�����ҽ��վ�Һ�ʱ��Դ������˳�򣬸�����˳�򣬲���ԡ���ʾ���кű𡱡�', '�����ֶ�1 ������ʽ(0-DESC��1-ASC)|�����ֶ�2 ������ʽ(0-DESC��1-ASC)|...', '',
         '��������Ҫ�������������ʾ�Һ��Ű�˳������', Null
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
--133584:���ϴ�,2018-11-30,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_�ҺŰ��żƻ�_Verify
(
  Id_In         In �ҺŰ��żƻ�.Id%Type,
  ������Ч_In   Number := 0,
  �ϴμƻ�ID_In In �ҺŰ��żƻ�.�ϴμƻ�Id%Type := Null
) Is
  Err_Item     Exception;
  v_Err_Msg    Varchar2(100);
  v_User_Name  ��Ա��.����%Type;
  n_Valied     Number(1);
  d_��Чʱ��   �ҺŰ��żƻ�.��Чʱ��%Type;
  n_�ϴμƻ�ID �ҺŰ��żƻ�.ID%Type;
Begin
  
  Select Nvl(Max(p.����),'') Into v_User_Name From �ϻ���Ա�� o, ��Ա�� p Where o.��Աid = p.Id And �û��� = User;
  If v_User_Name Is Null Then
    v_Err_Msg := '[ZLSOFT]��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����' || Chr(10) || Chr(13) ||
                 'ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�[ZLSOFT]';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Valied From �ҺŰ��żƻ� a Where Nvl(��Чʱ��, Sysdate) < Sysdate And a.Id = Id_In And Rownum < 2;
  If n_Valied = 1 And Nvl(������Ч_In, 0) = 0 Then
    v_Err_Msg := '[ZLSOFT]�üƻ����ŵ���Чʱ���Ѿ����ڣ����ܽ�����ˣ�[ZLSOFT]';
    Raise Err_Item;
  End If;
  
  if Nvl(�ϴμƻ�ID_In, 0) = 0 Then
    Select Max(Id)
    Into n_�ϴμƻ�id
    From (Select Max(Id) As Id, Max(ʧЧʱ��) As ʧЧʱ��, Count(1) As Count
           From �ҺŰ��żƻ�
           Where ����id = (Select Max(����id) From �ҺŰ��żƻ� Where Id = Id_In) And ���ʱ�� Is Not Null)
    Where Count = 1 And ʧЧʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd');
  Else
    n_�ϴμƻ�ID := �ϴμƻ�ID_In;
  End if;
  
  Update �ҺŰ��żƻ�
  Set ����� = v_User_Name, ���ʱ�� = Sysdate, �ϴμƻ�ID = n_�ϴμƻ�ID,
      ��Чʱ�� = Case Nvl(������Ч_In, 0) When 0 Then ��Чʱ�� Else Sysdate - 1 / 24 / 60 / 60 End
  Where Id = Id_In And ���ʱ�� Is Null
  Return ��Чʱ�� Into d_��Чʱ��;
  If Sql%Notfound Then
    v_Err_Msg := '[ZLSOFT]�üƻ������Ѿ���������˻�ɾ��,���������![ZLSOFT]';
    Raise Err_Item;
  End If;
  IF Nvl(n_�ϴμƻ�ID, 0) <> 0 Then
    Update �ҺŰ��żƻ� Set ʧЧʱ�� = d_��Чʱ�� Where ID = n_�ϴμƻ�ID;
  End IF;
  If Nvl(������Ч_In, 0) = 1 Then
    Begin
      Zl_�ҺŰ���_Autoupdate();
    Exception
      When Others Then
        Null;
    End;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�ҺŰ��żƻ�_Verify;
/

--128110:������,2018-11-30,�������������ջ��Զ�ִ��
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
    ҩƷid_In     ҩƷ�շ���¼.ҩƷid%Type,
    �ⷿid_In     ҩƷ�շ���¼.�ⷿid%Type,
    ����_In       ҩƷ�շ���¼.����%Type,
    ����_In       ������Ϣ.����%Type,
    �Է�����id_In ҩƷ�շ���¼.�Է�����id%Type,
    �շ����_In   סԺ���ü�¼.�շ����%Type,
    �������_In   Varchar
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
      Select v_Lngid, 1, ����, No_In, v_�շ����, �ⷿid, �Է�����id, ������id, -1, ҩƷid, Nvl(v_����, 0), ����, v_����, v_Ч��, v_��ǰ����,
             -1 * v_��ǰ����, -1 * v_��ǰ����, ���ۼ�, Round(-1 * v_��ǰ���� * v_��ǰ���� * ���ۼ�, v_Dec), '���ڷ����ջ�', v_��Ա����, �ջ�ʱ��_In, ����id_In,
             ����, Ƶ��, �÷�, ��ҩ��λid, ��������, ��׼�ĺ�, ���Ч��
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
                                r_Drug.����id, r_Drug.��ҳid, r_Drug.ҩƷid, r_Drug.�ⷿid, r_Drug.����, r_Drug.����, r_Drug.�Է�����id,
                                v_�������, v_�������);
                
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
                  
                    --סԺ���ü�¼:��������ջ����������ϴη�����,����ȷ
                    Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
                    If r_Other.�շ���� In ('4', '5', '6', '7') Then
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
                        �����շ���¼_Insert(v_����id, r_Otherdrug.����, r_Otherdrug.����, r_Otherdrug.����, r_Otherdrug.Ч��,
                                      r_Otherdrug.���Ч��, r_Otherdrug.�շ�id, r_Otherdrug.����id, r_Otherdrug.��ҳid,
                                      r_Otherdrug.ҩƷid, r_Otherdrug.�ⷿid, r_Otherdrug.����, r_Otherdrug.����,
                                      r_Otherdrug.�Է�����id, r_Other.�շ����, v_�������);
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
--000000:������,2018-11-29,����ê����Ϣ����
Create Or Replace Procedure Zl_����ҽ����¼_����
(
  ҽ��id_In     In ����ҽ����¼.Id%Type,
  Flag_In       In Number := 0,
  ҽ������_In   In ����ҽ����¼.ҽ������%Type := Null,
  ��������_In   In ����ҽ��״̬.��������%Type := Null,
  ����Ա���_In In ��Ա��.���%Type := Null,
  ����Ա����_In In ��Ա��.����%Type := Null
  --���ܣ�����סԺҽ����״̬�������Ͳ���(������������ͨ������Zl_����ҽ����¼_��������������) 
  --������ҽ��ID_IN=һ��ҽ��ID 
  --      FLAG_IN=�������ݡ�����ֹͣ��0=���ִ����ֹʱ��,1=�������е�ִ����ֹʱ�䡣 
  --      ҽ������_IN=�ù��̱��������˵���ʱ���ã����ڴ�����ʾ�� 
  --      ��������_IN=�ù��̱��������˵���ʱ���ã����ں˶Ի������ݡ�0-���˷���,n=���˾���ҽ������ 
) Is
  --����ָ��ҽ���Ĳ�����¼,��һ��ΪҪ���˵�����(״̬��������) 
  --���������˷��ͺ���Զ�ֹͣ,�ڻ��˷���ʱ�Զ�����ֹͣ���� 
  Cursor c_Rolladvice Is
    Select b.������Ա, b.����ʱ��, 0 As ���ͺ�, a.���, Null As NO, b.��������, 0 As ִ��״̬, Sysdate + Null As �״�ʱ��, Sysdate + Null As ĩ��ʱ��,
           a.�ϴ�ִ��ʱ��, a.ҽ����Ч, a.������� As ���, a.������Ŀid, Null As ����, a.����id, a.��ҳid, a.Ӥ��, 0 As ��¼����, 0 As �������, 0 As ��������id,
           a.��˱��, a.����ҽ��, a.ִ�п���id, Nvl(a.���id, a.Id) As ��id, a.���id, a.Id As ҽ��id, -null As ��������, Null As ��������
    From ����ҽ����¼ A, ����ҽ��״̬ B
    Where a.Id = b.ҽ��id And (a.Id = ҽ��id_In Or a.���id = ҽ��id_In) And
          (Nvl(a.ҽ����Ч, 0) = 0 And b.�������� Not In (1, 2, 3) Or Nvl(a.ҽ����Ч, 0) = 1 And b.�������� Not In (1, 2, 3, 8))
    Union
    Select b.������ As ������Ա, b.����ʱ�� As ����ʱ��, b.���ͺ�, a.���, b.No, -null As ��������, b.ִ��״̬, b.�״�ʱ��, b.ĩ��ʱ��, a.�ϴ�ִ��ʱ��, a.ҽ����Ч,
           c.���, a.������Ŀid, c.�������� As ����, a.����id, a.��ҳid, a.Ӥ��, b.��¼����, b.�������, a.��������id, a.��˱��, a.����ҽ��, a.ִ�п���id,
           Nvl(a.���id, a.Id) As ��id, a.���id, a.Id As ҽ��id, b.��������, b.��������
    From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C
    Where a.Id = b.ҽ��id And a.������Ŀid = c.Id And (a.Id = ҽ��id_In Or a.���id = ҽ��id_In)
    Order By ����ʱ�� Desc, ���ͺ�, ���;
  r_Rolladvice c_Rolladvice%RowType;

  --��ʽͬc_Rolladvice��ֻȡ���Ͳ��������Զ����˴��� 
  Cursor c_Rollsend(v_���ͺ� ����ҽ������.���ͺ�%Type) Is
    Select Distinct b.ҽ��id, b.����ʱ�� As ����ʱ��, b.���ͺ�, b.ִ��״̬, a.������� As ���, c.��ǰ����id As ���˲���id, a.���˿���id,
                    b.ִ�в���id As ִ�п���id
    From ����ҽ����¼ A, ����ҽ������ B, ������ҳ C
    Where a.Id = b.ҽ��id And b.���ͺ� = v_���ͺ� And a.����id = c.����id And a.��ҳid = c.��ҳid And
          (a.Id = ҽ��id_In Or a.���id = ҽ��id_In)
    Order By b.����ʱ�� Desc, b.���ͺ�;

  --����ҽ��������NO������λ���Ҫ���ʵķ��ü�¼ 
  --һ��ҽ�������Ƕ���д�˷��ͼ�¼,�ҿ���NO��ͬ(ҩƷ��,�÷��巨��һ����) 
  --���ܷ��ͼ�¼�ļƷ�״̬(��������Ʒ�),�з��ü�¼��Ȼ�������� 
  --����ֻ��۸񸸺�Ϊ�յ�,�Ա�ȡ������� 
  --ֻ�ܼ�¼״̬Ϊ1�ķ���,���������ʻ򲿷����ʵļ�¼,���ٴ�������"��¼״̬=3"�Ķ�ȡ�����������жϣ������� 
  Cursor c_Rollmoneyout
  (
    v_���ͺ�    ����ҽ������.���ͺ�%Type,
    v_ҽ��id    ����ҽ����¼.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.��¼״̬, a.No, a.���, a.�շ����, a.ִ��״̬, d.��������, a.ִ�в���id, a.��¼����
    From ������ü�¼ A, Table(t_Adviceids) B, ����ҽ������ C, �������� D
    Where c.ҽ��id = b.Column_Value And c.���ͺ� = v_���ͺ� And a.ҽ����� = b.Column_Value And
          (a.ҽ����� = v_ҽ��id Or Nvl(v_ҽ��id, 0) = 0) And a.��¼״̬ In (0, 1, 3) And a.No = c.No And a.��¼���� = c.��¼���� And
          a.�۸񸸺� Is Null And a.�շ�ϸĿid = d.����id(+)
    Order By a.No, a.���;

  Cursor c_Rollmoneyin
  (
    v_���ͺ�    ����ҽ������.���ͺ�%Type,
    v_ҽ��id    ����ҽ����¼.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.��¼״̬, a.No, a.���, a.�շ����, a.ִ��״̬, d.��������, a.ִ�в���id, a.��¼����
    From סԺ���ü�¼ A, Table(t_Adviceids) B, ����ҽ������ C, �������� D
    Where c.ҽ��id = b.Column_Value And c.���ͺ� = v_���ͺ� And a.ҽ����� = b.Column_Value And
          (a.ҽ����� = v_ҽ��id Or Nvl(v_ҽ��id, 0) = 0) And a.��¼״̬ In (0, 1, 3) And a.No = c.No And a.��¼���� = c.��¼���� And
          a.�۸񸸺� Is Null And a.�շ�ϸĿid = d.����id(+)
    Order By a.No, a.���;

  --ȡ����סԺ����ʱ�Զ����ŵ�����(��û�����ϵ�) 
  Cursor c_Stuff_Drug(v_����id ҩƷ�շ���¼.����id%Type) Is
    Select ID
    From ҩƷ�շ���¼
    Where ����id = v_����id And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0) And ����� Is Not Null
    Order By ҩƷid;

  --���ڴ�������ҽ���Ļ��� 
  Cursor c_Patilog
  (
    v_����id ���˱䶯��¼.����id%Type,
    v_��ҳid ���˱䶯��¼.��ҳid%Type
  ) Is
    Select *
    From ���˱䶯��¼
    Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� Is Null
    Order By ��ʼʱ�� Desc;
  r_Patilog c_Patilog%RowType;

  Cursor c_Adviceids Is
    Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In;
  t_Adviceids t_Numlist;

  v_ҽ��״̬     ����ҽ����¼.ҽ��״̬%Type;
  v_ҽ����Ч     ����ҽ����¼.ҽ����Ч%Type;
  v_����no       ����ҽ������.No%Type;
  v_�������     Varchar2(255);
  v_ĩ��ʱ��     ����ҽ������.ĩ��ʱ��%Type;
  v_����ʱ��     ����ҽ��״̬.����ʱ��%Type;
  v_��������     ������ĿĿ¼.��������%Type;
  v_ִ��Ƶ��     ������ĿĿ¼.ִ��Ƶ��%Type;
  v_�ϴ�ʱ��     ����ҽ����¼.�ϴ�ִ��ʱ��%Type;
  v_ִ��ʱ��     ����ҽ����¼.ִ��ʱ�䷽��%Type;
  v_��ʼִ��ʱ�� ����ҽ����¼.��ʼִ��ʱ��%Type;
  v_�ϴδ�ӡʱ�� ����ҽ����¼.�ϴδ�ӡʱ��%Type;
  v_Ƶ�ʼ��     ����ҽ����¼.Ƶ�ʼ��%Type;
  v_�����λ     ����ҽ����¼.�����λ%Type;
  v_���ͺ�       ����ҽ������.���ͺ�%Type;
  n_����ȼ�id   ���˱䶯��¼.����ȼ�id%Type;
  d_��ʼʱ��     ���˱䶯��¼.��ʼʱ��%Type;
  d_����ʱ��     ����ҽ��״̬.����ʱ��%Type;
  v_Tmp���ͺ�    ����ҽ������.���ͺ�%Type;
  n_ִ��         Number;

  Intdigit   Number(3);
  v_Update   Number(1);
  v_Count    Number(5);
  v_Temp     Varchar2(2000);
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;
  v_Time     Varchar2(4000);
  n_Blndo    Number;

  v_Error Varchar2(2000);
  Err_Custom Exception;

  Function Checkmoneyundo
  (
    v_No       סԺ���ü�¼.No%Type,
    v_��¼���� סԺ���ü�¼.��¼����%Type,
    v_���     סԺ���ü�¼.���%Type,
    n_����     Number := 0 --0סԺ��1���� 
  ) Return Number Is
    n_Num      Number;
    n_ִ��״̬ Number;
  Begin
    n_Num := 0;
    If n_���� = 0 Then
      Select Nvl(Sum(Nvl(����, 1) * ����), 0) As ����
      Into n_Num
      From סԺ���ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ In (2, 3);
      Select Nvl(ִ��״̬, 0)
      Into n_ִ��״̬
      From סԺ���ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ = 3;
    Else
      Select Nvl(Sum(Nvl(����, 1) * ����), 0) As ����
      Into n_Num
      From ������ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ In (2, 3);
      Select Nvl(ִ��״̬, 0)
      Into n_ִ��״̬
      From ������ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ = 3;
    End If;
    If n_Num <> 0 Then
      n_Num := 1;
    End If;
    --�������¼����ִ�У�����ִ�еģ����Զ��ˡ� 
    If n_ִ��״̬ <> 0 Then
      n_Num := 0;
    End If;
    Return(n_Num);
  End;
Begin
  v_Tmp���ͺ� := -1;
  Open c_Rolladvice;
  Loop
    Fetch c_Rolladvice
      Into r_Rolladvice;
    If c_Rolladvice%RowCount = 0 Then
      Close c_Rolladvice;
      v_Error := Nvl(ҽ������_In, '��ҽ��') || '��ǰû�п��Ի��˵����ݡ�';
      Raise Err_Custom;
    End If;
    Exit When c_Rolladvice%NotFound;
    Exit When d_����ʱ�� <> r_Rolladvice.����ʱ�� And d_����ʱ�� Is Not Null;
    d_����ʱ�� := r_Rolladvice.����ʱ��;
  
    --�������˵���ʱ�ж� 
    If ҽ������_In Is Not Null Then
      If Nvl(r_Rolladvice.��������, 0) <> Nvl(��������_In, 0) Then
        v_Error := Nvl(ҽ������_In, '��ҽ��') || '�����뵱ǰҽ��һ����ˣ����ܸ�ҽ���Ѿ�ִ��������������';
        Raise Err_Custom;
      End If;
    End If;
  
    --һ�鷢�ͺ�ִֻ��һ�� 
    If v_Tmp���ͺ� <> r_Rolladvice.���ͺ� Then
      v_Tmp���ͺ� := r_Rolladvice.���ͺ�;
      n_ִ��      := 1;
    Else
      n_ִ�� := 0;
    End If;
  
    If n_ִ�� = 1 Then
      Open c_Adviceids;
      Fetch c_Adviceids Bulk Collect
        Into t_Adviceids;
      Close c_Adviceids;
    
      If r_Rolladvice.���ͺ� = 0 Then
        --����ҽ��״̬����(��ʱ��ؼ���) 
        --4-���ϣ�5-������6-��ͣ��7-���ã�8-ֹͣ��9-ȷ��ֹͣ��10-Ƥ�Խ��;13-ͣ������ 
        ------------------------------------------------------------------ 
        --���ֻ���˻ص�У��״̬ 
        If r_Rolladvice.�������� = 3 Then
          v_Error := Nvl(ҽ������_In, '��ҽ��') || '��ǰ����ͨ��У��״̬�������ٻ��ˡ�';
          Raise Err_Custom;
        Elsif r_Rolladvice.�������� = 4 And Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
          If r_Rolladvice.��� = 'H' Then
            Select ��������, ִ��Ƶ�� Into v_��������, v_ִ��Ƶ�� From ������ĿĿ¼ Where ID = r_Rolladvice.������Ŀid;
            If v_�������� = '1' And v_ִ��Ƶ�� = '2' Then
              v_Error := '����ȼ����Ϻ����ٻ��ˡ�';
              Raise Err_Custom;
            End If;
          End If;
        End If;
      
        --����Ƿ�������������֮ǰ�Ĳ��� 
        If r_Rolladvice.�������� <> 5 Then
          --ȡ�������ʱ�� 
          Select Nvl(ҽ������ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD'))
          Into v_����ʱ��
          From ������ҳ
          Where ����id = r_Rolladvice.����id And ��ҳid = r_Rolladvice.��ҳid;
        
          If r_Rolladvice.����ʱ�� < v_����ʱ�� Then
            v_Error := '�ò������������֮ǰ�Ĳ��������ٻ��ˡ�';
            Raise Err_Custom;
          End If;
        End If;
      
        --ɾ��(����ҽ��)�����״̬������¼ 
        Delete /*+ Rule*/
        From ����ҽ��״̬
        Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And ����ʱ�� = r_Rolladvice.����ʱ��;
      
        --ȡɾ����Ӧ�ָ���ҽ��״̬ 
        Select ��������
        Into v_ҽ��״̬
        From ����ҽ��״̬
        Where ����ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = ҽ��id_In) And ҽ��id = ҽ��id_In;
      
        --�ָ�(����ҽ��)���˺��״̬ 
        Update ����ҽ����¼ Set ҽ��״̬ = v_ҽ��״̬ Where ID = ҽ��id_In Or ���id = ҽ��id_In;
      
        --��������Ĵ��� 
        If r_Rolladvice.�������� = 8 Then
          --�����ڷ����ջع���ҽ�� ���������������ģʽ�����ж϶�Ӧ�ġ����˷������ʡ������Ƿ�ȡ��������������ˣ��������� 
          --                       ����ǲ�����������ģʽ���������ٻ��ˡ� 
          --���ܳ��ڷ����ջ�ʱ��ȫ���ջ�(���ϴ�ִ��ʱ��) 
          Select /*+ Rule*/
           Nvl(Count(*), 0)
          Into v_Count
          From ����ҽ����¼ A, ����ҽ������ B
          Where b.ҽ��id = a.Id And (a.Id = ҽ��id_In Or a.���id = ҽ��id_In) And
                b.���ͺ� =
                (Select Max(���ͺ�) From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(t_Adviceids))) And
                a.ִ����ֹʱ�� Is Not Null And ((a.�ϴ�ִ��ʱ�� < b.ĩ��ʱ��) Or (a.�ϴ�ִ��ʱ�� Is Null And b.ĩ��ʱ�� Is Not Null));
          If v_Count > 0 Then
            If zl_GetSysParameter('�����ջز�����������', 1254) = '1' Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '�ѱ����ڷ����ջأ������ٳ���ֹͣ������';
              Raise Err_Custom;
            Else
              --����Ѿ�ȡ���������룬���������. 
              Select Count(1)
              Into v_Count
              From ���˷������� A, סԺ���ü�¼ B, ����ҽ����¼ C
              Where a.����id = b.Id And c.Id = b.ҽ����� And (c.Id = ҽ��id_In Or c.���id = ҽ��id_In);
              If v_Count > 0 Then
                v_Error := Nvl(ҽ������_In, '��ҽ��') || '�ѱ����ڷ����ջأ������ٳ���ֹͣ������';
                Raise Err_Custom;
              Else
                --�õ��ϴ�ִ��ʱ�����Ϣ 
                Select �ϴ�ִ��ʱ��, ִ��ʱ�䷽��, ��ʼִ��ʱ��, �ϴδ�ӡʱ��, Ƶ�ʼ��, �����λ
                Into v_�ϴ�ʱ��, v_ִ��ʱ��, v_��ʼִ��ʱ��, v_�ϴδ�ӡʱ��, v_Ƶ�ʼ��, v_�����λ
                From ����ҽ����¼
                Where ID = ҽ��id_In;
                v_�ϴ�ʱ�� := To_Date(To_Char(v_�ϴ�ʱ�� + 1 / 24 / 60 / 60, 'yyyy-MM-dd hh24:mi:ss'), 'yyyy-MM-dd hh24:mi:ss');
              
                --�޸��ϴ�ִ��ʱ��Ϊ�ջغ��ĩ��ִ��ʱ�䡣 
                v_ĩ��ʱ�� := Null;
                Begin
                  --һ��ҽ���ķ�����ĩʱ����ͬ,һ����ҩ��ȡ��С�� 
                  --ȡ���IDΪNULL��ҽ���ķ��ͼ�¼��ʱ�� 
                  --����ҩ;������ҩ�÷�����δ��д���ͼ�¼ 
                  Select /*+ Rule*/
                   ĩ��ʱ��, ���ͺ�
                  Into v_ĩ��ʱ��, v_���ͺ�
                  From ����ҽ������
                  Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And
                        ���ͺ� = (Select Max(���ͺ�)
                               From ����ҽ������
                               Where ҽ��id In (Select Column_Value From Table(t_Adviceids))) And Rownum = 1;
                Exception
                  When Others Then
                    Null;
                End;
                Update ����ҽ����¼ Set �ϴ�ִ��ʱ�� = v_ĩ��ʱ�� Where ID = ҽ��id_In Or ���id = ҽ��id_In;
              
                Select Count(1) Into v_Count From ����ҽ������ Where ҽ��id = ҽ��id_In And ���ͺ� = v_���ͺ�;
                If v_Count > 0 Then
                  --��ԭҽ��ִ��ʱ�� 
                  Select Zl_Adviceexetimes(ҽ��id_In, v_�ϴ�ʱ��, v_ĩ��ʱ��, v_ִ��ʱ��, v_��ʼִ��ʱ��, v_�ϴδ�ӡʱ��, v_Ƶ�ʼ��, v_�����λ, 0)
                  Into v_Time
                  From Dual;
                  Insert Into ҽ��ִ��ʱ��
                    (Ҫ��ʱ��, ҽ��id, ���ͺ�)
                    Select To_Date(Column_Value, 'yyyy-mm-dd hh24:mi:ss'), ҽ��id_In, v_���ͺ�
                    From Table(f_Str2list(v_Time));
                End If;
              End If;
            End If;
          End If;
        
          --����ȼ��䶯�������������䶯ʱ����������� 
          If r_Rolladvice.��� = 'H' And Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
            Select ��������, ִ��Ƶ�� Into v_��������, v_ִ��Ƶ�� From ������ĿĿ¼ Where ID = r_Rolladvice.������Ŀid;
            If v_�������� = '1' And v_ִ��Ƶ�� = '2' Then
              Select Count(*), Max(a.����ȼ�id), Max(a.��ʼʱ��)
              Into v_Count, n_����ȼ�id, d_��ʼʱ��
              From ���˱䶯��¼ A
              Where a.����id = r_Rolladvice.����id And a.��ҳid = r_Rolladvice.��ҳid And a.��ʼԭ�� = 6 And a.��ֹʱ�� Is Null And
                    a.���Ӵ�λ = 0;
              --���û���ҵ����һ���ǻ���ȼ��䶯���ֹ 
              If v_Count = 0 Then
                --ҽ������ȼ�����סʱ��Ļ���ȼ�һ��ʱҪ�����ж� 
                Select Count(*)
                Into v_Count
                From ���˱䶯��¼ A
                Where a.����id = r_Rolladvice.����id And a.��ҳid = r_Rolladvice.��ҳid And a.��ʼԭ�� = 6;
                If v_Count > 0 Then
                  v_Error := '���ڻ���ȼ�ҽ��ֹͣ��ò����Ѿ������������䶯��¼,���ܻ��˸�ҽ����ֹͣ������';
                  Raise Err_Custom;
                End If;
              Else
                --���n_����ȼ�IDΪNull�������Ƿ��ǵ�ǰ���˵�ҽ����Ӧ�ı䶯��¼,Ŀ�����ж������ȼ�ҽ��ʱҪ��˳����ˡ� 
                --���n_����ȼ�ID��ΪNull�����п�����У����һ������ȼ�ʱ���Զ�ֹͣ�ģ�δ�����䶯��¼�� 
                --     ����Ҫ��鵱ǰ���һ���䶯�Ļ���ȼ�ID�Ƿ��ǵ�ǰҽ���Ļ���ȼ�ID,Ŀ�����ж������ȼ�ҽ��ʱҪ��˳����ˣ����������Ҫ�ٳ������һ�α䶯��ֱ�ӻ���ҽ�����ɡ� 
                If n_����ȼ�id Is Null Then
                  Select Count(*)
                  Into v_Count
                  From ���˱䶯��¼ B, ����ҽ���Ƽ� C
                  Where b.����id = r_Rolladvice.����id And b.��ҳid = r_Rolladvice.��ҳid And c.ҽ��id = ҽ��id_In And
                        c.�շ�ϸĿid = b.����ȼ�id And b.��ֹʱ�� = d_��ʼʱ�� And b.��ֹԭ�� = 6 And b.���Ӵ�λ = 0;
                Else
                  --��ʼʱ��ֻȡ���ӶԱȣ�У�Ե�ʱ����ȼ��Ŀ�ʼʱ����ҽ����ʼʱ��+��ǰʱ������� 
                  Select Count(*)
                  Into v_Count
                  From ����ҽ���Ƽ� C, ����ҽ����¼ A
                  Where a.Id = c.ҽ��id And a.Id = ҽ��id_In And c.�շ�ϸĿid = n_����ȼ�id And
                        a.��ʼִ��ʱ�� = To_Date(To_Char(d_��ʼʱ��, 'yyyy-mm-dd hh24:mi'), 'yyyy-mm-dd hh24:mi');
                End If;
                If v_Count = 0 Then
                  v_Error := '�����˵�ҽ���������һ������ȼ�ҽ�����뽫����Ļ���ȼ�ҽ�����Ϻ��ٻ��˱���ҽ����';
                  Raise Err_Custom;
                End If;
              
                If n_����ȼ�id Is Null Then
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
                
                  Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, '1', Null, Null, '����ȼ��䶯');
                End If;
              End If;
            End If;
          End If;
        
          If r_Rolladvice.��� = 'Z' And Instr(',9,10,', ',' || r_Rolladvice.���� || ',') > 0 And
             Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
            --���˲���ҽ��ʱ�����ñ䶯��¼���� 
            Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, Null, Null, Null, '�����䶯');
          End If;
        
          --����ҽ��ֹͣʱ,���ͣ��ҽ����ʱ��,�����ʵϰҽʦ�������˵ģ���ָ������״̬ 
          Update ����ҽ����¼
          Set ִ����ֹʱ�� = Decode(Flag_In, 1, ִ����ֹʱ��, Null), ͣ��ҽ�� = Null, ͣ��ʱ�� = Null,
              ��˱�� = Decode(r_Rolladvice.��˱��, 3, 2, r_Rolladvice.��˱��)
          Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        Elsif r_Rolladvice.�������� = 9 Then
          --����ҽ��ȷ��ֹͣʱ,����Ƿ��Ѵ�ӡͣ��ʱ�� 
          Select /*+ Rule*/
           Count(*)
          Into v_Count
          From ����ҽ����ӡ
          Where ��ӡ��� = 1 And ҽ��id In (Select Column_Value From Table(t_Adviceids));
          If v_Count > 0 Then
            v_Error := Nvl(ҽ������_In, '��ҽ��') || '��ͣ��ʱ���Ѿ���ӡ�������ٳ���ȷ��ֹͣ������';
            Raise Err_Custom;
          End If;
        
          --����ҽ��ȷ��ֹͣʱ,���ͣ��ҽ����ʱ�� 
          Update ����ҽ����¼ Set ȷ��ͣ��ʱ�� = Null, ȷ��ͣ����ʿ = Null Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        Elsif r_Rolladvice.�������� = 10 Then
          --���˱�עƤ�Խ��,ͬʱɾ�������Ǽ�(+)��(-),���ݼ�¼ʱ�� 
          --�����ļ�¼��ҽ�������޹ۣ�����Ҫ���� 
          Delete From ���˹�����¼
          Where ����id = r_Rolladvice.����id And Nvl(��ҳid, 0) = Nvl(r_Rolladvice.��ҳid, 0) And ��¼ʱ�� = r_Rolladvice.����ʱ�� And
                Nvl(���, 0) = 0;
        
          Update ����ҽ����¼ Set Ƥ�Խ�� = Null Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        Elsif r_Rolladvice.�������� = 13 Then
          If Instr(r_Rolladvice.����ҽ��, '/') > 0 Then
            Update ����ҽ����¼ Set ��˱�� = 1 Where ID = ҽ��id_In Or ���id = ҽ��id_In;
          Else
            Update ����ҽ����¼ Set ��˱�� = Null Where ID = ҽ��id_In Or ���id = ҽ��id_In;
          End If;
        End If;
        --��������ҽ�����ϲ��� 
        --����ҽ�� 
        If r_Rolladvice.�������� = 4 And r_Rolladvice.��� = 'Z' Then
          Select Count(1) Into v_Count From ������ĿĿ¼ Where ID = r_Rolladvice.������Ŀid And �������� = '4';
          If v_Count = 1 Then
            b_Message.Zlhis_Cis_004(r_Rolladvice.����id, r_Rolladvice.��ҳid, ҽ��id_In);
          End If;
        End If;
      Else
        --����ҽ������(�Է��ͺŹؼ���) 
        ------------------------------------------------------------------ 
        --��ǰ������Ա 
        v_Temp     := Zl_Identity;
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      
        --����Ƿ�����Һ��Һ��¼�����Ƿ��Ѿ������������ѯ������˵������Һ��¼ 
        Begin
          Select Decode(Max(�Ƿ�����), 1, 1, 0)
          Into v_Count
          From ��Һ��ҩ��¼
          Where ҽ��id = ҽ��id_In And ���ͺ� = r_Rolladvice.���ͺ�;
        Exception
          When Others Then
            v_Count := -1;
        End;
      
        If v_Count = 1 Then
          v_Error := 'ҽ��"' || ҽ������_In || '"����ҺҩƷ���Ѿ�����Һ�����������������ܻ��˷��͡�';
          Raise Err_Custom;
        Elsif v_Count = 0 Then
          Zl_��Һ��ҩ��¼_ҽ������(ҽ��id_In, r_Rolladvice.���ͺ�, v_��Ա����, Sysdate);
        End If;
      
        --����Ƿ����δ��˵��������� 
        Select Count(*)
        Into v_Count
        From ����ҽ����¼ A, ����ҽ������ B, סԺ���ü�¼ C, ���˷������� D
        Where (a.Id = ҽ��id_In Or a.���id = ҽ��id_In) And a.Id = b.ҽ��id And b.ҽ��id = c.ҽ����� And c.Id = d.����id And
              c.��¼״̬ In (0, 1, 3) And d.״̬ = 0;
      
        If v_Count > 0 Then
          v_Error := 'ҽ��"' || ҽ������_In || '"����δ��˵��������룬��ȡ�����������������ٻ��˷��͡�';
          Raise Err_Custom;
        End If;
      
        --���ҽ���Ƿ������Ч��ҽ������ 
        Select Count(*)
        Into v_Count
        From ����ҽ������ A, סԺ���ü�¼ B
        Where a.ҽ��id = b.ҽ����� And a.No = b.No And b.��¼״̬ = 1 And b.ʵ�ս�� <> 0 And a.���ͺ� = r_Rolladvice.���ͺ� And
              a.ҽ��id In (Select Column_Value From Table(t_Adviceids));
        If v_Count > 0 Then
          v_Error := '��ҽ���»����ڸ�����Ŀ�����ȳ�����';
          Raise Err_Custom;
        End If;
      
        --���Ʒ����Զ�ִ��ʱ������Ҳ�Զ�����ִ��(����ʿվ�д˹���) 
        --�Ǹ������õ�����ҽ����ͬ��ͨҽ��ִ�д��� 
        Select ҽ����Ч Into v_ҽ����Ч From ����ҽ����¼ Where ID = ҽ��id_In;
        If Substr(zl_GetSysParameter('����ִ���Զ����', 1254), v_ҽ����Ч + 1, 1) = '1' Then
          Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into Intdigit From Dual;
        
          For r_Rollsend In c_Rollsend(r_Rolladvice.���ͺ�) Loop
            If Nvl(r_Rollsend.ִ��״̬, 0) = 1 And
               (Nvl(r_Rollsend.ִ�п���id, 0) = Nvl(r_Rollsend.���˲���id, 0) Or
                Nvl(r_Rollsend.ִ�п���id, 0) = Nvl(r_Rollsend.���˿���id, 0)) Then
            
              --ҽ����ִ��״̬ 
              Update ����ҽ������ Set ִ��״̬ = 0 Where ���ͺ� = r_Rollsend.���ͺ� And ҽ��id = r_Rollsend.ҽ��id;
              v_Update := 1;
            
              If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 0 Then
                --���õ�ִ��״̬ 
                For r_Rollmoney In c_Rollmoneyin(r_Rollsend.���ͺ�, r_Rollsend.ҽ��id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.��¼״̬ <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1) And
                       Not r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --��ͨ����ֱ��ȡ��ִ��״̬������ҩƷ�͸������õ����� 
                      Update סԺ���ü�¼
                      Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
                      Where NO = r_Rollmoney.No And ��¼���� = r_Rolladvice.��¼���� And ��¼״̬ = r_Rollmoney.��¼״̬ And
                            Nvl(�۸񸸺�, ���) = r_Rollmoney.��� And ҽ����� = r_Rollsend.ҽ��id;
                    Elsif r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1 Then
                      --�������õ����ģ����Զ����� 
                      For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                        Zl_�����շ���¼_��������(r_Stuff.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, 0, v_��Ա����);
                      End Loop;
                    Elsif r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --סԺ���ҷ�ҩ��ҩƷ�Զ���ҩ 
                      If r_Rollmoney.ִ�в���id = r_Rollsend.���˲���id Or r_Rollmoney.ִ�в���id = r_Rollsend.���˿���id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_ҩƷ�շ���¼_������ҩ(r_Drug.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 2);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              Else
                --סԺ���˷��÷��͵�����������������Դ����סԺ�� 
                For r_Rollmoney In c_Rollmoneyout(r_Rollsend.���ͺ�, r_Rollsend.ҽ��id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.��¼״̬ <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���, 1);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1) And
                       Not r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --��ͨ����ֱ��ȡ��ִ��״̬������ҩƷ�͸������õ����� 
                      Update ������ü�¼
                      Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
                      Where NO = r_Rollmoney.No And ��¼���� = r_Rolladvice.��¼���� And ��¼״̬ = r_Rollmoney.��¼״̬ And
                            Nvl(�۸񸸺�, ���) = r_Rollmoney.��� And ҽ����� = r_Rollsend.ҽ��id;
                    Elsif r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1 Then
                      --�������õ����ģ����Զ����� 
                      For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                        Zl_�����շ���¼_��������(r_Stuff.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, 0, v_��Ա����);
                      End Loop;
                    Elsif r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --�����ҷ�ҩ��ҩƷ�Զ���ҩ 
                      If r_Rollmoney.ִ�в���id = r_Rollsend.���˲���id Or r_Rollmoney.ִ�в���id = r_Rollsend.���˿���id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_ҩƷ�շ���¼_������ҩ(r_Drug.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 1);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              End If;
            End If;
          End Loop;
        End If;
        ------------------------------------------------------------------ 
        --�������ջصĳ���ҩƷҽ�����������(���˷��þͶ�����) 
        If Nvl(r_Rolladvice.ҽ����Ч, 0) = 0 Then
          If r_Rolladvice.�ϴ�ִ��ʱ�� Is Not Null And r_Rolladvice.ĩ��ʱ�� Is Not Null Then
            If r_Rolladvice.�ϴ�ִ��ʱ�� < r_Rolladvice.ĩ��ʱ�� Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '������ڷ��͵������ѱ��ջأ������ٻ��ˡ�';
              Raise Err_Custom;
            End If;
          Elsif r_Rolladvice.�ϴ�ִ��ʱ�� Is Null And r_Rolladvice.ĩ��ʱ�� Is Not Null Then
            --�������ܱ�ȫ�������ջ� 
            v_Error := Nvl(ҽ������_In, '��ҽ��') || 'δ�����ͣ����͵������ѱ�ȫ�������ջأ������ٻ��ˡ�';
            Raise Err_Custom;
          End If;
        End If;
      
        If Nvl(r_Rolladvice.ִ��״̬, 0) In (1, 3) And v_Update <> 1 Then
          --1-��ȫִ��;3-����ִ�� 
          v_Error := Nvl(ҽ������_In, '��ҽ��') || '������͵������Ѿ�ִ�л�����ִ�У����ܻ��ˡ�';
          Raise Err_Custom;
        Else
          --������ҽ����ִ�У���ҲҪ���ƻ��ˣ����磺����Ĳɼ���ʽ�� 
          Select /*+ Rule*/
           Count(1)
          Into v_Count
          From ����ҽ������
          Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And ִ��״̬ In (1, 3) And
                ���ͺ� =
                (Select Max(���ͺ�) From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(t_Adviceids)));
          If v_Count > 0 Then
            v_Error := Nvl(ҽ������_In, '��ҽ��') || '������͵������Ѿ�ִ�л�����ִ�У����ܻ��ˡ�';
            Raise Err_Custom;
          End If;
        End If;
      
        ------------------------------------------------------------------ 
        --������ҽ���ķ�������(��һ��ҽ�������в�ͬNO����) 
        --���ԭʼ�����ѱ�����(�򲿷�����),���ù��������ж� 
        v_����no   := Null;
        v_������� := Null;
        If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 0 Then
          For r_Rollmoney In c_Rollmoneyin(r_Rolladvice.���ͺ�, Null, t_Adviceids) Loop
            --��Ӧ�ķ�����ִ�� 
            If Nvl(r_Rollmoney.ִ��״̬, 0) <> 0 Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '���͵ķ��õ���"' || r_Rollmoney.No || '"�е������ѱ����ֻ���ȫִ�У����ܻ��ˡ�';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.��¼״̬ <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���);
            End If;
            If n_Blndo > 0 Then
              --���ֽ������жϲ�����ҩ 
              If v_����no <> r_Rollmoney.No And v_������� Is Not Null Then
                Zl_סԺ���ʼ�¼_Delete(v_����no, Substr(v_�������, 2), v_��Ա���, v_��Ա����, 2, 0, 0);
                v_������� := Null;
              End If;
              v_����no   := r_Rollmoney.No;
              v_������� := v_������� || ',' || r_Rollmoney.���;
            End If;
          End Loop;
        Else
          For r_Rollmoney In c_Rollmoneyout(r_Rolladvice.���ͺ�, Null, t_Adviceids) Loop
            --��Ӧ�ķ�����ִ�� 
            If Nvl(r_Rollmoney.ִ��״̬, 0) <> 0 And Not (Nvl(r_Rollmoney.ִ��״̬, 0) = -1 And Nvl(r_Rollmoney.��¼״̬, 0) = 0) Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '���͵ķ��õ���"' || r_Rollmoney.No || '"�е������ѱ����ֻ���ȫִ�У����ܻ��ˡ�';
              Raise Err_Custom;
            End If;
            --�շѵ������շ� 
            If r_Rollmoney.��¼״̬ = 1 And Not (r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 1) Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '���͵����ﵥ��"' || r_Rollmoney.No || '"���շѣ����ܻ��ˡ�';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.��¼״̬ <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���, 1);
            End If;
            If n_Blndo > 0 Then
              --���ֽ������жϲ�����ҩ 
              If v_����no <> r_Rollmoney.No And v_������� Is Not Null Then
                If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 1 Then
                  --סԺ����Ϊ�������(���������ҽ������Ϊ������ʣ�����ҽ��û�л��˹���) 
                  Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����);
                Else
                  Zl_���ﻮ�ۼ�¼_Delete(v_����no, Substr(v_�������, 2));
                End If;
                v_������� := Null;
              End If;
              v_����no   := r_Rollmoney.No;
              v_������� := v_������� || ',' || r_Rollmoney.���;
            End If;
          End Loop;
        End If;
        If v_������� Is Not Null And v_����no Is Not Null Then
          v_������� := Substr(v_�������, 2);
          If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 0 Then
            Zl_סԺ���ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����, 2, 0, 0);
          Elsif r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 1 Then
            --סԺ����Ϊ�������(���������ҽ������Ϊ������ʣ�����ҽ��û�л��˹���) 
            Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����);
          Else
            Zl_���ﻮ�ۼ�¼_Delete(v_����no, v_�������);
          End If;
        End If;
      
        --���˷��Ͳ�����ͨ�����ͼ�¼������Ϣ�����������ҽ��Ҫ��ǰ      
        For R In (Select a.����id, a.��ҳid, b.No, b.���ͺ�, b.��������, b.�״�ʱ��, b.ĩ��ʱ��, b.��������, a.Id, a.���id,
                         Nvl(a.���id, a.Id) As ��id, c.���, c.��������, a.ִ�п���id
                  From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C
                  Where a.Id = b.ҽ��id And a.������Ŀid = c.Id And b.���ͺ� = r_Rolladvice.���ͺ� And
                        b.ҽ��id In (Select Column_Value From Table(t_Adviceids))
                  Order By a.���) Loop
        
          --�˴������Ϣ����
          If r.��� = 'D' And r.���id Is Null Then
            --��� 
            b_Message.Zlhis_Cis_037(r.����id, r.��ҳid, Null, r.���ͺ�, r.��id, r.No, 2);
          Elsif r.��� = 'F' And r.���id Is Null Then
            --���� 
            b_Message.Zlhis_Cis_038(r.����id, r.��ҳid, Null, r.���ͺ�, r.��id, r.No);
          Elsif r.��� = 'K' And r.���id Is Null Then
            --��Ѫ 
            b_Message.Zlhis_Cis_039(r.����id, r.��ҳid, Null, r.���ͺ�, r.��id, r.No);
          Elsif r.��� = 'E' And r.�������� = '6' Then
            --����
            b_Message.Zlhis_Cis_036(r.����id, r.��ҳid, Null, r.���ͺ�, r.��id, r.No, 2);
          End If;
        
          Select Count(1) Into v_Count From ��������˵�� A Where a.����id = r.ִ�п���id And a.�������� = '����';
          If v_Count > 0 Then
            --����ִ��ҽ�����˷���
            b_Message.Zlhis_Cis_044(r.����id, r.��ҳid, r.���ͺ�, r.Id, r.No, r.��������, r.�״�ʱ��, r.ĩ��ʱ��, r.��������);
          End If;
        End Loop;
      
        --��Ѫҽ����ɾ������ҽ������ 
        Delete From ����ҽ������ Where ���ͺ� = r_Rolladvice.���ͺ� And ҽ��id = ҽ��id_In;
      
        --ɾ��ҽ��ִ��ʱ�� (����ҽ��ID�Ų����˼�¼) 
        Delete From ҽ��ִ��ʱ�� Where ���ͺ� = r_Rolladvice.���ͺ� And ҽ��id = ҽ��id_In;
      
        --ɾ�����ͼ�¼(����ҽ����) 
        Delete /*+ Rule*/
        From ����ҽ������
        Where ���ͺ� = r_Rolladvice.���ͺ� And ҽ��id In (Select Column_Value From Table(t_Adviceids));
      
        --���(����ҽ��)�ϴ�ִ��ʱ��(���ϴη��͵�ĩ��ִ��ʱ��) 
        --���г���(���������Գ���)����ʱ����д��ĩ��ʱ�� 
        --��������û�У���ֻ���ܷ�����һ�Ρ� 
        v_ĩ��ʱ�� := Null;
        Begin
          --һ��ҽ���ķ�����ĩʱ����ͬ,һ����ҩ��ȡ��С�� 
          --ȡ���IDΪNULL��ҽ���ķ��ͼ�¼��ʱ�� 
          --����ҩ;������ҩ�÷�����δ��д���ͼ�¼ 
          Select /*+ Rule*/
           ĩ��ʱ��
          Into v_ĩ��ʱ��
          From ����ҽ������
          Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And
                ���ͺ� =
                (Select Max(���ͺ�) From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(t_Adviceids))) And
                Rownum = 1;
        Exception
          When Others Then
            Null;
        End;
        Update ����ҽ����¼ Set �ϴ�ִ��ʱ�� = v_ĩ��ʱ�� Where ID = ҽ��id_In Or ���id = ҽ��id_In;
      
        --������������ʱ��ͬʱ�Զ�����ֹͣ 
        If Nvl(r_Rolladvice.ҽ����Ч, 0) = 1 Then
          --ɾ��(��������)�����ֹͣ״̬������¼ 
          Delete /*+ Rule*/
          From ����ҽ��״̬
          Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And
                ����ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = ҽ��id_In) And �������� = 8;
          --r_RollAdvice.����ʱ��:����ʱ����ܲ����Զ�ֹͣʱ����ͬ�� 
        
          --ȡɾ����Ӧ�ָ���ҽ��״̬ 
          Select ��������
          Into v_ҽ��״̬
          From ����ҽ��״̬
          Where ����ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = ҽ��id_In) And ҽ��id = ҽ��id_In;
        
          --�ָ�(����ҽ��)���˺��״̬ 
          Update ����ҽ����¼
          Set ҽ��״̬ = v_ҽ��״̬, ִ����ֹʱ�� = Null, ͣ��ҽ�� = Null, ͣ��ʱ�� = Null
          Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        End If;
      
        --סԺ����ҽ�����ͺ�Ļ���(3-ת��;5-��Ժ;6-תԺ,11-����) 
        If r_Rolladvice.��� = 'Z' And Instr(',3,5,6,11,', ',' || r_Rolladvice.���� || ',') > 0 And
           Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
          Open c_Patilog(r_Rolladvice.����id, r_Rolladvice.��ҳid);
          Fetch c_Patilog
            Into r_Patilog;
          If c_Patilog%Found Then
            If r_Rolladvice.���� = '3' And r_Patilog.��ʼԭ�� = 3 Then
              --ȡ������ת��״̬ 
              If r_Patilog.��ʼʱ�� Is Null Then
                --ת��ҽ�������⴦����һ������������ת��ҽ��ʱ��ֻ�ܻ��������һ��,70443 
                Select Count(1)
                Into v_Count
                From ����ҽ����¼ A, ������ĿĿ¼ B
                Where a.������Ŀid = b.Id And a.����id = r_Rolladvice.����id And a.��ҳid = r_Rolladvice.��ҳid And a.������� = 'Z' And
                      b.�������� = '3' And a.ҽ��״̬ = 8 And
                      a.��ʼִ��ʱ�� > (Select ��ʼִ��ʱ�� From ����ҽ����¼ Where ID = ҽ��id_In);
                If v_Count = 0 Then
                  Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, Null, Null, Null, 'ת��');
                Else
                  v_Error := '����ת���Ѿ���ƣ������ٻ��ˡ�';
                  Raise Err_Custom;
                End If;
              Else
                v_Error := '����ת���Ѿ���ƣ������ٻ��ˡ�';
                Raise Err_Custom;
              End If;
            Elsif r_Rolladvice.���� In ('5', '6', '11') And r_Patilog.��ʼԭ�� = 10 Then
              --ȡ������Ԥ��Ժ״̬ 
              Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, Null, Null, Null, 'Ԥ��Ժ');
            End If;
          End If;
          Close c_Patilog;
        End If;
      
        --���˲���ʱ�� 
        --1.�����¼�(ֻ��һ��ҽ����¼)��������7-����,8-����,11-���� 
        If r_Rolladvice.��� = 'F' Or r_Rolladvice.��� = 'Z' And Instr(',7,8,11,', ',' || r_Rolladvice.���� || ',') > 0 Then
          Zl_���Ӳ���ʱ��_Delete(r_Rolladvice.����id, r_Rolladvice.��ҳid, 'ҽ��', r_Rolladvice.��������id, ҽ��id_In);
        End If;
      
        --2.���⴦��֪��ͬ����(������ص�֪��ͬ�����ٴε��ã���Ϊ����������������Ŀ�����й�����֪��ͬ����) 
        If Instr('C,D,E,F,G,K,L', r_Rolladvice.���) > 0 Then
          For R In (Select a.Id, a.������� From ����ҽ����¼ A Where a.Id = ҽ��id_In Or a.���id = ҽ��id_In) Loop
            --���id��һ��ҽ����һ����������ģ�����Ҫ���ж�һ����� 
            If Instr('C,D,E,F,G,K,L', r.�������) > 0 Then
              Zl_���Ӳ���ʱ��_Delete(r_Rolladvice.����id, r_Rolladvice.��ҳid, 'ҽ��', r_Rolladvice.��������id, r.Id);
            End If;
          End Loop;
        End If;
      
        If r_Rolladvice.��� = 'Z' And r_Rolladvice.�������� = '6' Then
          --���� 
          b_Message.Zlhis_Cis_040(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'Z' And r_Rolladvice.�������� = '8' Then
          --���� 
          b_Message.Zlhis_Cis_041(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'Z' And r_Rolladvice.�������� = '11' Then
          --���� 
          b_Message.Zlhis_Cis_042(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'E' And r_Rolladvice.�������� = '5' Then
          --�������� 
          b_Message.Zlhis_Cis_043(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'H' And Nvl(r_Rolladvice.��������, '0') = '0' Then
          --������ 
          b_Message.Zlhis_Cis_007(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        End If;
      End If;
    End If;
    Exit When r_Rolladvice.���ͺ� = 0;
  End Loop;
  Close c_Rolladvice;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_����;
/
--126851:��С��,2018-11-28,�걾���պ�ı����ִ��״̬Ϊ0
Create Or Replace Procedure Zl_�����������_Update
(
  ҽ��ids_In  In Varchar2, --���ҽ��ID�ö��ŷָ�
  ִ��˵��_In In ����ҽ������.ִ��˵��%Type
) Is
  v_�������� ����ҽ������.��¼����%Type;

  Cursor c_Samplequest Is
    Select Distinct ID As ҽ��id, ������Դ
    From ����ҽ����¼ A
    Where a.Id In (Select * From Table(Cast(f_Num2list(ҽ��ids_In) As Zltools.t_Numlist)));
Begin
  --����ҽ��ִ��״̬
  Update ����ҽ������
  Set ִ��״̬ = 2, ִ��˵�� = ִ��˵��_In, ������ = Null, ����ʱ�� = Null, �ͼ��� = Null, �걾�ͳ�ʱ�� = Null, �걾�������� = Null, ������ = Null,
      ����ʱ�� = Null
  Where ҽ��id In (Select * From Table(Cast(f_Num2list(ҽ��ids_In) As Zltools.t_Numlist)));

  --�������ִ��״̬
  For r_Samplequest In c_Samplequest Loop
    If r_Samplequest.������Դ = 2 Then
      Select Decode(��¼����, 1, 1, Decode(�������, 1, 1, 2))
      Into v_��������
      From ����ҽ������
      Where ҽ��id = r_Samplequest.ҽ��id;
    Else
      v_�������� := 1;
    End If;
  
    If v_�������� = 2 Then
      Update סԺ���ü�¼
      Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
      Where �շ���� Not In ('5', '6', '7') And
            (ҽ�����, ��¼����, NO) In
            (Select ҽ��id, ��¼����, NO
             From ����ҽ������
             Where ҽ��id = r_Samplequest.ҽ��id
             Union All
             Select ҽ��id, ��¼����, NO
             From ����ҽ������
             Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Not Null));
    Else
      Update ������ü�¼
      Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
      Where �շ���� Not In ('5', '6', '7') And
            (ҽ�����, ��¼����, NO) In
            (Select ҽ��id, ��¼����, NO
             From ����ҽ������
             Where ҽ��id = r_Samplequest.ҽ��id
             Union All
             Select ҽ��id, ��¼����, NO
             From ����ҽ������
             Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Not Null));
    End If;
  
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����������_Update;
/



------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0039' Where ���=&n_System;
Commit;