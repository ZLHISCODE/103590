--10.35.0

--00000:������,2015-03-13,��������(���д���δ�Ǽ�BUG)
--84990:��˶,2015-06-12,��������ɾ�������ϴ����صȹ���
Delete zlPrograms where ϵͳ Is Null And ���=15;
--00000:��˶,2015-08-20,ģ�������Ȩ����ֶ�
alter table zltools.zlprograms drop column ����;
--00000:��˶,2015-05-04,��������(���д���δ�Ǽ�BUG)
alter table Zltools.zlParameters add ���� NUMBER(1);
alter table Zltools.zlParameters add ���� NUMBER(1);
Alter Table Zltools.zlParameters Add Constraint zlParameters_CK_���� Check (���� IN(0,1));
Alter Table Zltools.zlParameters Add Constraint zlParameters_CK_���� Check (���� IN(0,1));

alter table Zltools.zlParameters Add Ӱ�����˵�� varchar2(2000);
alter table Zltools.zlParameters add ����ֵ���� varchar2(2000);
alter table Zltools.zlParameters add ����˵�� varchar2(2000);
alter table Zltools.zlParameters add ����˵�� varchar2(2000);
alter table Zltools.zlParameters add ����˵�� varchar2(2000);
alter table Zltools.Zlparachangedlog modify �䶯���� varchar2(4000);

--89346:��˶,2015-10-20,�Զ�����
Insert Into Zlparameters
  (Id, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, -null, -null, 1, -null, -null, -null, 0, 0, 25, '�Զ�����', '5', '5', '���ָ���������Զ�����ϵͳ',
         '0��NUll���������Զ�������>0���Զ���������ķ�����', Null, Null, Null
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '�Զ�����');

--84990:��˶,2015-06-13,��������
Update zlParameters Set  Ӱ�����˵�� = '��¼�Զ���Ϣ������Ϣͣ��ʱ��(��)' , ����ֵ���� = '��λ����'  Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '�Զ���Ϣͣ��ʱ��';
Update zlParameters Set  Ӱ�����˵�� = '�������ʼ����������Ƿ���ʾ�Ѷ��ʼ�' Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '��ʾ�Ѷ��ʼ�';
Update zlParameters Set  Ӱ�����˵�� = '��¼���ʹ�õĲ�Ʒģ�飬�����ڵ���̨��ʷ�˵�����ʾ' Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '���ʹ��ģ��';
Update zlParameters Set  Ӱ�����˵�� = '�����Ƿ���䵱ǰ�û��Ľ��������Ա��´ν���ʱ������ǰ������,�������ڵ�λ�á���ߣ������п�˳��ȡ�' Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = 'ʹ�ø��Ի����';
Update zlParameters Set  Ӱ�����˵�� = '�����Ƿ�����ʼ���Ϣ֪ͨ' Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '�����ʼ���Ϣ';
Update zlParameters Set  ���� = 1 , Ӱ�����˵�� = '��¼Brower��񵼺�̨�����С����С����ֱ�Ϊ��0-9��,1-11��,2-12��' Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = 'zlBrwFontSize';
Update zlParameters Set  ���� = 1 , Ӱ�����˵�� = '����MDI��񵼺�̨��������ɫ'  Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = 'zlMdiFontColor';
Update zlParameters Set  ���� = 1 , Ӱ�����˵�� = '����MDI��񵼺�̨�ı���ͼƬ�ļ�·��' Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = 'zlMdiBackPic';
Update zlParameters Set  ���� = 1 , Ӱ�����˵�� = '����MDI��񵼺�̨�˵����з�ʽ' , ����ֵ���� = '0-�������У�1-��������' Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = 'zlMdiMenuArray';
Update zlParameters Set  ���� = 1 , Ӱ�����˵�� = '����Windows��񵼺�̨��������ɫ'  Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = 'zlWinFontColor';
Update zlParameters Set  ���� = 1 , Ӱ�����˵�� = '����Windows��񵼺�̨�ı���ͼƬ�ļ�·��'  Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = 'zlWinBackPic';
Update zlParameters Set  ���� = 1 , Ӱ�����˵�� = '��¼ʹ���������͵ĵ���̨��zlBrw��zlWin��zlMdi' , ����ֵ���� = '����̨����'  Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '����̨';
Update zlParameters Set  Ӱ�����˵�� = '�����Ƿ�������������ṩ�Զ����ع���'  Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '������������';
Update zlParameters Set  Ӱ�����˵�� = '�����סԺҽ��վ���Լ�ҩ��������ҩ�Ͳ��ŷ�ҩ�Ƚ��棬ҩƷ������ʾ�������浥����ϸ������������桢ֱ�ӽ����ҩƷѡ����ʱ��ҩƷ������ʾ��' , ����ֵ���� = '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��' , ����˵�� = '�ٴ�������Աϰ�߿�ҩƷͨ������ҩ����Աϰ�߿�ҩƷ��Ʒ��' Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = 'ҩƷ������ʾ';
Update zlParameters Set  Ӱ�����˵�� = '�����סԺҽ��վ�������շѺ�סԺ���ʵȷ�����ؽ��棬����ҩƷʱ�����ַ�ʽ��ʾ��ͨ��������뷽ʽ����ѡ����ʱҩƷ���Ƶ���ʾ��' , ����ֵ���� = '0-������ƥ����ʾ��1-�̶���ʾͨ��������Ʒ��'  Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '����ҩƷ��ʾ';
Update zlParameters Set  Ӱ�����˵�� = '�����ڴ��ڽ���Ĺ������л�����ƥ�䷽ʽ��������ʱ����������ʾ�л���ť��' Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '����ƥ�䷽ʽ�л�';
Update zlParameters Set  Ӱ�����˵�� = '����������������߶��������л����Զ������������ݿ�' , ����ֵ���� = '0-����⣬1-���' Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '��������Զ�����';
Update zlParameters Set  Ӱ�����˵�� = '������ʾ�ڵ���̨�������ϵĳ��ù���ģ��' , ����ֵ���� = 'ģ��1����ϵͳ,ģ��2����ϵͳ|ģ��1���,ģ��2���|ģ��1ͼ��,ģ��2ͼ��|ģ��1����,ģ��2����' , ���� = 0 , ��Ȩ = 0 , �̶� = 0 Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '���ù���ģ��';
Update zlParameters Set  Ӱ�����˵�� = '���ø���ҵ������в�������ʱ��ƥ�䷽��' , ����ֵ���� = '0-˫��ƥ�䣬1-����ƥ��' Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '����ƥ��';
Update zlParameters Set  Ӱ�����˵�� = '���ø���ҵ�������,��������ı�������Զ����������뷨����' , ����ֵ���� = '���뷨����'  Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '���뷨';
Update zlParameters Set  Ӱ�����˵�� = '���ø���ҵ�������,��������ʱ�ļ���ƥ�䷽ʽ' , ����ֵ���� = '0-ƴ����1-���'  Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '���뷽ʽ';
Update zlParameters Set  Ӱ�����˵�� = '�����Ƿ��˳�����ʱ�Զ��ر� Windows'  Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '�ر�Windows';
Update zlParameters Set  Ӱ�����˵�� = '�����Զ�����ʼ���Ϣ��ʱ����(��)' , ����ֵ���� = '��λ����' Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '�ʼ���Ϣ�������';
Update zlParameters Set  Ӱ�����˵�� = '���õ�¼ʱ�Ƿ����µ��ʼ���Ϣ' Where Nvl(ϵͳ, 0) = 0 And Nvl(ģ��, 0) = 0 And ������ = '��¼����ʼ���Ϣ';

Create Table Zltools.zlDeptParas(
    ����ID NUMBER(18),
    ����ID NUMBER(18),
    ����ֵ VARCHAR2(2000))
    PCTFREE 5
    Cache Storage(Buffer_Pool Keep);
Alter Table Zltools.zlDeptParas Add Constraint zlDeptParas_UQ_����ID Unique(����ID,����ID) Using Index PCTFREE 5;
Alter Table Zltools.zlDeptParas Add Constraint zlDeptParas_FK_����ID Foreign Key (����ID) References zlParameters(ID) On Delete Cascade;
--79998:��˶,2015-08-18,���븴���Կ���
Insert Into zlOptions(������,������,����ֵ,ȱʡֵ,����˵��) Values(20, '�Ƿ�������볤��', '','', '�Ƿ��������볤�ȿ���');
Insert Into zlOptions(������,������,����ֵ,ȱʡֵ,����˵��) Values(21, '���볤������', '','3', '�����������С����');
Insert Into zlOptions(������,������,����ֵ,ȱʡֵ,����˵��) Values(22, '���볤������', '','12', '�����������󳤶�');
Insert Into zlOptions(������,������,����ֵ,ȱʡֵ,����˵��) Values(23, '�Ƿ�������븴�Ӷ�', '','', '�Ƿ������������������һ����ĸ�����֡��������ַ������������ַ����ܵ����������롣');

--00000:������,2015-03-20,��������(���д���δ�Ǽ�BUG)
Create Or Replace Procedure Zltools.Zl_Parameters_Change_Value
(
  ����id_In     Zlparachangedlog.����id%Type,
  �䶯����_In   Zlparachangedlog.�䶯����%Type, --ԭֵ-->��ֵ
  �䶯ԭ��_In   Zlparachangedlog.�䶯ԭ��%Type,
  ����Ա����_In Zlparachangedlog.�䶯��%Type,
  �䶯ʱ��_In   Zlparachangedlog.�䶯ʱ��%Type
) Is
  n_Max��� Zlparachangedlog.���%Type;
Begin
  Select Nvl(Max(���), 1)+1 Into n_Max��� From Zlparachangedlog Where ����id = ����id_In;

  Insert Into Zlparachangedlog
    (����id, ���, �䶯˵��, �䶯����, �䶯��, �䶯ʱ��, �䶯ԭ��)
  Values
    (����id_In, n_Max���, 'ֵ�䶯', �䶯����_In, ����Ա����_In, �䶯ʱ��_In, �䶯ԭ��_In);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Parameters_Change_Value;
/

CREATE OR REPLACE Procedure Zltools.Zl_Parameters_Update_Batch
(
  ϵͳ���_In   Zlsystems.���%Type,
  �����б�_In   Varchar2, --ģ���1^������1^����ֵ1#ģ���2^������2^����ֵ2......
  ����Ա����_In Zlparachangedlog.�䶯��%Type
) Is
  t_ģ��   t_Numlist;
  t_������ t_Numlist;
  t_����ֵ t_Strlist;
Begin
  Select To_Number(C1), To_Number(Substr(C2, 1, Instr(C2, '^') - 1)), Substr(C2, Instr(C2, '^') + 1) Bulk Collect
  Into t_ģ��, t_������, t_����ֵ
  From Table(f_Str2list2(�����б�_In, '#', '^'));

  For Rs In (Select /*+ rule*/
              a.Id, a.����ֵ || '-->' || Substr(C2, Instr(C2, '^') + 1) As �䶯����, Sysdate As �䶯ʱ��
             From zlParameters A, Table(f_Str2list2(�����б�_In, '#', '^')) B
             Where a.ϵͳ = ϵͳ���_In And Nvl(a.ģ��, 0) = To_Number(b.C1) And
                   a.������ = To_Number(Substr(b.C2, 1, Instr(b.C2, '^') - 1)) And a.����˵�� Is Null) Loop
    --�о���˵���Ĺؼ������������ڽ����ṩ�䶯�Ǽ�
    Zl_Parameters_Change_Value(Rs.Id, Rs.�䶯����, '', ����Ա����_In, Rs.�䶯ʱ��);
  End Loop;

  Forall I In 1 .. t_������.Count
    Update zlParameters
    Set ����ֵ = t_����ֵ(I)
    Where ϵͳ = ϵͳ���_In And Nvl(ģ��, 0) = t_ģ��(I) And ������ = t_������(I);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Parameters_Update_Batch;
/
--83832:���Ʊ�,2015-04-07,��������
CREATE OR REPLACE Procedure Zltools.Zl_DeptParameters_Delete
(
  ����_In   Zlparameters.������%Type,
  ϵͳ_In   Zlparameters.ϵͳ%Type,
  ģ��_In   Zlparameters.ģ��%Type
  --���ܣ�ɾ���������͵Ķ�Ӧ���ŵ����в���
  --������
  --     ����_In�����봫���Nullֵ�����ַ���ʽ����Ĳ����Ż������,ע�����������Ϊ���֡�
  --     Ȩ��_IN������Ҫ����Ȩ�޿��ƵĲ�������ǰ�û��Ƿ���Ȩ������
) Is
  v_����id Zlparameters.Id%Type;
  v_˽��   Zlparameters.˽��%Type;
  v_����   Zlparameters.����%Type;
  v_��Ȩ   Zlparameters.��Ȩ%Type;
  v_������ Zluserparas.������%Type;
  v_����   Zlparameters.����%Type;
Begin
  --ȷ��������Ϣ
  Begin
    If Zl_To_Number(����_In) <> 0 Then
      --�Բ�����Ϊ׼����
      Select ID, ˽��, ����, ��Ȩ, Sys_Context('USERENV', 'TERMINAL'), ����
      Into v_����id, v_˽��, v_����, v_��Ȩ, v_������, v_����
      From zlParameters
      Where Nvl(ϵͳ, 0) = Nvl(ϵͳ_In, 0) And Nvl(ģ��, 0) = Nvl(ģ��_In, 0) And ������ = Zl_To_Number(����_In);
    Else
      --�Բ�����Ϊ׼����
      Select ID, ˽��, ����, ��Ȩ, Sys_Context('USERENV', 'TERMINAL'), ����
      Into v_����id, v_˽��, v_����, v_��Ȩ, v_������, v_����
      From zlParameters
      Where Nvl(ϵͳ, 0) = Nvl(ϵͳ_In, 0) And Nvl(ģ��, 0) = Nvl(ģ��_In, 0) And ������ = ����_In;
    End If;
  Exception
    When Others Then
      Return;
  End;

  If Nvl(v_����, 0) = 0 Then
    Return; --���ż�ģ�����
  End If;

  --���²���ֵ
  If v_����id Is Not Null Then
     Delete From zldeptparas Where ����id = v_����id;
  End If;
End Zl_DeptParameters_Delete;
/

--83832:���Ʊ�,2015-04-07,��������
Create Or Replace Procedure Zltools.Zl_Parameters_Update
(
  ����_In   Zlparameters.������%Type,
  ����ֵ_In Zlparameters.����ֵ%Type,
  ϵͳ_In   Zlparameters.ϵͳ%Type,
  ģ��_In   Zlparameters.ģ��%Type,
  Ȩ��_In   Number := 1,
  ����id_In zldeptparas.����id%Type := 0
  --���ܣ�����ϵͳ����ֵ��������û�˽�в��������û����Ե�ǰ��Ϊ׼
  --������
  --     ����_In�����봫���Nullֵ�����ַ���ʽ����Ĳ����Ż������,ע�����������Ϊ���֡�
  --     Ȩ��_IN������Ҫ����Ȩ�޿��ƵĲ�������ǰ�û��Ƿ���Ȩ������
) Is
  v_����id Zlparameters.Id%Type;
  v_˽��   Zlparameters.˽��%Type;
  v_����   Zlparameters.����%Type;
  v_��Ȩ   Zlparameters.��Ȩ%Type;
  v_������ Zluserparas.������%Type;
  v_����   Zlparameters.����%Type;
Begin
  --ȷ��������Ϣ
  Begin
    If Zl_To_Number(����_In) <> 0 Then
      --�Բ�����Ϊ׼����
      Select ID, ˽��, ����, ��Ȩ, Sys_Context('USERENV', 'TERMINAL'), ����
      Into v_����id, v_˽��, v_����, v_��Ȩ, v_������, v_����
      From zlParameters
      Where Nvl(ϵͳ, 0) = Nvl(ϵͳ_In, 0) And Nvl(ģ��, 0) = Nvl(ģ��_In, 0) And ������ = Zl_To_Number(����_In);
    Else
      --�Բ�����Ϊ׼����
      Select ID, ˽��, ����, ��Ȩ, Sys_Context('USERENV', 'TERMINAL'), ����
      Into v_����id, v_˽��, v_����, v_��Ȩ, v_������, v_����
      From zlParameters
      Where Nvl(ϵͳ, 0) = Nvl(ϵͳ_In, 0) And Nvl(ģ��, 0) = Nvl(ģ��_In, 0) And ������ = ����_In;
    End If;
  Exception
    When Others Then
      Return;
  End;

  --���Ȩ��
  If Nvl(Ȩ��_In, 0) = 0 Then
    If Nvl(v_����, 0) <> 0 Then
      Return; --���ż�ģ�����
    Elsif Nvl(ϵͳ_In, 0) <> 0 And Nvl(ģ��_In, 0) = 0 And Nvl(v_˽��, 0) = 0 And Nvl(v_����, 0) = 0 Then
      Return; --����ȫ�ֲ���,�̶���ҪȨ��
    Elsif Nvl(ϵͳ_In, 0) <> 0 And Nvl(ģ��_In, 0) <> 0 And Nvl(v_˽��, 0) = 0 And Nvl(v_����, 0) = 0 Then
      Return; --����ģ�����,�̶���ҪȨ��
    Elsif Nvl(ϵͳ_In, 0) <> 0 And Nvl(ģ��_In, 0) <> 0 And Nvl(v_˽��, 0) = 0 And Nvl(v_����, 0) = 1 And Nvl(v_��Ȩ, 0) = 1 Then
      Return; --Ҫ��Ȩ���Ƶı�������ģ��
    End If;
  End If;

  --���²���ֵ
  If v_����id Is Not Null Then
    If Nvl(v_����, 0) <> 0 Then
      Update zldeptparas Set ����ֵ = ����ֵ_In Where ����id = v_����id And ����ID= ����id_In;
      If Sql%RowCount = 0 Then
        Insert Into zldeptparas
          (����id, ����ID, ����ֵ)
        Values
          (v_����id,����id_In , ����ֵ_In);
      End If;
    elsIf Nvl(v_˽��, 0) = 0 And Nvl(v_����, 0) = 0 Then
      Update zlParameters Set ����ֵ = ����ֵ_In Where ID = v_����id;
    Else
      Update zlUserParas
      Set ����ֵ = ����ֵ_In
      Where ����id = v_����id And Nvl(�û���, 'NullUser') = Decode(v_˽��, 1, User, 'NullUser') And
            Nvl(������, 'NullMachine') = Decode(v_����, 1, v_������, 'NullMachine');
      If Sql%RowCount = 0 Then
        Insert Into zlUserParas
          (����id, �û���, ������, ����ֵ)
        Values
          (v_����id, Decode(v_˽��, 1, User, Null), Decode(v_����, 1, v_������, Null), ����ֵ_In);
      End If;
    End If;
  End If;
End Zl_Parameters_Update;
/
--84990:��˶,2015-09-22,��������
Create Or Replace Procedure Zltools.Zlparameters_Delall_Details
(
  �����б�_In Varchar2,
  n_����      Number := 0
  --n_����:1���������͵Ĳ�����0-�ǲ������͵Ĳ���
  --�����б�_In ϵͳ1^ģ��1^������1#ϵͳ2......, 
) Is
  t_����id t_Numlist;
Begin
  Select a.Id Bulk Collect
  Into t_����id
  From Zlparameters a,
       (Select Zl_To_Number(C1) ϵͳ, Zl_To_Number(Substr(C2, 1, Instr(C2, '^') - 1)) ģ��,
                Substr(C2, Instr(C2, '^') + 1) ������
         From Table(f_Str2list2(�����б�_In, '#', '^'))) b
  Where Nvl(a.ϵͳ, 0) = Nvl(b.ϵͳ, 0) And Nvl(a.ģ��, 0) = Nvl(b.ģ��, 0) And a.������ = b.������;
  If n_���� = 0 Then
    Forall i In 1 .. t_����id.Count
      Delete Zluserparas Where ����id = t_����id(i);
  Else
    Forall i In 1 .. t_����id.Count
      Delete Zldeptparas Where ����id = t_����id(i);
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zlparameters_Delall_Details;
/
--84990:��˶,2015-09-22,��������
CREATE OR REPLACE Procedure ZLTOOLS.Zlparameters_Add_Details
(
  ����id_In Zlparameters.Id%Type,
  �û���_In Varchar2,
  ������_In Varchar2,
  ����ֵ_In Varchar2
  --�û���_In �Զ��ŷָ�û�1,�û�2,
  --������_In �Զ��ŷָ����1,����2,
) Is
  n_���� Number(1);
Begin
  Select Nvl(����, 0) Into n_���� From Zlparameters Where Id = ����id_In;
  If n_���� = 0 Then
    Insert Into Zluserparas
      (����id, �û���, ������, ����ֵ)
      Select ����id, �û���, ������, ����ֵ
      From (Select ����id_In ����id, a.�û���, b.������, ����ֵ_In ����ֵ
             From (Select Distinct Column_Value �û��� From Table(f_Str2list(Nvl(�û���_In, ',')))) a,
                  (Select Distinct Column_Value ������ From Table(f_Str2list(Nvl(������_In, ',')))) b) c
      Where Not Exists
       (Select 1
             From Zluserparas
             Where ����id = c.����id And Nvl(�û���, '�տ�') = Nvl(c.�û���, '�տ�') And Nvl(������, '�տ�') = Nvl(c.������, '�տ�'));
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zlparameters_Add_Details;
/
--84990:��˶,2015-05-28,��������
Create Or Replace Procedure Zltools.Zlparameters_Update_Details
(
  ����id_In   Zlparameters.Id%Type,
  �����б�_In Varchar2
  --�����б�_In �û���1^������1^����ֵ1#�û���2^������2^����ֵ2......,
  --           �������Ͳ���������ID1,,����ֵ1#����ID2,,����ֵ2
) Is
  n_����   Number(1);
  t_����id t_Numlist;
  t_�û��� t_Strlist;
  t_������ t_Strlist;
  t_����ֵ t_Strlist;
Begin
  Select Nvl(����, 0) Into n_���� From zlParameters Where ID = ����id_In;
  If n_���� = 0 Then
    Select C1, Substr(C2, 1, Instr(C2, '^') - 1), Substr(C2, Instr(C2, '^') + 1) Bulk Collect
    Into t_�û���, t_������, t_����ֵ
    From Table(f_Str2list2(�����б�_In, '#', '^'));
  
    Forall I In 1 .. t_����ֵ.Count
      Update zlUserParas
      Set ����ֵ = t_����ֵ(I)
      Where ����id = ����id_In And Nvl(�û���, '�տ�') = Nvl(t_�û���(I), '�տ�') And Nvl(������, '�տ�') = Nvl(t_������(I), '�տ�');
  Else
    Select To_Number(C1), Substr(C2, 1, Instr(C2, '^') - 1), Substr(C2, Instr(C2, '^') + 1) Bulk Collect
    Into t_����id, t_������, t_����ֵ
    From Table(f_Str2list2(�����б�_In, '#', '^'));
  
    Forall I In 1 .. t_����ֵ.Count
      Update Zldeptparas Set ����ֵ = t_����ֵ(I) Where ����id = ����id_In And ����id = t_����id(I);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlparameters_Update_Details;
/
--84990:��˶,2015-09-28,��������
Create Or Replace Procedure Zltools.Zlparameters_Del_Details
(
  ����id_In   Zlparameters.Id%Type,
  �����б�_In Varchar2
  --�����б�_In �û���1^������1#�û���2^������2......,
  --           �������Ͳ���������ID1#����ID2
) Is
  n_����   Number(1);
  t_����id t_Numlist;
  t_�û��� t_Strlist;
  t_������ t_Strlist;
Begin
  Select Nvl(����, 0) Into n_���� From Zlparameters Where Id = ����id_In;
  If n_���� = 0 Then
    Select C1, C2 Bulk Collect Into t_�û���, t_������ From Table(f_Str2list2(�����б�_In, '#', '^'));
  
    Forall i In 1 .. t_�û���.Count
      Delete Zluserparas
      Where ����id = ����id_In And Nvl(�û���, '�տ�') = Nvl(t_�û���(i), '�տ�') And Nvl(������, '�տ�') = Nvl(t_������(i), '�տ�');
  Else
    Select To_Number(Column_Value) Bulk Collect Into t_����id From Table(f_Str2list(�����б�_In, '#'));
  
    Forall i In 1 .. t_����id.Count
      Delete Zldeptparas Where ����id = ����id_In And ����id = t_����id(i);
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zlparameters_Del_Details;
/
--84990:��˶,2015-05-28,��������
Create Or Replace Procedure Zltools.Zlparameters_Imp_Details
(
  ϵͳ_In     Zlparameters.ϵͳ%Type,
  ģ��_In     Zlparameters.ģ��%Type,
  ����_In     Zlparameters.������%Type,
  �����б�_In Varchar2
  --�����б�_In �û���1^������1^����ֵ1#�û���2^������2^����ֵ2......,
  --           �������Ͳ���������ID1,,����ֵ1#����ID2,,����ֵ2
  --�����б�Ϊ��ʱɾ��������ϸ����
) Is
  n_����id Zlparameters.Id%Type;
  n_����   Number(1);
  n_˽��   Number(1);
  n_����   Number(1);
  t_����id t_Numlist;
  t_�û��� t_Strlist;
  t_������ t_Strlist;
  t_����ֵ t_Strlist;
Begin
  --��ȡ����ID�벿������
  If Zl_To_Number(����_In) <> 0 Then
    Select Nvl(����, 0), Nvl(˽��, 0), Nvl(����, 0), ID
    Into n_����, n_˽��, n_����, n_����id
    From zlParameters
    Where ������ = Zl_To_Number(����_In) And Nvl(ģ��, 0) = Nvl(ģ��_In, 0) And Nvl(ϵͳ, 0) = Nvl(ϵͳ_In, 0);
  Else
    Select Nvl(����, 0), Nvl(˽��, 0), Nvl(����, 0), ID
    Into n_����, n_˽��, n_����, n_����id
    From zlParameters
    Where ������ = ����_In And Nvl(ģ��, 0) = Nvl(ģ��_In, 0) And Nvl(ϵͳ, 0) = Nvl(ϵͳ_In, 0);
  End If;
  If n_����id Is Not Null Then
    If n_���� = 0 Then
      If �����б�_In Is Null Then
        Delete zlUserParas Where ����id = n_����id;
        --˽�л򱾻��������Ų���
      Elsif n_˽�� = 1 Or n_���� = 1 Then
        Select C1, Substr(C2, 1, Instr(C2, '^') - 1), Substr(C2, Instr(C2, '^') + 1) Bulk Collect
        Into t_�û���, t_������, t_����ֵ
        From Table(f_Str2list2(�����б�_In, '#', '^'));
        --����ȫ��ɾ�����ٲ���
        Forall I In 1 .. t_����ֵ.Count
          Insert Into zlUserParas
            (����id, �û���, ������, ����ֵ)
          Values
            (n_����id, t_�û���(I), t_������(I), t_����ֵ(I));
      End If;
    Else
      If �����б�_In Is Null Then
        Delete Zldeptparas Where ����id = n_����id;
      Else
        Select To_Number(C1), Substr(C2, 1, Instr(C2, '^') - 1), Substr(C2, Instr(C2, '^') + 1) Bulk Collect
        Into t_����id, t_������, t_����ֵ
        From Table(f_Str2list2(�����б�_In, '#', '^'));
        --����ȫ��ɾ�����ٲ���
        Forall I In 1 .. t_����ֵ.Count
          Insert Into Zldeptparas (����id, ����id, ����ֵ) Values (n_����id, t_����id(I), t_����ֵ(I));
      End If;
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlparameters_Imp_Details;
/
--84990:��˶,2015-06-18,��������
--84598:��˶,2015-06-17,zl_GetSysParameter�������⴦��
Create Or Replace Function Zltools.zl_GetSysParameter
(
  ����_In   Zlparameters.������%Type,
  ģ��_In   Zlparameters.ģ��%Type := Null,
  ϵͳ_In   Zlparameters.ϵͳ%Type := 1,
  ����id_In Zldeptparas.����id%Type := 0
  --���ܣ���ȡ��ǰϵͳ��ָ�������Ĳ���ֵ 
  ----��������Ҫ���������̵��ã���zlParameters�ǹ�����������ʹ�ù��������Ĺ������� 
  ----����ʱע��,�������ֵΪ�ջ�û�иò���,�򷵻ؿ� 
  --������ 
  ----����_In�����봫���Nullֵ�����ַ���ʽ����Ĳ����Ż������,ע�����������Ϊ���֡� 
  ----ϵͳ_IN���Ǳ�׼��ϵͳ��Ҫ����ϵͳ�ţ�ע����ʾ��չ��ϵͳ�ţ���1��������100����ϵͳ����Null 
) Return Varchar2 As
  v_ϵͳ Zlparameters.ϵͳ%Type;
  v_˽�� Zlparameters.˽��%Type;
  v_���� Zlparameters.����%Type;
  v_���� Zlparameters.����%Type;

  v_����id Zluserparas.����id%Type;
  v_������ Zluserparas.������%Type;
  v_����ֵ Zlparameters.����ֵ%Type;
Begin
  --ȷ��ϵͳ,����û��ϵͳ(��˽��ȫ��) 
  If ϵͳ_In Is Not Null Then
    Select Min(���) Into v_ϵͳ From zlSystems Where Trunc(��� / 100) = ϵͳ_In;
  End If;

  --��ȡ������Ϣ 
  Begin
    If Zl_To_Number(����_In) <> 0 Then
      Select ID, Nvl(����ֵ, ȱʡֵ), ˽��, ����, ����
      Into v_����id, v_����ֵ, v_˽��, v_����, v_����
      From zlParameters
      Where ������ = Zl_To_Number(����_In) And Nvl(ģ��, 0) = Nvl(ģ��_In, 0) And Nvl(ϵͳ, 0) = Nvl(v_ϵͳ, 0);
    Else
      Select ID, Nvl(����ֵ, ȱʡֵ), ˽��, ����, ����
      Into v_����id, v_����ֵ, v_˽��, v_����, v_����
      From zlParameters
      Where ������ = ����_In And Nvl(ģ��, 0) = Nvl(ģ��_In, 0) And Nvl(ϵͳ, 0) = Nvl(v_ϵͳ, 0);
    End If;
  
  Exception
    When Others Then
      Return Null;
  End;
  If Nvl(v_����, 0) = 0 Then
    --��ȡ�ǲ��Ų���ֵ 
    If Nvl(v_˽��, 0) = 1 Or Nvl(v_����, 0) = 1 Then
      If Nvl(v_����, 0) = 1 Then
        Select Sys_Context('USERENV', 'TERMINAL') Into v_������ From Dual;
      End If;
      Begin
        Select Nvl(����ֵ, v_����ֵ)
        Into v_����ֵ
        From zlUserParas
        Where ����id = v_����id And (�û��� = User Or Nvl(v_˽��, 0) = 0) And (������ = Nvl(v_������, '�տ�') Or Nvl(v_����, 0) = 0);
      Exception
        When Others Then
          Return v_����ֵ;
      End;
    End If;
  Else
    Begin
      Select Nvl(����ֵ, v_����ֵ) Into v_����ֵ From Zldeptparas Where ����id = v_����id And ����id = ����id_In;
    Exception
      When Others Then
        Return v_����ֵ;
    End;
  End If;
  Return v_����ֵ;
End zl_GetSysParameter;
/
--84990:��˶,2015-05-20,��������
Create Or Replace Package Zltools.b_Runmana Is

  Type t_Refcur Is Ref Cursor;

  Procedure Get_Parameters
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Number := 0
  );

  Procedure Get_Parameter
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In Zlparameters.Id%Type
  );

  Procedure Get_Parachangedlog
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In Zlparachangedlog.����id%Type
  );

  Procedure Get_Job_Number
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Number
  );

  Procedure Get_Depict
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Zldatamove.ϵͳ%Type,
    ���_In    In Zldatamove.���%Type
  );

  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur);

  Procedure Get_Client
  (
    Cur_Out   Out t_Refcur,
    ����վ_In In Zlclients.����վ%Type := Null
  );

  Procedure Get_Client_Station(Cur_Out Out t_Refcur);

  Procedure Get_Project_No(Cur_Out Out t_Refcur);

  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur);

  Procedure Get_Resile
  (
    Cur_Out   Out t_Refcur,
    ������_In In Zlclientparaset.������%Type,
    ����_In   In Number := 0
  );

  Procedure Get_Zldatamove
  (
    Cur_Out Out t_Refcur,
    ϵͳ_In In Zldatamove.ϵͳ%Type
  );

  Procedure Get_Log
  (
    Cur_Out     Out t_Refcur,
    ��־����_In In Varchar2,
    Where_In    In Varchar2
  );

  Procedure Get_Log_Count
  (
    Cur_Out     Out t_Refcur,
    ��־����_In In Varchar2
  );

  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur);

  Procedure Get_Not_Regist(Cur_Out Out t_Refcur);

  Procedure Get_Zloption
  (
    Cur_Out   Out t_Refcur,
    ������_In In Zloptions.������%Type
  );

End b_Runmana;
/


--84990:��˶,2015-05-28,��������
Create Or Replace Package Body Zltools.b_Runmana Is

  --���ܣ�ȡ������Ϣ
  --frmParameters
  Procedure Get_Parameters
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Number := 0
  ) Is
  Begin
    If Nvl(ϵͳ_In, 0) = 0 Then
      Open Cursor_Out For
        Select a.Id, Nvl(a.ϵͳ, 0) ϵͳ, Nvl(a.ģ��, 0) ģ��, Nvl(a.˽��, 0) ˽��, a.������, a.������, a.����ֵ, a.ȱʡֵ, Nvl(a.����, 0) ����,
               a.Ӱ�����˵��, a.����ֵ����, a.����˵��, a.����˵��, a.����˵��, Nvl(a.����, 0) ����, Nvl(a.��Ȩ, 0) ��Ȩ, Nvl(a.�̶�, 0) �̶�,
               Nvl(a.����, 0) ����, b.���� As ģ������, zlSpellCode(b.����) As ģ�����
        From zlParameters A, zlPrograms B
        Where Nvl(a.ϵͳ, 0) = 0 And Nvl(a.ϵͳ, 0) = b.ϵͳ(+) And Nvl(a.ģ��, 0) = b.���(+);
    Else
      Open Cursor_Out For
        Select a.Id, Nvl(a.ϵͳ, 0) ϵͳ, Nvl(a.ģ��, 0) ģ��, Nvl(a.˽��, 0) ˽��, a.������, a.������, a.����ֵ, a.ȱʡֵ, Nvl(a.����, 0) ����,
               a.Ӱ�����˵��, a.����ֵ����, a.����˵��, a.����˵��, a.����˵��, Nvl(a.����, 0) ����, Nvl(a.��Ȩ, 0) ��Ȩ, Nvl(a.�̶�, 0) �̶�,
               Nvl(a.����, 0) ����, b.���� As ģ������, zlSpellCode(b.����) As ģ�����
        From zlParameters A, zlPrograms B,
             --����Ȩ�޲��֣�ֻ����Ȩ�Ĳ�����ʾ
             (Select Distinct f.���
               From zlProgFuncs F, zlRegFunc R
               Where Trunc(f.ϵͳ / 100) = r.ϵͳ(+) And f.��� = r.���(+) And f.���� = r.����(+) And
                     (r.���� Is Not Null Or r.���� Is Null And (f.��� Between 10000 And 19999)) And f.ϵͳ = ϵͳ_In And
                     1 = (Select 1 From Zlregaudit A Where a.��Ŀ = '��Ȩ֤��')
               Union All
               Select 0 As ���
               From Dual) M
        Where a.ϵͳ = Nvl(ϵͳ_In, 0) And Nvl(a.ϵͳ, 0) = b.ϵͳ(+) And Nvl(a.ģ��, 0) = b.���(+) And Nvl(a.ģ��, 0) = m.���;
    End If;
  End Get_Parameters;

  --���ܣ�����ָ���Ĳ���IDȡ������Ϣ
  --�����б�frmParameters;frmParaChangeSet
  Procedure Get_Parameter
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In Zlparameters.Id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select a.Id, Nvl(a.ϵͳ, 0) ϵͳ, Nvl(a.ģ��, 0) ģ��, Nvl(a.˽��, 0) ˽��, a.������, a.������, a.����ֵ, a.ȱʡֵ, Nvl(a.����, 0) ����,
             a.Ӱ�����˵��, a.����ֵ����, a.����˵��, a.����˵��, a.����˵��, Nvl(a.����, 0) ����, Nvl(a.��Ȩ, 0) ��Ȩ, Nvl(a.�̶�, 0) �̶�,
             Nvl(a.����, 0) ����, b.���� As ģ������, zlSpellCode(b.����) As ģ�����
      From zlParameters A, zlPrograms B
      Where a.Id = Nvl(����id_In, 0) And Nvl(a.ϵͳ, 0) = b.ϵͳ(+) And Nvl(a.ģ��, 0) = b.���(+);
  End Get_Parameter;
  --���ܣ�ȡ�����޸���Ϣ
  --�����б�frmParameters
  Procedure Get_Parachangedlog
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In Zlparachangedlog.����id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ����id, ���, �䶯˵��, �䶯����, �䶯��, �䶯ʱ��, �䶯ԭ��
      From Zlparachangedlog
      Where ����id = Nvl(����id_In, 0);
  
  End;
  --���ܣ�ȡZlAutoJob���к�
  Procedure Get_Job_Number
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Number
  ) Is
  Begin
    Open Cursor_Out For
      Select ��� + 1 As ���
      From zlAutoJobs
      Where Nvl(ϵͳ, 0) = ϵͳ_In And ���� = 3 And
            ��� + 1 Not In (Select ��� From zlAutoJobs Where Nvl(ϵͳ, 0) = ϵͳ_In And ���� = 3);
  End Get_Job_Number;

  --���ܣ�ȡZlDataMove����
  Procedure Get_Depict
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Zldatamove.ϵͳ%Type,
    ���_In    In Zldatamove.���%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ת������ From zlDataMove Where Nvl(ϵͳ, 0) = ϵͳ_In And ��� = ���_In;
  End Get_Depict;

  --���ܣ�ȡzlClients��MAX IP
  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Max(Ip) As Ip From zlClients;
  End Get_Client_Maxip;

  --���ܣ�ȡzlClients�ļ�¼
  Procedure Get_Client
  (
    Cur_Out   Out t_Refcur,
    ����վ_In In Zlclients.����վ%Type := Null
  ) Is
    v_Sql Varchar2(1000);
  Begin
    If Nvl(����վ_In, '��') = '��' Then
      v_Sql := 'Select a.Ip, a.����վ, a.Cpu, a.�ڴ�, a.Ӳ��, a.����ϵͳ, a.����, a.��;, a.˵��, a.������־, a.��ֹʹ��,
                             a.������, Decode(b.Terminal, Null, 0, 1) As ״̬, a.�ռ���־,a.����������,a.վ��,a.������ƵԴ
                From Zlclients a, (Select Distinct Terminal From V$session) b
                Where Upper(a.����վ) = Upper(b.Terminal(+))
                Order By a.Ip';
      Open Cur_Out For v_Sql;
    Else
      Open Cur_Out For
        Select Ip, ����վ, Cpu, �ڴ�, Ӳ��, ����ϵͳ, ����, ��;, ˵��, ������־, �ռ���־, ��ֹʹ��, ������, ����������, վ��, ������ƵԴ
        From zlClients
        Where Upper(����վ) = ����վ_In;
    End If;
  End Get_Client;

  --���ܣ�ȡzlClients��վ��
  Procedure Get_Client_Station(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Distinct Upper(����վ) || '[' || Ip || ']' As վ��, Upper(����վ) ����վ From zlClients;
  End Get_Client_Station;

  --���ܣ�ȡ������
  Procedure Get_Project_No(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select ������ From Zlclientparaset Where Rownum = 1;
  End Get_Project_No;

  --���ܣ�ȡ����
  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select ������, ������ || '-' || �������� As ��������, ��������, ����վ, �û��� From Zlclientscheme;
  End Get_Client_Scheme;

  --���ܣ�ȡ�ָ���Ϣ
  Procedure Get_Resile
  (
    Cur_Out   Out t_Refcur,
    ������_In In Zlclientparaset.������%Type,
    ����_In   In Number := 0
  ) Is
  Begin
    If ����_In = 0 Then
      Open Cur_Out For
        Select Distinct a.����վ || Decode(m.����վ, Null, ' ', '[' || m.Ip || ']') As ����վ, a.�û���, a.�ָ���־,
                        '[' || b.������ || ']' || b.�������� As ��������
        From Zlclientparaset A, Zlclientscheme B, zlClients M
        Where a.������ = b.������ And a.����վ = m.����վ(+) And a.������ = ������_In;
    End If;
  
    If ����_In = 1 Then
      Open Cur_Out For
        Select Distinct Upper(����վ) ����վ, Min(�ָ���־) �ָ���־
        From Zlclientparaset A
        Where a.������ = ������_In
        Group By ����վ;
    End If;
  
    If ����_In = 2 Then
      Open Cur_Out For
        Select Distinct Upper(�û���) �û���, Max(����վ) ����վ, Min(Decode(�ָ���־, 2, 0, �ָ���־)) �ָ���־
        From Zlclientparaset A
        Where a.������ = ������_In
        Group By �û���
        Order By �û���;
    End If;
  
  End Get_Resile;

  --���ܣ�ȡzldataMove����
  Procedure Get_Zldatamove
  (
    Cur_Out Out t_Refcur,
    ϵͳ_In In Zldatamove.ϵͳ%Type
  ) Is
  Begin
    Open Cur_Out For
      Select ���, ����, ˵��, �����ֶ�, ת������, �ϴ����� From zlDataMove Where ϵͳ = ϵͳ_In Order By ���;
  End Get_Zldatamove;

  --���ܣ�ȡ��־����
  Procedure Get_Log
  (
    Cur_Out     Out t_Refcur,
    ��־����_In In Varchar2,
    Where_In    In Varchar2
  ) Is
    v_Sql Varchar2(1000);
  Begin
    If ��־����_In = '������־' Then
      v_Sql := 'Select �Ự��,����վ,�û���,�������,������Ϣ,To_char(ʱ��,''yyyy-MM-dd hh24:mi:ss'') ʱ��
                     ,Decode(����,1,''�洢���̴���'',2,''������������'',3,''Ӧ�ó�������'',''�ͻ�����������'') ��������
                        From ZlErrorLog Where ' || Where_In;
      Open Cur_Out For v_Sql;
    End If;
    If ��־����_In = '������־' Then
      v_Sql := 'Select �Ự��,����վ,�û���,������,��������,To_char(����ʱ��,''yyyy-MM-dd hh24:mi:ss'') ����ʱ��
                                 ,To_char(�˳�ʱ��,''yyyy-MM-dd hh24:mi:ss'') �˳�ʱ��,Decode(�˳�ԭ��,1,''�����˳�'',''�쳣�˳�'') �˳�ԭ��
                                    From ZlDiaryLog Where ' || Where_In;
      Open Cur_Out For v_Sql;
    End If;
  End Get_Log;

  --���ܣ�ȡ��־��¼��
  Procedure Get_Log_Count
  (
    Cur_Out     Out t_Refcur,
    ��־����_In In Varchar2
  ) Is
  Begin
    If ��־����_In = '������־' Then
      Open Cur_Out For
        Select Count(*) ����
        From zlErrorLog
        Union All
        Select Nvl(To_Number(����ֵ), 0)
        From zlOptions
        Where ������ = 4;
    End If;
    If ��־����_In = '������־' Then
      Open Cur_Out For
        Select Count(*) ����
        From zlDiaryLog
        Union All
        Select Nvl(To_Number(����ֵ), 0)
        From zlOptions
        Where ������ = 2;
    
    End If;
  End Get_Log_Count;

  --���ܣ�ȡzlfilesupgradeg����
  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select ���, �ļ���, �汾��, �޸�����, �ļ�˵�� As ˵��,
             Decode(�ļ�����, 0, '��������', 1, 'Ӧ�ò���', 2, '�����ļ�', 3, '�����ļ�', 4, '��������', 5, 'ϵͳ�ļ�', '') As ����, ��װ·�� As ��װ·��,
             Md5 As Md5, ��������
      From zlFilesUpgrade
      Order By ���;
  End Get_Zlfilesupgrade;

  --���ܣ�ȡ��ע����Ŀ
  Procedure Get_Not_Regist(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select ��Ŀ, ����
      From zlRegInfo
      Where ��Ŀ Not In ('������', '�汾��', '������Ŀ¼', '�����û�', '��������', '�ռ�Ŀ¼', '�ռ�����', 'ע����', '��Ȩ֤��', '��Ȩ����', '��Ȩ�ʴ�');
  End Get_Not_Regist;

  --���ܣ�ȡ����ֵ
  Procedure Get_Zloption
  (
    Cur_Out   Out t_Refcur,
    ������_In In Zloptions.������%Type
  ) Is
  Begin
    Open Cur_Out For
      Select Nvl(����ֵ, ȱʡֵ) Option_Value From zlOptions Where ������ = ������_In;
  End Get_Zloption;

End b_Runmana;
/
--00000:��˶,2016-08-18,�������汾�����������µİ�ͷ����岻ƥ��
CREATE OR REPLACE Package b_Public Is
--��������
  Type t_Refcur Is Ref Cursor;
--���ܣ�ȡϵͳ����
--�����б�mdlMain.CurrentDate��clsDatabase.CurrentDate
  Procedure Get_Current_Date(Cursor_Out Out t_Refcur);
--���ܣ�ɾ��������־��������־
--�����б�mdlMain.DeleteAllLog
  Procedure Delete_All_Log(Runtimelog_In In Number := 0);
--���ܣ�ɾ����ǰ������־
--�����б�mdlMain.DeleteCurLog
  Procedure Delete_Diarylog
  (
    �Ự��_In   Number,
    �û���_In   Varchar2,
    ����վ_In   Varchar2,
    ������_In   Varchar2,
    ��������_In Varchar2,
    ����ʱ��_In Date
  );
--���ܣ�ɾ����ǰ������־
--�����б�mdlMain.DeleteCurLog
  Procedure Delete_Errorlog
  (
    �Ự��_In   Number,
    �û���_In   Varchar2,
    ����վ_In   Varchar2,
    ����_In     Number,
    �������_In Number,
    ʱ��_In     Date
  );
--���ܣ�ȡע����
--�����б�mdlMain.Getע����
  Procedure Get_Regcode(Cursor_Out Out t_Refcur);
--���ܣ�ȡ�汾��
--�����б�mdlMain.UpgradeManager
  Procedure Get_Ver(Cursor_Out Out t_Refcur);
--���ܣ����°汾��
--�����б�mdlMain.UpgradeManager
  Procedure Update_Ver(Verstring_In In Varchar2);
--���ܣ�ȡ��ϵͳ����������
--�����б�
--frmStatus.cmbsystem_Click��mdlMain.GetOwnerName��mdlMain.cmbSystem_Click
--frmAutoJobs.cmbSystem_Click��frmDataMove.cmbSystem_Click ��frmNoticeTools.cboSystem_Click
--frmProgPriv.ProgPriv��frmAppScript.cmbSystem_Click
  Procedure Get_Owner_Name
  (
    Cursor_Out Out t_Refcur,
    ���_In    In zlSystems.���%Type := 0
  );

--���ܣ�ȡע�������Ϣ
--�����б�
--frmAbout.GetUnitInfo��frmAutoJobs.From_load��frmClientsUpgrade.InitInfor
--frmFilesSet.ShowEdit��frmRegist.From_load��frmAppScript.From_Load
--frmFilesSendToServer.InitInfo
  Procedure Get_Reginfo
  (
    Cursor_Out Out t_Refcur,
    ��Ŀ_In    In zlRegInfo.��Ŀ%Type := Null
  );
--���ܣ�ȡzlGetSvrToolsg����
--�����б�frmMDIMain.MDIForm_Load
  Procedure Get_Zlsvrtools(Cursor_Out Out t_Refcur);
--���ܣ�ȡ�Ѱ�װϵͳ�嵥
--�����б�
--frmAppCheck.Form_Load��frmClearData.Form_Load��frmDataMove.Form_Load
--frmImp.FillSystem��frmLoadIn.FillSystem��frmLoadOut.FillSystem
--frmMDIMain.mnuFileRemove_Click��frmNoticeTools.Form_Activate��frmRoleGrant.FillSystem
--frmAppUpgrade.Form_Load��frmAppScript.Form_Load��frmExp.FillSystem
--frmInputTools.from_activate��fromRole.FillSystem��frmAutoJobs.From_load
--frmAppstart.sysCreated
  Procedure Get_Zlsystems
  (
    Cursor_Out Out t_Refcur,
    ������_In  In zlSystems.������%Type := Null
  );

End b_Public;
/
--84990:��˶,2015-05-28,��������
Create Or Replace Package Body Zltools.b_Public Is
  --���ܣ�ȡϵͳ����
  Procedure Get_Current_Date(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select Sysdate As ���� From Dual;
  End Get_Current_Date;

  --���ܣ�ɾ��������־��������־
  Procedure Delete_All_Log(Runtimelog_In In Number := 0) Is
    n_Count Number;
    n_Loop  Number;
  Begin
    If Runtimelog_In = 1 Then
      Select Count(����ʱ��) Into n_Count From zlDiaryLog;
      If n_Count > 1000 Then
        For n_Loop In 1 .. Ceil(n_Count - 1000) Loop
          Delete zlDiaryLog Where Rownum < 10001;
          Commit;
        End Loop;
      Else
        If n_Count > 0 Then
          Delete zlDiaryLog;
          Commit;
        End If;
      End If;
    Else
      Select Count(ʱ��) Into n_Count From zlErrorLog;
      If n_Count > 1000 Then
        For n_Loop In 1 .. Ceil(n_Count - 1000) Loop
          Delete zlErrorLog Where Rownum < 10001;
          Commit;
        End Loop;
      Else
        If n_Count > 0 Then
          Delete zlErrorLog;
          Commit;
        End If;
      End If;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Delete_All_Log;

  --���ܣ�ɾ����ǰ������־
  Procedure Delete_Diarylog
  (
    �Ự��_In   Number,
    �û���_In   Varchar2,
    ����վ_In   Varchar2,
    ������_In   Varchar2,
    ��������_In Varchar2,
    ����ʱ��_In Date
  ) Is
  Begin
    Delete zlDiaryLog
    Where �Ự�� = �Ự��_In And �û��� = �û���_In And ����վ = ����վ_In And ������ = ������_In And �������� = ��������_In And ����ʱ�� = ����ʱ��_In;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Delete_Diarylog;

  --���ܣ�ɾ����ǰ������־
  Procedure Delete_Errorlog
  (
    �Ự��_In   Number,
    �û���_In   Varchar2,
    ����վ_In   Varchar2,
    ����_In     Number,
    �������_In Number,
    ʱ��_In     Date
  ) Is
  Begin
    Delete zlErrorLog
    Where �Ự�� = �Ự��_In And �û��� = �û���_In And ����վ = ����վ_In And ���� = ����_In And ������� = �������_In And ʱ�� = ʱ��_In;
    Commit;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Delete_Errorlog;

  --���ܣ�ȡע����
  Procedure Get_Regcode(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select ���� From zlRegInfo Where ��Ŀ = 'ע����' Or ��Ŀ = '��Ȩ֤��' Order By �к�;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Regcode;

  --���ܣ�ȡ�汾��
  Procedure Get_Ver(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select ���� From zlRegInfo Where ��Ŀ = '�汾��';
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Ver;

  --���ܣ����°汾��
  Procedure Update_Ver(Verstring_In In Varchar2) Is
  Begin
    Update zlRegInfo Set ���� = Verstring_In Where ��Ŀ = '�汾��';
    If Sql%NotFound Then
      Insert Into zlRegInfo (��Ŀ, �к�, ����) Values ('�汾��', 1, Verstring_In);
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Update_Ver;

  --���ܣ�ȡ��ϵͳ����������
  Procedure Get_Owner_Name
  (
    Cursor_Out Out t_Refcur,
    ���_In    In Zlsystems.���%Type := 0
  ) Is
  Begin
    Open Cursor_Out For
      Select Upper(������) As ������ From zlSystems Where ��� = ���_In;
  End Get_Owner_Name;

  --���ܣ�ȡע�������Ϣ
  Procedure Get_Reginfo
  (
    Cursor_Out Out t_Refcur,
    ��Ŀ_In    In Zlreginfo.��Ŀ%Type := Null
  ) Is
  Begin
    If Trim(Nvl(��Ŀ_In, '��')) = '��' Then
      Open Cursor_Out For
        Select * From zlRegInfo;
    Else
      Open Cursor_Out For
        Select ���� From zlRegInfo Where ��Ŀ = ��Ŀ_In Order By �к�;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Reginfo;

  --���ܣ�ȡzlGetSvrToolsg����
  Procedure Get_Zlsvrtools(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select * From zlSvrTools Start With �ϼ� Is Null Connect By Prior ��� = �ϼ� Order By Level, ���;
  End Get_Zlsvrtools;

  --���ܣ�ȡ�Ѱ�װϵͳ�嵥
  Procedure Get_Zlsystems
  (
    Cursor_Out Out t_Refcur,
    ������_In  In Zlsystems.������%Type := Null
  ) Is
  Begin
    If Nvl(������_In, '��') = '��' Then
      Open Cursor_Out For
        Select ���, ����, �����, Upper(������) ������, ��װ����, ������װ, �汾�� From zlSystems Order By ���;
    Else
      Open Cursor_Out For
        Select ���, ����, �����, Upper(������) ������, ��װ����, ������װ, �汾��
        From zlSystems
        Where Upper(������) = Upper(������_In)
        Order By ���;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlsystems;

End b_Public;
/