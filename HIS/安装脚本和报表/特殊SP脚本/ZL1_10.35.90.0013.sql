----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--126035:����,2018-05-23,����Zlmenus��©�������
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, 'Ԥ��������ձ�', 'Ԥ��������ձ�', Null, 105, '����ͳ�Ʋ���Ա�Ĺ�������Ʊ��ʹ�ú��տ������', 100, 1104 From zlMenus Where ���� = 'סԺ���ת����ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '��Լ��λ����', '��Լ��λ����', Null, 105, '����Լ��λͳ���䲡�˵Ļ��ܷ��������Ƿ�������', 100, 1105 From zlMenus Where ���� = 'סԺ���ת����ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '�Һ�Ա�����ձ�', '�Һ�Ա�����ձ�', Null, 105, '����ͳ�ƹҺ�Ա�Ĺ�������Ʊ��ʹ�úͷ���Һ������', 100, 1112 From zlMenus Where ���� = '�ż���Һ�ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '�����ۺ�ͳ�Ʒ���', '�����ۺ�ͳ�Ʒ���', Null, 105, '�������������з���ͳ�Ʒ�������ı걾�˴Ρ���Ŀ�˴μ����õȡ�', 100, 1230 From zlMenus Where ϵͳ = 100 And ��� = 'ȱʡ' And ���� = '������Ϣϵͳ' And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, 'ѧ��ͳ��', 'ѧ��ͳ��', Null, 105, 'ָ����������Ŀ��ʱ�䷶Χ����������ѯ����������׼������ʵ�ͳ�����ݡ�', 100, 1231 From zlMenus Where ϵͳ = 100 And ��� = 'ȱʡ' And ���� = '������Ϣϵͳ' And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '������ҩ����ѯ', '������ҩ����ѯ', Null, 105, '�����ض�ҩ������ҩ���н顢���в�ѯ��', 100, 1232 From zlMenus Where ϵͳ = 100 And ��� = 'ȱʡ' And ���� = '������Ϣϵͳ' And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, 'ϸ��ҩ����ѯ', 'ϸ��ҩ����ѯ', Null, 105, 'ϸ���Կ����ص���ҩ���н顢���в�ѯ��', 100, 1233 From zlMenus Where ϵͳ = 100 And ��� = 'ȱʡ' And ���� = '������Ϣϵͳ' And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, 'ϸ��ҩ��ֲ�ͳ��', 'ϸ��ҩ��ֲ�ͳ��', Null, 105, 'ָ�������£�ָ��ĳϸ����ָ��������ҩ�����Խ����', 100, 1234 From zlMenus Where ϵͳ = 100 And ��� = 'ȱʡ' And ���� = '������Ϣϵͳ' And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '�����Լ�����ͳ��', '�����Լ�����ͳ��', Null, 105, 'ͳ��һ��ʱ���ڵļ����Լ����������', 100, 1235 From zlMenus Where ϵͳ = 100 And ��� = 'ȱʡ' And ���� = '������Ϣϵͳ' And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '������д���', '������д���', Null, 105, '��ѯ��סԺ���Ҳ�����д���', 100, 1279 From zlMenus Where ϵͳ = 100 And ��� = 'ȱʡ' And ���� = '�����ʿ�������ϵͳ' And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '�⹺ҩƷ���ܱ�', '�⹺ҩƷ����(��λ)', Null, 105, '����Ӧ�̻�ҩƷ��������⹺ҩƷ���ݣ��Թ���ѯ��', 100, 1312 From zlMenus Where ϵͳ = 100 And ��� = 'ȱʡ' And ���� = 'סԺҩ������ϵͳ' And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '����ҩƷ���ܱ�', '����ҩƷ���ܱ�', Null, 105, '��ҩƷ�����������ҩƷ���ݣ��Թ���ѯ��', 100, 1313 From zlMenus Where ϵͳ = 100 And ��� = 'ȱʡ' And ���� = 'סԺҩ������ϵͳ' And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '�������û��ܱ�', '���ŷ������', Null, 105, '���ܸ�����ҩ����һ��ʱ�������ҩƷ�����ݡ�', 100, 1314 From zlMenus Where ϵͳ = 100 And ��� = 'ȱʡ' And ���� = 'סԺҩ������ϵͳ' And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, 'ҩƷ�ƿ���ܱ�', '�Ƴ�����ͳ��', Null, 105, '��ӳһ��ʱ�����ҩƷ�ⷿ��ҩƷת�ƵĻ��������', 100, 1315 From zlMenus Where ϵͳ = 100 And ��� = 'ȱʡ' And ���� = 'סԺҩ������ϵͳ' And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, 'ҩƷ���ۻ��ܱ�', 'ҩƷ���ۻ��ܱ�', Null, 105, '��ѯҩƷ��һ��ʱ���ڵ��۱䶯���������', 100, 1316 From zlMenus Where ϵͳ = 100 And ��� = 'ȱʡ' And ���� = 'סԺҩ������ϵͳ' And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '�ƻ������', '�ƻ������', Null, 105, '��Ҫ��ѯ��Ʒ�ڼƻ���δ�����Ѹ�����Ϣ��', 100, 1325 From zlMenus Where ���� = 'ҩ�������ҩƷ���ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, 'ҩ������������', 'ҩ������������', Null, 105, '��ӳҩ����Ա�Ĺ����������', 100, 1346 From zlMenus Where ���� = 'סԺҩ������ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, 'ҽʦ��������ͳ�Ʊ�', 'ҽʦ��������ͳ�Ʊ�', Null, 105, 'ҽʦ��������ͳ�Ʊ�', 100, 1570 From zlMenus Where ϵͳ = 100 And ��� = 'ȱʡ' And ���� = '�����ʿ�������ϵͳ' And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '���Ҳ�������ͳ�Ʊ�', '���Ҳ�������ͳ�Ʊ�', Null, 105, '���Ҳ�������ͳ�Ʊ�', 100, 1571 From zlMenus Where ϵͳ = 100 And ��� = 'ȱʡ' And ���� = '�����ʿ�������ϵͳ' And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '�������ֽ���嵥', '�������ֽ���嵥', Null, 105, '�������ֽ���嵥', 100, 1572 From zlMenus Where ϵͳ = 100 And ��� = 'ȱʡ' And ���� = '�����ʿ�������ϵͳ' And ģ�� Is Null;
Insert into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Select 'ȱʡ',zlMenus_ID.Nextval,ID,'ҽ��ͳ�Ʊ���','סԺ���û��ܱ�',Null,105,'ҽ��ͳ�Ʊ���',100,1610 From zlMenus Where ϵͳ=100 And ���='ȱʡ' And ����='ҽ��֧��ϵͳ' And ģ�� is NULL;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '�⹺���Ļ��ܱ�', '�⹺������Դ�嵥', Null, 105, '����Ӧ�̻�������������⹺�������ݣ��Թ���ѯ��', 100, 1730 From zlMenus Where ���� = '�������Ϲ���ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '�������û��ܱ�', '���ŷ������', Null, 105, '���ܸ������ϲ���һ��ʱ��������������ϵ����ݡ�', 100, 1731 From zlMenus Where ���� = '�������Ϲ���ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '�����ƿ���ܱ�', '�Ƴ�����ͳ��', Null, 105, '��ӳһ��ʱ������������Ͽⷿ����������ת�ƵĻ��������', 100, 1732 From zlMenus Where ���� = '�������Ϲ���ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '����ֱ����֧���������ţ�', '����ֱ����֧����(����)', Null, 105, '����ֱ����֧����(����)', 100, 1733 From zlMenus Where ���� = '�������Ϲ���ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '�����������', '�ض�����ȥ��(����)', Null, 105, '�����Ż�ҽ������ָ������������ĵĲ�ͬȥ��', 100, 1734 From zlMenus Where ���� = '�������Ϲ���ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '���ϲ��Ź���������', '���ϲ��Ź���������', Null, 105, '��ӳ������Ա�Ĺ����������', 100, 1735 From zlMenus Where ���� = '�������Ϲ���ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '���ĳ�����ȱ����', '���ĳ�����ȱ����', Null, 105, '�������ĵĴ洢���޺����ޣ���ѯ������ĵĳ�����ȱ��Ϣ��', 100, 1736 From zlMenus Where ���� = '�������Ϲ���ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '���ĵ��ۻ��ܱ�', '���ĵ��ۻ��ܱ�', Null, 105, '��ѯ������һ��ʱ���ڵ��۱䶯���������', 100, 1737 From zlMenus Where ���� = '�������Ϲ���ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '����������ܱ�', '����������ܱ�', Null, 105, '�������������ҵ����ܸ������Ŀⷿ��ͬ���ĵĿ��仯���ݡ�', 100, 1738 From zlMenus Where ���� = '�������Ϲ���ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;
Insert into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Select 'ȱʡ',zlMenus_ID.Nextval,ID,'�����շ�����ܱ�','�����շ�����ܱ�',Null,105,'�����ķ���������ĵ��շ������Լ������Ͳ�ۡ�',100,1739 From zlMenus Where ϵͳ=100 And ���='ȱʡ' And ����='�������Ϲ���ϵͳ' And ģ�� is NULL;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '����Ч�ڱ�������', '����Ч�ڱ�������', Null, 105, '��ѯ�ڽ��һ��ʱ���ڽ�ʧЧ�����ĵĵ�ǰ��棬�Ա㼰ʱ����', 100, 1740 From zlMenus Where ���� = '�������Ϲ���ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;
Insert Into zlMenus(���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��) Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '�������ñ�������', '�������ñ�������', Null, 105, '�˽��ĳ��ʱ��������һֱû��ʹ�õĻ�ѹ���ġ�', 100, 1741 From zlMenus Where ���� = '�������Ϲ���ϵͳ' And ��� = 'ȱʡ' And ϵͳ = 100 And ģ�� Is Null;


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
--123845:����,2018-05-25,���������ӿ�Zl_Third_Getdepositbalance����ȡ���˿���Ԥ�����
Create Or Replace Procedure Zl_Third_Getdepositbalance
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  --------------------------------------------------------------------------------------------------
  --����:��ȡ����Ԥ�������
  --���:Xml_In:
  --    <IN>
  --        <BRID>����ID</BRID>
  --        <ZYID>��ҳID</ZYID> //סԺԤ����ѯʱ��Ч:������ҳID����ѯ�ڼ��ε�Ԥ�����
  --        <YJLX>Ԥ������</YJLX> //1-����Ԥ��;2-סԺԤ��;0-����Ԥ��
  --              ˵��:���Ԥ������û�д���,��ȱʡΪ0,��ȡ�����סԺԤ��
  --    </IN>
  --����:Xml_Out
  --  <OUTPUT>
  --     <YJYE>Ԥ�����</YJYE>
  --     �D�D�������д�������˵����ȷִ��
  --    <ERROR>
  --      <MSG>������Ϣ</MSG>
  --    </ERROR>
  --  </OUTPUT>
  --------------------------------------------------------------------------------------------------
  n_����id   ������Ϣ.����id%Type;
  n_��ҳid   ������ҳ.��ҳid%Type;
  n_Ԥ������ ����Ԥ����¼.Ԥ�����%Type;
  n_Ԥ����� �������.Ԥ�����%Type;
  n_������� �������.�������%Type;
  v_Temp     Varchar2(32767); --��ʱXML
  x_Templet  Xmltype; --ģ��XML
  v_Err_Msg  Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/BRID'), Extractvalue(Value(A), 'IN/ZYID'), Extractvalue(Value(A), 'IN/YJLX')
  Into n_����id, n_��ҳid, n_Ԥ������
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := 'δ���벡��ID,�޷���ɲ�ѯ!';
    Raise Err_Item;
  End If;

  If Nvl(n_Ԥ������, 0) = 1 Then
    Select Nvl(Sum(Ԥ�����), 0) - Nvl(Sum(�������), 0)
    Into n_Ԥ�����
    From �������
    Where ����id = n_����id And ���� = 1;
  Elsif Nvl(n_Ԥ������, 0) = 2 Then
    If Nvl(n_��ҳid, 0) = 0 Then
      Select Nvl(Sum(Ԥ�����), 0) - Nvl(Sum(�������), 0)
      Into n_Ԥ�����
      From �������
      Where ����id = n_����id And ���� = 2;
    Else
      Select Nvl(Sum(���), 0) - Nvl(Sum(��Ԥ��), 0)
      Into n_Ԥ�����
      From ����Ԥ����¼
      Where ����id = n_����id And ��ҳid = n_��ҳid And ��¼���� In (1, 11);
      Select Nvl(Sum(���), 0) Into n_������� From ����δ����� Where ����id = n_����id And ��ҳid = n_��ҳid;
      n_Ԥ����� := n_Ԥ����� - n_�������;
    End If;
  Else
    Select Nvl(Sum(Ԥ�����), 0) - Nvl(Sum(�������), 0) Into n_Ԥ����� From ������� Where ����id = n_����id;
  End If;
  If Nvl(n_Ԥ�����, 0) < 0 Then
    n_Ԥ����� := 0;
  End If;
  v_Temp := '<YJYE>' || n_Ԥ����� || '</YJYE>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getdepositbalance;
/

--125867:�ƽ�,2018-05-23,RIS�ӿڳ�Ժ������δ�ɷ��ò�����ִ�з���

Create Or Replace Package b_Zlxwinterface Is
  Type t_Refcur Is Ref Cursor;

  --1������RIS״̬�ı�
  Procedure Receiverisstate
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ״̬_In     Number,
    ������Ա_In ����ҽ������.�����%Type,
    ִ��ʱ��_In ����ҽ������.���ʱ��%Type := Null,
    ִ��˵��_In ����ҽ������.ִ��˵��%Type := Null,
    ����ִ��_In Number := 0
  );

  --2������ȷ��
  Procedure Ӱ�����ִ��
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  );

  --3��ȡ������ȷ��
  Procedure Ӱ�����ִ��_Cancel
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  );

  --4������RIS�ı���
  Procedure Receivereport
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ��������_In ���Ӳ�������.�����ı�%Type,
    �������_In ���Ӳ�������.�����ı�%Type,
    ���潨��_In ���Ӳ�������.�����ı�%Type,
    ����ҽ��_In ���Ӳ�����¼.������%Type
  );

  --5���޸����뵥��Ϣ
  Procedure Ӱ������Ϣ_�޸�
  (
    ҽ��id_In       ����ҽ����¼.Id%Type,
    ����_In         ������Ϣ.����%Type,
    �Ա�_In         ������Ϣ.�Ա�%Type,
    ����_In         ������Ϣ.����%Type,
    �ѱ�_In         ������Ϣ.�ѱ�%Type,
    ҽ�Ƹ��ʽ_In ������Ϣ.ҽ�Ƹ��ʽ%Type,
    ����_In         ������Ϣ.����%Type,
    ����_In         ������Ϣ.����״��%Type,
    ְҵ_In         ������Ϣ.ְҵ%Type,
    ���֤��_In     ������Ϣ.���֤��%Type,
    ��ͥ��ַ_In     ������Ϣ.��ͥ��ַ%Type,
    ��ͥ�绰_In     ������Ϣ.��ͥ�绰%Type,
    ��ͥ��ַ�ʱ�_In ������Ϣ.��ͥ��ַ�ʱ�%Type,
    ��������_In     ������Ϣ.��������%Type := Null
  );

  --6��ȡ�����뵥��Ϣ
  Procedure ȡ��������뵥
  (
    ҽ��id_In     ����ҽ��ִ��.ҽ��id%Type,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := 0,
    �ܾ�ԭ��_In   ����ҽ������.ִ��˵��%Type := Null
  );

  --7������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_Insert
  (
    ������Դ_In   In Risҽ��ʧ�ܼ�¼.������Դ%Type,
    ����id_In     In Risҽ��ʧ�ܼ�¼.����id%Type,
    ��ҳid_In     In Risҽ��ʧ�ܼ�¼.��ҳid%Type,
    �Һŵ���_In   In Risҽ��ʧ�ܼ�¼.�Һŵ���%Type,
    ���ͺ�_In     In Risҽ��ʧ�ܼ�¼.���ͺ�%Type,
    �������id_In In Risҽ��ʧ�ܼ�¼.�������id%Type,
    ��챨����_In In Risҽ��ʧ�ܼ�¼.��챨����%Type,
    ��������_In   In Risҽ��ʧ�ܼ�¼.��������%Type
  );

  --8������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_�ط�
  (
    Id_In       In Risҽ��ʧ�ܼ�¼.Id%Type,
    ��������_In In Number
  );

  --9�����˺��½�סԺ���˵���
  Procedure ����ҽ��_�ؽ�����
  (
    ҽ��id_In In ����ҽ������.ҽ��id%Type,
    No_In     In ����ҽ������.No%Type,
    Action_In In Number
  );

  --10����ӡRIS���ԤԼ֪ͨ��
  Procedure Ris���ԤԼ_��ӡ(ҽ��id_In In Ris���ԤԼ.ҽ��id%Type);

  --11������RIS�ֿ������ò���
  Procedure Ris���ÿ���_Update
  (
    �������_In Ris���ÿ���.�������%Type,
    ����_In     Ris���ÿ���.����%Type,
    ����ids_In  Varchar2,
    ��������_In Number
  );

  --12��ɾ��RIS�ֿ������ò���
  Procedure Ris���ÿ���_Delete;

  --13������Ԫ������ȡ��Ϣ
  Function Ris_Replace_Element_Value
  (
    Ԫ����_In   In ����������Ŀ.������%Type,
    ����id_In   In ���Ӳ�����¼.����id%Type,
    ����id_In   In ���Ӳ�����¼.��ҳid%Type,
    ������Դ_In In ���Ӳ�����¼.������Դ%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type
  ) Return Varchar2;

  --14��ɾ��RIS��Ժ���ò���
  Procedure Ris��Ժ����_Delete;

  --15������RISRis��Ժ���ò���
  Procedure Ris��Ժ����_Update
  (
    Id_In           Ris��Ժ����.Id%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    �û���_In       Ris��Ժ����.�û���%Type,
    ����_In         Ris��Ժ����.����%Type,
    ���ݿ������_In Ris��Ժ����.���ݿ������%Type
  );
End b_Zlxwinterface;
/

Create Or Replace Package Body b_Zlxwinterface Is

  --1������RIS״̬�ı�
  Procedure Receiverisstate
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ״̬_In     Number,
    ������Ա_In ����ҽ������.�����%Type,
    ִ��ʱ��_In ����ҽ������.���ʱ��%Type := Null,
    ִ��˵��_In ����ҽ������.ִ��˵��%Type := Null,
    ����ִ��_In Number := 0
  ) Is
  
    --������ҽ��ID_IN - ����ִ�е�ҽ��ID��
    --      ״̬_IN - -1-ɾ����0-ԤԼ��1-�Ǽǣ�3-�����ɣ�4-�����ֹ��9-�������棻12-������ˣ�15-����
    --     ����ִ��_In -0-ȫ��ִ�У�1-����ִ�У����ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ
  
    Cursor c_Adviceinfo Is
      Select a.Id, a.���id, Nvl(a.���id, a.Id) As ��id, a.�������, a.������Դ, a.ִ�п���id, b.ִ�й���
      From ����ҽ����¼ A, ����ҽ������ B
      Where a.Id = b.ҽ��id And ID = ҽ��id_In;
    r_Adviceinfo c_Adviceinfo%RowType;
  
    v_ִ��״̬ ����ҽ������.ִ��״̬%Type;
    v_ִ�й��� ����ҽ������.ִ�й���%Type;
    n_ִ��     Number; --����Ƿ���Ҫ����״̬��1����Ҫ���£���������Ҫ����
    v_Count    Number;
    v_�����   ����ҽ������.�����%Type;
    v_���ʱ�� ����ҽ������.���ʱ��%Type;
    v_Error    Varchar2(255);
    Err_Custom Exception;
  
  Begin
  
    v_ִ��״̬ := 0;
    v_ִ�й��� := 0;
  
    --��ȡҽ������ҽ��ID������ID
    Open c_Adviceinfo;
    Fetch c_Adviceinfo
      Into r_Adviceinfo;
    Close c_Adviceinfo;
  
    --����״̬_INִ��ҽ��
    ---1-ɾ����0-ԤԼ��1-�Ǽǣ�3-�����ɣ�4-�����ֹ��9-�������棻12-������ˣ�13-ȡ����ˣ�14-����ɾ����15-����
  
    If ״̬_In = -1 Or ״̬_In = 0 Then
      v_ִ��״̬ := 0; --δִ��
      v_ִ�й��� := 0;
    Elsif ״̬_In = 1 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 2; --�ѱ���
    Elsif ״̬_In = 3 Or ״̬_In = 14 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 3; --�Ѽ��
    Elsif ״̬_In = 4 Then
      --���ı�
      v_ִ��״̬ := v_ִ��״̬;
    Elsif ״̬_In = 9 Or ״̬_In = 13 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 4; --�ѱ���
    Elsif ״̬_In = 12 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 5; --�����
    Elsif ״̬_In = 15 Then
      v_ִ��״̬ := 1; --��ȫִ��
      v_ִ�й��� := 6; --�����
      v_�����   := ������Ա_In;
      v_���ʱ�� := ִ��ʱ��_In;
    End If;
  
    n_ִ�� := 1; --Ĭ�϶�Ҫ����״̬
  
    If ״̬_In = 13 Or ״̬_In = 14 Then
      --ɾ����Ӧ��������
      Delete From ���Ӳ�����¼
      Where ID = (Select ����id From ����ҽ������ Where ҽ��id = ҽ��id_In And Risid = Risid_In);
      Delete From ����ҽ������ Where ҽ��id = ҽ��id_In And Risid = Risid_In;
    
      --ɾ�����ж��Ƿ񻹴��ڱ��棬��������ҽ��״̬���ֲ��䣬������ȫ��ɾ�������ҽ��״̬
      Select Count(1) Into v_Count From ����ҽ������ Where ҽ��id = ҽ��id_In;
    
      If v_Count > 0 Then
        n_ִ�� := 0; --��������ҽ��״̬���ֲ���
      End If;
    End If;
  
    --����ǵǼǣ����жϴ˼���Ƿ�δִ��
    If ״̬_In = 1 Then
      If r_Adviceinfo.ִ�й��� >= 3 Then
        v_Error := '�����Ѿ���������ˣ������ظ��Ǽǡ�';
        Raise Err_Custom;
      End If;
    End If;
  
    --��ʼִ��ҽ��
    If n_ִ�� = 1 Then
      If Nvl(����ִ��_In, 0) = 1 Then
        -- ������λҽ������ִ��
        Update ����ҽ������
        Set ִ��״̬ = v_ִ��״̬, ִ�й��� = v_ִ�й���, ִ��˵�� = ִ��˵��_In, ����� = v_�����, ���ʱ�� = v_���ʱ��
        Where ҽ��id = ҽ��id_In;
      Else
        Update ����ҽ������
        Set ִ��״̬ = v_ִ��״̬, ִ�й��� = v_ִ�й���, ִ��˵�� = ִ��˵��_In, ����� = v_�����, ���ʱ�� = v_���ʱ��
        Where ҽ��id In (Select ID From ����ҽ����¼ Where (ID = r_Adviceinfo.��id Or ���id = r_Adviceinfo.��id));
      End If;
    End If;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Receiverisstate;

  --2������ȷ��
  Procedure Ӱ�����ִ��
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  ) Is
    --������ҽ��ID_IN=����ִ�е�ҽ��ID��
    --      ����ִ��_In=���ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ,0-������ִ��
    Cursor c_Advice Is
      Select ID, ���id, Nvl(���id, ID) As ��id, �������, ������Դ From ����ҽ����¼ Where ID = ҽ��id_In;
    r_Advice c_Advice%RowType;
  
    v_Temp     Varchar2(255);
    v_��Ա��� ��Ա��.���%Type;
    v_��Ա���� ��Ա��.����%Type;
    v_����id   ���ű�.Id%Type;
    v_�������� ����ҽ������.��¼����%Type;
    v_���ͺ�   ����ҽ������.���ͺ�%Type;
    v_ִ�й��� ����ҽ������.ִ�й���%Type;
    v_Error    Varchar2(255);
    Err_Custom Exception;
  Begin
  
    --ȡ��ҽ��ID
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
    Close c_Advice;
  
    Select ���ͺ�, ִ�й��� Into v_���ͺ�, v_ִ�й��� From ����ҽ������ Where ҽ��id = r_Advice.��id;
  
    --�ǼǺ���ɲ�ִ�з���  2-�Ǽǣ�3-��飬4-���棬5-��ˣ�6-���
    If v_ִ�й��� >= 2 Or v_ִ�й��� <= 6 Then
      --ȡ��ǰ������Ա
      If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null And ִ�в���id_In Is Not Null Then
        v_��Ա��� := ����Ա���_In;
        v_��Ա���� := ����Ա����_In;
        v_����id   := ִ�в���id_In;
      Else
        v_Temp     := Zl_Identity;
        v_����id   := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      End If;
    
      If r_Advice.������Դ = 2 Then
        Select Decode(��¼����, 1, 1, Decode(�������, 1, 1, 2))
        Into v_��������
        From ����ҽ������
        Where ���ͺ� = v_���ͺ� And ҽ��id = ҽ��id_In;
      Else
        v_�������� := 1;
      End If;
    
      --ִ�з��ú��Զ�����
      If v_�������� = 1 Then
        Zl_����ҽ��ִ��_Finish(ҽ��id_In, v_���ͺ�, ����ִ��_In, v_��Ա���, v_��Ա����, r_Advice.��id, r_Advice.�������, v_����id);
      Else
        Zl_סԺҽ��ִ��_Finish(ҽ��id_In, v_���ͺ�, ����ִ��_In, v_��Ա���, v_��Ա����, r_Advice.��id, r_Advice.�������, v_����id);
      End If;
    End If;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ӱ�����ִ��;

  --3��ȡ������ȷ��
  Procedure Ӱ�����ִ��_Cancel
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  ) Is
    --������
    --      ҽ��ID_IN=����ִ�е�ҽ��ID��
    --      ����ִ��_In=���ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ,0-������ִ��
  
    Cursor c_Advice Is
      Select ID, ���id, Nvl(���id, ID) As ��id From ����ҽ����¼ Where ID = ҽ��id_In;
    r_Advice c_Advice%RowType;
  
    v_���ͺ� ����ҽ������.���ͺ�%Type;
    v_Count  Number;
    v_Error  Varchar2(255);
    Err_Custom Exception;
  
  Begin
  
    --ȡ��ҽ��ID
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
    Close c_Advice;
  
    --�ȼ���Ƿ��Ѿ���Ժ��סԺ���ˣ��Ѿ�Ԥ��Ժ���߳�Ժ�ļ�����룬����ִ�з���
    Select Count(*)
    Into v_Count
    From ����ҽ����¼ A, ������ҳ B
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And (b.��Ժ���� Is Not Null Or b.״̬ = 3) And a.Id = r_Advice.��id;
  
    If v_Count > 0 Then
      v_Error := 'סԺ�����Ѿ���Ժ����Ԥ��Ժ������ȡ�����á�';
      Raise Err_Custom;
    End If;
  
    Select ���ͺ� Into v_���ͺ� From ����ҽ������ Where ҽ��id = r_Advice.��id;
  
    --����ͳһ��ҽ��ִ��Cancel����
    Zl_����ҽ��ִ��_Cancel(ҽ��id_In, v_���ͺ�, Null, ����ִ��_In, ִ�в���id_In, ����Ա���_In, ����Ա����_In);
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ӱ�����ִ��_Cancel;

  --4������RIS�ı���
  Procedure Receivereport
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ��������_In ���Ӳ�������.�����ı�%Type,
    �������_In ���Ӳ�������.�����ı�%Type,
    ���潨��_In ���Ӳ�������.�����ı�%Type,
    ����ҽ��_In ���Ӳ�����¼.������%Type
  ) Is
    --��ȡ����ҽ��������������Ϣ
    Cursor c_Advice
    (
      v_��id  Number,
      v_Risid Number
    ) Is
      Select e.Id, e.������Դ, e.����id, e.��ҳid, e.Ӥ��, e.���˿���id, e.�ļ�id, e.��������, e.��������, f.����id, e.ִ�п���id
      From (Select c.Id, c.������Դ, c.����id, c.��ҳid, c.Ӥ��, c.���˿���id, c.�ļ�id, d.���� ��������, d.���� ��������, c.ִ�п���id
             From (Select a.Id, a.������Դ, a.����id, a.��ҳid, a.Ӥ��, a.���˿���id, b.�����ļ�id �ļ�id, a.ִ�п���id
                    From ����ҽ����¼ A, ��������Ӧ�� B
                    Where a.Id = v_��id And a.������Ŀid = b.������Ŀid(+) And b.Ӧ�ó���(+) = Decode(a.������Դ, 2, 2, 4, 4, 1)) C,
                  �����ļ��б� D
             Where c.�ļ�id = d.Id(+)) E, ����ҽ������ F
      Where e.Id = f.ҽ��id(+) And f.Risid(+) = v_Risid;
  
    --�����ļ������Ԫ��
    Cursor c_File(v_File Number) Is
      Select a.Id, a.�ļ�id, a.��id, a.�������, a.��������, a.������, a.��������, a.��������, a.�����д�, a.�����ı�, a.�Ƿ���, a.Ԥ�����id, a.�������,
             a.ʹ��ʱ��, a.����Ҫ��id, a.�滻��, a.Ҫ������, a.Ҫ������, a.Ҫ�س���, a.Ҫ��С��, a.Ҫ�ص�λ, a.Ҫ�ر�ʾ, a.������̬, a.Ҫ��ֵ��
      From �����ļ��ṹ A
      Where a.�ļ�id = v_File
      Order By a.�������;
  
    Cursor c_Report(v_���Ӳ�����¼id Number) Is
      Select b.Id, a.�����ı�
      From ���Ӳ������� A, ���Ӳ������� B
      Where a.�������� = 3 And a.Id = b.��id And b.�������� = 2 And b.��ֹ�� = 0 And a.�ļ�id = v_���Ӳ�����¼id;
  
    Cursor c_Content
    (
      v_�ļ�id Number,
      v_���id Number
    ) Is
      Select a.Id, a.�ļ�id, a.��id, a.�������, a.��������, a.������, a.��������, a.��������, a.�����д�, a.�����ı�, a.�Ƿ���, a.Ԥ�����id, a.�������,
             a.ʹ��ʱ��, a.����Ҫ��id, a.�滻��, a.Ҫ������, a.Ҫ������, a.Ҫ�س���, a.Ҫ��С��, a.Ҫ�ص�λ, a.Ҫ�ر�ʾ, a.������̬, a.Ҫ��ֵ��
      From �����ļ��ṹ A
      Where �ļ�id = v_�ļ�id And ��id = v_���id;
  
    r_Advice        c_Advice%RowType;
    v_����id        ���Ӳ�������.�ļ�id%Type;
    v_��������id    ���Ӳ�������.Id%Type;
    v_��������idnew ���Ӳ�������.Id%Type;
    v_�������      ���Ӳ�������.�������%Type;
    v_��id          ���Ӳ�������.��id%Type;
    v_�����ı�      ���Ӳ�������.�����ı�%Type;
    v_�������id    ���Ӳ�������.�������id%Type;
    --v_��ʽ����    ���Ӳ�����ʽ.����%Type;
    v_Error Varchar2(255);
    Err_Custom Exception;
    v_��ҽ��id ����ҽ������.ҽ��id%Type;
    v_���     Varchar2(300);
    n_����     Number;
    n_Rptcount Number;
    v_�������� ���Ӳ�����¼.��������%Type;
    v_�Һŵ�id ���˹Һż�¼.Id%Type;
  
    Function Getrptno
    (
      v_ҽ��idin   ����ҽ������.ҽ��id%Type,
      v_��������in ���Ӳ�����¼.��������%Type
    ) Return Varchar As
      v_Return Number;
      v_No     Number;
      v_Count  Number;
    Begin
      Select Count(ҽ��id) + 1 Into v_No From ����ҽ������ Where ҽ��id = v_ҽ��idin;
      v_Count := 1;
      While v_Count = 1 Loop
        Select Count(ID)
        Into v_Count
        From ����ҽ������ A, ���Ӳ�����¼ B
        Where a.ҽ��id = v_ҽ��idin And a.����id = b.Id And b.�������� = v_��������in || v_No;
        If v_Count = 1 Then
          v_No := v_No + 1;
        End If;
      End Loop;
      v_Return := v_No;
      Return v_Return;
    End Getrptno;
  
  Begin
  
    -- ��ȡ��ҽ��ID ����ֹ��Ϊ���벿λҽ�������±��汣�����
    Select Nvl(���id, ID) As ��id Into v_��ҽ��id From ����ҽ����¼ Where ID = ҽ��id_In;
  
    Open c_Advice(v_��ҽ��id, Nvl(Risid_In, 0));
    Fetch c_Advice
      Into r_Advice;
  
    If Nvl(r_Advice.�ļ�id, 0) = 0 Then
      v_Error := '���μ����Ŀû�ж�Ӧ��صļ�鱨�棬�������Ա��ϵ��';
      Raise Err_Custom;
    Else
      If Nvl(r_Advice.����id, 0) > 0 Then
        ----����������
        --�ҳ��������д�ı�������к���"%����%","%����%","%����%","%���%",���ô���Ĳ�������
        For r_Report In c_Report(r_Advice.����id) Loop
          If r_Report.�����ı� Like '%����%' Then
            Update ���Ӳ������� Set �����ı� = ��������_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          Elsif r_Report.�����ı� Like '%���%' Then
            Update ���Ӳ������� Set �����ı� = �������_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          Elsif r_Report.�����ı� Like '%����%' Then
            Update ���Ӳ������� Set �����ı� = ���潨��_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          End If;
        End Loop;
        --���±���ʱ��
        Update ���Ӳ�����¼
        Set ���ʱ�� = Sysdate, ������ = ����ҽ��_In, ����ʱ�� = Sysdate
        Where ID = r_Advice.����id;
      Else
        --���жϵ������Ƿ��ж�Ӧ����ٺͱ��
        If Nvl(��������_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_����
          From �����ļ��ṹ A, �����ļ��ṹ B
          Where a.��id = b.Id And a.�������� = 3 And b.�������� = 1 And a.�����ı� Like '%����%' And a.�ļ�id = r_Advice.�ļ�id;
        
          If n_���� <= 0 Then
            v_Error := '�����Ƶ�����û���ҵ�����������Ӧ����ٻ�������ϵ����Ա���ã�';
            Raise Err_Custom;
          End If;
        End If;
        If Nvl(�������_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_����
          From �����ļ��ṹ A, �����ļ��ṹ B
          Where a.��id = b.Id And a.�������� = 3 And b.�������� = 1 And a.�����ı� Like '%���%' And a.�ļ�id = r_Advice.�ļ�id;
        
          If n_���� <= 0 Then
            v_Error := '�����Ƶ�����û���ҵ����������Ӧ����ٻ�������ϵ����Ա���ã�';
            Raise Err_Custom;
          End If;
        End If;
        If Nvl(���潨��_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_����
          From �����ļ��ṹ A, �����ļ��ṹ B
          Where a.��id = b.Id And a.�������� = 3 And b.�������� = 1 And a.�����ı� Like '%����%' And a.�ļ�id = r_Advice.�ļ�id;
        
          If n_���� <= 0 Then
            v_Error := '�����Ƶ�����û���ҵ������顿��Ӧ����ٻ�������ϵ����Ա���ã�';
            Raise Err_Custom;
          End If;
        End If;
      
        If r_Advice.������Դ = 1 Then
          --�����ȡ�Һŵ�ID
          Select Nvl(c.Id, 0)
          Into v_�Һŵ�id
          From ����ҽ����¼ B, ���˹Һż�¼ C
          Where b.�Һŵ� = c.No(+) And c.��¼״̬ In (1, 3) And b.Id = v_��ҽ��id;
        Else
          --����������޹Һŵ�ID��ֱ������Ϊ0
          v_�Һŵ�id := 0;
        End If;
      
        --�������Ӳ�����¼
        Select ���Ӳ�����¼_Id.Nextval Into v_����id From Dual;
        n_Rptcount := Getrptno(ҽ��id_In, r_Advice.��������);
        If n_Rptcount > 1 Then
          v_�������� := r_Advice.�������� || n_Rptcount;
        Else
          v_�������� := r_Advice.��������;
        End If;
        Insert Into ���Ӳ�����¼
          (ID, ������Դ, ����id, ��ҳid, Ӥ��, ����id, ��������, �ļ�id, ��������, ������, ����ʱ��, ���ʱ��, ������, ����ʱ��, ���汾, ǩ������)
        Values
          (v_����id, r_Advice.������Դ, r_Advice.����id, Decode(r_Advice.������Դ, 2, r_Advice.��ҳid, v_�Һŵ�id), r_Advice.Ӥ��,
           r_Advice.���˿���id, r_Advice.��������, r_Advice.�ļ�id, v_��������, ����ҽ��_In, Sysdate, Sysdate, ����ҽ��_In, Sysdate, 1, 2);
      
        --����ҽ�������¼
        Insert Into ����ҽ������ (ҽ��id, ����id, Risid) Values (v_��ҽ��id, v_����id, Risid_In);
      
        v_������� := 0;
      
        --�²�����������
        For r_File In c_File(r_Advice.�ļ�id) Loop
          Select ���Ӳ�������_Id.Nextval Into v_��������id From Dual;
          v_�����ı�   := r_File.�����ı�;
          v_�������id := 0;
        
          If Nvl(r_File.��������, 0) = 1 And Nvl(r_File.��id, 0) = 0 Then
            --���
            v_�������id := r_File.Id;
            v_��id       := v_��������id;
          End If;
        
          If Nvl(r_File.��������, 0) = 4 And r_File.Ҫ������ Is Not Null Then
            --Ԫ��
            v_�����ı� := Zl_Replace_Element_Value(r_File.Ҫ������, r_Advice.����id, r_Advice.��ҳid, r_Advice.������Դ, r_Advice.Id);
          End If;
        
          If Nvl(r_File.��id, 0) <> 0 Then
            v_�������id := 0;
          End If;
        
          v_������� := v_������� + 1;
        
          If Instr(v_���, '|' || r_File.��id || '|') > 0 Then
            Null;
          Else
            Insert Into ���Ӳ�������
              (ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id, �滻��,
               Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��, �������id)
            Values
              (v_��������id, v_����id, 1, 0, Decode(v_�������id, 0, v_��id, Null), v_�������, r_File.��������, r_File.������, r_File.��������,
               r_File.��������, Null, v_�����ı�, r_File.�Ƿ���, r_File.Ԥ�����id, r_File.�������, r_File.ʹ��ʱ��, r_File.����Ҫ��id,
               r_File.�滻��, r_File.Ҫ������, r_File.Ҫ������, r_File.Ҫ�س���, r_File.Ҫ��С��, r_File.Ҫ�ص�λ, r_File.Ҫ�ر�ʾ, r_File.������̬,
               r_File.Ҫ��ֵ��, Decode(v_�������id, 0, Null, v_�������id));
          End If;
        
          --Ϊ���ʱ�������ı�����
          If Nvl(r_File.��������, 0) = 3 And Nvl(r_File.��id, 0) <> 0 Then
            v_��� := v_��� || ',|' || r_File.Id || '|';
          
            If r_File.�����ı� Like '%����%' Then
              v_�����ı� := ��������_In || Chr(13) || Chr(13);
            Elsif r_File.�����ı� Like '%���%' Then
              v_�����ı� := �������_In || Chr(13) || Chr(13);
            Else
              v_�����ı� := ���潨��_In || Chr(13) || Chr(13);
            End If;
          
            For r_Con In c_Content(r_Advice.�ļ�id, r_File.Id) Loop
              Select ���Ӳ�������_Id.Nextval Into v_��������idnew From Dual;
              v_������� := v_������� + 1;
            
              Insert Into ���Ӳ�������
                (ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id,
                 �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��, �������id)
              Values
                (v_��������idnew, v_����id, 1, 0, v_��������id, v_�������, 2, r_Con.������, r_Con.��������, r_Con.��������, Null, v_�����ı�,
                 r_Con.�Ƿ���, r_Con.Ԥ�����id, r_Con.�������, r_Con.ʹ��ʱ��, r_Con.����Ҫ��id, r_Con.�滻��, r_Con.Ҫ������, r_Con.Ҫ������,
                 r_Con.Ҫ�س���, r_Con.Ҫ��С��, r_Con.Ҫ�ص�λ, r_Con.Ҫ�ر�ʾ, r_Con.������̬, r_Con.Ҫ��ֵ��,
                 Decode(v_�������id, 0, Null, v_�������id));
            End Loop;
          End If;
        End Loop;
      
        --����Ӳ�����ʽ�к����������ָ�ʽ�����ַ�������֮���������ֽ����ɼ�
        --Select ���� Into v_��ʽ���� From �����ļ���ʽ Where �ļ�ID=r_Advice.�ļ�ID;
        --Insert Into ���Ӳ�����ʽ (�ļ�ID,����) Values (v_����id,v_��ʽ����);
      
      End If;
    End If;
    Close c_Advice;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Receivereport;

  --5���޸����뵥��Ϣ
  Procedure Ӱ������Ϣ_�޸�
  (
    ҽ��id_In       ����ҽ����¼.Id%Type,
    ����_In         ������Ϣ.����%Type,
    �Ա�_In         ������Ϣ.�Ա�%Type,
    ����_In         ������Ϣ.����%Type,
    �ѱ�_In         ������Ϣ.�ѱ�%Type,
    ҽ�Ƹ��ʽ_In ������Ϣ.ҽ�Ƹ��ʽ%Type,
    ����_In         ������Ϣ.����%Type,
    ����_In         ������Ϣ.����״��%Type,
    ְҵ_In         ������Ϣ.ְҵ%Type,
    ���֤��_In     ������Ϣ.���֤��%Type,
    ��ͥ��ַ_In     ������Ϣ.��ͥ��ַ%Type,
    ��ͥ�绰_In     ������Ϣ.��ͥ�绰%Type,
    ��ͥ��ַ�ʱ�_In ������Ϣ.��ͥ��ַ�ʱ�%Type,
    ��������_In     ������Ϣ.��������%Type := Null
  ) As
  
    v_����     Varchar2(20);
    v_���䵥λ Varchar2(20);
    v_�������� Date;
    v_������Դ ����ҽ����¼.������Դ%Type;
    v_����id   ����ҽ����¼.����id%Type;
  Begin
    Begin
      Select ������Դ, ����id Into v_������Դ, v_����id From ����ҽ����¼ Where ID = ҽ��id_In;
    Exception
      When Others Then
        Return;
    End;
  
    If ��������_In Is Null And ����_In Is Not Null Then
      --�����������������
      v_���䵥λ := Substr(����_In, Length(����_In), 1);
      If Instr('��,��,��', v_���䵥λ) <= 0 Then
        v_���䵥λ := Null;
      Else
        v_���� := Replace(����_In, v_���䵥λ, '');
      End If;
      Begin
        v_���� := To_Number(v_����);
      Exception
        When Others Then
          v_���� := Null;
      End;
      If v_���� Is Not Null And v_���䵥λ Is Not Null Then
        Select Decode(v_���䵥λ, '��', Add_Months(Sysdate, -12 * v_����), '��', Add_Months(Sysdate, -1 * v_����), '��',
                       Sysdate - v_����)
        Into v_��������
        From Dual;
      End If;
    Else
      v_�������� := ��������_In;
    End If;
  
    If v_������Դ = 3 Then
      Update ������Ϣ
      Set ���� = ����_In, �Ա� = Nvl(�Ա�_In, �Ա�), ���� = ����_In, �������� = v_��������, �ѱ� = Nvl(�ѱ�_In, �ѱ�),
          ҽ�Ƹ��ʽ = Nvl(ҽ�Ƹ��ʽ_In, ҽ�Ƹ��ʽ), ���� = Nvl(����_In, ����), ����״�� = Nvl(����_In, ����״��), ְҵ = Nvl(ְҵ_In, ְҵ),
          ���֤�� = ���֤��_In, ��ͥ��ַ = ��ͥ��ַ_In, ��ͥ�绰 = ��ͥ�绰_In, ��ͥ��ַ�ʱ� = ��ͥ��ַ�ʱ�_In
      Where ����id = v_����id;
    
      --�޸Ķ�Ӧ��ҽ����¼
      Update ����ҽ����¼
      Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In
      Where ID = ҽ��id_In Or ���id = ҽ��id_In;
    Else
      Update ������Ϣ
      Set ���� = Nvl(����_In, ����), ����״�� = Nvl(����_In, ����״��), ְҵ = Nvl(ְҵ_In, ְҵ), ��ͥ��ַ = ��ͥ��ַ_In, ��ͥ�绰 = ��ͥ�绰_In,
          ��ͥ��ַ�ʱ� = ��ͥ��ַ�ʱ�_In
      Where ����id = v_����id;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ӱ������Ϣ_�޸�;

  --6��ȡ�����뵥��Ϣ
  Procedure ȡ��������뵥
  (
    ҽ��id_In     ����ҽ��ִ��.ҽ��id%Type,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := 0,
    �ܾ�ԭ��_In   ����ҽ������.ִ��˵��%Type := Null
  ) As
    --������ҽ��ID_IN=����ִ�е�ҽ��ID
  
    v_���ͺ� ����ҽ��ִ��.���ͺ�%Type;
  
  Begin
  
    Begin
      Select ���ͺ� Into v_���ͺ� From ����ҽ������ Where ҽ��id = ҽ��id_In;
    Exception
      When Others Then
        Return;
    End;
  
    Zl_����ҽ��ִ��_�ܾ�ִ��(ҽ��id_In, v_���ͺ�, ����Ա���_In, ����Ա����_In, ִ�в���id_In, �ܾ�ԭ��_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End ȡ��������뵥;

  --7������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_Insert
  (
    ������Դ_In   In Risҽ��ʧ�ܼ�¼.������Դ%Type,
    ����id_In     In Risҽ��ʧ�ܼ�¼.����id%Type,
    ��ҳid_In     In Risҽ��ʧ�ܼ�¼.��ҳid%Type,
    �Һŵ���_In   In Risҽ��ʧ�ܼ�¼.�Һŵ���%Type,
    ���ͺ�_In     In Risҽ��ʧ�ܼ�¼.���ͺ�%Type,
    �������id_In In Risҽ��ʧ�ܼ�¼.�������id%Type,
    ��챨����_In In Risҽ��ʧ�ܼ�¼.��챨����%Type,
    ��������_In   In Risҽ��ʧ�ܼ�¼.��������%Type
  ) Is
  Begin
    Insert Into Risҽ��ʧ�ܼ�¼
      (ID, ������Դ, ����id, ��ҳid, �Һŵ���, ���ͺ�, �������id, ��챨����, ��������, ����ʱ��, �ط�����)
    Values
      (Risҽ��ʧ�ܼ�¼_Id.Nextval, ������Դ_In, ����id_In, ��ҳid_In, �Һŵ���_In, ���ͺ�_In, �������id_In, ��챨����_In, ��������_In, Sysdate, 0);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Risҽ��ʧ�ܼ�¼_Insert;

  --8������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_�ط�
  (
    Id_In       In Risҽ��ʧ�ܼ�¼.Id%Type,
    ��������_In In Number
  ) Is
    v_�ط����� Risҽ��ʧ�ܼ�¼.�ط�����%Type;
  Begin
    --��������_In -- 1 �ط��ɹ���ɾ����¼��2--�ط�ʧ��
  
    If ��������_In = 1 Then
      Delete From Risҽ��ʧ�ܼ�¼ Where ID = Id_In;
    Else
      Select �ط����� Into v_�ط����� From Risҽ��ʧ�ܼ�¼ Where ID = Id_In;
      If v_�ط����� >= 99 Then
        v_�ط����� := 99;
      Else
        v_�ط����� := v_�ط����� + 1;
      End If;
      Update Risҽ��ʧ�ܼ�¼ Set ����ʱ�� = Sysdate, �ط����� = v_�ط����� Where ID = Id_In;
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Risҽ��ʧ�ܼ�¼_�ط�;

  --9�����˺��½�סԺ���˵���
  Procedure ����ҽ��_�ؽ�����
  (
    ҽ��id_In In ����ҽ������.ҽ��id%Type,
    No_In     In ����ҽ������.No%Type,
    Action_In In Number
  ) Is
    -- Action_In: 1 �ؽ����ݣ�2 ȡ���ؽ�����
    v_No ����ҽ������.No%Type;
  Begin
    If Action_In = 1 Then
      Select Nextno(14) Into v_No From Dual;
    
      Update ����ҽ������
      Set NO = v_No, �Ʒ�״̬ = 0
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In);
      Update סԺ���ü�¼ Set ҽ����� = Null Where NO = No_In;
    Elsif Action_In = 2 Then
      Update סԺ���ü�¼ Set ҽ����� = ҽ��id_In Where NO = No_In;
      Update ����ҽ������
      Set NO = No_In, �Ʒ�״̬ = 4
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In);
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End ����ҽ��_�ؽ�����;

  --10����ӡRIS���ԤԼ֪ͨ��
  Procedure Ris���ԤԼ_��ӡ(ҽ��id_In In Ris���ԤԼ.ҽ��id%Type) Is
    v_Temp     Varchar2(255);
    v_��Ա���� ��Ա��.����%Type;
  Begin
    --ȡ��ǰ������Ա
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  
    Update Ris���ԤԼ Set �Ƿ��ӡ = 1, ��ӡ�� = v_��Ա����, ��ӡʱ�� = Sysdate Where ҽ��id = ҽ��id_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris���ԤԼ_��ӡ;

  --11������RIS�ֿ������ò���
  Procedure Ris���ÿ���_Update
  (
    �������_In Ris���ÿ���.�������%Type,
    ����_In     Ris���ÿ���.����%Type,
    ����ids_In  Varchar2,
    ��������_In Number
  ) Is
  
    l_����id   t_Numlist := t_Numlist();
    v_����ris  Ris���ÿ���.�Ƿ�����ris%Type;
    v_����ԤԼ Ris���ÿ���.�Ƿ�����ԤԼ%Type;
  
    Cursor c_Dept(Dept_In Varchar2) Is
      Select Column_Value From Table(f_Num2list(Dept_In));
  Begin
  
    If ��������_In = 1 Then
      v_����ris  := 1;
      v_����ԤԼ := Null;
      Delete From Ris���ÿ��� Where ������� = �������_In And ���� = ����_In And �Ƿ�����ris = 1;
    Else
      v_����ris  := Null;
      v_����ԤԼ := 1;
      Delete From Ris���ÿ��� Where ������� = �������_In And ���� = ����_In And �Ƿ�����ԤԼ = 1;
    End If;
  
    If ����ids_In Is Null Then
      Insert Into Ris���ÿ���
        (ID, �������, ����, ����id, �Ƿ�����ris, �Ƿ�����ԤԼ)
      Values
        (Ris���ÿ���_Id.Nextval, �������_In, ����_In, Null, v_����ris, v_����ԤԼ);
    Else
      Open c_Dept(����ids_In);
      Fetch c_Dept Bulk Collect
        Into l_����id;
      Close c_Dept;
    
      Forall I In 1 .. l_����id.Count
        Insert Into Ris���ÿ���
          (ID, �������, ����, ����id, �Ƿ�����ris, �Ƿ�����ԤԼ)
        Values
          (Ris���ÿ���_Id.Nextval, �������_In, ����_In, l_����id(I), v_����ris, v_����ԤԼ);
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris���ÿ���_Update;

  --12��ɾ��RIS�ֿ������ò���
  Procedure Ris���ÿ���_Delete Is
  
  Begin
    Delete From Ris���ÿ���;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris���ÿ���_Delete;

  --13������Ԫ������ȡ��Ϣ
  Function Ris_Replace_Element_Value
  (
    Ԫ����_In   In ����������Ŀ.������%Type,
    ����id_In   In ���Ӳ�����¼.����id%Type,
    ����id_In   In ���Ӳ�����¼.��ҳid%Type,
    ������Դ_In In ���Ӳ�����¼.������Դ%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type
  ) Return Varchar2 Is
    v_Return Varchar2(4000) := Null;
    Cursor c_Patient Is
      Select ����, �Ա�, Decode(�Ա�, '��', 'M', 'Ů', 'F', 'O') As �Ա����, ��������, ����id, ��ϵ�˵�ַ, ��ͥ�绰, ��ϵ�˵绰, ����״��, ���֤��, ��ǰ����id,
             ��ǰ����id, ��ǰ���� As ����, ���￨��, ��Ժʱ��, ��Ժʱ��
      From ������Ϣ
      Where ����id = ����id_In;
    r_Patient c_Patient%RowType;
  
    Cursor c_Order Is
      Select ��ҳid, Ӥ��, Decode(������Դ, 1, 'OUTPAT', 2, 'INPAT', 'UNK') As ������Դ, ����ҽ��, ����ʱ��, У�Ի�ʿ, ҽ������, ������־, ִ�п���id
      From ����ҽ����¼
      Where ID = ҽ��id_In;
    r_Order c_Order%RowType;
  
    Cursor c_Diagnose Is
      Select ������� || Decode(Nvl(�Ƿ�����, 0), 0, '', ' (��)') As �ٴ����
      From �������ҽ�� A, ������ϼ�¼ B
      Where a.ҽ��id = ҽ��id_In And a.���id = b.Id;
    r_Diagnose c_Diagnose%RowType;
  
    --��ȡָ�����������
    Procedure p_Get_Rowtype(Table_In In Varchar2) Is
    Begin
      If Table_In = '������Ϣ' Then
        Open c_Patient;
        Fetch c_Patient
          Into r_Patient;
      Elsif Table_In = '����ҽ����¼' Then
        Open c_Order;
        Fetch c_Order
          Into r_Order;
      Elsif Table_In = '������ϼ�¼' Then
        Open c_Diagnose;
        Fetch c_Diagnose
          Into r_Diagnose;
      End If;
    Exception
      When Others Then
        Null;
    End p_Get_Rowtype;
  
  Begin
    Case
    --ֱ�ӷ��ص�����Ԫ��
      When Ԫ����_In = 'ҽ��ID' Then
        v_Return := ҽ��id_In;
      When Ԫ����_In = '����ID' Then
        v_Return := ����id_In;
      
    --�������Ա𵥶�����������Ӥ��
      When Instr(',����,�Ա�,�Ա����,��������,', ',' || Ԫ����_In || ',') > 0 Then
        p_Get_Rowtype('����ҽ����¼');
        p_Get_Rowtype('������Ϣ');
        If Nvl(r_Order.Ӥ��, 0) = 0 Then
          If Ԫ����_In = '����' Then
            v_Return := r_Patient.����;
          Elsif Ԫ����_In = '�Ա�' Then
            v_Return := r_Patient.�Ա�;
          Elsif Ԫ����_In = '�Ա����' Then
            v_Return := r_Patient.�Ա����;
          Elsif Ԫ����_In = '��������' Then
            v_Return := To_Char(r_Patient.��������, 'YYYYMMDDMISS');
          End If;
        Else
          If Ԫ����_In = '����' Then
            Select Decode(Ӥ������, Null, r_Patient.���� || '֮Ӥ' || Trim(To_Char(���, '9')), Ӥ������) As Ӥ������
            Into v_Return
            From ������������¼
            Where ����id = ����id_In And ��ҳid = r_Order.��ҳid And ��� = Nvl(r_Order.Ӥ��, 0);
          Elsif Instr('�Ա�', Ԫ����_In) > 0 Then
            Select Ӥ���Ա�
            Into v_Return
            From ������������¼
            Where ����id = ����id_In And ��ҳid = r_Order.��ҳid And ��� = Nvl(r_Order.Ӥ��, 0);
            If Ԫ����_In = '�Ա����' Then
              Select Decode(v_Return, '��', 'M', 'Ů', 'F', 'O') Into v_Return From Dual;
            End If;
          Elsif Ԫ����_In = '��������' Then
            Select ����ʱ��
            Into v_Return
            From ������������¼
            Where ����id = ����id_In And ��ҳid = r_Order.��ҳid And ��� = Nvl(r_Order.Ӥ��, 0);
            v_Return := To_Char(v_Return, 'YYYYMMDDMISS');
          End If;
        End If;
      
    --��ѯ������Ϣ���ص�Ԫ��
      When Instr(',��ϵ�˵�ַ,��ͥ�绰,��ϵ�˵绰,����״��,���֤��,����,���￨��,��Ժʱ��,��Ժʱ��,', ',' || Ԫ����_In || ',') > 0 Then
        p_Get_Rowtype('������Ϣ');
        Case Ԫ����_In
          When '��ϵ�˵�ַ' Then
            v_Return := r_Patient.��ϵ�˵�ַ;
          When '��ͥ�绰' Then
            v_Return := r_Patient.��ͥ�绰;
          When '��ϵ�˵绰' Then
            v_Return := r_Patient.��ϵ�˵绰;
          When '����״��' Then
            v_Return := r_Patient.����״��;
          When '���֤��' Then
            v_Return := r_Patient.���֤��;
          When '����' Then
            v_Return := r_Patient.����;
          When '���￨��' Then
            v_Return := r_Patient.���￨��;
          When '��Ժʱ��' Then
            v_Return := To_Char(r_Patient.��Ժʱ��, 'YYYYMMDDMISS');
          When '��Ժʱ��' Then
            v_Return := To_Char(r_Patient.��Ժʱ��, 'YYYYMMDDMISS');
          Else
            v_Return := '';
        End Case;
        --��ѯҽ�����ص�Ԫ��
      When Instr(',������Դ,����ҽ��,����ʱ��,У�Ի�ʿ,ҽ������,������־,������־����,', ',' || Ԫ����_In || ',') > 0 Then
        p_Get_Rowtype('����ҽ����¼');
        Case Ԫ����_In
          When '������Դ' Then
            v_Return := r_Order.������Դ;
          When '����ҽ��' Then
            v_Return := r_Order.����ҽ��;
          When '����ʱ��' Then
            v_Return := To_Char(r_Order.����ʱ��, 'YYYYMMDDMISS');
          When 'У�Ի�ʿ' Then
            v_Return := r_Order.У�Ի�ʿ;
          When 'ҽ������' Then
            v_Return := r_Order.ҽ������;
          When '������־' Then
            v_Return := r_Order.������־;
        End Case;
        --��ѯ��ϼ�¼���ص�Ԫ��
      When Ԫ����_In = '�ٴ����' Then
        p_Get_Rowtype('������ϼ�¼');
        v_Return := r_Diagnose.�ٴ����;
      
      Else
        --���в�ѯSQL����ֵ��Ԫ��
        If Ԫ����_In = 'ִ��վ��' Then
          p_Get_Rowtype('����ҽ����¼');
          Select Decode(վ��, 1, 'SITE0002', 2, 'SITE0001', 3, 'SITE0003', 'SITE0001')
          Into v_Return
          From ���ű�
          Where ID = r_Order.ִ�п���id;
        End If;
        If Ԫ����_In = '��ǰ��������' Then
          p_Get_Rowtype('������Ϣ');
          Select ���� Into v_Return From ���ű� Where ID = r_Patient.��ǰ����id;
        End If;
        If Ԫ����_In = '��������' Then
          p_Get_Rowtype('������Ϣ');
          Select ���� Into v_Return From ���ű� Where ID = r_Patient.��ǰ����id;
        End If;
        If Ԫ����_In = '��ʶ��' Then
          Select Decode(a.������Դ, 1, c.�����, 2, Decode(c.סԺ��, Null, c.�����, c.סԺ��), 4, c.������, c.�����)
          Into v_Return
          From ����ҽ����¼ A, ������Ϣ C
          Where a.����id = c.����id And a.Id = ҽ��id_In;
        End If;
    End Case;
  
    Return Trim(v_Return);
  Exception
    When Others Then
      Return Null;
  End Ris_Replace_Element_Value;

  --14��ɾ��RIS��Ժ���ò���
  Procedure Ris��Ժ����_Delete Is
  Begin
    Delete From Ris��Ժ����;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris��Ժ����_Delete;

  --15������RISRis��Ժ���ò���
  Procedure Ris��Ժ����_Update
  (
    Id_In           Ris��Ժ����.Id%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    �û���_In       Ris��Ժ����.�û���%Type,
    ����_In         Ris��Ժ����.����%Type,
    ���ݿ������_In Ris��Ժ����.���ݿ������%Type
  ) Is
  
  Begin
  
    Insert Into Ris��Ժ����
      (ID, ҽԺ����, ҽԺ����, �û���, ����, ���ݿ������)
    Values
      (Id_In, ҽԺ����_In, ҽԺ����_In, �û���_In, ����_In, ���ݿ������_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris��Ժ����_Update;

End b_Zlxwinterface;
/





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0013' Where ���=&n_System;
Commit;
