Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1288, 0, 0, 0, 0, 19,'XW�ؼ�ͼ���ַ', 0, 'http://127.0.0.1:8080/KeyImage.aspx?colid0=22&'||'colvalue0=[@STU_NO]', 'XW PACS WEB�������ĵ�ַ��'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1288 And ������ = 'XW�ؼ�ͼ���ַ');
  
--85463:����,2015-07-01,����б�Ӱ�����ͼ�鲿λ����
Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1290, 1, 1, 0, 1, 49,'Ӱ��������', 0, 0, '����б����ݰ�Ӱ��������'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1290 And ������ = 'Ӱ��������');
  
Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1291, 1, 1, 0, 1, 53,'Ӱ��������', 0, 0, '����б����ݰ�Ӱ��������'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1291 And ������ = 'Ӱ��������');
  
Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1294, 1, 1, 0, 1, 109,'Ӱ��������', 0, 0, '����б����ݰ�Ӱ��������'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1294 And ������ = 'Ӱ��������');
  
Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1290, 1, 1, 0, 1, 50,'��鲿λ����', 0, 0, '����б����ݰ���鲿λ����'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1290 And ������ = '��鲿λ����');
  
Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1291, 1, 1, 0, 1, 54,'��鲿λ����', 0, 0, '����б����ݰ���鲿λ����'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1291 And ������ = '��鲿λ����');
  
Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1294, 1, 1, 0, 1, 110,'��鲿λ����', 0, 0, '����б����ݰ���鲿λ����'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1294 And ������ = '��鲿λ����'); 
  
--84345:����,2015-06-30,�Ƿ���WorkList
Alter Table Ӱ�����¼ Add �Ƿ��� Number(1);

--84345:����,2015-06-30,�Ƿ���WorkList
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1290, '����', User, 'ZL_Ӱ�����¼_���Ͱ���', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1290 And ���� = '����' And Upper(����) = Upper('ZL_Ӱ�����¼_���Ͱ���'));
		 
--84345:����,2015-06-30,�Ƿ���WorkList
CREATE OR REPLACE Procedure ZL_Ӱ�����¼_���Ͱ���
( 
  ҽ��ID_In       Ӱ�����¼.ҽ��ID%Type,
  ���ͺ�_In       Ӱ�����¼.���ͺ�%Type,
  �Ƿ���_In     Ӱ�����¼.�Ƿ���%Type,
  ��鼼ʦ_In     Ӱ�����¼.��鼼ʦ%Type, 
  ��鼼ʦ��_In   Ӱ�����¼.��鼼ʦ��%Type, 
  ִ�м�_In       ����ҽ������.ִ�м�%Type
) As 
Begin 
 
  Update Ӱ�����¼ 
  Set    ��鼼ʦ = ��鼼ʦ_In, ��鼼ʦ�� = ��鼼ʦ��_In, �Ƿ��� = �Ƿ���_In
  Where  ҽ��ID = ҽ��ID_In and ���ͺ� =���ͺ�_In; 
 
  Update ����ҽ������ 
  Set ִ�м� = ִ�м�_In
  Where ҽ��ID=ҽ��ID_In and ���ͺ�=���ͺ�_In; 
Exception 
  When Others Then 
    Zl_Errorcenter(Sqlcode, Sqlerrm); 
End ZL_Ӱ�����¼_���Ͱ���;
/

--85773:����,2015-06-29,XW3D��Ƭ
Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1288, 0, 0, 0, 0, 17,'XWWEB��Ƭ��ַ', 0, 'http://127.0.0.1:8080/imageweb/imageAction.action?ColID0=22&ColValue0=[@STU_NO]', 'XW PACS WEB�������ĵ�ַ��'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1288 And ������ = 'XWWEB��Ƭ��ַ');
  
Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1288, 0, 0, 0, 0, 18,'XW3D��Ƭ����', 0, 'Study3D', 'XW PACS 3D��Ƭʱ�Ĺ�Ƭ���ͣ���Study3D��Ϊֱ�Ӵ򿪸ü���ͼ�񣬵�һ�����м���3D����SeriesList3D�� Ϊ�򿪸ü������У����û�ѡ�����ĳ�����е�3D���'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1288 And ������ = 'XW3D��Ƭ����');