CREATE OR REPLACE Procedure Zl_��ʷ����_����
(
  strDecode_IN  varchar2
)Is
  v_TempDecode varchar2(1000);
Begin

insert into ����걾��Ϣ(�걾ID, ҽ��ID, �걾����,�걾����,����,��������)
select ����걾��Ϣ_�걾ID.Nextval,a.ҽ��ID,a.�걾��λ,0,a.����,b.����ʱ��
from Ӱ����걾 a, Ӱ��걾����ȡ�� b where a.ҽ��id=b.ҽ��id 
and not exists(Select 1 From ����걾��Ϣ where ҽ��ID=a.ҽ��ID and �걾����=a.�걾��λ and ����=a.���� and ��������=b.����ʱ��);

--���Ӵ���Decode����
v_TempDecode:= 'insert into ��������Ϣ(����ҽ��ID,�����,ҽ��ID,�������,�޼�����,ʣ��λ��)
               select ��������Ϣ_����ҽ��ID.Nextval,�����,ҽ��ID,' || strDecode_IN || ',�޼�����,ʣ��걾λ��
               from Ӱ��걾����ȡ�� where ҽ��ID not in(select ҽ��ID from ��������Ϣ)';

Execute Immediate v_TempDecode;

insert into �����ͼ���Ϣ(ID,ҽ��ID,�ͼ쵥λ,�ͼ����,�ͼ���,�ͼ�����,�Ǽ���,����״̬,����ԭ��,��ע)
select �����ͼ���Ϣ_id.nextval,ҽ��ID,'��Ժ','', 'δ¼��',����ʱ��,decode(���ռ�ʦ,null,'δ¼��',���ռ�ʦ),decode(�������,'1','1','0'),����ԭ��,��ע
from Ӱ��걾����ȡ�� a where not exists(Select 1 From �����ͼ���Ϣ where ҽ��ID=a.ҽ��id and �ͼ�����=a.����ʱ�� and ����ԭ��=a.����ԭ��);

update ����걾��Ϣ a set �ͼ�ID=(select  ID from �����ͼ���Ϣ where ҽ��ID=a.ҽ��ID and rownum=1);


Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_��ʷ����_����;
