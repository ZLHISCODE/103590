--�걾����====================================================================================================


--��ӱ걾��Ϣ
CREATE OR REPLACE function Zl_����걾_����
(
  ҽ��ID_IN   ����걾��Ϣ.ҽ��ID%Type,     
  �걾����_IN ����걾��Ϣ.�걾����%Type,
  �걾����_IN ����걾��Ϣ.�걾����%Type,
  �ɼ���λ_IN ����걾��Ϣ.�ɼ���λ%Type,
  �걾����_IN ����걾��Ϣ.����%Type,
  �������_IN ����걾��Ϣ.�������%Type,
  ԭ�б��_IN ����걾��Ϣ.ԭ�б��%Type,
  ���λ��_IN ����걾��Ϣ.���λ��%Type,
  ��������_IN ����걾��Ϣ.��������%Type,
  ��ע��Ϣ_IN ����걾��Ϣ.��ע%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;

v_�걾ID ����걾��Ϣ.�걾ID%Type;

Begin
  select ����걾��Ϣ_�걾ID.NEXTVAL into  v_�걾ID  from dual;
     
  insert into ����걾��Ϣ(�걾ID,ҽ��ID,�걾����,�걾����,�ɼ���λ,����,�������,ԭ�б��,���λ��,��������,��ע)
  values(v_�걾ID, ҽ��ID_IN, �걾����_IN, �걾����_IN, �ɼ���λ_IN, �걾����_IN, �������_IN, ԭ�б��_IN, ���λ��_IN, ��������_IN, ��ע��Ϣ_IN);
  
  commit;
  
  return  v_�걾ID;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����걾_����;
/


--���±걾��Ϣ
CREATE OR REPLACE Procedure Zl_����걾_����
(
  �걾ID_IN   ����걾��Ϣ.�걾ID%Type,     
  �걾����_IN ����걾��Ϣ.�걾����%Type,
  �걾����_IN ����걾��Ϣ.�걾����%Type,
  �ɼ���λ_IN ����걾��Ϣ.�ɼ���λ%Type,
  �걾����_IN ����걾��Ϣ.���λ��%Type,
  �������_IN ����걾��Ϣ.�������%Type,
  ԭ�б��_IN ����걾��Ϣ.ԭ�б��%Type,
  ���λ��_IN ����걾��Ϣ.���λ��%Type,
  ��ע��Ϣ_IN ����걾��Ϣ.��ע%Type
) Is
Begin
  update ����걾��Ϣ 
  set �걾����=�걾����_IN,�걾����=�걾����_IN,�ɼ���λ=�ɼ���λ_IN,
      ����=�걾����_IN,�������=�������_IN,ԭ�б��=ԭ�б��_IN,
      ���λ��=���λ��_IN,��ע=��ע��Ϣ_IN
  where �걾ID=�걾ID_IN;   
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����걾_����;
/

--�����ͼ�걾
CREATE OR REPLACE Procedure Zl_����걾_����
(   
  ҽ��ID_IN   ��������Ϣ.ҽ��ID%Type,  
  �������_IN ��������Ϣ.�������%Type,   
  �ͼ쵥λ_IN �����ͼ���Ϣ.�ͼ쵥λ%Type,
  �ͼ����_IN �����ͼ���Ϣ.�ͼ����%Type,
  �ͼ���_IN   �����ͼ���Ϣ.�ͼ���%Type,
  �ͼ�����_IN �����ͼ���Ϣ.�ͼ�����%Type,
  ��ϵ��ʽ_IN �����ͼ���Ϣ.��ϵ��ʽ%Type,
  �Ǽ���_IN   �����ͼ���Ϣ.�Ǽ���%Type
) Is
  v_����� ��������Ϣ.�����%Type := null;
  v_������� ��������Ϣ.�������%Type;
Begin

  begin
    select ����� into v_����� from ��������Ϣ where ҽ��ID=ҽ��ID_IN;      
  exception
    When Others Then v_����� := null;	
  end;              
     
  if v_����� is null then    
     --û���ҵ���ҽ����Ӧ�Ĳ�����
     
     --���ɲ����
     Select Lpad(��������Ϣ_�����.NEXTVAL, 8, 0) into v_����� from dual;  
      
     --ȡ�õ�ǰ����������
     v_������� := �������_IN;
  
     --��Ӳ����ͼ���Ϣ
     insert into �����ͼ���Ϣ(ID, ҽ��ID,�ͼ쵥λ,�ͼ����,�ͼ���,�ͼ�����,��ϵ��ʽ,�Ǽ���,����״̬)
     values(�����ͼ���Ϣ_ID.NEXTVAL, ҽ��ID_IN, �ͼ쵥λ_IN, �ͼ����_IN, �ͼ���_IN, �ͼ�����_IN, ��ϵ��ʽ_IN, �Ǽ���_IN, 1);
     
     --��Ӳ�������Ϣ,���պ󣬼�����ȡ������
     insert into ��������Ϣ(�����, ҽ��ID, �������, ��ǰ����)
     values(v_�����, ҽ��ID_IN, v_�������, decode(v_�������, 3, 3, 1));
  else
    --���ü���ѱ����չ�ʱ����ֻ����ͼ���Ϣ   
     insert into �����ͼ���Ϣ(ID, ҽ��ID,�ͼ쵥λ,�ͼ����,�ͼ���,�ͼ�����,��ϵ��ʽ,�Ǽ���,����״̬)
     values(�����ͼ���Ϣ_ID.NEXTVAL, ҽ��ID_IN, �ͼ쵥λ_IN, �ͼ����_IN, �ͼ���_IN, �ͼ�����_IN, ��ϵ��ʽ_IN, �Ǽ���_IN, 1);    
  end if;  
  
  --���¶�Ӧҽ����ִ��˵��...             
  update ����ҽ������ 
  set ִ��˵��=ִ��˵�� || chr(13) || '�걾�ѱ����� [ ʱ��:'|| �ͼ�����_IN  || '    �Ǽ���:' || �Ǽ���_IN || '] '
  where ҽ��ID=ҽ��ID_IN;  
  
  --����ִ�й���
  update ����ҽ������ set ִ�й���=2, ����ʱ��=�ͼ�����_IN where ҽ��ID=ҽ��ID_IN and nvl(ִ�й���,0) < 2;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����걾_����;
/

--�����ͼ�걾 
CREATE OR REPLACE Procedure Zl_����걾_����
(
  ҽ��ID_IN   ��������Ϣ.ҽ��ID%Type,     
  �ͼ쵥λ_IN �����ͼ���Ϣ.�ͼ쵥λ%Type,
  �ͼ����_IN �����ͼ���Ϣ.�ͼ����%Type,
  �ͼ���_IN   �����ͼ���Ϣ.�ͼ���%Type,
  �ͼ�����_IN �����ͼ���Ϣ.�ͼ�����%Type,
  ��ϵ��ʽ_IN �����ͼ���Ϣ.��ϵ��ʽ%Type,
  �Ǽ���_IN   �����ͼ���Ϣ.�Ǽ���%Type,
  ����ԭ��_IN �����ͼ���Ϣ.����ԭ��%Type,
  ֪ͨ��_IN   �����ͼ���Ϣ.֪ͨ��%Type
) Is
Begin
     insert into �����ͼ���Ϣ(ID, ҽ��ID,�ͼ쵥λ,�ͼ����,�ͼ���,�ͼ�����,��ϵ��ʽ,�Ǽ���,����״̬, ����ԭ��, ֪ͨ��)
     values(�����ͼ���Ϣ_ID.NEXTVAL, ҽ��ID_IN, �ͼ쵥λ_IN, �ͼ����_IN, �ͼ���_IN, �ͼ�����_IN, ��ϵ��ʽ_IN, �Ǽ���_IN, 2, ����ԭ��_IN, ֪ͨ��_IN); 
     
     --���¶�Ӧҽ����ִ��˵��...   
     update ����ҽ������ 
     set ִ��˵��=ִ��˵�� || chr(13) || '�걾�ѱ����� [ ʱ��:'|| �ͼ�����_IN || '   ����ԭ��:' || ����ԭ��_IN || '    �Ǽ���:' || �Ǽ���_IN || '] '
     where ҽ��ID=ҽ��ID_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����걾_����;
/






--�������====================================================================================================





--��������
CREATE OR REPLACE function Zl_������_����
(
  ��������_IN ��������Ϣ.��������%Type,     
  ʹ���˷�_IN ��������Ϣ.ʹ���˷�%Type,
  �����˷�_IN ��������Ϣ.�����˷�%Type,
  ��������_IN ��������Ϣ.��������%Type,
  ��Ч��_IN   ��������Ϣ.��Ч��%Type,
  ��������_IN ��������Ϣ.��������%Type,
  ��¡��_IN   ��������Ϣ.��¡��%Type,
  ���ö���_IN ��������Ϣ.���ö���%Type,
  ������_IN ��������Ϣ.������%Type,
  Ӧ�����_IN ��������Ϣ.Ӧ�����%Type,
  �Ǽ���_IN   ��������Ϣ.�Ǽ���%Type,
  �Ǽ�ʱ��_IN ��������Ϣ.�Ǽ�ʱ��%Type,
  ��ע_IN     ��������Ϣ.��ע%Type
) return Number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_����ID ��������Ϣ.����ID%Type;
Begin
  select ��������Ϣ_����ID.NEXTVAL into v_����ID from dual;
       
  insert into ��������Ϣ(����ID,��������,ʹ���˷�,�����˷�,��������,��Ч��,��������,��¡��,���ö���,������,Ӧ�����,�Ǽ���,�Ǽ�ʱ��,ʹ��״̬,��ע)
  values(v_����ID, ��������_IN, ʹ���˷�_IN, �����˷�_IN, ��������_IN, 
         ��Ч��_IN, ��������_IN, ��¡��_IN, ���ö���_IN, ������_IN, Ӧ�����_IN,�Ǽ���_IN,�Ǽ�ʱ��_IN,1,��ע_IN);
         
  commit;                  
         
  return v_����ID;       
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_������_����;
/

--��ʹ�ù��Ŀ��壬���ܸ��¿������Ƶ�����
CREATE OR REPLACE Procedure Zl_������_����
(
  ����ID_IN   ��������Ϣ.����ID%Type,           
  ��������_IN ��������Ϣ.��������%Type,     
  ʹ���˷�_IN ��������Ϣ.ʹ���˷�%Type,
  �����˷�_IN ��������Ϣ.�����˷�%Type,
  ��������_IN ��������Ϣ.��������%Type,
  ��Ч��_IN   ��������Ϣ.��Ч��%Type,
  ��������_IN ��������Ϣ.��������%Type,
  ��¡��_IN   ��������Ϣ.��¡��%Type,
  ���ö���_IN ��������Ϣ.���ö���%Type,
  ������_IN ��������Ϣ.������%Type,
  Ӧ�����_IN ��������Ϣ.Ӧ�����%Type,
  �Ǽ���_IN   ��������Ϣ.�Ǽ���%Type,
  �Ǽ�ʱ��_IN ��������Ϣ.�Ǽ�ʱ��%Type,
  ��ע_IN     ��������Ϣ.��ע%Type
) Is
Begin
  update ��������Ϣ
  set ��������=��������_IN, ʹ���˷�=ʹ���˷�_IN,�����˷�=�����˷�_IN,��������=��������_IN,
      ��Ч��=��Ч��_IN,��������=��������_IN,��¡��=��¡��_IN,���ö���=���ö���_IN,
      ������=������_IN,Ӧ�����=Ӧ�����_IN,�Ǽ���=�Ǽ���_IN,�Ǽ�ʱ��=�Ǽ�ʱ��_IN,��ע=��ע_IN
  where ����ID=����ID_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_������_����;
/

--���¿���ʹ��״̬
CREATE OR REPLACE Procedure Zl_������_ʹ��״̬
(
  ����ID_IN ��������Ϣ.����ID%Type,           
  ʹ��״̬_IN ��������Ϣ.ʹ��״̬%Type
) Is
Begin
  update ��������Ϣ set ʹ��״̬=ʹ��״̬_IN where ����ID=����ID_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_������_ʹ��״̬;
/

--ɾ��������Ϣ
CREATE OR REPLACE Procedure Zl_������_ɾ��
(
  ����ID_IN ��������Ϣ.����ID%Type
) Is
Begin
  delete ��������Ϣ where ����ID=����ID_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_������_ɾ��; 
/


--�������巴��
CREATE OR REPLACE function Zl_�����巴��_����
(
  ����ID_IN      �����巴��.����ID%Type,   
  �ο������_IN  �����巴��.�ο������%Type,
  ʵ������_IN    �����巴��.ʵ������%Type,
  ��������_IN    �����巴��.��������%Type,
  ����ʱ��_IN    �����巴��.����ʱ��%Type,
  ����ҽ��_IN    �����巴��.����ҽ��%Type,
  �������_IN    �����巴��.�������%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id �����巴��.ID%Type;
Begin
  select �����巴��_Id.Nextval into v_id from dual;
  
  insert into �����巴��(ID, ����ID,�ο������,ʵ������,��������,����ʱ��,����ҽ��,�������)
  values(v_id, ����ID_IN, �ο������_IN, ʵ������_IN, ��������_IN, ����ʱ��_IN, ����ҽ��_IN, �������_IN);
  
  commit;
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����巴��_����;
/

--���¿��巴����Ϣ
CREATE OR REPLACE Procedure Zl_�����巴��_����
(
  ID_IN          �����巴��.ID%Type,   
  �ο������_IN  �����巴��.�ο������%Type,
  ʵ������_IN    �����巴��.ʵ������%Type,
  ��������_IN    �����巴��.��������%Type,
  ����ʱ��_IN    �����巴��.����ʱ��%Type,
  ����ҽ��_IN    �����巴��.����ҽ��%Type,
  �������_IN    �����巴��.�������%Type
) Is
Begin
  Update �����巴��
  set �ο������=�ο������_IN,ʵ������=ʵ������_IN,��������=��������_IN,
       ����ʱ��=����ʱ��_IN,����ҽ��=����ҽ��_IN,�������=�������_IN
  where ID=ID_IN;     
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����巴��_����;
/


--ɾ��������¼
CREATE OR REPLACE Procedure Zl_�����巴��_ɾ��
(
  ID_IN �����巴��.ID%Type
) Is
Begin
  Delete �����巴�� where ID=ID_IN;   
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����巴��_ɾ��;
/


--���������ײ�
CREATE OR REPLACE function Zl_�����ײ�_����
(
  �ײ�����_IN   �����ײ���Ϣ.�ײ�����%Type,   
  �ײ�˵��_IN   �����ײ���Ϣ.�ײ�˵��%Type,
  ����ʱ��_IN   �����ײ���Ϣ.����ʱ��%Type,
  ������_IN     �����ײ���Ϣ.������%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id �����ײ���Ϣ.�ײ�ID%Type;
Begin
  select �����ײ���Ϣ_�ײ�ID.Nextval into v_id from dual;
  
  insert into �����ײ���Ϣ(�ײ�ID,�ײ�����,�ײ�˵��,������,����ʱ��)
  values(v_id, �ײ�����_IN, �ײ�˵��_IN, ������_IN, ����ʱ��_IN);
  
  commit;
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����ײ�_����;
/


--���²����ײ�
CREATE OR REPLACE procedure Zl_�����ײ�_����
(
  �ײ�ID_IN     �����ײ���Ϣ.�ײ�ID%Type,     
  �ײ�����_IN   �����ײ���Ϣ.�ײ�����%Type,   
  �ײ�˵��_IN   �����ײ���Ϣ.�ײ�˵��%Type
)Is
Begin
  --�����ײ���Ϣ
  update  �����ײ���Ϣ set �ײ�����=�ײ�����_IN, �ײ�˵��=�ײ�˵��_IN where �ײ�ID=�ײ�ID_IN;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����ײ�_����;
/



--ɾ�������ײ�
CREATE OR REPLACE procedure Zl_�����ײ�_ɾ��
(
  �ײ�ID_IN     �����ײ���Ϣ.�ײ�ID%Type
)Is
Begin
  --ɾ���ײ���Ϣ
  delete �����ײ���Ϣ where �ײ�ID=�ײ�ID_IN;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����ײ�_ɾ��;
/
 

--�����ײͿ������
CREATE OR REPLACE function Zl_�����ײ͹���_����
(
  �ײ�ID_IN   �����ײ͹���.�ײ�ID%Type,   
  ����ID_IN   �����ײ͹���.����ID%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id �����ײ͹���.ID%Type;
Begin
  select �����ײ͹���_Id.Nextval into v_id from dual;
  
  insert into �����ײ͹���(ID, �ײ�ID,����ID) values(v_id, �ײ�ID_IN, ����ID_IN);
  
  commit;
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����ײ͹���_����;
/


--�����ײ�ID,ɾ���ײ͹����Ŀ���
CREATE OR REPLACE procedure Zl_�����ײ͹���_ɾ��
(
  �ײ�ID_IN   �����ײ͹���.�ײ�ID%Type
)Is
Begin
  --ɾ���������ײ���Ϣ          
  delete �����ײ͹��� where �ײ�ID=�ײ�ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����ײ͹���_ɾ��;
/


--���ݹ���IDɾ���ײ͹����Ŀ���
CREATE OR REPLACE procedure Zl_�����ײ͹���_ɾ��1
(
  �ײ͹���ID_IN   �����ײ͹���.ID%Type
)Is
Begin
  --ɾ���������ײ���Ϣ          
  delete �����ײ͹��� where ID=�ײ͹���ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����ײ͹���_ɾ��1;
/
 


--�걾ȡ��====================================================================================================


--�����Ѹ�
CREATE OR REPLACE function Zl_�����Ѹ�_��ʼ
(
  �걾ID_IN    �����Ѹ���Ϣ.�걾ID%Type,   
  ��ʼʱ��_IN  �����Ѹ���Ϣ.��ʼʱ��%Type,
  ����ʱ��_IN  �����Ѹ���Ϣ.����ʱ��%Type,
  ����Ա_IN    �����Ѹ���Ϣ.����Ա%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id �����Ѹ���Ϣ.ID%Type;
Begin
  select �����Ѹ���Ϣ_Id.Nextval into v_id from dual;
  
  insert into �����Ѹ���Ϣ(ID, �걾ID,��ʼʱ��,����ʱ��,��ǰ�״�,����Ա,���״̬)
  values(v_id, �걾ID_IN, ��ʼʱ��_IN, ����ʱ��_IN, 1, ����Ա_IN, 0);
  
  commit;
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����Ѹ�_��ʼ;
/



--�����Ѹƻ���
CREATE OR REPLACE procedure Zl_�����Ѹ�_����
(
  ID_IN        �����Ѹ���Ϣ.ID %Type,       
  ��ʼʱ��_IN  �����Ѹ���Ϣ.��ʼʱ��%Type,
  ����ʱ��_IN  �����Ѹ���Ϣ.����ʱ��%Type
)Is
Begin
  
  update �����Ѹ���Ϣ
  set ��ʼʱ��=��ʼʱ��_IN,����ʱ��=����ʱ��_IN,��ǰ�״�=��ǰ�״�+1
  where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����Ѹ�_����;
/


--�����ѸƳ���
CREATE OR REPLACE procedure Zl_�����Ѹ�_����
(
  ID_IN        �����Ѹ���Ϣ.ID %Type
) Is
Begin
  
  Delete �����Ѹ���Ϣ where ID=ID_IN and ���״̬=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����Ѹ�_����;
/


--�����Ѹ����
CREATE OR REPLACE procedure Zl_�����Ѹ�_���
(
  ID_IN        �����Ѹ���Ϣ.ID %Type
)Is
Begin
  
  update �����Ѹ���Ϣ set ���״̬=1 where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����Ѹ�_���;
/




--������ȡ��
CREATE OR REPLACE function Zl_����ȡ��_����
(
  �����_IN      ����ȡ����Ϣ.�����%Type,   
  ����ID_IN      ����ȡ����Ϣ.����ID%Type,
  �걾ID_IN      ����ȡ����Ϣ.�걾ID%Type,
  �걾����_IN    ����ȡ����Ϣ.�걾����%Type,
  ȡ��λ��_IN    ����ȡ����Ϣ.ȡ��λ��%Type,
  ������_IN      ����ȡ����Ϣ.������%Type,
  ��ȡҽʦ_IN    ����ȡ����Ϣ.��ȡҽʦ%Type,
  ��ȡҽʦ_IN    ����ȡ����Ϣ.��ȡҽʦ%Type,
  ��¼ҽʦ_IN    ����ȡ����Ϣ.��¼ҽʦ%Type,
  ȡ��ʱ��_IN    ����ȡ����Ϣ.ȡ��ʱ��%Type  
) return varchar2 Is
PRAGMA AUTONOMOUS_TRANSACTION;

v_id ����ȡ����Ϣ.�Ŀ�ID%Type;
v_seqNum ����ȡ����Ϣ.���%Type;

Begin                        
  --��ȡ���Ŀ�����  
  begin
    select  nvl(max(���), 0) into v_seqNum from ����ȡ����Ϣ where �����=�����_IN;
  exception
    When Others Then v_id := 0;	            
  end;    
  
  v_seqNum := v_seqNum + 1;
  select ����ȡ����Ϣ_�Ŀ�ID.Nextval into v_id from dual;
  
  --д��ȡ�ļ�¼    
  insert into ����ȡ����Ϣ(�Ŀ�ID, ���, �����, ����ID, �걾ID, �걾����,ȡ��λ��,������,��ȡҽʦ,��ȡҽʦ,��¼ҽʦ,ȡ��ʱ��)
  values(v_id, v_seqNum, �����_IN, ����ID_IN, �걾ID_IN, �걾����_IN, ȡ��λ��_IN, ������_IN,��ȡҽʦ_IN,��ȡҽʦ_IN,��¼ҽʦ_IN,ȡ��ʱ��_IN);
  
  --д����Ƭ��¼
  insert into ������Ƭ��Ϣ(ID,�����,�Ŀ�ID,����ID,��Ƭ����,��Ƭ��ʽ,��Ƭ��,��ǰ״̬)
  values(������Ƭ��Ϣ_ID.NEXTVAL,�����_IN,v_id,����ID_IN,0,0,������_IN,0);
  
  commit; 
  
  return v_id || '-' || v_seqNum;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ȡ��_����;
/



CREATE OR REPLACE Procedure Zl_����ȡ��_�������
(
  �Ŀ�ID_IN      ����ȡ����Ϣ.�����%Type,   
  ȡ��λ��_IN    ����ȡ����Ϣ.ȡ��λ��%Type,
  ������_IN      ����ȡ����Ϣ.������%Type,
  ��ȡҽʦ_IN    ����ȡ����Ϣ.��ȡҽʦ%Type,
  ��ȡҽʦ_IN    ����ȡ����Ϣ.��ȡҽʦ%Type 
)Is
Begin
  --����ȡ����Ϣ      
  update ����ȡ����Ϣ
  set ȡ��λ��=ȡ��λ��_IN,������=������_IN,��ȡҽʦ=��ȡҽʦ_IN,��ȡҽʦ=��ȡҽʦ_IN
  where �Ŀ�ID=�Ŀ�ID_IN;

  --������Ƭ��Ϣ
  update ������Ƭ��Ϣ set ��Ƭ��=������_IN where �Ŀ�ID=�Ŀ�ID_IN and ��ǰ״̬=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ȡ��_�������;
/



--����ϸ��ȡ��
CREATE OR REPLACE function Zl_����ȡ��_ϸ��
(
  �����_IN      ����ȡ����Ϣ.�����%Type,   
  ����ID_IN      ����ȡ����Ϣ.����ID%Type,
  �걾ID_IN      ����ȡ����Ϣ.�걾ID%Type,
  �걾����_IN    ����ȡ����Ϣ.�걾����%Type,
  ��״_IN        ����ȡ����Ϣ.��״%Type,
  ��ɫ_IN        ����ȡ����Ϣ.��ɫ%Type, 
  ����_IN        ����ȡ����Ϣ.����%Type,   
  �걾��_IN      ����ȡ����Ϣ.�걾��%Type,
  ϸ������_IN    ����ȡ����Ϣ.������%Type,
  ��ȡҽʦ_IN    ����ȡ����Ϣ.��ȡҽʦ%Type,
  ��ȡҽʦ_IN    ����ȡ����Ϣ.��ȡҽʦ%Type,
  ��¼ҽʦ_IN    ����ȡ����Ϣ.��¼ҽʦ%Type,
  ȡ��ʱ��_IN    ����ȡ����Ϣ.ȡ��ʱ��%Type  
) return varchar2 Is
PRAGMA AUTONOMOUS_TRANSACTION;

v_id ����ȡ����Ϣ.�Ŀ�ID%Type;
v_seqNum ����ȡ����Ϣ.���%Type;
v_slicesCount number;

Begin
  --��ȡ���Ŀ�����  
  begin
    select  nvl(max(���), 0) into v_seqNum from ����ȡ����Ϣ where �����=�����_IN;
  exception
    When Others Then v_id := 0;	            
  end;    
  
  v_seqNum := v_seqNum + 1;
  select ����ȡ����Ϣ_�Ŀ�ID.Nextval into v_id from dual;
  
  
  --д��ȡ�ļ�¼    
  insert into ����ȡ����Ϣ(�Ŀ�ID,���, �����, ����ID, �걾ID, �걾����,��״,��ɫ,����,�걾��,������,��ȡҽʦ,��ȡҽʦ,��¼ҽʦ,ȡ��ʱ��)
  values(v_id, v_seqNum, �����_IN, ����ID_IN, �걾ID_IN, �걾����_IN, ����_IN,��ɫ_IN,����_IN, �걾��_IN,ϸ������_IN,��ȡҽʦ_IN,��ȡҽʦ_IN,��¼ҽʦ_IN,ȡ��ʱ��_IN);

  if ϸ������_IN is null then
     v_slicesCount := 1;
  elsif ϸ������_IN <= 0 then
     v_slicesCount := 1;
  else
     v_slicesCount := ϸ������_IN;
  end if;  

  --д����Ƭ��¼
  insert into ������Ƭ��Ϣ(ID,�����,�Ŀ�ID,����ID,��Ƭ����,��Ƭ��ʽ,��Ƭ��,��ǰ״̬)
  values(������Ƭ��Ϣ_ID.NEXTVAL,�����_IN,v_id,����ID_IN,2, 0, v_slicesCount, 0);
  
  commit; 
  
  return v_id || '-' || v_seqNum;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ȡ��_ϸ��;
/


CREATE OR REPLACE Procedure Zl_����ȡ��_ϸ������
(
  �Ŀ�ID_IN      ����ȡ����Ϣ.�����%Type,   
  ��״_IN        ����ȡ����Ϣ.��״%Type,
  ��ɫ_IN        ����ȡ����Ϣ.��ɫ%Type, 
  ����_IN        ����ȡ����Ϣ.����%Type, 
  �걾��_IN      ����ȡ����Ϣ.�걾��%Type,   
  ϸ������_IN    ����ȡ����Ϣ.������%Type,
  ��ȡҽʦ_IN    ����ȡ����Ϣ.��ȡҽʦ%Type,
  ��ȡҽʦ_IN    ����ȡ����Ϣ.��ȡҽʦ%Type,
  ȡ��ʱ��_IN    ����ȡ����Ϣ.ȡ��ʱ��%Type  
)Is
v_slicesCount number;
Begin
  --����ȡ����Ϣ      
  update ����ȡ����Ϣ
  set ��״=��״_IN,��ɫ=��ɫ_IN,����=����_IN,�걾��=�걾��_IN,������=ϸ������_IN,��ȡҽʦ=��ȡҽʦ_IN,��ȡҽʦ=��ȡҽʦ_IN
  where �Ŀ�ID=�Ŀ�ID_IN;

  if ϸ������_IN is null then
     v_slicesCount := 1;
  elsif ϸ������_IN <= 0 then
     v_slicesCount := 1;
  else
     v_slicesCount := ϸ������_IN;     
  end if;  
  
  
  --������Ƭ��Ϣ
  update ������Ƭ��Ϣ  set ��Ƭ��=v_slicesCount where �Ŀ�ID=�Ŀ�ID_IN and ��ǰ״̬=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ȡ��_ϸ������;
/



--�������ȡ��
CREATE OR REPLACE function Zl_����ȡ��_����
(
  �����_IN      ����ȡ����Ϣ.�����%Type,   
  ����ID_IN      ����ȡ����Ϣ.����ID%Type,
  �걾ID_IN      ����ȡ����Ϣ.�걾ID%Type,
  �걾����_IN    ����ȡ����Ϣ.�걾����%Type,
  ȡ��λ��_IN    ����ȡ����Ϣ.ȡ��λ��%Type,
  �Ƿ����_IN    ����ȡ����Ϣ.�Ƿ����%Type,
  ������_IN      ����ȡ����Ϣ.������%Type,
  ��ȡҽʦ_IN    ����ȡ����Ϣ.��ȡҽʦ%Type,
  ��ȡҽʦ_IN    ����ȡ����Ϣ.��ȡҽʦ%Type,
  ��¼ҽʦ_IN    ����ȡ����Ϣ.��¼ҽʦ%Type,
  ȡ��ʱ��_IN    ����ȡ����Ϣ.ȡ��ʱ��%Type  
) return varchar2 Is
PRAGMA AUTONOMOUS_TRANSACTION;

v_id ����ȡ����Ϣ.�Ŀ�ID%Type;
v_seqNum ����ȡ����Ϣ.���%Type;

Begin
  --��ȡ���Ŀ�����  
  begin
    select  nvl(max(���), 0) into v_seqNum from ����ȡ����Ϣ where �����=�����_IN;
  exception
    When Others Then v_id := 0;	            
  end;    
  
  v_seqNum := v_seqNum + 1;
  select ����ȡ����Ϣ_�Ŀ�ID.Nextval into v_id from dual;
  
  --д��ȡ�ļ�¼    
  insert into ����ȡ����Ϣ(�Ŀ�ID, ���, �����, ����ID, �걾ID, �걾����,ȡ��λ��,�Ƿ����,������,��ȡҽʦ,��ȡҽʦ,��¼ҽʦ,ȡ��ʱ��)
  values(v_id, v_seqNum, �����_IN, ����ID_IN, �걾ID_IN, �걾����_IN, ȡ��λ��_IN, �Ƿ����_IN,������_IN,��ȡҽʦ_IN,��ȡҽʦ_IN,��¼ҽʦ_IN,ȡ��ʱ��_IN);


  --д����Ƭ��¼
  insert into ������Ƭ��Ϣ(ID,�����,�Ŀ�ID,����ID,��Ƭ����,��Ƭ��ʽ,��Ƭ��,��ǰ״̬)
  values(������Ƭ��Ϣ_ID.NEXTVAL,�����_IN,v_id,����ID_IN,1,0,������_IN,0);
  
  
  commit; 
  
  return v_id || '-' || v_seqNum;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ȡ��_����;
/



CREATE OR REPLACE Procedure Zl_����ȡ��_��������
(
  �Ŀ�ID_IN      ����ȡ����Ϣ.�����%Type,   
  ȡ��λ��_IN    ����ȡ����Ϣ.ȡ��λ��%Type,
  �Ƿ����_IN    ����ȡ����Ϣ.�Ƿ����%Type,
  ������_IN      ����ȡ����Ϣ.������%Type,
  ��ȡҽʦ_IN    ����ȡ����Ϣ.��ȡҽʦ%Type,
  ��ȡҽʦ_IN    ����ȡ����Ϣ.��ȡҽʦ%Type 
)Is
Begin
  
  --����ȡ�ļ�¼      
  update ����ȡ����Ϣ
  set ȡ��λ��=ȡ��λ��_IN,�Ƿ����=�Ƿ����_IN,������=������_IN,��ȡҽʦ=��ȡҽʦ_IN,��ȡҽʦ=��ȡҽʦ_IN
  where �Ŀ�ID=�Ŀ�ID_IN;


  --������Ƭ��Ϣ
  update ������Ƭ��Ϣ  set ��Ƭ��=������_IN  where �Ŀ�ID=�Ŀ�ID_IN and ��ǰ״̬=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ȡ��_��������;
/



--ɾ������ȡ�ļ�¼
CREATE OR REPLACE Procedure Zl_����ȡ��_ɾ��
(
  �Ŀ�ID_IN      ����ȡ����Ϣ.�Ŀ�ID%Type
)Is
Begin
      
  delete ����ȡ����Ϣ where �Ŀ�ID=�Ŀ�ID_IN;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ȡ��_ɾ��;
/

  


--����ȡ��ʱ�������Ϣ
CREATE OR REPLACE Procedure Zl_����ȡ��_��Ϣ����
(
  �����_IN      ��������Ϣ.�����%Type,   
  �޼�����_IN    ��������Ϣ.�޼�����%Type,
  ʣ��λ��_IN    ��������Ϣ.ʣ��λ��%Type
)Is
Begin
        
  update ��������Ϣ  set �޼�����=�޼�����_IN,ʣ��λ��=ʣ��λ��_IN  where �����=�����_IN;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ȡ��_��Ϣ����;
/


--ȡ��ȷ��
CREATE OR REPLACE Procedure Zl_����ȡ��_ȷ��
(
  �����_IN      ��������Ϣ.�����%Type
)Is
Begin
        
  --���²�����״̬
  update ��������Ϣ set ��ǰ����=2 where �����=�����_IN;
  
  --����в�ȡ���룬���������״̬  
  update ����������Ϣ set ����״̬=1 where ����״̬=0 and ��������=8 and �����=�����_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ȡ��_ȷ��;
/




--������Ƭ====================================================================================================

--������Ƭ�������ܸò������������Ƭ��
CREATE OR REPLACE Procedure Zl_������Ƭ_����
(
  �����_IN      ����ȡ����Ϣ.�����%Type,
  ��Ƭ��_IN      ������Ƭ��Ϣ.��Ƭ��%Type
)Is
Begin
        
  --������Ƭ״̬(Ϊ�������Ƭ���ܽ���)
  update ������Ƭ��Ϣ set ��ǰ״̬=1,��Ƭ��=��Ƭ��_IN  where ����� = �����_IN and ��ǰ״̬=0;
  
  --��������״̬��0-�����룬1-�ѽ��ܣ�2-����ɣ�
  update ����������Ϣ set ����״̬ = 1 
  where ����ID=(select distinct ����ID from ������Ƭ��Ϣ  where �����=�����_IN and ��ǰ״̬=0) 
        and ����״̬=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_������Ƭ_����;
/



--������Ƭ�������ܵ�ǰ�Ŀ����Ƭ��
CREATE OR REPLACE Procedure Zl_������Ƭ_�嵥��ӡ
(
  ID_IN      ������Ƭ��Ϣ.�Ŀ�ID%Type
)Is
Begin        
  
  --�����嵥״̬
  update ������Ƭ��Ϣ  set �嵥״̬=1 where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_������Ƭ_�嵥��ӡ;
/



--ȷ����Ƭ
CREATE OR REPLACE Procedure Zl_������Ƭ_ȷ��
(
  �����_IN      ����ȡ����Ϣ.�����%Type,
  ��Ƭʱ��_IN    ������Ƭ��Ϣ.��Ƭʱ��%Type
)Is
Begin
        
  --������Ƭ״̬����δ��ɵ���Ƭ��¼���޸�Ϊ�����״̬��δ���ܵ���Ƭ���ܽ���ȷ�ϣ�
  update ������Ƭ��Ϣ set ��ǰ״̬=2,��Ƭʱ��=��Ƭʱ��_IN where �����=�����_IN and ��ǰ״̬=1;  
  --where �Ŀ�id in(select �Ŀ�id from ����ȡ����Ϣ where �����=�����_IN) and ��ǰ״̬<>2;

  
  --�޸ļ��ĵ�ǰ����Ϊ��һ�׶Σ���Ƭ��ɺ����һ�׶�Ϊ��ϣ�
  update ��������Ϣ  set ��ǰ����=3 where �����=�����_IN;  
  
  
  --��������״̬���������������£�û����ִ�� 0-�����룬1-�ѽ��ܣ�2-����ɣ�
  update ����������Ϣ set ����״̬ = 2, ���ʱ��=��Ƭʱ��_IN 
  where ����ID=(select distinct ����ID from ������Ƭ��Ϣ where �����=�����_IN and ��ǰ״̬=1)  
        and ����״̬=1; 
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_������Ƭ_ȷ��;
/


--ȷ����Ƭ
/*CREATE OR REPLACE Procedure Zl_������Ƭ_ȷ��1
(
  �Ŀ�ID_IN      ������Ƭ��Ϣ.�Ŀ�ID%Type,
  ��Ƭʱ��_IN    ������Ƭ��Ϣ.��Ƭʱ��%Type
)Is
  v_count number;
Begin
        
  --������Ƭ״̬����δ��ɵ���Ƭ��¼���޸�Ϊ�����״̬��
  update ������Ƭ��Ϣ 
  set ��ǰ״̬=2,��Ƭʱ��=��Ƭʱ��_IN 
  where �Ŀ�ID = �Ŀ�ID_IN and ��ǰ״̬<>2;
  --where �Ŀ�id in(select �Ŀ�id from ����ȡ����Ϣ where �����=�����_IN) and ��ǰ״̬=1;
  
  v_count := 0;
  begin
    select sum(��Ƭ��) into v_count from ������Ƭ��Ϣ a, ����ȡ����Ϣ b
    where a.�Ŀ�id = b.�Ŀ�id and b.����� = (select ����� from ����ȡ����Ϣ where �Ŀ�id=�Ŀ�ID_IN) and a.��ǰ״̬ <> 2;
  exception
    when others then v_count := 0;          
  end;    
  
  --�޸ļ��ĵ�ǰ���̺���Ƭ״̬(�����вĿ����ȷ�Ϻ��޸ļ���ִ��״̬) 
  if v_count = 0 then
    update ��������Ϣ  set ��ǰ����=3, ��Ƭ״̬=2  where �����=(select ����� from ����ȡ����Ϣ where �Ŀ�id=�Ŀ�ID_IN);  
  end if;  
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_������Ƭ_ȷ��1;
*/



--�����ӳ�===================================================================================================


--��ӱ����ӳټ�¼
CREATE OR REPLACE function Zl_�������ӳ�_����
(
  �����_IN      �������ӳ�.�����%Type,   
  �ӳ�ԭ��_IN    �������ӳ�.�ӳ�ԭ��%Type,
  �ӳ�����_IN    �������ӳ�.�ӳ�����%Type,
  ��ʱ���_IN    �������ӳ�.��ʱ���%Type,
  ת����_IN      �������ӳ�.ת����%Type,  
  �Ǽ���_IN      �������ӳ�.�Ǽ���%Type,
  �Ǽ�ʱ��_IN    �������ӳ�.�Ǽ�ʱ��%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id �������ӳ�.ID%Type;
Begin
  --��ȡ�ӳٱ���ID
  select �������ӳ�_ID.NEXTVAL into v_id from dual;

  
  --д�뱨���ӳټ�¼    
  insert into �������ӳ�(ID, �����, �ӳ�ԭ��,�ӳ�����,��ʱ���,ת����,�Ǽ���,�Ǽ�ʱ��,��ǰ״̬)
  values(v_id, �����_IN, �ӳ�ԭ��_IN, �ӳ�����_IN, ��ʱ���_IN, ת����_IN,�Ǽ���_IN,�Ǽ�ʱ��_IN, 0);
  
  commit; 
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�������ӳ�_����;
/


--���±����ӳټ�¼
CREATE OR REPLACE procedure Zl_�������ӳ�_����
(
  ID_IN          �������ӳ�.ID%Type,   
  �ӳ�ԭ��_IN    �������ӳ�.�ӳ�ԭ��%Type,
  �ӳ�����_IN    �������ӳ�.�ӳ�����%Type,
  ��ʱ���_IN    �������ӳ�.��ʱ���%Type,
  ת����_IN      �������ӳ�.ת����%Type
) Is
Begin
  
  update �������ӳ�
  set �ӳ�ԭ��=�ӳ�ԭ��_IN,�ӳ�����=�ӳ�����_IN,��ʱ���=��ʱ���_IN,ת����=ת����_IN
  where ID=ID_IN;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�������ӳ�_����;
/


--ɾ�������ӳ�
CREATE OR REPLACE procedure Zl_�������ӳ�_ɾ��
(
  ID_IN          �������ӳ�.ID%Type
) Is
Begin

  --ɾ��δ��ӡ���ӳٱ��棬����Ѵ�ӡ����ɾ��
  delete �������ӳ� where ID=ID_IN; -- and ��ǰ״̬=0;  

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�������ӳ�_ɾ��;
/

--��ӡ�����ӳ�
CREATE OR REPLACE procedure Zl_�������ӳ�_��ӡ
(
  ID_IN          �������ӳ�.ID%Type
) Is
Begin

  --����ӡ���޸ı����ӳټ�¼�ĵ�ǰ״̬
  update �������ӳ� set ��ǰ״̬=1 where ID=ID_IN;  

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�������ӳ�_��ӡ;
/


--���̱���===================================================================================================


--��ӹ��̱����¼
CREATE OR REPLACE function Zl_������̱���_����
(
  �����_IN      ������̱���.�����%Type,   
  �걾����_IN    ������̱���.�걾����%Type,
  ��������_IN    ������̱���.��������%Type,
  ������_IN    ������̱���.������%Type,
  �����_IN    ������̱���.�����%Type,  
  ����ҽʦ_IN    ������̱���.����ҽʦ%Type,
  ��������_IN    ������̱���.��������%Type,
  ����ͼ��_IN    ������̱���.����ͼ��%Type,
  ��ע_IN        ������̱���.��ע%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id ������̱���.ID%Type;
Begin
  --��ȡ���̱���ID
  select ������̱���_ID.NEXTVAL into v_id from dual;

  
  --д����̱����¼    
  insert into ������̱���(ID, �����, �걾����,��������,�����,������,����ͼ��,����ҽʦ,��������,��ǰ״̬,��ע)
  values(v_id, �����_IN, �걾����_IN, ��������_IN, �����_IN, ������_IN,����ͼ��_IN,����ҽʦ_IN,��������_IN, 0,��ע_IN);
  
  commit; 
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_������̱���_����;
/ 

--���¹��̱���
CREATE OR REPLACE procedure Zl_������̱���_����
(
  ID_IN          ������̱���.ID%Type,   
  �걾����_IN    ������̱���.�걾����%Type,
  ��������_IN    ������̱���.��������%Type,
  ������_IN    ������̱���.������%Type,
  �����_IN    ������̱���.�����%Type,  
  ����ͼ��_IN    ������̱���.����ͼ��%Type,
  ��ע_IN        ������̱���.��ע%Type
)Is
Begin

  --���¹��̱����¼    
  update ������̱���
  set �걾����=�걾����_IN, ��������=��������_IN,������=������_IN,
      �����=�����_IN, ����ͼ��=����ͼ��_IN,��ע=��ע_IN
  where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_������̱���_����;
/


--ɾ�����̱���
CREATE OR REPLACE procedure Zl_������̱���_ɾ��
(
  ID_IN          ������̱���.ID%Type
)Is
Begin

  --ɾ�����̱����¼    
  delete ������̱��� where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_������̱���_ɾ��;
/


--���̱���״̬����
CREATE OR REPLACE procedure Zl_������̱���_״̬
(
  ID_IN          ������̱���.ID%Type,
  ��ǰ״̬_IN    ������̱���.��ǰ״̬%Type  
)Is
Begin

  --ɾ�����̱����¼    
  update ������̱��� set ��ǰ״̬=��ǰ״̬_IN where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_������̱���_״̬;
/

 


--�������===================================================================================================



--��Ӽ������
CREATE OR REPLACE function Zl_��������_����
(
  �����_IN      ����������Ϣ.�����%Type,   
  ������_IN      ����������Ϣ.������%Type,
  ����ʱ��_IN    ����������Ϣ.����ʱ��%Type,
  ��������_IN    ����������Ϣ.��������%Type,
  ��������_IN    ����������Ϣ.��������%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id ����������Ϣ.����ID%Type;
v_procedure ��������Ϣ.��ǰ����%Type;
Begin
  --��ȡ����ID
  select ����������Ϣ_����ID.NEXTVAL into v_id from dual;

  
  --д�������¼    
  insert into ����������Ϣ(����ID, �����, ������,����ʱ��,��������,��������,����״̬,�Ƿ��ӡ)
  values(v_id, �����_IN, ������_IN, ����ʱ��_IN, ��������_IN, ��������_IN,0,0);
  
  --���¼�����
  case 
    when ��������_IN = 0 then v_procedure := 4;  --�����黯
    when ��������_IN = 1 then v_procedure := 5;  --����Ⱦɫ
    when ��������_IN = 2 then v_procedure := 6;  --���Ӳ���
    when ��������_IN = 3 then v_procedure := 9;  --����Ƭ
    when ��������_IN = 4 then v_procedure := 8;  --��ȡ��
    else v_procedure := -1;
  end case;
  
  if v_procedure <= 0 then 
    Raise_Application_Error(-20101, '[ZLSOFT] �������ʱ��������Чȡ�ò�������Ϣ�еĵ�ǰ���̣���ִֹ�С�[ZLSOFT]');
  end if;
  
  
  update ��������Ϣ set ��ǰ����=v_procedure where �����=�����_IN;
    
  
  commit; 
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_��������_����;
/ 



--ɾ���������
CREATE OR REPLACE procedure Zl_��������_ɾ��
(
  ����ID_IN          ����������Ϣ.����ID%Type
)Is
Begin

  --ɾ�����̱����¼    
  delete ����������Ϣ where ����ID=����ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_��������_ɾ��;
/


--����ؼ���Ŀ
CREATE OR REPLACE function Zl_��������_�ؼ���Ŀ_����
(
  �����_IN      �����ؼ���Ϣ.�����%Type,         
  �Ŀ�ID_IN      �����ؼ���Ϣ.�Ŀ�ID%Type,   
  ����ID_IN      �����ؼ���Ϣ.����ID%Type,
  ����ID_IN      �����ؼ���Ϣ.����ID%Type,
  �ؼ�����_IN    �����ؼ���Ϣ.�ؼ�����%Type,
  �Ƿ���_IN    number := 0
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id �����ؼ���Ϣ.ID%Type;
v_���� �����ؼ���Ϣ.��������%Type;
Begin
  --��ȡ�ؼ���ϢID
  select �����ؼ���Ϣ_ID.NEXTVAL into v_id from dual;

  v_���� := -1;
  if �Ƿ���_IN = 0 then
     v_���� := 0;
  end if;
  
  --д�������¼    
  insert into �����ؼ���Ϣ(ID,�����,�Ŀ�ID,����ID,����ID,�ؼ�����,��������,��ǰ״̬,�嵥״̬)
  values(v_id,�����_IN, �Ŀ�ID_IN, ����ID_IN, ����ID_IN, �ؼ�����_IN, v_����, 0,0);
  
  --���¼�����
  if v_���� = -1 then
     update ��������Ϣ set ��ǰ����=decode(�ؼ�����_IN, 0, 4, 1, 5, 6) where �����=�����_IN;
  end if;
  
  commit; 
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_��������_�ؼ���Ŀ_����;
/ 


--ɾ���ؼ���Ŀ
CREATE OR REPLACE procedure Zl_��������_�ؼ���Ŀ_ɾ��
(
  ID_IN          �����ؼ���Ϣ.ID%Type
)Is
Begin

  --ɾ���ؼ���Ŀ(ֻ�����������Ŀ������ɾ��)
  delete �����ؼ���Ϣ where ID=ID_IN and ��ǰ״̬=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_��������_�ؼ���Ŀ_ɾ��;
/


--�ؼ���Ŀ����
CREATE OR REPLACE function Zl_��������_�ؼ���Ŀ_����
(
  ID_IN          �����ؼ���Ϣ.ID%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
  cursor c_SpeExamInf is
         select �����,�Ŀ�ID,����ID,�ؼ����� from �����ؼ���Ϣ where Id=ID_IN;
 
  r_SpeExamInf  c_SpeExamInf%RowType;         
  v_newid      �����ؼ���Ϣ.ID%Type;
  v_count   �����ؼ���Ϣ.��������%Type;
  
  v_Error    varchar(255);
  Err_Custom Exception;
Begin
    
  
  Open c_SpeExamInf;
  Fetch c_SpeExamInf Into r_SpeExamInf;
    
  If c_SpeExamInf%Rowcount = 0 Then
    Close c_SpeExamInf;
    v_Error := '������ȷ��ȡ�����ؼ���Ϣ��������ݣ�������ĿID�Ƿ�Ϊ��Ч���ݡ�';
    Raise Err_Custom;
  End If;  
  
  v_count := 0;
  begin
    select nvl(max(��������), 0) into v_count from �����ؼ���Ϣ 
    where �Ŀ�Id=r_SpeExamInf.�Ŀ�ID and ����Id=r_SpeExamInf.����ID and �ؼ�����=r_SpeExamInf.�ؼ�����; 
  exception
    when others then v_count := 0;           
  end;
  
  select �����ؼ���Ϣ_ID.NEXTVAL into v_newid from dual;

  --�ؼ���Ŀ����
  insert into �����ؼ���Ϣ(ID, �����,�Ŀ�ID,����ID,����ID,�ؼ�����,��������,��ǰ״̬,�嵥״̬) 
  select v_newid as ID,�����,�Ŀ�ID,����ID,����ID,�ؼ�����, v_count+1,0,0 from �����ؼ���Ϣ where ID=ID_IN;
  
  
  --���¼�����
  update ��������Ϣ set ��ǰ����=decode(r_SpeExamInf.�ؼ�����, 0, 4, 1, 5, 6) where �����=r_SpeExamInf.�����;

  commit;
  
  return v_newid;
  
  close c_SpeExamInf;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_��������_�ؼ���Ŀ_����;
/




--������Ƭ��Ŀ
CREATE OR REPLACE function Zl_��������_��Ƭ��Ŀ_����
(
  �����_IN      ������Ƭ��Ϣ.�����%Type,    
  �Ŀ�ID_IN      ������Ƭ��Ϣ.�Ŀ�ID%Type,   
  ����ID_IN      ������Ƭ��Ϣ.����ID%Type,
  ��Ƭ����_IN    ������Ƭ��Ϣ.��Ƭ����%Type,
  ��Ƭ��ʽ_IN    ������Ƭ��Ϣ.��Ƭ��ʽ%Type,
  ��Ƭ����_IN    ������Ƭ��Ϣ.��Ƭ��%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id       ������Ƭ��Ϣ.ID%Type;
Begin
  --��ȡ�ؼ���ϢID
  select ������Ƭ��Ϣ_ID.NEXTVAL into v_id from dual;

  
  --д����Ƭ��¼    
  insert into ������Ƭ��Ϣ(ID, �����, �Ŀ�ID,����ID,��Ƭ����,��Ƭ��,��Ƭ��ʽ,��ǰ״̬,�嵥״̬)
  values(v_id, �����_IN, �Ŀ�ID_IN, ����ID_IN, ��Ƭ����_IN, ��Ƭ����_IN, ��Ƭ��ʽ_IN, 0,0);
  
  commit; 
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_��������_��Ƭ��Ŀ_����;
/ 


--ɾ����Ƭ��Ŀ
CREATE OR REPLACE procedure Zl_��������_��Ƭ��Ŀ_ɾ��
(
  ID_IN          ������Ƭ��Ϣ.ID%Type
)Is
Begin

  --ɾ����Ƭ��Ŀ(ֻ��δ�������Ŀ������ɾ��)
  delete ������Ƭ��Ϣ where ID=ID_IN and ��ǰ״̬=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_��������_��Ƭ��Ŀ_ɾ��;
/




--�������===================================================================================================

--��Ӳ������
CREATE OR REPLACE function Zl_�������_����
(
  �����_IN      ���������Ϣ.�����%Type,   
  ����ҽʦ_IN    ���������Ϣ.����ҽʦ%Type,
  ���ﵥλ_IN    ���������Ϣ.���ﵥλ%Type,
  ����ҽʦ_IN    ���������Ϣ.����ҽʦ%Type,
  ����ʱ��_IN    ���������Ϣ.����ʱ��%Type,
  ��ֹʱ��_IN    ���������Ϣ.��ֹʱ��%Type,
  ��������_IN    ���������Ϣ.��������%Type,
  �������_IN    ���������Ϣ.�������%Type
) return number Is
PRAGMA AUTONOMOUS_TRANSACTION;
v_id ���������Ϣ.ID%Type;
Begin
  --��ȡ����ID
  select ���������Ϣ_ID.NEXTVAL into v_id from dual;

  
  --д�������¼    
  insert into ���������Ϣ(id,�����,����ҽʦ,���ﵥλ,����ҽʦ,����ʱ��,��ֹʱ��,��������,�������,��ǰ״̬)
  values(v_id, �����_IN, ����ҽʦ_IN, ���ﵥλ_IN, ����ҽʦ_IN, ����ʱ��_IN,��ֹʱ��_IN,��������_IN,�������_IN,0);
  
  commit; 
  
  return v_id;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�������_����;
/ 


--���ﷴ��
CREATE OR REPLACE procedure Zl_�������_����
(
  ID_IN          ���������Ϣ.ID%Type,
  ��Ͻ��_IN    ���������Ϣ.��Ͻ��%Type,
  ������_IN    ���������Ϣ.������%Type,
  ���ʱ��_IN    ���������Ϣ.���ʱ��%Type,
  ��ע_IN        ���������Ϣ.��ע%Type
)Is
Begin

  --���²�������¼ ���������¼��������״̬�޸�Ϊ���״̬��
  update ���������Ϣ 
  set ��Ͻ�� = ��Ͻ��_IN,������=������_IN,���ʱ��=���ʱ��_IN,��ע=��ע_IN,��ǰ״̬=2
  where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�������_����;
/



--ɾ��������Ĳ��������Ϣ
CREATE OR REPLACE procedure Zl_�������_ɾ��
(
  ID_IN          ���������Ϣ.ID%Type
)Is
Begin

  --ɾ����������¼ (����������ѳ����Ļ����¼�ɱ�ɾ��)
  delete ���������Ϣ where ID=ID_IN and (��ǰ״̬=0 or ��ǰ״̬=1);
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�������_ɾ��;
/



--���û��ﵱǰ״̬
CREATE OR REPLACE procedure Zl_�������_״̬
(
  ID_IN          ���������Ϣ.ID%Type,
  ��ǰ״̬_IN    ���������Ϣ.��ǰ״̬%Type  
)Is
Begin

  --���û����¼�ĵ�ǰ״̬
  update ���������Ϣ set ��ǰ״̬=��ǰ״̬_IN where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�������_״̬;
/



--�����ؼ�===================================================================================================


--�����ؼ촦��
CREATE OR REPLACE Procedure Zl_�����ؼ�_����
(
  �����_IN      ����ȡ����Ϣ.�����%Type,
  �ؼ�����_IN    �����ؼ���Ϣ.�ؼ�����%Type, 
  �ؼ�ҽʦ_IN    �����ؼ���Ϣ.�ؼ�ҽʦ%Type
)Is
Begin
  --��������״̬��0-�����룬1-�ѽ��ܣ�2-����ɣ�
  update ����������Ϣ set ����״̬ = 1 
  where ����ID=(select distinct ����ID from �����ؼ���Ϣ where �����=�����_IN and �ؼ�����=�ؼ�����_IN and ��ǰ״̬=0)
        and ����״̬=0;
          
        
  --�����ؼ�״̬��ֻ�е�ǰ״̬Ϊ0���ؼ���Ϣ�Ž��и��£�
  update �����ؼ���Ϣ 
  set ��ǰ״̬=1,�ؼ�ҽʦ=�ؼ�ҽʦ_IN 
  where �Ŀ�id in(select �Ŀ�id from ����ȡ����Ϣ where �����=�����_IN) and �ؼ�����=�ؼ�����_IN and ��ǰ״̬=0;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����ؼ�_����;
/


--�����ؼ촦��
CREATE OR REPLACE Procedure Zl_�����ؼ�_�嵥��ӡ
(
  ID_IN          �����ؼ���Ϣ.ID%Type
)Is 
Begin
        
  --�����ؼ�״̬��ֻ�е�ǰ״̬Ϊ0���ؼ���Ϣ�Ž��и��£�
  --update �����ؼ���Ϣ set ��ǰ״̬=1,�ؼ�ҽʦ=�ؼ�ҽʦ_IN where id=ID_IN and ��ǰ״̬=0;
  
    --�����嵥״̬
  update �����ؼ���Ϣ  set �嵥״̬=1 where ID=ID_IN;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����ؼ�_�嵥��ӡ;
/


--�ؼ���Ŀ¼��
CREATE OR REPLACE Procedure Zl_�����ؼ�_��Ŀ¼��
(
  ID_IN          �����ؼ���Ϣ.ID%Type,
  ��Ŀ���_IN    �����ؼ���Ϣ.��Ŀ���%Type 
)Is
Begin
        
  --�����ؼ����Ŀ����������ѽ��ܵ���Ŀ���ܽ���¼�룩
  update �����ؼ���Ϣ set ��Ŀ���=��Ŀ���_IN where ID=ID_IN; --and ��ǰ״̬=1;
  
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����ؼ�_��Ŀ¼��;
/


--�ؼ�ȷ��
CREATE OR REPLACE Procedure Zl_�����ؼ�_ȷ��
(
  �����_IN      ��������Ϣ.�����%Type,
  �ؼ�����_IN    �����ؼ���Ϣ.�ؼ�����%Type,
  ���ʱ��_IN    �����ؼ���Ϣ.���ʱ��%Type  
)Is
  v_count number;
Begin  
  v_count := 0;
  begin
    select count(id) into v_count from �����ؼ���Ϣ where �����=�����_IN and �ؼ�����=�ؼ�����_IN and ��ǰ״̬<>2;
  exception
    when others then v_count := 0;     
  end;
  
  if v_count <= 0 then
    --���¼�����
    update ��������Ϣ set ��ǰ����=3 where �����=�����_IN;
  end if;
  
  --��������״̬��0-�����룬1-�ѽ��ܣ�2-����ɣ�
  update ����������Ϣ set ����״̬ = 2,���ʱ��=���ʱ��_IN 
  where ����ID=(select distinct ����ID from �����ؼ���Ϣ where �����=�����_IN and �ؼ�����=�ؼ�����_IN and ��ǰ״̬=1) 
        and ����״̬=1;  
        
  --�����ؼ�״̬�������ѽ��ܵ���Ŀ���ܽ���ȷ�ϣ�
  update �����ؼ���Ϣ set ��ǰ״̬=2, ���ʱ��=���ʱ��_IN where �����=�����_IN and �ؼ�����=�ؼ�����_IN and ��ǰ״̬=1;        
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����ؼ�_ȷ��;
/



--�������===================================================================================================

--������
CREATE OR REPLACE Procedure Zl_������_���
(
  ҽ��ID_IN    ��������Ϣ.ҽ��ID%Type     
)Is
Begin
  update ��������Ϣ set ��ǰ����=10 where ҽ��id=ҽ��ID_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);  
End Zl_������_���;
/


--ȡ�����
CREATE OR REPLACE Procedure Zl_������_ȡ�����
(
  ҽ��ID_IN    ��������Ϣ.ҽ��ID%Type     
)Is
Begin
  update ��������Ϣ set ��ǰ����=3 where ҽ��id=ҽ��ID_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);  
End Zl_������_ȡ�����;
/




