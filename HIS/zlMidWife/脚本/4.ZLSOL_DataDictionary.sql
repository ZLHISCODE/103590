�����ֵ䣺
create table SOL_STD_FetalPosition--̥��λ
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table SOL_STD_Delivery--���䷽ʽ
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table SOL_STD_PerinealLaceration--�����������
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table SOL_STD_Anesthesia--����ʽ
(
code varchar2(10),
name varchar2(50),
Description varchar2(500)
);
create table SOL_STD_FetalPresentation--̥��¶
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);
create table SOL_STD_NeonatalAbnormality--�������쳣���
(
code varchar2(10),
name varchar2(50),
Description varchar2(100)
);

--̥��λ
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('01','����ǰ(LOA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('02','����ǰ(ROA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('03','�����(LOP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('04','�����(ROP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('05','�����(LOT)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('06','�����(ROT)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('07','���ǰ(LMA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('08','���ǰ(RMA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('09','����(LMP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('10','����(RMP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('11','����(LMT)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('12','����(RMT)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('13','����ǰ(LSA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('14','����ǰ(RSA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('15','������(LSP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('16','������(RSP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('17','������(LST)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('18','������(RST)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('19','���ǰ(LScA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('20','�Ҽ�ǰ(RscA)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('21','����(LScP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('22','�Ҽ��(RScP)','');
Insert Into SOL_STD_FetalPosition(code,name,Description) Values('99','����','');
--���䷽ʽ
Insert Into SOL_STD_Delivery(code,name,Description) Values('1','������Ȼ����','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('11','�����п�','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('12','����δ��','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('2','������������','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('21','��ǯ����','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('22','��λ����','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('23','̥ͷ����','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('3','�ʹ���','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('31','�ӹ��¶κ��п��ʹ���','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('32','�ӹ����ʹ���','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('33','��Ĥ���ʹ���','');
Insert Into SOL_STD_Delivery(code,name,Description) Values('9','����','');
--�����������
Insert Into SOL_STD_PerinealLaceration(code,name,Description) Values('1','������','');
Insert Into SOL_STD_PerinealLaceration(code,name,Description) Values('2','�������','');
Insert Into SOL_STD_PerinealLaceration(code,name,Description) Values('3','�������','');
Insert Into SOL_STD_PerinealLaceration(code,name,Description) Values('4','�������','');
Insert Into SOL_STD_PerinealLaceration(code,name,Description) Values('5','�����п�','');
--����ʽ
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('1','ȫ������','�������ʹȫ��������״̬');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('11','��������','������������ķ���ʹȫ��������״̬');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('12','��������','������ע�������ʹȫ��������״̬');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('13','��������','����ǰ��ʹ������־��ʧ�ķ���');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('2','׵��������','������ҩע��׵���ڴﵽ�ֲ�����Ч���ķ���');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('21','����Ĥ��ǻ��������','������ҩע������Ĥ��ǻ�ﵽ�ֲ�����Ч���ķ���');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('22','Ӳ��Ĥ��ǻ��������','������ҩע��Ӳ��Ĥ��ǻ�����ֲ�����Ч���ķ���');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('3','�ֲ�����','������ҩֱ��ע��ʩ����������֯�ڻ�������λ��Χ��������');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('31','�񾭴���������','���ֲ�����ҩע�����񾭴Ը�����ʹͨ���񾭴Ե��񾭼������ֲ�����������ֲ�����ķ���');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('32','�񾭽���������','���ֲ�����ҩע�����񾭽ڸ�����ʹͨ���񾭽ڵ��񾭼������ֲ�����������ֲ�����ķ���');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('33','����������','������ҩ��ע�����񾭸ɵ���Χ��ʹ���񾭷ֲ�����������������õķ���');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('34','������������','������ҩע��������Ұ���ܣ�ʹͨ������Ұ�Լ�������Ұ��������ĩ�ҽ��ܵ����͵ľֲ�������');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('35','�ֲ���������','������ҩ�������п��߷ֲ�ע����֯�ڣ���������֯�е���ĩ�ҵ�������');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('36','��������','������ҩֱ����ճĤ��Ƥ���Ӵ���ʹ֧��ò���ճĤ��Ƥ���ڵ���ĩ�ұ����͵�������');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('4','��������','��һ������ҩ�����ö�������������ǿ����Ч��');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('41','��������ȫ��','�����������������ͬ���ò�������Ч��');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('42','��ҩ��������','��������ҩ������ͬ���ò�������Ч��');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('43','�񾭴���ӲĤ�����͸�������','�񾭴����������Ӳ��Ĥ��ǻ��������ͬ���ò�������Ч��');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('44','ȫ�鸴��ȫ����','��ȫ�������ͬʱ�������ͻ���Ѫѹ');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('45','ȫ�鸴�Ͽ����Խ�ѹ','��ȫ�������ͬʱ���ͻ��ߵ�����');
Insert Into SOL_STD_Anesthesia(code,name,Description) Values('9','����������','����δ�ἰ������������');
--̥��¶
Insert Into SOL_STD_FetalPresentation(code,name,Description) Values('1','ͷ��¶','');
Insert Into SOL_STD_FetalPresentation(code,name,Description) Values('2','����¶','');
Insert Into SOL_STD_FetalPresentation(code,name,Description) Values('3','����¶','');
Insert Into SOL_STD_FetalPresentation(code,name,Description) Values('4','����¶','');
Insert Into SOL_STD_FetalPresentation(code,name,Description) Values('9','����','');
--�������쳣���
Insert Into SOL_STD_NeonatalAbnormality(code,name,Description) Values('1','��','');
Insert Into SOL_STD_NeonatalAbnormality(code,name,Description) Values('2','��������������','');
Insert Into SOL_STD_NeonatalAbnormality(code,name,Description) Values('3','����','');
Insert Into SOL_STD_NeonatalAbnormality(code,name,Description) Values('4','���','');
Insert Into SOL_STD_NeonatalAbnormality(code,name,Description) Values('5','��Ϣ','');
Insert Into SOL_STD_NeonatalAbnormality(code,name,Description) Values('6','�ͳ�������','');
Insert Into SOL_STD_NeonatalAbnormality(code,name,Description) Values('9','����','');
