delete from ǿ�ƽ��;
alter table ǿ�ƽ�� add Ĭ��ֵ text(100);
alter table ǿ�ƽ�� add Ĭ��ѡ�� bit;
alter table ǿ�ƽ�� add Ԫ������ text(5);
alter table ǿ�ƽ�� add ǿ�ƽ��ֵ text(100);

--Scheduled Procedure Step Ԥ�����̲���

insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('40','1','Ԥ������վAE','Scheduled Station AE Title','[CallingAT]',True,'AE',True,'[CallingAET]',True,'1');    
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('40','2','Ԥ�����̲��迪ʼ����','Scheduled Procedure Step Start Date ','[�״�����]',True,'DA',True,'[�״�����]',True,'1');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('40','3','Ԥ�����̲��迪ʼʱ��','Scheduled Procedure Step Start Time','[�״�ʱ��]',True,'TM',True,'[�״�ʱ��]',True,'1');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('8','60','Ӱ�����','Modality','[Ӱ�����]',True,'CS',True,'[Ӱ�����]',True,'1');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('40','6','Ԥ����ҽ������','Scheduled Performing Physician��s Name','',True,'PN',True,'',True,'2');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('40','7','Ԥ���Ĺ��̲�������','Scheduled Procedure Step Description','',True,'LO',True,'',True,'1C');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('40','10','Ԥ������վ����','Scheduled Station Name','[ִ�м�]',True,'SH',True,'[ִ�м�]',True,'2');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('40','11','Ԥ�����̲���λ��','Scheduled Procedure Step Location','[ִ�м�]',True,'SH',True,'[ִ�м�]',True,'2');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('40','8','Ԥ��Э���������','Scheduled Protocol Code Sequence','',True,'SQ',True,'',True,'1C');   
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('40','12','ҩ��Ԥ����','Pre-Medication','',True,'LO',True,'',True,'2C');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('40','9','Ԥ�����̲���ID','Scheduled Procedure Step ID','[ִ�й���]',True,'SH',True,'[ִ�й���]',True,'1');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('32','1070','���������Ӱ��','Requested Contrast Agent','',True,'LO',True,'',True,'2C');
	    
--Requested Procedure ����Ĺ���
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('40','1001','����Ĺ���ID','Requested Procedure ID','[ҽ��ID]_[���ͺ�]',False,'SH',True,'[ҽ��ID]_[���ͺ�]',True,'1');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('32','1060','����Ĺ�������','Requested Procedure Description','',False,'LO',True,'',True,'1C');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('32','1064','����Ĺ��̴�������','Requested Procedure Code Sequence','',False,'SQ',True,'',True,'1C');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('20','D','���UID','Study Instance UID','[ҽ��ID]',False,'UI',True,'[ҽ��ID]',True,'1');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('8','1110','�ο��������','Referenced Study Sequence','',False,'SQ',True,'',True,'2');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('40','1003','������̵����ȼ�','Requested Procedure Priority','',False,'SH',True,'',True,'2');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('40','1004','����ת�ư���','Patient Transport Arrangements','',False,'LO',True,'',True,'2');
	   	  	    	    
--Image Service Request ͼ���������
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('8','50','���','Accession Number','[ҽ��ID]',False,'SH',True,'[ҽ��ID]',True,'2');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('32','1032','�����ҽ������','Requesting Physician','',False,'PN',True,'',True,'2');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('8','90','�ο�ҽ������','Referring Physician��s Name','',False,'PN',True,'',True,'2');

--Visit Identification
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('38','10','���ID','Admission ID','',False,'LO',True,'',True,'2');
	    
--Visit Status Module
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('38','300','��ǰ����λ��','Current Patient Location','',False,'LO',True,'',True,'2');
	    	    	    	    	    	    	    	    	    	 
--Visit Relationship Module
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('8','1120','�ο���������','Referenced Patient Sequence','',False,'SQ',True,'',True,'2');
	    
--Patient Identification  ���˱�ʶ
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('10','10','��������','Patient��s Name','[Ӣ����]',False,'PN',True,'[Ӣ����]',True,'1');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('10','20','����ID','Patient ID','[��ʶ��]',False,'LO',True,'[��ʶ��]',True,'1');   
	    
--Patient Demographic  ����ͳ��ѧ��Ϣ	  
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('10','30','��������','Patient��s Birth Date ','[��������]',False,'DA',True,'[��������]',True,'2');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('10','40','�����Ա�','Patient��s Sex','[�Ա�]',False,'CS',True,'[�Ա�]',True,'2');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('10','1010','��������','Patient��s Age','[����]',False,'AS',True,'[����]',True,'3');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('10','1020','��������','Patient Size','',False,'DS',True,'',True,'3');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('10','1030','��������','Patient Weight','',False,'DS',True,'',True,'2');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('40','3001','�������ݱ���Ҫ��','Confidentiality constraint on patient data','',False,'LO',True,'',True,'2');
	    
--Patient Medical ����ҽ��ģ��
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('38','500','����״̬','Patient State','',False,'LO',True,'',True,'2');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('10','21C0','����״̬','Pregnancy Status','',False,'US',True,'',True,'2');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('10','2000','��ҩ����','Medical Alerts','',False,'LO',True,'',True,'2');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('10','2110','����','Contrast Allergies','',False,'LO',True,'',True,'2');
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('38','50','������Ҫ','Special Needs','',False,'LO',True,'',True,'2');
	    
--General Series Module ͨ������ģ��
insert into ǿ�ƽ��(���,Ԫ�غ�,���ı���,Ӣ�ı���,����ֵ,�Ƿ�Ƕ������,ֵ����,��ѡ��,Ĭ��ֵ,Ĭ��ѡ��,Ԫ������) 
	    values('18','15','��鲿λ','Body Part Examined','',False,'CS',False,'',False,'3');
	    
update �汾�� set �汾��='10.13.01';