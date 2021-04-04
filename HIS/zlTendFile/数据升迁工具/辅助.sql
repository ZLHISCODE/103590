/*****************************************
һ������°����л������ݣ�׼��������Ǩ
*/****************************************
--ֹͣ����Ա���ձ�����
alter table ���˻����ļ�  disable constraint ���˻����ļ�_FK_��ҳID;
alter table ���˻����ӡ  disable constraint ���˻����ӡ_FK_�ļ�ID;
alter table ���˻�����ϸ  disable constraint ���˻�����ϸ_FK_��¼ID ;
alter table ���˻�������  disable constraint ���˻�������_FK_�ļ�ID;
alter table ���˻�����Ŀ  disable constraint ���˻�����Ŀ_FK_�ļ�ID;

truncate table ���˻����ӡ;
truncate table ���˻�����ϸ;
truncate table ���˻�������;
truncate table ���˻����ļ�;
truncate table ������Ǩ��¼;

--�ָ����
alter table ���˻����ļ�  enable constraint ���˻����ļ�_FK_��ҳID;
alter table ���˻����ӡ  enable constraint ���˻����ӡ_FK_�ļ�ID;
alter table ���˻�������  enable constraint ���˻�������_FK_�ļ�ID;
alter table ���˻�����ϸ  enable constraint ���˻�����ϸ_FK_��¼ID;
alter table ���˻�����Ŀ  enable constraint ���˻�����Ŀ_FK_�ļ�ID;


/*****************************************
������ѯ��ռ�
*/****************************************
--�鿴��ռ��ռ�����
Select Tablespace_Name, ռ��, Sum(ʵ��) As ʵ��, Round(ռ�� / Sum(ʵ��) * 100, 2) As ����, Wm_Concat(File_Name) As Fn
From (Select a.Tablespace_Name, Round(Sum(a.Bytes) / 1024 / 1024, 2) As ռ��, Round(b.Bytes / 1024 / 1024, 2) As ʵ��,
              b.File_Name
       From Dba_Segments A, Dba_Data_Files B
       Where a.Tablespace_Name = b.Tablespace_Name
       Group By a.Tablespace_Name, b.Bytes, b.File_Name)
Group By Tablespace_Name, ռ��
Order By 4 Desc;

--ֱ�����������ļ�
Alter Tablespace ZL9EXPENSE Add Datafile 'F:\ORACLE\PRODUCT\10.2.0\ORADATA\ORCL\USERS01.DBF' Size 50M;

--�ڱ�ռ���������ļ���С
Alter Database Datafile 'D:\ORACLE\PRODUCT\10.2.0\ORADATA\ORCL\ZL9EPRDAT.DBF' Resize 50G;

--ֱ�ӵ�����ʱ��ռ�Ĵ�С
alter Database Tempfile 'F:\ORACLE\PRODUCT\10.2.0\ORADATA\ORCL\zltooltmp.dbf' Resize 300M;

--��ѯ���������ļ�
Select * From v$datafile;
