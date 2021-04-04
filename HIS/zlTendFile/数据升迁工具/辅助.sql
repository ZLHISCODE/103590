/*****************************************
一、清除新版所有护理数据，准备重新升迁
*/****************************************
--停止外键以便清空表数据
alter table 病人护理文件  disable constraint 病人护理文件_FK_主页ID;
alter table 病人护理打印  disable constraint 病人护理打印_FK_文件ID;
alter table 病人护理明细  disable constraint 病人护理明细_FK_记录ID ;
alter table 病人护理数据  disable constraint 病人护理数据_FK_文件ID;
alter table 病人护理活动项目  disable constraint 病人护理活动项目_FK_文件ID;

truncate table 病人护理打印;
truncate table 病人护理明细;
truncate table 病人护理数据;
truncate table 病人护理文件;
truncate table 护理升迁记录;

--恢复外键
alter table 病人护理文件  enable constraint 病人护理文件_FK_主页ID;
alter table 病人护理打印  enable constraint 病人护理打印_FK_文件ID;
alter table 病人护理数据  enable constraint 病人护理数据_FK_文件ID;
alter table 病人护理明细  enable constraint 病人护理明细_FK_记录ID;
alter table 病人护理活动项目  enable constraint 病人护理活动项目_FK_文件ID;


/*****************************************
二、查询表空间
*/****************************************
--查看表空间的占用情况
Select Tablespace_Name, 占用, Sum(实际) As 实际, Round(占用 / Sum(实际) * 100, 2) As 比率, Wm_Concat(File_Name) As Fn
From (Select a.Tablespace_Name, Round(Sum(a.Bytes) / 1024 / 1024, 2) As 占用, Round(b.Bytes / 1024 / 1024, 2) As 实际,
              b.File_Name
       From Dba_Segments A, Dba_Data_Files B
       Where a.Tablespace_Name = b.Tablespace_Name
       Group By a.Tablespace_Name, b.Bytes, b.File_Name)
Group By Tablespace_Name, 占用
Order By 4 Desc;

--直接增加数据文件
Alter Tablespace ZL9EXPENSE Add Datafile 'F:\ORACLE\PRODUCT\10.2.0\ORADATA\ORCL\USERS01.DBF' Size 50M;

--在表空间调整数据文件大小
Alter Database Datafile 'D:\ORACLE\PRODUCT\10.2.0\ORADATA\ORCL\ZL9EPRDAT.DBF' Resize 50G;

--直接调整临时表空间的大小
alter Database Tempfile 'F:\ORACLE\PRODUCT\10.2.0\ORADATA\ORCL\zltooltmp.dbf' Resize 300M;

--查询所有数据文件
Select * From v$datafile;
