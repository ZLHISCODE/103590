delete from 强制结果;
alter table 强制结果 add 默认值 text(100);
alter table 强制结果 add 默认选择 bit;
alter table 强制结果 add 元素类型 text(5);
alter table 强制结果 add 强制结果值 text(100);

--Scheduled Procedure Step 预定过程步骤

insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('40','1','预定工作站AE','Scheduled Station AE Title','[CallingAT]',True,'AE',True,'[CallingAET]',True,'1');    
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('40','2','预定过程步骤开始日期','Scheduled Procedure Step Start Date ','[首次日期]',True,'DA',True,'[首次日期]',True,'1');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('40','3','预定过程步骤开始时间','Scheduled Procedure Step Start Time','[首次时间]',True,'TM',True,'[首次时间]',True,'1');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('8','60','影像类别','Modality','[影像类别]',True,'CS',True,'[影像类别]',True,'1');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('40','6','预定的医生姓名','Scheduled Performing Physician’s Name','',True,'PN',True,'',True,'2');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('40','7','预定的过程步骤描述','Scheduled Procedure Step Description','',True,'LO',True,'',True,'1C');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('40','10','预定工作站名称','Scheduled Station Name','[执行间]',True,'SH',True,'[执行间]',True,'2');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('40','11','预定过程步骤位置','Scheduled Procedure Step Location','[执行间]',True,'SH',True,'[执行间]',True,'2');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('40','8','预定协议代码序列','Scheduled Protocol Code Sequence','',True,'SQ',True,'',True,'1C');   
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('40','12','药物预处理','Pre-Medication','',True,'LO',True,'',True,'2C');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('40','9','预定过程步骤ID','Scheduled Procedure Step ID','[执行过程]',True,'SH',True,'[执行过程]',True,'1');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('32','1070','被请求的造影剂','Requested Contrast Agent','',True,'LO',True,'',True,'2C');
	    
--Requested Procedure 请求的过程
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('40','1001','请求的过程ID','Requested Procedure ID','[医嘱ID]_[发送号]',False,'SH',True,'[医嘱ID]_[发送号]',True,'1');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('32','1060','请求的过程描述','Requested Procedure Description','',False,'LO',True,'',True,'1C');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('32','1064','请求的过程代码序列','Requested Procedure Code Sequence','',False,'SQ',True,'',True,'1C');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('20','D','检查UID','Study Instance UID','[医嘱ID]',False,'UI',True,'[医嘱ID]',True,'1');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('8','1110','参考检查序列','Referenced Study Sequence','',False,'SQ',True,'',True,'2');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('40','1003','请求过程的优先级','Requested Procedure Priority','',False,'SH',True,'',True,'2');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('40','1004','病人转移安排','Patient Transport Arrangements','',False,'LO',True,'',True,'2');
	   	  	    	    
--Image Service Request 图像服务请求
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('8','50','编号','Accession Number','[医嘱ID]',False,'SH',True,'[医嘱ID]',True,'2');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('32','1032','请求的医生姓名','Requesting Physician','',False,'PN',True,'',True,'2');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('8','90','参考医生姓名','Referring Physician’s Name','',False,'PN',True,'',True,'2');

--Visit Identification
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('38','10','许可ID','Admission ID','',False,'LO',True,'',True,'2');
	    
--Visit Status Module
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('38','300','当前病人位置','Current Patient Location','',False,'LO',True,'',True,'2');
	    	    	    	    	    	    	    	    	    	 
--Visit Relationship Module
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('8','1120','参考病人序列','Referenced Patient Sequence','',False,'SQ',True,'',True,'2');
	    
--Patient Identification  病人标识
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('10','10','病人姓名','Patient’s Name','[英文名]',False,'PN',True,'[英文名]',True,'1');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('10','20','病人ID','Patient ID','[标识号]',False,'LO',True,'[标识号]',True,'1');   
	    
--Patient Demographic  病人统计学信息	  
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('10','30','病人生日','Patient’s Birth Date ','[出生日期]',False,'DA',True,'[出生日期]',True,'2');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('10','40','病人性别','Patient’s Sex','[性别]',False,'CS',True,'[性别]',True,'2');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('10','1010','病人年龄','Patient’s Age','[年龄]',False,'AS',True,'[年龄]',True,'3');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('10','1020','病人体形','Patient Size','',False,'DS',True,'',True,'3');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('10','1030','病人体重','Patient Weight','',False,'DS',True,'',True,'2');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('40','3001','病人数据保密要求','Confidentiality constraint on patient data','',False,'LO',True,'',True,'2');
	    
--Patient Medical 病人医疗模型
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('38','500','病人状态','Patient State','',False,'LO',True,'',True,'2');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('10','21C0','怀孕状态','Pregnancy Status','',False,'US',True,'',True,'2');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('10','2000','用药警告','Medical Alerts','',False,'LO',True,'',True,'2');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('10','2110','过敏','Contrast Allergies','',False,'LO',True,'',True,'2');
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('38','50','特殊需要','Special Needs','',False,'LO',True,'',True,'2');
	    
--General Series Module 通用序列模型
insert into 强制结果(组号,元素号,中文标题,英文标题,数据值,是否嵌套数据,值类型,被选择,默认值,默认选择,元素类型) 
	    values('18','15','检查部位','Body Part Examined','',False,'CS',False,'',False,'3');
	    
update 版本表 set 版本号='10.13.01';