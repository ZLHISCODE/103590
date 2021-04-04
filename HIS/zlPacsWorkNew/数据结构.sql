
--1.病理检查信息
Create Table 病理检查信息(
    病理号 VARCHAR2(20),   
    医嘱ID Number(18),     
    检查类型 Number(1),
    当前过程 Number(2) default 0,
    巨检描述 Varchar2(2048),
    剩余位置 Varchar2(64),
    标本质量 Varchar2(10),
    制片质量 Varchar2(10))
    TABLESPACE zl9BaseItem; 
    
    
Alter Table 病理检查信息 Add Constraint 病理检查信息_PK Primary Key (病理号) Using Index Tablespace zl9indexhis;    
Alter Table 病理检查信息 Add Constraint 病理检查信息_FK_医嘱ID Foreign Key (医嘱ID) References 病人医嘱记录(ID) On Delete Cascade;  
Create Index 病理检查信息_IX_医嘱ID On 病理检查信息(医嘱ID) Pctfree 5 Tablespace zl9indexhis;   
Create Sequence 病理检查信息_病理号 Start With 1;



--2.病理标本信息
Create Table 病理标本信息(
    标本ID NUMBER(18),
    医嘱ID Number(18),
    标本名称 VARCHAR2(64) Not Null,
    材料类别 NUMBER(1) default 0,
    标本类型 NUMBER(1) default 0,
    采集部位 VARCHAR2(20),
    原有编号 VARCHAR2(20),
    数量 Number(2) Not Null,
    存放位置 VARCHAR2(64),
    接收日期 Date,
    备注 VARCHAR2(1024))
    TABLESPACE zl9BaseItem;    
    
Alter Table 病理标本信息 Add Constraint 病理标本信息_PK Primary Key (标本ID) Using Index Tablespace zl9indexhis;    
Alter Table 病理标本信息 Add Constraint 病理标本信息_FK_医嘱ID Foreign Key (医嘱ID) References 病人医嘱记录(ID) On Delete Cascade;
Create Index 病理标本信息_IX_医嘱ID On 病理标本信息(医嘱ID) Pctfree 5 Tablespace zl9indexhis;       
Create Sequence 病理标本信息_标本ID Start With 1;   


--3.病理送检信息
Create Table 病理送检信息(
    ID NUMBER(18),   
    医嘱ID NUMBER(18),
    送检单位 VARCHAR2(64),
    送检科室 VARCHAR2(64),
    送检人 VARCHAR2(64) Not Null,
    送检日期 DATE Not Null,
    联系方式 VARCHAR2(64),
    登记人 VARCHAR2(64) Not Null,
    核收状态 NUMBER(1) default 1,
    拒收原因 VARCHAR2(1024),
    通知人 VARCHAR2(64),
    备注 VARCHAR2(1024))
    TABLESPACE zl9BaseItem;
    

Alter Table 病理送检信息 Add Constraint 病理送检信息_PK Primary Key (ID) Using Index Tablespace zl9indexhis;    
Alter Table 病理送检信息 Add Constraint 病理送检信息_FK_医嘱ID Foreign Key (医嘱ID) References 病人医嘱记录(ID) On Delete Cascade;   
Create Index 病理送检信息_IX_医嘱ID On 病理送检信息(医嘱ID) Pctfree 5 Tablespace zl9indexhis; 
Create Sequence 病理送检信息_ID Start With 1;  
 
    
--4.病理申请信息    
Create Table 病理申请信息(
    申请ID Number(18),  
    病理号 Varchar2(20),  
    申请人 Varchar2(64) Not Null,
    申请时间 Date,        
    申请类型 Number(1) default 0,
    申请状态 Number(1) default 0,
    申请描述 Varchar2(1024),
    是否打印 Number(1) default 0,
    完成时间 Date)
    TABLESPACE zl9BaseItem;    
    
    
Alter Table 病理申请信息 Add Constraint 病理申请信息_PK Primary Key (申请ID) Using Index Tablespace zl9indexhis;  
Alter Table 病理申请信息 Add Constraint 病理申请信息_FK_病理号 Foreign Key (病理号) References 病理检查信息(病理号) On Delete Cascade;
Create Index 病理申请信息_IX_病理号 On 病理申请信息(病理号) Pctfree 5 Tablespace zl9indexhis;
Create Sequence 病理申请信息_申请ID Start With 1;  
  
    
--5.病理取材信息    
Create Table 病理取材信息(
    材块ID Number(18),
    序号 Number(18),
    病理号 Varchar2(20),
    申请ID Number(18),
    标本ID Number(18),
    标本名称 Varchar2(64),
    取材位置 Varchar2(64),
    形状 Varchar2(64),
    颜色 Varchar2(20),
    性质 Varchar2(20),
    标本量 Varchar2(20),
    蜡块数 Number(2) default 1,   
    是否冰余 Number(1) default 0,
    主取医师 Varchar2(64) Not Null,
    副取医师 Varchar2(64),
    记录医师 Varchar2(64) Not Null,
    取材时间 Date)
    TABLESPACE zl9BaseItem;   
    
    
Alter Table 病理取材信息 Add Constraint 病理取材信息_PK Primary Key (材块ID) Using Index Tablespace zl9indexhis;    
Alter Table 病理取材信息 Add Constraint 病理取材信息_FK_病理号 Foreign Key (病理号) References 病理检查信息(病理号) On Delete Cascade; 
--Alter Table 病理取材信息 Add Constraint 病理取材信息_FK_申请ID Foreign Key (申请ID) References 病理申请信息(申请ID) On Delete Cascade; --常规取材没有申请信息
Alter Table 病理取材信息 Add Constraint 病理取材信息_FK_标本ID Foreign Key (标本ID) References 病理标本信息(标本ID) On Delete Cascade; 
Create Index 病理取材信息_IX_病理号 On 病理取材信息(病理号) Pctfree 5 Tablespace zl9indexhis; 
Create Index 病理取材信息_IX_申请ID On 病理取材信息(申请ID) Pctfree 5 Tablespace zl9indexhis; 
Alter Table 病理取材信息 Add Constraint 病理取材信息_CK_是否冰余 Check (是否冰余 IN(0,1));
Create Sequence 病理取材信息_材块ID Start With 1; 
  
      
    
--6.病理脱钙信息    
Create Table 病理脱钙信息(
    ID Number(18),   
    标本ID Number(18),
    开始时间 Date,
    所需时长 Number(5),
    当前缸次 Number(2),
    完成状态 Number(1) default 0,
    操作员 Varchar2(64))
    TABLESPACE zl9BaseItem;     
    
    
Alter Table 病理脱钙信息 Add Constraint 病理脱钙信息_PK Primary Key (ID) Using Index Tablespace zl9indexhis;    
Alter Table 病理脱钙信息 Add Constraint 病理脱钙信息_FK_标本ID Foreign Key (标本ID) References 病理标本信息(标本ID) On Delete Cascade;   
Create Index 病理脱钙信息_IX_标本ID On 病理脱钙信息(标本ID) Pctfree 5 Tablespace zl9indexhis;    
Create Sequence 病理脱钙信息_ID Start With 1; 


--7.病理制片信息
Create Table 病理制片信息(
    ID Number(18),  
    病理号 Varchar(20), 
    材块ID Number(18),
    申请ID Number(18),
    制片类型 Number(1) default 0,
    制片方式 Number(1) default 0,
    制片时间 Date,
    制片数 Number(2),
    制片人 Varchar2(64),       
    当前状态 Number(1) default 0,
    清单状态 Number(1) default 0)
    TABLESPACE zl9BaseItem;     
    
    
Alter Table 病理制片信息 Add Constraint 病理制片信息_PK Primary Key (ID) Using Index Tablespace zl9indexhis;   
Alter table 病理制片信息 Add constraint 病理制片信息_FK_病理号 Foreign key(病理号) References 病理检查信息(病理号) On Delete Cascade;   
Alter Table 病理制片信息 Add Constraint 病理制片信息_FK_材块ID Foreign Key (材块ID) References 病理取材信息(材块ID) On Delete Cascade;  
Create Index 病理制片信息_IX_材块ID On 病理制片信息(材块ID) Pctfree 5 Tablespace zl9indexhis;     
Create Index 病理制片信息_IX_申请ID On 病理制片信息(申请ID, 材块ID) Pctfree 5 Tablespace zl9indexhis;    
Create Sequence 病理制片信息_ID Start With 1;  


--8.病理过程报告
Create Table 病理过程报告(
    ID Number(18),  
    病理号 Varchar2(20),
    标本名称 Varchar2(64),
    报告类型 Number(1),
    检查结果 Varchar2(2048),
    检查意见 Varchar2(2048),
    报告图像 Varchar2(2048),
    报告医师 Varchar2(64),        
    报告日期 Date,       
    当前状态 Number(1) default 0,
    备注 Varchar2(1024))
    TABLESPACE zl9BaseItem;  

Alter Table 病理过程报告 Add Constraint 病理过程报告_PK Primary Key (ID) Using Index Tablespace zl9indexhis;
Alter Table 病理过程报告 Add Constraint 病理过程报告_FK_病理号 Foreign Key (病理号) References 病理检查信息(病理号) on Delete Cascade;
Create Index 病理过程报告_IX_病理号 On 病理过程报告(病理号) Pctfree 5 Tablespace zl9indexhis;
Create Sequence 病理过程报告_ID Start With 1;


    
--9.病理抗体信息  
Create Table 病理抗体信息(
    抗体ID Number(18), 
    抗体名称 VARCHAR2(64) Not Null,
    使用人份 Number(5),
    已用人份 Number(5),
    生产日期 Date,
    有效期 Number(2),
    过期日期 Date,
    克隆性 Number(1),
    作用对象 Varchar2(20),
    理化性质 Varchar2(10),
    应用情况 Varchar2(1024),
    登记人 Varchar2(64)  Not Null,
    登记时间 Date,
    使用状态 Number(1) default 1,
    备注 Varchar2(1024))
    TABLESPACE zl9BaseItem;  
        
    
Alter Table 病理抗体信息 Add Constraint 病理抗体信息_PK Primary Key (抗体ID) Using Index Tablespace zl9indexhis;   
Create Sequence 病理抗体信息_抗体ID Start With 1;     


--10.病理特检信息    
Create Table 病理特检信息(
    ID Number(18),    
    病理号 Varchar(20) not null,
    材块ID Number(18) not null,
    申请ID Number(18),        
    抗体ID Number(18),
    特检类型 Number(1) default 0,
    制作类型 Number(1) default 0,
    当前状态 NUMBER(1) default 0,
    完成时间 Date,    
    特检医师 Varchar2(64),
    清单状态 Number(1) default 0,
    项目结果 Varchar2(20) null)
    TABLESPACE zl9BaseItem; 
    
    
Alter Table 病理特检信息 Add Constraint 病理特检信息_PK Primary Key (ID) Using Index Tablespace zl9indexhis;   
Alter table 病理特检信息 Add constraint 病理特检信息_FK_病理号 Foreign key(病理号) References 病理检查信息(病理号) On Delete Cascade;        
Alter Table 病理特检信息 Add Constraint 病理特检信息_FK_材块ID Foreign Key (材块ID) References 病理取材信息(材块ID) On Delete Cascade; 
Alter Table 病理特检信息 Add Constraint 病理特检信息_FK_申请ID Foreign Key (申请ID) References 病理申请信息(申请ID) On Delete Cascade; 
Alter Table 病理特检信息 Add Constraint 病理特检信息_FK_抗体ID Foreign Key (抗体ID) References 病理抗体信息(抗体ID) On Delete Cascade; 
Create Index 病理特检信息_IX_材块ID On 病理特检信息(材块ID, 抗体ID) Pctfree 5 Tablespace zl9indexhis;          
Create Index 病理特检信息_IX_申请ID On 病理特检信息(申请ID,材块ID,抗体ID) Pctfree 5 Tablespace zl9indexhis; 
Create Sequence 病理特检信息_ID Start With 1;   
         

--11.病理报告延迟
Create Table 病理报告延迟(
    ID Number(18),    
    病理号 Varchar2(20),
    延迟原因 Varchar2(1024) not null,        
    延迟天数 Number(2) not null,
    临时诊断 Varchar2(1024),
    转达人 Varchar2(64),
    登记人 Varchar2(64),
    登记时间 Date,    
    当前状态 Number(1) default 0)
    TABLESPACE zl9BaseItem; 
    
Alter Table 病理报告延迟 Add Constraint 病理报告延迟_PK Primary Key(ID) Using Index Tablespace zl9indexhis;
Alter Table 病理报告延迟 Add Constraint 病理报告延迟_FK_病理号 Foreign Key(病理号) References 病理检查信息(病理号) On Delete Cascade;
Create Index 病理报告延迟_IX_病理号 On 病理报告延迟(病理号) Pctfree 5 Tablespace zl9indexhis;
Create Sequence 病理报告延迟_ID Start With 1;


--12.病理会诊信息
Create Table 病理会诊信息(
    ID Number(18),    
    病理号 Varchar2(20),
    申请医师 Varchar2(64) not null,
    会诊医师 Varchar2(64),
    会诊单位 Varchar2(64),         
    会诊时间 Date not null,
    截止时间 Date not null,
    会诊类型 Number(1) default 0,
    检查描述 Varchar2(2048),
    诊断结果 Varchar2(2048),
    诊断意见 Varchar2(2048),    
    完成时间 Date,
    备注 Varchar2(1024),
    当前状态 Number(1) default 0)
    TABLESPACE zl9BaseItem; 
        
    
Alter Table 病理会诊信息 Add Constraint 病理会诊信息_PK Primary Key(ID) Using Index Tablespace zl9indexhis;
Alter Table 病理会诊信息 Add Constraint 病理会诊信息_FK_病理号 Foreign Key(病理号) References 病理检查信息(病理号) On Delete Cascade;
Create Index 病理会诊信息_IX_病理号 On 病理会诊信息(病理号) Pctfree 5 Tablespace zl9indexhis;   
Create Sequence 病理会诊信息_ID Start With 1;

      
    
--13.病理抗体反馈
Create Table 病理抗体反馈(
    ID Number(18),   
    抗体ID Number(18), 
    参考病理号 VARCHAR2(200),
    实验类型 Number(1) default 0,
    抗体评价 VARCHAR2(10),
    反馈意见 VARCHAR2(1024) Not Null,
    反馈医生 VARCHAR2(64) Not Null,
    反馈时间 Date)
    TABLESPACE zl9BaseItem;   
    
Alter Table 病理抗体反馈 Add Constraint 病理抗体反馈_PK Primary Key (ID) Using Index Tablespace zl9indexhis;    
Alter Table 病理抗体反馈 Add Constraint 病理抗体反馈_FK_抗体ID Foreign Key (抗体ID) References 病理抗体信息(抗体ID) On Delete Cascade; 
Create Index 病理抗体反馈_IX_抗体ID On 病理抗体反馈(抗体ID) Pctfree 5 Tablespace zl9indexhis;   
Create Sequence 病理抗体反馈_ID Start With 1;  


--14.病理套餐信息
Create Table 病理套餐信息(
    套餐ID Number(18), 
    套餐名称 VARCHAR2(64) not null,
    套餐说明 VARCHAR2(1024),
    创建人 VARCHAR2(64) Not Null,
    创建时间 Date)
    TABLESPACE zl9BaseItem;  
    
Alter Table 病理套餐信息 Add Constraint 病理套餐信息_PK Primary Key (套餐ID) Using Index Tablespace zl9indexhis;
Create Sequence 病理套餐信息_套餐ID Start With 1;


--15.病理套餐关联
 Create Table 病理套餐关联(
    ID Number(18),    
    套餐ID Number(18), 
    抗体ID Number(18))
    TABLESPACE zl9BaseItem;  

Alter Table 病理套餐关联 Add Constraint 病理套餐关联_PK Primary Key (ID) Using Index Tablespace zl9indexhis;
Alter Table 病理套餐关联 Add Constraint 病理套餐关联_FK_套餐ID Foreign Key (套餐ID) References 病理套餐信息(套餐ID) On Delete Cascade;
Alter Table 病理套餐关联 Add Constraint 病理套餐关联_FK_抗体ID Foreign Key (抗体ID) References 病理抗体信息(抗体ID) On Delete Cascade;
Create Index 病理套餐关联_IX_套餐ID On 病理套餐关联(套餐ID,抗体ID) Pctfree 5 Tablespace zl9indexhis;
Create Sequence 病理套餐关联_ID Start With 1;
  
    
--16.病理归档信息
Create Table 病理归档信息(
    档案ID Number(18), 
    病理号 Varchar2(20),
    资料类别 Number(1) default 0,
    资料编号 Varchar2(20) Not Null,
    资料数量 Number(2) Not null,
    可借数量 Number(2) Not null,
    遗失数量 Number(2),    
    入库时间 Date,
    档案状态 Number(1)  default 0,
    入库人 Varchar2(64) Not Null,
    存放位置 Varchar2(64),
    备注 Varchar2(1024))
    TABLESPACE zl9BaseItem; 
    
    
Alter Table 病理归档信息 Add Constraint 病理归档信息_PK Primary Key (档案ID) Using Index Tablespace zl9indexhis; 
Alter Table 病理归档信息 Add Constraint 病理归档信息_FK_病理号 Foreign Key (病理号) References 病理检查信息(病理号) On Delete Cascade;
Create Index 病理归档信息_IX_病理号 On 病理归档信息(病理号) Pctfree 5 Tablespace zl9indexhis; 
Create Index 病理归档信息_IX_资料编号 On 病理归档信息(资料编号) Pctfree 5 Tablespace zl9indexhis; 
Create Sequence 病理归档信息_档案ID Start With 1;            
    
   
--17.病理借阅信息    
Create Table 病理借阅信息(
    ID Number(18), 
    档案ID Number(18),
    借阅时间 Date Not Null,
    借阅人 Varchar2(64) Not Null,
    证件类型 Number(1) default 0,
    证件号码 Varchar2(20),
    联系电话 Varchar2(20),
    联系地址 Varchar2(128),
    押金 Number(16, 5),
    借阅数量 Number(2),
    借阅类型 Number(1) Default 0,
    借阅天数 Number(5),
    借阅原因 Varchar2(1024),
    登记人 Varchar2(64) Not Null,
    归还状态 Number(1)  default 1,
    归还日期 Date,
    退还押金 Number(16,5),
    外诊医院 Varchar2(64),
    外诊医师 Varchar2(64),
    外诊意见 Varchar2(2048))
    TABLESPACE zl9BaseItem;       

Alter Table 病理借阅信息 Add Constraint 病理借阅信息_PK Primary Key (ID) Using Index Tablespace zl9indexhis;    
Alter Table 病理借阅信息 Add Constraint 病理借阅信息_FK_档案ID Foreign Key (档案ID) References 病理归档信息(档案ID) On Delete Cascade; 
Create Index 病理借阅信息_IX_档案ID On 病理借阅信息(档案ID) Pctfree 5 Tablespace zl9indexhis; 
Create Sequence 病理借阅信息_ID Start With 1;
  
