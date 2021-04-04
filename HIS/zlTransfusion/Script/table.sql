Create Table 执行打印记录 (
       医嘱ID     Number(18),
       发送号         Number(18),
       流水号     Number(18),
       打印说明       Varchar2(1000),
       打印时间       Date,
       打印人         Varchar2(20))
       TABLESPACE zl9CisRec
       Pctfree 5 Pctused 85;

Create Table 暂存药品记录 (
       NO             VARCHAR2(8),
       序号           NUMBER(5),
       病人ID         Number(18),
       科室ID         Number(18),	
       医嘱ID         Number(18),	
       发送号         Number(18),
       药品ID         Number(18),	
       药品名称       Varchar2(80),	
       规格           Varchar2(40),
       执行分类       Number(2),    -- 0-其他治疗用 1-输液用 2-注射用 3-皮试用
       使用状态       Number(1),    -- 0-未用,1-已用
       摘要           Varchar2(200),	
       入出系数       Number(2),    -- 1-收暂存药品 -1-使用暂存药品
       单位           varchar2(20), -- 目录内的药品或医嘱药品为计算单位 ,目录外药品为门诊单位
       容量           Number(16,5),
       数量           Number(16,5), -- 不允许负数,目录内记录的是计算单位数量,目录外为门诊单位数量
       单价           Number(16,5),	
       金额           Number(16,5),	
       操作员         Varchar2(10),	
       登记时间       Date,	
       作废时间       Date) --	使用状态为1的记录不能作废
       TABLESPACE zl9CisRec
       Pctfree 5 Pctused 85;

Create Table 座位状况记录(
       病人ID         Number(18),
       科室ID         Number(18),
       编号           Varchar2(30), -- 座位编号
       类别           Number(1), -- 0-普通座位 1-加座 2-特殊药品座位 3-VIP座位  
       收费细目ID     Number(18), -- 如要收费，则存放对应的收费细目ID
       状态           Number(1), -- 0-空,1-在用,2-不可用,比如在维修
       备注           Varchar2(100),
       NO             Varchar2(8))
       TABLESPACE zl9CisRec
       Pctfree 5 Pctused 85;
       

Create Table 排队记录(
       病人ID         Number(18),	
       科室ID         Number(18),	
       日期           Date Default Sysdate,	
       顺序号         Number(5), -- 病人排队的顺序号
       加权号         Number(10), -- 特殊病人优先下改变顺序用
       状态           Number(2), -- 0-正常 1-完成 2-弃号 3-退号
       备注           Varchar2(100))
       TABLESPACE zl9CisRec
       Pctfree 5 Pctused 85;         
