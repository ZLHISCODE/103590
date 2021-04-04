--分类目录
--1.公共基础,2.医保基础,3.病人病案基础,4.费用基础,5.药品卫材基础
--6.临床基础,7.临床路径基础,8.病历基础,9.护理基础,10.检验基础
--11.检查基础,12.医保业务,13.病人病案业务,14.费用业务,15.药品卫材业务
--16.临床医嘱,17.临床路径,18.病历业务,19.护理业务,20.检验业务,21.检查业务
----------------------------------------------------------------------------
--[[1.公共基础]]
----------------------------------------------------------------------------
Create Table 病人费用异常记录
(
 费用ID       Number(18) Not Null,
 No           Varchar2(20),
 记录性质     Number(2),
 病人ID       Number(18),
 产生环节     Number(2) Not Null, -- 0-记费同步标志；1-作废同步标志；2-转费同步标志
 同步标志     Number(2), -- 1-药品/卫材单据未同步；2-药品/卫材单据收费状态未同步
 记录时间     Date,
 操作员姓名   Varchar2(100),
 工作站       Varchar2(100)
) TABLESPACE zl9Expense initrans 20;

Create Table 部门位置(
编码 VARCHAR2(4),
上级 VARCHAR2(4),
名称 varchar2(50),
简码 varchar2(50),
末级 number(1) DEFAULT 0
)Tablespace ZL9BASEITEM;

CREATE TABLE 三方接口配置(
  接口名 varchar2(50),
  参数号 Number(3),
  参数名 varchar2(50),
  参数值 varchar2(2000),
  说明 varchar2(200)
  )TABLESPACE zl9BaseItem;
Create Table 三方服务配置目录(
    系统标识 varchar2(100),
    服务名称 varchar2(100),
    服务地址 varchar2(300),
    是否启用 NUMBER(1))
    TABLESPACE zl9BaseItem
;
Create Table 手机号常用号段表
(
编码 varchar2(3),
名称 varchar2(20),
简码 varchar2(20),
号段 varchar2(10),
号码长度 Number(2)
)tablespace ZL9BASEITEM;
Create Table 诊断注释类型
(
  编码   VARCHAR2(2),
  名称   VARCHAR2(20),
  简码   VARCHAR2(8),
  缺省标志 NUMBER(1)
)Tablespace ZL9BASEITEM;
Create Table 手术注释类型
(
  编码   VARCHAR2(2),
  名称   VARCHAR2(20),
  简码   VARCHAR2(8),
  缺省标志 NUMBER(1)
)Tablespace ZL9BASEITEM;
Create Table 手术情况
(
编码 varchar2(2),
名称 varchar2(20),
简码 varchar2(8),
缺省标志 Number(1)
)tablespace ZL9BASEITEM;
create table ZLMSG_LISTS
(
  code       VARCHAR2(30),
  bz_type    VARCHAR2(10),
  name       VARCHAR2(30),
  key_define VARCHAR2(500),
  note       VARCHAR2(1000),
  using	     NUMBER(1)
)
tablespace ZLMSGDATA;
create table ZLMSG_TODO
(
  id          NUMBER(18),
  msg_code    VARCHAR2(30),
  Key_Value   VARCHAR2(1000),
  state       NUMBER(1),
  create_time DATE,
  creator     VARCHAR2(100)
)
tablespace ZLMSGDATA initrans 20;
Create Table 路径上报变异原因(
    编码       VARCHAR2(2),
    名称       VARCHAR2(80),
    简码       VARCHAR2(10),
    缺省标志   NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 部门扩展项目
(
  编码 Number(3),
  名称 Varchar2(20),
  是否图片  Number(1)
)TABLESPACE zl9BaseItem
 Cache Storage(Buffer_Pool Keep);
Create Table 部门扩展信息
(
  部门id Number(18),
  项目 Varchar2(20),
  内容 Varchar2(1000),
  图片 Blob
)TABLESPACE zl9BaseItem
 Cache Storage(Buffer_Pool Keep);
CREATE TABLE 过敏源
(
  编码 VARCHAR2(5),
  名称 VARCHAR2(100),
  简码 VARCHAR2(10),
  缺省标志 NUMBER(1) default 0
)TABLESPACE zl9BaseItem;
Create Table 诊断前后注释 (
    编码 varchar2(2),
    名称 varchar2(20),
    简码 varchar2(10),
    分类 VARCHAR2(20) Default Null )
    TABLESPACE ZL9BASEITEM
;
Create Table 手术前后注释 (
    编码 varchar2(2),
    名称 varchar2(20),
    简码 varchar2(10),
    分类 VARCHAR2(20) Default Null )
    TABLESPACE ZL9BASEITEM
;
CREATE TABLE 人员扩展项目
(
  编码 Number(3),
  名称 Varchar2(20),
  是否图片  Number(1)
)TABLESPACE zl9BaseItem

 Cache Storage(Buffer_Pool Keep);
Create Table 人员扩展信息
(
  人员id Number(18),
  项目 Varchar2(20),
  内容 Varchar2(1000),
  图片 Blob
)TABLESPACE zl9BaseItem
 Cache Storage(Buffer_Pool Keep);
Create Table 身份证未录原因(
    编码 VARCHAR2(2),
    名称 VARCHAR2(50),
    简码 VARCHAR2(10),
	缺省标志 NUMBER(1) default 0,
    说明 VARCHAR2(50))
    TABLESPACE zl9BaseItem;
Create Table 期间表(
    期间 VARCHAR2(6),
    开始日期 Date,
    终止日期 Date)
    TABLESPACE zl9BaseItem;
Create Table 时间段(
    时间段 varchar2(10),
    开始时间 date,
    终止时间 date,
    缺省时间 DATE,
    提前时间 Date,
    提前颜色 Varchar2(20),
    站点 Varchar2(3),
    号类 varchar2(10),
    出诊预留时间 number(18),
    休息时段 varchar2(200))
    TABLESPACE zl9BaseItem 
;
Create Table 号码控制表(
    项目序号 NUMBER(3),
    项目名称 VARCHAR2(20) Not Null,
    最大号码 VARCHAR2(20),
    自动补缺 NUMBER(1) Default 0,
		编号规则 NUMBER(1))
    TABLESPACE zl9BaseItem
    PCTFREE 20 initrans 20
    Cache Storage(Buffer_Pool Keep);
CREATE TABLE 科室号码表(
	项目序号 NUMBER(3),
	科室ID NUMBER(18),
	编号 VARCHAR2(1),
	最大号码 VARCHAR2(64))
    TABLESPACE zl9BaseItem
    PCTFREE 20 initrans 20
    Cache Storage(Buffer_Pool Keep);
CREATE TABLE 号码编号表(
	项目序号 NUMBER(3),
	号码前缀 VARCHAR2(20),
	日期 DATE,
	最大号码 VARCHAR2(20))
    TABLESPACE zl9BaseItem
    PCTFREE 20 initrans 20
    Cache Storage(Buffer_Pool Keep);
Create Table 床位编制分类(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(6),
	  符号 VARCHAR2(5),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 单据环节控制(
    单据 number(2),
    环节  number(2),
    内容 VARCHAR2(100))
    TABLESPACE zl9BaseItem;
Create Table 合约单位(
    ID NUMBER(18),
    上级id NUMBER(18),
    编码 VARCHAR2(10) Not Null,
    名称 VARCHAR2(100) Not Null,
    简码 VARCHAR2(10),
    末级 NUMBER(1) default 0,
    地址 VARCHAR2(100),
    电话 VARCHAR2(16),
    开户银行 VARCHAR2(50),
    帐号 VARCHAR2(50),
    联系人 VARCHAR2(20),
    电子邮件 varchar2(50),
    说明 varchar2(2000),
    站点 Varchar2(3),
    建档时间 Date,
    撤档时间 Date,
    社会信用代码 varchar2(50))
    TABLESPACE zl9BaseItem 
;
Create Table 社区目录(
		序号 NUMBER(5),
		名称 VARCHAR2(100),
		说明 VARCHAR2(200),
		启用 NUMBER(1),
		部件名 VARCHAR2(50))
		TABLESPACE zl9BaseItem;
Create Table 社区参数(
		社区 NUMBER(5),
		参数号 NUMBER(5),
		参数名 VARCHAR2(50),
		参数值 VARCHAR2(100),
		缺省值 VARCHAR2(100),
		参数说明 VARCHAR2(200))
		TABLESPACE zl9BaseItem;
Create Table 报表条件(
    编号 NUMBER(18),
    序号 NUMBER(3),
    条件 VARCHAR2(255))
    TABLESPACE zl9BaseItem;
Create Table 移动查房接口(
    编号 Number(2),
    名称 Varchar2(20),
		参数 Varchar2(1000),
    启用 Number(1))
    TABLESPACE zl9BaseItem;
Create Table 特殊符号(
	类别 VARCHAR2(10),
	编码 VARCHAR2(4),
	字符 VARCHAR2(2))
	TABLESPACE zl9BaseItem;
CREATE TABLE 排队号码表(
	队列日期 DATE ,
	排队名称 VARCHAR2 (100),
	最大号码  VARCHAR2(20))
TABLESPACE zl9BaseItem;
Create Table 排队叫号队列 (
    ID NUMBER(18),
    病人ID NUMBER(18),
    队列名称 VARCHAR2(60),
    业务id Number(18),
    科室ID NUMBER(18),
    排队号码 varchar2(20),
    排队标记 VARCHAR2(10),
    患者姓名 VARCHAR2(100),
    诊室 VARCHAR2(20),
    医生姓名 VARCHAR2(64),
    优先 NUMBER(1),
    回诊序号 Number(18),
    排队时间 DATE,
    排队状态 NUMBER(1) default 0,
    是否分时点 number(1) DEFAULT 0,
    呼叫医生 VARCHAR2(20),
    业务类型 number(5),
    呼叫时间 date,
    备注 Varchar2(64),
    排队序号 Varchar2(30))
    TABLESPACE zl9BaseItem
;
Create Table 排队优先原因
(
  编码 VARCHAR2(5),
  名称 VARCHAR2(64),
  简码 VARCHAR2(20),
  使用频率 number(5) default 0
)tablespace ZL9BASEITEM;
create table 排队语音呼叫
(
  ID       NUMBER(18),
  呼叫内容 VARCHAR2(1000),
  队列ID   NUMBER(18),
  队列名称 VARCHAR2(20),
  业务类型 number(5),
  生成时间 Date,
  站点 VARCHAR2(50))
  TABLESPACE zl9BaseItem
  PCTFREE 20 initrans 100;
create table 排队LED显示部件
(
  部件类型 NUMBER(3),
  部件名   VARCHAR2(20),
  启用     NUMBER(1),
  说明     VARCHAR2(50))
  TABLESPACE zl9BaseItem;
Create Table 部门性质分类(
    编码 VARCHAR2(2),
    名称 VARCHAR2(10),
    简码 VARCHAR2(6),
    服务病人 NUMBER(1),
    说明 VARCHAR2(200))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 人员性质分类(
    编码 VARCHAR2(2),
    名称 VARCHAR2(10),
    简码 VARCHAR2(6),
    说明 VARCHAR2(200))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 部门环境类别(
	编码 varchar2(3),
	名称 varchar2(10) not Null,
	简码 varchar2(10),
	范围 varchar2(100))
	TABLESPACE zl9BaseItem;
Create Table 部门表(
    ID NUMBER(18),
    上级id NUMBER(18),
    编码 VARCHAR2(10) Not Null,
    名称 varchar2(100) Not Null,
    简码 varchar2(100),
    位置 VARCHAR2(50),
    末级 NUMBER(1) default 0,
    建档时间 Date,
    撤档时间 Date,
    环境类别 VARCHAR2(10),
    部门负责人 number(18),
    站点 Varchar2(3),
    顺序 Number(3),
    最后修改时间 Date,
    别名 varchar2(100),
    位置编码 VARCHAR2(4))
    TABLESPACE zl9BaseItem
    Cache Storage(Buffer_Pool Keep) 
;
Create Table 人员表(
    ID NUMBER(18),
    编号 VARCHAR2(6) Not Null,
    姓名 VARCHAR2(20),
    简码 VARCHAR2(8),
    身份证号 VARCHAR2(18),
    出生日期 DATE,
    性别 VARCHAR2(4),
    民族 VARCHAR2(20),
    工作日期 DATE,
    办公室电话 VARCHAR2(20),
    电子邮件 VARCHAR2(20),
    执业类别 VARCHAR2(3),
    执业范围 VARCHAR2(20),
    执业证号 Varchar2(50),
    管理职务 VARCHAR2(30),
    专业技术职务 VARCHAR2(50),
    聘任技术职务 NUMBER(1),
    学历 VARCHAR2(10),
    所学专业 VARCHAR2(2),
    留学时间 NUMBER(2),
    留学渠道 VARCHAR2(10),
    接受培训 VARCHAR2(10),
    科研课题 VARCHAR2(10),
    个人简介 VARCHAR2(1000),
    建档时间 Date,
    撤档时间 Date,
    撤档原因 Varchar2(100),
    别名 Varchar2(100),
    签名 varchar2(20),
    签名图片 BLOB,
    资格证书号 varchar2(50),
    执业开始日期 date,
    处方权标志 number(1),
    手术等级 Varchar2(20),
    站点 Varchar2(3),
    移动电话 number(11),
    顺序 Number(3),
    最后修改时间 Date,
    门诊特殊医嘱权限 VARCHAR2(10),
    住院特殊医嘱权限 VARCHAR2(10),
    帐号到期时间 Date)
    TABLESPACE zl9BaseItem
    Cache Storage(Buffer_Pool Keep) 
;
Create Table 人员照片(
    人员ID NUMBER(18),
    照片 BLOB)
    TABLESPACE zl9BaseItem
    PCTFREE 20;
Create Table 人员证书记录(
    ID NUMBER(18),
    人员ID NUMBER(18),
    CertDN VARCHAR2(300),
    CertSN VARCHAR2(100),
    SIGNCERT VARCHAR2(3000),
    EncCert VARCHAR2(2000),
    注册时间 DATE,
    是否停用 Number(1),
    停用记录 XMLType,
    时间戳证书 varchar2(3000),
    签章信息 clob)
    TABLESPACE zl9BaseItem 
;
Create Table 部门人员(
    部门id NUMBER(18),
    人员id NUMBER(18),
    缺省 NUMBER(1))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 上机人员表(
    用户名 VARCHAR2(20),
    人员id NUMBER(18),
    系统升级锁定 number(1),
    登录密码 varchar2(100))
    TABLESPACE zl9BaseItem
    Cache Storage(Buffer_Pool Keep)
;
Create Table 部门性质说明(
    工作性质 VARCHAR2(10),
    部门id NUMBER(18),
    服务对象 NUMBER(3))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 人员性质说明(
    人员ID NUMBER(18),
    人员性质 VARCHAR2(10))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 部门安排(
    部门id NUMBER(18),
    星期 NUMBER(1),
    开始时间 date,
    终止时间 date)
    TABLESPACE zl9BaseItem;
Create Table 病区科室对应(
    病区id NUMBER(18),
    科室id NUMBER(18))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
CREATE TABLE 电子签名启用部门(
    部门ID NUMBER(18),
    场合 NUMBER(5))
    TABLESPACE ZL9BASEITEM
    Cache Storage(Buffer_Pool Keep);
Create Table 管理职务(
    编码 VARCHAR2(2),
    名称 VARCHAR2(30),
    简码 VARCHAR2(10))
    TABLESPACE zl9BaseItem;
Create Table 执业类别(
    编码 VARCHAR2(3),
    名称 VARCHAR2(20),
    简码 VARCHAR2(8),
    分类 VARCHAR2(16))
    TABLESPACE zl9BaseItem;
Create Table 专业技术职务(
    编码 VARCHAR2(3),
    名称 VARCHAR2(50),
    简码 VARCHAR2(10),
    是否选择 NUMBER(1),
    标识符 Varchar2(5))
    TABLESPACE zl9BaseItem
;
Create Table 业务消息类型(
    编码 VARCHAR2(100),
    名称 VARCHAR2(100),
    说明 VARCHAR2(4000),
    保留天数 number(6),
    是否三方消息 number(1))
    TABLESPACE zl9BaseItem 
;
Create Table RIS预约调整原因(
	编码   VARCHAR2(10),
	名称   VARCHAR2(20),
	简码   VARCHAR2(10),
	原因说明 VARCHAR2(100))
	TABLESPACE zl9BaseItem;
----------------------------------------------------------------------------
--[[2.医保基础]]
----------------------------------------------------------------------------
Create Table 保险类别(
    序号 NUMBER(3),
    名称 VARCHAR2(20),
    说明 VARCHAR2(100),
    医院编码 VARCHAR2(20),
    是否固定 NUMBER(1),
    是否禁止 NUMBER(1),
    具有中心 NUMBER(1),
    医保部件 VARCHAR2(30),
    外挂 NUMBER (1),
    项目提示 NUMBER (1) DEFAULT 0,
    医保包 varchar2(20),
    保险机构编码 varchar2(50))
    TABLESPACE zl9BaseItem
;
Create Table 保险中心目录(
    险类 NUMBER(3),
    序号 NUMBER(5),
    编码 VARCHAR2(6),
    名称 VARCHAR2(20))
    TABLESPACE zl9BaseItem;
Create Table 保险参数(
    险类 NUMBER(3),
    中心  NUMBER(5),
    参数名 VARCHAR2(20),
    参数值 VARCHAR2(40),
    序号 NUMBER(2),
    是否固定 NUMBER(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 保险人群(
    险类 NUMBER(3),
    序号 NUMBER(1),
    名称 VARCHAR2(10))
    TABLESPACE zl9BaseItem;
Create Table 保险费用档(
    险类 NUMBER(3),
    中心 NUMBER(5),
    档次 NUMBER(3),
    名称 VARCHAR2(25),
    下限 NUMBER(16,5),
    上限 NUMBER(16,5))
    TABLESPACE zl9BaseItem;
Create Table 保险年龄段(
	险类 NUMBER(3),
	中心 NUMBER(5),
	在职 NUMBER(1),
	年龄段 NUMBER(3),
	名称 VARCHAR2(20),
	下限 NUMBER(3),
	上限 NUMBER(3),
	全额统筹 number(1),
	无起付线 number(1),
	无封顶线 number(1))
    TABLESPACE zl9BaseItem;
Create Table 保险支付比例(
	险类 NUMBER(3),
	中心 NUMBER(5),
	在职 NUMBER(1),
	年龄段 NUMBER(3),
	档次 NUMBER(3),
	年度 NUMBER(4),
	比例 NUMBER(16,5))
    TABLESPACE zl9BaseItem;
Create Table 保险支付限额(
	险类 NUMBER(3),
	中心 NUMBER(5),
	年度 NUMBER(4),
	性质 VARCHAR2(1),
	金额 NUMBER(16,5))
    TABLESPACE zl9BaseItem;
Create Table 保险项目(
    险类 NUMBER(3),
    编码 VARCHAR2(20),
    名称 VARCHAR2(100),
    简码 VARCHAR2(30),
    大类编码 VARCHAR2(6),
    附注 VARCHAR2(50))
    TABLESPACE zl9BaseItem;
Create Table 保险支付大类(
	险类 NUMBER(3),
	ID NUMBER(18),
	编码 VARCHAR2(6),
	名称 VARCHAR2(40),
	简码 VARCHAR2(10),
	性质 NUMBER(3),
	算法 NUMBER(3),
	统筹比额 NUMBER(16,5),
	特准定额 NUMBER(16,5),
	特准天数 NUMBER(5),
	服务对象 NUMBER(1),
	是否医保 NUMBER(1) DEFAULT 1)
    TABLESPACE zl9BaseItem;
Create TABLE 大类档次比例(
	大类ID		number(18),
	档次		number(3),
	上限		number(16,5),
	下限            number(16,5),
	比例		number(16,5))
	TABLESPACE	ZL9BASEITEM;
CREATE TABLE 医保对照类别(
    险类 NUMBER(3),
	编码 NUMBER(1),		--default=0，表示缺省
	名称 VARCHAR2(20),
    说明 VARCHAR2(200))
    TABLESPACE ZL9BASEITEM;
CREATE TABLE 医保对照明细(
    险类 NUMBER(3),
	  类别 NUMBER(1) DEFAULT 0,		--default=0，表示缺省
	  收费细目ID NUMBER(18),
    项目编码 VARCHAR2(20),
	  说明 VARCHAR2(256))
    TABLESPACE ZL9BASEITEM;
Create Table 保险病种(
	险类 NUMBER(3),
	ID NUMBER(18),
	编码 VARCHAR2(6),
	名称 VARCHAR2(100),
	简码 VARCHAR2(10),
	类别 VARCHAR2(1),
	特殊封顶线 NUMBER(1),
	封顶线金额 NUMBER(16,5))
    TABLESPACE zl9BaseItem;
Create Table 保险支付项目(
    险类 NUMBER(3),
    收费细目ID NUMBER(18),
    大类ID NUMBER(18),
    项目编码 VARCHAR2(20),
    项目名称 VARCHAR2(100),
    附注 VARCHAR2(50),
    是否医保 NUMBER(1) DEFAULT 1,
    要求审批 NUMBER (1),
    保险费用等级 varchar2(50))
    TABLESPACE zl9BaseItem
;
Create Table 保险特准项目(
	病种ID NUMBER(18),
	收费细目ID NUMBER(18),
	大类 NUMBER(1) DEFAULT 0,
	性质 NUMBER(1) DEFAULT 0)
    TABLESPACE zl9BaseItem;
----------------------------------------------------------------------------
--[[3.病人病案基础]]
----------------------------------------------------------------------------
Create Table 病人实名信息变动    
(
实名ID    Number(18),
变更项目  Varchar2(20),  
原信息       Varchar2(100),  
新信息       Varchar2(100),  
变更时间  Date,  
变更人       Varchar2(20),  
变更原因  Varchar2(100)
) tablespace ZL9PATIENT;  

Create Table 实名认证接口    
(
ID       Number(18),
编号     Varchar2(5),  
接口名    Varchar2(50),  
部件名    Varchar2(100),
说明     Varchar2(100),  
是否启用 Number(1)
)tablespace ZL9PATIENT;

Create Table 实名认证接口日志    
(
ID        Number(18),  
实名ID    Number(18),  
接口ID    Number(18),  
入参      CLOB,  
出参      CLOB,  
认证者     Number(1),
调用结果  Number(1),
调用时间  Date,  
调用人    Varchar2(20),  
备注     Varchar2(100)
) tablespace ZL9PATIENT;

Create Table 病人实名信息 
(
   实名ID            Number(18),
   病人ID            Number(18),
   姓名              Varchar2(100),
   性别              Varchar2(4),
   出生日期          Date,
   国籍              Varchar2(30),
   民族              Varchar2(20) ,
   出生地点          Varchar2(100),
   住址              Varchar2(100),
   身份证号          Varchar2(18),
   身份证类型        Number(1),
   陪诊人姓名        Varchar2(100),
   陪诊人性别        Varchar2(4),
   陪诊人出生日期    Date,
   陪诊人国籍        Varchar2(30),
   陪诊人民族        Varchar2(20),
   陪诊人住址        Varchar2(100),
   陪诊人身份证号    Varchar2(18),
   陪诊人身份证类型	 Number(1),	
   陪诊人关系        Varchar2(30),
   手机号            Varchar2(50),
   备注              Varchar2(100),
   认证状态          Number(1),
   建档时间          Date,
   建档人            Varchar2(20),
   是否停用          Number(1),
   停用时间          Date,
   更新时间          Date,
   更新人            Varchar2(20)
) tablespace ZL9PATIENT;

Create Table 病人实名证件	
(
ID	          Number(18),	
实名ID	      Number(18),	
证件类型	    Varchar2(50),
证件号码	    Varchar2(20),
备注          Varchar2(100),
所有者	   Number(1)
) tablespace ZL9PATIENT;

Create Table 病人实名证件图片		
(
证件ID	  Number(18),	
序号	    Number(3),	
证件图片	Blob,	
备注	    Varchar2(100)
) tablespace ZL9PATIENT;

CREATE TABLE 切口部位(
  编码 VARCHAR2(5),
  名称 VARCHAR2(100),
  简码 VARCHAR2(100),
  缺省标志 NUMBER(1) default 0
)TABLESPACE zl9BaseItem;
Create Table 住院死亡原因(
    编码 VARCHAR2(2),
    名称 VARCHAR2(50),
    简码 VARCHAR2(10),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem ;
Create Table 区域(
    编码 VARCHAR2(15),
    上级编码 VARCHAR2(45),
    名称 VARCHAR2(100),
    简码 VARCHAR2(100),
    五笔码 varchar2(100),
    是否虚拟 number(1),
    是否不显示 number(1),
    级数 number(2),
    缺省标志 number(2) DEFAULT 0,
    邮编 varchar2(10))
    TABLESPACE zl9BaseItem 
;
Create Table 地区(
    编码 VARCHAR2(8),
    名称 VARCHAR2(50),
    简码 VARCHAR2(20),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 国籍(
	编码 VARCHAR2(3),
	名称 VARCHAR2(30),
	英文简称 varchar2(30),
	简码 VARCHAR2(10),
	缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 婚姻状况(
    编码 VARCHAR2(1),
    名称 VARCHAR2(4),
    简码 VARCHAR2(4),
    缺省标志 NUMBER(1) default 0,
	国标代码 varchar2(2))
    TABLESPACE zl9BaseItem;
Create Table 民族(
    编码 VARCHAR2(2),
    名称 VARCHAR2(20),
    简码 VARCHAR2(10),
    缺省标志 NUMBER(1) default 0,
	国标代码 varchar2(2),
	罗马拼写法 varchar2(20),
	字母代码 varchar2(10))
    TABLESPACE zl9BaseItem;
Create Table 社会关系(
    编码 VARCHAR2(2),
    名称 VARCHAR2(30),
    简码 VARCHAR2(10),
    缺省标志 NUMBER(1) default 0,
    是否停用 NUMBER(1))
    TABLESPACE zl9BaseItem
;
Create Table 身份(
    编码 VARCHAR2(2),
    名称 VARCHAR2(10),
    简码 VARCHAR2(10),
    优先级 NUMBER(1))
    TABLESPACE zl9BaseItem;
Create Table 性别(
    编码 VARCHAR2(1),
    名称 VARCHAR2(4),
    简码 VARCHAR2(4),
    缺省标志 NUMBER(1) default 0,
    国标代码 varchar2(2))
    TABLESPACE zl9BaseItem;
Create Table 学历(
    编码 VARCHAR2(2),
    名称 VARCHAR2(10),
    简码 VARCHAR2(10),
    缺省标志 NUMBER(1) default 0,
    国标代码 varchar2(2))
    TABLESPACE zl9BaseItem;
Create Table 血型(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(2),
    缺省标志 NUMBER(1) default 0,
	  国标代码 varchar2(2))
    TABLESPACE zl9BaseItem;
Create Table 职业(
    编码 VARCHAR2(3),
    名称 VARCHAR2(80),
    简码 VARCHAR2(10),
    病案名称 VARCHAR2(20),
    缺省标志 NUMBER(1) default 0,
    是否停用 NUMBER(1))
    TABLESPACE zl9BaseItem
;
Create Table 病人类型(
    编码 VARCHAR2(3),
    名称 VARCHAR2(50),
    简码 VARCHAR2(10),
    颜色 Number(18),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 分娩方式(
    编码 VARCHAR2(2),
    名称 VARCHAR2(20),
    简码 VARCHAR2(10),
    缺省标志 NUMBER(1) Default 0)
    TABLESPACE zl9BaseItem;
Create Table 胎儿状况(
    编码 VARCHAR2(2),
    名称 VARCHAR2(20),
    简码 VARCHAR2(10),
    缺省标志 NUMBER(1) Default 0)
    TABLESPACE zl9BaseItem;
--病人咨询
----------------------------------------------------------------------------
Create Table 咨询附加选项(
	序号 number(3),
	名称 varchar2(30),
	内容 varchar2(250))
	TABLESPACE zl9BaseItem;
Create Table 咨询表格元素(
	序号 number(5),
	名称 varchar2(30),
	列数 number(2),
	列宽 varchar2(250),
	行数 number(2),
	行高 varchar2(250),
	合并行 varchar2(250),
	合并列 varchar2(250),
	颜色 number(18))
	TABLESPACE zl9BaseItem;
Create Table 咨询表格内容(
	表号 number(5),
	行号 number(2),
	列号 number(2),
	内容 varchar2(200),
	对齐 number(1),
	颜色 number(18),
	字体 varchar2(200))
	TABLESPACE zl9BaseItem;
Create Table 咨询图片元素(
	序号 number(5),
	性质 number(2),
	名称 varchar2(30),
	类型 number(3),
	图形 blob,
	宽度 number(18),
	高度 number(18),
	固定 number(1) default 0,
	修改日期 Date)
	TABLESPACE zl9BaseItem;
Create Table 咨询页面目录(
	页面序号 number(18),
	上级序号 number(18),
	编码 varchar2(10),
	页面名称 varchar2(30),
	简码 varchar2(15),
	固定页面 number(1),
	页面风格 number(3),
	宣传标语 number(5),
	页面背景 number(5),
	背景音乐 number(18),
	命令参数 Varchar2(100),
	末级 number(1))
	TABLESPACE zl9BaseItem;
Create Table 咨询页面排列(
	序号 number(5),
	父序号 number(5),
	名称 varchar2(30),
	页面 number(18),
	页面图标 number(5),
	字体 varchar2(20),
	大小 Number(2),
	字形 Number(1),
	颜色 Number(18))
	TABLESPACE zl9BaseItem;
Create Table 咨询段落目录(
	页面序号 number(18),
	段落序号 number(5),
	段落类型 Number(3),
	标题文本 varchar2(30),
	标题图标 number(5),
	标题隐藏 number(1),
	标题位置 number(1),
	标题字体 varchar2(50),
	返回页首 number(1),
	段落文本 clob,
	段落字体 varchar2(50),
	插表序号 number(5),
	插表位置 number(1),
	插图序号 number(5),
	插图位置 number(1))
	TABLESPACE zl9BaseItem;
Create Table 咨询段落链接(
	页面序号 number(18),
	段落序号 number(5),
	链接页面 number(18),
	页内段号 number(18))
	TABLESPACE zl9BaseItem;
Create Table 咨询专家清单(
	序号 number(5),
	人员id number(18),
	科室id number(18))
	TABLESPACE zl9BaseItem;
Create Table 咨询广告序列(
	序号 number(5),
	图片序号 number(5))
	TABLESPACE zl9BaseItem;
--病案首页
----------------------------------------------------------------------------
Create Table ICU类型
(
  编码 VARCHAR2(20),
  名称 VARCHAR2(30),
  简码 VARCHAR2(20)
)
TABLESPACE zl9BaseItem;
CREATE TABLE 感染部位(
  编码 VARCHAR2(6),
  名称 VARCHAR2(20),
  简码 VARCHAR2(10),
  缺省标志 NUMBER(1) default 0
)TABLESPACE zl9BaseItem;
Create Table 器械导管目录
(
  编码 VARCHAR2(20),
  名称 VARCHAR2(30),
  简码 VARCHAR2(20)
)
TABLESPACE zl9BaseItem;
Create Table 医院感染目录
(
  编码 VARCHAR2(20),
  名称 VARCHAR2(30),
  简码 VARCHAR2(20)
)
TABLESPACE zl9BaseItem;
Create Table 病原学目录
(
  编码 VARCHAR2(20),
  名称 VARCHAR2(50),
  简码 VARCHAR2(20)
)
TABLESPACE zl9BaseItem;
CREATE TABLE  医学警示 (
	编码 VARCHAR2(4),
	名称 varchar2(20),
	简码 varchar2(10) ,
	缺省标志 NUMBER (1) DEFAULT 0) 
	TABLESPACE zl9BaseItem;
CREATE TABLE  证件类型 (
	编码 VARCHAR2(4),
	名称 varchar2(20),
	简码 varchar2(10), 
	缺省标志 NUMBER (1) DEFAULT 0) 
	TABLESPACE zl9BaseItem;
CREATE TABLE 病案项目(
	编码 varchar2(3),
	名称 varchar2(20),
	内容 varchar2(1000))
	TABLESPACE zl9BaseItem 
	Cache Storage(Buffer_Pool Keep);
Create Table 借阅理由(
    编码	Varchar2(10),
    名称	Varchar2(30),
    简码	Varchar2(30),
    缺省标志	Number(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 病案费目(
    编码 VARCHAR2(4),
    上级 VARCHAR2(4),
    名称 VARCHAR2(30),
    简码 VARCHAR2(10),
    末级 NUMBER(1) DEFAULT 0)
    TABLESPACE zl9BaseItem;
Create Table 病案类别(
    编码 VARCHAR2(2),
    名称 VARCHAR2(10),
    简码 VARCHAR2(6),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 医嘱审核用语(
  编码 VARCHAR2(2),
  名称 VARCHAR2(50),
  简码 VARCHAR2(10)
  ) TABLESPACE zl9BaseItem;
Create Table 病情(
    编码 VARCHAR2(1),
    名称 VARCHAR2(20),
    简码 VARCHAR2(4),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 住院死亡期间(
    编码 VARCHAR2(2),
    名称 VARCHAR2(50),
    简码 VARCHAR2(10),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem ;
Create Table 入院方式(
    编码 VARCHAR2(1),
    名称 VARCHAR2(8),
    简码 VARCHAR2(4),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 医疗机构(
    编码 VARCHAR2(4),
    名称 VARCHAR2(50),
    简码 VARCHAR2(25),
    上级 Varchar2(4),
    末级 Number(1))
    TABLESPACE zl9BaseItem 
;
Create Table 入院属性(
    编码 VARCHAR2(1),
    名称 VARCHAR2(8),
    简码 VARCHAR2(4),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 病人去向(
    编码 VARCHAR2(1),
    名称 VARCHAR2(20),
    简码 VARCHAR2(4),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 出院方式(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(4),
    缺省标志 NUMBER(1) default 0,
    是否停用 NUMBER(1))
    TABLESPACE zl9BaseItem
;
Create Table 出院转入(
    编码 VARCHAR2(5),
    名称 VARCHAR2(100),
    简码 VARCHAR2(50),
    缺省标志 NUMBER(1) default 0,
    上级 Varchar2(5),
    末级 Number(1))
    TABLESPACE zl9BaseItem 
;
Create Table 治疗结果(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(4),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 住院目的(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(6),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 感染因素(
  编码 VARCHAR2(3),
  名称 VARCHAR2(100))
  TABLESPACE zl9BaseItem;
Create Table 不良事件(
       编码 VARCHAR2(3),
       名称 VARCHAR2(100))
  TABLESPACE zl9BaseItem;
Create Table 手术切口愈合(
       编码 VARCHAR2(2),
       名称 VARCHAR2(100))
  TABLESPACE zl9BaseItem;
Create Table 医疗类别(
    编码 VARCHAR2(2),
    名称 VARCHAR2(10),
    简码 VARCHAR2(6),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 手术类型(
	编码 VARCHAR2(2),
	名称 VARCHAR2(20))
	TABLESPACE zl9BaseItem;
Create Table 手术操作类型(
    编码 Varchar2(2), 
    名称 Varchar2(20)
)TABLESPACE zl9BaseItem;
Create Table 抢救病因分类(
编码 Varchar2(2),
名称 Varchar2(50),
简码 Varchar2(10))
TABLESPACE ZL9BASEITEM 
Cache Storage(Buffer_Pool Keep);
Create Table 分化程度(
    编码 VARCHAR2(2),
    名称 VARCHAR2(20),
    简码 VARCHAR2(10),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE ZL9BASEITEM;
--疾病诊断
----------------------------------------------------------------------------
Create Table 疾病编码类别(
	编码 VARCHAR2(1),
	类别 VARCHAR2(20),
	说明 VARCHAR2(50),
	优先级 NUMBER(3),
	是否分类 NUMBER(1) default 1)
	TABLESPACE zl9BaseItem;
Create Table 疾病编码分类(
    ID NUMBER(18),
    上级ID NUMBER(18),
    序号 NUMBER(6),
    名称 VARCHAR2(150),
    简码 VARCHAR2(20),
    类别 VARCHAR2(1),
    编码范围 varchar2(60),
    是否病人 NUMBER(1) default 1,
    建档时间 Date,
    撤档时间 Date,
    章节 Varchar2(2),
    编码 VARCHAR2(20))
    TABLESPACE zl9BaseItem 
;
Create Table 疾病编码目录(
    ID number(18),
    编码 VARCHAR2(20),
    序号 NUMBER(3),
    附码 VARCHAR2(15),
    统计码 VARCHAR2(10),
    名称 VARCHAR2(150) not Null,
    简码 VARCHAR2(20),
    五笔码 Varchar2(20),
    适用范围 Number(1),
    说明 VARCHAR2(200),
    性别限制 VARCHAR2(4),
    疗效限制 VARCHAR2(4),
    手术类型 VARCHAR2(20),
    分娩 VARCHAR2(1),
    类别 VARCHAR2(1),
    分类ID NUMBER(18),
    建档时间 DATE,
    撤档时间 DATE,
    手术操作类型 Number(1),
    章节 Varchar2(2))
    TABLESPACE zl9BaseItem
    Cache Storage(Buffer_Pool Keep)
;
Create Table 疾病编码科室(
	疾病ID Number(18),
	科室ID Number(18),
	人员ID Number(18))
	TABLESPACE zl9BaseItem;
Create Table 疾病病种对应(
    疾病ID NUMBER(18),
    险类 NUMBER(3),
    病种ID NUMBER(18))
    TABLESPACE zl9BaseItem;
Create Table 最高诊断依据(
    编码 VARCHAR2(2),
    名称 VARCHAR2(30),
    简码 VARCHAR2(10),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE ZL9BASEITEM;
----------------------------------------------------------------------------
--[[4.费用基础]]
----------------------------------------------------------------------------
CREATE TABLE 电子票据类别(
	编号   number(3),
	名称   varchar2(50),
	简码   varchar2(20),
	是否启用   number(2),
	部件 varchar2(100),
	包名称 varchar2(100))
 TABLESPACE zl9Expense;

Create Table 电子票据站点控制(
 场合 Number(2),
 站点 varchar2(50))
 TABLESPACE zl9Expense;

Create Table 电子票据异常记录(
  ID Number(18),
  操作场景 number(2),
  业务类型 number(2),
  记录标志 number(2),
  单据号 varchar2(20),
  业务ID number(18),
  电子票据id number(18),
  病人ID number(18),
  姓名 varchar2(100),
  性别 varchar2(4),
  年龄 varchar2(20),
  门诊号 number(18),
  住院号 number(18),
  是否换开 number(2),
  票据信息 CLOB,
  操作员编号 varchar2(6),
  操作员姓名 varchar2(50),
  登记时间 Date)
 TABLESPACE zl9Expense;

CREATE TABLE 电子票据开票点(
  ID   Number(18),
  上级ID   Number(18),
  编码   varchar2(20),
  名称   varchar2(50),
  简码   varchar2(20),
  院区   varchar2(50),
  客户端   varchar2(50),
  部门ID   number(18),
  位置   varchar2(100),
  末级   number(2),
  建档时间   date,
  撤档时间   date)
 TABLESPACE zl9Expense;

CREATE TABLE 票据开票点对照(
    Id Number(18),
  开票点ID Number(18),
  人员ID Number(18),
  客户端 varchar2(50))
TABLESPACE zl9Expense;

Create Table 电子票据使用记录(
 ID Number(18),
 票种 number(2),
 记录状态 number(2),
 结算ID number(18),
 病人ID number(18),
 姓名 varchar2(100),
 性别 varchar2(4),
 年龄 varchar2(20),
 门诊号 number(18),
 住院号 number(18),
 代码 Varchar2(50),
 号码 Varchar2(50),
 检验码 Varchar2(20),
 凭证代码 Varchar2(50),
 凭证号码 Varchar2(50),
 凭证检验码 Varchar2(20),
 票据金额 number(16,5),
 生成时间 varchar2(30),
 URL内网  varchar2(2000),
 URL外网  varchar2(2000),
 原票据ID number(18),
 是否换开 number(2),
 纸质发票号 Varchar2(50),
 打印ID Number(18),
 退款id number(18),
 备注 varchar2(4000),
 开票点 varchar2(100),
 系统来源 varchar2(100),
 操作员编号 varchar2(6),
 操作员姓名 varchar2(50),
 登记时间 Date,
 待转出 number(3))
 TABLESPACE zl9Expense PCTFREE 5 initrans 20;

CREATE TABLE 电子票据二维码 (
 使用记录ID number(18),
 二维码 clob,
 待转出 number(3)) 
TABLESPACE zl9Expense PCTFREE 20;

Create Table 病人结算异常记录
(
  id              Number(18),
  操作场景        Number(2),
  是否作废        Number(2),
  业务ID          Number(18),
  是否病历费      Number(2),
  病人ID          Number(18),
  主页ID          Number(5),
  姓名            Varchar2(100),
  性别            Varchar2(4),
  年龄            Varchar2(20),
  门诊号          Number(18),
  住院号          Number(18),
  预交单号	      Varchar2(20),
  预交金额        Number(16,5),
  医疗卡单号      Varchar2(20),
  卡费金额        Number(16,5),
  卡类别ID	       Number(18),
  卡类别名称      Varchar2(50),
  发卡卡号	      Varchar2(100),
  同步状态	      Number(2),
  交易信息	      Clob,
  登记时间        Date,
  操作员姓名      Varchar2(20)
)Tablespace zl9Expense;

Create Table 押金类别
(编码 Varchar2(2),
名称  Varchar2(20),
简码 Varchar2(10),
缺省标志 Number(1)
)Tablespace ZL9BASEITEM;

Create Table 病人押金记录 (
    ID Number(18),
    NO Varchar2(8),
    记录状态 Number(3),
    实际票号 Varchar2(20),
    押金类别 Varchar2(20),
    病人ID Number(18),
    主页ID Number(5),
    科室ID Number(18),
    缴款单位 Varchar2(50),
    单位开户行 Varchar2(50),
    单位帐号 Varchar2(50),
    摘要 Varchar2(50),
    金额 Number(16,2),
    结算方式 Varchar2(20),
    结算号码 Varchar2(30),
    收款时间 Date,
    是否门诊 Number(2),
    卡类别ID Number(18),
    缴款组ID Number(18),
    卡号 Varchar2(50),
    交易流水号 Varchar2(50),
    交易说明 Varchar2(500),
    交易人员 Varchar2(20),
    交易时间 Date,
    校对标志 number(2),
    操作员编号 Varchar2(6),
    操作员姓名 Varchar2(20),
    待转出 Number(3),
    姓名 Varchar2(100),
    性别 Varchar2(4),
    年龄 Varchar2(20),
    门诊号 Number(18),
    住院号 Number(18),
    付款方式名称 Varchar2(20))
    Tablespace zl9Expense 
;

Create Table 费用结算对照(
    结帐ID Number(18),
    费用ID Number(18),
    是否重收 Number(1),
	门诊标志 Number(3),
    结帐金额 number(16,5),
    操作员编号 Varchar2(6),
    操作员姓名 Varchar2(20),
    待转出 Number(3)
    )
    TABLESPACE zl9Expense;
Create Table 三方交易记录(
    类别 varchar2(50),
    流水号 varchar2(50),
    状态 number(2),
    卡号 varchar2(50),
    交易说明 varchar2(500),
    业务结算ID number(18),
    交易时间 DATE,
    业务类型 number(2))
    TABLESPACE ZL9EXPENSE 
;
Create Global Temporary Table 临时票据打印内容(
  ID NUMBER(18),
  性质 NUMBER(3),
  NO VARCHAR2(8),
  开始票号 varchar2(50),
  领用ID number(18))
    ON COMMIT PRESERVE ROWS;
Create Table 收费价格等级(
 编码 Varchar2(2), 
 名称 Varchar2(30), 
 简码 Varchar2(10), 
 是否适用药品 Number(1) Default 0,
 是否适用卫材 Number(1) Default 0, 
 是否适用普通项目 Number(1) Default 0, 
 建档时间 Date, 
 撤档时间 Date) 
Tablespace Zl9baseitem;
Create Table 收费价格等级应用(
    价格等级 Varchar2(30),
    站点 Varchar2(3),
    医疗付款方式 Varchar2(20),
    性质 NUMBER(2))
    Tablespace Zl9baseitem
;
Create Table 人民币面额(
    编码 VARCHAR2(2),
    名称 VARCHAR2(10),
    简码 VARCHAR2(6),
    面额 NUMBER(16,2))
    TABLESPACE zl9BaseItem;
CREATE TABLE 常用预交摘要(
  编码 varchar2(4),
  名称 varchar2(50),
  简码 VARCHAR2(20),
  缺省标志 number(1))
  TABLESPACE zl9BaseItem;
  
Create Table 常用挂号摘要
(
  编码   VARCHAR2(4),
  名称   VARCHAR2(50),
  简码   VARCHAR2(25),
  缺省标志 NUMBER(1))
Tablespace ZL9BASEITEM;
CREATE TABLE 常用退费原因(
	编码 varchar2(4),
	名称 varchar2(50),
	简码 VARCHAR2(20),
	缺省标志 number(2)
	) TABLESPACE zl9BaseItem;
Create Table 医价接口(
    编号 number(2),
    名称 Varchar2(20),
    医疗 number(1), --是否支持医疗价目控制
    药品 number(1), --是否支持药品价目控制
    卫材 number(1), --是否支持卫材价目控制
    选用 number(1)) --当前选择使用标志
    TABLESPACE zl9BaseItem;
Create Table 标准医价规范(
    项目编码	varchar2(20),
    项目名称	varchar2(200),
    拼音码   varchar2(10),
    项目别名 varchar2(100),
    计价单位 varchar2(200),
    项目内涵 varchar2(1000),
    除外内容 varchar2(1000),
    项目说明 varchar2(1000),
    项目价格 number(20,2),
    重复标志 char(1),
    医院等级 char(1),
    注销标志 char(1),
    财务编码 char(1),
    最高限价 number(20,2),
    最低限价 number(20,2),
    调价日期 Date)
    TABLESPACE zl9BaseItem;
Create Table 票据使用类别(
    编码 varchar2(3),
    名称 VARCHAR2(50),
    简码  VARCHAR2(25),
    缺省标志 NUMBER (1) DEFAULT 0)
    TABLESPACE zl9BaseItem;
Create Table 结算场合(
    编码 VARCHAR2(2),
    名称 VARCHAR2(10),
    简码 VARCHAR2(4))
    TABLESPACE zl9BaseItem;
Create Table 结算方式(
 编码 VARCHAR2(2),
 名称 VARCHAR2(20),
 简码 VARCHAR2(4),
 性质 NUMBER(2),
 应收款 NUMBER(1),
 应付款 NUMBER(1),
 缺省标志 NUMBER(1) default 0,
 是否固定 number(1) default 0)
 TABLESPACE zl9BaseItem;
Create Table 结算方式应用(
    应用场合 VARCHAR2(10),
    结算方式 VARCHAR2(20),
    缺省标志 NUMBER(1) default 0,
    付款方式 Varchar2(20))
    TABLESPACE zl9BaseItem
;
Create Table 医疗付款方式(
    编码 VARCHAR2(2),
    名称 VARCHAR2(20),
    简码 VARCHAR2(10),
    缺省标志 NUMBER(1) default 0,
    是否医保 number(1) default 0,
    是否公费 number(1) default 0,
    是否停用 NUMBER(1))
    TABLESPACE zl9BaseItem
;
Create Table 费用类型(
	编码 VARCHAR2(2),
	名称 VARCHAR2(20),
	简码 VARCHAR2(10),
	性质 VARCHAR2(1),
	缺省标志 NUMBER(2) default 0)
	TABLESPACE zl9BaseItem;
Create Table 费别(
    编码 VARCHAR2(2),
    名称 VARCHAR2(10) Not Null,
    简码 VARCHAR2(4),
	有效开始 DATE,
	有效结束 DATE,
	适用科室 NUMBER(1),
	属性     NUMBER(1),
	仅限初诊 NUMBER(1),
    缺省标志 NUMBER(1) default 0,
	服务对象 NUMBER(3) default 3,
    说明 VARCHAR2(50))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 费别明细(
    费别 VARCHAR2(10),
    收入项目id NUMBER(18),
    收费细目ID Number(18),
    段号 NUMBER(3) default 1,
    应收段首值 NUMBER(16,5),
    应收段尾值 NUMBER(16,5) default 10000000000,
    实收比率 NUMBER(16,5) default 100,
    计算方法 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 费别适用科室(
    费别 VARCHAR2(10),
    科室id NUMBER(18))
    TABLESPACE zl9BaseItem;
Create Table 收入项目(
    ID NUMBER(18),
    上级id NUMBER(18),
    编码 VARCHAR2(8) Not Null,
    名称 VARCHAR2(20) Not Null,
    简码 VARCHAR2(10),
    末级 NUMBER(1) default 0,
    公费 NUMBER(1) default 0,
    收据费目 VARCHAR2(20),
    病案费目 VARCHAR2(30),
    建档时间 Date,
    撤档时间 Date)
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 收据费目(
    编码 VARCHAR2(8),
    名称 VARCHAR2(20),
    简码 VARCHAR2(10))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 收据费目对应(
    收入项目ID NUMBER(18),
    场合 NUMBER(1),
    收据费目 VARCHAR2(20))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
CREATE TABLE 收费项目类别(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(10),
    固定 NUMBER(1),
	序号 NUMBER(2))
    TableSpace zl9BaseItem;
CREATE TABLE 收费分类目录(
    ID NUMBER(18),
    上级id NUMBER(18),
    编码 VARCHAR2(15),
    名称 VARCHAR2(40),
    简码 VARCHAR2(10),
    撤档时间 Date)
    TableSpace zl9BaseItem;
Create Table 收费项目目录(
    类别 VARCHAR2(1),
    分类ID NUMBER(18),
    ID NUMBER(18),
    编码 VARCHAR2(20),
    名称 VARCHAR2(80),
    规格 VARCHAR2(100),
    产地 Varchar2(200),
    计算单位 VARCHAR2(20),
    说明 VARCHAR2(500),
    项目特性 NUMBER(3),
    费用类型 VARCHAR2(20),
    服务对象 NUMBER(3),
    屏蔽费别 NUMBER(1),
    是否变价 NUMBER(1),
    加班加价 NUMBER(1),
    补充摘要 NUMBER(1),
    费用确认 Number(1),
    执行科室 NUMBER(3),
    标识主码 VARCHAR2(20),
    标识子码 VARCHAR2(1),
    备选码 VARCHAR2(20),
    最低限价 NUMBER(20,2),
    最高限价 NUMBER(20,2),
    建档时间 DATE,
    撤档时间 DATE,
    录入限量 Number(16,5),
    计算方式 Number(1),
    站点 Varchar2(3),
    启用原因 Varchar2(100),
    停用原因 Varchar2(100),
    病案费目 varchar2(30))
    TableSpace zl9BaseItem
    Cache Storage(Buffer_Pool Keep) 
;
CREATE TABLE 收费项目别名(
    收费细目ID NUMBER(18),
    名称 VARCHAR2(80),
    性质 NUMBER(3),
    简码 VARCHAR2(40),
    码类 NUMBER(3))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
CREATE TABLE 收费适用科室(
    项目ID NUMBER(18),
    科室id NUMBER(18))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
CREATE TABLE 收费执行科室(
    收费细目ID NUMBER(18),
    病人来源 NUMBER(3) DEFAULT 1,
    开单科室id NUMBER(18),
    执行科室id NUMBER(18))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 收费价目(
    ID NUMBER(18),
    原价id NUMBER(18),
    收费细目id NUMBER(18),
    原价 NUMBER(16,7),
    现价 NUMBER(16,7),
    缺省价格 NUMBER(16,7),
    收入项目id NUMBER(18),
    加班加价率 NUMBER(16,5),
    附术收费率 NUMBER(16,5),
    变动原因 NUMBER(3) default 1,
    调价说明 VARCHAR2(100),
    调价ID NUMBER(18),
    调价人 VARCHAR2(20),
    执行日期 Date,
    终止日期 Date,
    No VARCHAR2(8),
    序号 NUMBER(5),
    调价汇总号 Varchar2(10),
    价格等级 Varchar2(30))
    TABLESPACE zl9BaseItem
    Cache Storage(Buffer_Pool Keep)
;
Create Table 收费从属项目(
    主项ID NUMBER(18),
    从项ID NUMBER(18),
    固有从属 NUMBER(2) default 0,
    从项数次 NUMBER(16,5))
    TABLESPACE zl9BaseItem;
Create Table 收费特定项目(
    特定项目 VARCHAR2(20),
    收费细目id NUMBER(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE  成套项目分类(
	ID		NUMBER(18),
	上级ID	Number(18),
	编码		Varchar2(10),
	名称		Varchar2(50),
	简码		Varchar2(20))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
CREATE TABLE 成套收费项目(
	ID		NUMBER(18),
	分类ID	Number(18),
	编码		Varchar2(10),
	名称		Varchar2(100),
	拼音		Varchar2(20),
	五笔          VARCHAR2(20),
	范围		Number(2),
	人员ID	Number(18),
	备注         VARCHAR2(100))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
CREATE TABLE	成套项目使用科室(
	成套ID	NUMBER(18),
	科室ID	Number(18))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
CREATE TABLE	成套收费项目组合(
	成套ID		NUMBER(18),
	收费细目ID	Number(18),
	序号			Number(18),
	从属父号		Number(18),
	付数                  Number(3),
	数量			Number(16,7),
	单价			Number(16,7),
	执行科室ID	Number(18))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
create table 收费项目组成
(
  ID         NUMBER(18),
  收费项目ID NUMBER(18),
  名称       VARCHAR2(80),
  单价       NUMBER(15,6),
  规格       VARCHAR2(100),
  计算单位   VARCHAR2(20),
  数量       NUMBER(18),
  说明       VARCHAR2(500)
)
    TABLESPACE zl9BaseItem;
Create Table 记帐报警线(
    病区id NUMBER(18),
    适用病人 VARCHAR2(20),
    报警方法 NUMBER(1),
    报警值 NUMBER(16,5),
    报警标志1 Varchar2(30),
    报警标志2 Varchar2(30),
    报警标志3 Varchar2(30),
    催款下限 NUMBER(16,5),
    催款标准 NUMBER(16,5))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 自动计价项目(
    病区id NUMBER(18),
    计算标志 NUMBER(1),
    收费细目id NUMBER(18),
    启用日期 Date)
    TABLESPACE zl9BaseItem;
Create Table 单据操作控制(
	人员ID Number(18),
	单据 Number(1),
	时间限制 Number(3),
	他人单据 Number(1),
	金额上限 Number(20,5))
	TABLESPACE zl9BaseItem;
--医疗卡
Create Table 医疗卡类别(
    ID Number(18),
    编码 VarChar2(4),
    名称 Varchar2(50),
    短名 Varchar2(4),
    前缀文本 Varchar2(6),
    卡号长度 Number(5),
    缺省标志 Number(1),
    是否固定 Number(1) DEFAULT 0,
    是否严格控制 Number(1) DEFAULT 0,
    是否自制 Number(1) DEFAULT 0,
    是否存在帐户 Number(1),
    是否退现 number(1) DEFAULT 1,
    是否全退 Number(1) DEFAULT 0,
    部件 Varchar2(100),
    备注 Varchar2(100),
    特定项目 Varchar2(6),
    结算方式 Varchar2(20),
    卡号密文 VARCHAR2(10),
    是否重复使用 number(1) DEFAULT 0,
    是否启用 NUMBER(1) DEFAULT 0,
    密码长度 number(2) DEFAULT 10,
    密码长度限制 NUMBER(2) DEFAULT 0,
    密码规则 NUMBER(2),
    是否模糊查找 NUMBER(1) DEFAULT 0,
    密码输入限制 Number(1) DEFAULT 0,
    是否缺省密码 Number(1) DEFAULT 0,
    是否制卡 NUMBER(1) DEFAULT 0,
    是否发卡 NUMBER(1) DEFAULT 0,
    是否写卡 NUMBER(1),
    险类 Number(3),
    发卡性质 Number(2) DEFAULT 0,
    是否转帐及代扣 number(1) DEFAULT 0,
    读卡性质 VARCHAR2(4) default '1000',
    键盘控制方式 number(3) default 0,
    是否证件 NUMBER(1) default 0,
    是否持卡消费 Number(3) Default 1,
    发送调用接口 Number(3) Default 0,
    是否退款验卡 Number(1) Default 0,
    退号退卡检查控制 Number(1) Default 1,
    设备是否启用回车 Number(1) Default 0,
    发卡控制 number(2),
    是否缺省退现 Number(1),
    是否独立结算 number(2),
    缺省有效时间 Varchar2(6),
    卡号识别规则 NUMBER (2),
    是否支持扫码付 number(2))
    TABLESPACE zl9BaseItem 
;
Create Table 消费卡类别目录(
    编号 number(6),
    名称 varchar2(50),
    系统 number(2),
    结算方式 VARCHAR2(20),
    部件 varchar2(100),
    启用 number(2),
    自制卡 number(2),
    前缀文本 Varchar2(4),
    卡号长度 Number(6),
    是否密文 NUMBER(1) DEFAULT 0,
    是否退现 number(1) DEFAULT 1,
    是否全退 Number(1) DEFAULT 0,
    密码长度 number(2),
    密码长度限制 Number(2) DEFAULT 0,
    密码规则 NUMBER(2),
    读卡性质 VARCHAR2(4) default '1000',
    键盘控制方式 number(3) default 0,
    限制类别 Varchar2(500),
    是否严格控制 Number(1) Default 0,
    是否特定病人 Number(1) Default 0,
    是否允许换卡 Number(1) Default 0,
    是否允许补卡 Number(1) Default 0,
    是否允许余额退款 Number(1) Default 0,
    应用场合 Varchar2(3) Default '000')
    TABLESPACE zl9BaseItem 
;
CREATE TABLE 消费卡类型(
	编码 varchar2(2),
	名称 varchar2(20),
	缺省面额 number(16,5),
	缺省折扣 number(16,5),
	缺省标志 number(2)
	) TABLESPACE zl9BaseItem;
CREATE TABLE 常用发卡原因(
	编码 varchar2(4),
	名称 varchar2(50),
	简码 VARCHAR2(20),
	缺省标志 number(2)
	) TABLESPACE zl9BaseItem;
CREATE TABLE 医疗卡挂失方式(
	编码 Varchar2 (4),
	名称 Varchar2 (20),
	简码 varchar2(10),
	有效天数 number (5),
	缺省标志 Number (1))
	TABLESPACE zl9BaseItem;
--挂号
Create Table 号类(
	编码 VARCHAR2(2),
	名称 VARCHAR2(10),
	简码 VARCHAR2(4),
	缺省标志 NUMBER(1) default 0,
	说明 VARCHAR2(50))
	TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 预约方式(
    编码 VARCHAR2(2),
    名称 VARCHAR2(10),
    简码 VARCHAR2(6),
    缺省标志 NUMBER(1) default 0,
    预约天数 number(5))
    TABLESPACE zl9BaseItem 
;
Create Table 挂号安排(
    ID NUMBER(18),
    号类 Varchar2(10),
    号码 VARCHAR2(5) Not Null,
    科室id NUMBER(18),
    项目ID NUMBER(18),
    医生姓名 VARCHAR2(20),
    医生ID NUMBER(18),
    序号 NUMBER(18),
    周日 VARCHAR2(4),
    周一 VARCHAR2(4),
    周二 VARCHAR2(4),
    周三 VARCHAR2(4),
    周四 VARCHAR2(4),
    周五 VARCHAR2(4),
    周六 VARCHAR2(4),
    病案必须 Number(1),
    分诊方式 Number(1),
    序号控制 Number(1) Default 0,
    开始时间 Date,
    终止时间 Date,
    停用日期 Date,
    执行时间 Date,
    执行计划ID Number(18),
    默认时段间隔 Number(5),
    是否删除 Number(1) Default 0,
    预约天数 Number(5))
    TABLESPACE zl9BaseItem
    Cache Storage(Buffer_Pool Keep)
;
	
Create Table 挂号安排时段( 
	 安排ID Number(18),
	 序号 Number(18),
    星期 VARCHAR2(10),
	 开始时间 Date,
	 结束时间 Date,
	 限制数量 Number(18),
	 是否预约 Number(1) DEFAULT 0)
	 Tablespace zl9BaseItem;
   
Create Table 挂号安排限制 (
	安排ID Number(18),
	限制项目 Varchar2(10),
	限号数 Number(5),
	限约数 Number(5))
	TableSpace zl9BaseItem;
Create Table 挂号序号状态(
    号码 Varchar2(5),
    日期 Date,
    序号 Number(5),
    状态 Number(1),
    操作员姓名 Varchar2(20),
    预约 Number(1) Default(0),
    备注 VARCHAR2(100),
    机器名 varchar2(200),
    登记时间 date)
    TABLESPACE zl9Patient 
	initrans 20	
;
Create Table 挂号安排诊室(
    号表ID NUMBER(18),
    门诊诊室 VARCHAR2(20),
	当前分配 Number(1))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 挂号安排计划(
    ID NUMBER(18),
    安排ID NUMBER (18),
    项目ID NUMBER(18),
    号码 VARCHAR2(5),
    生效时间 DATE,
    失效时间 Date DEFAULT To_date('3000-01-01','yyyy-mm-dd'),
    周日 VARCHAR2(4),
    周一 VARCHAR2(4),
    周二 VARCHAR2(4),
    周三 VARCHAR2(4),
    周四 VARCHAR2(4),
    周五 VARCHAR2(4),
    周六 VARCHAR2(4),
    分诊方式 NUMBER(1),
    序号控制 NUMBER(1),
    安排人 VARCHAR2(20),
    安排时间 DATE,
    审核人 VARCHAR2(20),
    审核时间 DATE,
    医生姓名 Varchar2(20),
    医生ID Number(18),
    实际生效 DATE DEFAULT To_date('3000-01-01','yyyy-mm-dd'),
    默认时段间隔 Number(5),
    上次计划ID number(18))
    TABLESPACE zl9BaseItem
    Cache Storage(Buffer_Pool Keep)
;
	
Create  Table 	挂号计划限制(
	计划ID Number(18),
	限制项目 Varchar2(10),
	限号数 Number(5),
	限约数 Number(5))
	Tablespace zl9BaseItem;
 
Create Table 挂号计划时段(
	计划ID Number(18),
	序号 Number(18),
  星期 VARCHAR2(10),
	开始时间 Date,
	结束时间 Date,
	限制数量 Number(18),
	是否预约 Number(1))
	Tablespace zl9BaseItem;
 
CREATE TABLE 挂号安排停用状态(
	安排ID	number(18),
	序号         number(18),
	开始停止时间  Date,
	结束停止时间   Date,
	制订人	varchar2(20),  
	制订日期     date,
	备注           varchar2(100))
    TABLESPACE zl9BaseItem;
CREATE TABLE 安排停用原因(
	编码 VARCHAR2(5),
	名称 VARCHAR2(50),
	简码 VARCHAR2(10),
	缺省标志 number (1))
	TABLESPACE zl9BaseItem;
Create Table 挂号计划诊室(
	计划ID	NUMBER(18),
	门诊诊室 VARCHAR2(20))
	TABLESPACE zl9BaseItem	
	Cache Storage(Buffer_Pool Keep);
 
Create Table 挂号合作单位(
    编码 Varchar2 (4),
    名称 Varchar2 (50),
    简码 varchar2(10),
    缺省标志 number (1),
    锁号时间 number(18))
    TABLESPACE zl9BaseItem
;
CREATE TABLE 合作单位安排控制(
	合作单位 Varchar2 (50),
	安排ID number(18),
	限制项目 varchar2(10),
	序号 Number(18),
	数量 number(10))
	TABLESPACE zl9BaseItem;
Create Table 合作单位计划控制(
	合作单位 VARCHAR2(50) ,
	计划ID   NUMBER(18),
	限制项目 VARCHAR2(10),
	序号     NUMBER(18),
	数量     NUMBER(10))
	Tablespace zl9BaseItem; 
Create Table 收费记帐单(
    ID NUMBER(18),
    编号 VARCHAR2(6),
    名称 VARCHAR2(50),
    收费项目数 NUMBER(18),
    适用范围 VARCHAR2(4),
    宽度 NUMBER(16),
    高度 NUMBER(16),
	背景色	NUMBER(18),
	字体	VARCHAR2(50))
    TABLESPACE zl9BaseItem;
Create Table 收费记帐单定义(
    记帐ID NUMBER(18),
    对应字段 VARCHAR2(50),
    序号 NUMBER(18),
    类型 VARCHAR2(20),
    定义值 VARCHAR2(30),
    顺序号 NUMBER(5),
	左边	NUMBER(16),
	顶边	NUMBER(16),
	宽度	NUMBER(16),
	高度	NUMBER(16),
	字体	VARCHAR2(50),
	前景色	NUMBER(18),
	背景色	NUMBER(18),
	是否显示	NUMBER(1),
	外形	NUMBER(1),
	边框线	NUMBER(1),
	透明	NUMBER(1))
    TABLESPACE zl9BaseItem;
Create Table 法定假日表(
    年份 Number(18),
    节日名称 varchar2(50),
    性质 Number(18),
    开始日期 Date,
    终止日期 Date,
    备注 varchar2(1000),
    允许挂号日期 varchar2(500),
    允许预约日期 varchar2(500))
    TABLESPACE zl9BaseItem 
;
CREATE TABLE 门诊诊室适用科室 (
	诊室ID number(18),
	科室ID number(18),
	缺省标志 number(2)) 
TABLESPACE zl9BaseItem ;
Create Table 临床出诊号源(
    ID number(18),
    号类 varchar2(10),
    号码 varchar2(5),
    科室id number(18),
    项目ID number(18),
    医生id number(18),
    医生姓名 varchar2(50),
    是否建病案 number(2) default 0,
    预约天数 number(3),
    出诊频次 number(3),
    假日控制状态 number(2),
    是否假日换休 number(2) default 0,
    是否临床排班 number(2) default 0,
    适用性别 varchar2(4),
    适用年龄段 varchar2(100),
    是否删除 number(2) default 0,
    建档时间 Date,
    撤档时间 Date,
    是否固定排班 Number(2),
    是否月排班 Number(2),
    是否周排班 Number(2))
    TABLESPACE zl9BaseItem 
;
Create Table 临床出诊号源限制(
   ID number(18),
   号源ID number(18),
   上班时段 varchar2(10),
   限号数 number(10),
   限约数 number(10),
   是否序号控制 number(2) default 0,
   是否分时段  NUMBER(2),
   预约控制 number(2),
   是否独占 number(2) default 0,   
   分诊方式 number(3),
   诊室ID number(18))
TABLESPACE zl9BaseItem ;
Create Table 临床出诊号源诊室(
   限制ID number(18),
   诊室ID number(18))
TABLESPACE zl9BaseItem ;
Create Table 临床出诊号源时段(
   限制ID number(18),
   序号 number(18),
   开始时间 Date,
   终止时间 Date,
   限制数量 number(10),
   是否预约 number(2))
TABLESPACE zl9BaseItem;
Create Table 临床出诊号源控制(
   限制ID number(18),
   类型 number(2),
   性质 number(2),
   名称 varchar2(50),
   序号 number(18),
   控制方式 number(2),
   数量 number(16,5))
TABLESPACE zl9BaseItem ;
Create Table 临床出诊表(
    ID number(18),
    排班方式 number(18),
    出诊表名 varchar2(50),
    年份 number(4),
    月份 number(2),
    周数 number(2),
    应用范围 number(2),
    科室ID number(18),
    备注 varchar2(100),
    创建人 varchar2(50),
    创建时间 Date,
    模板类型 Number(2),
    站点 Varchar2(3),
    关联ID Number(18))
    TABLESPACE zl9BaseItem 
;
Create Table 临床出诊安排(
    ID number(18),
    出诊ID number(18),
    号源ID number(18),
    项目ID number(18),
    医生id number(18),
    医生姓名 varchar2(50),
    排班规则 number(2),
    是否周六出诊 number(2),
    是否周日出诊 number(2),
    开始时间 Date,
    终止时间 Date,
    操作员姓名 varchar2(50),
    登记时间 Date,
    发布人 VARCHAR2(50),
    发布时间 Date,
    是否临时安排 number(2))
    TABLESPACE zl9BaseItem 
;
Create Table 临床出诊限制(
   ID     number(18),
   安排ID number(18),
   限制项目 varchar2(20),
   上班时段 varchar2(10),
   限号数 number(10),
   限约数 number(10),
   是否序号控制 number(2),
   是否分时段 NUMBER(2),
   预约控制 number(2),
   分诊方式 number(2),
   诊室ID number(18),
   是否独占 number(2))
TABLESPACE zl9BaseItem ;
Create Table 临床出诊诊室(
   限制ID number(18),
   诊室ID number(18))
TABLESPACE zl9BaseItem ;
Create Table 临床出诊时段(
   限制ID number(18),
   序号 number(18),
   开始时间 Date,
   终止时间 Date,
   限制数量 number(10),
   是否预约 number(2))
TABLESPACE zl9BaseItem ;
Create Table 临床出诊挂号控制(
   限制ID number(18),
   类型 number(2),
   性质 number(2),
   名称 varchar2(50),
   序号 number(18),
   控制方式 number(2),
   数量 number(16,5))
TABLESPACE zl9BaseItem ;
Create Table 临床出诊记录(
    ID number(18),
    安排ID number(18),
    号源ID number(18),
    出诊日期 Date,
    上班时段 varchar2(10),
    开始时间 Date,
    终止时间 Date,
    停诊开始时间 Date,
    停诊终止时间 Date,
    停诊原因 varchar2(50),
    缺省预约时间 Date,
    提前挂号时间 Date,
    限号数 number(10),
    已挂数 number(10),
    限约数 number(10),
    已约数 number(10),
    其中已接收 number(10),
    是否序号控制 number(2) default 0,
    是否分时段 number(2) default 0,
    预约控制 number(2),
    是否独占 number(2),
    项目ID number(18),
    科室ID number(18),
    医生id number(18),
    医生姓名 varchar2(50),
    替诊开始时间 Date,
    替诊终止时间 Date,
    替诊医生id number(18),
    替诊医生姓名 varchar2(50),
    分诊方式 number(2),
    诊室ID Number(18),
    是否锁定 number(2) default 0,
    是否临时出诊 number(2) default 0,
    登记人 varchar2(50),
    登记时间 Date,
    是否发布 number(2) default 0,
    相关ID Number(18))
    TABLESPACE zl9BaseItem
;
Create Table 临床出诊诊室记录(
   记录ID number(18),
   诊室ID Number(18),
   当前分配 number(1) default 0)
TABLESPACE zl9BaseItem ;
Create Table 临床出诊序号控制(
    记录ID number(18),
    序号 number(18),
    预约顺序号 number(18),
    开始时间 Date,
    终止时间 Date,
    数量 number(10),
    是否预约 number(2),
    挂号状态 number(2),
    锁号时间 Date,
    类型 number(2),
    名称 varchar2(50),
    操作员姓名 varchar2(50),
    工作站IP varchar2(20),
    工作站名称 varchar2(200),
    备注 varchar2(100),
    是否停诊 Number(2) Default 0)
    TABLESPACE zl9BaseItem 
	initrans 20
;
Create Table 临床出诊挂号控制记录(
   记录ID number(18),
   类型 number(2),
   性质 number(2),
   名称 varchar2(50),
   序号 number(18),
   控制方式 number(2),
   数量 number(16,5))
TABLESPACE zl9BaseItem ;
Create Table 临床出诊变动记录(
   ID number(18),
   记录ID number(18),
   变动类型 number(2),
   原预约控制 number(2),
   现预约控制 number(2),
   原数量 number(10),
   现数量 number(10),
   原分诊方式 number(2),
   原门诊诊室 varchar2(20),
   原诊室ID number(18),
   现分诊方式 number(2),
   现门诊诊室 varchar2(20),
   现诊室ID number(18),
   操作员姓名 varchar2(50),
   登记时间 Date)
TABLESPACE zl9BaseItem ;
Create Table 临床出诊变动明细(
   变动ID number(18),
   变动性质 number(2),
   类型 number(2),
   名称 varchar2(50),
   序号 number(18),
   控制方式 number(2),
   数量 number(10),
   诊室ID number(18),
   门诊诊室 varchar2(20))
TABLESPACE zl9BaseItem ;
Create Table 临床出诊停诊记录(
    ID number(18),
    记录ID number(18),
    开始时间 Date,
    终止时间 Date,
    停诊原因 varchar2(50),
    替诊医生ID number(18),
    替诊医生姓名 varchar2(50),
    申请人 varchar2(50),
    申请时间 Date,
    审批人 varchar2(50),
    审批时间 Date,
    取消人 varchar2(50),
    取消时间 Date,
    登记人 Varchar2(50),
    失效时间 Date,
    停诊号码 Varchar2(600))
    TABLESPACE zl9BaseItem 
;
Create Table 常用停诊原因(
   编码 varchar2(5),
   名称 varchar2(50),
   简码 varchar2(20),
   缺省标志 number(1) default 0)
TABLESPACE zl9BaseItem ;
----------------------------------------------------------------------------
--[[5.药品卫材基础]]
----------------------------------------------------------------------------
Create Table 药品三方事务接口(
    序号 NUMBER(4),
    名称 VARCHAR2(100),
    类型 NUMBER(1),
    设置 VARCHAR2(200),
    是否启动 NUMBER(1)
    )
    TABLESPACE zl9BaseItem;

Create Table 配送单号对照(
    id NUMBER(18),
    库房id NUMBER(18),
    单据 NUMBER(2),
    NO varchar2(8),
    配送单号 varchar2(20))
    TABLESPACE zl9medLst;

Create Table 中选药品对照(
    序号 NUMBER(4),
    中选药品id NUMBER(18),
    对照药品id NUMBER(18))
    TABLESPACE zl9BaseItem;

CREATE TABLE 药品存储库房(
    药品id NUMBER(18),
    库房id NUMBER(18),
    科室id NUMBER(18))
    Tablespace ZL9BASEITEM;
Create table 配置收费方案( 
  序号  NUMBER(18) not null,
  配药类型  varchar2(50),
  项目id  NUMBER(18),
  收费项目 varchar2(100),
  诊疗ID NUMBER(18))
  tablespace ZL9BASEITEM;
Create Table 药品包装单位(
  编码 VARCHAR2(3),
  名称 VARCHAR2(10),
  简码 VARCHAR2(5)
)TABLESPACE zl9BaseItem;
Create Table 药品使用说明项目(
  编码 VARCHAR2(2),
  名称 VARCHAR2(20),
  简码 VARCHAR2(20)
)TABLESPACE zl9BaseItem;
create table 入库验收结论
(
  编码   VARCHAR2(3),
  名称   VARCHAR2(100),
  缺省标志 NUMBER(1) default 0
)
tablespace ZL9BASEITEM;
Create Table 输液不配置药品(
    药品id number(18),
    名称 VARCHAR2(50))
    TABLESPACE zl9BaseItem;
Create Table 输液自备药清单
(
  序号     Number(3) ,
  药品id   Number(18) Not Null,
  是否检查库存 Number(1)
)
Tablespace zl9BaseItem;
Create Table 输液优先打印药品(
    药品id number(18),
    名称 VARCHAR2(50))
    TABLESPACE zl9BaseItem;  
Create Table 发药窗口(
    编码 VARCHAR2(1),
    名称 Varchar2(50),
    上班否 NUMBER(2),
    专家 NUMBER(1),
    药房id NUMBER(18),
    当前分配 NUMBER(1),
    叫号窗口 varchar2(50))
    TABLESPACE zl9BaseItem 
;
CREATE TABLE 药品外调单位(
	编码 VARCHAR2(2),
	名称 VARCHAR2(50),
	简码 VARCHAR2(10))
	TABLESPACE zl9BaseItem;
CREATE TABLE 药品外销单位(
	编码 VARCHAR2(2),
	名称 VARCHAR2(50),
	简码 VARCHAR2(10))
	TABLESPACE zl9BaseItem;
Create Table 药品库房货位(
    ID Number(18),
    编码 VARCHAR2(5),
    名称 VARCHAR2(50),
    简码 VARCHAR2(10),
    库房id number(18),
    备注 varchar2(100),
    上级id Number(18),
    末级 Number(1) Default 1)
    TABLESPACE zl9BaseItem 
;
Create Table 药品货位对照(
     库房ID number(18),
     药品ID number(18),
     货位ID number(18))
     TABLESPACE zl9BaseItem;
Create Table 药价管理级别(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(4))
    TABLESPACE zl9BaseItem;
Create Table 药品毒理分类(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(4))
    TABLESPACE zl9BaseItem;
Create Table 药品货源情况(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(4))
    TABLESPACE zl9BaseItem;
Create Table 药品剂型(
    编码 VARCHAR2(3),
    名称 VARCHAR2(20) Not Null,
    简码 VARCHAR2(10),
    标记码 VARCHAR2(1))
    TABLESPACE zl9BaseItem;
Create Table 发药类型(
    编码 VARCHAR2(3),
    名称 VARCHAR2(20) Not Null,
    简码 VARCHAR2(10))
    TABLESPACE zl9BaseItem;
Create Table 药品外观(
    编码 VARCHAR2(3),
    名称 VARCHAR2(20),
    简码 VARCHAR2(10),
    缺省标志 number(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 药品价值分类(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(4))
    TABLESPACE zl9BaseItem;
Create Table 药品存储温度(
    编码 VARCHAR2(2),
    名称 VARCHAR2(20),
    简码 VARCHAR2(4))
    TABLESPACE zl9BaseItem;
Create Table 药品来源分类(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(4))
    TABLESPACE zl9BaseItem;
create table 药品说明书
(
  ID    VARCHAR2(100),
  通用名称  VARCHAR2(4000),
  商品名   VARCHAR2(4000),
  汉语拼音  VARCHAR2(4000),
  英文名称  VARCHAR2(4000),
  药物规格  VARCHAR2(4000),
  药物剂型  VARCHAR2(4000),
  主要成分  VARCHAR2(4000),
  生产企业  VARCHAR2(4000),
  批准文号  VARCHAR2(4000),
  化学名称  CLOB,
  性状    CLOB,
  药理毒理  CLOB,
  药代动力学 CLOB,
  适应症   CLOB,
  用法用量  CLOB,
  不良反应  CLOB,
  禁忌症   CLOB,
  注意事项  CLOB,
  孕妇用药  CLOB,
  儿童用药  CLOB,
  老年人用药 CLOB,
  相互作用  CLOB,
  药物过量  CLOB,
  贮藏条件  CLOB
)
tablespace ZL9BASEITEM;
create table 基本药物说明
(
  编码 VARCHAR2(2),
  名称 VARCHAR2(30),
  简码 VARCHAR2(10)
)
tablespace ZL9BASEITEM;
Create Table 药品用药梯次(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(4))
    TABLESPACE zl9BaseItem;
Create Table 药品生产商(
    编码 VARCHAR2(10),
    名称 Varchar2(200),
    简码 VARCHAR2(10),
    站点 Varchar2(3))
    TABLESPACE zl9BaseItem 
;
Create Table 药品入出类别(
    ID NUMBER(18),
    编码 VARCHAR2(3),
    名称 VARCHAR2(20),
    系数 NUMBER(2))
    TABLESPACE zl9BaseItem;
Create Table 药品单据分类(
    编码 NUMBER(2),
    名称 VARCHAR2(16),
    性质 NUMBER(1),
    说明 VARCHAR2(200))
    TABLESPACE zl9BaseItem;
Create Table 药品单据性质(
    单据 NUMBER(2),
    类别id NUMBER(18))
    TABLESPACE zl9BaseItem;
Create Table 毁损发生原因(
    编码 VARCHAR2(1),
    名称 VARCHAR2(20),
    简码 VARCHAR2(4))
    TABLESPACE zl9BaseItem;
Create Table 毁损解决办法(
    编码 VARCHAR2(1),
    名称 VARCHAR2(20),
    简码 VARCHAR2(4))
    TABLESPACE zl9BaseItem;
Create Table 药品流向控制(
    所在库房ID NUMBER(18),
    对方库房ID NUMBER(18),
    流向 NUMBER(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 药品领用控制(
	领用部门ID NUMBER(18),
	对方库房ID NUMBER(18))
	TABLESPACE zl9BaseItem;
Create Table 药品库房单位(
    库房id NUMBER(18),
	适用范围 NUMBER(1),	--1-药库;2-门诊药房;3-住院药房;4-其它(制剂室)
    性质 NUMBER(1))  --1-售价单位,2-门诊单位,3-住院单位,4-药库单位
    TABLESPACE zl9BaseItem;
CREATE TABLE 药房配药控制(
    药房id NUMBER(18),
	门诊 NUMBER(1),		--1-门诊;2-住院
    配药 NUMBER(1),		--1-配药;其它不需配药
	自动发药天数 Number(3),
	配药确认 Number(1))
    TABLESPACE zl9BaseItem;
Create Table 供应商(
    ID NUMBER(18),
    上级id NUMBER(18),
    编码 VARCHAR2(8),
    名称 VARCHAR2(80),
    简码 VARCHAR2(10),
    末级 NUMBER(1) default 0,
    许可证号 VARCHAR2(30),
    许可证效期 DATE,
    执照号 VARCHAR2(30),
    执照效期 DATE,
    授权号 VARCHAR2(30),
    授权期 DATE,
    税务登记号 VARCHAR2(30),
    地址 VARCHAR2(50),
    电话 VARCHAR2(16),
    开户银行 VARCHAR2(50),
    帐号 varchar2(50),
    联系人 VARCHAR2(20),
    建档时间 Date,
    撤档时间 Date,
    类型 varchar2(10),
    信用期 number(6),
    信用额 number(18,5),
    销售委托人 varchar2(20),
    销售委托日期 date,
    质量认证号 varchar2(20),
    质量认证日期 date,
    药监局备案号 varchar2(20),
    药监局备案日期 date,
    站点 Varchar2(3),
    首营品种 Varchar2(200),
    备注 Varchar2(200))
    TABLESPACE zl9BaseItem 
;
Create Table 供应商照片(
    供应商ID NUMBER(18),
    许可证号照片 BLOB,
    执照号照片 BLOB,
    授权号照片 Blob)
    TABLESPACE zl9BaseItem;
Create Table 药品生产商对照 (
    药品id number(18),
    厂家名称 Varchar2(200),
    批准文号 VARCHAR2(40))
    tablespace ZL9BASEITEM 
;
Create Table 药品特性(
    药名ID NUMBER(18),
    药品剂型 VARCHAR2(20),
    毒理分类 VARCHAR2(10),
    货源情况 VARCHAR2(10),
    价值分类 VARCHAR2(10),
    用药梯次 VARCHAR2(10),
    急救药否 NUMBER(1),
    是否新药 NUMBER(1),
    是否皮试 NUMBER(1),
    是否原料 NUMBER(1),
    处方限量 NUMBER(16,5),
    处方职务 VARCHAR2(2),
    药品类型 NUMBER(1),
    品种医嘱 NUMBER(1),
    抗生素 number(1) DEFAULT 0,
    临床自管药 number(1),
    ATCCODE varchar2(50),
    是否肿瘤药 number(1),
    溶媒 number(1),
    是否原研药 Number(1),
    是否专利药 Number(1),
    是否单独定价 Number(1),
    是否辅助用药 number(1),
    严格控制用法用量 number(1))
    TableSpace zl9BaseItem
    Cache Storage(Buffer_Pool Keep) 
;
CREATE TABLE 药品规格(
    药名ID NUMBER(18),
    药品id NUMBER(18),
    剂量系数 NUMBER(16,5),
    门诊单位 VARCHAR2(8),
    门诊包装 NUMBER(16,5),
    住院单位 VARCHAR2(8),
    住院包装 NUMBER(16,5),
    药库单位 VARCHAR2(8),
    药库包装 NUMBER(16,5),
    最大效期 NUMBER(5),
    药品来源 VARCHAR2(10),
    协定药品 NUMBER(1),
    自制药品 NUMBER(1),
    批准文号 VARCHAR2(40),
    注册商标 VARCHAR2(50),
    标识码 VARCHAR2(29),
    药价级别 VARCHAR2(10),
    指导批发价 NUMBER(16,7),
    指导零售价 NUMBER(16,7),
    指导差价率 NUMBER(16,5),
    扣率 NUMBER(16,5),
    住院可否分零 NUMBER(3),
    动态分零 NUMBER(1),
    药库分批 NUMBER(1),
    药房分批 NUMBER(1),
    招标药品 NUMBER(1),
    差价让利比 NUMBER(16,5),
    GMP认证 NUMBER(1),
    成本价 number(16,7),
    管理费比例 NUMBER(16,5),
    申领单位 NUMBER(1),
    申领阀值 NUMBER(16,5),
    合同单位ID NUMBER(18),
    上次供应商ID NUMBER(18),
    上次产地     VARCHAR2(60),
    上次批号     VARCHAR2(20),
    上次生产日期 DATE,
    上次批准文号 VARCHAR2(40),
    发药类型  VARCHAR2(20),
    容量 NUMBER(16,5),
    增值税率 NUMBER(16,5),
    基本药物 varchar2(30),
    中药形态 Number(1) Default 0,  -- 0:散装;  1:中药饮片;  2:免煎剂
    是否常备 Number(1),
    门诊可否分零   NUMBER(3),
    DDD值 number(16,5),
    上次售价 number(16,7),
    高危药品 number(1),
    送货单位 varchar2(8),
    送货包装 number(16,5),
    加成率 number(16,5),
    图片 Blob,
    使用说明 Clob,
    是否摆药 number(1) default 1,
    是否零差价管理 Number(1),
	本位码 varchar2(20),
	原产地 varchar2(60),
	严格控制用法用量 number(1),
	是否易致跌倒 number(1),
	允许院外取药 number(1),
	是否带量采购 number(1),
	带量供应商ID number(18)
    ) 
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 药品规格扩展项目(
  编码 Number(3),
  名称 Varchar2(20))
  TABLESPACE zl9BaseItem
 Cache Storage(Buffer_Pool Keep);
Create Table 药品规格扩展信息(
  药品id Number(18),
  项目 Varchar2(20),
  内容 Varchar2(1000))
  TABLESPACE zl9BaseItem
 Cache Storage(Buffer_Pool Keep);
Create Table 药品加成方案(
    序号 NUMBER(18),
    最低价 NUMBER(16,5),
    最高价 NUMBER(16,5),
    加成率 NUMBER(16,5),
    差价额 NUMBER(16,5),
    说明 VARCHAR2(50),
    类型 number(1))
    tablespace ZL9BASEITEM;
Create Table 协定药品对照(
    药品ID NUMBER(18),
    协定药品ID NUMBER(18),
    分子 NUMBER(16,5),
    分母 NUMBER(16,5))
    TABLESPACE zl9BaseItem;
Create Table 自制药品构成(
    自制药品ID NUMBER(18),
    原料药品ID NUMBER(18),
    分子 NUMBER(16,5),
    分母 NUMBER(16,5))
    TABLESPACE zl9BaseItem;
Create Table 药品储备限额(
    库房id NUMBER(18),
    药品id NUMBER(18),
    上限 NUMBER(18,5),
    下限 NUMBER(18,5),
    盘点属性 VARCHAR2(4),
    库房货位 VARCHAR2(50),
    领用标志 Number(1) default 1)
    TABLESPACE zl9BaseItem;
Create Table 药品中标单位(
		药品id NUMBER(18),
		单位ID NUMBER(18),
		建档时间 Date,
		撤档时间 Date,
		中标序号 Varchar2(50))
    TABLESPACE zl9BaseItem;
CREATE TABLE 输液配药类型(
  编码 varchar2(4),
  名称 varchar2(50),
  简码 VARCHAR2(20)
  ) TABLESPACE zl9BaseItem;
  
CREATE TABLE 科室容量设置(
  科室id varchar2(20),
  科室名称 varchar2(20),
  配药批次 varchar2(20),
  容量 number(18),
  配置中心ID number(18)
  ) TABLESPACE zl9BaseItem;
CREATE TABLE 输液药品优先级(
  科室id varchar2(1000),
  科室名称 varchar2(2000),
  配药类型 varchar2(200),
  频次 VARCHAR2(200),
  有效 number(1),
  优先级 number(3)
  ) TABLESPACE zl9BaseItem;
Create Table 配药工作批次(
    批次 NUMBER(2),
    配药时间 Varchar2(20),
    给药时间 Varchar2(20),
    打包 Number(1) Default 0,
    启用 Number(1) Default 1,
    颜色 NUMBER(18),
    配置中心ID number(18),
    药品类型 varchar2(20))
    TABLESPACE zl9BaseItem 
;
Create Table 输液药品属性(
    药品ID NUMBER(18),
    存储温度 VARCHAR2(20),
    存储条件 NUMBER(1),
    配药类型 VARCHAR2(30),
    是否不予配置 NUMBER(1),
    输液注意事项 varchar2(200))
    TABLESPACE zl9BaseItem
;
Create Table 材料货源情况(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(4))
    TABLESPACE zl9BaseItem;
Create Table 材料价值分类(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(4))
    TABLESPACE zl9BaseItem;
Create Table 材料来源分类(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(4))
    TABLESPACE zl9BaseItem;
Create Table 材料材质分类(
    编码 VARCHAR2(4),
    名称 VARCHAR2(30),
    简码 VARCHAR2(10))
    TABLESPACE zl9BaseItem;
Create Table 材料存储条件(
    编码 VARCHAR2(4),
    名称 VARCHAR2(30),
    简码 VARCHAR2(10))
    TABLESPACE zl9BaseItem;
Create Table 材料生产商(
    编码 VARCHAR2(10),
    名称 VARCHAR2(60),
    简码 VARCHAR2(10),
    生产企业许可证 varchar2(40),
    生产企业许可证效期 Date,
    站点 Varchar2(3),
    经营许可证 varchar2(40),
    经营许可证效期 date,
    企业法人执照 varchar2(40),
    企业法人执照效期 date)
    TABLESPACE zl9BaseItem
;
Create Table 材料流向控制(
    所在库房ID NUMBER(18),
    对方库房ID NUMBER(18),
    流向 NUMBER(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 材料库房货位(
	编码 VARCHAR2(5),
	名称 VARCHAR2(50),
	简码 VARCHAR2(10))
	TABLESPACE zl9BaseItem;
Create Table 材料储备限额(
    库房id NUMBER(18),
    材料id NUMBER(18),
    上限 NUMBER(18,5),
    下限 NUMBER(18,5),
    盘点属性 VARCHAR2(4),
    库房货位 VARCHAR2(50))
    TABLESPACE zl9BaseItem;
Create Table 自制材料构成(
    自制材料ID NUMBER(18),
    原料材料ID NUMBER(18),
    分子 NUMBER(16,5),
    分母 NUMBER(16,5))
    TABLESPACE zl9BaseItem;
Create Table 材料特性(
    材料ID number(18),
    诊疗ID number(18),
    最大效期 number(5),
    灭菌效期 number(5),
    无菌性材料 number(1) default 0,
    一次性材料 number(1) default 0,
    原材料 number(1) DEFAULT 0,
    自制材料 number(1) default 0,
    货源情况 varchar2(10),
    材质分类 varchar2(30),
    存储条件 varchar2(30),
    许可证号 varchar2(50),
    许可证有效期 DATE,
    批准文号 VARCHAR2(40),
    注册商标 VARCHAR2(50),
    注册证号 Varchar2(50),
    包装单位 varchar2(8),
    换算系数 number(16,5),
    指导批发价 number(16,7),
    指导零售价 number(16,7),
    指导差价率 number(16,5),
    成本价 number(16,7),
    差价让利比 NUMBER(16,5),
    扣率 number(16,5),
    库房分批 number(1) default 0,
    在用分批 number(1) default 0,
    材料来源 varchar2(10),
    剂量单位 Varchar2(20),
    剂量系数 Number(16,5),
    招标材料 NUMBER(1),
    跟踪在用 number(1),
    跟踪病人 NUMBER(1) DEFAULT 0,
    核算材料 NUMBER(1) DEFAULT 0,
    增值税率 NUMBER(16,5),
    高值材料 number(1),
    是否条码管理 Number(1),
    上次售价 number(16,7),
    器械包卫材单件 number(1),
    上次供应商id number(18),
    上次产地 varchar2(60),
    注册证有效期 date,
    是否植入耗材 number(1),
    加成率 Number(16,5),
    是否分零 Number(1) Default 0,
    型号 varchar2(100))
    TableSpace zl9BaseItem
    Cache Storage(Buffer_Pool Keep) 
;
Create Table 材料中标单位(
    材料id NUMBER(18),
    单位ID NUMBER(18),
    成本价 NUMBER(16,5),
    中标序号 Varchar2(50))
    TABLESPACE zl9BaseItem;
CREATE TABLE 材料外销单位(
	编码 VARCHAR2(2),
	名称 VARCHAR2(50),
	简码 VARCHAR2(10))
	TABLESPACE zl9BaseItem;
Create Table 材料领用用途(
    编码 VARCHAR2(6),
    名称 varchar2(50),
    简码 varchar2(10),
    缺省标志 number(2) DEFAULT 0)
    TABLESPACE zl9BaseItem;
CREATE TABLE 材料加成方案(
	序号	number(18),
	最低价	number(16,5),
	最高价	number(16,5),
	加成率	number(16,5),
	计算方法 number(1),
	限价 number(16,5),
	说明 varchar2(50))
    TableSpace zl9BaseItem;
Create Table 药品卫材精度(
    性质        Number(1),
    类别	Number(1),
    内容	Number(1),
    单位	Number(1),
    精度	Number(1))
    TABLESPACE zl9BaseItem;
Create Table 药品出库检查(
	库房ID Number(18),
	检查方式 Number(1))
	TABLESPACE zl9BaseItem;
Create Table 材料出库检查(
	库房ID Number(18),
	检查方式 Number(1))
	TABLESPACE zl9BaseItem;
Create Table 虚拟库房对照(
  科室ID Number(18),
  库房ID Number(18),
  虚拟库房ID Number(18))
  TABLESPACE zl9BaseItem;
----------------------------------------------------------------------------
--[[6.临床基础]]
----------------------------------------------------------------------------
Create Table 医生交接班记录(
    记录id Number(18),
    科室id Number(18),
    交班医生 Varchar2(20),
    交班班次 Varchar2(10),
    交班开始时间 Date,
    交班结束时间 Date,
    接班医生 Varchar2(20),
    接班班次 Varchar2(10),
    接班开始时间 Date,
    接班结束时间 Date,
    交班状态 Number(1),
    接班状态 Number(1),
    记录人 Varchar2(20),
    完成时间 Date,
    审阅人 Varchar2(20),
    审阅时间 Date,
    审阅说明 Varchar2(100)
 )Tablespace ZL9BASEITEM Initrans 20;

Create Table 医生交接班内容(
    内容id Number(18),
    记录id Number(18),
    序号 Number(3),
    病人类型 Varchar2(100),
    病人id Number(18),
    主页id Number(5),
    姓名 Varchar2(100),
    性别 Varchar2(4),
    年龄 Varchar2(20),
    床号 Varchar2(10),
    标识号 Number(18),
    入院时间 Date,
    入院方式 Varchar2(8),
    交班描述 Varchar2(2000)
)Tablespace ZL9BASEITEM Initrans 20;

Create Table 医生交接班详情(
    内容id Number(18), 
    序号 Number(3), 
    项目 Varchar2(50), 
    内容 Varchar2(500)
)Tablespace ZL9BASEITEM Initrans 20;

Create Table 医生交接班汇总(
    记录id Number(18), 
    序号 Number(2), 
    项目 Varchar2(20), 
    数量 Number(5)
)Tablespace ZL9BASEITEM;

Create Table 医生值班班次(
    科室ID Number(18),
    值班班次 Varchar2(10),
    开始时间 Date,
    结束时间 Date
)Tablespace ZL9BASEITEM;

Create Table 医生交接班签名(
    签名ID Number(18),
    记录ID Number(18),
    签名类型 Number(1),
    签名人 Varchar2(20),
    签名时间 Date,
    证书ID Number(18),
    签名信息 Varchar2(4000),
    时间戳信息 Varchar2(4000),
    时间戳 Date
)Tablespace ZL9BASEITEM Initrans 20;

Create Table 医生交接班病人项目(
    病人简称 Varchar2(10),
    项目名称 Varchar2(20),
    序号 Number(3),
    项目类别 Varchar2(1),
    输入形式 Number(1),
    输入类型 Number(1),
    输入格式 Varchar2(20),
    输入值域 Varchar2(200),
    输入行数 Number(2),
    提取来源 Number(2),
    提取病历 Varchar2(100),
    提取sql Varchar2(4000),
    描述文字 Varchar2(20),
    是否只读 Number(1),
    死亡则隐藏 Number(1)
 )Tablespace ZL9BASEITEM;

Create Table 医生交接班病人类型(
    简称 Varchar2(10),
    名称 Varchar2(20),
    顺序 Number(2),
    起始描述 Varchar2(50),
    提取sql Varchar2(4000),
    是否停用 Number(1)
)Tablespace ZL9BASEITEM;

create table 急诊病人来源 
(
   编码                   varchar2(10),
   名称                   varchar2(50),
   缺省标志               NUMBER(1) default 0
)
tablespace zl9BaseItem;

create table 急诊意识状态 
(
   编码                   varchar2(10),
   名称                   varchar2(50),
   缺省标志               NUMBER(1) default 0
)
tablespace zl9BaseItem;

create table 急诊陪同人员 
(
   编码                   varchar2(10),
   名称                   varchar2(50),
   缺省标志               NUMBER(1) default 0
)
tablespace zl9BaseItem;

create table 急诊常见既往史 
(
   编码                   varchar2(10),
   名称                   varchar2(50)
)
tablespace zl9BaseItem;

create table 急诊常用主诉 
(
   编码                   varchar2(4),
   上级                   varchar2(4),
   名称                   varchar2(50),
   简码                  varchar2(20),  
   末级                  number(1) Default 0
)
tablespace zl9BaseItem;

create table 急诊病情级别 
(
   序号                 number(1),
   名称                 varchar2(5),
   严重程度                 varchar2(20),
   级别描述                 varchar2(400),
   响应要求说明               varchar2(400),
   再次评估时限               number(3),
   患者标识颜色               varchar2(6)
)
tablespace zl9BaseItem;

create table 急诊评分方法 
(
   ID                   number(18),
   英文名                  varchar2(100),
   中文名                  varchar2(100),
   说明                   varchar2(1000)
)
tablespace zl9BaseItem;

create table 急诊人工评定规则 
(
   ID                   number(18),
   分类                   varchar2(20),
   指标名称                 varchar2(200),
   适用人群                 varchar2(10),
   病情级别                 number(1)
)
tablespace zl9BaseItem;

create table 急诊评分方法分级 
(
   ID                       number(18),
   方法ID             number(18),
   分值上限                 number(5),
   分值下限                 number(5),
   运算符                  number(2),--1-等于 2-大于 3-小于 4-大于等于 5-小于等于  6-包含
   评分结果描述               varchar2(200),
   病情级别                 number(1)
)
tablespace zl9BaseItem;

create table 急诊评分方法规则 
(
   ID                   number(18),
   方法ID             number(18),
   指标ID                 number(18),
   指标年龄ID               number(18),
   指标值上限                number(10,2),
   指标值下限                number(10,2),
   运算符                    number(2),--1-等于 2-大于 3-小于 4-大于等于 5-小于等于  6-包含
   指标结果分值               number(5),
   指标结果描述               varchar2(100),
   病情级别                 number(1)
)
tablespace zl9BaseItem;

create table 急诊评分指标 
(
   ID                   number(18),
   方法ID               number(18),
   指标名称                 varchar2(100),
   值域类型                 number(1),
   值域范围                 varchar2(4000),
   值域单位                 varchar2(20)
)
tablespace zl9BaseItem;

create table 急诊评分指标年龄 
(
   ID                   number(18),
   指标ID               number(18),
   年龄上限                 number(4,1),
   年龄下限                 number(4,1),
   运算符                  number(1),
   年龄单位                 varchar2(4),
   年龄段描述                varchar2(20)
)
tablespace zl9BaseItem;


Create Table 聊天会话表(
    id          Number(18),
    对象标识    Varchar2(500), 
    对象内容    Varchar2(4000),
    病人id      Number(18),
    就诊id      NUMBER(18), 
    病人来源    Number(1),  
    创建人      Varchar2(50),
    创建时间    Date,
    接收人      Varchar2(50),
    状态        Number(1))
 Tablespace ZL9BASEITEM Initrans 20;

Create Table 聊天信息表(
    id          Number(18),
    会话id      Number(18),
    发送人      Varchar2(50),
    发送内容    Varchar2(1000),
    发送时间    Date,
    接收人      Varchar2(50),
    阅读时间    Date
 )Tablespace ZL9BASEITEM Initrans 20;

Create Table 压力性损伤分期
(
编码  VARCHAR2(3),     
名称	VARCHAR2(100),
简码  Varchar2(50),
缺省标志 Number(1))
tablespace	ZL9BASEITEM;

create table 医嘱执行组合
(
  条码   VARCHAR2(20),
  要求时间 DATE,
  发送号  NUMBER(18),
  医嘱id NUMBER(18), 
  说明  varchar2(1000),
  待转出  NUMBER(3)  
) TABLESPACE zl9CisRec
Initrans 20;


Create Table 病人中医诊断记录 
(
   诊断ID    NUMBER(18),
   病人ID    NUMBER(18),
   挂号单    varchar2(8),
   姓名      varchar2(100),
   门诊号    NUMBER(18),
   性别      varchar2(4),
   年龄      varchar2(20),
   民族      varchar2(50),
   出生日期  Date,
   就诊方式  NUMBER(1),  
   科别      varchar2(20), 
   疾病ID    Number(18),
   疾病名称  varchar2(50),
   证型ID    Number(18),
   证型名称  varchar2(50),
   中医诊断  varchar2(100),
   中医治法  varchar2(100),
   处方ID    NUMBER(18),
   操作时间  Date,
   操作人    varchar2(100),
   HIS诊断ID NUMBER(18),
   HIS医嘱ID NUMBER(18)
)TABLESPACE ZL9CISREC;

create table 病人中医处方记录 
(
   处方ID                 NUMBER(18),
   方剂ID                 NUMBER(18),
   方剂名称               Varchar2(50),
   付数                   NUMBER(3),
   中药用法               Varchar2(50),
   中药煎法               Varchar2(50),
   煎量                   Varchar2(50),
   用药频率               Varchar2(50), 
   频率次数               NUMBER(3),   
   频率间隔               NUMBER(3),    
   间隔单位               VARCHAR2(4),
   医生嘱托               Varchar2(200),
   HIS煎法ID              NUMBER(18),
   HIS用法ID              NUMBER(18),
   HIS药房ID              NUMBER(18) 
)TABLESPACE ZL9CISREC;

create table 病人中医处方明细
(
   处方明细ID              NUMBER(18), 
   处方ID              NUMBER(18),
   序号                NUMBER(5),
   草药ID              NUMBER(18),
   是否加药            NUMBER(1),
   来源                varchar2(50),
   草药名称            varchar2(50),
   用量                NUMBER(4,2),
   单位                varchar2(20),
   脚注                varchar2(100),
   HIS品种ID           NUMBER(18),
   HIS规格ID           NUMBER(18)
)TABLESPACE ZL9CISREC;

Create Table 草药目录 
(
   草药ID                 NUMBER(18),
   草药名称               Varchar2(50),
   简码                   Varchar2(50),
   别名                   Varchar2(50),
   别名简码               Varchar2(50),
   来源                   Varchar2(50),
   单位                   varchar2(20),
   草药描述               Varchar2(500),
   性状                   Varchar2(200),
   药性                   Varchar2(200),
   适应证                 Varchar2(200),
   用法                   Varchar2(500),
   服法                   Varchar2(500),
   禁忌                   Varchar2(1000),
   成分                   Varchar2(1000),
   药理作用               Varchar2(1000),
   HIS品种ID              NUMBER(18),
   创建人                 Varchar2(100),
   创建时间               date,
   最后修改人             Varchar2(100),
   最后修改时间           date
) TABLESPACE zl9BaseItem;

create table 中医疾病 
(
   疾病ID                NUMBER(18),
   疾病名称              varchar2(50),
   科别                  varchar2(20),
   简码                  varchar2(50),
   创建人                varchar2(100),
   创建时间              date,
   最后修改人            varchar2(100),
   最后修改时间          date
) TABLESPACE zl9BaseItem;

create table 中医证型 
(
   证型ID                 NUMBER(18),
   证型名称               varchar2(50),
   简码                   varchar2(50),
   疾病ID                 NUMBER(18),
   证型描述               varchar2(500),
   证型治法               varchar2(100),
   症状表现               varchar2(500),
   创建人                 varchar2(100),
   创建时间               date,
   最后修改人             varchar2(100),
   最后修改时间           date
) TABLESPACE zl9BaseItem;

create table 治法方剂 
(
   方剂ID                 NUMBER(18),
   方剂名称               varchar2(50),
   简码                   varchar2(50),
   别名                   varchar2(50),
   别名简码               varchar2(50),
   来源                   varchar2(50),
   组成摘要               varchar2(200),
   服法描述               varchar2(200),
   作用描述               varchar2(200),
   制法描述               varchar2(1000),
   适应证描述             varchar2(500),
   方剂组成作用描述       varchar2(2000),
   是否保密               NUMBER(1),
   创建人                 varchar2(100),
   创建时间               date,
   最后修改人             varchar2(100),
   最后修改时间           date
) TABLESPACE zl9BaseItem;

create table 证型方剂对照 
(
   对照ID                 NUMBER(18),
   证型ID                 NUMBER(18),
   方剂ID                 NUMBER(18),
   状态                   NUMBER(2),
   创建人                 varchar2(100),
   创建时间               date,
   最后修改人             varchar2(100),
   最后修改时间           date
) TABLESPACE zl9BaseItem;

create table 方剂构成 
(
   构成ID                 NUMBER(18),
   方剂ID                 NUMBER(18),
   草药ID                 NUMBER(18),
   用法备注               varchar2(100),
   古法用量               varchar2(50),
   用量                   NUMBER(16,5),
   创建人                 varchar2(100),
   创建时间               date
) TABLESPACE zl9BaseItem;

create table 临证加症 
(
   加症ID                 NUMBER(18),
   加症名称               varchar2(50),
   简码                   varchar2(50),
   状态                   NUMBER(2),
   创建人                 varchar2(100),
   创建时间               date,
   最后修改人             varchar2(100),
   最后修改时间           date
) TABLESPACE zl9BaseItem;

create table 加症治法 
(
   治法ID                 NUMBER(18),
   治法名称               varchar2(50),
   简码                   varchar2(50),
   加症ID                 NUMBER(18),
   状态                   NUMBER(2),
   创建人                 varchar2(100),
   创建时间               date,
   最后修改人             varchar2(100),
   最后修改时间           date
) TABLESPACE zl9BaseItem;

create table 加症用药 
(
   用药ID                 NUMBER(18),
   治法ID                 NUMBER(18),
   草药ID                 NUMBER(18),
   用量                   NUMBER(16,5),
   状态                   NUMBER(2),
   创建人                 varchar2(100),
   创建时间               date,
   最后修改人             varchar2(100),
   最后修改时间           date
) TABLESPACE zl9BaseItem;

Create Table 疾病编码章节
(
编码 Varchar2(2),
章节 Varchar2(50),
名称 Varchar2(200),
说明 Varchar2(50),
是否分类 Number(1),
优先级 Number(2)
)Tablespace ZL9BASEITEM;

Create Table 输血反应情况
(
编码 VarChar2(4),
名称 Varchar2(20),
简码 Varchar2(10),
缺省标志 Number(1)DEFAULT 0) 
TABLESPACE zl9BaseItem;

Create Table 电子病历授权访问人员 
(
   ID                     NUMBER(18),
   授权ID                 NUMBER(18),
   人员ID                 NUMBER(18)
)TABLESPACE zl9BaseItem;

Create Table 电子病历授权访问病人 
(
   ID                   NUMBER(18),
   授权ID               NUMBER(18),
   授权类型             Number(3),
   授权内容             VARCHAR2(50)
)TABLESPACE zl9BaseItem;

Create Table 电子病历申请访问病人 
(
   ID                     NUMBER(18),
   申请ID                 NUMBER(18),
   病人ID                 NUMBER(18)
)TABLESPACE zl9BaseItem;

Create Table 电子病历访问授权 
(
   ID                     NUMBER(18),
   授权类型               NUMBER(1),
   申请ID                 NUMBER(18),
   方案名                 VARCHAR2(50),
   访问病人               NUMBER(2),
   访问内容               XMLType,
   访问开始时间           DATE,
   访问结束时间           DATE,
   内容时限               NUMBER(1),
   授权人                 VARCHAR2(20),
   授权时间               DATE,
   作废人                 VARCHAR2(20),
   作废时间               DATE,
   备注                   VARCHAR2(100)
)TABLESPACE zl9BaseItem;

Create Table 电子病历访问日志 
(
   ID                   NUMBER(18),
   病人ID               NUMBER(18),
   就诊ID               NUMBER(18),
   病人来源             NUMBER(1),
   访问内容             VARCHAR2(100),
   内容ID               VARCHAR2(36),
   访问人               VARCHAR2(20),
   访问时间             DATE
)TABLESPACE zl9BaseItem;

Create Table 电子病历访问申请 
(
   ID                   NUMBER(18),
   访问内容             XMLType,
   访问开始时间         DATE,
   访问结束时间         DATE,
   内容时限             NUMBER(1),
   申请原因             VARCHAR2(100),
   审批状态             NUMBER(1),
   申请人               VARCHAR2(20),
   申请时间             DATE,
   撤消人               VARCHAR2(20),
   撤消时间             DATE,
   拒绝人               VARCHAR2(20),
   拒绝时间             DATE
)TABLESPACE zl9BaseItem;

Create Table 三方调用目录 (
    id number(5),
    编号 number(5),
    类别 varchar2(50),
    名称 varchar2(100),
    说明 varchar2(500),
    接入方式 number(1),
    浏览器类型 number(1),
    应用场合 varchar2(3),
    地址 VARCHAR2(500),
    是否停用 number(1),
    ftp地址 varchar2(100),
    ftp访问目录 varchar2(100),
    ftp用户名 varchar2(50),
    ftp密码 varchar2(100),
    ftp本地目录 varchar2(500),
    ftp端口 varchar2(10),
    ftp文件名 varchar2(50),
    菜单显示 number(1),
    工具栏显示 number(1),
    右键菜单显示 number(1),
    小图标 blob,
    大图标 blob,
    修改人 varchar2(50),
    修改时间 date)
    tablespace ZL9BASEITEM
;
Create Table 三方调用参数
(
  接口id number(5),
  序号   number(3),
  参数值  varchar2(200),
  备注   varchar2(500),
  SQLText  varchar2(2000)
)tablespace ZL9BASEITEM;
Create Table 病案首页填写情况
(
  病人id Number(18),
  主页id Number(5),
  项目   Varchar2(50)
)tablespace ZL9PATIENT;
Create Table 人员手术权限申请(
    ID number(18),
    申请人 VARCHAR2(20),
    申请时间 Date,
    授权人员ID number(18),
    诊疗项目ID number(18),
    权限 number(1),
    审核状态 number(1),
    审批人 VARCHAR2(20),
    审批时间 Date
)
TABLESPACE zl9BaseItem;
Create Table 医生常用诊断(
    ID Number(18),
    人员ID number(18),
    科室ID number(18),
    诊断名称 VARCHAR2(500),
    疾病ID number(18),
    诊断ID number(18),
    使用次数 number(18),
    诊断类型 number(2))
    TABLESPACE zl9BaseItem PCTFREE 10 initrans 20
    Cache Storage(Buffer_Pool Keep)
;
create table 医生常用医嘱(
   ID    Number(18),
   人员ID number(18), 
   科室ID number(18),
   诊断名称 VARCHAR2(500),
   疾病ID number(18), 
   诊断ID number(18),
   诊疗项目ID number(18), 
   药品ID number(18), 
   诊疗类别 Varchar2(1), 
   使用次数 number(18)
) TABLESPACE zl9BaseItem
  PCTFREE 10 initrans 20
  Cache Storage(Buffer_Pool Keep);
Create Table 药品用法用量 (
    药品ID number(18),
    用法ID number(18),
    频次 varchar2(3),
    成人剂量 number(16,5),
    小儿剂量 number(16,5),
    医生嘱托 varchar2(100),
    疗程 number(5),
    DDD值 number(16,5),
    性质 number(1))
    tablespace ZL9BASEITEM
;
CREATE TABLE 常用就诊摘要(
	编码 varchar2(4),
	名称 varchar2(4000),
	简码 VARCHAR2(4000),
  	人员ID number(18)
) TABLESPACE zl9BaseItem;
Create Table 输血性质(
       编码 Varchar2(2),  
       名称 Varchar2(50),  
       简码 Varchar2(25),
       缺省标志 Number(1)) 
    Tablespace zl9BaseItem;
Create Table 输液通道(
    编码 VARCHAR2(4),
    名称 VARCHAR2(20),
    简码 VARCHAR2(20))
    TABLESPACE zl9BaseItem;
Create Table 输血检验对照 (
    项目id number(18),
    检验项目id number(18),
    历史结果天数 number(2) default 7)
    TableSpace zl9BaseItem
;
create table 抗菌药物抽样用药(
   抽样ID number(18), 
   序号 number(18),
   药品序号 number(5),
   药名ID number(18), 
   图标 Varchar2(50),
   药名 Varchar2(50),
   单量 Varchar2(50),
   频次 Varchar2(50),
   途径 Varchar2(50),
   总量 Varchar2(50),
   起止日期 Varchar2(50)
) TABLESPACE zl9BaseItem;
create table 抗菌药物抽样手术(
   抽样ID number(18), 
   序号 number(18),
   手术ID number(18),
   手术名称 VARCHAR2(100),
   切口 Varchar2(20),
   开始时间 Date,
   结束时间 Date,
   预防用药期间 number(2),
   给药情况 number(1)
) TABLESPACE zl9BaseItem;
Create Table 抗菌用药评价项目(
    编码 Varchar2(5),
    序号 Number(5),
    名称 Varchar2(200),
    是否手术 Number(1),  
    是否合理 Number(1),  
    上级 Varchar2(5),
    末级 NUMBER(1)
)TABLESPACE zl9BaseItem;
Create Table 抗菌预防用药期间(
     编码 Number(5), 
     名称 Varchar2(200)
)TABLESPACE zl9BaseItem;
create table 抗菌药物抽样评价(
   抽样ID number(18), 
   序号 number(18),
   评价类型 number(1),
   是否合理 number(1),
   项目编码 Varchar2(5),
   行序号 number(5),
   项目值 Varchar2(200)
) TABLESPACE zl9BaseItem;
create table 抗菌药物抽样记录
(
   ID number(18),
   抽样人 Varchar2(20),
   抽样时间 Date,
   范围开始时间 Date,
   范围结束时间 Date
) TABLESPACE zl9BaseItem;
create table 抗菌药物抽样明细
(
   抽样ID number(18),
   病人ID number(18),
   主页ID Number(5),
   序号  Number(18),
   是否手术 Number(1), 
   病原学检测 	Number(1), 
   病原学检测日期	Date,	
   病原学检测标本	varchar2(50), 
   病原学检测检出细菌名	varchar2(100), 
   药敏试验 	Number(1), 
   药敏试验日期	Date,	
   药敏试验是否相符	Number(1), 
   用药前体温 varchar2(30),
   用药前白细胞计数 varchar2(30), 
   用药前中性粒细胞 varchar2(30),
   用药前C反应蛋白 varchar2(30),
   用药前丙谷转氨酶 varchar2(30),
   用药前肌酐 varchar2(30),
   用药后体温 varchar2(30),
   用药后白细胞计数 varchar2(30),
   用药后中性粒细胞 varchar2(30),
   用药后C反应蛋白 varchar2(30),
   用药后丙谷转氨酶 varchar2(30),
   用药后肌酐 varchar2(30),
   用药前体温日期 Date, 
   用药前白细胞计数日期 Date, 
   用药前中性粒细胞日期 Date,
   用药前C反应蛋白日期 Date,
   用药前丙谷转氨酶日期 Date,
   用药前肌酐日期 Date,
   用药后体温日期 Date,
   用药后白细胞计数日期 Date,
   用药后中性粒细胞日期 Date,
   用药后C反应蛋白日期 Date,
   用药后丙谷转氨酶日期 Date,
   用药后肌酐日期 Date, 
   影像学诊断 varchar2(200), 
   影像学诊断部位	varchar2(50),
   影像学诊断结论	varchar2(100),
   临床症状 number(18),  
   用药目的 number(2),     
   感染诊断 number(18),   
   治疗结果 number(2),    
   适应症 varchar2(500),   
   药物选择 varchar2(500),  
   单次剂量 varchar2(500),  
   每日给药频次 varchar2(500), 
   溶剂 varchar2(500),  
   给药途径 varchar2(500),  
   用药疗程 varchar2(500),  
   术前用药时间 varchar2(500),  
   术中用药 varchar2(500),   
   术后用药 varchar2(500),   
   联合用药 varchar2(500), 
   更换药物 varchar2(500),  
   备注 varchar2(500),
   是否打印 Number(1),
   是否编辑 Number(1),
   用药天数 NUMBER(5),
   抗菌药种数 NUMBER(5),
   是否用抗真菌药 Number(1)
) TABLESPACE zl9BaseItem;
CREATE TABLE 输血目的(
    编码 VARCHAR2(4),
    名称 VARCHAR2(100),
    简码 VARCHAR2(20)
)TABLESPACE zl9BaseItem; 
Create Table 临床部门(
    工作性质 VARCHAR2(10),
    部门id NUMBER(18))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 临床性质(
    编码 VARCHAR2(10),
    名称 VARCHAR2(30),
    简码 VARCHAR2(15),
		序号 NUMBER(4))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
Create Table 门诊诊室(
    编码 VARCHAR2(3),
    名称 VARCHAR2(20),
    简码 VARCHAR2(6),
    位置 VARCHAR2(40),
    站点 Varchar2(3),
    缺省标志 NUMBER(1) default 0,
    ID number(18))
    TABLESPACE zl9BaseItem
    Cache Storage(Buffer_Pool Keep) 
;
Create Table 临床医疗小组(
	ID NUMBER(18),
	科室ID NUMBER(18),
	名称 VARCHAR2(50),
	说明 VARCHAR2(200),
        建档时间 Date,
	撤档时间 Date)
    TABLESPACE zl9BaseItem
    Cache Storage(Buffer_Pool Keep);
Create Table 医疗小组人员(
	小组ID NUMBER(18),
	人员ID NUMBER(18),
	是否组长 Number(1))
    TABLESPACE zl9BaseItem
    Cache Storage(Buffer_Pool Keep);
 Create Table 人员抗菌药物权限(
	人员id Number(18), 
	级别 Number(1), 
	记录状态 Number(3) Default (1), 
	操作人员 Varchar2(20), 
	操作时间 Date,
	场合 Number(2) default(1)) 
	Tablespace Zl9baseitem;
Create Table 人员手术权限(
  人员id Number(18), 
  诊疗项目ID Number(18),
  记录性质 Number(3)) 
  Tablespace Zl9baseitem;  
CREATE TABLE 病生理情况(
	编码 VARCHAR2(3),
	名称 VARCHAR2(100))
TABLESPACE zl9BaseItem;
Create Table 单病种目录(
    编码 VARCHAR2(2),
    名称 VARCHAR2(50),
    简码 VARCHAR2(25),
    ICD编码 VARCHAR2(1000))
    TABLESPACE zl9BaseItem;
Create Table 常用体温说明(
    编码 VARCHAR2(2),
    名称 VARCHAR2(10),
    简码 VARCHAR2(6)
    )
    TABLESPACE zl9BaseItem;
Create table 体温标记说明 (
	编码 varchar2(2)  ,
	名称 varchar2(10) ,
	简码 varchar2(6)
	)
	TABLESPACE zl9BaseItem;
Create Table 医嘱内容定义(
    诊疗类别 VARCHAR2(1),
    医嘱内容 VARCHAR2(500))
    TABLESPACE zl9BaseItem;
Create Table 医嘱未执行原因(
    编码 VARCHAR2(5),
    名称 VARCHAR2(100),
    简码 VARCHAR2(10))
    TABLESPACE zl9BaseItem;
Create Table 医嘱常用原因(
    编码 VARCHAR2(4),
    名称 VARCHAR2(200),
    简码 VARCHAR2(200),
    性质 Number(1),
    人员 Varchar2(100))
    TABLESPACE zl9BaseItem;
Create Table 诊疗参考分类(
    ID NUMBER(18),
    上级id NUMBER(18),
    编码 VARCHAR2(8),
    名称 VARCHAR2(40),
    简码 VARCHAR2(10),
    类型 number(1))
    TableSpace zl9BaseItem;
Create Table 诊疗参考目录(
    ID NUMBER(18),
    分类ID NUMBER(18),
    编码 VARCHAR2(13),
    名称 VARCHAR2(60),
    说明 VARCHAR2(4000),
    编者 VARCHAR2(20),
    类型 NUMBER(1))
    TableSPACE zl9BaseItem;
Create Table 诊疗参考别名(
    参考目录ID NUMBER(18),
    名称 VARCHAR2(60),
    性质 NUMBER(1),
    简码 VARCHAR2(12),
    码类 NUMBER(1))
    TableSpace zl9BaseItem;
Create Table 诊疗参考内容(
    参考目录ID NUMBER(18),
    项目序号 NUMBER(5),
    参考项目 VARCHAR2(20),
    项目层次 NUMBER(1),
    内容行号 NUMBER(5),
    内容文本 VARCHAR2(4000),
    内容性质 NUMBER(3))
    TableSPACE zl9BaseItem;
Create Table 诊疗参考疾病(
    参考目录ID NUMBER(18),
    参考项目 VARCHAR2(20),
    内容行号 NUMBER(5),
    禁忌症ID NUMBER(18),
    禁忌类型 NUMBER(1))
    TableSPACE zl9BaseItem;
CREATE TABLE 诊疗项目类别(
    编码 VARCHAR2(1),
    名称 VARCHAR2(10),
    简码 VARCHAR2(10))
    TableSpace zl9BaseItem;
CREATE TABLE 诊疗分类目录(
    ID NUMBER(18),
    编码 VARCHAR2(20),
    名称 VARCHAR2(40),
    简码 VARCHAR2(10),
    上级id NUMBER(18),
    类型 NUMBER(1),
		建档时间 DATE,
		撤档时间 DATE)
    TableSpace zl9BaseItem;

Create Table 诊疗手术等级
(名称 Varchar2(50),
诊疗项目ID Number(18),
序号 Number(5),
手术类型 Number(1),
操作类型 Number(1)
)Tablespace ZL9BASEITEM;

Create Table 诊疗项目目录(
    类别 VARCHAR2(1),
    分类ID NUMBER(18),
    ID NUMBER(18),
    编码 VARCHAR2(20),
    名称 VARCHAR2(60),
    标本部位 VARCHAR2(60),
    计算单位 VARCHAR2(20),
    计算方式 NUMBER(1),
    计算规则 Number(1),
    执行频率 NUMBER(1),
    适用性别 NUMBER(1),
    单独应用 NUMBER(1),
    组合项目 NUMBER(1),
    操作类型 VARCHAR2(20),
    执行安排 NUMBER(1),
    执行科室 NUMBER(1),
    服务对象 NUMBER(1),
    计价性质 NUMBER(1),
    参考目录ID NUMBER(18),
    人员ID NUMBER(18),
    建档时间 DATE,
    撤档时间 DATE,
    录入限量 NUMBER(16,5),
    试管编码 Varchar2(4),
    执行分类 NUMBER(2),
    执行标记 NUMBER(1),
    站点 Varchar2(3),
    建档人 Varchar2(20),
    计算系数 number(16,5),
    诊疗频率编码 VARCHAR2(3),
    操作人员 varchar2(100),
    是否保密 NUMBER(1))
    TableSpace zl9BaseItem
    Cache Storage(Buffer_Pool Keep) 
;
Create Table 诊疗个人项目(
    人员ID NUMBER(18),
	诊疗项目ID NUMBER(18),
	收费细目ID NUMBER(18),
	频度 Number(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE 诊疗项目别名(
    诊疗项目id NUMBER(18),
    名称 VARCHAR2(60),
    性质 NUMBER(1),
    简码 VARCHAR2(30),
    码类 NUMBER(1))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
CREATE TABLE 诊疗互斥项目(
    组编号 NUMBER(18),
    组名称 VARCHAR2(30),
    项目ID NUMBER(18),
    类型 NUMBER(18))
    TableSpace zl9BaseItem;
CREATE TABLE 诊疗适用科室(
    项目ID NUMBER(18),
    科室ID NUMBER(18))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
CREATE TABLE 诊疗执行科室(
    诊疗项目id NUMBER(18),
    病人来源 NUMBER(1) DEFAULT 1,
    开单科室id NUMBER(18),
    执行科室id NUMBER(18))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
create table 诊疗用法用量
(
  项目id NUMBER(18),
  性质   NUMBER(3),
  用法id NUMBER(18),
  频次   VARCHAR2(3),
  成人剂量 NUMBER(16,5),
  小儿剂量 NUMBER(16,5),
  医生嘱托 VARCHAR2(100),
  疗程   NUMBER(5),
  ddd值 NUMBER(16,5)
)
tablespace ZL9BASEITEM;
CREATE TABLE 诊疗项目组合(
		诊疗组合ID NUMBER(18),
		序号 NUMBER(18),
		相关序号 NUMBER(18),
		期效 NUMBER(1),
		诊疗项目ID NUMBER(18),
		医嘱内容 Varchar2(1000),
		天数 NUMBER(16,5),
		单次用量 NUMBER(16,5),
		总给予量 NUMBER(16,5),
		收费细目ID NUMBER(18),
		标本部位 VARCHAR2(60),
		检查方法 Varchar2(30),
		医生嘱托 VARCHAR2(100),
		执行频次 VARCHAR2(20),
		频率次数 NUMBER(3),
		频率间隔 NUMBER(3),
		间隔单位 VARCHAR2(4),
		执行性质 NUMBER(1),
		执行标记 NUMBER(1),
		执行科室ID NUMBER(18),
		时间方案 VARCHAR2(50),
		配方ID Number(18),
		组合项目ID Number(18),
		配方类型 number(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 诊疗收费关系(
    诊疗项目id NUMBER(18),
    收费项目id NUMBER(18),
    收费数量 NUMBER(16,5) DEFAULT 1,
    固有对照 NUMBER(1),
    从属项目 Number(1),
		费用性质 Number(1) default 0 not Null,
    检查部位 Varchar2(30),
    检查方法 Varchar2(30),
		收费方式 Number(1),
		适用科室ID NUMBER(18),
		病人来源 NUMBER(1) Default 0 Not Null)
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
CREATE TABLE 诊疗频率项目(
    编码 VARCHAR2(3),
    名称 VARCHAR2(20),
    简码 VARCHAR2(10),
    英文名称 VARCHAR2(50),
    频率次数 NUMBER(3),
    频率间隔 NUMBER(3),
    间隔单位 VARCHAR2(4),
    适用范围 NUMBER(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 诊疗频率时间(
    执行频率 VARCHAR2(3),
    方案序号 NUMBER(3),
    时间方案 VARCHAR2(50),
    给药途径ID NUMBER(18))
    TABLESPACE zl9BaseItem;
Create Table 药比附加条件(
	类别 Varchar2(10),
	内容 Varchar2(4000)) 
	TABLESPACE zl9BaseItem;
CREATE TABLE 诊疗麻醉类型(
    编码 VARCHAR2(2),
    名称 VARCHAR2(20),
    简码 VARCHAR2(10),
    缺省标志 NUMBER(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 诊疗手术规模(
    编码 VARCHAR2(2),
    名称 VARCHAR2(10),
    简码 VARCHAR2(10),
    缺省标志 NUMBER(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 中药煎服脚注(
    编码 VARCHAR2(5),
    名称 VARCHAR2(20),
    简码 VARCHAR2(8))
    TABLESPACE zl9BaseItem;
Create Table 中药输入快捷(
	编码 VARCHAR2(2),
    名称 VARCHAR2(1) Not Null,
    数值 NUMBER(16,5) Not Null)
    TABLESPACE zl9BaseItem;
CREATE TABLE 诊治所见性质(
    编码 VARCHAR2(1),
    名称 VARCHAR2(20),
    简码 VARCHAR2(10),
    固定 NUMBER(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 常用嘱托(
    编码 VARCHAR2(5),
    名称 VARCHAR2(200),
    简码 VARCHAR2(200),
		人员 VARCHAR2(20))
    TABLESPACE zl9BaseItem;
CREATE TABLE 常用剂量比例(
    编码 VARCHAR2(3),
    名称 VARCHAR2(30),
    比例 Number(5,2))
    TABLESPACE zl9BaseItem;
CREATE TABLE 疾病参考项目(
    类别 NUMBER(1),
    主序号 NUMBER(3),
    子序号 NUMBER(3),
    名称 VARCHAR2(20),
    格式 NUMBER(1),
    层次 NUMBER(1),
    性质 NUMBER(1))
    TABLESPACE zl9BaseItem;
Create Table 诊疗参考项目(
    类型 Number(1),
    序号 NUMBER(3),
    层次 NUMBER(1),
    名称 VARCHAR2(20),
    性质 NUMBER(1))
    TableSPACE zl9BaseItem;
CREATE TABLE 疾病诊断分类(
    ID NUMBER(18),
    上级ID NUMBER(18),
    编码 VARCHAR2(6),
    名称 VARCHAR2(40),
    简码 VARCHAR2(10),
    类别 NUMBER(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 疾病诊断属类(
    分类ID NUMBER(18),
    诊断ID NUMBER(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE 疾病诊断目录(
    ID NUMBER(18),
    类别 NUMBER(1),
    编码 VARCHAR2(10),
    名称 VARCHAR2(40),
    适用范围 Number(1),
    说明 VARCHAR2(4000),
    编者 VARCHAR2(20),
    疑似 NUMBER(5),
    临床 NUMBER(5),
    建档时间 DATE ,
    撤档时间 DATE)
    TABLESPACE zl9BaseItem   
    Cache Storage(Buffer_Pool Keep);
Create TABLE 疾病诊断科室(
    诊断ID NUMBER(18),
    科室ID NUMBER(18),
    人员ID Number(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE 疾病诊断别名(
    诊断id NUMBER(18),
    名称 VARCHAR2(40),
    性质 NUMBER(1),
    简码 VARCHAR2(12),
    码类 NUMBER(1))
    TABLESPACE zl9BaseItem   
    Cache Storage(Buffer_Pool Keep);
CREATE TABLE 疾病诊断参考(
    诊断ID NUMBER(18),
    项目序号 NUMBER(5),
    参考项目 VARCHAR2(20),
    项目层次 NUMBER(1),
    项目格式 NUMBER(1),
    证候ID NUMBER(18),
    证候序号 NUMBER(5),
    证候名称 VARCHAR2(20),
    内容行号 NUMBER(5),
    内容文本 VARCHAR2(4000),
    内容性质 NUMBER(1))
    TABLESPACE zl9BaseItem;
Create Table 诊断病种对应(
    诊断ID NUMBER(18),
    险类 NUMBER(3),
    病种ID NUMBER(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE 疾病诊疗措施(
    诊断ID NUMBER(18),
    参考项目 VARCHAR2(20),
    证候名称 VARCHAR2(20),
    内容行号 NUMBER(5),
    诊疗项目ID NUMBER(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE 疾病诊断规则(
    诊断ID NUMBER(18),
    分组号 NUMBER(3),
    分组名 VARCHAR2(20),
    条件号 NUMBER(3),
    项目ID NUMBER(18),
    关系式 VARCHAR2(10),
    条件值 VARCHAR2(250),
    怀疑度 NUMBER(3))
    TABLESPACE zl9BaseItem;
CREATE TABLE 疾病诊断对照(
    疾病ID NUMBER(18),
    诊断ID NUMBER(18),
    手术ID NUMBER(18))
    TABLESPACE zl9BaseItem;
----------------------------------------------------------------------------
--[[7.临床路径基础]]
----------------------------------------------------------------------------
CREATE TABLE 临床路径审核(
    路径ID NUMBER(18),
    版本号 NUMBER(3),
    操作时间 Date,
    操作状态 Number(1),
    操作人员 Varchar2(20),  
    操作说明 Varchar2(200))
    TABLESPACE ZL9CISREC;

CREATE TABLE 路径报表目录(
	ID		NUMBER(18),
	编码    VARCHAR2(64),
	名称	VARCHAR2(100),
	是否固定 NUMBER(1)
	)
	TABLESPACE zl9BaseItem;
CREATE TABLE 路径报表结构(		
	报表ID	NUMBER(18),
	行号	NUMBER(5),
	项目序号	NUMBER(5),
	项目文本1 VARCHAR2(100),
	项目文本2 VARCHAR2(100),
	SQL文本 VARCHAR2(4000),
	页数 number(3),
	路径ID number(18),
	多选序号 number(5)
	)
    TABLESPACE zl9BaseItem;
Create Table 路径报表序号 (
   报表ID number(18),
   行号  NUMBER(5),
   路径ID number(18),
   序号 Number(8)
) TABLESPACE zl9BaseItem;
Create Table 标准路径目录(
 ID NUMBER(8),
 科室名称 Varchar2(100),
 编码 Varchar2(8),   
 路径名称 Varchar2(80),
 类别  NUMBER(2),
 版本说明 Varchar2(20) 
)
 tablespace ZL9BASEITEM
;
create table 标准路径流程(
	  标准路径id NUMBER(8) ,
	  序号     NUMBER(3) ,
	  标题     VARCHAR2(100),
	  内容     VARCHAR2(4000)
	)
	tablespace ZL9BASEITEM;
Create Table 标准路径病种(
    标准路径id NUMBER(8),
    疾病编码 varchar2(200),
    手术编码 VARCHAR2(100))
    tablespace ZL9BASEITEM
;
create table 标准路径表单(
	  标准路径id NUMBER(8),
	  表单序号   NUMBER(3),
	  表单名称   VARCHAR2(100),
	  表单表头   Varchar2(500),
	  分类序号   NUMBER(3),
	  分类名称   VARCHAR2(50),
	  阶段序号   NUMBER(3),
	  阶段名称   VARCHAR2(100),
	  路径内容   VARCHAR2(2000)
	)
	tablespace ZL9BASEITEM;
CREATE TABLE 路径项目顺序(
	顺序 number(2),
	医嘱期效 NUMBER(1),
	诊疗类别 VARCHAR2(1),
	操作类型 VARCHAR2(20),
	执行分类 NUMBER(2))
TableSpace zl9BaseItem       
    Cache Storage(Buffer_Pool Keep);
Create Table 临床病例分型(
    编码 VARCHAR2(1),
    名称 VARCHAR2(20),
    简码 VARCHAR2(10))
    TABLESPACE zl9BaseItem;
Create Table 路径结果性质(
    编码 NUMBER(2),
    名称 VARCHAR2(20),
    简码 VARCHAR2(10))
    TABLESPACE zl9BaseItem;
Create Table 路径常见结果(
    编码 VARCHAR2(5),
    名称 VARCHAR2(100),
    简码 VARCHAR2(10),
		上级 VARCHAR2(5),
		末级 NUMBER(1),
		基本 NUMBER(1),
		性质 NUMBER(2))
    TABLESPACE zl9BaseItem;
Create Table 变异常见原因(
    编码 VARCHAR2(6),
    名称 VARCHAR2(200),
    简码 VARCHAR2(20),
	上级 VARCHAR2(6),
	末级 NUMBER(1),
	性质 NUMBER(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 临床路径图标(
	ID NUMBER(18),
	图标 BLOB,
	性质 NUMBER(1))
	LOB(图标) Store as (Cache)
    TABLESPACE zl9BaseItem;
----------------------------------------------------------------------------
--[[8.病历基础]]
----------------------------------------------------------------------------
CREATE TABLE 家系称谓关系(
    序号 NUMBER(3),
    父亲 NUMBER(3),
    母亲 NUMBER(3),
    称谓 VARCHAR2(10),
    关系 VARCHAR2(12),
    性别 VARCHAR2(4),
    唯一关系 NUMBER(3),
    辈分等级 NUMBER(5),
    亲属等级 NUMBER(3),
    血缘关系 NUMBER(3),
	国标代码 varchar2(2))
    TABLESPACE zl9BaseItem;
CREATE TABLE 诊治所见分类(
    性质 VARCHAR2(1),
    ID NUMBER(18),
    上级ID NUMBER(18),
    编码 VARCHAR2(6),
    名称 VARCHAR2(40),
    简码 VARCHAR2(8))
    TABLESPACE zl9BaseItem;
CREATE TABLE 诊治所见项目(
    ID NUMBER(18),
    分类ID NUMBER(18),
    编码 VARCHAR2(13),
    中文名 VARCHAR2(60),
    英文名 VARCHAR2(40),
    替换域 NUMBER(1),
    类型 NUMBER(3),
    长度 NUMBER(3),
    小数 NUMBER(3),
    单位 VARCHAR2(20),
    临床意义 VARCHAR2(250),
    表示法 NUMBER(1),
    性别域 NUMBER(1),
    数值域 VARCHAR2(1000),
    正常域 VARCHAR2(1000),
    初始值 VARCHAR2(1000),
    文字表述 NUMBER(1),
    空值文字 VARCHAR2(100),
    必填 Number(1) Default 0,
	动态域 Number(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 病历标记图形(
    编码 VARCHAR2(4),
    名称 VARCHAR2(30),
    简码 VARCHAR2(10),
    图形 BLOB)	
    TABLESPACE zl9BaseItem;
CREATE TABLE 病历常用样式(
    编号 NUMBER(3),
    名称 VARCHAR2(30),
    段落样式 VARCHAR2(4000),
    字体样式 VARCHAR2(4000),
    系统 NUMBER(1))
    TABLESPACE zl9BaseItem;
Create Table 病历书写事件 (
  种类 NUMBER(3),
  编号 NUMBER(3),
  名称 Varchar2(20),
  简码 Varchar2(10),
  说明 Varchar2(100),
  事前病历 Number(1),
  循环病历 Number(1))
  TABLESPACE zl9BaseItem;
CREATE TABLE 病历页面格式(
    种类 NUMBER(3),
    编号 VARCHAR2(3),
    名称 VARCHAR2(30),
    报表 NUMBER(1),
    格式 VARCHAR2(4000),
    页眉 VARCHAR2(1000),
    页脚 VARCHAR2(1000),
    图形 BLOB,
	页眉文件 BLOB,
	页脚文件 BLOB)
    TABLESPACE zl9BaseItem;
Create Table 病历文件列表(
    ID NUMBER(18),
    种类 NUMBER(3),
    子类 Varchar2(10),
    编号 VARCHAR2(3),
    名称 VARCHAR2(30),
    说明 VARCHAR2(2000),
    页面 VARCHAR2(3),
    保留 NUMBER(5),
    通用 NUMBER(3),
    格式 Number(5))
    TABLESPACE zl9BaseItem 
;
CREATE TABLE 病历时限要求(
    文件ID NUMBER(18),
    事件 VARCHAR2(20),
    必须 NUMBER(1),
    唯一 NUMBER(1),
    书写时限 NUMBER(5),
    审阅时限 NUMBER(5),
    诊断时限 NUMBER(5),
    一般周期 NUMBER(5),
    病重周期 NUMBER(5),
    病危周期 NUMBER(5))
    TABLESPACE zl9BaseItem;
CREATE TABLE 病历替代关系(
    文件ID NUMBER(18),
    替代ID NUMBER(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE 病历应用科室(
    文件ID NUMBER(18),
    科室ID NUMBER(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE 疾病报告前提(
    文件ID NUMBER(18),
    疾病ID NUMBER(18),
    诊断ID NUMBER(18),
	报告病种 Varchar2(80))
    TABLESPACE zl9BaseItem;
Create Table 病历单据应用(
    诊疗项目ID Number(18),
    应用场合 Number(3),
    病历文件ID Number(18))
    Tablespace zl9BaseItem;
Create Table 病历单据附项(
    文件ID Number(18),
    项目 Varchar2(30),
    必填 Number(1),
    排列 Number(5),
    要素ID Number(18),
    只读 number(1),
    内容 Varchar2(200))
    Tablespace zl9BaseItem;
    
create table 病历附项模板(
    ID NUMBER(18),
    病历文件Id NUMBER(18),
    单据附项 VARCHAR2(30),
    模板标题 VARCHAR2(30),
    模板内容 VARCHAR2(512),
    使用次数 number(8)       
)TABLESPACE zl9BaseItem;
    
CREATE TABLE 病历文件格式(
    文件ID NUMBER(18),
    内容 BLOB)
	LOB(内容) Store as (Cache)
    TABLESPACE zl9BaseItem;
CREATE TABLE 病历文件结构(
    ID NUMBER(18),
    文件ID NUMBER(18),
    父ID NUMBER(18),
    对象序号 NUMBER(18),
    对象类型 NUMBER(1),
    对象标记 NUMBER(18),
    保留对象 NUMBER(1),
    对象属性 VARCHAR2(1000),
    内容行次 NUMBER(18),
    内容文本 VARCHAR2(4000),
    是否换行 NUMBER(1),
    预制提纲ID NUMBER(18),
    复用提纲 NUMBER(1),
    使用时机 VARCHAR2(8),
    诊治要素ID NUMBER(18),
		替换域 NUMBER(1),
    要素名称 VARCHAR2(40),
    要素类型 NUMBER(3),
    要素长度 NUMBER(3),
    要素小数 NUMBER(3),
    要素单位 VARCHAR2(50),
    要素表示 NUMBER(3),
    输入形态 NUMBER(3),
    要素值域 VARCHAR2(4000))
    TABLESPACE zl9BaseItem;
CREATE TABLE 病历文件图形(
    对象ID NUMBER(18),
    图形 BLOB)
	LOB(图形) Store as (Cache)
    TABLESPACE zl9BaseItem;
Create Table 病历词句分类(
    ID Number(18),
    上级ID Number(18),
    编码 Varchar2(8),
    名称 Varchar2(30),
    说明 Varchar2(200),
    范围 Varchar2(8))
    Tablespace zl9BaseItem;
Create Table 病历词句示范(
    ID Number(18),
    分类ID Number(18),
    编号 Varchar2(13),
    名称 Varchar2(60),
    通用级 Number(1),
    科室id Number(18),
    人员id Number(18))
    Tablespace zl9BaseItem;
CREATE TABLE 病历词句组成(
    词句ID NUMBER(18),
    排列次序 NUMBER(5),
    内容性质 NUMBER(3),
    内容文本 VARCHAR2(4000),
    诊治要素ID NUMBER(18),
	替换域 NUMBER(1),
    要素名称 VARCHAR2(40),
    要素类型 NUMBER(3),
    要素长度 NUMBER(3),
    要素小数 NUMBER(3),
    要素单位 VARCHAR2(10),
    要素表示 NUMBER(3),
    要素值域 VARCHAR2(4000),
    输入形态 NUMBER(3),
    对象属性 Varchar2(1000))
    TABLESPACE zl9BaseItem;
Create Table 病历词句条件(
    词句ID Number(18),
    条件项 Varchar2(20),
    条件值 Varchar2(2000))
    Tablespace zl9BaseItem;
Create Table 病历提纲词句(
    提纲ID Number(18),
    词句分类ID Number(18))
    Tablespace zl9BaseItem;
CREATE TABLE 病历范文目录(
    ID NUMBER(18),
    文件ID NUMBER(18),
    编号 VARCHAR2(5),
    名称 VARCHAR2(30),
    简码 VARCHAR2(10),
    分类 Varchar2(50),
    性质 NUMBER(1),
    说明 VARCHAR2(100),
    通用级 NUMBER(1),
    科室id NUMBER(18),
    人员id NUMBER(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE 病历范文格式(
    文件ID NUMBER(18),
    内容 BLOB)
	LOB(内容) Store as (Cache)
    TABLESPACE zl9BaseItem;
CREATE TABLE 病历范文内容(
    ID NUMBER(18),
    文件ID NUMBER(18),
    父ID NUMBER(18),
    对象序号 NUMBER(18),
    对象类型 NUMBER(1),
    对象标记 NUMBER(18),
    保留对象 NUMBER(1),
    对象属性 VARCHAR2(1000),
    内容行次 NUMBER(18),
    内容文本 VARCHAR2(4000),
    是否换行 NUMBER(1),
    预制提纲ID NUMBER(18),
		定义提纲ID Number(18),
    复用提纲 NUMBER(1),
    使用时机 VARCHAR2(2),
    诊治要素ID NUMBER(18),
		替换域 NUMBER(1),
    要素名称 VARCHAR2(40),
    要素类型 NUMBER(3),
    要素长度 NUMBER(3),
    要素小数 NUMBER(3),
    要素单位 VARCHAR2(50),
    要素表示 NUMBER(3),
    输入形态 NUMBER(3),
    要素值域 VARCHAR2(4000))
    TABLESPACE zl9BaseItem;
CREATE TABLE 病历范文图形(
    对象ID NUMBER(18),
    图形 BLOB)
	LOB(图形) Store as (Cache)
    TABLESPACE zl9BaseItem;
Create Table 病历范文条件(
    范文ID Number(18),
    条件项 Varchar2(20),
    条件值 Varchar2(2000))
    Tablespace zl9BaseItem;
Create Table 病历范文包(
       ID Number(18),
       编号 Varchar2(5),
       名称 Varchar2(30),
       简码 Varchar2(10),
       说明 Varchar2(100),
       通用级 Number(1),
       科室ID Number(18),
       人员ID Number(18))
    TABLESPACE zl9BaseItem;
Create Table 病历范文包组成(
       范文包ID Number(18),
       范文ID  Number(18))
    TABLESPACE zl9BaseItem;
--病案审查
Create Table 病案审查方案(
    ID		Number(18),
    名称	Varchar2(50),
    总分	Number(5,2),
    分段线	Number(5,2),
    启用时间	Date,
    停用时间	Date,
    说明 VARCHAR2(200))
    TableSpace zl9BaseItem;
Create Table 病案审查分类(
    ID		Number(18),
    上级id	Number(18),
    编码	Varchar2(10),
    名称	Varchar2(30),
    方案ID	Number(18))
    TableSpace zl9BaseItem;
Create Table 病案审查目录(
    ID		Number(18),
    分类ID	Number(18),
    编码	Varchar2(10),
    名称	Varchar2(255),
    简码	Varchar2(255),
    说明	Varchar2(2000),
    审查依据	Varchar2(4000),
    适用对象	Number(3),
    文件ID	Varchar2(2000),
    适用环节	Varchar2(1),
    分值	Number(5,2),
    分制	Number(1),
	数据源 Number(1) Default 0)
    TableSpace zl9BaseItem;
----------------------------------------------------------------------------
--[[9.护理基础]]
----------------------------------------------------------------------------
Create table 护理内容导入定义
(
类别 number(1),
名称 varchar2(100),
格式 varchar2(500)
)tablespace zl9BaseItem;
Create Table 输血类型(
    编码 VARCHAR2(2),
    名称 VARCHAR2(20),
    简码 VARCHAR2(10),
    缺省标志 NUMBER(1) Default 0)
    TABLESPACE zl9BaseItem;
Create Table 输血申请项目(
    医嘱ID NUMBER(18),
    诊疗项目ID number(18),
    申请量 number(16,5),
    申请血型 varchar2(10),
    申请RH varchar2(10),
    血液信息 varchar2(200),
    待转出 NUMBER(3),
    建议输注速度 varchar2(20),
    建议输注单位 varchar2(10))
    tablespace ZL9CISREC;
Create Table 护理记录项目(
    项目序号 NUMBER(5),
    项目名称 VARCHAR2(20),
    保留项目 NUMBER(1),
    项目类型 NUMBER(3),
    项目长度 NUMBER(3),
    项目小数 NUMBER(3),
    项目单位 VARCHAR2(10),
    项目表示 NUMBER(3),
    项目值域 VARCHAR2(4000),
    项目性质 Number(3),
    项目ID NUMBER(18),
    护理等级 NUMBER(3),
    适用科室 Number(3),
    适用病人 Number(3),
    应用方式 Number(3),
    操作类型 VARCHAR2(20),
    分组名 VARCHAR2(20),
    说明 VARCHAR2(1000),
    应用场合 number(1) DEFAULT 0,
    缺省值 VARCHAR2(100),
    分组汇总 VARCHAR2(100))
    TABLESPACE zl9BaseItem;
    --无样本数据,,同一块上的并发事务可能较多
CREATE TABLE 护理项目模板(
	科室ID NUMBER (18),
	模板名称 VARCHAR2 (50),
	护理等级 NUMBER (3),	--0-特级;1-一级;2-二级;3-三级;-1不限护理等级
	项目序号 NUMBER (5),
	排列序号 NUMBER (3))
	TABLESPACE zl9BaseItem;
CREATE TABLE 体温部件(
	名称 VARCHAR2 (50),
	适用地区 VARCHAR2 (50),
	部件 VARCHAR2 (50),
	新部件 VARCHAR2 (50),
	启用 NUMBER (1) DEFAULT 0)
	TABLESPACE zl9BaseItem;
CREATE TABLE 体温部位(
    项目序号 NUMBER (5),
    部位 VARCHAR2 (50),
    标记符号 VARCHAR2 (10),
    标记颜色 NUMBER (18),
    标记图形 BLOB,
	缺省项 NUMBER (1) DEFAULT 0,
	固定项 NUMBER(1) DEFAULT 0)
    TABLESPACE zl9BaseItem;
CREATE TABLE 体温重叠标记(
    序号	NUMBER(5),
    上级序号	NUMBER(5),
    项目序号	NUMBER(5),
    体温部位	VARCHAR2(10),
    重叠数目	Number(5),
    重叠项目	Varchar2(2000),
    标记符号	Varchar2(10),
    标记颜色	Number(18),
    标记图形	Blob)
    TABLESPACE zl9BaseItem;
CREATE TABLE 护理适用科室(
    项目序号 NUMBER(5),
    科室id NUMBER(18))
    TABLESPACE zl9BaseItem;
Create Table 体温记录项目(
    项目序号 NUMBER(5),
    排列序号 NUMBER(3),
    记录名 VARCHAR2(20),
    记录法 NUMBER(3),
    记录符 Varchar2(20),
    记录色 NUMBER(18),
    最大值 NUMBER(16,5),
    最小值 NUMBER(16,5),
    单位值 NUMBER(16,5),
    记录频次 Number(1),
    单位 Varchar2(10),
    最高行 NUMBER(5),
    刻度间隔 NUMBER (16,5),
    警示线 NUMBER (16,5),
    临界值 Varchar2(30),
    入院首测 NUMBER (1) DEFAULT 0)
    TABLESPACE zl9BaseItem 
;
CREATE TABLE 护理汇总时段(
	名称 VARCHAR2 (20),
	开始 VARCHAR2 (5),
	结束 VARCHAR2 (5),
	类别 NUMBER (1) DEFAULT 1,
	单据 NUMBER (1) DEFAULT 1,
	科室ID NUMBER(18))
	TABLESPACE zl9BaseItem;
CREATE TABLE 护理汇总项目(
	序号 NUMBER (5),
	父序号 NUMBER (5))
	TABLESPACE zl9BaseItem;
CREATE TABLE 护理波动项目(
	项目序号 NUMBER (5),
	项目名称 varchar2(20))
	TABLESPACE zl9BaseItem;
CREATE TABLE 体温同步项目(
	护理等级 NUMBER(3),
	年龄范围 Varchar2(50),
	禁用项目 Varchar2(100),
	适用科室 varchar2(200))
	TABLESPACE zl9BaseItem;
Create Table 病区标记内容(
    病区ID NUMBER (18),
    主题序号 NUMBER(18),
    标记序号 NUMBER (5),
    说明 VARCHAR2(20),
    图形索引 NUMBER (5),
    有效天数 NUMBER (5),
    是否特殊 number(1))
    TABLESPACE zl9BaseItem
;
CREATE TABLE 病区公告栏样式(
	ID NUMBER(18),
	病区ID NUMBER (18),
	名称 VARCHAR2 (20) NOT NULL,
	别名 VARCHAR2 (20),
	诊疗项目 XMLTYPE,
	行号 NUMBER (18),
	位置 NUMBER (1) DEFAULT 1,
	是否固定 NUMBER (1),
	是否隐藏 NUMBER (1),
	内容 VARCHAR2 (500),
	时间 DATE )
	TABLESPACE zl9BaseItem;
CREATE TABLE 护理项目频次(
	频次 NUMBER (1),
	序号 NUMBER (1),
	开始 VARCHAR2 (5),
	结束 VARCHAR2 (5),
	类别 NUMBER (1) DEFAULT 1)
	TABLESPACE zl9BaseItem;
CREATE TABLE 产程部件(
	名称 VARCHAR2 (50),
	适用地区 VARCHAR2 (50),
	部件 VARCHAR2 (50),
	启用 NUMBER (1) DEFAULT 0)
	TABLESPACE zl9BaseItem;
CREATE TABLE 产程要素内容(
	文件ID NUMBER (18),
	婴儿 NUMBER (1),
	名称 VARCHAR2 (60),
	内容 VARCHAR2 (100),
	待转出 Number(3))
	TABLESPACE zl9BaseItem;
CREATE TABLE 麻醉记录项目(
    序号 NUMBER(3),
    记录名 VARCHAR2(20),
    记录符 VARCHAR2(2),
    记录色 NUMBER(18),
    最大值 NUMBER(16,5),
    最小值 NUMBER(16,5),
    单位 Varchar2(10),
    记录法 number(3),
    保留 number(1),
    项目ID NUMBER(18))
    TABLESPACE zl9BaseItem;
----------------------------------------------------------------------------
--[[10.检验基础]]
----------------------------------------------------------------------------
CREATE TABLE 检验报告项目(
    ID         Number(18),
    诊疗项目ID NUMBER(18),
    检验标本 VARCHAR2(20),
    排列序号 NUMBER(5),
    报告项目ID NUMBER(18),
    细菌ID NUMBER(18))
    TABLESPACE zl9BaseItem;
Create Table 检验备注文字(
	编码 varchar2(10),
	名称 varchar2(100) not Null,
	简码 varchar2(10),
	说明 varchar2(80),
	分类 varchar2(20))
    TABLESPACE zl9BaseItem;
Create Table 检验抗生素组(
	ID number(18),
	编码 varchar2(10),
	名称 varchar2(50),
	英文 Varchar2(50),
	简码 Varchar2(20))
    TABLESPACE zl9BaseItem;
Create Table 检验用抗生素(
	ID number(18),
	编码 varchar2(10),
	中文名 varchar2(50),
	英文名 varchar2(50),
	简码 varchar2(20),
	说明 varchar2(100),
	药敏方法 Number(1),
	WHONET码 Varchar2(10),
	用法用量1 Varchar2(30),
	血药浓度1 Varchar2(30),
	尿药浓度1 Varchar2(30),
	用法用量2 Varchar2(30),
	血药浓度2 Varchar2(30),
	尿药浓度2 Varchar2(30))
    TABLESPACE zl9BaseItem;
Create Table 检验抗生素用药(
	抗生素ID number(18),
	抗生素分组ID number(18))
    TABLESPACE zl9BaseItem;
Create Table 检验培养文字(
	编码 varchar2(10),
	名称 varchar2(100) not Null,
	简码 varchar2(10),
	说明 varchar2(80))
    TABLESPACE zl9BaseItem;
Create Table 检验评语文字(
    编码 VARCHAR2(3),
    名称 VARCHAR2(50),
    简码 VARCHAR2(10),
    说明 VARCHAR2(80),
    分类 VARCHAR2(20))
    TABLESPACE zl9BaseItem;
Create Table 检验细菌类型(
	ID number(18),
	编码 varchar2(13),
	中文名称 varchar2(40),
	英文名称 varchar2(40),
	简码 varchar2(10))
    TABLESPACE zl9BaseItem;
Create Table 检验细菌类别(
	编码 varchar2(8),
	名称 varchar2(30),
	简码 varchar2(20),
	缺省标志 NUMBER(1) DEFAULT 0)
	TABLESPACE zl9BaseItem;
Create Table 检验细菌菌属(
	编码 varchar2(8),
	名称 varchar2(30),
	简码 varchar2(20))
	TABLESPACE zl9BaseItem;
Create Table 革兰染色分类(
	编码 varchar2(8),
	名称 varchar2(30),
	简码 varchar2(20),
	缺省标志 NUMBER(1) DEFAULT 0)
	TABLESPACE zl9BaseItem;
Create Table 检验细菌(
	ID number(18),
	编码 varchar2(10),
	中文名 varchar2(100),
	英文名 varchar2(100),
	类型ID number(18),
	简码 varchar2(10),
	默认药敏 Varchar2(1),
	默认方法 Varchar2(20),
	WHONET码 Varchar2(10),
	默认结果 varchar2(200),
	细菌类别 varchar2(30),
	细菌菌属 varchar2(30),
	革兰氏分类  varchar2(30))
    TABLESPACE zl9BaseItem;
Create Table 检验细菌抗生素(
	细菌ID number(18),
	抗生素分组ID number(18),
	缺省标志 number(18))
    TABLESPACE zl9BaseItem;
Create Table 检验项目(
	诊治项目ID number(18),
	缩写 varchar2(40),
	报告代号 varchar2(10),
	项目类别 number(1),
	结果类型 number(1),
	单位 varchar2(20),
	打印类型 number(18),
	打印序号 number(18),
	计算公式 varchar2(500),
	检验方法 varchar2(40),
	合并后代码 varchar2(10),
	结果异常条件 varchar2(10),
	结果范围 Varchar2(20),
	默认值 Varchar2(200),
	警戒下限 Number(16,5),
	警戒上限 Number(16,5),
	变异报警率 Number(16,5),
	比对警示率 Number(16,5),
	比对失控率 Number(16,5),
	取值序列 Varchar2(200),
	隐私项目 Number(1),
	阳性公式 varchar2(50),
	弱阳性公式 varchar2(50),
	CutOff公式 varchar2(50),
	排列序号 Number(18),
	变异警示率 Number(16,5),
	临床意义 varchar2(4000),
	多参考 number(1))
    TABLESPACE zl9BaseItem;
Create Table 检验项目参考(
	ID     Number(18),
	项目ID number(18),
	标本类型 varchar2(20),
	性别域 number(1),
	年龄上限 number(18),
	年龄下限 number(18),
	年龄单位 varchar2(10),
	参考高值 number(21,4),
	参考低值 number(21,4),
	备注 varchar2(50),
	仪器ID number(18),
	申请科室ID Number(18),
	临床特征 Varchar2(30),
	可偏移率 Number(10,2),
	默认 number(1),
	警示上限 NUMBER(16,5),
	警示下限 NUMBER(16,5),
	复查上限 NUMBER(16,5),
	复查下限 NUMBER(16,5))
    TABLESPACE zl9BaseItem;
Create Table 检验项目取值(
	项目ID number(18),
	编码 varchar2(10),
	取值 varchar2(10),
	结果标志 number(1))
    TABLESPACE zl9BaseItem;
Create Table 检验标本形态(
	编码 varchar2(10),
	名称 varchar2(50) not Null,
	说明 varchar2(100))
    TABLESPACE zl9BaseItem;
Create Table 检验仪器(
	ID number(18),
	编码 varchar2(10),
	名称 varchar2(20),
	简码 varchar2(10),
	连接计算机 varchar2(40),
	通讯程序名 varchar2(40),
	通讯端口 varchar2(10),
	波特率 number(6),
	数据 number(2),
	停止位 number(2,1),
	校验位 varchar2(4),
	仪器类型 varchar2(50),
	仪器标志色 varchar2(10),
	使用小组ID number(18),
	质控标本号 varchar2(40),
	备注 VARCHAR2(100),
	质控周期 Number(5),
	周期单位 Varchar2(2),
	质控水平数 Number(1),
	上次质控日 Date,
	QC码 Varchar2(8),
	试剂来源 Varchar2(30),
	校准物来源 Varchar2(30),
	微生物 Number(1),
	转换日期 Date,
	转换仪器ID Number(18),
	波长 varchar2(60),
	振板频率 varchar2(30),
	振板时间 varchar2(5),
	进板方式 varchar2(30),
	空白形式 varchar2(30),
	对数质控图 Number(1),
	发送时指定杯号 Number(1))
    TABLESPACE zl9BaseItem;
Create Table 检验仪器项目(
	项目ID number(18),
	仪器ID number(18),
	通道编码 varchar2(20),
	小数位数 number(18),
	结果位数 number(18),
	缺省仪器 number(1),
	加算值 Number(16,5),
	换算比 Number(16,5),
	抗生素ID number(18),
	糖耐量项目 number(1))
    TABLESPACE zl9BaseItem;
Create Table 检验质控规则(
    ID Number(18),
    种类 Number(1),
    编码 Varchar2(3),
    名称 Varchar2(20),
    说明 Varchar2(100),
    形式 Number(1),
    多水平 Number(1),
    N Number(2),
    X Number(5, 1),
    M Number(2),
    P Number(5, 3),
    K Number(5, 1),
    H Number(5, 1))
    Tablespace zl9BaseItem;
Create Table 检验质控品(
    ID Number(18),
    仪器ID Number(18),
    名称 Varchar2(50),
    批号 Varchar2(10),
    浓度 Varchar2(30),
    水平 Number(1),
    方法 Varchar2(30),
    开始日期 Date,
    结束日期 Date,
    非定值 Number(1),
    标本号 Varchar2(40),
    试剂 varchar2(30),
    校准物 varchar2(30))
    Tablespace zl9BaseItem;
Create Table 检验质控品项目(
	质控品ID NUMBER(18),
	项目ID   NUMBER(18),
	靶值     NUMBER(18,4),
	SD       NUMBER(18,4),
	CV       NUMBER(18,4),
	项目QC码 VARCHAR2(8),
	方法QC码 VARCHAR2(8),
	方法     VARCHAR2(30),
	取值序列 VARCHAR2(500),
	序列值 VarChar2(500),
	质控取值 VarChar2(100))
    TABLESPACE zl9BaseItem;
Create Table 检验小组(
  ID NUMBER(18),
  编码 VARCHAR2(10),
  名称 VARCHAR2(50))
    TABLESPACE zl9BaseItem;
Create Table 检验小组成员(
  小组ID NUMBER(18),
  人员ID NUMBER(18),
  默认小组 NUMBER(1),
  备注   VARCHAR2(100))
    TABLESPACE zl9BaseItem;
Create Table 检验小组仪器(
  小组ID NUMBER(18),
  仪器ID NUMBER(18),
  查看   Number(1),
  更改   Number(1),
  条码输入 Number(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 诊疗检验标本(
    编码 VARCHAR2(2),
    名称 VARCHAR2(20),
    简码 VARCHAR2(8),
		适用性别 VARCHAR2(4))
    TABLESPACE zl9BaseItem;
CREATE TABLE 诊疗检验类型(
    编码 VARCHAR2(2),
    名称 VARCHAR2(20),
    简码 VARCHAR2(8),
    缺省标志 NUMBER(1),
    管码 Varchar2(2))
    TABLESPACE zl9BaseItem;
Create Table 检验试剂关系(
    ID     Number(18),
    项目ID Number(18),
    材料ID Number(18),
    仪器ID Number(18),
	数量 number(16,5),
	固定 number(1))
    TABLESPACE zl9BaseItem;
Create Table 采血管类型(
    编码 Varchar2(4),
    名称 varchar2(30),
    简码 Varchar2(10),
    添加剂 Varchar2(30),
    采血量 Varchar2(30),
    规格 Varchar2(30),
    颜色 number(10),
    材料ID number(18))
    TABLESPACE zl9BaseItem;
Create Table 检验结果描述(
    编码 VARCHAR2(3),
    名称 VARCHAR2(200),
    简码 VARCHAR2(10),
    分类 VARCHAR2(20))
    TABLESPACE zl9BaseItem;
Create Table 临床特征(
    编码 VARCHAR2(2),
    名称 VARCHAR2(30),
    简码 VARCHAR2(10))
    TABLESPACE zl9BaseItem;
Create Table 检验项目选项(
    诊疗项目ID Number(18),
    门诊仪器ID Number(18),
    门诊仪器分解 Number(1),
    住院仪器ID Number(18),
    住院仪器分解 Number(1),
    体检仪器ID Number(18),
    体检仪器分解 Number(1),    
    跟踪天数 Number(5),
    耗时标准 Number(5),
    耗时单位 Varchar2(4),
    取报告地点 Varchar2(50),
    附加说明 Varchar2(200),
    送检时限 number(4),
    急诊耗时 Number(5))
    TABLESPACE zl9BaseItem;
Create Table 仪器细菌对照(
    仪器ID Number(18),
    通道编码 Varchar2(50),
    细菌ID Number(18),
    抗生素ID Number(18))
    TABLESPACE zl9BaseItem;
Create Table 检验仪器规则(
    ID Number(18),
    上级ID Number(18),
    仪器ID Number(18),
    项目ID Number(18),
    判断 Varchar2(80),
    规则ID Number(18),
    性质 Varchar2(1),
    批范围 Number(3),
    多水平 Number(1),
    Y标记级 Number(1),
    Y规则 Varchar2(20),
    Y结束 Number(1),
    Y提示 Varchar2(500),
    N标记级 Number(1),
    N规则 Varchar2(20),
    N结束 Number(1),
    N提示 Varchar2(500),
    是否使用 Number(1))
    Tablespace zl9BaseItem;
Create Table 细菌检测方法(
    编码 VARCHAR2(2),
    名称 VARCHAR2(20),
    简码 VARCHAR2(10))
    TABLESPACE zl9BaseItem;
Create Table 细菌耐药机制(
    编码 VARCHAR2(4),
    名称 VARCHAR2(100),
    简码 VARCHAR2(20))
    TABLESPACE zl9BaseItem;
Create Table 检验申请项目(
    标本ID Number(18),
    诊疗项目ID Number(18),
    序号 Number(3),
	待转出 Number(3))
    TABLESPACE zl9BaseItem;
Create Table 检验模板目录(
    ID Number(18),
    编码 Varchar2(6),
	名称 Varchar2(30),
	简码 Varchar2(10),
    诊疗项目ID Number(18),
    说明 Varchar2(50),
    编制人 Varchar2(20),
    编制时间 Date,
    检验评语 Varchar2(100),
    检验备注 Varchar2(400))
    TABLESPACE zl9BaseItem;
Create Table 检验模板内容(
    ID Number(18),
    模板ID Number(18),
    项目ID Number(18),
    检验结果 Varchar2(60),
    细菌ID Number(18),
    培养描述 Varchar2(50))
    TABLESPACE zl9BaseItem;
Create Table 检验模板药敏(
    细菌结果ID Number(18),
    抗生素ID Number(18),
    结果 Varchar2(20),
    结果类型 Varchar2(20),
    药敏方法 Number(3))
    TABLESPACE zl9BaseItem;
Create Table 检验合并规则(
	主项目ID number(18),
	合并项目ID number(18) not null)
    TABLESPACE zl9BaseItem;
Create Table 质控检验方法(
    编码 Varchar2(6),
    名称 Varchar2(30),
    简码 Varchar2(10))
    Tablespace zl9BaseItem;
Create Table 质控试剂来源(
    编码 Varchar2(6),
    名称 Varchar2(30),
    简码 Varchar2(10),
    QC编码 Varchar2(8))
    Tablespace zl9BaseItem;
Create Table 质控报告词句(
    编码 Varchar2(3),
    名称 Varchar2(80),
    简码 Varchar2(10),
    分组 Varchar2(4))
    Tablespace zl9BaseItem;
Create Table 质控即刻法(
    N Number(3),
    N3S Number(6,2),
    N2S Number(6,2))
    Tablespace zl9BaseItem;
Create Table 质控控界系数(
    规则 Varchar2(20),
    基础 Number(1),
    N2 Number(6,2),
    N3 Number(6,2),
    N4 Number(6,2),
    N6 Number(6,2),
    N7 Number(6,2),
    N10 Number(6,2),
    N12 Number(6,2),
    N16 Number(6,2),
    N20 Number(6,2),
    行次 Number(2))
    Tablespace zl9BaseItem;
Create Table 检验仪器状态(
    仪器ID Number(18),
    项目ID Number(18),
    失控标记 Varchar2(100),
    失控日期 Date)
    Tablespace zl9BaseItem;
Create Table 检验质控均值(
    质控品ID Number(18),
    项目ID Number(18),
    期间 Varchar2(20),
    开始日期 Date,
    结束日期 Date,    
    均值 Number(18,4),
    SD Number(18,4),
    CV Number(18,4),
    设置日期 Date,
    设置人 Varchar2(20))
    Tablespace zl9BaseItem;
Create Table 检验质控范则(
    范例名 Varchar2(30),
    水平数 Number(1),
    序号 Number(18),
    上级 Number(18),
    判断 Varchar2(80),
    规则名 Varchar2(20),
    形式 Number(1),
    N Number(2),
    X Number(5, 1),
    M Number(2),
    性质 Varchar2(1),
    批范围 Number(3),
    多水平 Number(1),
    Y标记级 Number(1),
    Y规则 Varchar2(20),
    Y结束 Number(1),
    Y提示 Varchar2(500),
    N标记级 Number(1),
    N规则 Varchar2(20),
    N结束 Number(1),
    N提示 Varchar2(500))
    Tablespace zl9BaseItem;
CREATE TABLE 检验审核类别(
    编码 VARCHAR2(2),
    名称 VARCHAR2(20),
    简码 VARCHAR2(8),
    缺省标志 NUMBER(1))
    TABLESPACE zl9BaseItem;
create table 检验审核规则(
  ID NUMBER(18),
  编码  Varchar2(3),
  名称  VARCHAR2(30),
  分类   VARCHAR2(30),
  项目ID NUMBER(18),
  仪器ID NUMBER(18),
  科室ID NUMBER(18),
  病人类型 VARCHAR2(1),
  性别   VARCHAR2(4),
  年龄下限 NUMBER(9),
  年龄上限 NUMBER(9),
  年龄单位 VARCHAR2(4),
  诊断 VARCHAR2(500),
  规则   VARCHAR2(4000),
  特殊规则 VARCHAR2(4000),
  规则关系 VARCHAR2(3),
  提示信息 VARCHAR2(200),
  急诊   VARCHAR2(1),
  有效   VARCHAR2(1),
  审核   VARCHAR2(1),
  备注   VARCHAR2(200))
  Tablespace zl9BaseItem;
create table 检验细菌抗生素参考
(
  细菌ID       NUMBER(18),
  抗生素分组ID NUMBER(18),
  抗生素ID     NUMBER(18),
  药敏方法     NUMBER(1),
  参考低值     NUMBER(21,4),
  参考高值     NUMBER(21,4),
  判断方式     NUMBER(1),   ---- 1-包含参考值,0-参考值除外
  备注         VARCHAR2(500),
  低值结果     VARCHAR2(30),
  中间结果     VARCHAR2(30),
  高值结果     VARCHAR2(30)
)tablespace ZL9BASEITEM;
Create Table 检验拒收理由(
	编码 varchar2(10),
	名称 varchar2(200) not Null)
    TABLESPACE zl9BaseItem;    
 Create Table 检验分析用途(
	编码 varchar2(10),
	名称 varchar2(200) not Null)
    TABLESPACE zl9BaseItem;
Create Table 检验酶标模板(
	ID	 NUMBER(18),
	编号	 Number(3),
	名称     VARCHAR2(20),
	项目     VARCHAR2(1000),	--项目格式：项目A;项目B;...项目H 共8个项目
	内容     VARCHAR2(2000))	--内容格式: 编号1;编号2...编号12|编号1;编号2...编号12 共8行每行12个编号
    TABLESPACE zl9BaseItem;
----------------------------------------------------------------------------
--[[11.检查基础]]
----------------------------------------------------------------------------
create table RIS分院设置
(
  ID     NUMBER(18),
  医院名称   VARCHAR2(100),
  医院代码   VARCHAR2(20),
  用户名    VARCHAR2(30),
  密码     VARCHAR2(70),
  数据库服务名 VARCHAR2(30)
)
tablespace ZL9CISREC;
create table RIS启用控制
(
  ID Number(18),
  检查类型   VARCHAR2(20),
  场合   Number(1),
  部门ID   NUMBER(18),
  是否启用RIS     Number(1),
  是否启用预约     Number(1)
)Tablespace zl9BaseItem;
Create Table RIS医嘱失败记录
(
   ID Number(18),
   医嘱id NUMBER(18),
   病人来源 Number(1),
   病人ID Number(18),
   主页ID Number(5),
   挂号单号 Varchar2(8),
   发送号 Number(18),
   体检任务ID Number(18),
   体检报到号 VARCHAR2(20),
   发送类型 Number(1),
   发送时间 date,
   重发次数 Number(2)
)Tablespace ZL9CISREC;
Create Table RIS接口日志记录
(
   ID Number(18),
   时间 Date,
   站点 Varchar2(50),
   用户 Varchar2(100),
   类型 Number(1),
   标题 Varchar2(100),
   函数 Varchar2(100),
   内容 Varchar2(4000)
)Tablespace ZL9CISREC;
Create Table 医技执行房间(
       科室id Number(18),
       执行间 varchar2(20),
       简码   varchar2(20),
       当前分配 Number(1),
       检查设备 Varchar2(3),
       号码前缀 varchar2(10),
       分组ID Number(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE 诊疗检查类型(
    编码 VARCHAR2(2),
    名称 VARCHAR2(20),
    简码 VARCHAR2(8),
    建病案 NUMBER(1))
    TABLESPACE zl9BaseItem;
Create Table 诊疗项目部位(
    ID Number(18),
    项目ID Number(18),
    类型 Varchar2(20),
    部位 Varchar2(30),
    方法 Varchar2(30),
    默认 Number(1),
    上级方法 Varchar2(30))
    Tablespace zl9BaseItem;
Create Table 诊疗检查部位(
    类型 Varchar2(20),
    编码 Varchar2(4),
    名称 Varchar2(30),
    分组 Varchar2(30),
    备注 Varchar2(200),
    方法 Varchar2(1000),
    适用性别 Number(1))
    Tablespace zl9BaseItem;
Create Table 造影剂(
    编码 Varchar2(2),
    名称 Varchar2(30),
    简码 Varchar2(10))
    Tablespace zl9BaseItem;
Create Table 快捷功能信息(
    ID Number(18), 
    项目 varchar2(128),
    模块号 number(18),
    菜单分组 varchar2(128),
    分组序号 number(18) default 0,
    菜单ID Number(18),    
    菜单说明 varchar2(128),
    控制键 Number(18),
    字符键 Number(18),
    默认键 varchar2(64),
    组合名 varchar2(64)
    )
    TABLESPACE zl9BaseItem; 
Create Table 快捷功能关联
(
   ID number(18),
   快捷功能ID number(18),
   用户ID number(18),
   控制键 number(18),
   字符键 number(18),
   组合名 varchar2(64))
   TABLESPACE zl9BaseItem;
    
Create Table 影像执行分组(
       ID Number(18),
       科室id Number(18),
       组名 Varchar2(30),
       分组前缀  Varchar2(10)
       )
    TABLESPACE zl9BaseItem;
    
Create Table 影像分组关联(
       ID Number(18),
       科室ID Number(18),
       分组ID Number(18),
       诊疗项目ID Number(18)
       )
    TABLESPACE zl9BaseItem;
           
create table 影像滤镜模板
(
  ID         NUMBER(18),
  影像类型   VARCHAR2(30),
  滤镜名称   VARCHAR2(50),
  增强强度增加 NUMBER(3),
  增强强度减少 NUMBER(3),
  增强幅度增加 NUMBER(3),
  增强幅度减少 NUMBER(3),
  平滑增加     NUMBER(3),
  平滑减少     NUMBER(3))
    TABLESPACE zl9BaseItem;
CREATE TABLE 影像图像备注(
    编码 VARCHAR2(5),
    名称 VARCHAR2(100),
    简码 VARCHAR2(20),
    人员 VARCHAR2(20))
    TABLESPACE zl9BaseItem;
Create Table 影像申请常用词句(
     ID NUMBER(18),
     项目分类 VARCHAR2(30),
     词句内容 VARCHAR2(200),
     是否通用 NUMBER(1),
     科室ID NUMBER(18),  
     创建人员ID NUMBER(18))
     TABLESPACE zl9BaseItem;
create table 影像MWL部位对码
(
  ID           number(18),
  服务ID       number(18),
  PACS部位名称 varchar2(30),
  设备部位名称 varchar2(64),
  设备部位代码 varchar2(64)
)
    TABLESPACE zl9BaseItem;
Create Table 影像操作记录(
  IP地址 VARCHAR2 (15),
  类型 VARCHAR2 (20),
  启动时间 DATE)
    TABLESPACE zl9BaseItem;
Create Table 影像插件挂接
(
  ID         Number(18),
  名称       Varchar2(30),
  版本       Varchar2(30),
  路径       Varchar2(100),
  程序集     Varchar2(100),
  执行类型   Number(1),
  是否启用   Number(1),
  所属模块   Number(18)
)TableSpace zl9BaseItem;
Create Table 影像插件功能
(
  ID                Number(18),
  插件ID            Number(18),
  功能序号          Number(3),
  名称              Varchar2(30),
  方法              Varchar2(500),
  方法参数          Varchar2(4000),
  是否启用          Number(1),
  是否加入右键菜单  Number(1) default 0,
  是否加入工具栏    Number(1) default 0,
  自动执行时机      Number(5) default 0,
  VBS脚本           Varchar2(4000)
)TableSpace zl9BaseItem;
    
Create Table 影像查询方案(
       Id Number(18),
       方案名称 varchar2(30),
       方案说明 varchar2(512),
       查询语句 varchar2(1024),
       是否默认 Number(1) default 0,
       使用状态 Number(1) default 1,
       方案序号 Number(18),
       所属科室 Number(18) default 0,
	   是否启用规则 Number(1) default 0,
       是否系统查询 Number(1) default 0,
	   是否常用 Number(1),
	   所属模块 Number(18),
	   方案内容 Clob,
	   版本 Number(5)
)TABLESPACE zl9BaseItem;

Create Table 影像查询关联(
       Id Number(18),
       用户ID Number(18),
       查询方案ID Number(18),       
       是否默认 Number(1),
       是否常用 Number(1),
       所属站点 varchar2(64),
	   默认加载站点 Varchar2(512)
)TABLESPACE zl9BaseItem;  
Create Table 影像查询特性(
       Id Number(18),
       用户ID Number(18),
       查询方案ID Number(18),
       条件配置 Varchar2(4000),
       过滤配置 Varchar2(4000),
       列表配置 Varchar2(4000)
)TABLESPACE zl9BaseItem;  
Create Table 影像查询资源(
       Id Number(18),
       资源名称 Varchar2(64),
       资源类型 Number(1),
       图标 Blob
)TABLESPACE zl9BaseItem; 
Create Table 影像查询配置(
       Id Number(18),
       方案ID Number(18),
       录入项目 varchar2(30),
       录入类型 Number(1),
       默认值   varchar2(512),
       数据来源 varchar2(1024),
       录入顺序 number(18)
)TABLESPACE zl9BaseItem;
Create Table 影像诊断分类(
  编码 VARCHAR2(2),
  名称 VARCHAR2(20),
  简码 VARCHAR2(8),
  科室名称 VARCHAR2(20))
    TABLESPACE zl9BaseItem;
create table 影像流程参数
(
    ID     NUMBER(18),
    科室ID NUMBER(18),
    参数名 VARCHAR2(100),
    参数值 VARCHAR2(1000))
    TABLESPACE zl9BaseItem;
Create Table 影像检查类别(
    编码 varchar2(10),
    名称 varchar2(20),
    简码 varchar2(10),
    排列 number(3),
    最大号码 varchar2(64),
	诊疗类型 VARCHAR2(20))
    TABLESPACE zl9BaseItem;
Create Table 影像检查项目(
    诊疗项目id number(18),
    影像类别 varchar2(10),
    可行病检 number(1),
    可发胶片 number(1),
    检查准备 varchar2(50),
    报告图象 number(1))
    TABLESPACE zl9BaseItem;
Create Table 影像设备目录(
    设备号 varchar2(3),
    设备名 varchar2(100),
    类型 number(1),
    IP地址 varchar2(15),
    端口号 varchar2(5),
    本机目录 varchar2(100),
    FTP目录 varchar2(100),
    目录满 number(1),
    FTP用户名 varchar2(20),
    FTP密码 varchar2(20),
    共享目录用户名 VARCHAR2(20),
    共享目录密码 VARCHAR2(20),
    共享目录 VARCHAR2(100),
    本地AE varchar2(20),
    设备AE varchar2(20),
    状态 NUMBER(1),
    影像类别 VARCHAR2(20))
    TABLESPACE zl9BaseItem;
--影像参数
Create Table 影像颜色清单(
    序号 number(5),
    颜色 varchar2(20),
    颜色内容 varchar2(4000),
    系统方案 number(1))
    TABLESPACE zl9BaseItem;
Create Table 影像打印格式(
    编码 number(5),
    名称 varchar2(50),
    格式 varchar2(20),
    参数1 number(3),
    参数2 number(3),
    参数3 number(3),
    参数4 number(3),
    参数5 number(3),
    参数6 number(3),
    参数7 number(3))
    TABLESPACE zl9BaseItem;
Create Table 影像胶片规格(
    编码 number(5),
    名称 varchar2(20),
    胶片宽度 varchar2(20),
    胶片长度 varchar2(20),
    单位 varchar2(20))
    TABLESPACE zl9BaseItem;
Create Table 影像标注存储表(
    编号 number(5),
    VGroup varchar2(20),
    Element varchar2(20),
    VR varchar2(20),
    标注属性 varchar2(20))
    TABLESPACE zl9BaseItem;
Create Table 影像图像信息表(
    ID number(5),
    开始地址 varchar2(20),
    结束地址 varchar2(20),
    英文名称 varchar2(50),
    中文名称 varchar2(50),
    中文简称 varchar2(50),
    英文简称 varchar2(50),
    常用 number(1),
    被选用 number(1),
    位置 number(2),
    角内序号 number(2),
    可导出 number(1),
    使用计算 number(1))
    TABLESPACE zl9BaseItem;
Create Table 影像界面参数表(
    人员ID number(18),
    正常图像边框颜色 varchar2(20),
    正常图像边框线型 varchar2(5),
    正常图像边框线宽 varchar2(5),
    选中图像边框颜色 varchar2(20),
    选中序列边框颜色 varchar2(20),
    选中图像边框线型 varchar2(5),
    选中图像边框线宽 varchar2(5),
    图像标记颜色 varchar2(20),
    图像标记大小 varchar2(5),
    标注选择句柄颜色 varchar2(20),
    标注选择句柄大小 varchar2(5),
    定位线颜色 varchar2(20),
    定位线线型 varchar2(5),
    定位线间距 varchar2(5),
    序列间间隔 varchar2(5),
    横向最大序列 varchar2(5),
    纵向最大序列 varchar2(5),
    图像间距 varchar2(5),
    显示多余边框 varchar2(5),
    背景颜色 varchar2(20),
    程序背景颜色 varchar2(20),
    标注正常颜色 varchar2(20),
    标注正常线型 varchar2(5),
    标注正常线宽 varchar2(5),
    标注选中颜色 varchar2(20),
    标注选中线型 varchar2(5),
    标注选中线宽 varchar2(5),
    标注文字大小 varchar2(5),
    测量显示面积 varchar2(5),
    测量显示平均值 varchar2(5),
    测量显示均方差 varchar2(5),
    测量显示中文 varchar2(20),
    文字X方向偏移 varchar2(5),
    文字Y方向偏移 varchar2(5),
    文字随图像缩放 varchar2(5),
    显示体位标记 varchar2(20),
    中文体位标记 varchar2(20),
    显示标尺 varchar2(5),
    标尺左右边距 varchar2(5),
    标尺上下边距 varchar2(5),
    标尺宽度 varchar2(5),
    标尺高度 varchar2(5),
    标尺线宽 varchar2(5),
    标尺颜色 varchar2(20),
    窗宽窗位位置 varchar2(5),
    鼠标穿梭步长 varchar2(5),
    鼠标漫游步长 varchar2(5),
    鼠标调窗步长 varchar2(5),
    鼠标缩放步长 varchar2(5),
    病人信息上下边距 varchar2(5),
    病人信息左右边距 varchar2(5),
    病人信息颜色 varchar2(20),
    病人信息显示最小值 varchar2(5),
    病人信息随图像缩放 varchar2(5),
    病人信息字体 varchar2(50),
    病人信息题头 varchar2(20),
    直接照相 varchar2(5),
    工具栏图标大小 varchar2(5),
    工具栏位置 varchar2(5),
    工具栏显示 varchar2(5),
    状态栏字体大小 varchar2(5),
    正常血管阈值 varchar2(10),
    狭窄血管阈值 varchar2(10),
    血管壁宽度 varchar2(10),
    测量显示周长 varchar2(10),
    测量显示最大值 varchar2(10),
    测量显示最小值 varchar2(10),
    鼠标滚轮操作 varchar2(2),
    显示打印标记 VARCHAR2(2))
    TABLESPACE zl9BaseItem;
Create Table 影像鼠标按钮分配(
    人员ID number(18),
    直线 varchar2(200),
    矩形 varchar2(200),
    椭圆 varchar2(200),
    箭头 varchar2(200),
    多边形 varchar2(200),
    多边线 varchar2(200),
    角度 varchar2(200),
    文字 varchar2(200),
    穿梭定位 varchar2(200),
    窗宽窗位 varchar2(200),
    漫游 varchar2(200),
    缩放 varchar2(200),
    裁剪_标注调整 varchar2(200),
    自适应调窗 varchar2(200),
    三维鼠标 varchar2(200),
    画标注 varchar2(200))
    TABLESPACE zl9BaseItem;
Create Table 影像图像消隐表(
    ID  number(18),
    人员ID number(18),
    影像类型 varchar2(30),
    消隐类型 varchar2(30),
    圆心X number(10),
    圆心Y number(10),
    圆形半径 number(10),
    矩形左边界 number(10),
    矩形右边界 number(10),
    矩形上边界 number(10),
    矩形下边界 number(10),
    多边形顶点 varchar2(50),
    消隐颜色 number(20))
    TABLESPACE zl9BaseItem;
Create Table 影像预设窗宽窗位(
    ID  number(18),
    人员ID  NUMBER(18),
    影像类型 varchar2(30),
    快捷键 number(5),
    窗口名称 varchar2(50),
    窗口英文名 varchar2(60),
    窗宽 number(10),
    窗位 number(10),
    是否默认 number(5))
    TABLESPACE zl9BaseItem;
Create Table 影像屏幕布局(
    ID  number(18),
    人员ID number(18),
    影像类型 varchar2(30),
    自动序列布局 number(5),
    自动图像布局 number(5),
    序列行数 number(5),
    序列列数 number(5),
    图像行数 number(5),
    图像列数 number(5),
    自动反白 number(1),
    显示病人信息 number(1),
    选择定位线 number(1),
    选择序列同步 number(1),
    插值模式 number(1),
	图像排序 number(1))
    TABLESPACE zl9BaseItem;
Create Table 影像打印机设置(
    ID  number(18),
    打印机名 varchar2(50),
    IP地址 varchar2(18),
    端口号 number(5),
    AE名称 varchar2(50),
    打印格式 varchar2(50),
    优先级 varchar2(30),
    打印份数 number(5),
    介质 varchar2(30),
    方向 varchar2(30),
    胶片规格 varchar2(30),
    选用片盒 varchar2(30),
    分辨率 varchar2(30),
    放大模式 varchar2(30),
    平滑模式 varchar2(30),
    修整 varchar2(30),
    最小密度 varchar2(30),
    最大密度 varchar2(30),
    空白密度 varchar2(30),
    边框密度 varchar2(30),
    极性 varchar2(30),
    图像位数 number(5),
    用户AE名称 VARCHAR2(50),
    图像边框宽度 NUMBER(2),
    图片分辨率 number(3))
    TABLESPACE zl9BaseItem;
Create Table 影像胶片打印字体(
    影像类别 varchar2(50),
    字体大小 number(5),
    是否随图像缩放 number(5),
    体位标注字体大小 NUMBER(5),
    体位标注随图像缩放 NUMBER(5),
    字体反色 NUMBER(1) default 0,
    字体阴影 NUMBER(1) default 0,
    字体背景透明 NUMBER(1) default 1)
    TABLESPACE zl9BaseItem;
Create Table 服用造影剂(
       医嘱ID         NUMBER(18),
       造影剂         Varchar2(30),
       用量           Varchar2(30),
       浓度           Varchar2(30))
    TABLESPACE zl9CisRec
    PCTFREE 5;
create table 影像DICOM服务对
(
  服务ID   number(18),
  服务名    varchar2(20),
  设备号   varchar2(3),
  服务功能 varchar2(20),
  PACS角色 varchar2(3),
  PACSIP地址    varchar2(15),
  PACSAE名称   varchar2(20),
  PACS端口 varchar2(5),
  设备IP地址    varchar2(15),
  设备AE名称   varchar2(20),
  设备端口 varchar2(5)
  )
  TABLESPACE zl9BaseItem;
create table 影像DICOM服务参数
(
  服务参数ID  number(18),
  服务ID   number(18),
  参数名称  varchar2(100),
  参数值    varchar2(1000))
  TABLESPACE zl9BaseItem;
create table 影像MWL结果集
(
  ID         number(18),
  服务ID    number(18),
  组号       varchar2(4),
  元素号     varchar2(4),
  上级ID	 number(18),
  中文标题   varchar2(50),
  英文标题   varchar2(50),
  数据值     varchar2(100),
  是否嵌套数据 number(1),
  是否递增     number(1),
  值类型	   varchar2(10),
  选中         number(1),
  元素类型     varchar2(5),
  强制结果值   varchar2(100),
  默认值       varchar2(100),
  默认选中     number(1),
  默认强制结果值 varchar2(100))
  TABLESPACE zl9BaseItem;
create table 影像接入设备
(
  接入ID   number(18),
  IP地址   varchar2(20),
  设备名称 varchar2(100),
  影像类别 varchar2(20))
  TABLESPACE zl9BaseItem;
    
Create Table 影像收藏类别(
    ID   NUMBER(18),      
    上级ID   NUMBER(18),    
    收藏类别  Varchar2(64),   
    是否共享 NUMBER(1),        
    创建人ID   NUMBER(18),    
    创建时间 Date             
)TABLESPACE zl9BaseItem;
--病理基础
Create Table 影像病理类别
(
    编码 varchar2(10),
    名称 varchar2(20),
    简码 varchar2(10),
    前导标记 varchar2(1),
    最大号码 number(18))
    TABLESPACE zl9BaseItem;
Create Table 病理类型(
    编码 VARCHAR2(2),
    名称 VARCHAR2(10),
    简码 VARCHAR2(6),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;
Create Table 病理号码规则(
       ID Number(5),
       类型 Number(1) default -1,
       前缀 Varchar2(5),
       年   Number(1) default 0,
       月   Number(1) default 0,
       日   Number(1) default 0,
       年份位数 Number(2) default 4,
       序号位数 Number(2) default 5,
       起始数   Number(18) default 1,
       相同规则 Number(1) default 0,
	   名称     Varchar2(30) 
)TABLESPACE zl9BaseItem;
Create Table 病理号码记录(
       ID Number(5),
       类型 Number(1) default -1,
       年   Number(4) default 0,
       月   Number(2) default 0,
       日   Number(2) default 0,
       当前序号 Number(18) default 1,
       号码规则ID  Number(5)
)TABLESPACE zl9BaseItem;
Create Table 病理检查标本(
       ID Number(18),
       标本名称 Varchar2(64),
       标本部位 Varchar2(20),
       标本类型 Number(1) default 0,
       默认标本量   Varchar2(20),
       默认制片数 Number(2) default 1,
       简码    varchar2(10),
       备注     Varchar2(255)       
) TABLESPACE zl9BaseItem;
Create Table 病理套餐信息(
    套餐ID Number(18), 
    套餐名称 VARCHAR2(64),
    套餐类别 VARCHAR2(64),
    套餐说明 VARCHAR2(1024),
    创建人 VARCHAR2(64),
    创建时间 Date)
    TABLESPACE zl9BaseItem;  
    
Create Table 病理套餐关联(
    ID Number(18),    
    套餐ID Number(18), 
    抗体ID Number(18),
    抗体顺序 Number(5))
    TABLESPACE zl9BaseItem;  
    
    
Create Table 病理档案分类(
       ID Number(18),
       分类名称 Varchar2(64),
       材料类型 Number(1),
       报表名称 Varchar2(30),
       创建人 Varchar2(64),
       创建时间 date,
       备注 Varchar2(1024)
  )TABLESPACE zl9BaseItem;    
--Pacs报告编辑器
Create Table 影像报告元素分类(
       ID               Raw(16),
	   上级ID		  Raw(16),	
       编码             Varchar2(6),
       名称             Varchar2(80),
       说明             Varchar2(200),
       最后编辑时间     Date
)TABLESPACE zlPacsBaseTab;
Create Table 影像报告值域清单(
       ID                 Raw(16),
       分类ID             Raw(16),
       编码               Varchar2(20),
       名称               Varchar2(80),
       说明               Varchar2(200),
       数据类型           Varchar2(2),    
       值域种类           Number(1), 
       值域描述           Xmltype,
       最后编辑时间       Date
)TABLESPACE zlPacsBaseTab;
Create Table 影像报告元素清单(
       ID                  Raw(16),
       分类ID              Raw(16),
       编码                Varchar2(30),
       名称                Varchar2(80),
	   前缀                Varchar2(80),
	   后缀                Varchar2(80),
       说明                Varchar2(200),
       数据类型            Varchar2(2),
       数值形态            Varchar2(3),
       最小长度            Number(8),
       最大长度            Number(8),
       最小小数位          Number(8),
       最大小数位          Number(8),
       计量单位            Varchar2(20),
       扩展描述            Xmltype,
       值域ID              Raw(16),
       值域种类            Number(1), 
       最后编辑时间        Date
)TABLESPACE zlPacsBaseTab;
Create Table 影像报告计量单位(
       编码             Varchar2(20),
       名称             Varchar2(80),
       说明             Varchar2(200),
       前缀             Varchar2(10)
)TABLESPACE zlPacsBaseTab;
Create Table 影像报告组句清单(
       ID               Raw(16),
       编码             Varchar2(30),
       名称             Varchar2(80),
       说明             Varchar2(200),
       分组             Varchar2(60),
       多组             Number(1),
       组成             Xmltype,
       编辑人           Varchar2(100),         
       最后编辑时间     Date
)TABLESPACE zlPacsBaseTab;
Create Table 影像报告片段清单(
       ID               Raw(16),
       上级ID           Raw(16),    
       编码             Varchar2(10),
       名称             Varchar2(80),
       说明             Varchar2(200),
       节点类型         Number(1),
       组成             Xmltype,
	   适应条件			Xmltype,
       学科             Varchar2(200),
       标签             Varchar2(200),
       是否私有         Number(1),          
       作者             Varchar2(100),         
       最后编辑时间     Date
)TABLESPACE zlPacsBaseTab;
Create Table 影像报告预备提纲(
       ID                    Raw(16),
       编码                  Varchar2(3),
       名称                  Varchar2(80),
       说明                  Varchar2(200),
       最后编辑时间          Date
)TABLESPACE zlPacsBaseTab;
Create Table 影像报告原型清单(
       ID                    Raw(16),
       种类                  Varchar2(2),   
       编码                  Varchar2(30),
       名称                  Varchar2(80),
	   设备号				  Varchar2(3),
       说明                  Varchar2(200),
	   分组					 Varchar2(60),
       内容                  Xmltype,
       可否重置页面          Number(1),
       可否重置格式          Number(1),
	   可否书写多份          Number(1),
	   是否禁用			     Number(1),
       专用插件              Xmltype,
       控制选项              Xmltype,
	   词句加载时机			Number(1),
	   插件加载时机			Number(1),
       创建人                Varchar2(100),
       创建时间              Date,
       修改人                Varchar2(100),  
       修改时间              Date,
	   使用次数				 Number(18)
)TABLESPACE zlPacsBaseTab;
Create Table 影像报告原型应用(
       诊疗项目ID         Number(18),
       应用场合           Number(3),
       报告原型ID         Raw(16)
)TABLESPACE zlPacsBaseTab;
Create Table 影像报告原型片段(
       原型ID      Raw(16),
       片段ID      Raw(16)
)TABLESPACE zlPacsBaseTab;
Create Table 影像报告事件(
       ID              Raw(16),
       种类            Number(1),
       原型ID          Raw(16), 
       编号            Number(8),
       名称            Varchar2(80),
       说明            Varchar2(200),
       元素IID         Varchar2(36),
       扩展标记        Varchar2(200)
)TABLESPACE zlPacsBaseTab;
Create Table 影像报告动作(
       ID               Raw(16),
       原型ID           Raw(16), 
       事件ID           Raw(16),       
       动作类型         Number(1),
       名称             Varchar2(80),
       说明             Varchar2(200),
       可否手工执行     Number(1),
       序号             Number(8),
       内容             Xmltype
)TABLESPACE zlPacsBaseTab;
Create Table 影像报告插件(
       ID               Raw(16),
       编码             Varchar2(30), 
       名称             Varchar2(80),       
       说明             Varchar2(200),       
       显示样式         Number(1),       
       种类             Number(1),
       类名             Varchar2(100),   
       库名             Varchar2(100),   
       是否禁用         Number(1)
)TABLESPACE zlPacsBaseTab;
Create Table 影像报告种类(
       编码             Varchar2(2), 
       名称             Varchar2(80),       
       说明             Varchar2(200)
)TABLESPACE zlPacsBaseTab;
Create Table 影像报告范文清单(
       ID               Raw(16),
       原型ID           Raw(16),   
       编号             Number(8),
       名称             Varchar2(80),
       说明             Varchar2(200),
       内容             Xmltype,
       学科             Varchar2(200),
       标签             Varchar2(200),
       是否私有         Number(1),
       作者             Varchar2(100),  
       最后编辑时间     Date
)TABLESPACE zlPacsBaseTab;
--由于影像报告记录表涉及到xml处理，且11g和10g的xml选项有所区别，因此需要包含11g和10g的表结构创建脚本，11g的创建脚本放在10g之前。
--11g的影像报告记录表结构创建脚本
Create Table 影像报告记录(
			   ID             Raw(16),
			   医嘱ID		  Number(18),
			   原型ID         Raw(16),   
			   文档标题       Varchar2(60),
			   设备号		   Varchar2(3),
			   报告内容       Xmltype,
			   报告状态       Number(1),
			   待处理人	  Varchar2(100),
			   创建时间       Date,
			   创建人         Varchar2(100),
			   最后编辑时间   Date,
			   最后编辑人     Varchar2(100),
			   锁定人         Varchar2(100),       
			   诊断意见       Varchar2(1024),
			   检查部位       Varchar2(500),
			   编辑日志       Xmltype,
			   报告质量   	NUMBER(1),
			   结果阳性		NUMBER(1),
			   报告打印   NUMBER(1) default 0,
               报告发放   NUMBER(1) default 0,
               报告发放人  VARCHAR2(100),
               最后审核人  VARCHAR2(100),
               最后审核时间 DATE,
			   记录人    VARCHAR2(100),
			   待转出		  Number(3)
		)TABLESPACE zlPacsBizTab
		 XMLTYPE COLUMN 报告内容 STORE AS SECUREFILE BINARY XML(
			TABLESPACE zlPacsBizXml 
			DISABLE STORAGE IN ROW  
			NOCACHE LOGGING 
			COMPRESS HIGH)
		 XMLTYPE COLUMN 编辑日志 STORE AS SECUREFILE BINARY XML(
			TABLESPACE zlPacsBizXml 
			DISABLE STORAGE IN ROW  
			NOCACHE LOGGING 
			COMPRESS HIGH);
--10g的影像报告记录表结构创建脚本
Create Table 影像报告记录(
			   ID             Raw(16),
			   医嘱ID		  Number(18),
			   原型ID         Raw(16),   
			   文档标题       Varchar2(60),
			   设备号		   Varchar2(3),
			   报告内容       Xmltype,
			   报告状态       Number(1),
			   待处理人	  Varchar2(100),
			   创建时间       Date,
			   创建人         Varchar2(100),
			   最后编辑时间   Date,
			   最后编辑人     Varchar2(100),
			   锁定人         Varchar2(100),       
			   诊断意见       Varchar2(1024),
			   检查部位       Varchar2(500),
			   编辑日志       Xmltype,			   
			   报告质量   	NUMBER(1),
			   结果阳性		NUMBER(1),
			   报告打印   NUMBER(1) default 0,
			   报告发放   NUMBER(1) default 0,
               报告发放人  VARCHAR2(100),
               最后审核人  VARCHAR2(100),
               最后审核时间 DATE,
			   记录人    VARCHAR2(100),
			   待转出		  Number(3)
		)TABLESPACE zlPacsBizTab
		 XMLTYPE COLUMN 报告内容 STORE AS Clob(
			TABLESPACE zlPacsBizXml 
			DISABLE STORAGE IN ROW  
			NOCACHE LOGGING)
		 XMLTYPE COLUMN 编辑日志 STORE AS Clob(
			TABLESPACE zlPacsBizXml 
			DISABLE STORAGE IN ROW  
			NOCACHE LOGGING);
Create Table 影像参数说明(
       ID               Raw(16),
	   PID				Raw(16),
	   系统			  Number(5),		 					
       模块             Varchar2(100),   
       分组             Varchar2(60),
	   参数序号	        Number(18),   	
       参数名           Varchar2(100),
       默认值           Varchar2(4000),
       参数级别         Number(1),
       取值范围         Varchar2(4000),
	   启用条件	        Varchar2(255),
       说明             Varchar2(255)
)TABLESPACE zlPacsBaseTab;
Create Table 影像参数取值(
       ID              Raw(16),
       参数ID          Raw(16),   
       参数标识        Varchar2(100),
       参数值          Varchar2(4000)
)TABLESPACE zlPacsBaseTab;
Create Table 影像字典清单(
       ID            Raw(16),
       分组          Varchar2(60),
       编号          Varchar2(20),
       名称	     Varchar2(80),
       说明          Varchar2(500),
       是否系统      Number(1)
)TABLESPACE zlPacsBaseTab;
Create Table 影像字典内容(
       字典ID        Raw(16),
       编号          Varchar2(20),
       名称	         Varchar2(80),
       简码          Varchar2(10),
       说明          Varchar2(500)
)TABLESPACE zlPacsBaseTab;
Create Table 影像报告操作记录(
       	   ID			   Raw(16),
	   报告ID		   Raw(16),
	   医嘱ID		   Number(18),
	   文档标题 		   VARCHAR2(60),
	   操作类型                Number(1),
	   操作人		   Varchar2(100),
	   操作时间	           Date,
	   作废人		   Varchar2(100),
	   作废时间	           Date,
	   作废说明	           Varchar2(500),
	   待转出		   Number(3)
)TABLESPACE zlPacsBaseTab;
----------------------------------------------------------------------------
--[[12.医保业务]]
----------------------------------------------------------------------------
CREATE TABLE 就诊登记记录(
	险类 NUMBER(18),
	病人ID NUMBER(18),
	主页ID NUMBER(18),
	就诊时间 DATE ,
	状态 NUMBER(2),		--1-就诊中;0-未就诊
	医疗类别 VARCHAR2(3),
	帐户余额 NUMBER(16,5),
	病种ID NUMBER(18),
	病种名称 VARCHAR2(100),
	并发症 VARCHAR2(200),
	IC卡信息 VARCHAR2(200),
	HIS流水号 VARCHAR2(30),
	YB流水号 VARCHAR2(30),
	记录ID NUMBER(18),	--结帐ID，门诊可用此字段来关联，住院不必
	备注 VARCHAR2(200),
	确认 NUMBER(1))
    TABLESPACE ZL9BASEITEM;
CREATE TABLE 医保病人关联表(
	险类 NUMBER (3),
	中心 NUMBER (5),
	医保号 VARCHAR2 (30),
	病人ID NUMBER (18),
	就诊时间 DATE,
	标志 NUMBER (1) DEFAULT 0)
    TABLESPACE ZL9BASEITEM;
CREATE TABLE 结算日志(
	性质 NUMBER(1) DEFAULT 0,	--1-门诊
	NO VARCHAR2(20),
	医保号 VARCHAR2 (50),
	姓名 VARCHAR2(100),
	费用总额 NUMBER(16,5),
	结算时间 DATE )
    TABLESPACE ZL9BASEITEM;
Create Table 医保病人档案(
    险类 NUMBER(3),
    中心 NUMBER(5),
    卡号 VARCHAR2(25),
	医保号 VARCHAR2(30),
    密码 VARCHAR2(8),
    人员身份 VARCHAR2(8),
    单位编码 VARCHAR2(15),
    顺序号 VARCHAR2(20),
	退休证号 VARCHAR2(26),
    帐户余额 NUMBER(16,5),
    当前状态 NUMBER(2),
    病种ID NUMBER(18),
    在职 NUMBER(1),
    年龄段 NUMBER(3),
    灰度级 VARCHAR2(1),
	就诊时间 DATE)
    TABLESPACE zl9Patient
    PCTFREE 5;
Create Table 帐户年度信息(
	病人ID NUMBER(18),
	险类 NUMBER(3),
	年度 NUMBER(4),
	帐户增加累计 NUMBER(16,5),
	帐户支出累计 NUMBER(16,5),
	进入统筹累计 NUMBER(16,5),
	统筹报销累计 NUMBER(16,5),
	住院次数累计 NUMBER(3),
	本次起付线   NUMBER(16,5),
	基本统筹限额 NUMBER(16,5),
	大额统筹限额 NUMBER(16,5),
	起付线累计   NUMBER(16,5),
	大额统筹累计 NUMBER(16,5),
	封销信息  VARCHAR2(100))
    TABLESPACE zl9Patient
    PCTFREE 5;
Create Table 保险结算记录(
    性质 NUMBER(2),
    记录ID NUMBER(18),
    冲销ID NUMBER (18),
    冲销时间 DATE,
    险类 NUMBER(3),
    病人ID NUMBER(18),
    年度 NUMBER(4),
    帐户累计增加 NUMBER(16,5),
    帐户累计支出 NUMBER(16,5),
    累计进入统筹 NUMBER(16,5),
    累计统筹报销 NUMBER(16,5),
    住院次数 NUMBER(5),
    起付线 NUMBER(16,5),
    封顶线 NUMBER(16,5),
    实际起付线 NUMBER(16,5),
    发生费用金额 NUMBER(16,5),
    全自付金额 NUMBER(16,5),
    首先自付金额 NUMBER(16,5),
    进入统筹金额 NUMBER(16,5),
    统筹报销金额 NUMBER(16,5),
    大病自付金额 NUMBER(16,5),
    超限自付金额 NUMBER(16,5),
    个人帐户支付 NUMBER(16,5),
    支付顺序号 VARCHAR2(20),
    中途结帐 NUMBER(1),
    主页ID NUMBER(5),
    是否上传 NUMBER(1),
    备注 VARCHAR2(500),
    校正 NUMBER(1),
    就诊流水号 VARCHAR2(30),
    结算时间 DATE,
    工作站 VARCHAR2(50),
    版本号 VARCHAR2(15),
    医疗类别 VARCHAR2(3),
    病种ID NUMBER(18),
    病种名称 VARCHAR2(100),
    并发症 VARCHAR2(200),
    确认 NUMBER(1),
    序号 Number(18),
    卡类别ID number(18),
	NO Varchar2(8))
    TABLESPACE zl9Expense PCTFREE 5 
;
Create Table 保险结算计算(
	结帐ID NUMBER(18),
	档次 NUMBER(3),
	进入统筹金额 NUMBER(16,5),
	统筹报销金额 NUMBER(16,5),
	比例 NUMBER(3))
    TABLESPACE zl9Expense
    PCTFREE 5;
Create Table 保险结算明细(
	结帐ID number(18),
	结算方式 varchar2(20),
	金额 number(16,5),
	标志 NUMBER(1) DEFAULT 0)
	TABLESPACE zl9Expense
	PCTFREE 5;
Create Table 保险模拟结算(
    病人ID Number(18),
    主页ID Number(5),
    结算方式 Varchar2(20),
    金额 Number(16,5),
    更新时间 Date)
    TABLESPACE zl9Expense
    PCTFREE 5;
----------------------------------------------------------------------------
--[[13.病人病案业务]]
----------------------------------------------------------------------------
CREATE TABLE 不良行为分类(
  编码 varchar2(2),
  名称  varchar2(20),
  简码 varchar2(10),
  是否固定 number(1),
  有效期限 number(5) 
  )TABLESPACE zl9BaseItem;
CREATE TABLE 常用不良行为原因(
  编码 varchar2(5),
  名称  varchar2(50),
  简码 varchar2(10),
  是否固定 number(1)
  )TABLESPACE zl9BaseItem;
CREATE TABLE 不良行为控制(
  应用场合 varchar2(10),
  行为类别  varchar2(20),
  预约方式 varchar2(20),
  序号 Number(5),
  控制规则 varchar2(50),
  控制方式 number(1)
  )TABLESPACE zl9BaseItem;
CREATE TABLE 病人不良记录(
  ID  number(18),
  行为类别  varchar2(20),
  病人ID  number(18),
  发生时间 date,
  加入原因 varchar2(50),
  加入说明 varchar2(500),
  加入时间 date,
  附加信息 varchar2(50),
  登记人 varchar2(20),
  撤消原因 varchar2(500),
  撤消人 varchar2(20),
  撤消时间 date
  )TABLESPACE zl9Patient;
CREATE TABLE 病人身份关联(
	关联ID Number(18),
	病人ID Number(18),
  操作人员 varchar2(20),
  操作时间 date
  )TABLESPACE ZL9PATIENT;
Create Table 病人自动计算(
    Id Number(18),
    病人ID number(18) Not Null,
    主页ID number(5) Not Null,
    性质 Number(2),     
    开始时间 Date,
	开始原因 number(2),
    附加床位 number(1),
    科室ID  Number(18),
    病区ID  number(18),
    护理等级id number(18),
    床位等级id number(18),
    床号 VARCHAR2(10),
    终止人员 varchar2(20),
	终止原因 number(2),
    终止时间 Date,
    操作员编号 varchar2(6),
    操作员姓名 varchar2(20),
    上次计算时间 Date)
    TABLESPACE zl9Patient
    initrans 20;
Create Table 病人信息(
    病人ID NUMBER(18),
    主页ID NUMBER(5),
    门诊号 NUMBER(18),
    住院号 NUMBER(18),
    就诊卡号 VARCHAR2(50),
    卡验证码 VARCHAR2(50),
    费别 VARCHAR2(10),
    医疗付款方式 Varchar2(20),
    姓名 VARCHAR2(100),
    性别 VARCHAR2(4),
    年龄 varchar2(20),
    出生日期 Date,
    出生地点 VARCHAR2(100),
    身份证号 VARCHAR2(18),
    其他证件 VARCHAR2(20),
    身份 VARCHAR2(10),
    职业 VARCHAR2(80),
    民族 VARCHAR2(20),
    国籍 VARCHAR2(30),
    籍贯 VARCHAR2(100),
    区域 VARCHAR2(100),
    学历 VARCHAR2(10),
    婚姻状况 VARCHAR2(4),
    家庭地址 VARCHAR2(100),
    家庭电话 VARCHAR2(20),
    家庭地址邮编 VARCHAR2(6),
    监护人 VARCHAR2(64),
    联系人姓名 VARCHAR2(64),
    联系人关系 VARCHAR2(30),
    联系人地址 VARCHAR2(100),
    联系人电话 VARCHAR2(20),
    户口地址 VARCHAR2(100),
    户口地址邮编 VARCHAR2(6),
    Email Varchar2(30),
    QQ Varchar2(12),
    合同单位id NUMBER(18),
    工作单位 VARCHAR2(100),
    单位电话 VARCHAR2(20),
    单位邮编 VARCHAR2(6),
    单位开户行 VARCHAR2(50),
    单位帐号 Varchar2(50),
    担保人 VARCHAR2(100),
    担保额 NUMBER(16,5),
    担保性质 NUMBER(1),
    就诊时间 Date,
    就诊状态 Number(1) Default 0,
    就诊诊室 Varchar2(20),
    住院次数 number(3),
    当前科室id number(18),
    当前病区id number(18),
    当前床号 VARCHAR2(10),
    入院时间 DATE,
    出院时间 Date,
    在院 number(1),
    IC卡号 varchar2(50),
    健康号 varchar2(50),
    医保号 VARCHAR2(30),
    险类 NUMBER(3),
    查询密码 Varchar2(50),
    登记时间 Date,
    停用时间 Date,
    锁定 Number(1),
    联系人身份证号 varchar2(18),
    病人类型 Varchar2(50),
    手机号 Varchar2(50),
    单位地址 varchar2(100))
    TABLESPACE zl9Patient initrans 20 
;
Create Table 病人家属(
  病人ID Number(18),
  家属ID Number(18),
  关系 Varchar2(30),
  登记人 varchar2(100),
  登记时间 Date,
  撤档人  varchar2(100),
  撤档时间 date)
  Tablespace zl9Patient ;
Create Table 在院病人
(
	病区ID NUMBER(18),
	科室ID NUMBER(18),
	病人ID NUMBER(18),
	主页ID NUMBER(5)
)
TABLESPACE zl9Patient
Initrans 20;
CREATE TABLE 病人信息变动(
	病人ID Number(18),
	变动项目 VARCHAR2(10) not NULL,
	原信息 VARCHAR2(100),
	新信息 VARCHAR2(100),
	变动时间 DATE,
	变动人 Varchar2(20),
	变动模块 Varchar2(100),
	说明 varchar2(4000)
	)TABLESPACE zl9Patient;
Create Table 病人医疗卡信息(
    病人ID number(18),
    卡类别ID Number(18),
    卡号 Varchar2(50),
    密码 Varchar2(50),
    状态 Number(2) DEFAULT 0,
    挂失时间 Date,
    挂失方式 Varchar2(20),
    挂失人 varchar2(20),
    发卡日期 Date,
    发卡人 Varchar2(20),
    终止使用时间 Date,
    二维码 varchar2(200))
    TABLESPACE zl9Patient;
Create Table 病人医疗卡属性(
	病人ID Number(18),
	卡类别ID Number(18),
	卡号 varchar2(50),
	信息名 Varchar2(20),
	信息值 Varchar2(100))
	Tablespace zl9Patient ;
Create Table 病人信息从表(
	病人ID Number(18),
	就诊ID Number(18),
	信息名 Varchar2(20),
	信息值 Varchar2(100))
	Tablespace zl9Patient ;
Create Table 病人医疗卡变动(
    ID Number(18),
    病人ID Number(18),
    卡类别ID Number(18),
    卡号 VarChar2(50),
    变动ID Number(18),
    变动类别 Number(3),
    原密码 VARCHAR2(50),
    现密码 VARCHAR2(50),
    变动时间 Date,
    变动原因 Varchar2(100),
    挂失方式 Varchar2(30),
    操作员姓名 Varchar2(20),
    登记时间 Date,
    卡费 number(16,5),
    病历费 number(16,5),
    费用单号 varchar2(8),
    终止使用时间 date)
    TABLESPACE zl9Patient 
;
Create Table 病人照片(
    病人ID NUMBER(18),
    照片 blob)
    TABLESPACE zl9Patient
    PCTFREE 20;
    --照片新增时需要使用较多的预留空间
 
Create Table 病人担保记录(
    病人ID      NUMBER(18),
	主页ID		NUMBER(5),
    担保人      VARCHAR2(64),
    担保额      NUMBER(16,5),
	担保性质    NUMBER(1),
	担保原因	   VARCHAR2(50),
	累计号      NUMBER(5),
    操作员编号  VARCHAR2(6),
    操作员姓名  VARCHAR2(20),
    登记时间    Date,
	到期时间	Date,
	删除标志  NUMBER(1) default 1,
	删除操作员编号 VARCHAR2(6),
	删除操作员姓名 VARCHAR2(20),
	删除时间  Date)
    TABLESPACE zl9Patient;
Create Table 病人合并记录(
    病人ID NUMBER(18),
    原信息 VARCHAR2(1000),
    合并原因 VARCHAR2(250),
    操作员姓名 VARCHAR2(20),
    合并时间 Date,
    原病人id NUMBER(18))
    TABLESPACE zl9Patient
;
CREATE TABLE 病人社区信息(
		病人ID NUMBER(18),
		社区 NUMBER(5),
		社区号 VARCHAR2(20),
		标志 NUMBER(1),
		就诊类型 NUMBER(1),
		就诊时间 DATE)
		TABLESPACE zl9Patient;
Create Table 门诊病案记录(
    病人ID NUMBER(18),
    病案号 NUMBER(18),
    建立日期 Date,
    病案类别 VARCHAR2(10),
    存储状态 VARCHAR2(4),
    存放位置 VARCHAR2(20))
    TABLESPACE zl9Patient
    initrans 20;
Create Table 住院病案记录(
    病人ID NUMBER(18),
    主页ID NUMBER(5),
    病案号 VARCHAR2(20),
		档案号 VARCHAR2(20),
    建立日期 Date,
    病案类别 VARCHAR2(10),
    存储状态 VARCHAR2(8),
    存放位置 VARCHAR2(20))
    TABLESPACE zl9Patient
    PCTFREE 5;
Create Table 病案主页(
    病人ID NUMBER(18),
    主页ID NUMBER(5),
    住院号 NUMBER(18),
    留观号 number(18),
    病人性质 NUMBER(1),
    医疗付款方式 VARCHAR2(20),
    费别 VARCHAR2(10),
    再入院 NUMBER(1),
    入院病区ID NUMBER(18),
    入院科室id NUMBER(18),
    医疗小组id NUMBER(18),
    入院日期 Date,
    入院病况 VARCHAR2(20),
    入院方式 VARCHAR2(8),
    入院属性 VARCHAR2(8),
    二级院转入 VARCHAR2(1),
    住院目的 VARCHAR2(10),
    入院病床 VARCHAR2(10),
    是否陪伴 NUMBER(1),
    当前病况 VARCHAR2(20),
    当前病区id NUMBER(18),
    护理等级id NUMBER(18),
    出院科室id NUMBER(18),
    出院病床 VARCHAR2(10),
    出院日期 Date,
    住院天数 NUMBER(4),
    出院方式 VARCHAR2(10),
    是否确诊 NUMBER(1),
    确诊日期 Date,
    新发肿瘤 number(1),
    血型 VARCHAR2(10),
    抢救次数 NUMBER(5),
    成功次数 NUMBER(5),
    随诊标志 NUMBER(1),
    随诊期限 NUMBER(5),
    尸检标志 NUMBER(1),
    门诊医师 VARCHAR2(20),
    责任护士 VARCHAR2(20),
    住院医师 VARCHAR2(20),
    病案号 VARCHAR2(20),
    编目员编号 VARCHAR2(6),
    编目员姓名 VARCHAR2(20),
    编目日期 Date,
    状态 NUMBER(3),
    费用和 NUMBER(16,5),
    姓名 VARCHAR2(100),
    性别 VARCHAR2(4),
    年龄 varchar2(20),
    身高 NUMBER(16,5),
    体重 NUMBER(16,5),
    婚姻状况 VARCHAR2(4),
    职业 VARCHAR2(80),
    国籍 VARCHAR2(30),
    学历 VARCHAR2(10),
    单位电话 VARCHAR2(20),
    单位邮编 VARCHAR2(6),
    单位地址 VARCHAR2(100),
    区域 VARCHAR2(100),
    家庭地址 VARCHAR2(100),
    家庭电话 VARCHAR2(20),
    家庭地址邮编 VARCHAR2(6),
    联系人姓名 VARCHAR2(64),
    联系人关系 VARCHAR2(30),
    联系人地址 VARCHAR2(100),
    联系人电话 VARCHAR2(20),
    联系人身份证号 VARCHAR2(18),
    户口地址 VARCHAR2(100),
    户口地址邮编 VARCHAR2(6),
    中医治疗类别 VARCHAR2(4),
    险类 NUMBER(3),
    社区 Number(5),
    审核标志 NUMBER(1),
    审核人 VARCHAR2(20),
    审核日期 DATE,
    是否上传 NUMBER(1),
    数据转出 Number(1),
    登记人 Varchar2(20),
    登记时间 Date,
    备注 Varchar2(100),
    病案状态 Number(3),
    病人类型 Varchar2(50),
    封存时间 Date,
    路径状态 number(1),
    单病种 varchar2(2),
    待转出 Number(3),
    婴儿科室ID NUMber(18),
    婴儿病区ID NUMber(18),
    母婴转科标志 varchar2(100),
    医嘱重整时间 Date,
    是否禁止自动记帐 Number(1),
    入科时间 Date,
    挂号ID number(18),
    是否转科婴儿 NUMBER (1),
    审核说明 varchar2(200),
    预出院日期 date,
    工作单位 varchar2(100))
    TABLESPACE zl9Patient initrans 20 
;
Create Table 病案主页从表(
	病人ID NUMBER(18),
	主页ID NUMBER(5),
	信息名 VARCHAR2(20),
	信息值 VARCHAR2(100))
    TABLESPACE zl9Patient
    initrans 20;
Create Table 病人变动记录(
    Id Number(18),
    病人ID number(18) Not Null,
    主页ID number(5) Not Null,
    开始时间 Date,
    开始原因 number(3),
    附加床位 number(1),
    病区id number(18),
    科室id number(18),
    医疗小组id number(18),
    护理等级id number(18),
    床位等级id number(18),
    床号 VARCHAR2(10),
    责任护士 varchar2(20),
    经治医师 varchar2(20),
    主治医师 varchar2(20),
    主任医师 varchar2(20),
    病情         varchar2(20),
    终止人员 varchar2(20),
    终止时间 Date,
    终止原因 number(3),
    操作员编号 varchar2(6),
    操作员姓名 varchar2(20),
    上次计算时间 Date)
    TABLESPACE zl9Patient
    PCTFREE 5 initrans 20;
Create Table 病人过敏药物(
    病人ID NUMBER(18),
    过敏药物id NUMBER(18),
    过敏药物 VARCHAR2(60),
	过敏反应 varchar2(100))
    TABLESPACE zl9Patient
    PCTFREE 5;
Create Table 床位状况记录(
    病区id NUMBER(18),
    床号 VARCHAR2(10),
    科室id NUMBER(18),
    房间号 VARCHAR2(10),
    性别分类 VARCHAR2(10),
    床位编制 VARCHAR2(10),
    等级id NUMBER(18),
    状态 VARCHAR2(4),
    病人id NUMBER(18),
    共用 NUMBER(1) Default 0,
    顺序号 NUMBER(10,1))
    TABLESPACE zl9Patient PCTFREE 20 initrans 20
    Cache Storage(Buffer_Pool Keep)
;
Create Table 床位增减记录(
    日期 Date,
    变动 NUMBER(5),
    病区id NUMBER(18),
    床号 VARCHAR2(10),
    科室id NUMBER(18),
	床位编制 VARCHAR2(10))
    TABLESPACE zl9Patient;
Create Table 病人审批项目(
    病人ID      NUMBER(18),
    主页ID	NUMBER(5),
    项目ID      NUMBER(18),
    审批人      VARCHAR2(20),
    审批时间	Date,
    使用限量	NUMBER(16,5),
    已用数量	NUMBER(16,5))
    TABLESPACE zl9Patient;
CREATE TABLE 审批项目模板(
	编码  number(5),
	名称  varchar2(20),
	项目ID NUMBER(18))
	TABLESPACE zl9BaseItem;
Create Table 病人来源(
    编码 VARCHAR2(1),
    名称 VARCHAR2(20),
    简码 VARCHAR2(4),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;    
Create Table 病人备注信息(
    Id Number(18),
    病人ID number(18) Not Null,
    主页ID number(5) Not Null,
    内容 varchar2(200),
    登记时间 Date,
    登记人 varchar2(20),
    是否完成 Number(1),
    完成时间 Date,
    完成人 varchar2(20))
    TABLESPACE zl9Patient;
Create Table 病人地址信息(
    病人ID NUMBER(18),
    主页ID NUMBER(5),
    地址类别 Number(5),
    省 varchar2(100),
    市 varchar2(100),
    县 varchar2(100),
    乡镇 Varchar2(100),
    其他 varchar2(100),
    区划代码 Varchar2(15)) 
    tablespace zl9Patient;
--病人病案
----------------------------------------------------------------------------
CREATE TABLE 病人过敏记录(
    ID NUMBER(18),
    病人ID NUMBER(18),
    主页ID NUMBER(18),
    记录来源 NUMBER(1),
    药物ID NUMBER(18),
    药物名 VARCHAR2(60),
    结果 NUMBER(1),
    过敏时间 DATE,
    记录时间 DATE,
    记录人 VARCHAR2(20),
    过敏反应 varchar2(100),
    过敏源编码 Varchar2(10),
    待转出 Number(3))
    TABLESPACE zl9CisRec
    PCTFREE 5;
CREATE TABLE 病人症状记录(
    病人ID NUMBER(18),  
    主页ID NUMBER(18),	--门诊病人填：挂号ID
    序号   NUMBER(4),
    编码   VARCHAR2(10),
    名称   VARCHAR2(100),
    开始日期 DATE,
    结束日期 DATE,
    记录人 VARCHAR2(20),
    记录时间 DATE)
    TABLESPACE zl9CisRec;
CREATE TABLE  病人免疫记录 (
	病人ID NUMBER(18),
	接种时间 Date,
	接种名称 varchar2(200)) 
	TABLESPACE zl9Patient;	
Create Table 病人诊断记录(
    ID NUMBER(18),
    病人ID NUMBER(18),
    主页ID NUMBER(18),
    医嘱ID NUMBER(18),
    记录来源 NUMBER(1),
    诊断次序 NUMBER(2) DEFAULT 1,
    编码序号 NUMBER(2) DEFAULT 1,
    病历ID NUMBER(18),
    病例ID NUMBER(18),
    诊断类型 NUMBER(2),
    疾病ID NUMBER(18),
    诊断ID NUMBER(18),
    证候ID NUMBER(18),
    诊断描述 VARCHAR2(500),
    入院病情 varchar2(30),
    出院情况 VARCHAR2(10),
    是否未治 NUMBER(1),
    是否疑诊 NUMBER(1),
    备注 varchar2(200),
    记录日期 DATE,
    记录人 VARCHAR2(20),
    取消时间 DATE,
    取消人 VARCHAR2(20),
    发病时间 date,
    待转出 Number(3),
    前注释 VARCHAR2(200),
    后注释 VARCHAR2(200),
    录入次序 VARCHAR2(4),
    编码类别 Varchar2(2))
    TABLESPACE zl9CisRec 
;
CREATE TABLE 病人诊断医嘱(
    诊断ID NUMBER(18),
    医嘱ID NUMBER(18),
	待转出 Number(3))
    TABLESPACE zl9CisRec
    PCTFREE 5;
Create Table 诊断符合情况(
	病人ID number(18),
	主页ID number(5),
	符合类型 number(2),
	符合情况 number(2))
	TABLESPACE zl9CisRec
	PCTFREE 5;
Create Table 病人手麻记录(
    ID NUMBER(18),
    病人ID NUMBER(18),
    主页ID NUMBER(5),
    记录来源 NUMBER(1),
    手术次序 NUMBER(2),
    手术日期 DATE,
    准备天数 NUMBER(3),
    手术情况 VARCHAR2(8),
    再次手术 NUMBER(1),
    手术开始时间 DATE,
    手术结束时间 DATE,
    抗菌用药时间 Date,
    拟行手术 VARCHAR2(100),
    手术操作ID NUMBER(18),
    诊疗项目ID NUMBER(18),
    已行手术 varchar2(300),
    主刀医师 VARCHAR2(20),
    助产护士 VARCHAR2(20),
    第一助手 VARCHAR2(20),
    第二助手 VARCHAR2(20),
    手术护士 VARCHAR2(20),
    麻醉开始时间 DATE,
    麻醉结束时间 DATE,
    麻醉方式 NUMBER(18),
    ASA分级 VARCHAR2(20),
    NNIS分级 VARCHAR2(20),
    手术级别 number(2),
    麻醉类型 VARCHAR2(20),
    麻醉质量 VARCHAR2(6),
    输液总量 NUMBER(5),
    麻醉医师 VARCHAR2(20),
    输氧开始时间 DATE,
    输氧结束时间 DATE,
    切口 VARCHAR2(2),
    愈合 VARCHAR2(6),
    切口部位 VARCHAR2(100),
    重返计划 NUMBER(1),
    重返目的 VARCHAR2(100),
    切口感染 NUMBER(1),
    并发症 NUMBER(1),
    术前抗菌用药 NUMBER(1),
    抗菌用药天数 NUMBER(5),
    非预期的二次手术 NUMBER(1),
    麻醉并发症 NUMBER(1),
    术中异物遗留 NUMBER(1),
    手术并发症 NUMBER(1),
    术后出血或血肿 NUMBER(1),
    手术伤口裂开 NUMBER(1),
    术后深静脉血栓 NUMBER(1),
    术后生理代谢紊乱 NUMBER(1),
    术后呼吸衰竭 NUMBER(1),
    术后肺栓塞 NUMBER(1),
    术后败血症 NUMBER(1),
    术后髋关节骨折 NUMBER(1),
    记录日期 DATE,
    记录人 VARCHAR2(20),
    取消时间 DATE,
    取消人 VARCHAR2(20),
    待转出 Number(3),
    数据来源 number(1),
    前注释 VARCHAR2(200),
    后注释 VARCHAR2(200),
    手术类型 VARCHAR2(20))
    TABLESPACE zl9CisRec 
;
Create Table 病人新生儿记录(
    病人ID NUMBER(18),
    主页ID NUMBER(18),
    序号 NUMBER(3),
    婴儿姓名 VARCHAR2(100),
    婴儿性别 VARCHAR2(4),
    分娩次数 NUMBER(3),
    分娩方式 VARCHAR2(20),
    胎儿状况 VARCHAR2(20),
    出生时间 DATE,
    身长 number(16,5),
    体重 number(16,5),
    血型 varchar2(10),
    备注说明 varchar2(100),
    死亡时间 Date,
    登记时间 Date,
    登记人 VARCHAR2(20),
    婴儿病人ID NUMBER (18),
    婴儿主页ID NUMBER (5))
    TABLESPACE zl9CisRec PCTFREE 5 
;
Create Table 病人分娩信息(
    病人ID number(18),
    主页ID number(5),
    胎儿次序 number(5),
    分娩方式 varchar2(20) not Null,
    出生胎位 varchar2(10) not Null,
    分娩情况 varchar2(10) not Null,
    出生缺陷 number(1) not Null,
    婴儿性别 varchar2(10) not Null,
    婴儿体重 varchar2(10),
    Apgar评分 varchar2(10),
    分娩时间 Date)
    TABLESPACE zl9CisRec PCTFREE 5 
;
Create Table 新生儿诊断记录(
    病人ID number(18),
    主页ID number(5),
    胎儿次序 number(5),
    诊断次序 number(5),
    疾病ID number(5),
    描述信息 varchar2(100),
    编码类别 VARCHAR2(2),
    录入次序 VARCHAR2(4))
    TABLESPACE zl9CisRec PCTFREE 5 
;
Create Table 病人抗生素记录(
       病人Id NUMBER(18),
       主页Id NUMBER(5),
       药名id NUMBER(18),
       药品名称 VARCHAR2(80),
       用药目的 VARCHAR2(200),
       使用阶段 VARCHAR2(30),
       使用天数 NUMBER(18,2),
       记录人 VarCHAR2(30),
       记录时间 Date,
       一类切口预防用 Number(1),
       DDD数 Number(16,4),
       联合用药 varchar2(30))
  TABLESPACE zl9CisRec;
Create Table 病案重症监护情况(
	病人ID number(18),
	主页ID number(5),
	序号   number(18), 
	监护室名称 varchar2(100),
	进入时间 Date,
	退出时间 Date,
	再入住计划 number(1),
	再入住原因 varchar2(100),
	人工气道脱出   NUMBER(1),
    重返重症医学科  NUMBER(1),
    重返间隔时间   VARCHAR2(30)
)TABLESPACE zl9CisRec;
Create Table 病案化疗记录(
  病人ID NUMBER(18),
  主页ID NUMBER(5),
  序号   number(18),
  疾病ID NUMBER(18),
  开始日期 DATE,
  结束日期 DATE,
  疗程数   number(16,5),
  总量     number(16,5),
  化疗方案 VARCHAR2(50),
  化疗效果 VARCHAR2(10))
    TABLESPACE zl9CisRec ;
    
Create Table 病案放疗记录(
  病人ID NUMBER(18),
  主页ID NUMBER(5),
  序号   number(18),
  疾病ID NUMBER(18),
  开始日期 DATE,
  结束日期 DATE,
  设野部位 VARCHAR2(50),
  放射剂量   NUMBER(16,5),
  累计量     NUMBER(16,5),
  放疗效果 VARCHAR2(10))
    TABLESPACE zl9CisRec ;
Create Table 病案精神治疗(
	病人ID NUMBER(18),
	主页ID NUMBER(5),
	序号   number(18),
	药品ID number(18),
	药物名称 varchar2(200),
	疗程	varchar2(50),
	最高日量 varchar2(50),
	特殊反应 VARCHAR2(100),
	疗效 VARCHAR2(50))
    TABLESPACE zl9CisRec ;
Create Table 器械导管使用情况(
	病人ID number(18),
	主页ID number(5),
	序号 number(18),
	监护室名称 VARCHAR2(50),
	器械及导管 Varchar2(20),
	开始使用时间 Date,
	结束使用时间 Date,
	感染累计时间 varchar2(20))
TABLESPACE zl9CisRec;
Create Table 病人感染记录(
	序号 number(5),
	病人ID NUMBER(18),
	主页ID NUMBER(5),
	登记时间 Date,
	登记人 VARCHAR2(20),
	确诊日期 Date,
	感染部位 VARCHAR2(20),
	感染名称 VARCHAR2(30)
)TABLESPACE zl9CisRec PCTFREE 5;
Create Table 病人病原学检查(
	序号 number(5),
	病人ID NUMBER(18),
	主页ID NUMBER(5),
	登记时间 Date,
	登记人 VARCHAR2(20),
	标本 VARCHAR2(20),
	病原学代码 VARCHAR2(20),
	送检日期 Date
)TABLESPACE zl9CisRec PCTFREE 5;
----------------------------------------------------------------------------
--[[14.费用业务]]
----------------------------------------------------------------------------
Create table 预交单据余额 
(
预交ID number(18),
病人ID Number(18),
预交类别 number(1),
预交余额 number(16,5)
) Tablespace zl9Expense 
  PCTFREE 5 initrans 20;
Create Table 消费卡入库记录(
    ID Number(18),
    接口编号 Number(6),
    前缀文本 Varchar2(2),
    开始卡号 Varchar2(50),
    终止卡号 Varchar2(50),
    入库数量 Number(18),
    剩余数量 Number(18),
    备注 Varchar2(200),
    登记人 Varchar2(20),
    登记时间 Date,
    是否存在卡 Number(1) Default 0,
    批次 varchar2(20))
    Tablespace Zl9expense 
;
Create Table 消费卡领用记录(
    ID Number(18),
    接口编号 Number(6),
    领用人 Varchar2(20),
    前缀文本 Varchar2(2),
    开始卡号 Varchar2(50),
    终止卡号 Varchar2(50),
    使用方式 Number(1),
    登记时间 Date,
    使用时间 Date,
    登记人 Varchar2(20),
    当前卡号 Varchar2(50),
    剩余数量 Number(18),
    批次 Varchar2(20),
    核对人 Varchar2(20),
    核对时间 Date,
    核对结果 Number(1),
    核对模式 Number(1),
    备注 Varchar2(200),
    签字人 Varchar2(20),
    签字时间 Date,
    入库ID number(18))
    Tablespace Zl9expense 
;
Create Table 消费卡报损记录(
    ID Number(18), 
    入库id Number(18), 
    开始卡号 Varchar2(50), 
    终止卡号 Varchar2(50), 
    数量 Number(18),
    报损原因 Varchar2(200), 
    报损人 Varchar2(20), 
    报损时间 Date
) Tablespace Zl9expense;
Create Table 消费卡使用记录(
    ID Number(18), 
    卡号 Varchar2(50), 
    性质 Number(1), 
    原因 Number(1), 
    领用id Number(18), 
    回收次数 Number(3),
    接口编号 Number(6), 
    使用时间 Date, 
    使用人 Varchar2(20), 
    核对人 Varchar2(20), 
    核对时间 Date, 
    核对结果 Number(1),
    备注 Varchar2(200)
 ) Tablespace Zl9expense;
Create Table 消费卡变动记录(
    ID Number(18), 
    消费卡id Number(18), 
    卡号 Varchar2(50), 
    变动类型 Number(3), 
    变动原因 Varchar2(100),
    原密码 Varchar2(50), 
    现密码 Varchar2(50), 
    原卡号 Varchar2(50), 
    操作员姓名 Varchar2(20), 
    登记时间 Date
 ) Tablespace Zl9expense;
Create Table 帐户缴款余额(
    性质 Number(2),
    结算方式 Varchar2(20),
    余额 Number(16, 5),
    扣率 Number(16, 5),
    实际缴款 Number(16, 5),
    有效期 Date,
    消费卡id Number(18),
    交易序号 Number(18),
    卡类别id Number(18),
    卡号 Varchar2(50),
    交易流水号 Varchar2(50),
    交易说明 Varchar2(500),
    缴款时间 Date
) Tablespace Zl9expense;
Create Table 费用变动记录(
 ID Number(18),
 记录状态 Number(3),
 病人id Number(18),
 主页id Number(5),
 变动时间 Date,
 原变动id Number(18),
 目标变动id Number(18),
 原病区id Number(18),
 目标病区id Number(18),
 费用id Number(18),
 NO Varchar2(8),
 收费类别 Varchar2(1),
 收费细目id Number(18),
 医嘱序号 Number(18),
 数量 Number(16, 5),
 单价 Number(16, 5),
 应收金额 Number(16, 5),
 实收金额 Number(16, 5),
 状态 Number(2),
 摘要 Varchar2(500),
 操作员编号 Varchar2(6),
 操作员姓名 Varchar2(20),
 待转出 Number(3))
Tablespace Zl9expense Pctfree 5;
Create Table 常用退号原因
(
  编码   VARCHAR2(4),
  名称   VARCHAR2(50),
  简码   VARCHAR2(25),
  缺省标志 NUMBER(1))
Tablespace ZL9BASEITEM;
Create Table 病人服务信息记录(
   ID number(18),
   通知类型 number(18),
   记录ID number(18),
   挂号ID number(18),
   号源ID number(18),
   号码 varchar2(10),
   科室ID number(18),
   项目ID number(18),
   医生ID number(18),
   医生姓名 varchar2(50),
   病人ID number(18),
   复诊方式 number(2),
   数量 number(10),
   开始时间 Date,
   终止时间 Date,
   通知原因 varchar2(100),
   登记人 varchar2(50),
   登记时间 Date,
   处理说明 varchar2(100),
   处理人 varchar2(50),
   处理时间 Date)
TABLESPACE zl9BaseItem ;
Create Table 财务组组长构成(
    组ID Number(18),
    组长ID Number(18),
    上次轧帐时间 Date)
    TABLESPACE zl9BaseItem
    PCTFREE 5;
Create Table 三方退款信息(
    结帐ID Number(18),
    记录ID Number(18),
    金额 Number(16,5),
    卡号 Varchar2(50),
    交易流水号 Varchar2(50),
    交易说明 Varchar2(500),
    待转出 Number(3),
    性质 Number(1),
    原交易流水号 varchar2(50),
    原交易说明 VARCHAR2(500),
    是否未退 number(1),
    是否转帐 number(1) ,
    卡类别ID number(18))
    TABLESPACE zl9Expense 
;
create table 费用清单打印
(
  ID   number(18),
  门诊标志 Number(3),
  记录性质 number(3),
  No   VARCHAR2(30),
  序号 number(18),
  记录状态 number(18),
  收费细目ID number(18),
  打印次数 number(5),
  病人ID number(18),
  主页ID number(18),
  上次打印时间 Date,
  打印站点 Varchar2(100),
  打印站点IP Varchar2(100),
  打印人 Varchar2(100),
  打印时间 Date,
  待转出  NUMBER(3)
)
TABLESPACE zl9Expense PCTFREE 5 initrans 20;
Create Table 医保结算明细 (
    结帐id Number(18),
    NO Varchar2(8),
    结算方式 Varchar2(20),
    金额 Number(16,5),
    备注 Varchar2(200),
    待转出 Number(3),
    卡类别ID number(18),
    关联交易ID number(18),
    交易流水号 varchar2(50),
    交易说明 varchar2(100))
    Tablespace zl9Expense 
;
Create Table 就诊变动记录(
    ID Number(18),
    类别 Number(2),
    挂号单 Varchar2(8),
	病人ID Number(18),
    变动原因 Varchar2(200),
    原号码 Varchar2(5),
    现号码 Varchar2(5),
    原科室ID Number(18),
    现科室ID Number(18),
    原项目ID	Number(18),
    现项目ID  Number(18),
    原医生ID	Number(18),
    现医生ID	Number(18),
    原医生姓名	Varchar2(50),
    现医生姓名	Varchar2(50),
    原诊室	Varchar2(20),
    现诊室	Varchar2(20),
    原号序	Number(5),
    现号序	Number(5),
    原预约时间	Date,
    现预约时间	Date,
    登记时间	Date,
    操作员姓名	Varchar2(100),
    操作员编号	Varchar2(6))
    TABLESPACE zl9Patient;
Create Table 费用补充记录(
    记录性质 number(3),
    NO varchar2(8),
    记录状态 Number (3),
    实际票号 varchar2(50),
    结算ID Number(18),
    病人ID number(18),
    收费结帐ID Number(18),
    费用状态 number(1),
    操作员编号 varchar2(6),
    操作员姓名 varchar2(100),
    备注 varchar2(100),
    登记时间 Date,
    缴款组ID number(18),
    结算序号 number(18),
    附加标志 number(3),
    待转出 number(3),
    姓名 Varchar2(100),
    性别 Varchar2(4),
    年龄 Varchar2(20),
    门诊号 Number(18),
    住院号 Number(18),
    付款方式名称 Varchar2(20))
    TABLESPACE zl9Expense 
;
Create Table 凭条打印记录(
  NO varchar2(8),
  记录性质 NUMBER(3),
  打印时间 Date,
  打印类型 NUMBER(3),
  打印人 varchar2(100),
  机器名 varchar2(100),
  IP地址 varchar2(100),
  备注 varchar2(500),
  待转出 NUMBER(3)) 
TABLESPACE zl9Expense PCTFREE 5 initrans 20;
Create Table 病人挂号记录(
    ID NUMBER(18),
    NO VARCHAR2(8),
    记录性质 number(3) default(1),
    记录状态 NUMBER(3)default(1),
    病人ID NUMBER(18),
    门诊号 NUMBER(18),
    姓名 VARCHAR2(100),
    性别 VARCHAR2(4),
    年龄 varchar2(20),
    复诊 NUMBER(1),
    号别 VARCHAR2(5),
    号序 NUMBER(5),
    急诊 NUMBER(1),
    诊室 VARCHAR2(20),
    附加标志 NUMBER(1),
    执行部门ID NUMBER(18),
    执行人 VARCHAR2(20),
    执行状态 NUMBER(1),
    执行时间 DATE,
    完成时间 DATE,
    登记时间 DATE,
    发生时间 DATE,
    操作员编号 VARCHAR2(6),
    操作员姓名 VARCHAR2(20),
    传染病上传 Number(1),
    发病时间 Date,
    发病地址 varchar2(200),
    分诊时间 Date,
    社区 Number(5),
    摘要 Varchar2(1000),
    转诊号别 VARCHAR2(5),
    转诊科室ID Number(18),
    转诊诊室 VARCHAR2(20),
    转诊医生 VARCHAR2(20),
    转诊状态 Number(1),
    续诊科室ID Number(18),
    病生理情况 VARCHAR2(250),
    预约 number(2),
    预约方式 varchar2(10),
    记录标志 number(2),
    退号审核人 VARCHAR2(20),
    退号审核时间 DATE,
    预约时间 DATE,
    接收人 VARCHAR2(20),
    接收时间 Date,
    交易流水号 VARCHAR2(50),
    交易说明 VARCHAR2(50),
    合作单位 VARCHAR2(50),
    预约操作员 VARCHAR2(20),
    预约操作员编号 VARCHAR2(6),
    险类 number(3),
    待转出 Number(3),
    医疗付款方式 Varchar2(20),
    出诊记录ID number(18),
    收费单 varchar2(2000),
    路径状态 number(1),
    取号标志 NUMBER(2),
    挂号项目ID NUMBER(18),
    费别 VARCHAR2(10),
    号类 varchar2(10),
    结算模式 number(2))
    TABLESPACE zl9Patient PCTFREE 5 initrans 20 
;
Create Table 病人转诊记录(
    挂号ID NUMBER(18),
    NO VARCHAR2(8),
    申请科室ID NUMBER(18),
    申请医生 VARCHAR2(20),
    接收科室ID NUMBER(18),
    接收医生 VARCHAR2(20),
	接收时间 Date,
	待转出 Number(3)
	)
    TABLESPACE zl9Patient;
Create Table 病人挂号汇总(
    日期 date,
    科室id NUMBER(18),
    项目ID NUMBER(18),
    医生姓名 VARCHAR2(20),
    医生ID NUMBER(18),
    号码 VARCHAR2(5),
    已挂数 NUMBER(5),
    已约数 NUMBER(5),
    其中已接收 Number(5),
	待转出 Number(3))
    TABLESPACE zl9Expense
    PCTFREE 5 initrans 20;
Create  Table 合作单位挂号汇总(
	日期 Date,
	号码 Varchar2(5),
	合作单位 Varchar2(50),
	序号 Number(5),
	已约数 Number(10),
	已接数 Number(10)
	) Tablespace zl9Expense
	PCTFREE 5 initrans 20;
Create Table 病人余额(
    病人id NUMBER(18),
    性质 NUMBER(1),
    类型 NUMBER(2) DEFAULT 2,
    预交余额 NUMBER(16,5),
    费用余额 NUMBER(16,5))
    TABLESPACE zl9Expense
    initrans 20;
Create Table 病人缴款记录(
    ID NUMBER(18),
    病人ID NUMBER(18),
    No VARCHAR2(8),
    记录状态 Number(3),
    结算方式 VARCHAR2(20),
    结算号 VARCHAR2(10),
    金额 NUMBER(16,5),
    摘要 VARCHAR2(50),
    登记时间 Date,
    登记人 VARCHAR2(20))
    TABLESPACE zl9Expense
    PCTFREE 5;
Create Table 病人缴款对照(
    缴款单 VARCHAR2(8),
    结帐ID NUMBER(18),
    金额 NUMBER(16,5))
    TABLESPACE zl9Expense
    PCTFREE 5;
Create Table 病人催款记录(
    ID		NUMBER(18),
    病人ID  NUMBER(18),
    主页ID  NUMBER(18),
    预交余额  NUMBER(16,5),
    未结费用  NUMBER(16,5),
    自费金额  NUMBER(16,5),
    医保预结  NUMBER(16,5),
    当前余额  NUMBER(16,5),
    催款下限  NUMBER(16,5),
    催款标准  NUMBER(16,5),
    催款金额  NUMBER(16,5), 
    打印日期 DATE ,
    打印人     VARCHAR2(20))
    TABLESPACE zl9Expense;
Create Table 病人结帐记录(
    ID NUMBER(18),
    NO VARCHAR2(8),
    实际票号 VARCHAR2(20),
    记录状态 NUMBER(3),
    中途结帐 NUMBER(1),
    病人id NUMBER(18),
    操作员编号 VARCHAR2(6),
    操作员姓名 VARCHAR2(20),
    备注 VARCHAR2(50),
    原因 VARCHAR2(100),
    收费时间 Date,
    开始日期 Date,
    结束日期 Date,
    缴款组ID number(18),
    结帐类型 NUMBER(1),
    待转出 Number(3),
    结算状态 number(2),
    主页ID number(18),
    住院次数 varchar2(2000),
    结帐金额 number(16,5),
    姓名 Varchar2(100),
    性别 Varchar2(4),
    年龄 Varchar2(20),
    门诊号 Number(18),
    住院号 Number(18),
    付款方式名称 Varchar2(20),
    是否电子票据 number(2))
    TABLESPACE zl9Expense PCTFREE 5 
;
Create Table 住院费用记录(
    ID NUMBER(18),
    记录性质 NUMBER(3),
    NO VARCHAR2(8),
    实际票号 VARCHAR2(50),
    记录状态 NUMBER(3),
    序号 NUMBER(18),
    从属父号 NUMBER(5),
    价格父号 NUMBER(5),
    多病人单 NUMBER(1) default 0,
    记帐单ID NUMBER(18) default 0,
    病人id NUMBER(18),
    主页id NUMBER(5),
    医嘱序号 NUMBER(18),
    门诊标志 NUMBER(3) default 1,
    记帐费用 NUMBER(1) default 0,
    姓名 VARCHAR2(100),
    性别 VARCHAR2(4),
    年龄 varchar2(20),
    标识号 NUMBER(18),
    床号 VARCHAR2(10),
    病人病区id NUMBER(18),
    病人科室id NUMBER(18),
    费别 VARCHAR2(10),
    收费类别 VARCHAR2(1),
    收费细目id NUMBER(18),
    计算单位 VARCHAR2(20),
    付数 NUMBER(3) default 1,
    发药窗口 VARCHAR2(50),
    数次 NUMBER(16,5),
    加班标志 NUMBER(1),
    附加标志 NUMBER(1),
    婴儿费 NUMBER(1),
    收入项目id NUMBER(18),
    收据费目 VARCHAR2(20),
    标准单价 NUMBER(16,5),
    应收金额 NUMBER(16,5),
    实收金额 NUMBER(16,5),
    划价人 VARCHAR2(20),
    开单部门id NUMBER(18),
    开单人 varchar2(41),
    发生时间 Date,
    登记时间 Date,
    执行部门id NUMBER(18),
    执行人 VARCHAR2(20),
    执行状态 NUMBER(2),
    执行时间 date,
    结论 Varchar2(500),
    操作员编号 VARCHAR2(6),
    操作员姓名 VARCHAR2(20),
    结帐id NUMBER(18),
    结帐金额 NUMBER(16,5),
    保险大类id number(18),
    保险项目否 number(1),
    保险编码 varchar2(20),
    费用类型 varchar2(20),
    统筹金额 number(16,5),
    是否上传 number(1),
    摘要 Varchar2(1000),
    是否急诊 Number(1) Default 0,
    缴款组ID number(18),
    医疗小组ID NUMBER(18),
    待转出 Number(3),
    费用状态 NUMBER(2),
    记费同步标志bak Number(1),
    作废同步标志bak Number(1),
    转费同步标志bak Number(1),
    批次 Number(18),
    领药部门ID Number(18),
    医嘱期效 Number(1),
    是否保密 Number(1),
    是否附费 number(1))
    TABLESPACE zl9Expense initrans 20 
;
Create Table 门诊费用记录(
    ID NUMBER(18),
    记录性质 NUMBER(3),
    NO VARCHAR2(8),
    实际票号 VARCHAR2(50),
    记录状态 NUMBER(3),
    序号 NUMBER(18),
    从属父号 NUMBER(5),
    价格父号 NUMBER(5),
    记帐单ID NUMBER(18) default 0,
    病人id NUMBER(18),
    医嘱序号 NUMBER(18),
    门诊标志 NUMBER(3) default 1,
    记帐费用 NUMBER(1) default 0,
    姓名 VARCHAR2(100),
    性别 VARCHAR2(4),
    年龄 varchar2(20),
    标识号 NUMBER(18),
    付款方式 VARCHAR2(10),
    病人科室id NUMBER(18),
    费别 VARCHAR2(10),
    收费类别 VARCHAR2(1),
    收费细目id NUMBER(18),
    计算单位 VARCHAR2(20),
    付数 NUMBER(3) default 1,
    发药窗口 VARCHAR2(50),
    数次 NUMBER(16,5),
    加班标志 NUMBER(1),
    附加标志 NUMBER(1),
    婴儿费 NUMBER(1),
    收入项目id NUMBER(18),
    收据费目 VARCHAR2(20),
    标准单价 NUMBER(16,5),
    应收金额 NUMBER(16,5),
    实收金额 NUMBER(16,5),
    划价人 VARCHAR2(20),
    开单部门id NUMBER(18),
    开单人 varchar2(41),
    发生时间 Date,
    登记时间 Date,
    执行部门id NUMBER(18),
    执行人 VARCHAR2(20),
    执行状态 NUMBER(2),
    执行时间 date,
    结论 VARCHAR2(1000),
    操作员编号 VARCHAR2(6),
    操作员姓名 VARCHAR2(20),
    结帐id NUMBER(18),
    结帐金额 NUMBER(16,5),
    保险大类id number(18),
    保险项目否 number(1),
    保险编码 varchar2(20),
    费用类型 varchar2(20),
    统筹金额 number(16,5),
    是否上传 number(1),
    摘要 Varchar2(1000),
    是否急诊 Number(1) Default 0,
    缴款组ID number(18),
    费用状态 number(4),
    待转出 Number(3),
    挂号ID Number(18),
    主页ID Number(5),
    病人病区id Number(18),
    记费同步标志bak Number(1),
    作废同步标志bak Number(1),
    批次 Number(18),
    医嘱期效 Number(1),
    是否保密 Number(1),
    是否附费 number(1))
    TABLESPACE zl9Expense initrans 20 
;
Create Table 病人费用销帐(
    费用ID NUMBER(18),
    申请类别 number(2) DEFAULT 0,
    收费细目id NUMBER(18),
    申请部门id NUMBER(18),
    审核部门id NUMBER(18),
    数量 NUMBER(16,5),
    申请人 VARCHAR2(20),
    申请时间 Date,
    审核人 VARCHAR2(20),
    审核时间 Date,
    核查人 VARCHAR2(20),
    核查日期 Date,
    状态 NUMBER(1),
    销帐原因 Varchar2(200),
    待转出 Number(3))
    TABLESPACE zl9Expense PCTFREE 5 
;
Create Table 病人退费申请(
    NO VARCHAR2(8),
    记录性质 NUMBER(3),
    申请人 VARCHAR2(20),
    申请时间 Date,
    审核人 VARCHAR2(20),
    审核时间 Date,
    申请原因 Varchar2(100),
    审核原因 Varchar2(100),
    状态 Number(2))
    TABLESPACE zl9Expense PCTFREE 5 
;
Create Table 病人结帐汇总(    
	结帐时间 date,
	病人ID  NUMBER(18),
	主页ID NUMBER(5),
	结帐id NUMBER(18),
	病人病区id NUMBER(18),
	病人科室id NUMBER(18),
	开单部门id NUMBER(18),
	执行部门id NUMBER(18),
	收入项目id NUMBER(18),    
	应收金额 NUMBER(16,5),
	实收金额 NUMBER(16,5),
	结帐金额 NUMBER(16,5))
	TABLESPACE zl9Expense
	PCTFREE 5;
CREATE TABLE  费用审核记录(
	性质   number(2),
	费用ID NUMBER(18),
	病人ID NUMBER(18),
	主页ID NUMBER(18),
	审核人 VARCHAR2(20),
	审核日期 Date ,
	转出ID number(18),
	转出人 VARCHAR2(20),
	转出时间 DATE,
	记录状态 NUMBER(2))
     TABLESPACE zl9Expense;
Create Table 医生收入汇总(
    日期 date,
    开单人 VARCHAR2(20),
    执行人 VARCHAR2(20),
    病人病区id NUMBER(18),
    病人科室id NUMBER(18),
    开单部门id NUMBER(18),
    执行部门id NUMBER(18),
    收入项目id NUMBER(18),
    来源途径 NUMBER(3),
    记帐费用 NUMBER(1),
    应收金额 NUMBER(16,5),
    实收金额 NUMBER(16,5),
    结帐金额 NUMBER(16,5))
	TABLESPACE zl9Expense
	PCTFREE 5;
Create Table 病人费用汇总(
    日期 date,
    病人病区id NUMBER(18),
    病人科室id NUMBER(18),
    开单部门id NUMBER(18),
    执行部门id NUMBER(18),
    收入项目id NUMBER(18),
    来源途径 NUMBER(3),
    记帐费用 NUMBER(1),
    应收金额 NUMBER(16,5),
    实收金额 NUMBER(16,5),
    结帐金额 NUMBER(16,5))
    TABLESPACE zl9Expense
    PCTFREE 5;
Create Table 病人未结费用(
    病人id NUMBER(18),
    主页id NUMBER(5),
    病人病区id NUMBER(18),
    病人科室id NUMBER(18),
    开单部门id NUMBER(18),
    执行部门id NUMBER(18),
    收入项目id NUMBER(18),
    来源途径 NUMBER(3),
    金额 NUMBER(16,5))
    TABLESPACE zl9Expense;
Create Table 病人预交记录(
    ID NUMBER(18),
    记录性质 NUMBER(3),
    NO VARCHAR2(8),
    实际票号 VARCHAR2(20),
    记录状态 NUMBER(3),
    病人id NUMBER(18),
    主页id NUMBER(18),
    科室id NUMBER(18),
    缴款单位 VARCHAR2(50),
    单位开户行 VARCHAR2(50),
    单位帐号 Varchar2(50),
    摘要 VARCHAR2(50),
    金额 NUMBER(16,5),
    结算方式 VARCHAR2(20),
    结算号码 VARCHAR2(30),
    收款时间 Date,
    操作员编号 VARCHAR2(6),
    操作员姓名 VARCHAR2(20),
    冲预交 NUMBER(16,5),
    结帐id NUMBER(18),
    缴款 NUMBER(16,5),
    找补 NUMBER(16,5),
    缴款组ID number(18),
    预交类别 number(1),
    卡类别ID number(18),
    结算卡序号 number(18),
    卡号 varchar2(50),
    交易流水号 varchar2(50),
    交易说明 varchar2(500),
    合作单位 VARCHAR2(50),
    结算序号 NUMBER(18),
    校对标志 number(2),
    待转出 Number(3),
    结算性质 Number(2),
    会话号 Varchar2(45),
    关联交易Id number(18),
    交易时间 Date,
    交易人员 Varchar2(20),
    附加标志 number(2),
    是否转入预交 number(1),
    姓名 Varchar2(100),
    性别 Varchar2(4),
    年龄 Varchar2(20),
    门诊号 Number(18),
    住院号 Number(18),
    付款方式名称 Varchar2(20),
    是否电子票据 number(2),
    预交电子票据 number(2))
    TABLESPACE zl9Expense initrans 20 
;
 
Create Table 三方结算交易(
    交易ID Number(18),
    交易项目 varchar2(50),
    交易内容 VARCHAR2(100),
    待转出 Number(3),
    原预交ID Number(18),
    性质 Number(1))
    TABLESPACE zl9Expense 
;
Create Table 票据入库记录(
    ID NUMBER(18),
    票种 NUMBER(1),
    使用类别 VARCHAR2(50),
    有无票据 number(1),
    前缀文本 VARCHAR2(2),
    开始号码 VARCHAR2(50),
    终止号码 VARCHAR2(50),
    入库数量 NUMBER(10),
    剩余数量 NUMBER(10),
    备注 VARCHAR2(200),
    登记人 VARCHAR2(20),
    登记时间 DATE,
    批次 varchar2(20),
    类别名称 varchar2(50),
    是否下载 number(2))
    TABLESPACE zl9Expense 
;
Create Table 票据报损记录(
    ID NUMBER(18),
    入库ID NUMBER(18),
    开始号码 VARCHAR2(50),
    终止号码 VARCHAR2(50),
    数量 NUMBER(10),
    报损原因 VARCHAR2(200),
    报损人 VARCHAR2(20),
    报损时间 DATE)
    TABLESPACE zl9Expense;
Create Table 票据领用记录(
    ID NUMBER(18),
    票种 NUMBER(1),
    使用类别 VARCHAR2(50),
    领用人 VARCHAR2(20),
    前缀文本 VARCHAR2(2),
    开始号码 VARCHAR2(50),
    终止号码 VARCHAR2(50),
    使用方式 NUMBER(1),
    登记时间 DATE,
    使用时间 DATE,
    登记人 VARCHAR2(20),
    当前号码 VARCHAR2(50),
    剩余数量 NUMBER(10),
    批次 VARCHAR2(20),
    核对人 VARCHAR2(20),
    核对时间 DATE,
    核对结果 NUMBER(1),
    核对模式 NUMBER(1),
    签字人 varchar2(20),
    签字时间 DATE,
    备注 VARCHAR2(200),
    待转出 Number(3),
    入库ID number(18),
    类别名称 varchar2(50),
    是否下载 number(2))
    TABLESPACE zl9Expense PCTFREE 5 initrans 20 
;
Create Table 票据使用明细(
    ID Number(18),
    票种 NUMBER(1),
    号码 VARCHAR2(50),
    性质 NUMBER(1),
    原因 NUMBER(1),
    领用ID NUMBER(18),
    回收次数 NUMBER(3),
    打印ID NUMBER(18),
    使用时间 DATE,
    使用人 VARCHAR2(20),
    核对人 VARCHAR2(20),
    核对时间 DATE,
    核对结果 NUMBER(1),
    备注 VARCHAR2(200),
    待转出 Number(3),
    票据金额 Number(16,5),
    电子票据ID number(18))
    TABLESPACE zl9Expense PCTFREE 5 initrans 20 
;
Create Table 票据打印内容(
    ID NUMBER(18),
    数据性质 NUMBER(3),
    NO VARCHAR2(8),
    待转出 Number(3),
    打印类型 NUMBER(2))
    TABLESPACE zl9Expense PCTFREE 5 initrans 20
;
Create Table 票据打印明细(
	使用ID	NUMBER(18),
	票种	NUMBER(1),
	NO	VARCHAR2(8),
	票号	VARCHAR2(50),
	是否回收 NUMBER(1),
	关联票号序号 NUMBER(18), 
	序号	VARCHAR2(4000),
	待转出 Number(3))
TABLESPACE zl9Expense
PCTFREE 5 initrans 20;
Create Table 人员缴款余额(
    收款员 VARCHAR2(20),
    结算方式 VARCHAR2(20),
    性质 NUMBER,
    余额 NUMBER(16,5),
    上次轧帐时间 DATE)
    TABLESPACE zl9Expense
    PCTFREE 20 initrans 20;
Create Table 人员收缴记录(
    ID Number(18),
    记录性质 Number(2),
    NO varchar2(20),
    收款员 varchar2(20),
    收款部门ID Number(18),
    冲预交款 Number(16,5),
    借入合计 Number(16,5),
    借出合计 Number(16,5),
    摘要 varchar2(500),
    开始时间 Date,
    终止时间 Date,
    缴款组ID Number(18),
    登记人 varchar2(20),
    登记时间 Date,
    小组收款人 varchar2(20),
    小组收款时间 Date,
    小组收款ID Number(18),
    小组轧账ID Number(18),
    财务收款人 varchar2(20),
    财务收款时间 Date,
    财务收款ID Number(18),
    作废人 varchar2(20),
    作废时间 Date,
    收缴标志 number(2),
    待转出 Number(3),
    是否挂号 Number(1),
    是否就诊卡 Number(1),
    是否消费卡 Number(1),
    是否收费 Number(1),
    预交类别 Number(2),
    是否结帐 Number(1),
    是否押金 Number(2))
    TABLESPACE zl9Expense PCTFREE 20 
;
Create Table 人员收缴明细(
	收缴ID	 number(18),
	结算方式 Varchar2(20),
	结算号	 Varchar2(10),
	金额	 number(16,5),
	余额	 number(16,5),
	待转出	 number(3))
TABLESPACE zl9Expense
PCTFREE 5;
Create Table 人员收缴票据(
    收缴ID Number (18),
    票种 Number(2),
    性质 number(2),
    序号 Number(18),
    票据张数 Number(18),
    开始票号 Varchar2(50),
    终止票号 Varchar2(50),
    金额 number(16,5),
    发生时间 date,
    待转出 Number(3),
    批次 Varchar2(20))
    TABLESPACE zl9Expense PCTFREE 5
;
Create Table 人员收缴对照(
	收缴ID Number(18),
	性质   Number(2),
	记录ID Number(18),
	待转出 Number(3))
TABLESPACE zl9Expense
PCTFREE 5;
  
Create Table 人员暂存记录(
	ID	 number(18),
	收缴ID	 number(18),
	记录性质 number(2),
	NO	 varchar2(20),
	结算方式  varchar2(20),
	金额	 number(16,5),
	收款员	 varchar2(20),
	领用时间 Date,
	收回人   varchar2(20),
	收回时间 Date,	
	备注     varchar2(50),
	登记人   varchar2(20),
	登记时间 Date,	
	待转出 number(3))
TABLESPACE zl9Expense
PCTFREE 5;
CREATE TABLE 人员借款记录(
	ID number(18),
	借款金额 number(16,5),
	备注 varchar2(100),
	借款人 varchar2(20),
	申请时间 Date,
	结算方式 VARCHAR2(20)  NOT NULL,
	借出人 varchar2(20),
	借出时间 date,
	取消时间 DATE,
	取消原因 varchar2(100),
	待转出 Number(3))
	TABLESPACE zl9Expense
	PCTFREE 5;
Create Table 财务缴款分组(
    ID NUMBER(18),
    组名称 VARCHAR2(50),
    简码     VARCHAR2(20),
    说明	VARCHAR2(50),
    负责人ID Number(18),
    删除日期 Date ,
    上次轧帐时间 Date )
    TABLESPACE zl9BaseItem
    Cache Storage(Buffer_Pool Keep);
Create Table 缴款成员组成(
    组ID NUMBER(18),
    成员ID number(18))
    TABLESPACE zl9BaseItem
    Cache Storage(Buffer_Pool Keep);
Create Table 收费清点记录(
    日期 Date,
    收款员 VARCHAR2(20),
    性质 NUMBER(1),  --1-预交款,2-结帐,3-收费,4-挂号
    开始时间 Date,
    终止时间 Date)
    TABLESPACE zl9Expense
    PCTFREE 5;
--消费卡事务
Create Table 消费卡信息(
    ID Number(18),
    接口编号 number(6),
    卡类型 Varchar2(20),
    卡号 Varchar2(20),
    序号 Number(18),
    密码 Varchar2(50),
    限制类别 Varchar2(500),
    可否充值 Number(2) DEFAULT 0,
    有效期 Date,
    发卡原因 Varchar2(50),
    发卡人 Varchar2(20),
    领卡部门ID number(18),
    领卡人 Varchar2(20),
    发卡时间 Date,
    回收人 Varchar2(20),
    回收时间 Date,
    停用人 VARCHAR2(20),
    停用日期 DATE,
    当前状态 Number(2) DEFAULT 1,
    备注 varchar2(100),
    卡面金额 Number(16,5),
    销售金额 Number(16,5),
    充值折扣率 Number(16,5),
    余额 Number(16,5),
    发卡序号 number(18),
    回收组ID number(18),
    病人id Number(18),
    相关id Number(18),
    领用id Number(18))
    TABLESPACE zl9Expense 
;
Create Table 病人卡结算记录 (
    ID Number(18),
    接口编号 NUMBER(18),
    消费卡ID Number(18),
    序号 number(18),
    记录状态 number(18),
    结算方式 Varchar2(20),
    实收金额 Number(16,5),
    卡号 Varchar2(50),
    交易流水号 Varchar2(50),
    交易时间 DATE,
    备注 Varchar2(100),
    结算标志 number(2) DEFAULT 1,
    待转出 Number(3),
    记录性质 Number(3),
    结算id Number(18),
    应收金额 Number(16,5),
    扣率 Number(16,5),
    缴款 Number(16,5),
    找补 Number(16,5),
    缴款组id Number(18),
    缴款人姓名 Varchar2(20),
    病人id Number(18),
    单位开户行 Varchar2(50),
    单位帐号 Varchar2(50),
    结算号码 Varchar2(30),
    操作员编号 Varchar2(6),
    操作员姓名 Varchar2(20),
    登记时间 Date,
    交易说明 Varchar2(500),
    卡类别id Number(18),
    结算卡号 Varchar2(50),
    结算序号 Number(18),
    交易序号 Number(18))
    TABLESPACE zl9Expense PCTFREE 5 
;
----------------------------------------------------------------------------
--[[15.药品卫材业务]]
----------------------------------------------------------------------------
Create Table 药品采购途径(
    编码 VARCHAR2(2),
    名称 Varchar2(50),
    简码 VARCHAR2(10))
    TABLESPACE zl9BaseItem;
CREATE TABLE 药品结存汇总
(
  结存id NUMBER(18),
  入出系数 NUMBER(2),
  入出类别id NUMBER(18),
  库房id NUMBER(18),
  药品id NUMBER(18),
  批次 NUMBER(18),
  数量 NUMBER(16,5),
  金额 NUMBER(16,5),
  差价 NUMBER(16,5))
  TABLESPACE zl9MedLst
  PCTFREE 5;
Create table 未审药品记录
(收发id number(18),
单据 Number(2),
库房id number(18),
药品id Number(18),
批次 Number(18),
填制日期 date,
待转出 Number(3)) 
TABLESPACE zl9MedLst
    initrans 20;
Create Table 材料结存记录(
    Id Number(18),
    库房id Number(18),
    期初日期 Date,
    期末日期 Date,
    填制人 Varchar2(20),
    填制日期 Date,
    审核人 Varchar2(20),
    审核日期 Date,
    上次结存ID Number(18),
    期间 varchar2(6),
    性质 Number(1),
    取消人 varchar2(20),
    取消日期 date)
    TABLESPACE zl9MedLst ;
Create Table 材料结存明细(
    结存id Number(18),
    库房id Number(18),
    材料id Number(18),
    批次 Number(18),
    期初数量 Number(16,5),
    期初金额 Number(16,5),
    期初差价 Number(16,5),
    期末数量 Number(16,5),
    期末金额 Number(16,5),
    期末差价 Number(16,5))
    TABLESPACE zl9MedLst;
Create Table 材料结存误差(
    Id Number(18),
    结存id Number(18),
    库房id Number(18),
    材料id Number(18),
    批次 Number(18),
    数量差 Number(16,5),
    金额差 Number(16,5),
    差价差 Number(16,5))
    TABLESPACE zl9MedLst;
CREATE TABLE 配液台(
  id number(4),
  名称 varchar2(50),
  部门id VARCHAR2(20)
  ) TABLESPACE zl9BaseItem;
CREATE TABLE 配液台药品对照(
  配药台id number(4),
  药品id number(18),
  部门id number(18)
  ) TABLESPACE zl9BaseItem;
CREATE TABLE 配液工作安排(
  部门id number(18),
  日期 date,
  配药台id number(4),
  批次 number(18),
  审核人 varchar2(20),
  摆药人 varchar2(20),
  核对人 varchar2(20),
  配液人 varchar2(20),
  复核人 varchar2(20)
  ) TABLESPACE zl9BaseItem;
Create Table 卫材条码打印记录
(NO Varchar2(8), 
单据 number(2),
库房id Number(18), 
材料id Number(18), 
序号 Number(5), 
商品条码 Varchar2(50), 
内部条码 Varchar2(50), 
入库数量 Number(18, 5), 
打印数量 Number(18, 5), 
入库时间 Date) Tablespace Zl9medlst;
Create Table 药品设备接口(
  ID Number(18), 
  编号 Varchar2(10) Not Null, 
  名称 Varchar2(20), 
  类型 Number(2), 
  启用日期 Date,
  停用日期 Date, 
  连接信息 Varchar2(2000), 
  扩展信息 Xmltype, 
  备注 Varchar2(200)
)
Tablespace zl9BaseItem;
Create Table 药品收发门诊标志(
  处方号 Varchar2(8), 
  单据 Number(2),
  库房ID Number(18),
  业务分类 Number(2), 
  标志 Number(2),
  待转出 Number(3)
) Initrans 20
Tablespace Zl9medlst;
Create Table 药品收发住院标志(
  收发ID NUMBER(18), 
  业务分类 Number(2), 
  标志 Number(2),
  待转出 Number(3)
) Initrans 20
Tablespace Zl9medlst;
Create Table 材料质量主表
(
  id number(18),
  NO varchar2(8),
  登记人 varchar2(20),
  登记日期 date,
  处理人 varchar2(20),
  处理日期 date,
  备注  varchar2(200)
)TABLESPACE zl9MedLst;
Create Table 材料质量记录
(
  质量id number(18),
  库房id number(18),
  材料id number(18),
  批次 number(18),
  批号 varchar2(20),
  产地 varchar2(60),
  成本价 number(16,7),
  零售价 number(16,7),
  毁损数量 number(16,5),
  毁损原因 varchar2(20),
  解决办法 varchar2(20),
  供药单位id number(18)
)TABLESPACE zl9MedLst;
create table 药品入库信息 
(
       药品id number(18),
       库房id number(18),
       批次 number(18),
       入库日期 date
) tablespace ZL9MEDLST;
Create Table 药品收发主表 (
    ID number(18),
    no varchar2(8),
    单据 number(2),
    库房id number(18),
    打印状态 number(2),
    对方部门id Number(18),
    打印时间 date,
    打印人 varchar2(20))
    tablespace ZL9MEDLST 
;
Create table 入出类别对照
(
单据 number(2),
入类别id number (18),
出类别id number(18)
)TABLESPACE zl9BaseItem;
create table 药品验收记录
(
id number(18),
NO varchar2(8),
库房id number(18),
供药单位id number(18),
验收人 varchar2(200),
验收日期 date,
复核人 varchar2(200),
复核日期 date,
是否合格 number(1),
备注  varchar2(1000)
) TABLESPACE zl9MedLst;
Create Table 药品验收明细 (
    验收id number(18),
    药品id number(18),
    成本价 number(16,7),
    零售价 number(16,7),
    进药数量 number(16,5),
    批号 varchar2(20),
    生产日期 date,
    效期 date,
    产地 varchar2(60),
    批准文号 varchar2(40),
    进药日期 date,
    是否合格 number(1),
    验收结论 varchar2(100),
    导入标记 number(1))
    TABLESPACE zl9MedLst
;
Create Table 处方审查参数(
       机器名 Varchar2(15), 
       服务对象 Number(1),
       是否开启审方 Number(1), 
       最后操作时间 Date,
       来源科室 Varchar2(4000)) 
Pctfree 10 Initrans 20 
Tablespace Zl9baseitem;
Create Table 处方审查条件(
       ID Number(18), 
       类别 Number(2), 
       药名id Number(18),
       科室id Number(18), 
       医生id Number(18), 
       诊断id Number(18), 
       疾病id Number(18),
       业务 number(1)) 
Tablespace Zl9baseitem;
Create Table 处方审查项目(
       ID Number(18), 
       类别 Number(1),
       编码 Varchar2(10), 
       简称 Varchar2(50),
       内容 Varchar2(500), 
       是否门诊启用 Number(1), 
       是否住院启用 Number(1), 
       服务对象 Number(1),
       PASS结果 Varchar2(50),
       操作人 Varchar2(100), 
       操作时间 Date, 
       作废时间 Date)
Tablespace Zl9baseitem;
Create Table 处方审查常用理由(
       用户名 Varchar2(20),
       内容 Varchar2(500))
Tablespace zl9BaseItem;
Create Table 处方审查记录(
       ID Number(18),
       病人id Number(18), 
       挂号id Number(18), 
       主页id Number(18),
       提交人 Varchar2(100), 
       提交时间 Date, 
       提交科室id Number(18),
       审查结果 Number(1),
       审查人 Varchar2(100), 
       审查时间 Date, 
       发药药房id Number(18), 
       综合理由 Varchar2(500),
       状态 Number(2),
       锁定用户 Varchar2(20),
       锁定时间 Date,
       待转出 Number(3)) 
Pctfree 10 Initrans 20 
Tablespace Zl9medlst;
Create Table 处方审查明细(
       审方id Number(18), 
       医嘱id Number(18), 
       最后提交 Number(1), 
       待转出 Number(3)) 
Pctfree 10 Initrans 20 
Tablespace Zl9medlst;
Create Table 处方审查结果(
       审方id Number(18), 
       医嘱id Number(18), 
       审查项目id Number(18),
       最后提交 Number(1),
       药师审查 Number(1),
       自动审查 Number(1),
       理由 Varchar2(500),
       待转出 Number(3)) 
Pctfree 10 Initrans 20 
Tablespace Zl9medlst;
Create Table 收费调价记录(
    id NUMBER(18),
    原价ID NUMBER(18),
    收费细目id NUMBER(18),
    原价 NUMBER(16,7),
    现价 NUMBER(16,7),
    缺省价格 NUMBER(16,7),
    收入项目id NUMBER(18),
    加班加价率 NUMBER(16,5),
    附术收费率 NUMBER(16,5),
    变动原因 NUMBER(3),
    调价说明 VARCHAR2(100),
    调价ID NUMBER(18),
    填制人 VARCHAR2(20),
    填制日期 date,
    执行日期 DATE,
    终止日期 DATE,
    NO VARCHAR2(8),
    序号 NUMBER(5),
    审核人 VARCHAR2(20),
    审核日期 DATE,
    审核标志 number(1),
    说明 varchar2(200),
    价格等级 Varchar2(30))
    TABLESPACE zl9BaseItem
;
Create Table 药品批号对照 (
    药品id number(18),
    生产厂家 Varchar2(200),
    批号 varchar2(20),
    批次 number(18),
    成本价 number(16,7),
    售价 number(16,7),
    供应商ID Number(18))
    TABLESPACE zl9MedLst 
;
create table 药品财务审核
(
库房id number(18),
单据 number(2),
原始NO varchar2(8),
上次NO varchar2(8),
本次NO varchar2(8),
审核人 varchar2(100),
审核日期 date,
摘要 varchar2(1000))
TABLESPACE zl9MedLst;
create table 调价汇总记录 
(
       调价号 varchar2(10),
       类型 number(1),
       执行日期 date,
       填制日期 date,
       填制人 varchar2(20),
       说明 varchar2(100),
       分类 number(1)
) tablespace zl9MedLst;
Create Table 药品采购计划(
    ID NUMBER(18),
    No varchar2(8),
    计划类型 NUMBER(3),
    期间 VARCHAR2(8),
    库房id NUMBER(18),
    药房id NUMBER(18),
    编制方法 NUMBER(3),
    编制说明 VARCHAR2(250),
    编制人 VARCHAR2(20),
    编制日期 date,
    审核人 VARCHAR2(20),
    审核日期 date,
    复核人 VARCHAR2(20),
    复核日期 date,
    来源库房 varchar2(200),
    来源药房 varchar2(200),
    合并计划id number(18),
    采购途径 varchar2(50))
    TABLESPACE zl9MedLst PCTFREE 5 
;
Create Table 药品计划内容(
    计划ID NUMBER(18),
    药品id NUMBER(18),
    序号 NUMBER(5),
    前期数量 NUMBER(16,5),
    上期数量 NUMBER(16,5),
    上期销量 NUMBER(16,5),
    本期销量 NUMBER(16,5),
    库存数量 NUMBER(16,5),
    计划数量 NUMBER(16,5),
    执行数量 NUMBER(16,5),
    单价 NUMBER (19,7),
    金额 NUMBER(18,5),
    上次供应商 VARCHAR2(50),
    上次生产商 Varchar2(200),
    说明 Varchar2(100),
    售价 NUMBER (19,7),
    售价金额 NUMBER(18,5),
    是否上传 Number(1) Default 0,
    送货数量 number(16,5),
    批准文号 varchar2(40))
    TABLESPACE zl9MedLst PCTFREE 5 
;
Create Table 药品退药计划(
    ID NUMBER(18),
    No VARCHAR2(8),
    序号 NUMBER(5),
    药品id NUMBER(18),
    供药单位id NUMBER(18),
    实际数量 NUMBER(16,5),
    成本价 NUMBER(16,7),
    成本金额 NUMBER(16,5),
    产地 Varchar2(200),
    批号 VARCHAR2(20),
    效期 Date,
    填制人 VARCHAR2(20),
    填制日期 Date,
    摘要 VARCHAR2(100),
    审核人 VARCHAR2(20),
    审核日期 Date)
    TABLESPACE zl9MedLst PCTFREE 5 
;
Create Table 材料采购计划(
    ID NUMBER(18),
    单据 NUMBER(2) DEFAULT 0,
    No varchar2(8),
    计划类型 NUMBER(3),
    期间 VARCHAR2(6),
    库房id NUMBER(18),
    部门ID NUMBER(18),
    编制方法 NUMBER(3),
    编制说明 VARCHAR2(250),
    编制人 VARCHAR2(20),
    编制日期 date,
    审核人 VARCHAR2(20),
    审核日期 date)
    TABLESPACE zl9MedLst
    PCTFREE 5;
Create Table 材料计划内容(
    计划ID NUMBER(18),
    材料id NUMBER(18),
    序号 NUMBER(5),
    前期数量 NUMBER(16,5),
    上期数量 NUMBER(16,5),
    库存数量 NUMBER(16,5),
    请购数量 number(16,5),
    计划数量 NUMBER(16,5),
    单价 NUMBER (19,7),
    金额 NUMBER(18,5),
    上次供应商 VARCHAR2(50),
    上次生产商 VARCHAR2(60),
    上期销量 number(16,5),
    本期销量 number(16,5),
    执行数量 number(16,5))
    TABLESPACE zl9MedLst PCTFREE 5
;
Create Table 药品库存(
    库房id NUMBER(18),
    药品id NUMBER(18),
    批次 NUMBER(18),
    效期 DATE,
    性质 NUMBER(1),
    可用数量 NUMBER(18,5),
    实际数量 NUMBER(18,5),
    实际金额 NUMBER(18,5),
    实际差价 NUMBER(18,5),
    上次供应商id NUMBER(18),
    上次采购价 NUMBER(16,7),
    上次批号 VARCHAR2(20),
    上次生产日期 date,
    上次产地 Varchar2(200),
    灭菌效期 Date,
    批准文号 VARCHAR2(40),
    零售价 NUMBER(16,7),
    上次扣率 NUMBER(16,7),
    商品条码 Varchar2(50),
    内部条码 Varchar2(50),
    平均成本价 number(16,7),
    原产地 varchar2(60))
    TABLESPACE zl9MedLst initrans 20 
;
Create Table 药品结存(
    库房id NUMBER(18),
    药品id NUMBER(18),
    批次 NUMBER(18),
    填制日期 Date,
    结存日期 Date,
    实际数量 NUMBER(18,5),
    实际金额 NUMBER(18,5),
    实际差价 NUMBER(18,5),
    入库日期 Date,
    结存标志 Number(1),
	是否初始 Number(1))
    TABLESPACE zl9MedLst;
Create Table 药品留存(
    期间 varchar2(8),
    科室id NUMBER(18),
    库房id NUMBER(18),
    药品id NUMBER(18),
    可用数量 NUMBER(18,5),
    实际数量 NUMBER(18,5),
    实际金额 NUMBER(18,5))
    TABLESPACE zl9MedLst
    PCTFREE 5;
Create Table 药品留存计划(
    部门id NUMBER(18),
    库房id NUMBER(18),
    药品id NUMBER(18),
    留存数量 NUMBER(18,5),
    留存ID Number(18),
    状态 Number(1),
    登记人 Varchar2(20),
    登记时间 Date,
    待转出 Number(3),
    实际数量 number(18))
    TABLESPACE zl9MedLst
    PCTFREE 5;
CREATE TABLE 药品签名记录(
	ID NUMBER(18),
	签名规则 NUMBER(2),
	签名信息 VARCHAR2(4000),
	时间戳 DATE,
	时间戳信息 Varchar2(4000),
	证书ID	NUMBER(18),
	签名时间 DATE,
	签名人 VARCHAR2(20),
	环节 NUMBER(2),
	待转出 Number(3))
	TABLESPACE zl9MedLst;
Create TABLE 药品签名明细(
	签名ID NUMBER(18),
	收发ID NUMBER(18),
	待转出 Number(3))
	TABLESPACE zl9MedLst;
Create Table 药品收发汇总(
    日期 Date,
    库房id NUMBER(18),
    药品id NUMBER(18),
    类别id NUMBER(18),
    单据 NUMBER(2),
    数量 NUMBER(18,5),
    金额 NUMBER(18,5),
    差价 NUMBER(18,5))
    TABLESPACE zl9MedLst
    PCTFREE 5;
Create Table 未发药品记录(
    单据 NUMBER(2),
    No VARCHAR2(8),
    病人id NUMBER(18),
    主页id NUMBER(18),
    姓名 VARCHAR2(100),
    优先级 NUMBER(1),
    对方部门id NUMBER(18),
    库房id NUMBER(18),
    发药窗口 Varchar2(50),
    填制日期 Date,
    已收费 NUMBER(1),
    配药人 VARCHAR2(20),
    打印状态 NUMBER(1) Default 0,
    未发数 NUMBER(5),
    处方类型 Number(2),
    领药号 Varchar2(20),
    排队状态 Number(1),
    呼叫时间 date,
    呼叫内容 varchar2(100),
    紧急标志 Number(1),
    呼叫终端 varchar2(50))
    TABLESPACE zl9MedLst initrans 20
;
Create Table 成本价调价信息(
    Id NUMBER(18),
    收发id NUMBER(18),
    供药单位id NUMBER(18),
    库房id NUMBER(18),
    药品id NUMBER(18),
    批次 NUMBER(18),
    批号 VARCHAR2(20),
    效期 DATE,
    产地 VARCHAR2(60),
    灭菌效期 Date,
    原成本价 NUMBER(16,7),
    新成本价 NUMBER(16,7),
    发票号 VARCHAR2(200),
    发票日期 Date,
    发票金额 NUMBER(18,5),
    应付款变动 Number(1),
    执行日期 Date,
    调价汇总号 Varchar2(10))
    TABLESPACE zl9MedLst
    PCTFREE 5;
Create Table 药品价格记录(
    ID NUMBER(18),
    原价id NUMBER(18),
    价格类型 NUMBER(1),
    药品ID NUMBER(18),
    库房ID NUMBER(18),
    批次 NUMBER(18),
    原价 NUMBER(16,7),
    现价 NUMBER(16,7),
    供药单位id NUMBER(18),
    批号 VARCHAR2(20),
    效期 DATE,
    产地 Varchar2(200),
    灭菌效期 Date,
    发票号 VARCHAR2(200),
    发票日期 Date,
    发票金额 NUMBER(18,5),
    应付款变动 Number(1),
    执行日期 Date,
    终止日期 Date,
    记录状态 NUMBER(1),
    调价类型 NUMBER(1),
    调价说明 VARCHAR2(100),
    调价人 VARCHAR2(20),
    调价汇总号 Varchar2(10),
    收发id NUMBER(18),
    调价序号 number(4))
    TABLESPACE zl9MedLst PCTFREE 5 
;
Create Table 药品收发记录(
    ID NUMBER(18),
    记录状态 NUMBER(3),
    单据 NUMBER(2),
    No VARCHAR2(8),
    序号 NUMBER(5),
    库房id NUMBER(18),
    供药单位id NUMBER(18),
    入出类别id NUMBER(18),
    对方部门id NUMBER(18),
    入出系数 NUMBER(2),
    药品id NUMBER(18),
    批次 NUMBER(18),
    产地 Varchar2(200),
    批号 VARCHAR2(20),
    生产日期 date,
    效期 Date,
    付数 NUMBER(3) default 1,
    填写数量 NUMBER(16,5),
    实际数量 NUMBER(16,5),
    成本价 NUMBER(16,7),
    成本金额 NUMBER(16,5),
    扣率 NUMBER(16,7),
    零售价 NUMBER(16,7),
    零售金额 NUMBER(16,5),
    差价 NUMBER(16,5),
    摘要 VARCHAR2(1000),
    填制人 VARCHAR2(20),
    填制日期 Date,
    配药人 VARCHAR2(20),
    配药日期 DATE,
    审核人 VARCHAR2(20),
    审核日期 Date,
    价格id NUMBER(18),
    费用id NUMBER(18),
    单量 NUMBER(18,7),
    频次 VARCHAR2(20),
    用法 VARCHAR2(30),
    外观 VARCHAR2(100),
    灭菌日期 Date,
    灭菌效期 Date,
    产品合格证 VARCHAR2(100),
    发药方式 NUMBER(1),
    发药窗口 Varchar2(50),
    领用人 VARCHAR2(20),
    批准文号 VARCHAR2(40),
    汇总发药号 NUMBER(18),
    注册证号 varchar2(50),
    库房货位 Varchar2(50),
    商品条码 Varchar2(50),
    内部条码 Varchar2(50),
    核查人 Varchar2(200),
    核查日期 date,
    签到确认人 varchar2(20),
    签到时间 date,
    待转出 Number(3),
    计划id number(18),
    是否未取药 Number(1),
    取药确认人员 VARCHAR2(20),
    取药时间 Date,
    验收结论 varchar2(100),
    原产地 Varchar2(60),
    修改人 VARCHAR2(20),
    修改日期 DATE,
    紧急标志 Number(1),
    冲销原因 varchar2(200),
    采购途径 varchar2(50),
    姓名 VARCHAR2(100),
    性别 VARCHAR2(4),
    年龄 VARCHAR2(20),
    出生日期 Date,
    身份证号 VARCHAR2(18),
    病人ID NUMBER(18),
    主页ID NUMBER(5),
    病人科室id NUMBER(18),
    病人病区id NUMBER(18),
    婴儿序号 NUMBER(3),
    病人来源 NUMBER(3),
    医嘱id NUMBER(18),
    身份 VARCHAR2(10),
    处方类型 NUMBER(2),
    皮试结果 VARCHAR2(20),
    诊断描述 VARCHAR2(500),
    已收费 NUMBER(1),
    费用来源 NUMBER(1))
    TABLESPACE zl9MedLst initrans 20 
;
Create Table 收发记录补充信息(
	收发ID number(18),
	科室 varchar2(20),
	病人姓名 varchar2(100),
	住院号 number(18),
	床号 varchar2(10),
	待转出 Number(3))
    TABLESPACE zl9MedLst
    PCTFREE 5;
Create Table 暂存药品记录 (
       NO             VARCHAR2(8),
       序号           NUMBER(5),
       病人ID         Number(18),
       科室ID         Number(18),
       医嘱ID         Number(18),
       发送号         Number(18),
       药品ID         Number(18),
       药品名称       Varchar2(80),
       规格           Varchar2(100),
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
       TABLESPACE zl9CisRec;
Create Table 材料领用信息(
    收发ID NUMBER(18),
    材料ID NUMBER(18),
    病人ID NUMBER(18),
    主页ID NUMBER(18),
    姓名 VARCHAR2(100),
    性别 VARCHAR2(4),
    年龄 varchar2(20),
    床号 VARCHAR2(10),
    医疗付款方式 VARCHAR2(20),
    当前科室id NUMBER(18),
    当前病区id NUMBER(18),
    使用时间 DATE,
    条码 VARCHAR2(20))
    TABLESPACE zl9MedLst
;
Create Table 药品质量记录(
    ID NUMBER(18),
    库房id number(18),
    药品id NUMBER(18),
    批次 NUMBER(18),
    产地 VARCHAR2(60),
    批号 VARCHAR2(20),
    毁损原因 VARCHAR2(20),
    毁损数量 NUMBER(16,5),
    成本单价 number(16,7),
    成本金额 number(16,7),
    销售单价 number(16,7),
    销售金额 number(16,7),
    说明	VARCHAR2(50),
    供药单位id number(18),
    登记人 VARCHAR2(20),
    登记时间 Date,
    解决办法 VARCHAR2(20),
    处理人 VARCHAR2(20),
    处理时间 Date,
    出库单NO VARCHAR2(8))
    TABLESPACE zl9MedLst
    PCTFREE 5;
Create Table 药品结存记录(
    Id Number(18),
    库房id Number(18),
    期初日期 Date,
    期末日期 Date,
    填制人 Varchar2(20),
    填制日期 Date,
    审核人 Varchar2(20),
    审核日期 Date,
    上次结存ID Number(18),
    期间 varchar2(6),
    性质 Number(1),
    取消人 varchar2(20),
    取消日期 date)
    TABLESPACE zl9MedLst 
;
Create Table 药品结存明细(
    结存id Number(18),
    库房id Number(18),
    药品id Number(18),
    批次 Number(18),
    期初数量 Number(16,5),
    期初金额 Number(16,5),
    期初差价 Number(16,5),
    期末数量 Number(16,5),
    期末金额 Number(16,5),
    期末差价 Number(16,5))
    TABLESPACE zl9MedLst;
Create Table 药品结存误差(
    Id Number(18),
    结存id Number(18),
    库房id Number(18),
    药品id Number(18),
    批次 Number(18),
    数量差 Number(16,5),
    金额差 Number(16,5),
    差价差 Number(16,5))
    TABLESPACE zl9MedLst;
Create Table 输液配药记录(
    ID NUMBER(18),
    部门ID NUMBER(18),
    序号 NUMBER(18),
    配药批次 Number(2),
    姓名 VARCHAR2(100),
    性别 VARCHAR2(4),
    年龄 varchar2(20),
    住院号 NUMBER(18),
    床号 VARCHAR2(10),
    病人病区id NUMBER(18),
    病人科室id NUMBER(18),
    执行时间 Date,
    瓶签号 Varchar2(20),
    打印标志 number(5),
    医嘱ID Number(18),
    发送号 Number(18),
    是否打包 Number(1),
    摆药单号 NUMBER(18),
    优先级 VARCHAR2(30),
    打包时间 date,
    是否锁定 number(1),
    是否调整批次 number(1),
    手工调整批次 number(1),
    操作状态 number(2),
    操作人员 varchar2(20),
    操作时间 date,
    打印序号 number(5),
    打印时间 date,
    待转出 Number(3),
    是否确认调整 number(1),
    批次标记 number(1),
    工作人员 varchar2(100),
    打印流水号 number(5),
    配药台 varchar2(20),
    病人ID NUMBER(18),
    主页ID NUMBER(18),
    发送时间 Date)
    TABLESPACE zl9MedLst initrans 20 
;
Create Table 输液配药状态(
    配药id Number(18),
    操作类型 Number(2),
    操作人员 Varchar2(20),
    操作时间 Date,
    操作说明 Varchar2(200),
    待转出 Number(1),
    实际工作人员 varchar2(20))
    Tablespace Zl9medlst Initrans 20 
;
CREATE TABLE 输液配药附费(
    配药ID NUMBER(18),
    NO VARCHAR2(8),
	病人id number(18),
	待转出 Number(3))
    TABLESPACE zl9MedLst
    initrans 20;
Create Table 输液配药内容(
    记录ID NUMBER(18),
    收发ID NUMBER(18),
    数量 NUMBER(16,5),
	待转出 Number(3))
    TABLESPACE zl9MedLst
    initrans 20;
Create Table 库房确认记录(
    库房id NUMBER(18),
    月份 VARCHAR2(6),
    性质 NUMBER(1),  --1-药品,2-卫材
    开始时间 Date,
    终止时间 Date)
    TABLESPACE zl9MedLst;
Create Table 应付记录(
    ID number(18),
    记录性质 number(3),
    记录状态 NUMBER(3),
    NO varchar2(8),
    项目id number(18),
    序号 number(18),
    收发id NUMBER(18),
    单位ID NUMBER(18),
    库房ID Number(18),
    品名 varchar2(80),
    规格 varchar2(100),
    产地 varchar2(50),
    批号 varchar2(20),
    计量单位 varchar2(8),
    入库单据号 varchar2(8),
    单据金额 number(16,5),
    数量 number(16,5),
    采购价 number(19,5),
    采购金额 number(16,5),
    随货单号 VARCHAR2(200),
    发票号 VARCHAR2(200),
    发票日期 Date,
    发票金额 NUMBER(18,5),
    发票修改时间 Date,
    制定日期 Date,
    计划金额 number(16,5),
    计划人 varchar2(20),
    计划日期 Date,
    填制人 varchar2(20),
    填制日期 Date,
    审核人 VARCHAR2(20),
    审核日期 Date,
    摘要 varchar2(1000),
    付款序号 number(18),
    计划序号 number(18) Default 0,
    付款标志 number(1) default 0,
    预审 number(1) default 0,
    系统标识 number(1),
    发票代码 varchar2(20),
    随货日期 date)
    TABLESPACE zl9DueRec PCTFREE 5
;
Create Table 应付余额(
    单位id NUMBER(18),
    性质 NUMBER(1),
    金额 NUMBER(18,5))
    TABLESPACE zl9DueRec;
Create Table 付款记录(
    ID NUMBER(18),
    记录状态 NUMBER(3),
    No VARCHAR2(8),
    序号 NUMBER(5),
    预付款 NUMBER(1),
    单位id NUMBER(18),
    金额 NUMBER(16,5),
    结算方式 VARCHAR2(20),
    结算号码 VARCHAR2(10),
    摘要 VARCHAR2(50),
    填制人 VARCHAR2(20),
    填制日期 Date,
    预审人 VARCHAR2(20),
    预审日期 Date,
    审核人 VARCHAR2(20),
    审核日期 Date,
    付款序号 NUMBER(18),
	拒付标志 number(1) Default 0)
    TABLESPACE zl9DueRec
    PCTFREE 5;
----------------------------------------------------------------------------
--[[16.临床医嘱]]
----------------------------------------------------------------------------
CREATE TABLE 门诊病案项目(
  编码 varchar2(3),
  名称 varchar2(20),
  内容 varchar2(1000))
  TABLESPACE zl9BaseItem 
  Cache Storage(Buffer_Pool Keep);

Create Table 急诊分诊记录 (
    ID number(18),
    就诊ID number(18),
    分诊次数 number(2),
    自动病情级别 number(1),
    人工病情级别 number(1),
    人工评级说明 varchar2(100),
    修改说明 varchar2(100),
    分诊科室ID number(18),
    分诊科室名称 varchar2(100),
    收缩压 number(3),
    舒张压 number(3),
    心率 number(3),
    呼吸频率 number(3),
    指氧饱和度 number(3,1),
    体温 number(3,1),
    血糖 number(5,2),
    血钾 number(5,2),
    体征测量时间 date,
    登记人 varchar2(100),
    登记时间 date,
    待转出 Number(3))
    tablespace zl9CisRec 
;

Create Table 急诊病人评分 (
    ID number(18),
    分诊ID number(18),
    方法ID number(18),
    评分方法分值 number(5),
    评分结果描述 varchar2(100),
    病情级别 number(1),
    待转出 Number(3))
    tablespace zl9CisRec 
;

Create Table 急诊病人评分指标 (
    评分ID number(18),
    指标ID number(18),
    指标结果文本 varchar2(50),
    待转出 Number(3))
    tablespace zl9CisRec 
;

Create Table 急诊就诊记录 (
    ID number(18),
    病人ID number(18),
    病人年龄 VARCHAR2(20),
    年龄数值 number(3),
    年龄单位 VARCHAR2(4),
    挂号ID number(18),
    分诊科室ID number(18),
    保险类别 varchar2(50),
    病情级别 number(1),
    分诊病情级别 number(1),
    修订说明 varchar2(50),
    修订时间 date,
    修订人员 varchar2(100),
    到院时间 date,
    主诉 varchar2(50),
    是否三无人员 number(1),
    是否绿色通道 number(1),
    陪同人员 varchar2(10),
    病人来源 varchar2(50),
    既往病史 varchar2(500),
    意识状态 varchar2(50),
    是否成批就诊 number(1),
    成批就诊人数 number(5),
    是否复合伤 number(1),
    备注 varchar2(500),
    登记人 varchar2(100),
    登记时间 DATE,
    待转出 Number(3))
    tablespace zl9CisRec 
;

Create Table 路径通用诊疗项目(
    路径ID Number(18),
    版本号 Number(5),
    诊疗项目ID NUMBER(18))
    TABLESPACE zl9BaseItem;

Create Table 药嘱禁忌说明(
    医嘱A NUMBER(18),
    医嘱B NUMBER(18),
    禁忌类型 varchar2(50),
    操作时间 date,
    操作人员 VARCHAR2(20),
    用药说明 varchar2(500),
    待转出 Number(3))
    TABLESPACE zl9CisRec;

Create Table 病人用药清单(
  ID  NUMBER(18),
  病人ID  NUMBER(18), 
  主页ID  NUMBER(5),  
  组号  NUMBER(18),
  用药来源  NUMBER(1),
  药品类别  VARCHAR2(1),
  用药内容  VARCHAR2(1000),
  诊疗项目ID  NUMBER(18),    
  收费细目ID  NUMBER(18),    
  天数  Number(16,5),
  开始时间  DATE,
  终止时间  DATE,
  登记时间  DATE,
  登记人  VARCHAR2(20),
  总给予量   NUMBER(16,5),
  单次用量  NUMBER(16,5),   
  执行频次  VARCHAR2(20),   
  频率次数  NUMBER(3),   
  频率间隔  NUMBER(3),    
  间隔单位  VARCHAR2(4),
  用法ID  NUMBER(18),
  煎法ID  NUMBER(18),
  备注  VARCHAR2(1000),
  待转出     NUMBER(3)
)TABLESPACE zl9CisRec;
Create Table 病人用药配方(
  配方ID  NUMBER(18),
  序号  NUMBER(3),  
  诊疗项目ID  NUMBER(18),    
  收费细目ID  NUMBER(18),  
  单量  NUMBER(16,5),    
  脚注  VARCHAR2(100),
  待转出     NUMBER(3)
)TABLESPACE zl9CisRec;
Create Global Temporary Table 中联合理用药参数(参数内容 clob) On Commit Delete Rows;
Create Table 病人危急值记录(
  ID Number(18),
  数据来源 varchar2(100),    
  病人ID number(18),
  主页ID NUMBER(5),
  挂号单 VARCHAR2(8),
  婴儿 number(3),
  姓名 VARCHAR2(100),
  性别 VARCHAR2(4),
  年龄 varchar2(20),    
  医嘱ID number(18),
  标本ID NUMBER(18),   
  危急值描述 varchar2(2000),       
  报告时间 date,
  报告科室ID number(18),
  报告人 VARCHAR2(20),    
  处理情况 varchar2(2000),
  确认时间 date,          
  确认人 VARCHAR2(20),
  确认科室ID number(18),       
  状态 number(3),      
  是否危急值 number(1),  
  待转出 Number(3)
) TABLESPACE zl9CisRec;
CREATE TABLE 病人危急值医嘱(
    危急值ID NUMBER(18),
    医嘱ID NUMBER(18),
    待转出 Number(3))
    TABLESPACE zl9CisRec;
CREATE TABLE 病人危急值病历(
    危急值ID NUMBER(18),
    文档ID VARCHAR2(32),
    子文档ID VARCHAR2(32),
    标题 varchar2(100),
    完成人 varchar2(20),
    完成时间 date,
    待转出 Number(3))
    TABLESPACE zl9EprDat;    
CREATE TABLE 自定义申请单文件(
  文件ID NUMBER(18),
  文件名 VARCHAR2(200),
  类别 number(2),
  内容 CLOB,
  创建人 VARCHAR2(20),
  创建时间 DATE
  )
TABLESPACE zl9EprLob;
CREATE TABLE 医嘱申请单文件(
  医嘱ID NUMBER(18),
  文件ID NUMBER(18),
  文件名 VARCHAR2(200),
  类别 number(2),
  内容 CLOB,
  待转出 Number(3)
  )
TABLESPACE zl9EprLob;
Create Table 医嘱报告内容(
    ID Number(18),
    类型 Number(2),
    报告名 Varchar2(100),
    报告说明 Varchar2(100),
    内容 BLOB,
    创建人 VARCHAR2(20),
    创建时间 DATE,
    待转出 Number(3),
    打印次数 Number(5),
    是否禁止打印 number(1),
    来源id Varchar2(36),
    报告状态 Number(1))
    LOB(内容) Store as (Cache) Tablespace zl9CISRec PCTFREE 5 
;
Create Table 停嘱原因(编码 Varchar2(2),名称  Varchar2(50),简码 Varchar2(25), 缺省标志 Number(1)) Tablespace zl9BaseItem;
create table 传染病目录(
   编码 VARCHAR2(20),
   名称 VARCHAR2(150), 
   简码 VARCHAR2(20), 
   说明 VARCHAR2(200)
) TABLESPACE zl9BaseItem;
Create Table 疾病阳性记录(
    ID Number(18),
    病人ID number(18),
    主页id NUMBER(5),
    挂号单 VARCHAR2(8),
    送检时间 date,
    送检科室ID number(18),
    送检医生 VARCHAR2(100),
    标本名称 VARCHAR2(64),
    反馈结果 VARCHAR2(1000),
    传染病名称 VARCHAR2(200),
    检查时间 date,
    登记时间 date,
    登记人 VARCHAR2(100),
    登记科室ID number(18),
    记录状态 number(2),
    处理人 VARCHAR2(100),
    处理时间 date,
    处理情况说明 VARCHAR2(1000),
    文件ID number(18),
    待转出 Number(3),
    医嘱ID number(18))
    TABLESPACE zl9EprDat
;
create table 疾病报告反馈(
   文件ID NUMBER(18),
   登记时间 date, 
   登记人 VARCHAR2(100),
   记录状态 NUMBER(3),
   反馈内容 VARCHAR2 (500),
   处理人 VARCHAR2(100),
   处理时间 date,
   处理情况说明 VARCHAR2(500),
   待转出 Number(3)
) TABLESPACE zl9EprDat;
Create Table 输血检验结果(
  医嘱ID number(18),
  序号   number(18),
  检验项目ID number(18),
  指标代码 varchar2(20),
  指标中文名 varchar2(60),
  指标英文名 varchar2(40),
  指标结果 varchar2(500),
  结果单位 varchar2(50),
  结果标志 varchar2(10),
  结果参考 varchar2(500),
  取值序列 varchar2(4000),
  是否人工填写 number(1),
  待转出 Number(3)
) TABLESPACE zl9CisRec;
Create Table 病人医嘱记录(
    ID NUMBER(18),
    相关ID NUMBER(18),
    前提ID Number(18),
    病人来源 NUMBER(1),
    病人id NUMBER(18),
    主页id NUMBER(5),
    挂号单 VARCHAR2(8),
    婴儿 NUMBER(3),
    姓名 VARCHAR2(100),
    性别 VARCHAR2(4),
    年龄 varchar2(20),
    病人科室id NUMBER(18),
    序号 NUMBER(18),
    医嘱状态 NUMBER(3),
    医嘱期效 NUMBER(1),
    诊疗类别 VARCHAR2(1),
    诊疗项目id NUMBER(18),
    标本部位 VARCHAR2(60),
    检查方法 Varchar2(30),
    收费细目id NUMBER(18),
    天数 Number(16,5),
    单次用量 NUMBER(16,5),
    首次用量 NUMBER(16,5),
    总给予量 NUMBER(16,5),
    医嘱内容 VARCHAR2(1000),
    医生嘱托 VARCHAR2(200),
    执行科室id NUMBER(18),
    皮试结果 VARCHAR2(10),
    执行频次 VARCHAR2(20),
    频率次数 NUMBER(3),
    频率间隔 NUMBER(3),
    间隔单位 VARCHAR2(4),
    执行时间方案 VARCHAR2(100),
    计价特性 NUMBER(1),
    执行性质 NUMBER(1),
    执行标记 Number(1),
    审核标记 Number(1),
    可否分零 NUMBER(3),
    紧急标志 NUMBER(1),
    开始执行时间 DATE,
    执行终止时间 DATE,
    上次执行时间 DATE,
    上次打印时间 Date,
    开嘱科室id NUMBER(18),
    开嘱医生 VARCHAR2(41),
    开嘱时间 DATE,
    校对护士 VARCHAR2(20),
    校对时间 DATE,
    停嘱医生 VARCHAR2(201),
    停嘱时间 DATE,
    确认停嘱时间 Date,
    确认停嘱护士 Varchar2(20),
    手术时间 Date,
    是否上传 number(1),
    审查结果 Number(1),
    屏蔽打印 Number(1),
    摘要 Varchar2(1000),
    零费记帐 Number(1),
    用药目的 Number(1),
    用药理由 Varchar2(1000),
    审核状态 Number(1),
    申请序号 NUMBER(18),
    超量说明 varchar2(1000),
    是否费用审核 number(1),
    配方ID Number(18),
    手术情况 number(2),
    组合项目ID Number(18),
    重整标志 Number(1),
    新开签名ID NUMBER(18),
    待转出 Number(3),
    药师审核标志 number(1),
    药师审核时间 date,
    药师审核原因 varchar2(500),
    禁忌药品说明 varchar2(100),
    审核药师 varchar2(20),
    处方序号 Number(18),
    皮试阳性说明 VARCHAR2(1000),
    会诊医嘱ID number(18))
    TABLESPACE zl9CisRec initrans 20 
;
CREATE TABLE 病人医嘱计价(
		医嘱ID NUMBER(18),
		收费细目ID NUMBER(18),
		数量 NUMBER(16,5),
		单价 NUMBER(16,5),
		从项 Number(1),
		执行科室ID Number(18),
		费用性质 Number(1),
		收费方式 Number(1),
		待转出 Number(3))
    TABLESPACE zl9CisRec
    initrans 20;
CREATE TABLE 病人医嘱状态(
    医嘱ID NUMBER(18),
    操作类型 NUMBER(3),
    操作人员 VARCHAR2(20),
    操作时间 DATE,
	操作说明 VARCHAR2(200),
	签名ID Number(18),
	待转出 Number(3))
    TABLESPACE zl9CisRec
    initrans 20;
Create Table 病人医嘱附件(
    医嘱ID Number(18),
    项目 Varchar2(30),
    必填 Number(1),
    排列 Number(5),
    要素ID Number(18),
    内容 Varchar2(4000),
	待转出 Number(3))
    Tablespace zl9CISRec
    PCTFREE 5;
Create Table 病人医嘱报告(
    医嘱ID Number(18),
    病历ID Number(18),
    检查报告ID Raw(16),
    RISID Number(18),
    查阅状态 Number(1),
    待转出 Number(3),
    报告ID Number(18))
    Tablespace zl9CISRec PCTFREE 5
;
Create Table 报告查阅记录(
    医嘱ID Number(18),
    病历ID Number(18),
    检查报告ID Raw(16),
    查阅人 Varchar2(20),
    查阅时间 Date,
    查阅次数 Number(5),
    取消时间 Date,
    待转出 Number(3),
    RISID Number(18),
    报告ID Number(18))
    Tablespace zl9CISRec PCTFREE 5 
;
Create Table 医嘱签名记录(
    ID NUMBER(18),
    签名规则 NUMBER(2),
    签名信息 VARCHAR2(4000),
    时间戳 DATE,
    证书ID NUMBER(18),
    签名时间 DATE,
    签名人 VARCHAR2(20),
    待转出 Number(3),
    时间戳信息 Varchar2(4000))
    TABLESPACE zl9CisRec
;

create table 病人医嘱异常记录 (
    医嘱ID   NUMBER(18),
    发送号   NUMBER(18),
    NO       VARCHAR2(8),
    记录性质 NUMBER(3),
	配药ID   NUMBER(18),
    病人ID   NUMBER(18),
    产生环节 NUMBER(3),
    记录时间 Date,
    操作员姓名 varchar2(100),
    工作站 varchar2(100)
) TABLESPACE zl9CisRec initrans 20;

Create Table 病人医嘱发送(
    医嘱ID NUMBER(18),
    发送号 NUMBER(18),
    记录性质 NUMBER(3),
    门诊记帐 NUMBER(1),
    NO VARCHAR2(8),
    记录序号 NUMBER(18),
    发送数次 NUMBER(16,5),
    发送人 VARCHAR2(20),
    发送时间 DATE,
    首次时间 DATE,
    末次时间 DATE,
    安排时间 Date,
    执行状态 NUMBER(3),
    执行部门id NUMBER(18),
    完成人 Varchar2(20),
    完成时间 Date,
    计费状态 NUMBER(3),
    执行间 varchar2(20),
    报到时间 Date,
    执行过程 Number(1),
    采样人 varchar2(20),
    采样时间 DATE,
    样本条码 VARCHAR2(18),
    结果阳性 Number(1),
    执行说明 Varchar2(1000),
    接收人 varchar2(20),
    接收时间 date,
    接收批次 number(18),
    送检人 varchar2(20),
    条码打印 number(3),
    标本送出时间 Date,
    重采标本 number(1),
    待转出 Number(3),
    标本发送批号 number(18),
    状态说明 Varchar2(200),
	领药号      VARCHAR2(20)
  )    TABLESPACE zl9CisRec initrans 20 ;

CREATE TABLE 病人医嘱执行(
    医嘱ID NUMBER(18),
    发送号 NUMBER(18),
    要求时间 DATE,
    本次数次 NUMBER(16,5),
    执行摘要 VARCHAR2(200),
    执行人 VARCHAR2(20),
    执行时间 DATE,
    执行结果 number(1),
    登记人 VARCHAR2(20),
    登记时间 DATE,
    核对人 VARCHAR2(20),
    核对时间 date,
    流水号	Number(18), -- 记录哪几组医嘱一起执行的
    接单人	Varchar2(20),
    配药人	Varchar2(20),
    组数	Number(18), -- 保存本次执行一共有几组
    组次	Number(18), -- 保存次序
    滴速	Number(10,5), -- 本组的滴速
    滴系数	Number(10,5), -- 本组的滴系数
    液体量	Number(16,5), -- 药品的液体量
    耗时	Number(10), --执行完需要用的时间，单位秒
    提醒	Number(10), --提前多少时间进行提醒，单位秒,-1表示不提醒，0表示到期提醒，>0表示提前的时间
    说明 	Varchar2(200), --接单护士填写药品执行时的相关说明，如先输，避光
    配液时间 Date,
	执行科室ID Number(18),
    待转出 Number(3),
    输液通道 Varchar2(20),
	记录来源 Number(1),
	执行方式 Number(1))
    TABLESPACE zl9CisRec
    initrans 20;

Create Table 医嘱执行时间
(
要求时间 DATE,
医嘱ID NUMBER(18),
发送号 NUMBER(18),
待转出 Number(3)
)
TABLESPACE zl9CisRec
Initrans 20;
Create Table 医嘱执行计价(
    医嘱ID NUMBER(18),
    发送号 NUMBER(18),
    要求时间 DATE,
    收费细目ID NUMBER(18),
    费用性质 NUMBER(1) default(0),
    数量 NUMBER(16,5),
    待转出 Number(3),
    执行状态 number(1),
    费用id Number(18),
    执行部门ID Number(18))
    TABLESPACE zl9CisRec initrans 20 
;
Create Table 医嘱执行打印
(
医嘱ID   NUMBER(18),    
报表ID   NUMBER(18),
上次打印时间 Date,
待转出 Number(3)
)
TABLESPACE zl9CisRec
Initrans 20;
CREATE TABLE 病人执行单打印(
		病人ID Number(18),
		主页ID Number(18),
		婴儿 Number(3),
		报表ID Number(18),
		末页末行号 Number(3))
    TABLESPACE zl9CisRec;
CREATE TABLE 病人医嘱附费(
    医嘱ID NUMBER(18),
    发送号 NUMBER(18),
    记录性质 NUMBER(3),
    NO VARCHAR2(8),
	待转出 Number(3))
    TABLESPACE zl9CisRec
    PCTFREE 5
    initrans 20;
Create Table 病人医嘱打印(
    医嘱ID NUMBER(18),
    页号 NUMBER(5),
    行号 NUMBER(5),
    行数 NUMBER(5),
    病人ID Number(18),
    主页ID Number(18),
    婴儿 Number(3),
    期效 Number(1),
    打印标记 Number(1),
    打印时间 DATE,
    打印人 VARCHAR2(20),
    特殊医嘱 number(1),
    待转出 Number(3),
    医嘱内容 VARCHAR2(1000),
    类型 Number(1))
    TABLESPACE zl9CisRec 
;
CREATE TABLE 诊疗单据打印(
    记录性质 NUMBER(3),
    NO VARCHAR2(8),
	打印内容 Number(1),
	打印人 Varchar2(20),
	打印时间 Date,
	待转出 Number(3))
    TABLESPACE zl9CisRec;
Create Table 输血申请记录(
    医嘱ID number(18),
    是否待诊 number(2),
    输血性质 number(2),
    即往输血史 number(2),
    受血者属地 number(2),
    输血血型 number(2),
    RHD number(2),
    受血者血型 number(2),
    HCT number(10,2),
    ALT number(10,2),
    HBSAG number(2),
    梅毒 number(2),
    血红蛋白 number(10,2),
    血小板 number(10,2),
    ANTIHCV number(2),
    ANTIHIV12 number(2),
    待转出 Number(3),
    输血类型 varchar2(100),
    输血目的 varchar2(100),
    既往输血反应史 number(2),
    输血禁忌及过敏史 number(2),
    孕产情况 varchar2(5),
    是否签订同意书 number(1),
    是否已评估 number(1),
    允许失血量 number(4),
    当前失血量 number(4))
    TABLESPACE zl9CisRec 
;
Create Table 执行打印记录 (
       医嘱ID     Number(18),
       发送号         Number(18),
       流水号     Number(18),
       打印说明       Varchar2(1000),
       打印时间       Date,
       打印人         Varchar2(20),
	   待转出 Number(3))
       TABLESPACE zl9CisRec
       Pctfree 5;
Create Table 座位状况记录(
       病人ID         Number(18),
       科室ID         Number(18),
       分类           Varchar2(30), -- 分类，用户可自已输入
       编号           Varchar2(30), -- 座位编号
       类别           Number(1),    -- 0-普通座位 1-加座 2-特殊药品座位 3-VIP座位
       收费细目ID     Number(18), -- 如要收费，则存放对应的收费细目ID
       状态           Number(1), -- 0-空,1-在用,2-不可用,比如在维修
       类型           Number(1), -- 0-座位,1-床位
       备注           Varchar2(100),
       NO             Varchar2(8),
       呼叫器编号  varchar2(50))
       TABLESPACE zl9CisRec;
Create Table 排队记录(
    病人ID Number(18),
    科室ID Number(18),
    日期 Date Default Sysdate,
    顺序号 Number(5),
    加权号 Number(10),
    状态 Number(2),
    开始操作员 Varchar2(20),
    开始时间 Date,
    结束操作员 Varchar2(20),
    结束时间 Date,
    挂号单 Varchar2(8),
	主页ID number(18),
    呼叫标志 NUMBER(1) default 0 not null,
    备注 Varchar2(100),
    穿刺台 number(2))
    TABLESPACE zl9CisRec
;
Create Table 门诊穿刺台(
    ID Number(18),
    科室ID Number(18),
    序号 Number(2),
    有效 Number(1),
    呼叫器编号 Varchar2(50),
	待穿病人id number(18))
    TABLESPACE zl9CisRec 
;
Create Table 呼叫器日志(
	ID	Number(18),
	科室ID	Number(18),
	类别	Number(1),
	呼叫源	Varchar2(20),
	呼叫器编号	Varchar2(50),
	呼叫代码	Varchar2(20),
	呼叫时间	date,
	呼叫类别	number(2),
	医嘱ID	number(18),
	发送号	number(18),
	要求时间	date,	
	剩余液体量	number(18),
	响应人	Varchar2(20),
	响应时间	date)
	TABLESPACE zl9CisRec;
Create Table 门诊输液操作日志(
    ID Number(18),
    科室ID Number(18),
    挂号单 Varchar2(8),
    类别 Number(2),
    时间 Date,
    内容 Varchar2(4000),
    操作员 Varchar2(20),
    病人ID number(18),
    主页ID number(18))
    TABLESPACE zl9CisRec PCTFREE 5 
;
create table 业务消息清单
(
   ID Number(18),
   病人ID number(18),
   就诊ID Number(18),
   就诊科室ID Number(18),
   就诊病区ID Number(18),
   病人来源 Number(1),
   消息内容 Varchar2(4000),
   提醒场合 varchar2(50),
   类型编码 varchar2(100),
   业务标识  varchar2(4000),
   优先程度  Number(3),
   是否已阅  Number(1),
   登记时间 Date
) TABLESPACE zl9CisRec initrans 20;
create table 业务消息提醒部门
(
   消息ID Number(18),
   部门ID number(18) 
) TABLESPACE zl9CisRec initrans 20;
create table 业务消息提醒人员
(
   消息ID Number(18),
   提醒人员 varchar2(20)
) TABLESPACE zl9CisRec initrans 20;
create table 业务消息状态
(
   消息ID Number(18),
   阅读场合 number(3),
   阅读人 varchar2(20),
   阅读时间 date,
   阅读部门ID number(18)
) TABLESPACE zl9CisRec initrans 20;
Create Table RIS检查预约 (
    医嘱ID NUMBER(18),
    预约ID NUMBER(18),
    预约日期 DATE,
    检查设备ID NUMBER(18),
    检查设备名称 VARCHAR2(64),
    预约开始时间 DATE,
    预约结束时间 DATE,
    预约开始时间段 DATE,
    预约结束时间段 DATE,
    待转出 NUMBER(3),
    是否打印 NUMBER(1),
    序号 Number(18),
	预约来源 NUMBER(1),
	是否调整 NUMBER(1),
	打印时间 DATE,
	打印人 VARCHAR2(100))
    TABLESPACE zl9CisRec
;
----------------------------------------------------------------------------
--[[17.临床路径]]
----------------------------------------------------------------------------
CREATE TABLE 病人路径医嘱变异(
  路径执行ID NUMBER(18),
  医嘱内容ID NUMBER(18),
  变异原因 VARCHAR2(6),
  待转出 Number(3))
  TABLESPACE zl9CISRec;
Create Global Temporary Table 路径打印记录(
    路径执行ID NUMBER(18),
    分类 varchar2(100),
    列号 NUMBER(18),
    行号 NUMBER(18),
    内容 varchar2(1000),
    阶段ID NUMBER(18))
    On Commit Delete Rows;
Create Table 门诊变异常见原因(
    编码 VARCHAR2(6),
    名称 VARCHAR2(200),
    简码 VARCHAR2(20),
	上级 VARCHAR2(6),
	末级 NUMBER(1),
	性质 NUMBER(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 门诊路径目录(
    ID NUMBER(18),
    分类 VARCHAR2(50),
    编码 VARCHAR2(5),
    名称 VARCHAR2(100),
    通用 NUMBER(1),
    最新版本 NUMBER(3),
    适用性别 NUMBER(1),
    适用年龄 VARCHAR2(10),
    说明 VARCHAR2(200),
    最大间隔时间 NUMBER(3))
    TABLESPACE zl9BaseItem;
CREATE TABLE 门诊路径版本(
    路径ID NUMBER(18),
    版本号 NUMBER(3),
    标准治疗时间 VARCHAR2(10),
    标准费用 VARCHAR2(20),
    版本说明 VARCHAR2(200),
    创建人 VARCHAR2(20),
    创建时间 DATE,
    审核人 VARCHAR2(20),
    审核时间 DATE,
    停用人 VARCHAR2(20),
    停用时间 DATE)
    TABLESPACE zl9BaseItem;
CREATE TABLE 门诊路径分类(
    路径ID NUMBER(18),
    版本号 NUMBER(3),
    序号 NUMBER(5),
    名称 VARCHAR2(50))
    TABLESPACE zl9BaseItem;
CREATE TABLE 门诊路径科室(
    路径ID NUMBER(18),
    科室ID NUMBER(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE 门诊路径病种(
    路径ID NUMBER(18),
    疾病ID NUMBER(18),
    诊断ID NUMBER(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE 门诊路径文件(
    路径ID NUMBER(18),
    文件名 VARCHAR2(200),
    内容 BLOB,
    创建人 VARCHAR2(20),
    创建时间 DATE,
    类别 number(2))
    TABLESPACE zl9BaseItem; 
CREATE TABLE 门诊路径阶段(
    ID NUMBER(18),
    路径ID NUMBER(18),
    版本号 NUMBER(3),
    父ID NUMBER(18),
    序号 NUMBER(5),
    名称 VARCHAR2(50),
    开始天数 NUMBER(3),
    结束天数 NUMBER(3),
    分类 VARCHAR2(50),
    说明 VARCHAR2(200))
    TABLESPACE zl9BaseItem;
CREATE TABLE 门诊路径评估(
    ID NUMBER(18),
    路径ID NUMBER(18),
    版本号 NUMBER(3),
    阶段ID NUMBER(18),
    评估类型 NUMBER(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 门诊路径评估指标(
    ID NUMBER(18),
    评估ID NUMBER(18),
    序号 NUMBER(5),
    评估指标 VARCHAR2(200),
    指标类型 NUMBER(1),
    指标结果 VARCHAR2(500))
    TABLESPACE zl9BaseItem;
Create Table 门诊路径项目(
    ID NUMBER(18),
    路径ID NUMBER(18),
    版本号 NUMBER(3),
    阶段ID NUMBER(18),
    分类 VARCHAR2(50),
    项目序号 NUMBER(5),
    项目内容 VARCHAR2(1000),
    执行方式 NUMBER(1),
    项目结果 VARCHAR2(500),
    图标ID NUMBER(18),
    导入参考 varchar2(1500),
    导入结果 number(1),
    内容要求 NUMBER(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 门诊路径评估条件(
    评估ID NUMBER(18),
    指标ID NUMBER(18),
    项目ID NUMBER(18),
    关系式 VARCHAR2(5),
    条件值 VARCHAR2(50),
    条件组合 NUMBER(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 门诊路径医嘱内容(
    ID NUMBER(18),
    相关ID NUMBER(18),
    序号 NUMBER(5),
    期效 NUMBER(1),
    诊疗项目ID NUMBER(18),
    收费细目ID NUMBER(18),
    医嘱内容 VARCHAR2(1000),
    单次用量 NUMBER(16,5),
    总给予量 NUMBER(16,5),
    标本部位 VARCHAR2(60),
    检查方法 VARCHAR2(30),
    医生嘱托 VARCHAR2(1000),
    执行频次 VARCHAR2(20),
    频率次数 NUMBER(3),
    频率间隔 NUMBER(3),
    间隔单位 VARCHAR2(4),
    执行性质 NUMBER(1),
    执行标记 NUMBER(1),
    执行科室ID NUMBER(18),
    时间方案 VARCHAR2(50),
    是否缺省 Number(1) Default 0,
    是否备选 number(1) default(0),
    配方ID Number(18),
    组合项目ID Number(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE 门诊路径医嘱(
    路径项目ID NUMBER(18),
    医嘱内容ID NUMBER(18))
    TABLESPACE zl9BaseItem;
Create Table 门诊路径病历(
    项目ID NUMBER(18),
    文件ID NUMBER(18),
    原型ID VARCHAR2(32),
    名称 varchar2(100),
    序号 Number(5))
    TABLESPACE zl9BaseItem;
CREATE TABLE 门诊路径医嘱变动(
    项目ID  NUMBER(18),
    操作时间  Date,
    操作员  VARCHAR2(100),
    医嘱内容ID  NUMBER(18),
    相关ID NUMBER(18),
    序号 NUMBER(5),
    期效 NUMBER(1),
    诊疗项目ID NUMBER(18),
    收费细目ID NUMBER(18),
    医嘱内容 VARCHAR2(1000),
    单次用量 NUMBER(16,5),
    总给予量 NUMBER(16,5),
    标本部位 VARCHAR2(60),
    检查方法 VARCHAR2(30),
    医生嘱托 VARCHAR2(1000),
    执行频次 VARCHAR2(20),
    频率次数 NUMBER(3),
    频率间隔 NUMBER(3),
    间隔单位 VARCHAR2(4),
    执行性质 NUMBER(1),
    执行标记 NUMBER(1),
    执行科室ID NUMBER(18),
    时间方案 VARCHAR2(50),
    是否缺省 Number(1) Default 0,
    是否备选 number(1) default 0,
    配方ID Number(18),
    组合项目ID Number(18),
    审核状态 number(1),
    审核人 varchar2(100),
    审核时间 date)
     TABLESPACE zl9CISRec;
Create Table 标准门诊路径目录(
     ID NUMBER(18),
     科室名称 Varchar2(100),
     编码 Varchar2(8),   
     路径名称 Varchar2(80),
     类别  NUMBER(2),
     版本说明 Varchar2(20))
    TABLESPACE ZL9BASEITEM;
Create Table 标准门诊路径病种(
    标准路径id NUMBER(18),
    疾病编码   VARCHAR2(100),
    手术编码   VARCHAR2(100))
    TABLESPACE ZL9BASEITEM;
create table 标准门诊路径表单(
    标准路径id NUMBER(18),
    表单序号   NUMBER(3),
    表单名称   VARCHAR2(100),
    表单表头   Varchar2(500),
    分类序号   NUMBER(3),
    分类名称   VARCHAR2(50),
    阶段序号   NUMBER(3),
    阶段名称   VARCHAR2(100),
    路径内容   VARCHAR2(2000))
    tablespace ZL9BASEITEM;
create table 标准门诊路径流程(
    标准路径id NUMBER(18) ,
    序号     NUMBER(3) ,
    标题     VARCHAR2(100),
    内容     VARCHAR2(4000))
    tablespace ZL9BASEITEM;
CREATE TABLE 病人门诊路径(
    ID NUMBER(18),
    病人ID NUMBER(18),
    挂号ID NUMBER(18),
    科室ID NUMBER(18),
    路径ID NUMBER(18),
    版本号 NUMBER(3),
    导入人 VARCHAR2(20),
    导入时间 DATE,
    导入说明 VARCHAR2(1000),
    未导入原因 Varchar2(6),
    开始时间 DATE,
    结束时间 DATE,
    状态 NUMBER(1),
    当前天数   NUMBER(18),
    当前阶段ID NUMBER(18),
    前一阶段ID NUMBER(18),
    诊断类型 NUMBER(2),
    诊断来源 NUMBER(1),
    疾病ID NUMBER(18),
    诊断ID NUMBER(18),
    待转出 Number(3))
    TABLESPACE zl9CISRec
    PCTFREE 5;
CREATE TABLE 病人门诊路径记录(
    路径记录ID NUMBER(18),
    挂号ID NUMBER(18),
    待转出 Number(3))   
    TABLESPACE zl9CISRec
    PCTFREE 5;
CREATE TABLE 病人门诊路径评估(
    路径记录ID NUMBER(18),
    阶段ID NUMBER(18),
    日期 DATE,
    天数 NUMBER(5),
    评估人 VARCHAR2(50),
    评估时间 DATE,
    评估结果 NUMBER(2),
    评估说明 VARCHAR2(1000),
    变异原因 Varchar2(6),
    时间进度 Number(1) Default 0,
    登记人 VARCHAR2(20),
    登记时间 DATE,
    变异审核人 Varchar2(20),
    变异审核时间 Date,
    待转出 Number(3))
    TABLESPACE zl9CISRec
    PCTFREE 5;
CREATE TABLE 病人门诊路径变异(
    路径记录ID NUMBER(18),
    阶段ID NUMBER(18),
    日期 DATE,
    变异原因 VARCHAR2(6),
    待转出 Number(3))
    TABLESPACE zl9CISRec;
Create Table 病人门诊路径执行(
    ID NUMBER(18),
    路径记录ID NUMBER(18),
    阶段ID NUMBER(18),
    日期 DATE,
    天数 NUMBER(5),
    分类 VARCHAR2(50),
    项目ID NUMBER(18),
    项目序号 NUMBER(5),
    项目内容 VARCHAR2(1000),
    项目结果 VARCHAR2(500),
    变异原因 Varchar2(6),
    添加原因 VARCHAR2(1000),
    图标ID NUMBER(18),
    执行人 VARCHAR2(20),
    执行时间 DATE,
    执行结果 VARCHAR2(50),
    执行说明 VARCHAR2(200),
    登记人 VARCHAR2(20),
    登记时间 DATE,
    待转出 Number(3))
    TABLESPACE zl9CISRec PCTFREE 5;
CREATE TABLE 病人门诊路径指标(
    路径记录ID NUMBER(18),
    阶段ID NUMBER(18),
    日期 DATE,
    天数 NUMBER(5),
    评估类型 NUMBER(1),
    评估指标 VARCHAR2(200),
    指标类型 NUMBER(1),
    指标结果 VARCHAR2(50),
    待转出 Number(3))
    TABLESPACE zl9CISRec
    PCTFREE 5;
CREATE TABLE 病人门诊路径医嘱(
    路径执行ID NUMBER(18),
    病人医嘱ID NUMBER(18),
    待转出 Number(3))
    TABLESPACE zl9CISRec
    PCTFREE 5;
CREATE TABLE 门诊路径报表目录(
	ID		NUMBER(18),
	编码    VARCHAR2(64),
	名称	VARCHAR2(100),
	是否固定 NUMBER(1)
	)
	TABLESPACE zl9CISRec;
CREATE TABLE 门诊路径报表结构(		
	报表ID	NUMBER(18),
	行号	NUMBER(5),
	项目序号	NUMBER(5),
	项目文本1 VARCHAR2(100),
	项目文本2 VARCHAR2(100),
	SQL文本 VARCHAR2(4000),
	页数 number(3),
	路径ID number(18),
	多选序号 number(5)
	)
    TABLESPACE zl9CISRec;
Create Table 门诊路径报表序号 (
   报表ID number(18),
   行号  NUMBER(5),
   路径ID number(18),
   序号 Number(8)
) TABLESPACE zl9CISRec;
CREATE TABLE 门诊路径报表文件(
	ID		 NUMBER(18),	
	报表ID	 NUMBER(18),
	期间	 VARCHAR2(20),
	开始时间 DATE,
	结束时间 DATE,
	路径ID	 NUMBER(18),	
	填写人	 VARCHAR2(20),	
	填写时间 DATE
	)
    TABLESPACE zl9CISRec;
CREATE TABLE 门诊路径报表记录(
	文件ID	NUMBER(18),	
	行号	NUMBER(3),
	项目值	VARCHAR2(100),
	备注	VARCHAR2(1000)
	)
    TABLESPACE zl9CISRec;
CREATE TABLE 病人门诊出径记录(
	病人ID		NUMBER(18),
	挂号ID		NUMBER(18),
	行号		NUMBER(5),	
	路径记录ID	number(18),
	数字值		NUMBER(18),
	字符值		VARCHAR2(100),
	日期值		Date,
	备注		VARCHAR2(1000),
	登记人		VARCHAR2(20),
	登记时间	DATE,
	待转出 Number(3)
	)
TABLESPACE zl9CISRec;
CREATE TABLE 病人门诊路径取消(
    操作时间 Date,
    操作人  VARCHAR2(20),
    审核人  VARCHAR2(20),
    病人ID    NUMBER(18),
    挂号ID    NUMBER(18))
    TABLESPACE zl9CISRec;
create table 病人路径病历
(
路径执行ID  Number(18),
任务ID     varchar2(32)
)
TABLESPACE zl9CISRec;
Create Table 临床路径目录(
    ID NUMBER(18),
    分类 VARCHAR2(50),
    编码 VARCHAR2(5),
    名称 VARCHAR2(100),
    通用 NUMBER(1),
    最新版本 NUMBER(3),
    病例分型 VARCHAR2(20),
    适用病情 VARCHAR2(20),
    适用性别 NUMBER(1),
    适用年龄 VARCHAR2(10),
    说明 VARCHAR2(200),
    确诊天数 NUMBER(3),
    结束路径控制 number(1),
    性质 NUMBER(1) default(0),
    变异系数 NUMBER(3),
    审核状态 NUMBER (3),
    是否循环 Number(1))
    TABLESPACE zl9BaseItem 
;
CREATE TABLE 临床路径分支(
    ID  NUMBER(18),
    路径ID NUMBER(18),
    版本号 NUMBER(3),
    名称  VARCHAR2(50),
    说明  VARCHAR2(200),
    前一阶段ID NUMBER(18),
    标准住院日 VARCHAR2(10),
    标准费用 VARCHAR2(20),
    创建人 VARCHAR2(20),
    创建时间 DATE)
    TABLESPACE zl9BaseItem;
CREATE TABLE 临床路径病种(
    路径ID NUMBER(18),
    疾病ID NUMBER(18),
	诊断ID NUMBER(18),
	性质 number(2))
    TABLESPACE zl9BaseItem;
CREATE TABLE 临床路径科室(
    路径ID NUMBER(18),
    科室ID NUMBER(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE 临床路径文件(
    路径ID NUMBER(18),
	文件名 VARCHAR2(200),
    内容 BLOB,
	创建人 VARCHAR2(20),
	创建时间 DATE,
	类别 number(2)
	)
    TABLESPACE zl9BaseItem;
CREATE TABLE 临床路径版本(
    路径ID NUMBER(18),
    版本号 NUMBER(3),
    标准住院日 VARCHAR2(10),
    标准费用 VARCHAR2(20),
    版本说明 VARCHAR2(200),
    创建人 VARCHAR2(20),
    创建时间 DATE,
    审核人 VARCHAR2(20),
    审核时间 DATE,
	停用人 VARCHAR2(20),
    停用时间 DATE,
	药剂科审核人 VARCHAR2(20),
	药剂科审核时间 DATE)
    TABLESPACE zl9BaseItem;
CREATE TABLE 临床路径阶段(
		ID NUMBER(18),
    路径ID NUMBER(18),
    版本号 NUMBER(3),
		父ID NUMBER(18),
	分支ID NUMBER(18),
    序号 NUMBER(5),
    名称 VARCHAR2(50),
    开始天数 NUMBER(3),
    结束天数 NUMBER(3),
    标志 VARCHAR2(10),
		分类 VARCHAR2(50),
    说明 VARCHAR2(200))
    TABLESPACE zl9BaseItem;
CREATE TABLE 临床路径分类(
    路径ID NUMBER(18),
    版本号 NUMBER(3),
    序号 NUMBER(5),
		名称 VARCHAR2(50),
	分支ID NUMBER(18))
    TABLESPACE zl9BaseItem;
Create Table 临床路径项目(
    ID NUMBER(18),
    路径ID NUMBER(18),
    版本号 NUMBER(3),
    阶段ID NUMBER(18),
    分支ID NUMBER(18),
    分类 VARCHAR2(50),
    项目序号 NUMBER(5),
    项目内容 VARCHAR2(1000),
    执行方式 NUMBER(1),
    执行者 NUMBER(1),
    项目结果 VARCHAR2(500),
    图标ID NUMBER(18),
    导入参考 varchar2(1500),
    导入结果 number(1),
    内容要求 NUMBER(1),
    生成者 number(1))
    TABLESPACE zl9BaseItem
;
Create Table 路径医嘱变动(
    项目ID NUMBER(18),
    操作时间 Date,
    操作员 VARCHAR2(100),
    医嘱内容ID NUMBER(18),
    相关ID NUMBER(18),
    序号 NUMBER(5),
    期效 NUMBER(1),
    诊疗项目ID NUMBER(18),
    收费细目ID NUMBER(18),
    医嘱内容 VARCHAR2(1000),
    单次用量 NUMBER(16,5),
    总给予量 NUMBER(16,5),
    标本部位 VARCHAR2(60),
    检查方法 VARCHAR2(30),
    医生嘱托 VARCHAR2(1000),
    执行频次 VARCHAR2(20),
    频率次数 NUMBER(3),
    频率间隔 NUMBER(3),
    间隔单位 VARCHAR2(4),
    执行性质 NUMBER(1),
    执行标记 NUMBER(1),
    执行科室ID NUMBER(18),
    时间方案 VARCHAR2(50),
    是否必选 number(1) default 0,
    是否缺省 Number(1) Default 0,
    是否备选 number(1) default 0,
    配方ID Number(18),
    组合项目ID Number(18),
    审核状态 number(1),
    审核人 varchar2(100),
    审核时间 date,
    药剂审核人 VARCHAR2(100),
    药剂审核时间 date,
    父项ID number(18))
    TABLESPACE zl9CISRec 
;
Create Table 路径医嘱内容(
    ID NUMBER(18),
    相关ID NUMBER(18),
    序号 NUMBER(5),
    期效 NUMBER(1),
    诊疗项目ID NUMBER(18),
    收费细目ID NUMBER(18),
    医嘱内容 VARCHAR2(1000),
    单次用量 NUMBER(16,5),
    总给予量 NUMBER(16,5),
    标本部位 VARCHAR2(60),
    检查方法 VARCHAR2(30),
    医生嘱托 VARCHAR2(1000),
    执行频次 VARCHAR2(20),
    频率次数 NUMBER(3),
    频率间隔 NUMBER(3),
    间隔单位 VARCHAR2(4),
    执行性质 NUMBER(1),
    执行标记 NUMBER(1),
    执行科室ID NUMBER(18),
    时间方案 VARCHAR2(50),
    是否必选 number(1) default 0,
    是否缺省 Number(1) Default 0,
    是否备选 number(1) default(0),
    配方ID Number(18),
    组合项目ID Number(18),
    父项ID number(18))
    TABLESPACE zl9BaseItem
;
CREATE TABLE 临床路径医嘱(
		路径项目ID NUMBER(18),
    医嘱内容ID NUMBER(18))
    TABLESPACE zl9BaseItem;
Create Table 临床路径病历(
    项目ID NUMBER(18),
    文件ID NUMBER(18),
    原型ID VARCHAR2(32),
    名称 varchar2(100),
    序号 Number(5))
    TABLESPACE zl9BaseItem;
CREATE TABLE 临床路径评估(
		ID NUMBER(18),
    路径ID NUMBER(18),
    版本号 NUMBER(3),
		阶段ID NUMBER(18),
		评估类型 NUMBER(1),
	分支ID NUMBER(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE 路径评估指标(
		ID NUMBER(18),
    评估ID NUMBER(18),
    序号 NUMBER(5),
		评估指标 VARCHAR2(200),
		指标类型 NUMBER(1),
		指标结果 VARCHAR2(500))
    TABLESPACE zl9BaseItem;
CREATE TABLE 路径评估条件(
		评估ID NUMBER(18),
    指标ID NUMBER(18),
    项目ID NUMBER(18),
		关系式 VARCHAR2(5),
		条件值 VARCHAR2(50),
		条件组合 NUMBER(1))
    TABLESPACE zl9BaseItem;
CREATE TABLE 病人临床路径(
		ID NUMBER(18),
		病人ID NUMBER(18),
		主页ID NUMBER(5),
		科室ID NUMBER(18),
		路径ID NUMBER(18),
		版本号 NUMBER(3),
		导入人 VARCHAR2(20),
		导入时间 DATE,
		导入说明 VARCHAR2(1000),
		未导入原因 Varchar2(6),
		开始时间 DATE,
		结束时间 DATE,
		状态 NUMBER(1),
		当前天数   NUMBER(18),
		当前阶段ID NUMBER(18),
		前一阶段ID NUMBER(18),
		诊断类型 NUMBER(2),
		诊断来源 NUMBER(1),
		疾病ID NUMBER(18),
		诊断ID NUMBER(18),
        合并路径个数 NUMBER(2),
		待转出 Number(3))
    TABLESPACE zl9CISRec
    PCTFREE 5;
Create Table 病人合并路径(
    ID         NUMBER(18),      
    病人ID     NUMBER(18),
    主页ID     NUMBER(5),
    科室ID     NUMBER(18),
    路径ID     NUMBER(18),
    版本号     NUMBER(3),
    导入人     VARCHAR2(20),
    导入时间     DATE,
    导入说明     VARCHAR2(1000),
    当前天数     NUMBER(18),
    当前阶段ID   NUMBER(18),
    前一阶段ID   NUMBER(18),
    诊断类型     NUMBER(2),
    诊断来源     NUMBER(1),
    疾病ID     NUMBER(18),
    首要路径记录ID  NUMBER(18),
    首要路径阶段ID NUMBER(18),
    首要路径天数   NUMBER(18),
    结束时间     DATE,
	待转出 Number(3)) 
TABLESPACE zl9CISRec;
Create Table 病人合并路径评估(
    路径记录ID  NUMBER(18),      
    阶段ID NUMBER(18),
	日期 DATE,
    合并路径记录ID  NUMBER(18),
    合并路径阶段ID NUMBER(18),
    合并路径天数   NUMBER(18),
    登记时间 date,
	待转出 Number(3)) 
TABLESPACE zl9CISRec;
CREATE TABLE 病人路径变异(
	路径记录ID NUMBER(18),
	阶段ID NUMBER(18),
	日期 DATE,
    	变异原因 VARCHAR2(6),
	待转出 Number(3))
    TABLESPACE zl9CISRec;
Create Table 病人路径执行(
    ID NUMBER(18),
    路径记录ID NUMBER(18),
    阶段ID NUMBER(18),
    日期 DATE,
    天数 NUMBER(5),
    分类 VARCHAR2(50),
    项目ID NUMBER(18),
    项目序号 NUMBER(5),
    项目内容 VARCHAR2(1000),
    执行者 NUMBER(1),
    项目结果 VARCHAR2(500),
    变异原因 Varchar2(6),
    添加原因 VARCHAR2(1000),
    图标ID NUMBER(18),
    执行人 VARCHAR2(20),
    执行时间 DATE,
    执行结果 VARCHAR2(50),
    执行说明 VARCHAR2(200),
    登记人 VARCHAR2(20),
    登记时间 DATE,
    合并路径记录ID NUMBER(18),
    合并路径阶段ID NUMBER(18),
    生成者 number(1),
    生成时间性质 number(1),
    待转出 Number(3)
    )
    TABLESPACE zl9CISRec PCTFREE 5
;
CREATE TABLE 病人路径评估(
		路径记录ID NUMBER(18),
		阶段ID NUMBER(18),
		日期 DATE,
		天数 NUMBER(5),
		评估人 VARCHAR2(50),
		评估时间 DATE,
		评估结果 NUMBER(2),
		评估说明 VARCHAR2(1000),
		变异原因 Varchar2(6),
		时间进度 Number(1) Default 0,
		登记人 VARCHAR2(20),
		登记时间 DATE,
		变异审核人 Varchar2(20),
		变异审核时间 Date,
		跳转审核人 varchar2(20),
		跳转审核时间 date,
		原路径ID NUMBER(18),
		原路径版本 NUMBER(3),
		待转出 Number(3))
    TABLESPACE zl9CISRec
    PCTFREE 5;
CREATE TABLE 病人路径指标(
		路径记录ID NUMBER(18),
		阶段ID NUMBER(18),
		日期 DATE,
		天数 NUMBER(5),
		评估类型 NUMBER(1),
		评估指标 VARCHAR2(200),
		指标类型 NUMBER(1),
		指标结果 VARCHAR2(50),
		合并路径记录ID Number(18),
		待转出 Number(3))
    TABLESPACE zl9CISRec
    PCTFREE 5;
CREATE TABLE 病人路径医嘱(
	路径执行ID NUMBER(18),
    病人医嘱ID NUMBER(18),
	待转出 Number(3))
    TABLESPACE zl9CISRec
    PCTFREE 5;
CREATE TABLE 病人出径记录(
	病人ID		NUMBER(18),
	主页ID		NUMBER(18),
	行号		NUMBER(5),	
    路径记录ID  number(18),
	数字值		NUMBER(18),
	字符值		VARCHAR2(100),
	日期值		Date,
	备注		VARCHAR2(1000),
	登记人		VARCHAR2(20),
	登记时间	DATE,
	待转出 Number(3)
	)
TABLESPACE zl9CISRec;
CREATE TABLE 病人路径取消(
  操作时间 Date,
  操作人  VARCHAR2(20),
  审核人  VARCHAR2(20),
  病人ID    NUMBER(18),
  主页ID    NUMBER(18)
  )
TABLESPACE zl9CISRec;
CREATE TABLE 路径报表文件(
	ID		 NUMBER(18),	
	报表ID	 NUMBER(18),
	期间	 VARCHAR2(20),
	开始时间 DATE,
	结束时间 DATE,
	路径ID	 NUMBER(18),	
	填写人	 VARCHAR2(20),	
	填写时间 DATE
	)
    TABLESPACE zl9CISRec;
CREATE TABLE 路径报表记录(
	文件ID	NUMBER(18),	
	行号	NUMBER(3),
	项目值	VARCHAR2(100),
	备注	VARCHAR2(1000)
	)
    TABLESPACE zl9CISRec;
----------------------------------------------------------------------------
--[[18.病历业务]]
----------------------------------------------------------------------------
Create Table 电子病历记录(
    ID NUMBER(18),
    序号 NUMBER(4),
    病人来源 NUMBER(3),
    病人ID NUMBER(18),
    主页ID NUMBER(18),
    婴儿 NUMBER(5),
    科室ID NUMBER(18),
    病历种类 NUMBER(3),
    文件ID NUMBER(18),
    病历名称 VARCHAR2(30),
    创建人 VARCHAR2(20),
    创建时间 DATE,
    完成时间 DATE,
    保存人 VARCHAR2(20),
    保存时间 Date,
    最后版本 NUMBER(5),
    签名级别 NUMBER(1),
    归档人 VARCHAR2(20),
    归档日期 DATE,
    处理状态 NUMBER(3),
    打印人 Varchar2(20),
    打印时间 Date,
    编辑方式 Number(1) Default 0,
    路径执行ID Number(18),
    待转出 Number(3),
    门诊路径执行ID Number(18),
	输出PDF状态 Number(2))
    TABLESPACE zl9EprDat
    Initrans 20
;
CREATE TABLE 电子病历格式(
    文件ID NUMBER(18),
    内容 BLOB,
    文本内容 CLOB,
	待转出 Number(3))
	LOB(内容) Store as (Cache)
    TABLESPACE zl9EprLob
    PCTFREE 20
    Initrans 20
;
CREATE TABLE 电子病历附件(
    病历ID NUMBER(18),
    序号 NUMBER(5),
    文件名 VARCHAR2(50),
    内容 BLOB,
    大小 NUMBER(12,2),
    创建人 VARCHAR2(20),
    日期 Date,
	待转出 Number(3))
	LOB(内容) Store as (Cache)
    TABLESPACE zl9EprLob
    Initrans 20
;
CREATE TABLE 电子病历内容(
    ID NUMBER(18),
    文件ID NUMBER(18),
    开始版 NUMBER(5),
    终止版 NUMBER(5),
    父ID NUMBER(18),
    对象序号 NUMBER(18),
    对象类型 NUMBER(1),
    对象标记 NUMBER(18),
    保留对象 NUMBER(1),
    对象属性 VARCHAR2(1000),
    内容行次 NUMBER(18),
    内容文本 VARCHAR2(4000),
    是否换行 NUMBER(1),
    预制提纲ID NUMBER(18),
		定义提纲ID Number(18),
    复用提纲 NUMBER(1),
    使用时机 VARCHAR2(2),
    诊治要素ID NUMBER(18),
		替换域 NUMBER(1),
    要素名称 VARCHAR2(40),
    要素类型 NUMBER(3),
    要素长度 NUMBER(3),
    要素小数 NUMBER(3),
    要素单位 VARCHAR2(50),
    要素表示 NUMBER(3),
    输入形态 NUMBER(3),
    要素值域 VARCHAR2(4000),
	待转出 Number(3))
    TABLESPACE zl9EprDat
    Initrans 20
;
CREATE TABLE 电子病历图形(
    对象ID NUMBER(18),
    图形 BLOB,
	待转出 Number(3))
	LOB(图形) Store as (Cache)
    TABLESPACE zl9EprLob
    Initrans 20
;
CREATE TABLE 病历变动原因(
	ID      Number(18),
	病历文件id  Number(18),
	变动原因  Number(1),
	原因要件id  Number(18),
	原因要素  Varchar2(40),
	原因内容  Varchar2(50))
	TABLESPACE zl9EprDat;
CREATE TABLE 病历变动结果(
	ID          Number(18),
	变动原因id  Number(18),
	变动结果    Number(1),
	病历提纲id  Number(18),
	结果要件id  Number(18),
	结果要素    Varchar2(40),
	结果值域  Varchar2(500),
	原始值域  Varchar2(500))
	TABLESPACE zl9EprDat;
Create Table 电子病历时机(
    ID        Number(18),
    病人ID    Number(18),
    主页ID    Number(18),
    病人来源  Number(1),
    科室ID    Number(18),
    责任人    Varchar2(64),
    文件ID    Number(18),
    病历种类  Number(3),
    病历编号  Varchar2(3),
    病历名称  Varchar2(30),
    事件      Varchar2(1000),
    必须      Number(1),
    唯一      Number(1),
    事件时间   Date,
    开始时间   Date,
    到期时间   Date,
    一般周期   Number(5),
    病重周期   Number(5),
    病危周期   Number(5),
    周期号     Number(5),
    完成记录ID Number(18),
    完成时间   Date)
    TABLESPACE zl9EprDat
	PCTFREE 20 initrans 20;
CREATE TABLE 电子病历打印(
    ID NUMBER(18),
    文件ID NUMBER(18),
    种类 NUMBER(3),
    病人ID NUMBER(18),
    主页ID NUMBER(18),
    打印人 Varchar2(64),
    打印时间	Date)
    TABLESPACE zl9EprDat
    Initrans 20
;
Create Table 疾病申报记录(
    文件ID NUMBER(18),
    处理状态 NUMBER(3),
    收拒人 VARCHAR2(20),
    收拒时间 DATE,
    收拒说明 VARCHAR2(100),
    报送人 VARCHAR2(20),
    报送时间 DATE,
    报送单位 VARCHAR2(30),
    报送备注 VARCHAR2(100),
    登记人 VARCHAR2(20),
    登记时间 DATE,
    姓名 VARCHAR2(100),
    性别 VARCHAR2(4),
    年龄 varchar2(20),
    职业 VARCHAR2(80),
    家庭地址 VARCHAR2(100),
    家庭电话 VARCHAR2(20),
    发病日期 DATE,
    确诊日期 DATE,
    诊断描述1 VARCHAR2(150),
    诊断描述2 VARCHAR2(150),
    填报备注 VARCHAR2(100),
    文档ID Varchar2(32),
    待转出 Number(3),
    报卡类型 VARCHAR2(50),
    报告医生 VARCHAR2(100),
    撤档人 VARCHAR2(100),
    撤档时间 Date,
	病人ID NUMBER(18),
	主页ID NUMBER(18),
	病人来源 NUMBER(3))
    TABLESPACE zl9EprDat;
Create Table 疾病申报对应(
    申报项目 VARCHAR2(30),
    对应要素 VARCHAR2(40))
    TABLESPACE zl9BaseItem;
Create Table 疾病申报反馈(
	申报ID Number(18),
	反馈信息 VARCHAR2(500),
	登记人 VARCHAR2(20),
	登记时间 date,
	处理情况说明 VARCHAR2(500),
	待转出 Number(3))
	TABLESPACE zl9EprDat;
Create Global Temporary Table 临时病历内容(
		ID NUMBER(18),
		文件ID NUMBER(18),
		父ID NUMBER(18),
		对象序号 NUMBER(18),
		对象类型 NUMBER(1),
		对象标记 NUMBER(18),
		保留对象 NUMBER(1),
		对象属性 VARCHAR2(1000),
		开始版 NUMBER(5),
		终止版 NUMBER(5),
		内容行次 NUMBER(18),
		内容文本 VARCHAR2(4000),
		是否换行 NUMBER(1),
		预制提纲ID NUMBER(18),
		定义提纲ID Number(18),
		复用提纲 NUMBER(1),
		使用时机 VARCHAR2(2),
		诊治要素ID NUMBER(18),
		替换域 NUMBER(1),
		要素名称 VARCHAR2(40),
		要素类型 NUMBER(3),
		要素长度 NUMBER(3),
		要素小数 NUMBER(3),
		要素单位 VARCHAR2(50),
		要素表示 NUMBER(3),
		输入形态 NUMBER(3),
		要素值域 VARCHAR2(4000))
    On Commit Delete Rows;
Create Global Temporary Table 病历时限监测(
    病人ID NUMBER(18),
    主页ID NUMBER(18),
    病人来源 NUMBER(3),
    变动事件 VARCHAR2(40),
    事件时间 DATE,
    文件ID NUMBER(18),
    病历种类 NUMBER(3),
    病历编号 VARCHAR2(3),
    病历名称 VARCHAR2(30),
    唯一 NUMBER(1),
    科室ID NUMBER(18),
    责任人 VARCHAR2(20),
    周期号 NUMBER(3),
    基点时间 DATE,
    要求时间 DATE,
    完成记录ID NUMBER(18),
    完成时间 DATE)
	On Commit Preserve Rows;
Create Global Temporary Table 病历内容监测(
    病人id NUMBER(18),
    主页id NUMBER(18),
    病人来源 NUMBER(3),
    病历记录id NUMBER(18),
    病历种类 NUMBER(3),
    病历名称 VARCHAR2(30),
    完成日期 DATE,
    提纲id NUMBER(18),
    提纲父id NUMBER(18),
    提纲层次 NUMBER(3),
    提纲序号 NUMBER(5),
    提纲文本 VARCHAR2(200),
    提示级别 NUMBER(1),	--说明：0-无;1-提示;2-严重
    提示内容 VARCHAR2(4000))
	On Commit Preserve Rows;
--审查归档
Create Table 病案提交记录(
    ID			Number(18),
    病人id		Number(18),
    主页id		Number(5),
    记录状态	Number(3),
    提交人		Varchar2(20),
    提交时间	Date,
    接收人		Varchar2(20),
    接收时间	Date,
    归档人		Varchar2(20),
    归档时间	Date,
    拒审人		Varchar2(20),
    拒审时间	Date,
    拒审理由	Varchar2(255))
    TableSpace zl9CISAudit;
Create Table 病案接收记录(
	ID NUMBER(18),
	病人id NUMBER(18),
	主页ID NUMBER(18),
	运送人 varchar2(20),
	接收人 varchar2(20),
	接收时间 Date,
	记录时间 date)
	TABLESPACE zl9CISAudit;
Create Table 病案审阅书签(
    提交id		Number(18),
    审阅对象	Number(3),
    文件id		Number(18),
    审阅时间	Date)
    TableSpace zl9CISAudit;
Create Table 病案打印记录(
	病人id		Number(18),
	主页id		Number(5),
	打印次数	Number(5),
	打印序号	Number(5),
	打印内容	Varchar2(100), 
	打印人		Varchar2(20),	
	打印时间	Date)
	TABLESPACE zl9CISAudit;
Create Table 病案反馈记录(
    ID			Number(18),
    相关id		Number(18),
    提交id		Number(18),
    病人id		Number(18),
    主页id		Number(5),
    反馈对象	Number(3),
    文件id		Varchar2(32),
    医嘱id		Number(18),
    科室id		Number(18),
    记录性质	Number(3),
    记录状态	Number(3),
    反馈意见	Varchar2(255),
    反馈项目id	Number(18),
    反馈人		Varchar2(20),
    反馈时间	Date,
    处理期限	Date,
    处理说明	Varchar2(255),
    处理人		Varchar2(20),
    处理时间	Date,
    分值 NUMBER(8,2),
    补充说明 VARCHAR2(255),
    反馈次数 NUMBER(5),
    评分级别 Varchar2(1),
    分制 Number(1),
    反馈记录 Varchar2(200),
	子文档ID Varchar2(32))
    TableSpace zl9CISAudit
    PCTFREE 5;
Create Table 病案反馈历史(
    ID			Number(18),
    相关id		Number(18),
    提交id		Number(18),
    病人id		Number(18),
    主页id		Number(5),
    反馈对象		Number(3),
    文件id		Varchar2(32),
    医嘱id		Number(18),
    科室id		Number(18),
    记录性质	Number(3),
    记录状态	Number(3),
    反馈意见	Varchar2(255),
    反馈项目id	Number(18),
    反馈人		Varchar2(20),
    反馈时间	Date,
    处理期限	Date,
    处理说明	Varchar2(255),
    处理人		Varchar2(20),
    处理时间	Date,
    分值 NUMBER(8,2),
    补充说明 VARCHAR2(255),
    反馈次数 NUMBER(5),
    评分级别 Varchar2(1),
    分制 Number(1),
    反馈记录 Varchar2(200),
	子文档ID Varchar2(32))
    TableSpace zl9CISAudit
    PCTFREE 5;
Create Table 病案借阅记录(
	ID		Number(18),
	No		Varchar2(10),
	记录状态	Number(3),   
	申请人	Varchar2(20),	
	申请理由	Varchar2(255),
	申请时间	Date,
	申请期限	Date,
	借阅时间	Date,
	借阅期限	Date,
	批准人	Varchar2(20),
	批准时间	Date,
	拒借理由	Varchar2(255),
	拒借人	Varchar2(20),
	拒借时间	Date,
	登记时间	Date,
	收回人		Varchar2(20),
	归还时间	Date)
	TABLESPACE zl9CISAudit;
Create Table 病案借阅内容(
    借阅id		Number(18),
    病人id		Number(18),
    主页id		Number(5))
    TableSpace zl9CISAudit;
Create Table 病案借阅人员(
    借阅id		Number(18),
    人员id		Number(18))
    TableSpace zl9CISAudit;
Create Table 病案封存记录(
    ID			Number(18),
    病人id		Number(18),
    主页id		Number(5),
    记录状态	Number(3),
    封存人		Varchar2(20),
    封存时间	Date,
    封存理由	Varchar2(255),
    解封人		Varchar2(20),
    解封时间	Date)
    TableSpace zl9CISAudit;
--病案评分
Create Table 病案评分方案(
	ID number(18),
	名称 varchar2(50),
	总分 number(8,2) default 100,
	上值 number(8,2),
	下值 number(8,2),
	类型 varchar2(10),
	分制 varchar2(10),
	选用 number(1) default 0,
	启用时间 Date,
	停用时间 Date)
    	TABLESPACE zl9BaseItem;
Create Table 病案评分标准(
	ID number(18),
	上级ID number(18),
	方案ID number(18),
	名称 varchar2(50),
	描述 varchar2(4000),
	标准分值 number(8,2),
	缺陷等级 varchar2(2),
	评分单位 varchar2(8),
	上级序号 NUMBER(18),
	序号 NUMBER(18),
	判断依据 Varchar2(4000),
	否决等级 varchar2(2),
	数据源 Number(1) Default 0)
    	TABLESPACE zl9BaseItem;
Create Table 病案评分结果(
	ID number(18),
	病人ID number(18),
	主页ID number(5),
	方案ID number(18),
	总分 number(8,2),
	等级 varchar2(2),
	返回修改 number(1),
	病理类型 varchar2(20),
	备注	varchar2(50),
	评分人 varchar2(20),
	评分时间 Date,
	审核人 varchar2(20),
	审核时间 Date)
	TABLESPACE zl9CISAudit;
Create Table 病案评分明细(
	ID number(18),
	主表ID number(18),
	评分标准ID number(18),
	单项分数 number(8,2),
	缺陷等级 varchar2(2),
	可否修改 Number(1) Default 0,
	备注	varchar2(500))
	TABLESPACE zl9BaseItem;
CREATE TABLE 隐私保护项目(
    项目ID NUMBER(18))
    TABLESPACE zl9BaseItem;
CREATE TABLE 疾病报送单位(
    编码 VARCHAR2(2),
    名称 VARCHAR2(30),
    简码 VARCHAR2(10))
    TABLESPACE zl9BaseItem;
----------------------------------------------------------------------------
--[[19.护理业务]]
----------------------------------------------------------------------------
CREATE TABLE 体温单打印
(
文件ID number(18),
开始时间 date,
打印页号 number(5),
打印人 varchar2(20),
打印时间 date
) TABLESPACE ZL9EPRDAT;
CREATE TABLE 病人护理文件(
	ID NUMBER (18),
	科室ID NUMBER (18),
	病人ID NUMBER (18),
	主页ID NUMBER (18),
	婴儿 NUMBER (3),
	格式ID NUMBER (18),
	文件名称 VARCHAR2 (50),
	开始时间 DATE ,
	结束时间 DATE,
	续打ID NUMBER (18),			--同一个病人的文件中,只有文件ID相同的文件才允许续打，体温单除外（合并时不模拟计算）
	归档人 VARCHAR2 (20),
	归档时间 DATE ,
	创建人 VARCHAR2 (20),
	创建时间 DATE,
	待转出 Number(3))
	PCTFREE 20 initrans 10  
	TABLESPACE ZL9EPRDAT;
CREATE TABLE 病人护理打印(
	文件ID NUMBER (18),
	记录ID NUMBER (18),
	发生时间 DATE ,
	行数 NUMBER (3),
	开始页号 NUMBER (5),
	开始行号 NUMBER (5),
	结束页号 NUMBER (5),
	结束行号 NUMBER (5),
	行差 NUMBER (5) DEFAULT 0,	--记录与上次修改后相差的数据行，0表示行数未发生变化
	打印人 VARCHAR2 (20),
	打印时间 DATE,
	打印页号 NUMBER (5),
	打印行号 NUMBER (5),
	打印标识 NUMBER(1),
	打印结束页号 NUMBER(5),
	待转出 Number(3))
	PCTFREE 20 initrans 10   
	TABLESPACE ZL9EPRDAT;
CREATE TABLE  病人护理诊断(
	ID    NUMBER(18) ,
	病人ID NUMBER(18) , 
	主页ID NUMBER(18),
	文件ID NUMBER(18) ,
	诊断类型 NUMBER(2) ,
	诊断内容 VARCHAR2(200),
	是否疑诊 NUMBER(1),
	标记时间 DATE,
	待转出 NUMBER(3))
	PCTFREE 20 initrans 10   
	TABLESPACE ZL9EPRDAT;
CREATE TABLE 病人护理活动项目(
	文件ID NUMBER (18),
	页号 NUMBER (5),
	列号 NUMBER (5),
	列头名称 VARCHAR2 (100),		--此列的表头内容，缺省居中对齐
	序号 NUMBER(1),				--项目在当前列的排列序号，从1开始，最大为2
	项目序号 NUMBER (5),		--每列只能绑定一个项目或两个选择类型的项目
	部位 VARCHAR2 (50),
	操作员 VARCHAR2 (20),
	操作时间 DATE,
	待转出 Number(3))
	PCTFREE 20 INITRANS 10  
	TABLESPACE ZL9EPRDAT;
CREATE TABLE 病人护理数据(
	ID NUMBER (18),
	文件ID NUMBER (18),
	发生时间 DATE ,
	显示 NUMBER(2) DEFAULT 0,
	最后版本 NUMBER (5),
	保存人 VARCHAR2 (20),
	保存时间 DATE,
	签名人 VARCHAR2 (50),
	交班签名人 VARCHAR2(20),
	签名时间 VARCHAR2 (50),
	签名级别 NUMBER(3),
	汇总类别 NUMBER(3) DEFAULT 0,
	汇总文本 VARCHAR2(50),
	汇总标记 NUMBER(2),
	开始时点 VARCHAR2(5),
	结束时点 VARCHAR2(5),
	待转出 Number(3))
	PCTFREE 20 initrans 10  
	TABLESPACE ZL9EPRDAT;
CREATE TABLE 病人护理明细(
	ID NUMBER (18),
	记录ID NUMBER (18),
	记录类型 NUMBER (3),	--护理项目=1，上标说明=2，手术日标记=4，签名记录=5，下标说明=6，审签记录=15
	项目分组 VARCHAR2 (20),
	项目ID NUMBER (18),
	相关序号 NUMBER (5),
	项目序号 NUMBER (5),
	项目名称 VARCHAR2 (20),
	项目类型 NUMBER (3),
	记录内容 VARCHAR2 (4000),
	项目单位 VARCHAR2 (10),
	记录标记 NUMBER (3),	--通常填写为0，物理降温填写为1；脉搏短绌的心率填写为1
	体温部位 VARCHAR2 (10),
	记录组号 NUMBER (3),
	复试合格 NUMBER (1),
	数据来源 NUMBER (1) DEFAULT 0 ,	--0-手工录入;1-来源于记录单;2-来源于体温单;3-来源于PDA;9-历史数据，为了保证汇总数据不按新的上下级方式汇总，避免升级后体温单查看不正确
	来源ID NUMBER(18),				--明细ID
	共用 NUMBER (1) DEFAULT 0,		--1表示被其它记录使用,便于快速同步更新
	未记说明 VARCHAR2 (4000),
	开始版本 NUMBER (5),
	终止版本 NUMBER (5),
	记录人 VARCHAR2 (20),
	记录时间 DATE ,
	显示 NUMBER(1) DEFAULT 0,
	待转出 Number(3))
	PCTFREE 20 initrans 10  
	TABLESPACE ZL9EPRDAT;
CREATE TABLE 病人护理要素内容
(
  文件id NUMBER(18),
  页号   NUMBER(5),
  名称 VARCHAR2(60),
  内容 varchar2(1000),
  操作员  VARCHAR2(20),
  操作时间 DATE,
  待转出  NUMBER(3)
)tablespace ZL9EPRDAT;
Create Table 病区标记记录(
    病区ID NUMBER(18),
    病人ID NUMBER(18),
    主页ID NUMBER(18),
    主题序号 NUMBER(18),
    标记序号 NUMBER(5),
    日期 DATE,
    主题病区ID number(18),
    标记顺序 number(1))
    TABLESPACE ZL9EPRDAT 
;
Create Table 病人护理记录(
    ID NUMBER(18),
    病人来源 NUMBER(3),
    病人ID NUMBER(18),
    主页ID NUMBER(18),
    婴儿 NUMBER(5),
    科室ID NUMBER(18),
    护理级别 NUMBER(1),
    发生时间 DATE,
    最后版本 Number(5),
    归档人 Varchar2(20),
    归档时间 Date,
    保存人 VARCHAR2(20),
    保存时间 DATE,
    待转出 number(3))
    TABLESPACE zl9EprDat
;
Create Table 病人护理内容(
    ID NUMBER(18),
    记录ID NUMBER(18),
    记录类型 NUMBER(3),
    项目分组 VARCHAR2(20),
    项目ID NUMBER(18),
    项目序号 NUMBER(5),
    项目名称 VARCHAR2(20),
    项目类型 NUMBER(3),
    记录内容 VARCHAR2(4000),
    项目单位 VARCHAR2(10),
    记录标记 NUMBER(3),
    体温部位 Varchar2(10),
    记录组号 Number(3),
    复试合格 Number(1),
    未记说明 varchar2(4000),
    开始版本 Number(5),
    终止版本 Number(5),
    记录人 VARCHAR2(20),
    修改时间 DATE,
    待转出 number(3))
    TABLESPACE zl9EprDat 
;
----------------------------------------------------------------------------
--[[20.检验业务]]
----------------------------------------------------------------------------
Create TABLE 检验流水线标本(
  ID        NUMBER(18),
  标本ID    NUMBER(18),
  仪器是否审核  number(1),
  待转出 Number(3)
) TABLESPACE zl9CisRec;
Create TABLE 检验流水线指标(
  ID        NUMBER(18),
  标本ID    NUMBER(18),
  项目id    NUMBER(18),
  仪器是否审核  number(1),
  审核内容  varchar2(4000),
  待转出 Number(3)
) TABLESPACE zl9CisRec;
Create Table 检验标本记录(
	ID number(18),
	医嘱ID number(18),
	标本序号 varchar2(20) not Null,
	采样时间 Date,
	采样人 varchar2(20),
	标本类型 varchar2(200),
	核收人 varchar2(20),
	核收时间 Date,
	样本状态 number(1),
	检验人 varchar2(20),
	检验时间 Date,
	审核人 varchar2(20),
	审核时间 Date,
	合并报告号 number(18),
	打印次数 number(18),
	申请类型 number(1),
	仪器ID number(18),
	样本条码 varchar2(18),
	报告结果 number(2),
	备注 varchar2(4000),
	未通过审核原因 varchar2(40),
	申请时间 Date,
	标本形态 varchar2(50),
	是否质控品 number(1),
	执行科室id number(18),
	微生物标本 Number(1),
	No Varchar2(20),
	是否传送 NUMBER(1),
	标本类别 NUMBER(1),
	检验备注 VARCHAR2(400),
	病人来源 NUMBER(1),
	病人id NUMBER(18),  --如果为普通病人医嘱,对应的病人信息记录;如果为婴儿医嘱,表示其母亲对应的病人信息
	婴儿 NUMBER(3),	    --是第几个婴儿医嘱产生的申请,普通为0
	姓名 VARCHAR2(100),
	性别 VARCHAR2(4),
	年龄 VARCHAR2(20),
	年龄数字 NUMBER(4),
	年龄单位 VARCHAR2(10),
	申请人 Varchar2(20),
	申请科室id Number(20),
	合并ID number(18),
	床号 VARCHAR2(10),
	标识号 number(18),
	病人科室 varchar2(24),
	紧急 number(1),
	挂号单 varchar2(8),
	门诊号 number(18),
	住院号 number(18),
	出生日期 date,
	主页ID number(5),
	检验项目 varchar2(1000),
	操作类型 varchar2(20),
	接收人 varchar2(20),
	接收时间 date,
	杯号 varchar2(20),
	初审人 varchar2(20),
	初审时间 date,
	一级报告 varchar2(1000),
	二级报告 varchar2(1000),
	三级报告 varchar2(1000),
	审核未通过 varchar2(2000),
	结果为空 number(3),
	保存人 varchar2(30),
	保存时间 date,
	保存位置 varchar2(100),
	保存环境 varchar2(500),
	销毁人 varchar2(30),
	销毁时间 date,
	销毁方式 varchar2(100),
	待转出 Number(3))
    TABLESPACE zl9CisRec;
--检验项目分布
Create Table 检验项目分布(
	ID		Number(18),
	标本id		Number(18),
	项目id		Number(18),
	医嘱id		Number(18),
	细菌ID		number(18),
	范围		Number(1),
	待转出 Number(3))		--此字段暂时未用
	TABLESPACE zl9CisRec
  PCTFREE 5;
Create Table 检验试剂记录(
	医嘱id		Number(18),
	No		Varchar2(20),
	序号		Number(18),
	材料id		Number(18),
	数量		Number(16,5),
	固定 number(1),
	待转出 Number(3))
    TABLESPACE zl9CisRec
    PCTFREE 5;
Create Table 检验普通结果(
	ID number(18),
	检验标本ID number(18) not Null,
	检验项目ID number(18),
	检验结果 varchar2(500),
	结果标志 number(1),
	结果参考 varchar2(500),
	修改者 varchar2(20),
	修改时间 Date,
	记录类型 number(2),
	原始结果 varchar2(500),
	原始记录时间 Date,
	记录者 varchar2(20),
	是否检验 number(1),
	修改原因 number(1),
	细菌ID number(18),
	仪器ID number(18),
	培养描述 varchar2(50),
	诊疗项目ID number(18),
	排列序号 NUMBER(5),
	OD varchar2(20),
	CUTOFF varchar2(20),
	SCO varchar2(20),
	酶标板ID number(18),
	弃用结果 number(3),
	耐药机制 varchar2(100),
	药敏组ID number(18),
	稀释倍数 NUMBER(16,5),
	检验备注 varchar2(4000),
	待转出 Number(3))
    TABLESPACE zl9CisRec
    PCTFREE 5;
Create Table 检验药敏结果(
	细菌结果ID number(18),
	抗生素ID number(18),
	修改者 varchar2(20),
	修改时间 Date,
	结果 varchar2(20),
	结果类型 varchar2(10),
	记录类型 number(2),
	仪器ID number(18),
	药敏方法 Number(1),
	待转出 Number(3))
    TABLESPACE zl9CisRec
    PCTFREE 5;
Create Table 检验质控记录(
    标本ID Number(18),
    标本序号 Varchar2(20),
    检验人 Varchar2(20),
    仪器ID Number(18),
    检验时间 Date,
    时间 Varchar2(8),
    质控品ID Number(18),
    测试次数 Number(3),
    弃用记录 Number(3),
    新批号试剂 Number(1),
    新包装试剂 Number(1),
    新批号校准物 Number(1),
    新包装校准物 Number(1),
    新包装控制物 Number(1),
    仪器维护更新 Number(1),
	待转出 Number(3))
    Tablespace zl9CisRec;
Create Table 检验质控报告(
    结果ID Number(18),
    标记 Number(1),
    规则 Varchar2(100),
    提示 Varchar2(500),
    原因 Varchar2(500),
    措施 Varchar2(500),
    结论 Varchar2(500),
    报告人 Varchar2(20),
    报告时间 Date,
    归档人 Varchar2(20),
    归档时间 Date,
    项目ID number(18),
	待转出 Number(3))
    Tablespace zl9CisRec;
Create Table 检验弃用报告(
    结果ID Number(18),
    原因 Varchar2(500),
    报告人 Varchar2(20),
    报告时间 Date)
    Tablespace zl9CisRec;
Create Table 检验图像结果(
	ID			Number(18),
	标本id		Number(18),
	图像类型		varchar2(20),
	待转出		Number(3),
	图像点		CLOB,
	图像位置		varchar2(4000))		
	TABLESPACE zl9CisRec PCTFREE 5;
Create Table 检验酶标记录(
	ID		Number(18),
	板号		varchar2(20) Not Null,
	测试时间	Date,
	波长		varchar2(10),
	参考波长	varchar2(10),
	振板频率	varchar2(10),
	振板时间        varchar2(10),
	进板方式        varchar2(10),
	空白形式	varchar2(10),
	试剂批号	varchar2(20),
	试剂效期	Date,
	试剂厂商	varchar2(50),
	测试方法	varchar2(30),
	仪器ID		number(18),
	是否发送	Number(1),		--是否发送到技师工作站
	OD减空白	Number(1),		--是否减去空白孔质 1=要减
	存放位置	varchar2(50),		--用于保存存放位置
	单板单项	Number(1),		--是否是进行的单版单项测试 1=单版单项
	测试项目	varchar2(300),		--如是单板单项就只有一个项目 如单版多项格式：A_ID;B_ID:C_ID... 共8个项目
	阳性公式	varchar2(1000),		--如是单板单项就只有一个公式 如单版多项格式: A_公式;B_公式;C_公式...共8个公式
	弱阳性公式	varchar2(1000),		--如是单板单项就只有一个公式 如单版多项格式: A_公式;B_公式;C_公式...共8个公式
	CutOff公式	varchar2(1000),		--如是单板单项就只有一个公式 如单版多项格式: A_公式;B_公式;C_公式...共8个公式
	测试结果	varchar2(3000),		--编号1^结果;编号2^结果...编号12^结果|编号1^结果;编号2结果...编号12^结果 共8行每行12个结果为空填为"^"
	试剂记录	varchar2(1000)		--记录试剂:试剂批号;试剂效期;试剂厂商;测试方法|试剂批号;试剂效期;试剂厂商;测试方法|.....
	)
	TABLESPACE zl9CisRec;
Create Table 检验操作记录(
	ID Number(18),
	标本ID number(18),
	操作类型 number(2),  -- 0=审核 1=取消审核
	操作员 varchar2(20),
	操作时间 date,
	待转出 Number(3))
  TABLESPACE zl9CisRec;
Create Table 检验拒收记录(
    ID number(18),
    医嘱ID number(18),
    拒收人 varchar2(20),
    拒收时间 date,
    拒收理由 varchar2(200),
    重采人 varchar2(20),
    重采时间 date,
    待转出 Number(3),
    拒收标本接收人 VARCHAR2(20),
    拒收标本接收时间 date,
    通知护士 varchar(100),
    护士工号 varchar(20))
    TABLESPACE zl9CisRec 
;
Create Table 检验酶标试剂
(
  试剂批号 VARCHAR2(30),
  试剂效期 DATE,
  试剂厂商 VARCHAR2(100),
  测试方法 VARCHAR2(100),
  测试项目ID number(18)
) TABLESPACE zl9CisRec;
Create Table 检验分析记录(
	ID     NUMBER(18),
	标本ID number(18),
	用途 varchar2(10),
	待转出 Number(3))
  TABLESPACE zl9CisRec;
Create Table 检验签名记录(
    检验标本ID NUMBER(18),
    签名规则 NUMBER(2),
    签名信息 VARCHAR2(4000),
    时间戳 DATE,
    证书ID NUMBER(18),
    签名时间 DATE,
    签名人 VARCHAR2(20),
    待转出 Number(3),
    时间戳信息 varchar2(4000))
    TABLESPACE zl9CisRec
;
Create global temporary Table 检验酶标板打印(
  类型     VARCHAR2(20),
  列名	   VARCHAR2(20),
  Col1	   VARCHAR2(20),
  Col2	   VARCHAR2(20),
  Col3	   VARCHAR2(20),
  Col4	   VARCHAR2(20),
  Col5	   VARCHAR2(20),
  Col6	   VARCHAR2(20),
  Col7	   VARCHAR2(20),
  Col8	   VARCHAR2(20),
  Col9	   VARCHAR2(20),
  Col10	   VARCHAR2(20),
  Col11	   VARCHAR2(20),
  Col12	   VARCHAR2(20))
  on commit delete rows;
Create global temporary Table 质控即刻法打印(
  检验日期 VARCHAR2(10),
  次数     Varchar2(2),
  测定值   Varchar2(18),
  均值     Varchar2(18),
  SD       Varchar2(18),
  SI上限   Varchar2(18),
  SI下限   Varchar2(18),
  结果     VARCHAR2(10),
  检验者   VARCHAR2(30))
  on commit preserve rows;    
create global temporary table 质控即刻图打印
(
  项目     VARCHAR2(10),
  A01      varchar2(10),
  A02      varchar2(10),
  A03      varchar2(10),
  A04      varchar2(10),
  A05      varchar2(10),
  A06      varchar2(10),
  A07      varchar2(10),
  A08      varchar2(10),
  A09      varchar2(10),
  A10      varchar2(10),
  A11      varchar2(10),
  A12      varchar2(10),
  A13      varchar2(10),
  A14      varchar2(10),
  A15      varchar2(10),
  A16      varchar2(10),
  A17      varchar2(10),
  A18      varchar2(10),
  A19      varchar2(10),
  A20      varchar2(10)
)
on commit preserve rows;
----------------------------------------------------------------------------
--[[21.检查业务]]
----------------------------------------------------------------------------
Create Table 影像检查记录(
    医嘱ID number(18),
    发送号 number(18),
    影像类别 varchar2(10),
    执行科室ID number(18),
    检查号 varchar2(64),
    姓名 varchar2(100),
    英文名 varchar2(100),
    性别 varchar2(4),
    年龄 varchar2(20),
    出生日期 date,
    身高 number(16,5),
    体重 number(16,5),
    病理检查 number(1),
    检查UID varchar2(64),
    位置一 varchar2(3),
    位置二 varchar2(3),
    位置三 varchar2(3),
    检查设备 varchar2(30),
    是否打印 number(1),
    检查技师 Varchar2(20),
    检查技师二 Varchar2(20),
    影像质量 Varchar2(10),
    报告质量 Varchar2(10),
    危急状态 number(1),
    符合情况 Varchar2(10),
    附加主述 Varchar2(200),
    报告图象 varchar2(4000),
    接收日期 DATE,
    报到人 varchar2(20),
    完成人 varchar2(20),
    报告操作 Varchar2(20),
    绿色通道 Number(1),
    报告打印 Number(2),
    报告人 Varchar2(64),
    复核人 Varchar2(64),
    随访描述 Varchar2(200),
    诊断分类 VARCHAR2(100),
    发放胶片 number(1),
    报告发放 number(1),
    关联ID NUMBER(18),
    报告发放人 varchar2(10),
    胶片发放人 varchar2(10),
    图像位置 Number(1),
    图像数量 Number(5),
    是否技师确认 Number(1),
    是否电子胶片 Number(1),
	是否安排 Number(1),
    待转出 Number(3),
	待处理人 Varchar2(64),
	校对日期 date,
	校对状态 Number(1)) 
    TABLESPACE zl9CisRec;
create table 影像危急值记录(
    id number(18),
    医嘱id number(18),
    登记人 varchar2(30),
    登记时间 date,
    通知时间 date,
    通知方式 varchar2(20),
    接受科室 varchar2(30),
    接受人员 varchar2(30),
    处理结果 varchar2(512),
	待转出 Number(3))
    tablespace zl9CisRec;
Create Table 影像检查序列(
    序列UID varchar2(64),
    检查UID varchar2(64),
    序列号 number(10),
    序列描述 varchar2(64),
    采集时间 Date,
	待转出 Number(3))
    TABLESPACE zl9CisRec;
Create Table 影像检查图象(
    图像UID varchar2(64),
    序列UID varchar2(64),
    图像号 number(10),
    图像描述 varchar2(64),
    采集时间 date,
    图像时间 date,
    层厚 VARCHAR2(20),
    图像位置病人 VARCHAR2(64),
    图像方向病人 VARCHAR2(120),
    参考帧UID VARCHAR2(64),
    切片位置 VARCHAR2(20),
    行数 VARCHAR2(10),
    列数 VARCHAR2(10),
    像素距离 VARCHAR2(64),
    动态图 NUMBER(1),
    胶片打印 NUMBER(1),
    编码名称 varchar2(64),
    录制长度 number(18),
	待转出 Number(3),
	报告图 NUMBER(3),
	校对结果 Number(1),
	更新时间 date)
    TABLESPACE zl9CisRec
    PCTFREE 5;
Create Table 影像临时记录(
    影像类别 varchar2(10),
    检查号 varchar2(64),
    姓名 varchar2(100),
    英文名 varchar2(100),
    性别 varchar2(4),
    年龄 varchar2(20),
    出生日期 date,
    身高 number(5),
    体重 number(5),
    病理检查 number(1),
    发放胶片 number(1),
    检查UID varchar2(64),
    位置一 varchar2(3),
    位置二 varchar2(3),
    位置三 varchar2(3),
    检查设备 varchar2(64),
    报告图象 varchar2(2000),
    接收日期 DATE)
    TABLESPACE zl9CisRec
;
Create Table 影像临时序列(
    序列UID varchar2(64),
    检查UID varchar2(64),
    序列号 number(10),
    序列描述 varchar2(64),
    采集时间 Date)
    TABLESPACE zl9CisRec;
Create Table 影像临时图象(
    图像UID varchar2(64),
    序列UID varchar2(64),
    图像号 number(10),
    图像描述 varchar2(64),
    采集时间 date,
    图像时间 date,
    层厚 VARCHAR2(20),
    图像位置病人 VARCHAR2(64),
    图像方向病人 VARCHAR2(120),
    参考帧UID VARCHAR2(64),
    切片位置 VARCHAR2(20),
    行数 VARCHAR2(10),
    列数 VARCHAR2(10),
    像素距离 VARCHAR2(64),
    动态图 NUMBER(1),
    编码名称 varchar2(64),
    录制长度 number(18))    
    TABLESPACE zl9CisRec;
    
Create Table 影像报告驳回(
    ID Number(18),
    医嘱ID Number(18),
    病历ID Number(18),
    检查报告ID Raw(16),
    驳回理由 Varchar2(512),
    驳回时间 Date,
    驳回人 Varchar2(64),
    是否撤销 Number(1),
    待转出 Number(3),
    RISID Number(18),
    报告ID Number(18))
    Tablespace zl9CisRec 
;
Create Table 影像归档作业(
    编码 number(10),
    名称 varchar2(20),
    执行时间 Date,
    源设备 varchar2(1),
    目的设备 varchar2(1),
    指定设备 varchar2(3),
    是否迁移 number(1),
    是否删除 number(1),
    开始时间 Date,
    结束时间 Date,
    规则编码 number(10),
    自动备份 number(1),
    执行过程 number(1),
    检索条件 varchar2(250))
    TABLESPACE zl9CisRec;
Create Table 胶片打印记录(
    ID	Number(18),
    相关ID	Number(18),
    医嘱ID	Number(18),
    胶片大小	varchar2(20),
    打印人	varchar2(64),
    打印时间	Date)
    TABLESPACE zl9CisRec;
Create Table 影像收藏内容(
    ID   NUMBER(18),       
    收藏ID  NUMBER(18), 
    医嘱ID  NUMBER(18), 
    收藏时间 Date,
	待转出 Number(3)
)TABLESPACE zl9CISREC;
create table 影像申请单图像
(
    ID          NUMBER(18),      
    医嘱ID      NUMBER(18),    
    申请单图像  varchar2(64),           
    FTP路径     varchar2(100),
    设备号      varchar2(3),
    扫描人      varchar2(20),
    扫描时间    date,
	待转出 Number(3)
)
TABLESPACE zl9CISREC;
create table 影像预约设备
(
  id    NUMBER(18),
  设备名称  VARCHAR2(64),
  影像设备号 VARCHAR2(3),
  影像类别  VARCHAR2(10),
  设备说明  VARCHAR2(200),
  是否启用  NUMBER(1),
  是否默认  NUMBER(1),
  科室id  NUMBER(18)
)
tablespace ZL9CISREC;
create table 影像预约记录
(
  id      NUMBER(18),
  医嘱id    NUMBER(18),
  序号      VARCHAR2(64),
  预约设备id  NUMBER(18),
  预约设备名称  VARCHAR2(64),
  诊室名称    VARCHAR2(30),
  预约开始时间段 DATE,
  预约结束时间段 DATE,
  预约开始时间  DATE,
  预约结束时间  DATE,
  是否打印    NUMBER(1),
  是否检查    NUMBER(1),
  待转出      NUMBER(3),
  检查注意    VARCHAR2(2000),
  开单人      VARCHAR2(100),
  开单时间    DATE,
  是否收费    NUMBER(1),
  打印时间    DATE,
  打印人      VARCHAR2(100)
)
tablespace ZL9CISREC;
create table 影像预约项目
(
  id     NUMBER(18),
  影像类别   VARCHAR2(10),
  预约设备id NUMBER(18),
  诊疗项目id NUMBER(18),
  检查时长   NUMBER(4),
  注意事项   VARCHAR2(1000)
)
tablespace ZL9CISREC;
create table 影像预约方案
(
  id      NUMBER(18),
  预约设备id  NUMBER(18),
  方案名称    VARCHAR2(100),
  方案类型    NUMBER(1) not null,
  方案内容    VARCHAR2(10),
  间隔	      NUMBER(3),
  是否按日历休息 NUMBER(1),
  是否启用    NUMBER(1),
  开始时间    Date
)
tablespace ZL9CISREC;
create table 影像预约日历
(
  年月  NUMBER(6),
  休息日 VARCHAR2(100)
)
tablespace ZL9CISREC;
create table 影像预约时间计划
(
  id     NUMBER(18),
  预约方案id NUMBER(18),
  开始时间   DATE,
  结束时间   DATE,
  预约容量   NUMBER(3),
  计算方法   NUMBER(1)
)
tablespace ZL9CISREC;
create table 影像预约启用控制
(
  ID      	  NUMBER(18),
  检查科室ID  NUMBER(18),
  场合      	NUMBER(1),
  预约科室ID  NUMBER(18),
  是否必须预约  NUMBER(1)
)
tablespace ZL9CISREC;
---病理业务
----------------------------------------------------------------------------
Create Table 病理检查信息(
    病理医嘱ID Number(18), 
    医嘱ID Number(18),   
    病理号 VARCHAR2(20),           
    检查类型 Number(1),
    取材过程 Number(1) default 0,
    制片过程 Number(1) default 0,
    免疫过程 Number(1) default 0,
    特染过程 Number(1) default 0,
    分子过程 Number(1) default 0,
    巨检描述 Varchar2(2048),
    剩余位置 Varchar2(64),
    后续处理 Varchar2(64),
    综合质量 Varchar2(10),
    综合意见 varchar2(255),
    报到时间 Date,
	号码规则ID Number(5))
    TABLESPACE zl9CisRec; 
Create Table 病理质量信息(
    ID Number(18),   
    病理医嘱ID Number(18), 
    评价项目 VARCHAR2(20),   
    评价结果 VARCHAR2(10),   
    评价意见 Varchar2(255),
    改进方法 Varchar2(255),
    备注 Varchar2(1024),
    评价人 Varchar2(64),
    评价时间 date)
    TABLESPACE zl9CisRec; 
Create Table 病理标本信息(
    标本ID NUMBER(18),
    医嘱ID Number(18),
    送检ID Number(18),
    标本名称 VARCHAR2(64),
    材料类别 NUMBER(1) default 0,
    标本类型 NUMBER(1) default 0,
    采集部位 VARCHAR2(20),
    原有编号 VARCHAR2(20),
    数量 Number(2) default 0,
    存放位置 VARCHAR2(64),
    接收日期 Date,
    备注 VARCHAR2(1024))
    TABLESPACE zl9CisRec;    
Create Table 病理送检信息(
    ID NUMBER(18),   
    医嘱ID NUMBER(18),
    送检单位 VARCHAR2(64),
    送检科室 VARCHAR2(64),
    送检人 VARCHAR2(64),
    送检日期 DATE Not Null,
    联系方式 VARCHAR2(64),
    登记人 VARCHAR2(64),
    核收状态 NUMBER(1) default 1,
    拒收原因 VARCHAR2(1024),
    通知人 VARCHAR2(64),
    备注 VARCHAR2(1024))
    TABLESPACE zl9CisRec;
    
Create Table 病理申请信息(
    申请ID Number(18),  
    病理医嘱ID Number(18), 
    申请人 Varchar2(64),
    申请时间 Date,        
    申请类型 Number(1) default 0,
    申请细目 Number(1) default 0,
    申请状态 Number(1) default 0,
    申请描述 Varchar2(1024),
    是否打印 Number(1) default 0,
    补费状态 Number(1) default 0,
    完成时间 Date)
    TABLESPACE zl9CisRec;    
   
Create Table 病理取材信息(
    材块ID Number(18),
    序号 Number(18),
    病理医嘱ID Number(18), 
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
    是否脱钙 Number(1) default 0,
    是否蜡块 Number(1) default 0,
    主取医师 Varchar2(64),
    副取医师 Varchar2(64),
    记录医师 Varchar2(64),
    确认状态 Number(1) default 0,
    归档状态 number(1) default 0,
    取材时间 Date)
    TABLESPACE zl9CisRec;   
   
Create Table 病理脱钙信息(
    ID Number(18),   
    标本ID Number(18),
    开始时间 Date,
    所需时长 Number(5),
    当前缸次 Number(2),
    完成状态 Number(1) default 0,
    操作员 Varchar2(64))
    TABLESPACE zl9CisRec;     
Create Table 病理制片信息(
    ID Number(18),  
    病理医嘱ID Number(18), 
    材块ID Number(18),
    申请ID Number(18),
    制片类型 Number(1) default 0,
    制片方式 Number(1) default 0,
    制片时间 Date,
    制片数 Number(2),
    制片人 Varchar2(64),       
    当前状态 Number(1) default 0,
    归档状态 number(1) default 0,
    清单状态 Number(1) default 0)
    TABLESPACE zl9CisRec;     
Create Table 病理过程报告(
    ID Number(18),  
    病理医嘱ID Number(18), 
    标本名称 Varchar2(64),
    报告类型 Number(1),
    报告子项 Number(1),
    检查结果 Varchar2(2048),
    检查意见 Varchar2(2048),
    报告图像 Varchar2(2048),
    报告医师 Varchar2(64),        
    报告日期 Date,       
    当前状态 Number(1) default 0,
    备注 Varchar2(1024))
    TABLESPACE zl9CisRec;  
Create Table 病理抗体信息(
    抗体ID Number(18), 
    抗体名称 VARCHAR2(64),
    使用人份 Number(5),
    已用人份 Number(5),
    生产日期 Date,
    有效期 Number(4),
    过期日期 Date,
    克隆性 Number(1),
    作用对象 Varchar2(20),
    理化性质 Varchar2(10),
    应用情况 Varchar2(1024),
    登记人 Varchar2(64),
    登记时间 Date,
    使用状态 Number(1) default 1,
    备注 Varchar2(1024))
    TABLESPACE zl9CisRec;  
   
Create Table 病理特检信息(
    ID Number(18),    
    病理医嘱ID Number(18), 
    材块ID Number(18),
    申请ID Number(18),        
    抗体ID Number(18),
    项目顺序 Varchar2(20),
    特检类型 Number(1) default 0,
    特检细目 Number(1) default 0,
    制作类型 Number(1) default 0,
    当前状态 NUMBER(1) default 0,
    完成时间 Date,    
    特检医师 Varchar2(64),
    清单状态 Number(1) default 0,
    归档状态 number(1) default 0,
    项目结果 Varchar2(20) null)
    TABLESPACE zl9CisRec; 
Create Table 病理报告延迟(
    ID Number(18),    
    病理医嘱ID Number(18), 
    延迟原因 Varchar2(1024),        
    延迟天数 Number(2) default 0,
    临时诊断 Varchar2(1024),
    转达人 Varchar2(64),
    登记人 Varchar2(64),
    登记时间 Date,    
    当前状态 Number(1) default 0)
    TABLESPACE zl9CisRec; 
Create Table 病理会诊信息(
    ID Number(18),    
    病理医嘱ID Number(18), 
    申请医师 Varchar2(64),
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
    TABLESPACE zl9CisRec; 
Create Table 病理抗体反馈(
    ID Number(18),   
    抗体ID Number(18), 
    参考病理号 VARCHAR2(200),
    实验类型 Number(1) default 0,
    抗体评价 VARCHAR2(10),
    反馈意见 VARCHAR2(1024),
    反馈医生 VARCHAR2(64),
    反馈时间 Date)
    TABLESPACE zl9CisRec;   
Create Table 病理档案信息(
       ID Number(18),
       档案名称 Varchar2(64),
       档案编号 Varchar2(20),
       分类ID Number(18),
       检查范围 Varchar2(64),
       开始日期 date,
       结束日期 date,
       创建人 Varchar2(64),
       创建日期 date,
       档案说明 Varchar2(1024),
       所属房间 Varchar2(32),
       所属柜号 Varchar2(32),       
       所属抽屉 Varchar2(32),
       详细地址 Varchar2(128),
       档案状态 number(1) default 0,
       归档时间 Date
  )TABLESPACE zl9CisRec;
Create Table 病理归档信息(
    ID Number(18), 
    资料来源 Number(1) default 0,
    病理医嘱ID Number(18),
    材块ID   Number(18),
    制片ID   Number(18),
    特检ID   Number(18),
    档案ID   Number(18),
    存放状态 Number(1)  default 0,
    借阅状态 Number(1)  default 0
  )TABLESPACE zl9CisRec; 
  
Create Table 病理借阅信息(
    ID Number(18), 
    借阅人 Varchar2(64),
    借阅时间 Date,
    证件类型 Number(1) default 0,
    证件号码 Varchar2(20),
    联系电话 Varchar2(20),
    联系地址 Varchar2(128),
    押金 Number(16, 5),
    借阅类型 Number(1),
    借阅天数 Number(5),
    借阅原因 Varchar2(1024),
    登记人 Varchar2(64) ,
    归还状态 Number(1)  default 0,
    确认状态 Number(1)  default 0,
    备注 Varchar2(1024)
  )TABLESPACE zl9CisRec;       
Create Table 病理遗失信息(
       ID Number(18),
       借阅ID number(18),
       归档ID Number(18),
       遗失数量 Number(18),
       遗失原因 Varchar2(1024),
       遗失日期 date,
       登记人  Varchar2(64),
       备注 Varchar2(1024)
  )TABLESPACE zl9CisRec;
Create Table 病理归还信息(
       ID Number(18),
       借阅ID Number(18),
       归还人 Varchar2(64),
       归还日期 date,
       退还押金 Number(16,5),
       外诊医院 Varchar2(64),
       外诊医师 Varchar2(64),
       外诊意见 Varchar2(2048),
       登记人  Varchar2(64),
       备注 Varchar2(1024)
  )TABLESPACE zl9CisRec;
Create Table 病理借阅关联(
       借阅ID Number(18),
       归档ID Number(18),
       归还数量 Number(2),
       借阅数量 Number(2),
       归还状态 Number(1) default 0
  )TABLESPACE zl9CisRec;  
Create Table 病理玻片信息(
       Id Number(18),
       来源ID Number(18),
       来源类型 Number(1),
       材块ID Number(18),
       病理医嘱ID Number(18),
       条码号 Varchar2(30),
       归档状态 Number(1),
       玻片质量 Varchar2(10),
       评审人 Varchar2(30),
       评审日期 date,
       备注   varchar2(512)
)TABLESPACE zl9CisRec;   
  