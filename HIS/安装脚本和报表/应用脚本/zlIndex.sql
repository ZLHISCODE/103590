--分类目录
--1.公共基础,2.医保基础,3.病人病案基础,4.费用基础,5.药品卫材基础
--6.临床基础,7.临床路径基础,8.病历基础,9.护理基础,10.检验基础
--11.检查基础,12.医保业务,13.病人病案业务,14.费用业务,15.药品卫材业务
--16.临床医嘱,17.临床路径,18.病历业务,19.护理业务,20.检验业务,21.检查业务

----------------------------------------------------------------------------
--[[1.公共基础]]
----------------------------------------------------------------------------
Create Index 诊断前后注释_IX_分类 on 诊断前后注释(分类) Tablespace zl9Indexhis;

Create Index 手术前后注释_IX_分类 on 手术前后注释(分类) Tablespace zl9Indexhis;

create index ZLMSG_TODO_IX_CREATE_TIME on ZLMSG_TODO (CREATE_TIME) tablespace ZLMSGDATA;

Create Index 部门扩展信息_IX_项目 On 部门扩展信息(项目) Tablespace zl9Indexhis;
Create Index 人员扩展信息_IX_项目 On 人员扩展信息(项目) Tablespace zl9Indexhis;
Create Index 区域_IX_上级编码 On 区域(上级编码) Tablespace zl9Indexhis;
Create Index 人员表_IX_签名 On 人员表(签名) Tablespace zl9Indexhis;
Create Index 人员性质说明_IX_人员性质 On 人员性质说明(人员性质) Tablespace zl9Indexhis;
Create Index 人员证书记录_IX_人员ID On 人员证书记录(人员ID) Tablespace zl9Indexhis;
Create Index 部门性质说明_IX_工作性质 On 部门性质说明(工作性质) Tablespace zl9Indexhis;
Create Index 部门人员_IX_人员ID On 部门人员(人员ID) Tablespace zl9Indexhis;
Create Index 临床部门_IX_部门ID On 临床部门(部门ID) Tablespace zl9Indexhis;
Create Index 病区科室对应_IX_科室ID On 病区科室对应(科室ID) Tablespace zl9Indexhis;

Create Index 排队叫号队列_IX_队列名称 On 排队叫号队列(队列名称) Tablespace zl9Indexhis;
Create Index 排队叫号队列_IX_科室ID On 排队叫号队列(科室id) Tablespace zl9Indexhis;
Create Index 排队叫号队列_IX_病人ID On 排队叫号队列(病人ID) Tablespace zl9Indexhis;
create index 排队叫号队列_IX_业务ID on 排队叫号队列(业务ID,业务类型) tablespace zl9indexhis;
create index 排队语音呼叫_IX_队列ID on 排队语音呼叫(队列ID,站点) Tablespace zl9indexhis;
Create Index 上机人员表_IX_人员ID On 上机人员表(人员id)   Tablespace zl9indexhis;
----------------------------------------------------------------------------
--[[2.医保基础]]
----------------------------------------------------------------------------
Create Index 保险结算记录_IX_病人ID On 保险结算记录(病人ID) Tablespace zl9Indexhis;
Create Index 保险结算记录_IX_结算时间 On 保险结算记录(结算时间) Tablespace zl9Indexhis;
Create Index 保险结算记录_IX_卡类别ID On 保险结算记录(卡类别ID) Tablespace zl9Indexhis;
Create Index 保险支付项目_IX_大类ID On 保险支付项目(大类ID,险类) Tablespace zl9Indexhis;
Create Index 保险支付项目_IX_项目编码 On 保险支付项目(项目编码,险类) Tablespace zl9Indexhis;
Create Index 审批项目模板_IX_项目ID On 审批项目模板(项目ID) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[3.病人病案基础]]
----------------------------------------------------------------------------
create index 病人实名信息_IX_建档时间 on 病人实名信息 (建档时间)  tablespace ZL9INDEXHIS;
create index 病人实名信息_IX_陪诊人身份证号 on 病人实名信息 (陪诊人身份证号) tablespace ZL9INDEXHIS;
create index 病人实名信息_IX_手机号 on 病人实名信息 (手机号) tablespace ZL9INDEXHIS;
create index 病人实名信息_IX_姓名 on 病人实名信息 (姓名) tablespace ZL9INDEXHIS;
create index 病人实名证件_IX_证件号码 on 病人实名证件 (证件号码) tablespace ZL9INDEXHIS;
create index 实名认证接口日志_IX_实名ID on 实名认证接口日志 (实名ID) tablespace ZL9INDEXHIS;
create index 实名认证接口日志_IX_接口ID on 实名认证接口日志 (接口ID) tablespace ZL9INDEXHIS;
create index 实名认证接口日志_IX_调用时间 on 实名认证接口日志 (调用时间) tablespace ZL9INDEXHIS;

Create Index 医疗机构_IX_上级 on 医疗机构(上级) Tablespace zl9Indexhis;

Create Index 出院转入_IX_上级 on 出院转入(上级) Tablespace zl9Indexhis;

Create Index 疾病编码分类_IX_上级ID On 疾病编码分类(上级ID) Tablespace zl9Indexhis;
Create Index 疾病编码目录_IX_分类ID On 疾病编码目录(分类ID) Tablespace zl9Indexhis;
Create Index 疾病编码目录_IX_名称 On 疾病编码目录(名称) Tablespace zl9Indexhis;
Create Index 疾病编码目录_IX_简码 On 疾病编码目录(简码) Tablespace zl9Indexhis;
Create Index 疾病编码目录_IX_五笔码 On 疾病编码目录(五笔码) Tablespace zl9Indexhis;
Create Index 疾病编码科室_IX_科室ID On 疾病编码科室(科室ID) Tablespace zl9Indexhis;
Create Index 疾病编码科室_IX_人员ID On 疾病编码科室(人员ID) Tablespace zl9Indexhis;
Create Index 疾病诊断科室_IX_科室ID On 疾病诊断科室(科室ID) Tablespace zl9Indexhis;
Create Index 疾病诊断科室_IX_人员ID On 疾病诊断科室(人员ID) Tablespace zl9Indexhis;
Create Index 疾病诊断分类_IX_上级ID On 疾病诊断分类(上级ID) Tablespace zl9Indexcis;
Create Index 疾病诊断别名_IX_诊断ID On 疾病诊断别名(诊断id) Tablespace zl9Indexcis;
Create Index 疾病诊断别名_IX_名称 On 疾病诊断别名(名称) Tablespace zl9Indexcis;
Create Index 疾病诊断别名_IX_简码 On 疾病诊断别名(简码) Tablespace zl9Indexcis;
Create Index 疾病诊疗措施_IX_诊疗项目ID On 疾病诊疗措施(诊疗项目ID) Tablespace zl9Indexcis;
Create Index 疾病诊断规则_IX_项目ID On 疾病诊断规则(项目ID) Tablespace zl9Indexcis;
Create Index 疾病诊断对照_IX_诊断ID On 疾病诊断对照(诊断ID) Tablespace zl9Indexcis;
Create Index 疾病诊断对照_IX_手术ID On 疾病诊断对照(手术ID) Tablespace zl9Indexcis;

Create Index 咨询表格内容_IX_表号 On 咨询表格内容(表号) Tablespace zl9Indexhis;
Create Index 咨询广告序列_IX_图片序号 On 咨询广告序列(图片序号) Tablespace zl9Indexhis;
Create Index 咨询页面目录_IX_宣传标语 On 咨询页面目录(宣传标语) Tablespace zl9Indexhis;
Create Index 咨询页面目录_IX_页面背景 On 咨询页面目录(页面背景) Tablespace zl9Indexhis;
Create Index 咨询页面目录_IX_上级序号 On 咨询页面目录(上级序号) Tablespace zl9Indexhis;
Create Index 咨询页面排列_IX_页面 On 咨询页面排列(页面) Tablespace zl9Indexhis;
Create Index 咨询页面排列_IX_父序号 On 咨询页面排列(父序号) Tablespace zl9Indexhis;
Create Index 咨询页面排列_IX_页面图标 On 咨询页面排列(页面图标) Tablespace zl9Indexhis;
Create Index 咨询段落目录_IX_页面序号 On 咨询段落目录(页面序号) Tablespace zl9Indexhis;
Create Index 咨询段落目录_IX_标题图标 On 咨询段落目录(标题图标) Tablespace zl9Indexhis;
Create Index 咨询段落目录_IX_插表序号 On 咨询段落目录(插表序号) Tablespace zl9Indexhis;
Create Index 咨询段落目录_IX_插图序号 On 咨询段落目录(插图序号) Tablespace zl9Indexhis;
Create Index 咨询段落链接_IX_链接 On 咨询段落链接(页面序号,段落序号) Tablespace zl9Indexhis;
Create Index 咨询段落链接_IX_链接页面 On 咨询段落链接(链接页面) Tablespace zl9Indexhis;
Create Index 咨询专家清单_IX_人员id On 咨询专家清单(人员id) Tablespace zl9Indexhis;
Create Index 咨询专家清单_IX_科室id On 咨询专家清单(科室id) Tablespace zl9Indexhis;


----------------------------------------------------------------------------
--[[4.费用基础]]
----------------------------------------------------------------------------
CREATE INDEX 电子票据异常记录_IX_登记时间 ON 电子票据异常记录(登记时间) TABLESPACE zl9Indexhis; 

CREATE INDEX 电子票据异常记录_IX_病人ID ON 电子票据异常记录(病人ID) TABLESPACE zl9Indexhis;
CREATE INDEX 电子票据异常记录_IX_电子票据id ON 电子票据异常记录(电子票据id) TABLESPACE zl9Indexhis;
CREATE INDEX 电子票据异常记录_IX_单据号 ON 电子票据异常记录(单据号,操作场景) TABLESPACE zl9Indexhis; 
CREATE INDEX 电子票据开票点_IX_简码 ON 电子票据开票点(简码) TABLESPACE zl9Indexhis;

CREATE INDEX 电子票据开票点_IX_部门ID ON 电子票据开票点(部门ID) TABLESPACE zl9Indexhis;
CREATE INDEX 票据开票点对照_IX_人员ID ON 票据开票点对照(人员ID) TABLESPACE zl9Indexhis;

CREATE INDEX 票据开票点对照_IX_客户端 ON 票据开票点对照(客户端) TABLESPACE zl9Indexhis;
CREATE INDEX 合约单位_IX_名称 ON 合约单位(名称) TABLESPACE zl9Indexhis;

CREATE INDEX 电子票据使用记录_IX_登记时间 ON 电子票据使用记录(登记时间) TABLESPACE zl9Indexhis;

CREATE INDEX 电子票据使用记录_IX_生成时间 ON 电子票据使用记录(生成时间) TABLESPACE zl9Indexhis;
CREATE INDEX 电子票据使用记录_IX_结算ID ON 电子票据使用记录(结算ID) TABLESPACE zl9Indexhis;
CREATE INDEX 电子票据使用记录_IX_原票据ID ON 电子票据使用记录(原票据ID) TABLESPACE zl9Indexhis;
Create Index 电子票据使用记录_IX_待转出 On 电子票据使用记录(待转出) Tablespace zl9Indexcis;
Create Index 电子票据二维码_IX_待转出 On 电子票据二维码(待转出) Tablespace zl9Indexcis;

Create Index 病人费用异常记录_IX_NO On 病人费用异常记录(NO,记录性质) Pctfree 5 Tablespace zl9Indexhis;

Create Index 病人费用异常记录_IX_病人ID On 病人费用异常记录(病人ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人结算异常记录_IX_登记时间 On 病人结算异常记录(登记时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人结算异常记录_IX_病人id On 病人结算异常记录(病人id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人结算异常记录_IX_预交单号 On 病人结算异常记录(预交单号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人结算异常记录_IX_医疗卡单号 On 病人结算异常记录(医疗卡单号) Pctfree 5 Tablespace zl9Indexhis;

Create Index 病人押金记录_IX_病人ID On 病人押金记录(病人ID) Pctfree 5 Tablespace zl9Indexhis;

Create Index 病人押金记录_IX_主页ID On 病人押金记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人押金记录_IX_缴款组ID On 病人押金记录(缴款组ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人押金记录_IX_收款时间 On 病人押金记录(收款时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人押金记录_IX_交易时间 On 病人押金记录(交易时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人押金记录_IX_待转出 On 病人押金记录(待转出) Pctfree 5 Tablespace zl9Indexhis;
Create Index 费用结算对照_IX_费用ID on 费用结算对照(费用ID) Tablespace zl9Indexhis;
Create Index 费用结算对照_IX_待转出 On 费用结算对照(待转出) Tablespace Zl9indexhis;

Create Index 三方交易记录_IX_交易时间 On 三方交易记录(交易时间) Tablespace zl9Indexhis;

Create Index 费别明细_IX_收费细目id On 费别明细(费别, 收费细目id) Tablespace zl9Indexhis;
Create Index 收费分类目录_IX_上级ID On 收费分类目录(上级ID) Tablespace zl9Indexhis;
Create Index 收费项目目录_IX_分类ID On 收费项目目录(分类ID) Tablespace zl9Indexhis;
Create Index 收费项目别名_IX_名称 On 收费项目别名(名称) Tablespace zl9Indexhis;
Create Index 收费项目别名_IX_简码 On 收费项目别名(简码) Tablespace zl9Indexhis;
Create Index 收费执行科室_IX_开单科室ID On 收费执行科室(开单科室ID) Tablespace zl9Indexhis;
Create Index 收费执行科室_IX_执行科室ID On 收费执行科室(执行科室ID) Tablespace zl9Indexhis;
Create Index 收费价目_IX_收费细目id On 收费价目(收费细目id) Tablespace zl9Indexhis;
Create Index 收费价目_IX_价格等级 on 收费价目(价格等级) Tablespace zl9Indexhis;
Create Index 收费价目_IX_变动原因 On 收费价目(变动原因) Tablespace zl9Indexhis;
Create Index 成套项目分类_IX_简码 On 成套项目分类(简码) Tablespace zl9Indexhis;
Create Index 成套收费项目_IX_拼音 On 成套收费项目(拼音) Tablespace zl9Indexhis;
Create Index 成套收费项目_IX_五笔 On 成套收费项目(五笔) Tablespace zl9Indexhis;
Create Index 成套收费项目_IX_分类ID On 成套收费项目(分类ID) Tablespace zl9Indexhis;

Create Index 挂号安排_IX_执行计划ID On 挂号安排(执行计划ID) Tablespace zl9Indexhis;
Create Index 挂号安排计划_IX_安排时间 On 挂号安排计划(安排时间) Tablespace zl9Indexhis;
Create Index 挂号安排计划_IX_审核时间 On 挂号安排计划(审核时间) Tablespace zl9Indexhis;
Create Index 挂号安排计划_IX_生效时间 On 挂号安排计划(生效时间) Tablespace zl9Indexhis;
Create Index 挂号安排计划_IX_失效时间 On 挂号安排计划(失效时间) Tablespace zl9Indexhis;
Create Index 挂号安排计划_IX_实际生效 On 挂号安排计划(实际生效) Tablespace zl9Indexhis;
Create Index 挂号安排计划_IX_安排ID On 挂号安排计划(安排ID) Tablespace zl9Indexhis;
Create Index 挂号安排计划_IX_上次计划ID on 挂号安排计划(上次计划ID) Tablespace zl9Indexhis;
Create Index 挂号安排停用状态_IX_开始时间 On 挂号安排停用状态(开始停止时间) Tablespace zl9Indexhis;
Create Index 挂号安排停用状态_IX_结束时间 On 挂号安排停用状态(结束停止时间) Tablespace zl9Indexhis;
Create Index 常用退费原因_IX_简码 On 常用退费原因(简码) Tablespace zl9Indexhis;

Create Index 常用发卡原因_IX_简码 On 常用发卡原因(简码) Tablespace zl9Indexhis;
Create Index 消费卡信息_IX_发卡序号 On 消费卡信息(发卡序号) Tablespace zl9Indexhis;
Create Index 消费卡信息_IX_有效期 On 消费卡信息(有效期) Tablespace zl9Indexhis;
Create Index 消费卡信息_IX_发卡时间 On 消费卡信息(发卡时间) Tablespace zl9Indexhis;
Create Index 消费卡信息_IX_回收时间 On 消费卡信息(回收时间) Tablespace zl9Indexhis;
Create Index 消费卡信息_IX_当前状态 On 消费卡信息(当前状态) Tablespace zl9Indexhis;
Create Index 消费卡信息_IX_停用日期 On 消费卡信息(停用日期) Tablespace zl9Indexhis;
Create Index 消费卡信息_Ix_病人id On 消费卡信息(病人id) Tablespace Zl9indexhis;
Create Index 消费卡信息_Ix_领用id On 消费卡信息(领用id) Tablespace Zl9indexhis;

----------------------------------------------------------------------------
--[[5.药品卫材基础]]
----------------------------------------------------------------------------
Create Index 药品存储库房_IX_库房ID On 药品存储库房(库房ID) Tablespace zl9Indexhis;
Create Index 药品存储库房_IX_科室ID On 药品存储库房(科室ID) Tablespace zl9Indexhis;
Create Index 药品规格_IX_药名ID On 药品规格(药名ID) Tablespace zl9Indexhis;
Create Index 药品规格_IX_标识码 On 药品规格(标识码) Tablespace zl9Indexhis;
Create Index 药品规格_IX_带量供应商ID On 药品规格(带量供应商ID) Tablespace zl9Indexhis;
Create Index 材料特性_IX_诊疗ID On 材料特性(诊疗ID) Tablespace zl9Indexhis;
Create Index 材料领用信息_IX_主页ID On 材料领用信息(病人ID,主页ID) Tablespace zl9Indexhis;
Create Index 材料领用用途_IX_简码 On 材料领用用途(简码) Tablespace zl9Indexhis;
Create Index 供应商_IX_上级ID On 供应商(上级ID) Tablespace zl9Indexhis;
Create Index 供应商_IX_简码 On 供应商(简码) Tablespace zl9Indexhis;
Create Index 收费价目_IX_调价汇总号 On 收费价目(调价汇总号) Tablespace zl9Indexhis;
Create Index 卫材条码打印记录_IX_入库时间 on 卫材条码打印记录(入库时间) Tablespace zl9Indexhis;
Create Index 卫材条码打印记录_IX_材料id on 卫材条码打印记录(材料id) Tablespace zl9Indexhis;
Create Index 药品规格扩展信息_IX_项目 On 药品规格扩展信息(项目) Tablespace zl9Indexhis;
Create Index 药品库房货位_IX_库房id on 药品库房货位(库房id) Tablespace zl9Indexhis;
Create Index 药品库房货位_IX_上级id on 药品库房货位(上级id) Tablespace zl9Indexhis;
Create Index 药品货位对照_IX_药品ID on 药品货位对照(药品ID) Tablespace zl9Indexhis;
Create Index 药品货位对照_IX_货位ID on 药品货位对照(货位ID) Tablespace zl9Indexhis;
Create Index 入出类别对照_IX_入类别id on 入出类别对照(入类别id) Tablespace zl9Indexhis;
Create Index 入出类别对照_IX_出类别id on 入出类别对照(出类别id) Tablespace zl9Indexhis;
Create Index 药品批号对照_IX_供应商ID on 药品批号对照(供应商ID) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[6.临床基础]]
----------------------------------------------------------------------------
Create Index 医生交接班记录_IX_接班开始时间 On 医生交接班记录(接班开始时间)  Tablespace zl9Indexhis;

Create Index 医生交接班记录_IX_接班结束时间 On 医生交接班记录(接班结束时间)  Tablespace zl9Indexhis;
Create Index 医生交接班内容_IX_病人ID On 医生交接班内容(病人ID,主页ID) Tablespace zl9Indexhis;

Create Index 医生交接班签名_IX_记录ID On 医生交接班签名(记录ID)  Tablespace zl9Indexhis;

Create Index 急诊常用主诉_IX_上级 on 急诊常用主诉(上级) Tablespace zl9Indexhis;

Create Index 聊天会话表_IX_创建时间 On 聊天会话表(创建时间)  Tablespace zl9Indexhis;

Create Index 聊天会话表_IX_接收人 On 聊天会话表(接收人)  Tablespace zl9Indexhis;
Create Index 聊天会话表_IX_病人ID On 聊天会话表(病人ID,就诊ID) Tablespace zl9Indexhis;
Create Index 聊天信息表_IX_阅读时间 On 聊天信息表(阅读时间)  Tablespace zl9Indexhis;

Create Index 聊天信息表_IX_接收人 On 聊天信息表(接收人)  Tablespace zl9Indexhis;
Create Index 聊天信息表_IX_会话id On 聊天信息表(会话id)  Tablespace zl9Indexhis;
Create Index 证型方剂对照_IX_方剂ID on 证型方剂对照(方剂ID) Tablespace zl9indexhis;

Create Index 方剂构成_IX_草药ID on 方剂构成(草药ID) Tablespace zl9indexhis;

Create Index 加症治法_IX_加症ID on 加症治法(加症ID) Tablespace zl9indexhis;

Create Index 加症用药_IX_草药ID on 加症用药(草药ID) Tablespace zl9indexhis;

Create Index 医嘱执行组合_IX_待转出 On 医嘱执行组合(待转出) Tablespace zl9Indexcis;

Create Index 医嘱执行组合_Ix_要求时间 On 医嘱执行组合(要求时间) Pctfree 5 Tablespace Zl9indexcis;

Create Index 病人中医处方记录_IX_方剂ID on 病人中医处方记录(方剂ID) Tablespace zl9indexhis;
Create Index 病人中医处方记录_IX_煎法ID on 病人中医处方记录(HIS煎法ID) Tablespace zl9indexhis;
Create Index 病人中医处方记录_IX_用法ID on 病人中医处方记录(HIS用法ID) Tablespace zl9indexhis;
Create Index 病人中医处方记录_IX_药房ID on 病人中医处方记录(HIS药房ID) Tablespace zl9indexhis;

Create Index 病人中医诊断记录_IX_病人ID on 病人中医诊断记录(病人ID) Tablespace zl9indexhis;
Create Index 病人中医诊断记录_IX_处方ID on 病人中医诊断记录(处方ID) Tablespace zl9indexhis;
Create Index 病人中医诊断记录_IX_疾病ID on 病人中医诊断记录(疾病ID) Tablespace zl9indexhis;
Create Index 病人中医诊断记录_IX_证型ID on 病人中医诊断记录(证型ID) Tablespace zl9indexhis;

Create Index 病人中医诊断记录_IX_挂号单 on 病人中医诊断记录(挂号单) Tablespace zl9indexhis;
Create Index 病人中医诊断记录_IX_门诊号 on 病人中医诊断记录(门诊号) Tablespace zl9indexhis;
Create Index 病人中医诊断记录_IX_HIS诊断ID on 病人中医诊断记录(HIS诊断ID) Tablespace zl9indexhis;
Create Index 病人中医诊断记录_IX_HIS医嘱ID on 病人中医诊断记录(HIS医嘱ID) Tablespace zl9indexhis;
Create Index 病人中医诊断记录_IX_操作时间 on 病人中医诊断记录(操作时间) Tablespace zl9indexhis;
Create Index 病人中医处方明细_IX_诊疗项目ID on 病人中医处方明细(HIS品种ID) Tablespace zl9indexhis;
Create Index 病人中医处方明细_IX_草药ID on 病人中医处方明细(草药ID) Tablespace zl9indexhis;

Create Index 病人中医处方明细_IX_规格ID on 病人中医处方明细(HIS规格ID) Tablespace zl9indexhis;
Create Index 草药目录_IX_简码 on 草药目录(简码) Tablespace zl9indexhis;

Create Index 草药目录_IX_HIS品种ID on 草药目录(HIS品种ID) Tablespace zl9indexhis;
Create Index 中医证型_IX_疾病ID on 中医证型(疾病ID) Tablespace zl9indexhis;
Create Index 中医疾病_IX_科别 on 中医疾病(科别) Tablespace zl9indexhis;

Create Index 电子病历访问申请_IX_申请人 on 电子病历访问申请(申请人) Tablespace zl9indexhis;

Create Index 电子病历访问申请_IX_申请时间 on 电子病历访问申请(申请时间) Tablespace zl9indexhis;
Create Index 电子病历访问申请_IX_审批状态 on 电子病历访问申请(审批状态) Tablespace zl9indexhis;
Create Index 电子病历访问授权_IX_授权人 on 电子病历访问授权(授权人) Tablespace zl9indexhis;

Create Index 电子病历访问授权_IX_申请ID on 电子病历访问授权(申请ID) Tablespace zl9indexhis;
Create Index 电子病历访问授权_IX_授权时间 on 电子病历访问授权(授权时间) Tablespace zl9indexhis;
Create Index 电子病历访问授权_IX_开始时间 on 电子病历访问授权(访问开始时间) Tablespace zl9indexhis;
Create Index 电子病历访问授权_IX_结束时间 on 电子病历访问授权(访问结束时间) Tablespace zl9indexhis;
Create Index 电子病历访问日志_IX_访问人 on 电子病历访问日志(访问人) Tablespace zl9indexhis;

Create Index 电子病历访问日志_IX_访问时间 on 电子病历访问日志(访问时间) Tablespace zl9indexhis;
Create Index 电子病历授权访问病人_IX_授权id on 电子病历授权访问病人(授权id) Tablespace zl9indexhis;

Create Index 电子病历申请访问病人_IX_申请ID on 电子病历申请访问病人(申请ID) Tablespace zl9indexhis;

Create Index 电子病历授权访问人员_IX_授权id on 电子病历授权访问人员(授权id) Tablespace zl9indexhis;

Create Index RIS检查预约_IX_待转出 On RIS检查预约(待转出) Tablespace zl9Indexcis;
Create Index RIS检查预约_IX_预约开始时间 On RIS检查预约(预约开始时间) Tablespace zl9Indexcis;
Create Index RIS检查预约_IX_预约日期 On RIS检查预约(预约日期) Tablespace zl9Indexcis;
Create Index 医生常用诊断_IX_疾病ID On 医生常用诊断(疾病ID) Tablespace zl9Indexcis;
Create Index 医生常用诊断_IX_诊断ID On 医生常用诊断(诊断ID) Tablespace zl9Indexcis;
Create Index 医生常用医嘱_IX_疾病ID On 医生常用医嘱(疾病ID) Tablespace zl9Indexcis;
Create Index 医生常用医嘱_IX_诊断ID On 医生常用医嘱(诊断ID) Tablespace zl9Indexcis;
Create Index 医生常用医嘱_IX_诊疗项目ID On 医生常用医嘱(诊疗项目ID) Tablespace zl9Indexcis;
Create Index 医生常用医嘱_IX_药品ID On 医生常用医嘱(药品ID) Tablespace zl9Indexcis;
Create Index 医生常用医嘱_IX_人员ID On 医生常用医嘱(人员ID) Tablespace zl9Indexcis;
Create Index 输血检验对照_IX_检验项目id On 输血检验对照(检验项目id) Tablespace zl9Indexhis;
Create Index 抗菌药物抽样记录_IX_抽样时间 On 抗菌药物抽样记录(抽样时间) Tablespace zl9Indexcis;
Create Index 抗菌药物抽样明细_IX_病人ID On 抗菌药物抽样明细(病人ID,主页ID) Tablespace zl9Indexcis;
Create Index 抗菌药物抽样明细_IX_临床症状 On 抗菌药物抽样明细(临床症状) Tablespace zl9Indexcis;
Create Index 抗菌药物抽样明细_IX_感染诊断 On 抗菌药物抽样明细(感染诊断) Tablespace zl9Indexcis;
Create Index 抗菌药物抽样手术_IX_手术ID On 抗菌药物抽样手术(手术ID) Tablespace zl9Indexcis;
Create Index 诊疗分类目录_IX_上级ID On 诊疗分类目录(上级ID) Tablespace zl9Indexhis;
Create Index 诊疗项目目录_IX_分类ID On 诊疗项目目录(分类ID) Tablespace zl9Indexhis;
Create Index 诊疗项目别名_IX_名称 On 诊疗项目别名(名称) Tablespace zl9Indexhis;
Create Index 诊疗项目别名_IX_简码 On 诊疗项目别名(简码) Tablespace zl9Indexhis;
Create Index 诊疗执行科室_IX_开单科室ID On 诊疗执行科室(开单科室ID) Tablespace zl9Indexcis;
Create Index 诊疗执行科室_IX_执行科室ID On 诊疗执行科室(执行科室ID) Tablespace zl9Indexcis;
Create Index 诊疗项目组合_IX_诊疗项目ID On 诊疗项目组合(诊疗项目ID) Tablespace zl9Indexcis;
Create Index 诊疗项目组合_IX_配方ID On 诊疗项目组合(配方ID) Tablespace zl9Indexcis;
Create Index 诊疗项目部位_IX_部位 on 诊疗项目部位(部位,类型) Tablespace zl9indexhis;
Create Index 诊疗收费关系_IX_收费项目ID On 诊疗收费关系(收费项目id) Tablespace zl9Indexcis;
Create Index 人员抗菌药物权限_Ix_人员id On 人员抗菌药物权限(人员id) Tablespace Zl9Indexhis;
Create Index 人员手术权限申请_IX_诊疗项目ID On 人员手术权限申请(诊疗项目ID) Tablespace zl9Indexhis;
Create Index 人员手术权限申请_IX_授权人员ID On 人员手术权限申请(授权人员ID) Tablespace zl9Indexhis;
Create Index 人员手术权限_IX_诊疗项目ID On 人员手术权限(诊疗项目ID) Tablespace zl9Indexhis;
Create Index 常用就诊摘要_IX_人员ID On 常用就诊摘要(人员ID) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[7.临床路径基础]]
----------------------------------------------------------------------------
Create Index 临床路径项目_IX_版本号 On 临床路径项目(路径ID,版本号) Tablespace zl9Indexcis;
Create Index 临床路径项目_IX_阶段ID On 临床路径项目(阶段ID) Tablespace zl9Indexcis;
Create Index 临床路径项目_IX_图标ID On 临床路径项目(图标ID) Tablespace zl9Indexcis;
Create Index 路径医嘱内容_IX_相关ID On 路径医嘱内容(相关ID) Tablespace zl9Indexcis;
Create Index 路径医嘱内容_IX_诊疗项目ID On 路径医嘱内容(诊疗项目ID) Tablespace zl9Indexcis;
Create Index 路径医嘱内容_IX_收费细目ID On 路径医嘱内容(收费细目ID) Tablespace zl9Indexcis;
Create Index 路径医嘱内容_IX_执行科室ID On 路径医嘱内容(执行科室ID) Tablespace zl9Indexcis;
Create Index 路径医嘱内容_IX_配方ID On 路径医嘱内容(配方ID) Tablespace zl9Indexcis;
Create Index 临床路径分支_IX_前一阶段ID On 临床路径分支(前一阶段ID) Tablespace zl9Indexhis;
Create Index 临床路径阶段_IX_分支ID On 临床路径阶段(分支ID) Tablespace zl9Indexhis;
Create Index 临床路径阶段_IX_父ID On 临床路径阶段(父ID) Tablespace zl9Indexcis;
Create Index 临床路径分类_IX_分支ID On 临床路径分类(分支ID) Tablespace zl9Indexhis;
Create Index 临床路径项目_IX_分支ID On 临床路径项目(分支ID) Tablespace zl9Indexhis;
Create Index 临床路径评估_IX_分支ID On 临床路径评估(分支ID) Tablespace zl9Indexhis;
Create Index 临床路径评估_IX_阶段ID On 临床路径评估(阶段ID) Tablespace zl9Indexcis;
Create Index 路径评估条件_IX_评估ID On 路径评估条件(评估ID) Tablespace zl9Indexcis;
Create Index 路径评估条件_IX_项目ID On 路径评估条件(项目ID) Tablespace zl9Indexcis;
Create Index 临床路径医嘱_IX_医嘱内容ID On 临床路径医嘱(医嘱内容ID) Tablespace zl9Indexcis;

----------------------------------------------------------------------------
--[[8.病历基础]]
----------------------------------------------------------------------------
Create Index 诊治所见项目_IX_分类ID On 诊治所见项目(分类ID) Tablespace zl9Indexcis;
Create Index 病历提纲词句_IX_词句分类ID On 病历提纲词句(词句分类ID) Tablespace zl9Indexcis;
Create Index 病历替代关系_IX_替代ID On 病历替代关系(替代ID) Tablespace zl9Indexcis;
Create Index 病历应用科室_IX_科室ID On 病历应用科室(科室ID) Tablespace zl9Indexcis;
Create Index 疾病报告前提_IX_疾病ID On 疾病报告前提(疾病ID) Tablespace zl9Indexcis;
Create Index 疾病报告前提_IX_诊断ID On 疾病报告前提(诊断ID) Tablespace zl9Indexcis;
Create Index 病历单据应用_IX_病历文件ID On 病历单据应用(病历文件ID) Tablespace zl9Indexcis;
Create Index 病历附项模板_IX_病历文件Id On 病历附项模板(病历文件Id,单据附项) Tablespace zl9Indexhis;
Create Index 病历文件结构_IX_父ID On 病历文件结构(父ID) Tablespace zl9Indexcis;
Create Index 病历文件结构_IX_预制提纲ID On 病历文件结构(预制提纲ID) Tablespace zl9Indexcis;
Create Index 病历文件结构_IX_诊治要素ID On 病历文件结构(诊治要素ID) Tablespace zl9Indexcis;
Create Index 病历词句示范_IX_科室id On 病历词句示范(科室id) Tablespace zl9Indexcis;
Create Index 病历词句示范_IX_人员id On 病历词句示范(人员id) Tablespace zl9Indexcis;
Create Index 病历词句示范_IX_编号 On 病历词句示范(编号) Tablespace zl9Indexcis;
Create Index 病历词句示范_IX_名称 On 病历词句示范(名称) Tablespace zl9Indexcis;
Create Index 病历词句组成_IX_内容文本 On 病历词句组成(内容文本) Tablespace zl9Indexcis;
Create Index 病历范文目录_IX_科室id On 病历范文目录(科室id) Tablespace zl9Indexcis;
Create Index 病历范文目录_IX_人员id On 病历范文目录(人员id) Tablespace zl9Indexcis;
Create Index 病历范文内容_IX_父ID On 病历范文内容(父ID) Tablespace zl9Indexcis;
Create Index 病历范文内容_IX_预制提纲ID On 病历范文内容(预制提纲ID) Tablespace zl9Indexcis;
Create Index 病历范文内容_IX_诊治要素ID On 病历范文内容(诊治要素ID) Tablespace zl9Indexcis;

Create Index 病案审查分类_IX_上级id On 病案审查分类(上级id) Tablespace zl9Indexcis;
Create Index 病案审查分类_IX_方案id On 病案审查分类(方案id) Tablespace zl9Indexcis;
Create Index 病案审查目录_IX_分类id On 病案审查目录(分类id) Tablespace zl9Indexcis;
----------------------------------------------------------------------------
--[[9.护理基础]]
----------------------------------------------------------------------------
Create Index 体温重叠标记_IX_上级序号 On 体温重叠标记(上级序号) Tablespace zl9Indexcis;
Create Index 护理适用科室_IX_科室ID On 护理适用科室(科室ID) Tablespace zl9Indexcis;
----------------------------------------------------------------------------
--[[10.检验基础]]
----------------------------------------------------------------------------
Create Index 检验细菌_IX_简码 On 检验细菌(简码) Tablespace zl9Indexcis;
Create Index 检验试剂关系_IX_材料id On 检验试剂关系(材料id) Tablespace zl9Indexcis;
Create Index 检验备注文字_IX_分类 On 检验备注文字(分类) Tablespace zl9Indexcis;
Create Index 检验评语文字_IX_分类 On 检验评语文字(分类) Tablespace zl9Indexcis;
Create Index 检验报告项目_IX_细菌ID On 检验报告项目(细菌id) Tablespace zl9Indexcis;
Create Index 检验报告项目_IX_报告项目ID On 检验报告项目(报告项目ID) Tablespace zl9Indexcis;

Create Index 检验模板目录_IX_诊疗项目ID On 检验模板目录(诊疗项目ID) Tablespace zl9Indexcis;
Create Index 检验模板内容_IX_模板ID On 检验模板内容(模板ID) Tablespace zl9Indexcis;
Create Index 检验模板内容_IX_项目ID On 检验模板内容(项目ID) Tablespace zl9Indexcis;
Create Index 检验模板内容_IX_细菌ID On 检验模板内容(细菌ID) Tablespace zl9Indexcis;
Create Index 检验模板药敏_IX_抗生素ID On 检验模板药敏(抗生素ID) Tablespace zl9Indexcis;
Create Index 检验合并规则_IX_主项目ID On 检验合并规则(主项目ID) Tablespace zl9Indexcis;
Create Index 检验合并规则_IX_合并项目ID On 检验合并规则(合并项目ID) Tablespace zl9Indexcis;

Create Index 检验仪器_IX_使用小组ID On 检验仪器(使用小组ID) Tablespace zl9Indexcis;
Create Index 检验仪器抗生素_IX_抗生素ID On 检验仪器项目(抗生素id) Tablespace zl9Indexcis;
Create Index 检验仪器抗生素_IX_项目ID On 检验仪器项目(项目id) Tablespace zl9Indexcis;
Create Index 检验仪器状态_IX_项目ID On 检验仪器状态(项目ID) Tablespace zl9Indexcis;
Create Index 检验仪器规则_IX_上级ID On 检验仪器规则(上级ID) Tablespace zl9Indexcis;
Create Index 检验仪器规则_IX_仪器ID On 检验仪器规则(仪器ID) Tablespace zl9Indexcis;
Create Index 检验仪器规则_IX_规则ID On 检验仪器规则(规则ID) Tablespace zl9Indexcis;

----------------------------------------------------------------------------
--[[11.检查基础]]
----------------------------------------------------------------------------
Create Index RIS接口日志记录_IX_时间 On RIS接口日志记录(时间) Tablespace zl9Indexcis;
Create Index 病理号码记录_IX_号码规则ID On 病理号码记录(号码规则ID) Tablespace zl9Indexcis;
Create Index 病理号码记录_IX_年 On 病理号码记录(年) Pctfree 5 Tablespace zl9Indexcis;
Create Index 影像查询方案_IX_所属科室 On 影像查询方案(所属科室) Tablespace zl9Indexhis;
Create Index 影像查询配置_IX_方案ID On 影像查询配置(方案ID) Tablespace zl9Indexhis;
Create Index 快捷功能信息_IX_模块号 On 快捷功能信息(模块号,项目) Tablespace zl9Indexhis;
create index 医技执行房间_IX_分组ID on 医技执行房间(分组ID) Tablespace zl9Indexhis;
create index 影像分组关联_IX_分组ID on 影像分组关联(分组ID) Tablespace zl9Indexhis;
create index 影像执行分组_IX_科室ID on 影像执行分组(科室ID) Tablespace zl9Indexhis;
Create Index 影像申请常用词句_IX_科室ID On 影像申请常用词句(科室ID) Tablespace zl9Indexhis;
Create Index 影像申请常用词句_IX_创建人员ID On 影像申请常用词句(创建人员ID) Tablespace zl9Indexhis;
----------------------------------------------------------------------------
--[[12.医保业务]]
----------------------------------------------------------------------------
Create Index 医保病人档案_IX_就诊时间 On 医保病人档案(就诊时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 医保病人关联表_IX_病人ID On 医保病人关联表(病人ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人审批项目_IX_项目ID On 病人审批项目(项目ID) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[13.病人病案业务]]
----------------------------------------------------------------------------
Create Index 病人新生儿记录_IX_婴儿病人ID On 病人新生儿记录(婴儿病人ID,婴儿主页ID) Pctfree 5 Tablespace zl9Indexhis;

Create Index 常用不良行为原因_IX_简码 On 常用不良行为原因(简码) Tablespace zl9Indexhis;
Create Index 不良行为控制_IX_行为类别 On 不良行为控制(行为类别) Tablespace zl9Indexhis;
Create Index 病人不良记录_IX_加入时间 On 病人不良记录(加入时间) Tablespace zl9Indexhis;
Create Index 病人不良记录_IX_发生时间 On 病人不良记录(发生时间) Tablespace zl9Indexhis;
Create Index 病人不良记录_IX_撤消时间 On 病人不良记录(撤消时间) Tablespace zl9Indexhis;
Create Index 病人不良记录_IX_病人ID On 病人不良记录(病人ID) Tablespace zl9Indexhis;
Create Index 病人不良记录_IX_行为类别 On 病人不良记录(行为类别) Tablespace zl9Indexhis;
Create Index 病人自动计算_IX_病人ID On 病人自动计算(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人自动计算_IX_开始时间 On 病人自动计算(开始时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人自动计算_IX_终止时间 On 病人自动计算(终止时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 床位增减记录_IX_病区ID On 床位增减记录(病区ID) Tablespace zl9Indexhis;
Create Index 床位状况记录_IX_科室ID On 床位状况记录(科室ID) Tablespace zl9Indexhis;
Create Index 床位状况记录_IX_病人ID On 床位状况记录(病人ID) Tablespace zl9Indexhis;

Create Index 病人信息_IX_姓名 On 病人信息(姓名) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人信息_IX_登记时间 On 病人信息(登记时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人信息_IX_身份证号 On 病人信息(身份证号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人信息_IX_IC卡号 On 病人信息(IC卡号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人信息_IX_医保号 On 病人信息(医保号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人信息_IX_合同单位id On 病人信息(合同单位id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人信息_IX_在院 On 病人信息(在院) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人信息_IX_手机号 on 病人信息(手机号) Tablespace zl9Indexhis;
Create Index 病人信息_IX_当前科室ID On 病人信息(当前科室ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人信息_IX_联系人身份证号 On 病人信息(联系人身份证号 ) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人身份关联_IX_关联ID On 病人身份关联(关联ID) Tablespace zl9Indexhis;
Create Index 在院病人_IX_病人ID On 在院病人(病人ID,主页ID) Tablespace zl9Indexhis;
Create Index 病人合并记录_IX_病人ID On 病人合并记录(病人id) Tablespace zl9Indexhis;
Create Index 病人合并记录_IX_原病人id On 病人合并记录(原病人id) Tablespace Zl9indexhis;
Create Index 病人担保记录_IX_主页ID On 病人担保记录(病人ID,主页ID) Tablespace zl9Indexhis;
Create Index 病人变动记录_IX_病人ID On 病人变动记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人变动记录_IX_医疗小组ID On 病人变动记录(医疗小组ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人变动记录_IX_开始时间 On 病人变动记录(开始时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人变动记录_IX_终止时间 On 病人变动记录(终止时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_入院日期 On 病案主页(入院日期) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_出院日期 On 病案主页(出院日期) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_医疗小组ID On 病案主页(医疗小组ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_住院号 On 病案主页(住院号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_病案号 On 病案主页(病案号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病案主页_IX_留观号 On 病案主页(留观号) Tablespace zl9Indexhis;
Create Index 病案主页_IX_待转出 On 病案主页(待转出) Tablespace zl9Indexhis;
Create Index 病案主页_IX_挂号ID On 病案主页(挂号ID)  Tablespace zl9Indexcis;

Create Index 病人家属_IX_家属ID On 病人家属(家属ID) Tablespace zl9Indexhis;

Create Index 住院病案记录_IX_病人ID On 住院病案记录(病案号) PCTFREE 5 Tablespace zl9Indexhis;
Create Index 住院病案记录_IX_档案号 On 住院病案记录(档案号) PCTFREE 5 Tablespace zl9Indexhis;

Create Index 病人过敏记录_IX_病人ID On 病人过敏记录(病人ID,主页ID) Tablespace zl9Indexcis;
Create Index 病人过敏记录_IX_待转出 On 病人过敏记录(待转出) Tablespace zl9Indexcis;
Create Index 病人诊断记录_IX_病人ID On 病人诊断记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人诊断记录_IX_医嘱id On 病人诊断记录(医嘱id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人诊断记录_IX_病历ID On 病人诊断记录(病历ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人诊断记录_IX_病例ID On 病人诊断记录(病例ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人诊断记录_IX_待转出 On 病人诊断记录(待转出) Tablespace zl9Indexcis;
Create Index 病人诊断记录_IX_疾病id On 病人诊断记录(疾病id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人诊断记录_IX_诊断id On 病人诊断记录(诊断id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人诊断记录_IX_证候id On 病人诊断记录(证候id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人诊断医嘱_IX_医嘱ID On 病人诊断医嘱(医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人诊断医嘱_IX_待转出 On 病人诊断医嘱(待转出) Tablespace zl9Indexcis;
Create Index 病人手麻记录_IX_主页ID On 病人手麻记录(病人ID,主页ID ) Tablespace zl9Indexcis;
Create Index 病人手麻记录_IX_待转出 On 病人手麻记录(待转出) Tablespace zl9Indexcis;
Create Index 病人抗生素记录_IX_药名id On 病人抗生素记录(药名id) Tablespace zl9Indexcis;

Create Index 病案化疗记录_IX_开始日期 On 病案化疗记录(开始日期) PCTFREE 5 Tablespace zl9Indexcis;
Create Index 病案化疗记录_IX_结束日期 On 病案化疗记录(结束日期) PCTFREE 5 Tablespace zl9Indexcis;
Create Index 病案放疗记录_IX_开始日期 On 病案放疗记录(开始日期) PCTFREE 5 Tablespace zl9Indexcis;
Create Index 病案放疗记录_IX_结束日期 On 病案放疗记录(结束日期) PCTFREE 5 Tablespace zl9Indexcis;
Create Index 病案精神治疗_IX_登记时间 On 病案精神治疗(药物名称) PCTFREE 5 Tablespace zl9Indexcis;

----------------------------------------------------------------------------
--[[14.费用业务]]
----------------------------------------------------------------------------
Create Index 卡类别ID_IX_卡类别ID on 三方退款信息(卡类别ID) Tablespace zl9Indexhis;
Create Index 三方退款信息_IX_卡类别ID On 三方退款信息(卡类别ID) Tablespace Zl9indexhis;

Create Index 预交单据余额_IX_病人ID on 预交单据余额(病人ID) Tablespace zl9Indexhis;

Create Index 消费卡信息_Ix_病人id On 消费卡信息(病人id) Tablespace Zl9indexhis;
Create Index 消费卡信息_Ix_领用id On 消费卡信息(领用id) Tablespace Zl9indexhis;

Create Index 消费卡入库记录_Ix_登记人 On 消费卡入库记录(登记人) Tablespace Zl9indexhis;
Create Index 消费卡入库记录_Ix_登记时间 On 消费卡入库记录(登记时间) Tablespace Zl9indexhis;
Create Index 消费卡入库记录_Ix_是否存在卡 On 消费卡入库记录(是否存在卡) Tablespace Zl9indexhis;
Create Index 消费卡入库记录_IX_批次 On 消费卡入库记录(批次) Tablespace zl9Indexhis;

Create Index 消费卡领用记录_Ix_领用人 On 消费卡领用记录(领用人) Tablespace Zl9indexhis;
Create Index 消费卡领用记录_Ix_批次 On 消费卡领用记录(批次) Tablespace Zl9indexhis;
Create Index 消费卡领用记录_Ix_登记时间 On 消费卡领用记录(登记时间) Tablespace Zl9indexhis;
Create Index 消费卡领用记录_IX_入库ID On 消费卡领用记录(入库ID) Tablespace zl9Indexhis;

Create Index 消费卡报损记录_Ix_入库id On 消费卡报损记录(入库id) Tablespace Zl9indexhis;
Create Index 消费卡报损记录_Ix_报损人 On 消费卡报损记录(报损人) Tablespace Zl9indexhis;
Create Index 消费卡报损记录_Ix_报损时间 On 消费卡报损记录(报损时间) Tablespace Zl9indexhis;

Create Index 消费卡使用记录_Ix_领用id On 消费卡使用记录(领用id, 性质) Tablespace Zl9indexhis;
Create Index 消费卡使用记录_Ix_使用时间 On 消费卡使用记录(使用时间) Tablespace Zl9indexhis;

Create Index 帐户缴款余额_Ix_交易序号 On 帐户缴款余额(交易序号) Tablespace Zl9indexhis;

Create Index 病人缴款记录_IX_病人ID On 病人缴款记录(病人ID) Tablespace zl9indexhis;
Create Index 病人缴款记录_IX_登记时间 On 病人缴款记录(登记时间) Tablespace zl9indexhis;
Create Index 病人缴款对照_IX_结帐Id On 病人缴款对照(结帐Id) Tablespace zl9indexhis;

Create Index 费用变动记录_Ix_目标变动id On 费用变动记录(目标变动id) Tablespace Zl9indexhis;
Create Index 费用变动记录_IX_待转出 On 费用变动记录(待转出) Tablespace Zl9indexhis;

Create Index 费用变动记录_Ix_收费细目id On 费用变动记录(收费细目id) Tablespace Zl9indexhis;
Create Index 费用变动记录_Ix_费用id On 费用变动记录(费用id) Tablespace Zl9indexhis;
Create Index 费用变动记录_Ix_病人id On 费用变动记录(病人id, 主页id) Tablespace Zl9indexhis;
Create Index 病人服务信息记录_IX_登记时间 on 病人服务信息记录(登记时间) Tablespace zl9Indexhis;

Create Index 病人服务信息记录_IX_处理时间 on 病人服务信息记录(处理时间) Tablespace zl9Indexhis;
Create Index 病人服务信息记录_IX_病人ID on 病人服务信息记录(病人ID) Tablespace zl9Indexhis;
Create Index 病人服务信息记录_IX_挂号ID on 病人服务信息记录(挂号ID) Tablespace zl9Indexhis;
Create Index 病人服务信息记录_IX_号码ID on 病人服务信息记录(号码) Tablespace zl9Indexhis;
Create Index 病人服务信息记录_IX_号源ID on 病人服务信息记录(号源ID) Tablespace zl9Indexhis;
Create Index 病人服务信息记录_IX_记录ID on 病人服务信息记录(记录ID) Tablespace zl9Indexhis;
Create Index 病人服务信息记录_IX_项目ID on 病人服务信息记录(项目ID) Tablespace zl9Indexhis;
Create Index 病人服务信息记录_IX_医生ID on 病人服务信息记录(医生ID) Tablespace zl9Indexhis;
Create Index 临床出诊变动记录_IX_记录ID on 临床出诊变动记录(记录ID) Tablespace zl9Indexhis;

Create Index 临床出诊变动记录_IX_登记时间 on 临床出诊变动记录(登记时间) Tablespace zl9Indexhis;
Create Index 临床出诊变动记录_IX_原诊室ID on 临床出诊变动记录(原诊室ID) Tablespace zl9Indexhis;
Create Index 临床出诊变动记录_IX_现诊室ID on 临床出诊变动记录(现诊室ID) Tablespace zl9Indexhis;
Create Index 临床出诊变动明细_IX_诊室ID on 临床出诊变动明细(诊室ID) Tablespace zl9Indexhis;

Create Index 临床出诊停诊记录_IX_记录ID on 临床出诊停诊记录(记录ID) Tablespace zl9Indexhis;

Create Index 临床出诊停诊记录_IX_替诊医生ID on 临床出诊停诊记录(替诊医生ID) Tablespace zl9Indexhis;
Create Index 临床出诊停诊记录_IX_申请时间 on 临床出诊停诊记录(申请时间) Tablespace zl9Indexhis;
Create Index 临床出诊停诊记录_IX_审批时间 on 临床出诊停诊记录(审批时间) Tablespace zl9Indexhis;
Create Index 门诊诊室适用科室_IX_科室id on 门诊诊室适用科室(科室id) Tablespace zl9Indexhis;

Create Index 临床出诊记录_IX_诊室ID on 临床出诊记录(诊室ID) Tablespace zl9Indexhis;
Create Index 临床出诊记录_IX_安排ID on 临床出诊记录(安排ID) Tablespace zl9Indexhis;
Create Index 临床出诊记录_IX_号源ID on 临床出诊记录(号源ID) Tablespace zl9Indexhis;
Create Index 临床出诊记录_IX_替诊医生id on 临床出诊记录(替诊医生id) Tablespace zl9Indexhis;
Create Index 临床出诊记录_IX_医生id on 临床出诊记录(医生id) Tablespace zl9Indexhis;
Create Index 临床出诊记录_IX_项目ID on 临床出诊记录(项目ID) Tablespace zl9Indexhis;
Create Index 临床出诊记录_IX_科室ID on 临床出诊记录(科室ID) Tablespace zl9Indexhis;
Create Index 临床出诊记录_IX_相关ID On 临床出诊记录(相关ID) Tablespace zl9Indexhis;
Create Index 临床出诊表_IX_科室ID on 临床出诊表(科室ID) Tablespace zl9Indexhis;
Create Index 临床出诊表_IX_关联ID On 临床出诊表(关联ID) Tablespace zl9Indexhis;
Create Index 临床出诊安排_IX_项目ID on 临床出诊安排(项目ID) Tablespace zl9Indexhis;
Create Index 临床出诊安排_IX_医生id on 临床出诊安排(医生id) Tablespace zl9Indexhis;
Create Index 临床出诊安排_IX_号源ID on 临床出诊安排(号源ID) Tablespace zl9Indexhis;
Create Index 临床出诊安排_IX_出诊ID on 临床出诊安排(出诊ID) Tablespace zl9Indexhis;
Create Index 临床出诊限制_IX_诊室ID on 临床出诊限制(诊室ID) Tablespace zl9Indexhis;
Create Index 临床出诊诊室记录_IX_诊室ID on 临床出诊诊室记录(诊室ID) Tablespace zl9Indexhis;
Create Index 临床出诊诊室_IX_诊室ID on 临床出诊诊室(诊室ID) Tablespace zl9Indexhis;
create Index 临床出诊号源限制_IX_诊室ID on 临床出诊号源限制(诊室ID) Tablespace zl9Indexhis;
create Index 临床出诊号源诊室_IX_诊室ID on 临床出诊号源诊室(诊室ID) Tablespace zl9Indexhis;
Create Index 临床出诊号源_IX_项目ID on 临床出诊号源(项目ID) Tablespace zl9Indexhis;

Create Index 临床出诊号源_IX_医生id on 临床出诊号源(医生id) Tablespace zl9Indexhis;
Create Index 临床出诊号源_IX_医生姓名 on 临床出诊号源(医生姓名) Tablespace zl9Indexhis;
Create Index 三方退款信息_IX_待转出 On 三方退款信息(待转出) Tablespace zl9Indexhis;
Create Index 三方退款信息_IX_记录id On 三方退款信息(记录id) Tablespace Zl9indexhis;
Create Index 费用清单打印_IX_主页ID On 费用清单打印(病人ID,主页ID) Tablespace zl9Indexhis;

Create Index 费用清单打印_IX_打印时间 On 费用清单打印(打印时间) Tablespace zl9Indexhis;
Create Index 费用清单打印_IX_待转出 On 费用清单打印(待转出) Tablespace zl9Indexhis;
Create Index 医保结算明细_Ix_待转出 On 医保结算明细(待转出) Tablespace Zl9indexhis;
Create Index 医保结算明细_IX_NO On 医保结算明细(NO) Pctfree 5 Tablespace zl9Indexhis;
Create Index 医保结算明细_IX_卡类别ID On 医保结算明细(卡类别ID) Tablespace zl9Indexhis;

Create Index 就诊变动记录_IX_登记时间 On 就诊变动记录(登记时间) Tablespace zl9indexhis;
Create Index 就诊变动记录_IX_病人ID On 就诊变动记录(病人ID) Tablespace zl9indexhis;

Create Index 费用补充记录_IX_结算ID On 费用补充记录(结算ID) Tablespace zl9indexhis;
  Create Index 费用补充记录_Ix_缴款组id On 费用补充记录(缴款组id) Tablespace Zl9indexhis;

Create Index 费用补充记录_IX_收费结帐ID On 费用补充记录(收费结帐ID) Tablespace zl9indexhis;
Create Index 费用补充记录_IX_结算序号 On 费用补充记录(结算序号) Tablespace zl9indexhis;
Create Index 费用补充记录_IX_费用状态 On 费用补充记录(费用状态) Tablespace zl9indexhis;
Create Index 费用补充记录_IX_待转出 On 费用补充记录(待转出) Tablespace zl9indexhis;
Create Index 费用补充记录_IX_登记时间 On 费用补充记录(登记时间) Tablespace zl9indexhis;
Create Index 费用补充记录_IX_病人id On 费用补充记录(病人id) Tablespace zl9indexhis;
Create Index 凭条打印记录_IX_待转出 On 凭条打印记录(待转出) Tablespace zl9Indexhis;
Create Index 三方结算交易_IX_待转出 On 三方结算交易(待转出) Tablespace zl9Indexhis;
Create Index 三方结算交易_IX_原预交id On 三方结算交易(原预交id) Tablespace zl9Indexhis;
Create Index 病人卡结算记录_IX_交易时间 On 病人卡结算记录(交易时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人卡结算记录_IX_待转出 On 病人卡结算记录(待转出) Tablespace zl9Indexhis;
Create Index 病人卡结算记录_Ix_结算id On 病人卡结算记录(结算id) Tablespace Zl9indexhis;
Create Index 病人卡结算记录_Ix_结算序号 On 病人卡结算记录(结算序号) Tablespace Zl9indexhis;
Create Index 病人卡结算记录_Ix_交易序号 On 病人卡结算记录(交易序号) Tablespace Zl9indexhis;
Create Index 病人卡结算记录_Ix_病人id On 病人卡结算记录(病人id) Tablespace Zl9indexhis;
Create Index 病人卡结算记录_Ix_登记时间 On 病人卡结算记录(登记时间) Tablespace Zl9indexhis;
Create Index 病人医疗卡变动_IX_变动ID On 病人医疗卡变动(变动ID) Tablespace zl9Indexhis;
Create Index 病人医疗卡变动_IX_卡号 On 病人医疗卡变动(卡号,卡类别ID,变动时间) Tablespace zl9Indexhis;
Create Index 病人医疗卡变动_IX_费用单号 On 病人医疗卡变动(费用单号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人医疗卡信息_IX_挂失时间 On 病人医疗卡信息(挂失时间) Tablespace zl9Indexhis;
Create Index 病人医疗卡信息_IX_发卡日期 On 病人医疗卡信息(发卡日期) Tablespace zl9Indexhis;
Create Index 病人医疗卡信息_IX_终止使用时间 on 病人医疗卡信息(终止使用时间) Tablespace zl9Indexhis;
Create Index 病人医疗卡信息_IX_二维码 On 病人医疗卡信息(二维码) Pctfree 5 Tablespace zl9Indexhis;

Create Index 病人挂号汇总_IX_号码 On 病人挂号汇总(号码) Tablespace zl9Indexhis;
Create Index 病人挂号汇总_IX_项目ID On 病人挂号汇总(项目ID) Tablespace zl9Indexhis;
Create Index 病人挂号汇总_IX_待转出 On 病人挂号汇总(待转出) Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_病人ID On 病人挂号记录(病人ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_接收时间 On 病人挂号记录(接收时间) Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_登记时间 On 病人挂号记录(登记时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_预约时间 On 病人挂号记录(预约时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_发生时间 On 病人挂号记录(发生时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_执行时间 On 病人挂号记录(执行时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_执行状态 On 病人挂号记录(执行状态) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_待转出 On 病人挂号记录(待转出) Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_出诊记录ID on 病人挂号记录(出诊记录ID) Tablespace zl9Indexhis;
Create Index 病人挂号记录_IX_挂号项目ID on 病人挂号记录(挂号项目ID) Tablespace zl9IndexHis;
Create Index 挂号序号状态_IX_日期 On 挂号序号状态(日期) Initrans 20 Tablespace zl9Indexhis;
Create Index 挂号序号状态_IX_登记时间 On 挂号序号状态(登记时间) Initrans 20 Tablespace zl9indexhis;
Create Index 挂号序号状态_IX_号码 On 挂号序号状态(号码) Initrans 20 Tablespace zl9Indexhis;
Create Index 病人转诊记录_IX_待转出 On 病人转诊记录(待转出) Tablespace zl9Indexhis;

Create Index 人员收缴记录_IX_收款员 On 人员收缴记录(收款员) Tablespace zl9Indexhis;
Create Index 人员收缴记录_IX_缴款组ID On 人员收缴记录(缴款组ID) Tablespace zl9Indexhis;
Create Index 人员收缴记录_IX_小组收款ID On 人员收缴记录(小组收款ID) Tablespace zl9Indexhis;
Create Index 人员收缴记录_IX_小组轧账ID On 人员收缴记录(小组轧账ID) Tablespace zl9Indexhis;
Create Index 人员收缴记录_IX_财务收款ID On 人员收缴记录(财务收款ID) Tablespace zl9Indexhis;
Create Index 人员收缴记录_IX_作废时间 On 人员收缴记录(作废时间) Tablespace zl9Indexhis;
Create Index 人员收缴记录_IX_登记时间 On 人员收缴记录(登记时间) Tablespace zl9Indexhis;
Create Index 人员收缴记录_IX_待转出 On 人员收缴记录(待转出) Tablespace zl9Indexhis;
Create Index 人员收缴明细_IX_待转出 On 人员收缴明细(待转出) Tablespace zl9Indexhis;
Create Index 人员收缴票据_IX_待转出 On 人员收缴票据(待转出) Tablespace zl9Indexhis;
Create Index 人员收缴对照_IX_记录ID On 人员收缴对照(记录ID, 性质) Tablespace zl9Indexhis;
Create Index 人员收缴对照_IX_待转出 On 人员收缴对照(待转出) Tablespace zl9Indexhis;
Create Index 人员暂存记录_IX_收缴ID On 人员暂存记录(收缴ID) Tablespace zl9Indexhis;
Create Index 人员暂存记录_IX_收回时间 On 人员暂存记录(收回时间) Tablespace zl9Indexhis;
Create Index 人员暂存记录_IX_登记时间 On 人员暂存记录(登记时间) Tablespace zl9Indexhis;
Create Index 人员暂存记录_IX_领用时间 On 人员暂存记录(领用时间) Tablespace zl9Indexhis;
Create Index 人员暂存记录_IX_收款员 On 人员暂存记录(收款员) Tablespace zl9Indexhis;
Create Index 人员暂存记录_IX_待转出 On 人员暂存记录(待转出) Tablespace zl9Indexhis;
Create Index 人员借款记录_IX_借款人 On 人员借款记录(借款人) Tablespace zl9Indexhis;
Create Index 人员借款记录_IX_申请时间 On 人员借款记录(申请时间) Tablespace zl9Indexhis;
Create Index 人员借款记录_IX_借出人 On 人员借款记录(借出人) Tablespace zl9Indexhis;
Create Index 人员借款记录_IX_借出时间 On 人员借款记录(借出时间) Tablespace zl9Indexhis;
Create Index 人员借款记录_IX_待转出 On 人员借款记录(待转出) Tablespace zl9Indexhis;
Create Index 病人催款记录_IX_病人ID On 病人催款记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人催款记录_IX_打印日期 On 病人催款记录(打印日期) Pctfree 5 Tablespace zl9Indexhis;
Create Index 财务组组长构成_IX_组长ID On 财务组组长构成(组长ID) Pctfree 5 Tablespace zl9indexhis;

Create Index 病人结帐记录_IX_收费时间 On 病人结帐记录(收费时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人结帐记录_IX_病人id On 病人结帐记录(病人id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人结帐记录_IX_待转出 On 病人结帐记录(待转出) Tablespace zl9Indexhis;
Create Index 病人结帐记录_IX_结算状态 On 病人结帐记录(结算状态) Tablespace zl9indexhis;
Create Index 住院费用记录_IX_收费细目id On 住院费用记录(收费细目id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_收入项目id On 住院费用记录(收入项目id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_医嘱序号 On 住院费用记录(医嘱序号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_结帐ID On 住院费用记录(结帐ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_登记时间 On 住院费用记录(登记时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_发生时间 On 住院费用记录(发生时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_病人id On 住院费用记录(病人id,主页ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_保险大类ID On 住院费用记录(保险大类ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 住院费用记录_IX_待转出 On 住院费用记录(待转出) Tablespace zl9Indexhis;

Create Index 门诊费用记录_IX_收费细目id On 门诊费用记录(收费细目id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_收入项目id On 门诊费用记录(收入项目id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_医嘱序号 On 门诊费用记录(医嘱序号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_结帐ID On 门诊费用记录(结帐ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_登记时间 On 门诊费用记录(登记时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_发生时间 On 门诊费用记录(发生时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_病人id On 门诊费用记录(病人id,主页id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_保险大类ID On 门诊费用记录(保险大类ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_挂号ID On 门诊费用记录(挂号ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 门诊费用记录_IX_待转出 On 门诊费用记录(待转出) Tablespace zl9Indexhis;

Create Index 病人费用销帐_IX_申请时间 On 病人费用销帐(申请时间) Tablespace zl9Indexhis;
Create Index 病人费用销帐_IX_核查日期 On 病人费用销帐(核查日期) Tablespace zl9Indexhis;
Create Index 病人费用销帐_IX_待转出 On 病人费用销帐(待转出) Tablespace Zl9indexhis;
Create Index 费用审核记录_IX_病人ID On 费用审核记录(病人ID,主页ID) Tablespace zl9Indexhis;
Create Index 费用审核记录_IX_审核日期 On 费用审核记录(审核日期) Tablespace zl9Indexhis;
Create Index 病人费用销帐_IX_审核时间 On 病人费用销帐(审核时间) Tablespace zl9Indexhis;
Create Index 病人费用销帐_IX_状态 On 病人费用销帐(状态) Tablespace zl9Indexhis;
Create Index 病人退费申请_IX_申请时间 On 病人退费申请(申请时间) Tablespace zl9Indexhis;
Create Index 病人退费申请_IX_审核时间 On 病人退费申请(审核时间) Tablespace zl9Indexhis;
Create Index 病人费用汇总_IX_收入项目id On 病人费用汇总(收入项目id) Tablespace zl9Indexhis;
Create Index 病人结帐汇总_IX_结帐ID On 病人结帐汇总(结帐ID) Tablespace zl9Indexhis;
Create Index 病人结帐汇总_IX_收入项目id On 病人结帐汇总(收入项目id) Tablespace zl9Indexhis;
Create Index 病人结帐汇总_IX_病人id On 病人结帐汇总(病人id,主页id) Tablespace zl9Indexhis;
Create Index 医生收入汇总_IX_执行人 On 医生收入汇总(日期,执行人) Tablespace zl9Indexhis;
Create Index 医生收入汇总_IX_收入项目id On 医生收入汇总(收入项目id) Tablespace zl9Indexhis;
Create Index 病人未结费用_IX_病人id On 病人未结费用(病人id,主页ID) Tablespace zl9Indexhis;
Create Index 病人未结费用_IX_收入项目id On 病人未结费用(收入项目id) Tablespace zl9Indexhis;
Create Index 病人预交记录_IX_主页ID On 病人预交记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人预交记录_IX_结帐id On 病人预交记录(结帐id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人预交记录_IX_收款时间 On 病人预交记录(收款时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人预交记录_IX_结算序号 On 病人预交记录(结算序号) Tablespace zl9Indexhis;
Create Index 病人预交记录_IX_待转出 On 病人预交记录(待转出) Tablespace zl9Indexhis;
Create Index 病人预交记录_IX_交易时间 on 病人预交记录(交易时间) Tablespace zl9Indexhis;
Create Index 病人预交记录_IX_关联交易ID on 病人预交记录(关联交易ID) Tablespace zl9Indexhis;

Create Index 票据入库记录_IX_登记人 On 票据入库记录(登记人) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据入库记录_IX_登记时间 On 票据入库记录(登记时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据入库记录_IX_有无票据 On 票据入库记录(有无票据) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据入库记录_IX_批次 On 票据入库记录(批次) Tablespace zl9Indexhis;
Create Index 票据报损记录_IX_报损人 On 票据报损记录(报损人) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据报损记录_IX_报损时间 On 票据报损记录(报损时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据报损记录_IX_入库ID On 票据报损记录(入库ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据领用记录_IX_领用人 On 票据领用记录(领用人) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据领用记录_IX_批次 On 票据领用记录(批次,票种) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据领用记录_IX_登记时间 On 票据领用记录(登记时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据领用记录_IX_待转出 On 票据领用记录(待转出) Tablespace zl9Indexhis;
Create Index 票据领用记录_IX_入库ID On 票据领用记录(入库ID) Tablespace zl9Indexhis;
Create Index 票据打印内容_IX_NO On 票据打印内容(NO) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据使用明细_IX_领用ID On 票据使用明细(领用ID,票种,性质) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据使用明细_IX_使用时间 On 票据使用明细(使用时间) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据使用明细_IX_打印ID On 票据使用明细(打印ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据使用明细_IX_待转出 On 票据使用明细(待转出) Tablespace zl9Indexhis;
CREATE INDEX 票据使用明细_IX_电子票据ID ON 票据使用明细(电子票据ID) TABLESPACE zl9Indexhis;
Create Index 票据打印明细_IX_使用ID On 票据打印明细(使用ID,NO) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据打印明细_IX_关联票号序号 On 票据打印明细(关联票号序号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 票据打印明细_IX_待转出 On 票据打印明细(待转出) Tablespace zl9Indexhis;
Create Index 票据打印内容_IX_待转出 On 票据打印内容(待转出) Tablespace zl9Indexhis;
Create Index 缴款成员组成_IX_成员ID On 缴款成员组成(成员ID) Pctfree 5 Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[15.药品卫材业务]]
---------------------------------------------------------------------------- 

Create Index 未审药品记录_IX_药品id On 未审药品记录(药品id) Tablespace zl9Indexhis;   
Create Index 未审药品记录_IX_填制日期 On 未审药品记录(填制日期) Tablespace zl9Indexhis;
Create Index 未审药品记录_IX_待转出 On 未审药品记录(待转出) Tablespace zl9Indexhis;  
Create Index 材料结存记录_IX_上次结存id On 材料结存记录(上次结存id) Tablespace zl9Indexhis;
Create Index 材料结存记录_IX_填制日期 On 材料结存记录(填制日期) Tablespace zl9Indexhis;
Create Index 材料结存记录_IX_结存日期 On 材料结存记录(审核日期) Tablespace zl9Indexhis;

Create Index 材料结存明细_IX_结存id On 材料结存明细(结存id) Tablespace zl9Indexhis;
Create Index 材料结存明细_IX_材料id On 材料结存明细(材料id) Tablespace zl9Indexhis;

Create Index 材料结存误差_IX_结存id On 材料结存误差(结存id) Tablespace zl9Indexhis;
Create Index 材料结存误差_IX_材料id On 材料结存误差(材料id) Tablespace zl9Indexhis;

Create Index 配液工作安排_IX_配药台id On 配液工作安排(配药台id) Tablespace ZL9INDEXHIS;

Create Index 药品收发门诊标志_IX_待转出 ON 药品收发门诊标志(待转出) Tablespace Zl9indexhis;
Create Index 药品收发住院标志_IX_待转出 ON 药品收发住院标志(待转出) Tablespace Zl9indexhis;

Create Index 材料质量记录_IX_质量id On 材料质量记录(质量id) Tablespace zl9Indexhis;
Create Index 材料质量记录_IX_材料id On 材料质量记录(材料id) Tablespace zl9Indexhis;
Create Index 材料质量记录_IX_供药单位id On 材料质量记录(供药单位id) Tablespace zl9Indexhis;

Create Index 药品验收记录_IX_供药单位id On 药品验收记录(供药单位id) Tablespace zl9Indexhis;   
Create Index 药品验收记录_IX_NO On 药品验收记录(NO) Tablespace zl9Indexhis;  
Create Index 药品验收明细_IX_药品id On 药品验收明细(药品id) Tablespace zl9Indexhis;  

Create Index 药品用法用量_IX_用法id On 药品用法用量(用法ID) Tablespace zl9Indexhis;
Create Index 材料储备限额_IX_材料ID On 材料储备限额(材料ID) Tablespace zl9Indexhis;

Create Index 处方审查条件_IX_科室ID ON 处方审查条件(科室ID) Tablespace Zl9indexhis;
Create Index 处方审查条件_IX_医生ID ON 处方审查条件(医生ID) Tablespace Zl9indexhis;
Create Index 处方审查条件_IX_诊断ID ON 处方审查条件(诊断ID) Tablespace Zl9indexhis;
Create Index 处方审查条件_IX_疾病ID ON 处方审查条件(疾病ID) Tablespace Zl9indexhis;
Create Index 处方审查条件_IX_药名ID ON 处方审查条件(药名ID) Tablespace Zl9indexhis;
Create Index 处方审查记录_Ix_挂号id On 处方审查记录(挂号id) Tablespace Zl9indexhis;

Create Index 处方审查记录_Ix_病人id On 处方审查记录(病人id, 主页id) Tablespace Zl9indexhis;
Create Index 处方审查记录_Ix_审查时间 On 处方审查记录(审查时间) Tablespace Zl9indexhis;
Create Index 处方审查记录_Ix_状态 On 处方审查记录(状态) Tablespace Zl9indexhis;
Create Index 处方审查记录_Ix_锁定用户 On 处方审查记录(锁定用户) Tablespace Zl9indexhis;
Create Index 处方审查记录_Ix_待转出 On 处方审查记录(待转出) Tablespace Zl9indexhis;
Create Index 处方审查明细_Ix_医嘱id On 处方审查明细(医嘱id) Tablespace Zl9indexhis;

Create Index 处方审查明细_IX_待转出 ON 处方审查明细(待转出) Tablespace Zl9indexhis;
Create Index 处方审查结果_Ix_审查项目id On 处方审查结果(审查项目id) Tablespace Zl9indexhis;
Create Index 处方审查结果_Ix_医嘱id On 处方审查结果(医嘱id) Tablespace Zl9indexhis;

Create Index 处方审查结果_IX_待转出 ON 处方审查结果(待转出) Tablespace Zl9indexhis;
Create Index 收费调价记录_IX_收费细目id On 收费调价记录(收费细目id) Tablespace zl9Indexhis;
Create Index 收费调价记录_IX_价格等级 on 收费调价记录(价格等级) Tablespace zl9Indexhis;

Create Index 收费调价记录_IX_收入项目id On 收费调价记录(收入项目id) Tablespace zl9Indexhis;
Create Index 收费调价记录_IX_审核标志 On 收费调价记录(审核标志) Tablespace zl9Indexhis;
Create Index 收费调价记录_IX_填制日期 On 收费调价记录(填制日期) Tablespace zl9Indexhis;
Create Index 药品财务审核_IX_审核日期 On 药品财务审核(审核日期) Tablespace zl9Indexcis;
Create Index 药品采购计划_IX_编制日期 On 药品采购计划(编制日期) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品采购计划_IX_审核日期 On 药品采购计划(审核日期) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品采购计划_IX_复核日期 On 药品采购计划(复核日期) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品采购计划_IX_合并计划id On 药品采购计划(合并计划id) Tablespace zl9Indexhis;
Create Index 药品计划内容_IX_药品id On 药品计划内容(药品id) Tablespace zl9Indexhis;
Create Index 药品退药计划_IX_药品id On 药品退药计划(药品id) Tablespace zl9Indexhis;
Create Index 药品退药计划_IX_供药单位id On 药品退药计划(供药单位id) Tablespace zl9Indexhis;
Create Index 药品退药计划_IX_填制日期 On 药品退药计划(填制日期) Tablespace zl9Indexhis;
Create Index 药品退药计划_IX_审核日期 On 药品退药计划(审核日期) Tablespace zl9Indexhis;
Create Index 材料采购计划_IX_NO On 材料采购计划(no) Tablespace zl9Indexhis;
Create Index 材料计划内容_IX_材料id On 材料计划内容(材料id) Tablespace zl9Indexhis;
Create Index 药品留存计划_IX_状态 On 药品留存计划(部门ID,状态) Tablespace zl9Indexhis;
Create Index 药品留存计划_IX_留存ID On 药品留存计划(留存ID) Tablespace zl9Indexhis;
Create Index 药品留存计划_IX_待转出 On 药品留存计划(待转出) Tablespace zl9Indexhis;

Create Index 药品库存_IX_药品id On 药品库存(药品id) Tablespace zl9Indexhis;
Create Index 药品库存_IX_商品条码 On 药品库存(商品条码) Tablespace zl9Indexhis;
Create Index 药品库存_IX_内部条码 On 药品库存(内部条码) Tablespace zl9Indexhis;
Create Index 药品结存汇总_IX_药品id On 药品结存汇总(药品id) Tablespace zl9Indexhis;
Create Index 药品结存汇总_IX_结存id On 药品结存汇总(结存id) Tablespace zl9Indexhis;
Create Index 药品结存汇总_IX_库房id On 药品结存汇总(库房id) Tablespace zl9Indexhis;
Create Index 药品结存汇总_IX_入出类别id On 药品结存汇总(入出类别id) Tablespace zl9Indexhis;
Create Index 药品结存_IX_药品id On 药品结存(药品id) Tablespace zl9Indexhis;
Create Index 药品留存_IX_药品id On 药品留存(药品id) Tablespace zl9Indexhis;
Create Index 药品收发汇总_IX_药品id On 药品收发汇总(药品id) Tablespace zl9Indexhis;
Create Index 药品收发汇总_IX_类别id On 药品收发汇总(类别id) Tablespace zl9Indexhis;
Create Index 未发药品记录_IX_填制日期 On 未发药品记录(填制日期) Tablespace zl9Indexhis;
Create Index 未发药品记录_IX_对方部门ID On 未发药品记录(对方部门ID) Tablespace zl9Indexhis;
Create Index 未发药品记录_IX_主页ID On 未发药品记录(病人ID,主页ID) Tablespace zl9Indexhis;
Create Index 未发药品记录_IX_排队状态 On 未发药品记录(排队状态) Tablespace zl9Indexcis;
Create Index 药品收发记录_IX_费用id On 药品收发记录(费用id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_药品id On 药品收发记录(药品id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_入出类别id On 药品收发记录(入出类别id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_供药单位id On 药品收发记录(供药单位id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_填制日期 On 药品收发记录(填制日期) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_审核日期 On 药品收发记录(审核日期) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_价格ID On 药品收发记录(价格ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_汇总发药号 On 药品收发记录(汇总发药号) Pctfree 5 Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_商品条码 On 药品收发记录(商品条码) Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_内部条码 On 药品收发记录(内部条码) Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_待转出 On 药品收发记录(待转出) Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_计划id On 药品收发记录(计划id) Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_病人ID On 药品收发记录(病人ID,主页ID) Tablespace zl9Indexhis;
Create Index 药品收发记录_IX_医嘱id On 药品收发记录(医嘱id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 收发记录补充信息_IX_待转出 On 收发记录补充信息(待转出) Tablespace zl9Indexhis;

Create Index 药品签名记录_IX_证书ID On 药品签名记录(证书ID) Tablespace zl9Indexhis;
Create Index 药品签名记录_IX_待转出 On 药品签名记录(待转出) Tablespace zl9Indexhis;
Create Index 药品签名明细_IX_收发ID On 药品签名明细(收发ID) Tablespace zl9Indexhis;
Create Index 药品签名明细_IX_待转出 On 药品签名明细(待转出) Tablespace zl9Indexhis;

Create Index 成本价调价信息_IX_执行日期 On 成本价调价信息(执行日期) Tablespace zl9Indexhis;
Create Index 成本价调价信息_IX_药品ID On 成本价调价信息(药品ID) Tablespace zl9Indexhis;
Create Index 成本价调价信息_IX_收发id On 成本价调价信息(收发id) Tablespace zl9Indexhis;
Create Index 成本价调价信息_IX_供药单位ID On 成本价调价信息(供药单位ID) Tablespace zl9Indexhis;

Create Index 药品价格记录_IX_药品ID On 药品价格记录(药品ID) Tablespace zl9Indexhis;
Create Index 药品价格记录_IX_原价id On 药品价格记录(原价id) Tablespace zl9Indexhis;
Create Index 药品价格记录_IX_收发id On 药品价格记录(收发id) Tablespace zl9Indexhis;
Create Index 药品价格记录_IX_供药单位ID On 药品价格记录(供药单位ID) Tablespace zl9Indexhis;
Create Index 药品价格记录_IX_调价汇总号 On 药品价格记录(调价汇总号) Tablespace zl9Indexhis;
Create Index 药品价格记录_IX_记录状态 On 药品价格记录(记录状态) Tablespace zl9Indexhis;

Create Index 药品质量记录_IX_药品id On 药品质量记录(药品id) Tablespace zl9Indexhis;
Create Index 药品质量记录_IX_供药单位id On 药品质量记录(供药单位id) Tablespace zl9Indexhis;
Create Index 药品质量记录_IX_登记时间 On 药品质量记录(登记时间) Tablespace zl9Indexhis;
Create Index 药品质量记录_IX_处理时间 On 药品质量记录(处理时间) Tablespace zl9Indexhis;

Create Index 药品结存记录_IX_填制日期 On 药品结存记录(填制日期) Tablespace zl9Indexhis;
Create Index 药品结存记录_IX_结存日期 On 药品结存记录(审核日期) Tablespace zl9Indexhis;
Create Index 药品结存明细_IX_药品id On 药品结存明细(药品id) Tablespace zl9Indexhis;
Create Index 药品结存误差_IX_药品id On 药品结存误差(药品id) Tablespace zl9Indexhis;
Create Index 药品结存误差_IX_结存id On 药品结存误差(结存id) Tablespace zl9Indexhis;
Create Index 暂存药品记录_IX_病人ID On 暂存药品记录(病人ID) Tablespace zl9Indexhis;
Create Index 暂存药品记录_IX_登记时间 On 暂存药品记录(登记时间) Tablespace zl9Indexhis;
Create Index 暂存药品记录_IX_医嘱ID On 暂存药品记录(医嘱ID, 发送号) Tablespace zl9Indexhis;

Create Index 输液配药记录_IX_执行时间 On 输液配药记录(执行时间) Tablespace zl9Indexhis;
Create Index 输液配药记录_IX_操作时间 On 输液配药记录(操作时间,操作状态) Tablespace zl9Indexhis;
Create Index 输液配药记录_IX_摆药单号 On 输液配药记录(摆药单号) Pctfree 20 Tablespace zl9Indexhis;
Create Index 输液配药记录_IX_瓶签号 On 输液配药记录(瓶签号) Tablespace zl9Indexhis;
Create Index 输液配药记录_IX_待转出 On 输液配药记录(待转出) Tablespace zl9Indexhis;
Create Index 输液配药记录_IX_打印时间 On 输液配药记录(打印时间) Tablespace zl9Indexcis;
Create Index 输液配药记录_IX_病人ID On 输液配药记录(病人ID,主页ID) Tablespace zl9Indexhis;
Create Index 输液配药记录_IX_发送时间 On 输液配药记录(发送时间) Tablespace zl9Indexhis;

Create Index 输液配药状态_IX_操作时间 On 输液配药状态(操作时间,操作类型) Tablespace zl9Indexhis;
Create Index 输液配药状态_IX_待转出 On 输液配药状态(待转出) Tablespace zl9Indexhis;
Create Index 输液配药内容_IX_收发ID On 输液配药内容(收发ID) Tablespace zl9Indexhis;
Create Index 输液配药内容_IX_待转出 On 输液配药内容(待转出) Tablespace zl9Indexhis;
Create Index 输液配药附费_IX_待转出 On 输液配药附费(待转出) Tablespace zl9Indexhis;
Create Index 输液配药附费_IX_病人id On 输液配药附费(病人id) Tablespace zl9Indexcis;

Create Index 配置收费方案_IX_诊疗ID On 配置收费方案(诊疗ID) Tablespace zl9Indexhis;

Create Index 应付记录_IX_收发ID On 应付记录(收发ID) Tablespace zl9Indexhis;
Create Index 应付记录_IX_单位ID On 应付记录(单位ID) Tablespace zl9Indexhis;
Create Index 应付记录_IX_付款序号 On 应付记录(付款序号) Tablespace zl9Indexhis;
Create Index 应付记录_IX_审核日期 On 应付记录(审核日期) Tablespace zl9Indexhis;
Create Index 应付记录_IX_发票号 On 应付记录(发票号) Tablespace zl9Indexhis;
Create Index 应付记录_IX_随货单号 On 应付记录(随货单号) Tablespace zl9Indexhis;
Create Index 应付记录_IX_入库单据号 On 应付记录(入库单据号) Tablespace zl9Indexhis;
Create Index 付款记录_IX_单位id On 付款记录(单位id) Tablespace zl9Indexhis;
Create Index 付款记录_IX_填制日期 On 付款记录(填制日期) Tablespace zl9Indexhis;
Create Index 付款记录_IX_预审日期 On 付款记录(预审日期) Tablespace zl9Indexhis;
Create Index 付款记录_IX_审核日期 On 付款记录(审核日期) Tablespace zl9Indexhis;
Create Index 付款记录_IX_付款序号 On 付款记录(付款序号) Tablespace zl9Indexhis;

Create Index 调价汇总记录_IX_执行日期 On 调价汇总记录(执行日期) Tablespace zl9Indexhis;
Create Index 调价汇总记录_IX_填制日期 On 调价汇总记录(填制日期) Tablespace zl9Indexhis;
Create Index 成本价调价信息_IX_调价汇总号 On 成本价调价信息(调价汇总号) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[16.临床医嘱]]
----------------------------------------------------------------------------
Create Index 急诊分诊记录_IX_就诊ID On 急诊分诊记录(就诊ID) Tablespace zl9IndexCis;
Create Index 急诊分诊记录_IX_登记时间 On 急诊分诊记录(登记时间) Tablespace zl9IndexCis;
Create Index 急诊分诊记录_IX_待转出 On 急诊分诊记录(待转出) Tablespace zl9IndexCis;
Create Index 急诊病人评分指标_IX_待转出 On 急诊病人评分指标(待转出) Tablespace zl9IndexCis;
Create Index 急诊病人评分_IX_分诊ID On 急诊病人评分(分诊ID) Tablespace zl9IndexCis;
Create Index 急诊病人评分_IX_待转出 On 急诊病人评分(待转出) Tablespace zl9IndexCis;
Create Index 急诊就诊记录_IX_病人ID On 急诊就诊记录(病人ID) Tablespace zl9IndexCis;
Create Index 急诊就诊记录_IX_挂号ID On 急诊就诊记录(挂号ID) Tablespace zl9IndexCis;
Create Index 急诊就诊记录_IX_登记时间 On 急诊就诊记录(登记时间) Tablespace zl9IndexCis;
Create Index 急诊就诊记录_IX_待转出 On 急诊就诊记录(待转出) Tablespace zl9IndexCis;
Create Index 路径通用诊疗项目_IX_诊疗项目ID On 路径通用诊疗项目(诊疗项目ID) Tablespace zl9Indexcis;

Create Index 药嘱禁忌说明_IX_医嘱B On 药嘱禁忌说明(医嘱B) Tablespace zl9Indexcis;

Create Index 药嘱禁忌说明_IX_待转出 On 药嘱禁忌说明(待转出) Tablespace zl9Indexcis;
Create Index 病人用药清单_IX_病人ID on 病人用药清单(病人ID, 主页ID) Tablespace zl9indexhis;
Create Index 病人用药清单_IX_用法ID on 病人用药清单(用法ID) Tablespace zl9indexhis;
Create Index 病人用药清单_IX_煎法ID on 病人用药清单(煎法ID) Tablespace zl9indexhis;

Create Index 病人用药清单_IX_收费细目ID on 病人用药清单(收费细目ID) Tablespace zl9indexhis;
Create Index 病人用药清单_IX_诊疗项目ID on 病人用药清单(诊疗项目ID) Tablespace zl9indexhis;
Create Index 病人用药清单_IX_开始时间 on 病人用药清单(开始时间) Tablespace zl9indexhis;
Create Index 病人用药清单_IX_待转出 on 病人用药清单(待转出) Tablespace zl9indexhis;
Create Index 病人用药配方_IX_收费细目ID on 病人用药配方(收费细目ID) Tablespace zl9indexhis;

Create Index 病人用药配方_IX_诊疗项目ID on 病人用药配方(诊疗项目ID) Tablespace zl9indexhis;
Create Index 病人用药配方_IX_待转出 on 病人用药配方(待转出) Tablespace zl9indexhis;
Create Index 病人危急值记录_IX_病人ID On 病人危急值记录(病人ID,主页ID)  Tablespace zl9Indexcis;
Create Index 病人危急值记录_IX_挂号单 On 病人危急值记录(挂号单)  Tablespace zl9Indexcis;
Create Index 病人危急值记录_IX_医嘱ID On 病人危急值记录(医嘱ID) Tablespace zl9Indexcis;
Create Index 病人危急值记录_IX_报告时间 On 病人危急值记录(报告时间)  Tablespace zl9Indexcis;
Create Index 病人危急值记录_IX_待转出 On 病人危急值记录(待转出) Tablespace zl9Indexcis;
Create Index 病人危急值医嘱_IX_医嘱ID On 病人危急值医嘱(医嘱ID) Tablespace zl9Indexcis;

Create Index 病人危急值医嘱_IX_待转出 On 病人危急值医嘱(待转出) Tablespace zl9Indexcis;
Create Index 病人危急值病历_IX_待转出 On 病人危急值病历(待转出) Tablespace zl9Indexcis;

Create Index 医嘱申请单文件_IX_待转出 On 医嘱申请单文件(待转出) Tablespace zl9Indexcis;

Create Index 医嘱申请单文件_IX_文件ID On 医嘱申请单文件(文件ID) Tablespace zl9Indexcis;
Create Index 医嘱报告内容_IX_待转出 On 医嘱报告内容(待转出) Tablespace zl9Indexcis;

Create Index 疾病申报反馈_IX_登记时间 On 疾病申报反馈(登记时间) Tablespace zl9Indexcis;

Create Index 疾病申报反馈_IX_待转出 On 疾病申报反馈(待转出) Tablespace zl9Indexcis;
Create Index 疾病阳性记录_IX_病人ID On 疾病阳性记录(病人ID,主页ID)  Tablespace zl9Indexcis;
Create Index 疾病阳性记录_IX_登记时间 On 疾病阳性记录(登记时间)  Tablespace zl9Indexcis;
Create Index 疾病阳性记录_IX_挂号单 On 疾病阳性记录(挂号单)  Tablespace zl9Indexcis;
Create Index 疾病阳性记录_IX_待转出 On 疾病阳性记录(待转出) Tablespace zl9Indexcis;
Create Index 疾病阳性记录_IX_文件ID On 疾病阳性记录(文件ID) Tablespace zl9Indexcis;
Create Index 疾病阳性记录_IX_医嘱ID On 疾病阳性记录(医嘱ID) Tablespace zl9Indexcis;
Create Index 疾病报告反馈_IX_待转出 On 疾病报告反馈(待转出) Tablespace zl9Indexcis;
Create Index 疾病报告反馈_IX_登记时间 On 疾病报告反馈(登记时间) Tablespace zl9Indexcis;

Create Index 输血检验结果_IX_检验项目ID On 输血检验结果(检验项目ID) Tablespace zl9Indexcis;
Create Index 输血检验结果_IX_待转出 On 输血检验结果(待转出) Tablespace zl9Indexcis;
Create Index 排队记录_IX_病人ID On 排队记录(病人ID) Tablespace zl9Indexcis;
Create Index 排队记录_IX_日期 On 排队记录(日期) Tablespace zl9Indexcis;
Create Index 排队记录_IX_呼叫标志 On 排队记录(呼叫标志) Tablespace zl9Indexcis;
Create Index 座位状况记录_IX_病人ID On 座位状况记录(病人ID) Tablespace zl9Indexcis;
Create Index 座位状况记录_IX_收费细目id On 座位状况记录(收费细目id) Tablespace zl9Indexcis;
Create Index 门诊输液操作日志_IX_时间 On 门诊输液操作日志(时间) Tablespace ZL9INDEXCIS;
Create Index 门诊输液操作日志_IX_操作员 On 门诊输液操作日志(操作员) Tablespace ZL9INDEXCIS;
Create Index 门诊输液操作日志_IX_挂号单 On 门诊输液操作日志(挂号单) Tablespace ZL9INDEXCIS;
Create Index 门诊穿刺台_Ix_待穿病人id On 门诊穿刺台(待穿病人id) Pctfree 5 Tablespace Zl9indexcis;

Create Index 病人医嘱记录_IX_相关ID On 病人医嘱记录(相关ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_主页ID On 病人医嘱记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_诊疗项目ID On 病人医嘱记录(诊疗项目id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_收费细目ID On 病人医嘱记录(收费细目id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_挂号单 On 病人医嘱记录(挂号单) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_开嘱时间 On 病人医嘱记录(开嘱时间,医嘱状态) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_开始执行时间 On 病人医嘱记录(开始执行时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_手术时间 On 病人医嘱记录(手术时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_审核状态 On 病人医嘱记录(审核状态) Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_申请序号 On 病人医嘱记录(申请序号) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_配方ID On 病人医嘱记录(配方ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_待转出 On 病人医嘱记录(待转出) Tablespace zl9Indexcis;
Create Index 病人医嘱记录_IX_处方序号 On 病人医嘱记录(处方序号) Pctfree 5 Tablespace zl9Indexcis;

Create Index 病人医嘱状态_IX_操作时间 On 病人医嘱状态(操作时间,操作类型) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱状态_IX_签名ID On 病人医嘱状态(签名ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱状态_IX_待转出 On 病人医嘱状态(待转出) Tablespace zl9Indexcis;
Create Index 病人医嘱计价_IX_收费细目ID On 病人医嘱计价(收费细目ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱计价_IX_待转出 On 病人医嘱计价(待转出) Tablespace zl9Indexcis;
Create Index 病人医嘱发送_IX_发送号 On 病人医嘱发送(发送号) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱发送_IX_发送时间 On 病人医嘱发送(发送时间,执行状态) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱发送_IX_首次时间 On 病人医嘱发送(首次时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱发送_IX_报到时间 On 病人医嘱发送(报到时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱发送_IX_样本条码 On 病人医嘱发送(样本条码) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱发送_IX_接收批次 On 病人医嘱发送(接收批次) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱发送_IX_待转出 On 病人医嘱发送(待转出) Tablespace zl9Indexcis;
Create Index 病人医嘱报告_IX_病历ID On 病人医嘱报告(病历ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱报告_IX_待转出 On 病人医嘱报告(待转出) Tablespace zl9Indexcis;
Create Index 病人医嘱报告_IX_RISID On 病人医嘱报告(RISID)  Tablespace zl9Indexcis;
Create Index 病人医嘱报告_IX_报告ID On 病人医嘱报告(报告ID) Tablespace zl9Indexcis;
Create Index 病人医嘱发送_IX_发送执行 On 病人医嘱发送(发送时间,执行部门id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱发送_IX_标本发送批号 On 病人医嘱发送(标本发送批号) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱异常记录_IX_病人ID On 病人医嘱异常记录(病人ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱异常记录_IX_NO On 病人医嘱异常记录(NO,记录性质) Pctfree 5 Tablespace zl9Indexcis;

Create Index 病人医嘱附费_IX_NO	On 病人医嘱附费(NO,记录性质) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱附费_IX_待转出 On 病人医嘱附费(待转出) Tablespace zl9Indexcis;
Create Index 病人医嘱附件_IX_待转出 On 病人医嘱附件(待转出) Tablespace zl9Indexcis;
Create Index 医嘱签名记录_IX_证书ID On 医嘱签名记录(证书ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 医嘱签名记录_IX_待转出 On 医嘱签名记录(待转出) Tablespace zl9Indexcis;
Create Index 病人医嘱打印_IX_主页ID On 病人医嘱打印(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱打印_IX_打印时间 On 病人医嘱打印(打印时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人医嘱打印_IX_待转出 On 病人医嘱打印(待转出) Tablespace zl9Indexcis;
Create Index 医嘱执行时间_Ix_要求时间 On 医嘱执行时间(要求时间) Pctfree 5 Tablespace Zl9indexcis;
Create Index 医嘱执行时间_IX_待转出 On 医嘱执行时间(待转出) Tablespace zl9Indexcis;
Create Index 医嘱执行计价_IX_收费细目id On 医嘱执行计价(收费细目id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 医嘱执行计价_IX_待转出 On 医嘱执行计价(待转出) Tablespace zl9Indexcis;
Create Index 医嘱执行打印_IX_待转出 On 医嘱执行打印(待转出) Tablespace zl9Indexcis;
Create Index 执行打印记录_IX_流水号 On 执行打印记录(流水号) Tablespace zl9Indexcis;
Create Index 执行打印记录_IX_待转出 On 执行打印记录(待转出) Tablespace zl9Indexcis;

Create Index 病人医嘱执行_IX_执行时间 On 病人医嘱执行(执行时间) Tablespace zl9Indexcis;
Create Index 病人医嘱执行_IX_流水号 On 病人医嘱执行(流水号) Tablespace zl9Indexcis;
Create Index 病人医嘱执行_IX_待转出 On 病人医嘱执行(待转出) Tablespace zl9Indexcis;
Create Index 诊疗单据打印_IX_待转出 On 诊疗单据打印(待转出) Tablespace zl9Indexcis;
Create Index 输血申请记录_IX_待转出 On 输血申请记录(待转出) Tablespace zl9Indexcis;
Create Index 报告查阅记录_IX_病历ID On 报告查阅记录(病历ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 报告查阅记录_IX_待转出 On 报告查阅记录(待转出) Tablespace zl9Indexcis;
Create index 输血申请项目_IX_诊疗项目ID on 输血申请项目 (诊疗项目ID) tablespace ZL9INDEXCIS;
Create index 输血申请项目_IX_待转出 on 输血申请项目 (待转出) tablespace ZL9INDEXCIS;

Create Index 业务消息清单_IX_病人ID On 业务消息清单(病人ID,就诊ID) Tablespace zl9Indexcis;
Create Index 业务消息清单_IX_登记时间 On 业务消息清单(登记时间) Tablespace zl9Indexcis;
Create Index 业务消息状态_IX_阅读时间 On 业务消息状态(阅读时间) Tablespace zl9Indexcis;

----------------------------------------------------------------------------
--[[17.临床路径]]
----------------------------------------------------------------------------
Create Index 门诊路径报表文件_IX_路径ID On 门诊路径报表文件(路径ID) Tablespace zl9Indexcis;

Create Index 门诊路径报表序号_IX_报表ID On 门诊路径报表序号(报表ID) Tablespace zl9Indexcis;

Create Index 门诊路径病历_IX_文件ID On 门诊路径病历(文件ID) Tablespace zl9Indexcis;
Create Index 门诊路径阶段_IX_父ID On 门诊路径阶段(父ID) Tablespace zl9Indexcis;

Create Index 门诊路径评估_IX_阶段ID On 门诊路径评估(阶段ID) Tablespace zl9Indexcis;

Create Index 门诊路径项目_IX_版本号 On 门诊路径项目(路径ID,版本号) Tablespace zl9Indexcis;

Create Index 门诊路径项目_IX_阶段ID On 门诊路径项目(阶段ID) Tablespace zl9Indexcis;
Create Index 门诊路径项目_IX_图标ID On 门诊路径项目(图标ID) Tablespace zl9Indexcis;
Create Index 门诊路径评估条件_IX_评估ID On 门诊路径评估条件(评估ID) Tablespace zl9Indexcis;

Create Index 门诊路径评估条件_IX_项目ID On 门诊路径评估条件(项目ID) Tablespace zl9Indexcis;
Create Index 门诊路径医嘱内容_IX_相关ID On 门诊路径医嘱内容(相关ID) Tablespace zl9Indexcis;

Create Index 门诊路径医嘱内容_IX_诊疗项目ID On 门诊路径医嘱内容(诊疗项目ID) Tablespace zl9Indexcis;
Create Index 门诊路径医嘱内容_IX_收费细目ID On 门诊路径医嘱内容(收费细目ID) Tablespace zl9Indexcis;
Create Index 门诊路径医嘱内容_IX_执行科室ID On 门诊路径医嘱内容(执行科室ID) Tablespace zl9Indexcis;
Create Index 门诊路径医嘱内容_IX_配方ID On 门诊路径医嘱内容(配方ID) Tablespace zl9Indexcis;
Create Index 门诊路径医嘱_IX_医嘱内容ID On 门诊路径医嘱(医嘱内容ID) Tablespace zl9Indexcis;

Create Index 门诊路径医嘱变动_IX_诊疗项目ID On 门诊路径医嘱变动(诊疗项目ID) Tablespace zl9Indexcis;

Create Index 门诊路径医嘱变动_IX_收费细目ID On 门诊路径医嘱变动(收费细目Id)   Tablespace zl9Indexcis;
Create Index 门诊路径医嘱变动_IX_配方ID On 门诊路径医嘱变动(配方ID)  Tablespace zl9Indexcis;
Create Index 病人门诊路径_IX_病人ID On 病人门诊路径(病人ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人门诊路径_IX_挂号ID On 病人门诊路径(挂号ID) Pctfree 5 Tablespace zl9Indexcis;

Create Index 病人门诊路径_IX_路径ID On 病人门诊路径(路径ID,版本号) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人门诊路径_IX_导入时间 On 病人门诊路径(导入时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人门诊路径_IX_疾病ID On 病人门诊路径(疾病ID) Tablespace zl9Indexcis;
Create Index 病人门诊路径_IX_诊断ID On 病人门诊路径(诊断ID) Tablespace zl9Indexcis;
Create Index 病人门诊路径_IX_未导入原因 On 病人门诊路径(未导入原因) Tablespace zl9Indexcis;
Create Index 病人门诊路径_IX_待转出 On 病人门诊路径(待转出) Tablespace zl9Indexcis;
Create Index 病人门诊路径记录_IX_待转出 On 病人门诊路径记录(待转出) Tablespace zl9Indexcis;
Create Index 病人门诊路径记录_IX_挂号ID On 病人门诊路径记录(挂号ID) Tablespace zl9Indexcis;

Create Index 病人门诊路径评估_IX_日期 On 病人门诊路径评估(日期) Pctfree 5 Tablespace zl9Indexcis;

Create Index 病人门诊路径评估_IX_登记时间 On 病人门诊路径评估(登记时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人门诊路径评估_IX_阶段ID On 病人门诊路径评估(阶段ID) Tablespace zl9Indexcis;
Create Index 病人门诊路径评估_IX_变异原因 On 病人门诊路径评估(变异原因) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人门诊路径评估_IX_待转出 On 病人门诊路径评估(待转出) Tablespace zl9Indexcis;
Create Index 病人门诊路径变异_IX_变异原因 On 病人门诊路径变异(变异原因) Tablespace zl9Indexcis;

Create Index 病人门诊路径变异_IX_待转出 On 病人门诊路径变异(待转出) Tablespace zl9Indexcis;
Create Index 病人门诊路径执行_IX_日期 On 病人门诊路径执行(日期) Pctfree 5 Tablespace zl9Indexcis;

Create Index 病人门诊路径执行_IX_路径记录ID On 病人门诊路径执行(路径记录ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人门诊路径执行_IX_阶段ID On 病人门诊路径执行(阶段ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人门诊路径执行_IX_项目ID On 病人门诊路径执行(项目ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人门诊路径执行_IX_图标ID On 病人门诊路径执行(图标ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人门诊路径执行_IX_登记时间 On 病人门诊路径执行(登记时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人门诊路径执行_IX_变异原因 On 病人门诊路径执行(变异原因) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人门诊路径执行_IX_待转出 On 病人门诊路径执行(待转出) Tablespace zl9Indexcis;
Create Index 病人门诊路径指标_IX_日期 On 病人门诊路径指标(日期) Pctfree 5 Tablespace zl9Indexcis;

Create Index 病人门诊路径指标_IX_阶段ID On 病人门诊路径指标(阶段ID) Tablespace zl9Indexcis;
Create Index 病人门诊路径指标_IX_待转出 On 病人门诊路径指标(待转出) Tablespace zl9Indexcis;
Create Index 病人门诊路径医嘱_IX_病人医嘱ID On 病人门诊路径医嘱(病人医嘱ID) Pctfree 5 Tablespace zl9Indexhis;

Create Index 病人门诊路径医嘱_IX_待转出 On 病人门诊路径医嘱(待转出) Tablespace zl9Indexcis;
Create Index 病人门诊出径记录_IX_路径记录ID On 病人门诊出径记录(路径记录ID) Tablespace zl9Indexhis;

Create Index 病人门诊出径记录_IX_待转出 On 病人门诊出径记录(待转出) Tablespace zl9Indexhis;
Create Index 病人门诊出径记录_IX_病人ID On 病人门诊出径记录(病人ID) Tablespace zl9Indexcis;
Create Index 病人门诊出径记录_IX_挂号ID On 病人门诊出径记录(挂号ID) Tablespace zl9Indexcis;
Create Index 病人门诊路径取消_IX_病人ID On 病人门诊路径取消(病人ID) Tablespace zl9Indexcis;

Create Index 病人门诊路径取消_IX_挂号ID On 病人门诊路径取消(挂号ID) Tablespace zl9Indexcis;
Create Index 路径医嘱变动_IX_诊疗项目ID On 路径医嘱变动(诊疗项目ID) Tablespace zl9Indexcis;
Create Index 路径医嘱变动_IX_收费细目ID On 路径医嘱变动(收费细目Id)   Tablespace zl9Indexcis;
Create Index 路径医嘱变动_IX_配方ID On 路径医嘱变动(配方ID)  Tablespace zl9Indexcis;
Create Index 病人临床路径_IX_病人ID On 病人临床路径(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人临床路径_IX_路径ID On 病人临床路径(路径ID,版本号) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人临床路径_IX_导入时间 On 病人临床路径(导入时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人临床路径_IX_疾病ID On 病人临床路径(疾病ID) Tablespace zl9Indexcis;
Create Index 病人临床路径_IX_诊断ID On 病人临床路径(诊断ID) Tablespace zl9Indexcis;
Create Index 病人临床路径_IX_未导入原因 On 病人临床路径(未导入原因) Tablespace zl9Indexcis;
Create Index 病人临床路径_IX_待转出 On 病人临床路径(待转出) Tablespace zl9Indexcis;

Create Index 病人路径变异_IX_变异原因 On 病人路径变异(变异原因) Tablespace zl9Indexcis;
Create Index 病人路径变异_IX_待转出 On 病人路径变异(待转出) Tablespace zl9Indexcis;
Create Index 病人路径医嘱变异_IX_变异原因 On 病人路径医嘱变异(变异原因) Tablespace zl9Indexcis;
Create Index 病人路径医嘱变异_IX_待转出 On 病人路径医嘱变异(待转出) Tablespace zl9Indexcis;
Create Index 病人路径医嘱_IX_病人医嘱ID On 病人路径医嘱(病人医嘱ID) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病人路径医嘱_IX_待转出 On 病人路径医嘱(待转出) Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_日期 On 病人路径执行(日期) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_路径记录ID On 病人路径执行(路径记录ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_阶段ID On 病人路径执行(阶段ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_项目ID On 病人路径执行(项目ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_图标ID On 病人路径执行(图标ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_登记时间 On 病人路径执行(登记时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_变异原因 On 病人路径执行(变异原因) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径执行_IX_待转出 On 病人路径执行(待转出) Tablespace zl9Indexcis;

Create Index 病人路径评估_IX_日期 On 病人路径评估(日期) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径评估_IX_登记时间 On 病人路径评估(登记时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径评估_IX_阶段ID On 病人路径评估(阶段ID) Tablespace zl9Indexcis;
Create Index 病人路径评估_IX_变异原因 On 病人路径评估(变异原因) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径评估_IX_待转出 On 病人路径评估(待转出) Tablespace zl9Indexcis;
Create Index 病人合并路径_IX_主页ID On 病人合并路径(病人ID,主页ID) Tablespace zl9Indexhis;
Create Index 病人合并路径_IX_版本号 On 病人合并路径(路径ID,版本号) Tablespace zl9Indexhis;
Create Index 病人合并路径_IX_疾病ID On 病人合并路径(疾病ID) Tablespace zl9Indexhis;
Create Index 病人合并路径_IX_当前阶段ID On 病人合并路径(当前阶段ID) Tablespace zl9Indexhis;
Create Index 病人合并路径_IX_前一阶段ID On 病人合并路径(前一阶段ID) Tablespace zl9Indexhis;
Create Index 病人合并路径_IX_首要路径阶段ID On 病人合并路径(首要路径阶段ID) Tablespace zl9Indexhis;
Create Index 病人合并路径_IX_首要路径记录ID On 病人合并路径(首要路径记录ID) Tablespace zl9Indexhis;
Create Index 病人合并路径_IX_待转出 On 病人合并路径(待转出) Tablespace zl9Indexcis;
Create Index 病人合并路径评估_IX_待转出 On 病人合并路径评估(待转出) Tablespace zl9Indexcis;

Create Index 病人路径执行_IX_合并路径阶段ID On 病人路径执行(合并路径阶段ID) Tablespace zl9Indexhis;
Create Index 病人路径执行_IX_合并路径记录ID On 病人路径执行(合并路径记录ID) Tablespace zl9Indexhis;
Create Index 病人路径指标_IX_合并路径阶段ID On 病人路径指标(合并路径记录ID) Tablespace zl9Indexhis;
Create Index 病人路径指标_IX_日期 On 病人路径指标(日期) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人路径指标_IX_阶段ID On 病人路径指标(阶段ID) Tablespace zl9Indexcis;
Create Index 病人路径指标_IX_待转出 On 病人路径指标(待转出) Tablespace zl9Indexcis;
Create Index 病人出径记录_IX_路径记录ID On 病人出径记录(路径记录ID) Tablespace zl9Indexhis;
Create Index 病人出径记录_IX_待转出 On 病人出径记录(待转出) Tablespace zl9Indexhis;
Create Index 病人路径取消_IX_病人ID On 病人路径取消(病人ID,主页ID) Tablespace zl9Indexcis;

----------------------------------------------------------------------------
--[[18.病历业务]]
----------------------------------------------------------------------------
Create Index 电子病历记录_IX_病人ID On 电子病历记录(病人ID,主页ID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index 电子病历记录_IX_文件ID On 电子病历记录(文件ID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index 电子病历记录_IX_完成时间 On 电子病历记录(完成时间,病历种类,科室id) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index 电子病历记录_IX_创建时间 On 电子病历记录(创建时间) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index 电子病历记录_IX_路径执行ID On 电子病历记录(路径执行ID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index 电子病历记录_IX_待转出 On 电子病历记录(待转出) Initrans 20 Tablespace zl9Indexcis;
Create Index 电子病历记录_IX_门诊路径执行ID On 电子病历记录(门诊路径执行ID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;

Create Index 电子病历内容_IX_父ID On 电子病历内容(父ID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index 电子病历内容_IX_预制提纲ID On 电子病历内容(预制提纲ID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index 电子病历内容_IX_诊治要素ID On 电子病历内容(诊治要素ID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index 电子病历内容_IX_待转出 On 电子病历内容(待转出)  Initrans 20 Tablespace zl9Indexcis;
Create Index 病历变动原因_IX_病历文件id On 病历变动原因(病历文件id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病历变动原因_IX_原因要件id On 病历变动原因(原因要件id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病历变动原因_IX_原因要素 On 病历变动原因(原因要素) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病历变动结果_IX_变动原因id On 病历变动结果(变动原因id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病历变动结果_IX_结果要件id On 病历变动结果(结果要件id) Pctfree 5 Tablespace zl9Indexhis;
Create Index 病历变动结果_IX_结果要素 On 病历变动结果(结果要素) Pctfree 5 Tablespace zl9Indexhis;
Create Index 电子病历打印_IX_病人ID On 电子病历打印(病人ID,主页ID) Pctfree 5 Initrans 20 Tablespace zl9Indexcis;
Create Index 电子病历时机_IX_病人ID On 电子病历时机(病人ID,主页ID) Pctfree 20 Tablespace zl9Indexcis;
Create Index 电子病历时机_IX_文件ID On 电子病历时机(文件ID) Pctfree 20 Tablespace zl9Indexcis;
Create Index 疾病申报记录_IX_文档ID On 疾病申报记录(文档ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 疾病申报记录_IX_待转出 On 疾病申报记录(待转出) Tablespace zl9Indexcis;
Create Index 疾病申报记录_IX_姓名 On 疾病申报记录(姓名) Tablespace zl9Indexcis;
Create Index 疾病申报记录_IX_病人ID On 疾病申报记录(病人ID,主页ID) Tablespace zl9Indexcis;

Create Index 电子病历附件_IX_待转出 On 电子病历附件(待转出) Initrans 20 Tablespace zl9Indexcis;
Create Index 电子病历格式_IX_待转出 On 电子病历格式(待转出) Initrans 20 Tablespace zl9Indexcis;
Create Index 电子病历图形_IX_待转出 On 电子病历图形(待转出) Initrans 20 Tablespace zl9Indexcis;

--临时表,不要指定表空间,Pctfree等参数
Create Index 病历时限监测_IX_病人id On 病历时限监测(病人ID,主页ID,病人来源);
Create Index 病历内容监测_IX_病人id On 病历内容监测(病人ID,主页ID,病人来源);

--病案审查归档
Create Index 病案提交记录_IX_主页ID On 病案提交记录(病人ID,主页ID) Tablespace zl9Indexcis;
Create Index 病案提交记录_IX_提交时间 On 病案提交记录(提交时间) Tablespace zl9Indexcis;
Create Index 病案打印记录_IX_主页ID On 病案打印记录(病人id,主页id) Tablespace zl9Indexcis;
Create Index 病案打印记录_IX_打印时间 On 病案打印记录(打印时间) Tablespace zl9Indexcis;
Create Index 病案审阅书签_IX_提交id On 病案审阅书签(提交id) Tablespace zl9Indexcis;
Create Index 病案反馈记录_IX_主页ID On 病案反馈记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈记录_IX_提交id On 病案反馈记录(提交id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈记录_IX_相关id On 病案反馈记录(相关id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈记录_IX_反馈时间 On 病案反馈记录(反馈时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈记录_IX_处理时间 On 病案反馈记录(处理时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈记录_IX_医嘱id On 病案反馈记录(医嘱id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈记录_IX_科室id On 病案反馈记录(科室id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈历史_IX_主页ID On 病案反馈历史(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈历史_IX_医嘱id On 病案反馈历史(医嘱id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈历史_IX_科室id On 病案反馈历史(科室id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈历史_IX_提交id On 病案反馈历史(提交id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈历史_IX_相关id On 病案反馈历史(相关id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈历史_IX_反馈时间 On 病案反馈历史(反馈时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案反馈历史_IX_处理时间 On 病案反馈历史(处理时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病案封存记录_IX_主页ID On 病案封存记录(病人ID,主页ID) Tablespace zl9Indexcis;
Create Index 病案借阅内容_IX_主页ID On 病案借阅内容(病人ID,主页ID) Tablespace zl9Indexcis;
Create Index 病案借阅内容_IX_借阅id On 病案借阅内容(借阅id) Tablespace zl9Indexcis;
Create Index 病案借阅内容_IX_病人id On 病案借阅内容(病人id) Tablespace zl9Indexcis;
Create Index 病案借阅人员_IX_借阅id On 病案借阅人员(借阅id) Tablespace zl9Indexcis;
Create Index 病案评分标准_IX_方案ID On 病案评分标准(方案ID) Tablespace zl9Indexcis;
Create Index 病案评分标准_IX_上级ID On 病案评分标准(上级ID) Tablespace zl9Indexcis;
Create Index 病案评分结果_IX_方案ID On 病案评分结果(方案ID) Tablespace zl9Indexcis;
Create Index 病案评分明细_IX_结果ID On 病案评分明细(主表ID) Tablespace zl9Indexcis;
Create Index 病案评分明细_IX_评分标准ID On 病案评分明细(评分标准ID) Tablespace zl9Indexcis;
Create Index 病案借阅记录_IX_登记时间 On 病案借阅记录(登记时间) Tablespace zl9Indexcis;


----------------------------------------------------------------------------
--[[19.护理业务]]
----------------------------------------------------------------------------
Create Index 病人护理文件_IX_主页ID On 病人护理文件(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理文件_IX_待转出 On 病人护理文件(待转出) Tablespace zl9Indexcis;
Create index 病人护理文件_IX_续打ID On 病人护理文件(续打ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理记录_IX_待转出 On 病人护理记录(待转出) Tablespace zl9Indexcis;
Create Index 病人护理记录_IX_主页ID On 病人护理记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理记录_IX_发生时间 On 病人护理记录(发生时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理内容_IX_待转出 On 病人护理内容(待转出) Tablespace zl9Indexcis;
Create Index 病人护理内容_IX_记录id On 病人护理内容(记录id) Pctfree 5 Tablespace zl9Indexcis;

Create Index 病人护理数据_IX_文件ID On 病人护理数据(文件ID,发生时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理数据_IX_待转出 On 病人护理数据(待转出) Tablespace zl9Indexcis;
Create Index 病人护理明细_IX_记录ID On 病人护理明细(记录ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理明细_IX_来源ID On 病人护理明细(来源ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理明细_IX_待转出 On 病人护理明细(待转出) Tablespace zl9Indexcis;

Create Index 病人护理打印_IX_文件ID On 病人护理打印(文件ID,发生时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理打印_IX_待转出 On 病人护理打印(待转出) Tablespace zl9Indexcis;
Create Index 病区标记记录_IX_主页ID On 病区标记记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理活动项目_IX_待转出 On 病人护理活动项目(待转出) Tablespace zl9Indexcis;
Create Index 产程要素内容_IX_待转出 On 产程要素内容(待转出) Tablespace zl9Indexcis;
Create Index 病人护理要素内容_IX_待转出 On 病人护理要素内容(待转出) Tablespace zl9Indexcis;
Create Index 病人护理诊断_IX_病人ID On 病人护理诊断(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理诊断_IX_文件ID On 病人护理诊断(文件ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病人护理诊断_IX_待转出 ON 病人护理诊断 (待转出) Tablespace zl9Indexcis;


----------------------------------------------------------------------------

--[[20.检验业务]]

----------------------------------------------------------------------------
Create Index 检验流水线标本_IX_待转出 On 检验流水线标本(待转出) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验流水线指标_IX_待转出 On 检验流水线指标(待转出) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验流水线指标_IX_项目ID On 检验流水线指标(项目ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验流水线标本_IX_标本ID On 检验流水线标本(标本ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验流水线指标_IX_标本ID On 检验流水线指标(标本ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_医嘱ID On 检验标本记录(医嘱id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_检验时间 On 检验标本记录(检验时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_申请时间 On 检验标本记录(申请时间) Pctfree 5 Tablespace ZL9INDEXCIS;
Create Index 检验标本记录_IX_审核时间 On 检验标本记录(审核时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_年龄数字 On 检验标本记录(年龄数字) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_挂号单 On 检验标本记录(挂号单) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_合并ID On 检验标本记录(合并ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_主页ID On 检验标本记录(病人ID,主页ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_标识号 On 检验标本记录(标识号) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_待转出 On 检验标本记录(待转出) Tablespace zl9Indexcis;
Create Index 检验标本记录_IX_NO On 检验标本记录(NO) Tablespace zl9Indexcis;

Create Index 检验普通结果_IX_细菌ID On 检验普通结果(细菌ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验普通结果_IX_仪器ID On 检验普通结果(仪器ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验普通结果_IX_检验标本ID On 检验普通结果(检验标本ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验普通结果_IX_药敏组ID On 检验普通结果(药敏组ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验普通结果_IX_酶标板ID On 检验普通结果(酶标板ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验普通结果_IX_项目id On 检验普通结果(检验项目ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验普通结果_IX_待转出 On 检验普通结果(待转出) Tablespace zl9Indexcis;
Create Index 检验项目分布_IX_标本id On 检验项目分布(标本id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验项目分布_IX_项目id On 检验项目分布(项目id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验项目分布_IX_医嘱id On 检验项目分布(医嘱id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验项目分布_IX_细菌ID On 检验项目分布(细菌id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验项目分布_IX_待转出 On 检验项目分布(待转出) Tablespace zl9Indexcis;

Create Index 检验质控记录_IX_仪器ID On 检验质控记录(仪器ID) Tablespace zl9Indexcis;
Create Index 检验质控记录_IX_质控品ID On 检验质控记录(质控品ID) Tablespace zl9Indexcis;
Create Index 检验质控记录_IX_待转出 On 检验质控记录(待转出) Tablespace zl9Indexcis;
Create Index 检验图像结果_IX_标本id On 检验图像结果(标本id) Pctfree 5 Tablespace zl9Indexcis;
Create Index 检验酶标记录_IX_测试时间 On 检验酶标记录(测试时间) Tablespace zl9Indexhis;
Create Index 检验操作记录_IX_标本id On 检验操作记录(标本id) Tablespace zl9Indexcis;
Create Index 检验操作记录_IX_待转出 On 检验操作记录(待转出) Tablespace zl9Indexhis;
Create Index 检验分析记录_IX_标本ID On 检验分析记录(标本ID) Tablespace zl9Indexhis;
Create Index 检验分析记录_IX_用途 On 检验分析记录(用途) Tablespace zl9Indexhis;
Create Index 检验分析记录_IX_待转出 On 检验分析记录(待转出) Tablespace zl9Indexcis;
Create Index 检验拒收记录_IX_医嘱ID On 检验拒收记录(医嘱ID) Tablespace zl9Indexcis;
Create Index 检验拒收记录_IX_待转出 On 检验拒收记录(待转出) Tablespace zl9Indexcis;

Create Index 检验申请项目_IX_待转出 On 检验申请项目(待转出) Tablespace zl9Indexcis;
Create Index 检验试剂记录_IX_待转出 On 检验试剂记录(待转出) Tablespace zl9Indexcis;
Create Index 检验质控报告_IX_待转出 On 检验质控报告(待转出) Tablespace zl9Indexcis;
Create Index 检验药敏结果_IX_待转出 On 检验药敏结果(待转出) Tablespace zl9Indexcis;
Create Index 检验签名记录_IX_待转出 On 检验签名记录(待转出) Tablespace zl9Indexhis;

----------------------------------------------------------------------------
--[[21.检查业务]]
----------------------------------------------------------------------------
Create Index 影像报告驳回_IX_医嘱ID On 影像报告驳回(医嘱ID,病历ID,检查报告ID,RISID,报告ID) Tablespace ZL9INDEXCIS;
Create Index 影像报告驳回_IX_待转出 On 影像报告驳回(待转出) Tablespace zl9Indexcis;
Create Index 影像报告驳回_IX_检查报告ID On 影像报告驳回(检查报告ID) Tablespace zlPacsBaseIndex;
Create Index 影像检查记录_IX_检查号 On 影像检查记录(检查号, 影像类别) Tablespace zl9Indexcis;
Create Index 影像检查记录_IX_位置一 On 影像检查记录(位置一) Tablespace zl9Indexcis;
Create Index 影像检查记录_IX_位置二 On 影像检查记录(位置二) Tablespace zl9Indexcis;
Create Index 影像检查记录_IX_位置三 On 影像检查记录(位置三) Tablespace zl9Indexcis;
Create Index 影像检查记录_IX_接收日期 On 影像检查记录(接收日期) Pctfree 5 Tablespace zl9Indexcis;
Create Index 影像检查记录_Ix_执行科室id On 影像检查记录(执行科室id) Pctfree 5 Tablespace Zl9Indexcis;
Create Index 影像检查记录_IX_关联ID On 影像检查记录(关联ID) Pctfree 5 Tablespace Zl9Indexcis;
Create Index 影像检查记录_IX_待转出 On 影像检查记录(待转出) Tablespace zl9Indexcis;
Create Index 影像检查记录_IX_校对状态 On 影像检查记录(校对状态)  Tablespace zl9Indexcis;

Create Index 影像临时记录_IX_检查号 On 影像临时记录(检查号, 影像类别) Tablespace zl9Indexcis;
Create Index 影像临时记录_IX_位置一 On 影像临时记录(位置一) Tablespace zl9Indexcis;
Create Index 影像临时记录_IX_位置二 On 影像临时记录(位置二) Tablespace zl9Indexcis;
Create Index 影像临时记录_IX_位置三 On 影像临时记录(位置三) Tablespace zl9Indexcis;
Create Index 影像临时记录_IX_接收日期 On 影像临时记录(接收日期) Tablespace zl9Indexcis;
Create Index 胶片打印记录_IX_相关ID On 胶片打印记录(相关ID) Tablespace zl9Indexhis;
Create Index 胶片打印记录_IX_打印时间 On 胶片打印记录(打印时间) Tablespace zl9Indexhis;
Create Index 影像收藏类别_IX_上级ID On 影像收藏类别(上级ID) Tablespace zl9Indexcis;
Create Index 影像收藏类别_IX_创建人ID On 影像收藏类别(创建人ID) Tablespace zl9Indexcis;
Create Index 影像申请单图像_IX_医嘱ID On 影像申请单图像(医嘱ID) Tablespace zl9Indexcis;
Create Index 影像申请单图像_IX_待转出 On 影像申请单图像(待转出) Tablespace zl9Indexcis;
Create Index 影像收藏内容_IX_医嘱ID On 影像收藏内容(医嘱ID) Tablespace zl9Indexcis;
Create Index 影像收藏内容_IX_待转出 On 影像收藏内容(待转出) Tablespace zl9Indexcis;

Create Index 影像检查图象_IX_待转出 On 影像检查图象(待转出) Tablespace zl9Indexcis;
Create Index 影像检查序列_IX_待转出 On 影像检查序列(待转出) Tablespace zl9Indexcis;
Create Index 影像危急值记录_IX_待转出 On 影像危急值记录(待转出) Tablespace zl9Indexcis;

Create Index 影像预约记录_IX_预约设备ID On 影像预约记录(预约设备ID) Tablespace zl9Indexcis;
Create Index 影像预约记录_IX_预约开始时间 On 影像预约记录(预约开始时间) Tablespace zl9Indexcis;
Create Index 影像预约记录_IX_医嘱ID On 影像预约记录(医嘱ID) Tablespace zl9Indexcis;
Create Index 影像预约记录_IX_待转出 On 影像预约记录(待转出) Tablespace zl9Indexcis;
Create Index 影像预约项目_IX_预约设备ID On 影像预约项目(预约设备ID) Tablespace zl9Indexcis;
Create Index 影像预约项目_IX_诊疗项目ID On 影像预约项目(诊疗项目ID) Tablespace zl9Indexcis;
Create Index 影像预约方案_IX_预约设备ID On 影像预约方案(预约设备ID) Tablespace zl9Indexcis;
Create Index 影像预约时间计划_IX_预约方案ID On 影像预约时间计划(预约方案ID) Tablespace zl9Indexcis;

Create Index 病理检查信息_IX_号码规则ID On 病理检查信息(号码规则ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理检查信息_IX_医嘱ID On 病理检查信息(医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理检查信息_IX_报到时间 On 病理检查信息(报到时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理质量信息_IX_病理医嘱ID On 病理质量信息(病理医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理标本信息_IX_医嘱ID On 病理标本信息(医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理标本信息_IX_送检ID On 病理标本信息(送检ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理送检信息_IX_医嘱ID On 病理送检信息(医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理申请信息_IX_病理医嘱ID On 病理申请信息(病理医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理申请信息_IX_申请时间 On 病理申请信息(申请时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理取材信息_IX_病理医嘱ID On 病理取材信息(病理医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理取材信息_IX_申请ID On 病理取材信息(申请ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理取材信息_IX_标本ID On 病理取材信息(标本ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理取材信息_IX_取材时间 On 病理取材信息(取材时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理脱钙信息_IX_标本ID On 病理脱钙信息(标本ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理制片信息_IX_材块ID On 病理制片信息(材块ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理制片信息_IX_申请ID On 病理制片信息(申请ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理制片信息_IX_病理医嘱ID On 病理制片信息(病理医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理制片信息_IX_制片时间 On 病理制片信息(制片时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理过程报告_IX_病理医嘱ID On 病理过程报告(病理医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理特检信息_IX_申请ID On 病理特检信息(申请ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理特检信息_IX_抗体ID On 病理特检信息(抗体ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理特检信息_IX_材块ID On 病理特检信息(材块ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理特检信息_IX_完成时间 On 病理特检信息(完成时间) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理报告延迟_IX_病理医嘱ID On 病理报告延迟(病理医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理会诊信息_IX_病理医嘱ID On 病理会诊信息(病理医嘱ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理抗体反馈_IX_抗体ID On 病理抗体反馈(抗体ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理套餐关联_IX_抗体ID On 病理套餐关联(抗体ID) Pctfree 5 Tablespace zl9Indexcis;
Create Index 病理档案信息_IX_分类ID On 病理档案信息(分类ID) Tablespace zl9Indexcis;
Create Index 病理档案信息_IX_创建日期 On 病理档案信息(创建日期) Tablespace zl9Indexcis;
Create Index 病理归档信息_IX_材块ID On 病理归档信息(材块ID) Pctfree 5 TableSpace zl9Indexcis;
Create Index 病理归档信息_IX_制片ID On 病理归档信息(制片ID) Pctfree 5 TableSpace zl9Indexcis;
Create Index 病理归档信息_IX_特检ID On 病理归档信息(特检ID) Pctfree 5 TableSpace zl9Indexcis;
Create Index 病理归档信息_IX_档案ID On 病理归档信息(档案ID) Pctfree 5 TableSpace zl9Indexcis;
Create Index 病理借阅信息_IX_借阅时间 On 病理借阅信息(借阅时间) TableSpace zl9Indexcis;
Create Index 病理借阅信息_IX_证件号码 On 病理借阅信息(证件号码,证件类型) TableSpace zl9Indexcis;
Create Index 病理遗失信息_IX_借阅ID On 病理遗失信息(借阅ID) TableSpace zl9Indexcis;
Create Index 病理遗失信息_IX_归档ID On 病理遗失信息(归档ID) TableSpace zl9Indexcis;
Create Index 病理遗失信息_IX_遗失日期 On 病理遗失信息(遗失日期) TableSpace zl9Indexcis;
Create Index 病理归还信息_IX_借阅ID On 病理归还信息(借阅ID) TableSpace zl9Indexcis;
Create Index 病理借阅关联_IX_借阅ID On 病理借阅关联(借阅ID) TableSpace zl9Indexcis;
Create Index 病理玻片信息_IX_材块Id On 病理玻片信息(材块Id) Tablespace zl9Indexcis;
Create Index 病理玻片信息_IX_来源ID On 病理玻片信息(来源ID) Tablespace zl9Indexcis;
Create Index 病理玻片信息_IX_病理医嘱ID On 病理玻片信息(病理医嘱ID) Tablespace zl9Indexcis;


Create Index 影像报告值域清单_IX_分类ID On 影像报告值域清单(分类ID) Tablespace zlPacsBaseIndex;
Create Index 影像报告元素清单_IX_分类ID On 影像报告元素清单(分类ID) Tablespace zlPacsBaseIndex;
Create Index 影像报告元素清单_IX_值域ID On 影像报告元素清单(值域ID) Tablespace zlPacsBaseIndex;
Create Index 影像报告片段清单_IX_上级ID On 影像报告片段清单(上级ID) Tablespace zlPacsBaseIndex;
Create Index 影像报告动作_IX_原型ID On 影像报告动作(原型ID) Tablespace zlPacsBaseIndex;
Create Index 影像报告动作_IX_事件ID On 影像报告动作(事件ID) Tablespace zlPacsBaseIndex;
Create Index 影像报告记录_IX_原型ID On 影像报告记录(原型ID) Tablespace zlPacsBizIndex;
Create Index 影像报告记录_IX_待转出 On 影像报告记录(待转出) Tablespace zlPacsBizIndex;
Create Index 影像报告记录_IX_医嘱ID On 影像报告记录(医嘱ID) Tablespace zlPacsBizIndex;
Create Index 影像参数说明_IX_PID On 影像参数说明(PID) Tablespace zlPacsBaseIndex;
Create Index 影像报告操作记录_IX_报告ID On 影像报告操作记录(报告ID) Tablespace zlPacsBaseIndex;
Create Index 影像报告操作记录_IX_待转出 On 影像报告操作记录(待转出) Tablespace zlPacsBaseIndex;
Create Index 影像报告操作记录_IX_医嘱ID On 影像报告操作记录(医嘱ID) Tablespace zlPacsBaseIndex;

Create Index 影像查询方案_IX_所属模块 On 影像查询方案(所属模块) Tablespace zl9Indexhis;
Create Index 影像查询关联_IX_用户ID On 影像查询关联(用户ID) Tablespace zl9Indexhis;
Create Index 影像查询关联_IX_查询方案ID On 影像查询关联(查询方案ID) Tablespace zl9Indexhis;
Create Index 影像查询特性_IX_用户ID On 影像查询特性(用户ID) Tablespace zl9Indexhis;
Create Index 影像查询特性_IX_查询方案ID On 影像查询特性(查询方案ID) Tablespace zl9Indexhis;