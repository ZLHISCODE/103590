

INSERT INTO 电子票据类别 (编号,名称,简码,是否启用,部件,包名称) values(2,'中联模拟接口(V1.0.0)','BSDZBJPT',1,'zlEInvoice.clsEInvoice_Test','b_Einvoice_Request_Test');
 
insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('中联模拟接口(V1.0.0)',1,upper('URL_Type'),'HTTP',NULL);
insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('中联模拟接口(V1.0.0)',2,upper('URL_Address'),'','<ip>:<port>/<service>/api/medical/接口服务标识');
insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('中联模拟接口(V1.0.0)',3,'应用帐号','','即Appid');
insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('中联模拟接口(V1.0.0)',4,'签名私钥','','即KEY值');
insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('中联模拟接口(V1.0.0)',5,'支持版本','V2.0.3','目前只支持:V2.0.3和V3.1.0');
insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('中联模拟接口(V1.0.0)',6,'数据传输方式','application/json','提交和返回数据可以为JSON格式（Content-Type: application/json）');
insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('中联模拟接口(V1.0.0)',7,'字符编码','UTF-8','统一采用UTF-8字符编码');

insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('中联模拟接口(V1.0.0)',8,'缺省卡类别ID','','缺省读取的卡类别');
insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('中联模拟接口(V1.0.0)',9,'身份证作卡类型编号','999998','使用身份证作为上传的卡类型的编号');
insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('中联模拟接口(V1.0.0)',10,'病人无卡的卡类别编号','999999','病人无任何卡时上传的卡类型编号');
insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('中联模拟接口(V1.0.0)',11,'病人无卡的卡号','-','病人无任何卡时上传的卡号');

--insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('中联模拟接口(V1.0.0)',12,'误差项目','','因小数与HIS不匹配时，产生误差费的收费项目');



CREATE TABLE 收费渠道对照(
	结算方式 varchar2(20),
	卡类别ID Number(18),
	渠道编码 varchar2(20))
TABLESPACE zl9Expense;
Alter Table 收费渠道对照 Add Constraint 收费渠道对照_UQ_开票点ID  Unique(结算方式,卡类别ID, 渠道编码)  Using Index Tablespace zl9Indexhis;
Alter Table 收费渠道对照 Add Constraint 收费渠道对照_FK_卡类别ID Foreign Key (卡类别ID) References 医疗卡类别(ID) on delete cascade;
 

--用于统计用
CREATE TABLE 开票结算对照(
	结算方式 varchar2(20), --HIS结算方式
	开票结算方式 varchar2(20)) --只有三个:个人账户支付,医保统筹基金支付,其它医保支付
TABLESPACE zl9Expense;
Alter Table 开票结算对照 Add Constraint 开票结算对照_UQ_结算方式 Unique(结算方式,开票结算方式 )  Using Index Tablespace zl9Indexhis; 


--药品支付类别对码
CREATE TABLE 支付类别对照(
	保险大类ID number(18),  
	大类编码 varchar2(20),
	大类名称 varchar2(50)) --1：无自负/甲;2：有自负/乙;3：全自负/丙
TABLESPACE zl9Expense;
Alter Table 支付类别对照 Add Constraint 支付类别对照_UQ_保险大类ID Unique(保险大类ID,大类编码)  Using Index Tablespace zl9Indexhis; 
Alter Table 支付类别对照 Add Constraint 支付类别对照FK_保险大类ID Foreign Key (保险大类ID) References 保险支付大类(ID) on delete cascade;


Create Table 电子票据核对记录(
 业务日期 Date,
 开票点 Varchar2(100),
 His开票数 Number(18),
 His开票金额 Number(18),
 平台开票数 Number(16, 5),
 平台开票金额 Number(16, 5),
 核对类型 Number(1),
 核对人 Varchar2(50),
 核对时间 Date,
 核对结果 Number(1),
 核对说明 Varchar2(4000)
) Tablespace Zl9Expense;

Alter Table 电子票据核对记录 Add Constraint 电子票据核对记录_Uq_业务日期 Unique(业务日期, 开票点, 核对类型, 核对人) Using Index Tablespace Zl9indexhis;
Alter Table 电子票据核对记录 Modify 业务日期 Constraint 电子票据核对记录_NN_业务日期 Not Null;
Alter Table 电子票据核对记录 Modify 核对类型 Constraint 电子票据核对记录_NN_核对类型 Not Null;
Alter Table 电子票据核对记录 Modify 核对人 Constraint 电子票据核对记录_NN_核对人 Not Null;
Alter Table 电子票据核对记录 Modify 核对时间 Constraint 电子票据核对记录_NN_核对时间 Not Null;
Alter Table 电子票据核对记录 Modify 核对结果 Constraint 电子票据核对记录_NN_核对结果 Not Null;
 
Create Table 电子票据修正记录(
 业务日期 Date,
 电子票据ID Number(18),
 业务流水号 Varchar2(50),
 HIS开票点 Varchar2(100),
 HIS开票金额 Number(16,5),
 HIS票据状态 Number(1),
 平台开票点 Varchar2(100),
 平台开票金额 Number(16,5),
 平台票据状态 Number(1),
 修正方式 NUmber(1),
 修正人 Varchar2(50),
 修正时间 Date,
 修正结果 Number(1),
 修正说明 Varchar2(4000)
) Tablespace Zl9Expense;

Alter Table 电子票据修正记录 Add Constraint 电子票据修正记录_Uq_业务日期 Unique(业务日期, 电子票据ID, 业务流水号, 平台票据状态) Using Index Tablespace Zl9indexhis;
Alter Table 电子票据修正记录 Modify 业务日期 Constraint 电子票据修正记录_NN_业务日期 Not Null;
Alter Table 电子票据修正记录 Modify 修正方式 Constraint 电子票据修正记录_NN_修正方式 Not Null;
Alter Table 电子票据修正记录 Modify 修正人 Constraint 电子票据修正记录_NN_修正人 Not Null;
Alter Table 电子票据修正记录 Modify 修正时间 Constraint 电子票据修正记录_NN_修正时间 Not Null;
Alter Table 电子票据修正记录 Modify 修正结果 Constraint 电子票据修正记录_NN_修正结果 Not Null;

CONNECT sys@system AS sysdba;
grant execute on dbms_crypto to zlhis

CONNECT zlhis@his ;

Create Or Replace Package b_Einvoice_Request_Test Is
  ------------------------------------------------------------------
  --电子票据业务处理(主要是针对博思电子票据的处理)
  --一.公共函数说明
  --  0.get_charset-获取当前数据库的字符集
  --  1.zlJsonStr-处理要组合成Json串的字符串中的特殊符号
  --  2.MD5_Clob-针对Clob进行MD5加密
  --  3.Request-发起业务请求   
  --  3.GetRequestData_Encode-获取请求数据(加密)
  --  4.Getrequestdata_Decode-获取业务数据(解密)
  --  5.Get_ParaInfo-获取博思相关参数配置
  --  6.Get_Version-获取当前博思的版本号
  --  7 Get_IDENTITY-获取操作员信息
  --二.电子票据处理
  --1.获取业务数据相关
  --1.1 Get_Einvoice_Node-获取开票点函数
  --1.2 Get_Chargedata_Create-获取收费开票票据
  --1.3 Get_SendCarddata_Create-获取发卡开票票据
  --1.4 Get_Registerdata_Create-获取挂号开票票据
  --1.5 Get_MZBalancedata_Create-获取门诊结帐开票票据
  --1.6 Get_ZYBalancedata_Create-获取住院结帐开票票据
  --1.7 Get_Depositdata_Create-获取预交开票票据
  --2.电子票据操作相关
  --2.1.Einvoice_Start-电子票据是否启用
  --2.2.EInvoice_Create-电子票据开具
  --2.3.Einvoice_Cancel_Check-电子票据作废前检查(返回:1-合法;0-不合法)
  --2.4.Einvoice_Cancel-电子票据作废(返回1-成功;0-失败)
  ------------------------------------------------------------------
  Function Get_Charset Return Number;

  Function zlJsonStr
  (
    Str_In  In Varchar2,
    Type_In Number := 0
  ) Return Varchar2;

  Procedure Get_Identity
  (
    人员id_Out     Out 人员表.Id%Type,
    人员编号_Out   Out 人员表.编号%Type,
    操作员姓名_Out Out 人员表.姓名%Type
  );

  Function Md5_Clob(Souce_In In Clob) Return Varchar2;

  Function Getrequestdata_Encode
  (
    Reqdata_In Clob,
    Appid_In   Varchar2,
    Key_In     Varchar2,
    Version_In Varchar2 := '1.0'
  ) Return Clob;

  Function Getrequestdata_Decode
  (
    Datasouce_In   Clob,
    Appid_In       Varchar2,
    Key_In         Varchar2,
    Datadecode_Out Out Clob,
    Errmsg_Out     Out Varchar2,
    Version_In     Varchar2 := '1.0'
  ) Return Number;

  Procedure Get_Parainfo
  (
    Version_Out     Out Varchar2,
    Url_Out         Out Varchar2,
    Url_Type_Out    Out Varchar2,
    Appid_Out       Out Varchar2,
    Key_Out         Out Varchar2,
    Contenttype_Out Out Varchar2,
    Charset_Out     Out Varchar2
  );
  Function Get_Version Return Varchar2;

  Function Request
  (
    Reqdata_In    Clob,
    Servername_In Varchar2,
    Respdata_Out  Out Clob,
    Errmsg_Out    Out Varchar2,
    Version_In    Varchar2 := '1.0'
  ) Return Number;

  Function Get_Einvoice_Node
  (
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    人员id_In     人员表.Id%Type := Null
    
  ) Return Varchar2;

  Function Einvoice_Start
  (
    业务场景_In Integer,
    险类_In     保险结算记录.险类%Type,
    类型_In     Integer := Null
  ) Return Number;

  Procedure Get_Chargedata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  );

  Procedure Get_Sendcarddata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  );

  Procedure Get_Registerdata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  );

  Procedure Get_Zybalancedata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  );

  Procedure Get_Mzbalancedata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  );

  Procedure Get_Depositdata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  );

  Function Einvoice_Create
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    冲销id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number;

  --电子票据作废检查
  Function Einvoice_Cancel_Check
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number;

  --电子票据作废
  Function Einvoice_Cancel
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number;

End b_Einvoice_Request_Test;
/


Create Or Replace Package Body b_Einvoice_Request_Test Is
  ------------------------------------------------------------------
  --电子票据业务处理(主要是针对博思电子票据的处理)
  --一.公共函数说明
  --  0.get_charset-获取当前数据库的字符集
  --  1.zlJsonStr-处理要组合成Json串的字符串中的特殊符号
  --  2.MD5_Clob-针对Clob进行MD5加密
  --  3.Request-发起业务请求   
  --  3.GetRequestData_Encode-获取请求数据(加密)
  --  4.Getrequestdata_Decode-获取业务数据(解密)
  --  5.Get_ParaInfo-获取博思相关参数配置
  --  6.Get_Version-获取当前博思的版本号
  --  7 Get_IDENTITY-获取操作员信息
  --二.电子票据处理
  --1.获取业务数据相关
  --1.1 Get_Einvoice_Node-获取开票点函数
  --1.2 Get_Chargedata_Create-获取收费开票票据
  --1.3 Get_SendCarddata_Create-获取发卡开票票据
  --1.4 Get_Registerdata_Create-获取挂号开票票据
  --1.5 Get_MZBalancedata_Create-获取门诊结帐开票票据
  --1.6 Get_ZYBalancedata_Create-获取住院结帐开票票据
  --1.7 Get_Depositdata_Create-获取预交开票票据
  --2.电子票据操作相关
  --2.1.Einvoice_Start-电子票据是否启用
  --2.2.EInvoice_Create-电子票据开具
  --2.3.Einvoice_Cancel_Check-电子票据作废前检查(返回:1-合法;0-不合法)
  --2.4.Einvoice_Cancel-电子票据作废(返回1-成功;0-失败)
  ------------------------------------------------------------------
  Mv_开票点     电子票据使用记录.开票点%Type;
  Mv_人员id     人员表.Id%Type;
  Mv_操作员编号 人员表.编号%Type;
  Mv_操作员姓名 人员表.姓名%Type;
  Mv_Url_Type   三方接口配置.参数值%Type;

  Mv_Url         三方接口配置.参数值%Type;
  Mv_Appid       三方接口配置.参数值%Type;
  Mv_Key         三方接口配置.参数值%Type;
  Mv_Version     三方接口配置.参数值%Type;
  Mv_Contenttype 三方接口配置.参数值%Type;
  Mv_Charset     三方接口配置.参数值%Type;
  Mn_Charset     Number(2); --1.ZHS16GBK;0-AL32UTF8

  Function Get_Charset Return Number As
    --功能：获取当前数据库的字符集
    --返回:1.ZHS16GBK;0-AL32UTF8
    n_Charset Number(2);
  Begin
    If Mn_Charset Is Not Null Then
      Return Mn_Charset;
    End If;
    Begin
      Select Nvl(Max(1), 0)
      Into n_Charset
      From Nls_Database_Parameters
      Where Parameter = 'NLS_CHARACTERSET' And Value Like '%ZHS16GBK%';
    Exception
      When Others Then
        n_Charset := 1;
    End;
    Mn_Charset := n_Charset;
  End Get_Charset;

  Function zlJsonStr
  (
    Str_In  In Varchar2,
    Type_In Number := 0
  ) Return Varchar2 Is
    --功能：处理要组合成Json串的字符串中的特殊符号
    --参数：
    --     Type_In=数据类型:0-字符,1-数值
    v_Temp Varchar2(32767);
  Begin
    If Str_In Is Null Then
      If Nvl(Type_In, 0) = 1 Then
        Return '0';
      Else
        Return Null;
      End If;
    Elsif Nvl(Type_In, 0) = 1 Then
      If Instr(Str_In, '.') = 0 Then
        Return Str_In;
      Else
        Return To_Char(Str_In, 'FM9999999990.0999999999');
      End If;
    Else
      --注意顺序
      v_Temp := Str_In;
      v_Temp := Replace(v_Temp, '\', '\\');
      v_Temp := Replace(v_Temp, '"', '\"');
      v_Temp := Replace(v_Temp, Chr(13), '\r');
      v_Temp := Replace(v_Temp, Chr(10), '\n');
      v_Temp := Replace(v_Temp, Chr(9), '\t');
      Return v_Temp;
    End If;
  End zlJsonStr;
  Procedure Get_Identity
  (
    人员id_Out     Out 人员表.Id%Type,
    人员编号_Out   Out 人员表.编号%Type,
    操作员姓名_Out Out 人员表.姓名%Type
  ) Is
  Begin
    If Mv_操作员编号 Is Not Null Then
      人员id_Out     := Mv_人员id;
      人员编号_Out   := Mv_操作员编号;
      操作员姓名_Out := Mv_操作员姓名;
      Return;
    End If;
    Select p.Id, p.编号, p.姓名
    Into 人员id_Out, 人员编号_Out, 操作员姓名_Out
    From 上机人员表 O, 人员表 P
    Where o.用户名 = User And o.人员id = p.Id And Rownum < 2;
  
    Mv_人员id     := 人员id_Out;
    Mv_操作员编号 := 人员编号_Out;
    Mv_操作员姓名 := 操作员姓名_Out;
  Exception
    When Others Then
      Null;
  End Get_Identity;

  Function Md5_Clob(Souce_In In Clob) Return Varchar2 Is
  Begin
    If Souce_In Is Null Then
      Return Null;
    End If;
    Return Rawtohex(Dbms_Crypto.Hash(Souce_In, Dbms_Crypto.Hash_Md5));
  End Md5_Clob;

  Function Getrequestdata_Encode
  (
    Reqdata_In Clob,
    Appid_In   Varchar2,
    Key_In     Varchar2,
    Version_In Varchar2 := '1.0'
  ) Return Clob As
    ------------------------------------------------------------------
    --业务数据加密处理
    --入参:
    --   ReqData_In-请求的数据
    --   Appid_In-应用帐号
    --   Key_In-签名私钥
    --   version_In-请求的版本号，缺省为:1.0
    --返回:
    --    加密后的数据
    ------------------------------------------------------------------
  
    v_Guid    Varchar2(1000);
    v_Sign    Varchar2(32676);
    c_Data    Clob;
    c_Temp    Clob;
    n_Charset Number(2);
  Begin
  
    v_Guid := Sys_Guid();
  
    n_Charset := Get_Charset;
  
    --base64编码
    c_Data := Replace(Zltools.Zlbase64.Decode(Reqdata_In, n_Charset), Chr(13), '');
  
    --stringA=appid+app1+data=ewogI&noise=ibuaiVcKdpRxkhJA+key=192006250b4c09247ec02f6a2d+version=1.0
    c_Temp := To_Clob('appid=' || Appid_In || 'data=') || c_Data ||
              To_Clob('noise=' || v_Guid || 'key=' || Key_In || 'version=' || Version_In);
  
    --MD5加密
    v_Sign := Upper(Md5_Clob(c_Temp));
  
    --业务参数构建
    c_Temp := To_Clob('{');
    c_Temp := c_Temp || To_Clob('"appid":"' || Appid_In || '",'); --应用帐号
    c_Temp := c_Temp || To_Clob('"data":"' || c_Data || '",'); --请求业务参数
    c_Temp := c_Temp || To_Clob('"noise":"' || v_Guid || '",'); --请求随机标识
    c_Temp := c_Temp || To_Clob('"version":"' || Version_In || '",'); --版本
    c_Temp := c_Temp || To_Clob('"sign":"' || v_Sign || '"'); --签名
    c_Temp := c_Temp || To_Clob('}');
  
    Return c_Temp;
  End Getrequestdata_Encode;

  Function Getrequestdata_Decode
  (
    Datasouce_In   Clob,
    Appid_In       Varchar2,
    Key_In         Varchar2,
    Datadecode_Out Out Clob,
    Errmsg_Out     Out Varchar2,
    Version_In     Varchar2 := '1.0'
  ) Return Number As
    ------------------------------------------------------------------
    --发起业务请求(解密)
    --入参:
    --   dataSouce_In-需要解码的原数据
  
    --出参:
    --    DataDecode_Out-返回解码后的数据
    --    ErrMsg_out-返回的错误内容
    ------------------------------------------------------------------
    j_Json PLJson;
    c_Data       Clob;
    v_Temp       Varchar2(32767);
    c_Temp       Clob;
    v_Result     Varchar2(1000);
    n_Charset    Number(2);
    v_Guid       Varchar2(32767);
    v_Sign       Varchar2(100);
    v_Sign_Check Varchar2(100);
  Begin
    j_Json := PLJson(Datasouce_In);
  
    --  返回结果参数  data  String  是  各API调用返回的内容不同，实际内容以各接口API为准，参数值为JSON格式做base64编码，编码字符集UTF-8 
    --  返回随机标识  noise  String  是  每次请求返回一个唯一编号，全局唯一（建议采用UUID）。
    --  签名  sign  String  是  MD5摘要结果转换成大写
    c_Data := j_Json.Get_Clob('data');
    v_Guid := j_Json.Get_String('noise');
    v_Sign := j_Json.Get_String('sign');
  
    --stringA=appid+app1+data=ewogI&noise=ibuaiVcKdpRxkhJA+key=192006250b4c09247ec02f6a2d+version=1.0
    c_Temp := To_Clob('appid=' || Appid_In || 'data=') || c_Data ||
              To_Clob('noise=' || v_Guid || 'key=' || Key_In || 'version=' || Version_In);
    --MD5加密
    v_Sign_Check := Upper(b_Einvoice_Request_Test.Md5_Clob(c_Temp));
    If Nvl(v_Sign_Check, '') <> Nvl(v_Sign, '') Then
      --不一致，可能被串改，返回0
      Errmsg_Out := '签名信息不正确，请与产品提供商联系！';
      Return 0;
    End If;
  
    n_Charset := b_Einvoice_Request_Test.Get_Charset;
    c_Data    := Zltools.Zlbase64.Decode(c_Data, n_Charset); --base64解码
  
    j_Json   := PLJson(c_Data);
    v_Result := j_Json.Get_String('result'); --除S0000表示成功标识，其它都为错误标识
    If Nvl(v_Result, '') <> 'S0000' Then
      --{"result":"E0001","message":"BASE64(错误信息)"}。
      v_Temp         := j_Json.Get_String('message');
      v_Temp         := Zltools.Zlbase64.Decode(v_Temp, n_Charset);
      Errmsg_Out     := v_Result || '-' || v_Temp;
      Datadecode_Out := Null;
      Return 0;
    End If;
  
    --请求成功:
    --{"result":"S0000","message":"BASE64(响应业务参数)"}
    Datadecode_Out := j_Json.Get_Clob('message');
    --base64解码
    Datadecode_Out := Zltools.Zlbase64.Decode(Datadecode_Out, n_Charset);
    Return 1;
  Exception
    When Others Then
      Errmsg_Out := 'JSON内容错误：' || To_Char(SQLCode) || ':' || SQLErrM;
      Return 0;
  End Getrequestdata_Decode;

  Procedure Get_Parainfo
  (
    Version_Out     Out Varchar2,
    Url_Out         Out Varchar2,
    Url_Type_Out    Out Varchar2,
    Appid_Out       Out Varchar2,
    Key_Out         Out Varchar2,
    Contenttype_Out Out Varchar2,
    Charset_Out     Out Varchar2
  ) Is
  Begin
    If Mv_Url Is Null Then
    
      For c_配置 In (Select 参数号, 参数名, 参数值 From 三方接口配置 Where 接口名 = '中联模拟接口(V1.0.0)') Loop
        If c_配置.参数名 = Upper('URL_Type') Then
          Mv_Url_Type := c_配置.参数值;
        Elsif c_配置.参数名 = Upper('URL_Address') Then
          Mv_Url := c_配置.参数值;
        Elsif c_配置.参数名 = Upper('应用帐号') Then
          Mv_Appid := c_配置.参数值;
        Elsif c_配置.参数名 = Upper('签名私钥') Then
          Mv_Key := c_配置.参数值;
        Elsif c_配置.参数名 = Upper('支持版本') Then
          Mv_Version := c_配置.参数值;
        Elsif c_配置.参数名 = Upper('数据传输方式') Then
          Mv_Contenttype := c_配置.参数值;
        Elsif c_配置.参数名 = Upper('字符编码') Then
          Mv_Charset := c_配置.参数值;
        End If;
      End Loop;
    
      If Mv_Charset Is Null Then
        Mv_Charset := 'UTF-8';
      End If;
      If Mv_Url_Type Is Null Then
        Mv_Url_Type := 'HTTP';
      End If;
      Mv_Url := Mv_Url_Type || '://' || Mv_Url;
      If Mv_Contenttype Is Null Then
        Mv_Contenttype := 'application/json';
      End If;
    End If;
    Version_Out     := Mv_Version;
    Url_Out         := Mv_Url;
    Url_Type_Out    := Mv_Url_Type;
    Appid_Out       := Mv_Appid;
    Key_Out         := Mv_Key;
    Contenttype_Out := Mv_Contenttype;
    Charset_Out     := Mv_Charset;
  End Get_Parainfo;

  Function Get_Version Return Varchar2 As
    v_Version     三方接口配置.参数值%Type;
    v_Url         三方接口配置.参数值%Type;
    v_Url_Type    三方接口配置.参数值%Type;
    v_Appid       三方接口配置.参数值%Type;
    v_Key         三方接口配置.参数值%Type;
    v_Contenttype 三方接口配置.参数值%Type;
    v_Charset     三方接口配置.参数值%Type;
  Begin
    Get_Parainfo(v_Version, v_Url, v_Url_Type, v_Appid, v_Key, v_Contenttype, v_Charset);
    Return v_Version;
  End Get_Version;

  Function Request
  (
    Reqdata_In    Clob,
    Servername_In Varchar2,
    Respdata_Out  Out Clob,
    Errmsg_Out    Out Varchar2,
    Version_In    Varchar2 := '1.0'
  ) Return Number As
    ------------------------------------------------------------------
    --发起业务请求
    --入参:
    --   ReqData_In-发送的报文内容
    --   ServerName_in-请求服务名
    --   version_In-请求的版本号，缺省为:1.0
    --出参:
    --    RespData_Out-响应报文内容
    --    ErrMsg_out-返回的错误内容
    ------------------------------------------------------------------
  
    o_Http_Req  Utl_Http.Req; --http请求对象
    o_Http_Resp Utl_Http.Resp; --http响应对象
  
    Err_Item Exception;
    v_Err_Msg Varchar2(255);
    c_Temp    Clob; --响应报文
  
    v_Buffer_Text Varchar2(32767); --缓存
    n_Deftimeout  Integer Default 3600;
    --HttP相关变量
    v_Url         Varchar2(4000); --发送的Http
    v_Url_Type    三方接口配置.参数值%Type;
    v_Appid       三方接口配置.参数值%Type;
    v_Key         三方接口配置.参数值%Type;
    v_Version     三方接口配置.参数值%Type;
    v_Charset     三方接口配置.参数值%Type;
    v_Contenttype Varchar2(1000);
    v_Temp        Varchar2(32767);
    n_Amount      Pls_Integer := 3900;
    n_Offset      Pls_Integer := 1;
    n_Count       Number(18);
    c_Data        Clob;
  Begin
  
    Get_Parainfo(v_Version, v_Url, v_Url_Type, v_Appid, v_Key, v_Contenttype, v_Charset);
    --URL：http://[ip]:[port]/[service]/api/medical/ [接口服务标识]
    v_Url     := v_Url || '/' || Servername_In;
    v_Version := Version_In;
  
    --组织业务请求的加密数据
    c_Data := Getrequestdata_Encode(Reqdata_In, v_Appid, v_Key, v_Version);
  
    -- 初始化HTTP请求参数.
    Utl_Http.Set_Transfer_Timeout(n_Deftimeout);
    o_Http_Req := Utl_Http.Begin_Request(v_Url, 'POST');
    Utl_Http.Set_Header(o_Http_Req, 'Content-Type', v_Contenttype);
    Utl_Http.Set_Header(o_Http_Req, 'Content-Length', Lengthb(c_Data));
    Utl_Http.Set_Body_Charset(o_Http_Req, v_Charset);
  
    n_Count := Dbms_Lob.Getlength(c_Data);
  
    If n_Count > 30000 Then
      --分块发送HTTP请求
      Utl_Http.Set_Header(o_Http_Req, 'Transfer-Encoding', 'chunked'); --分块
      While (n_Offset < n_Count) Loop
        Dbms_Lob.Read(c_Data, n_Amount, n_Offset, v_Buffer_Text);
      
        Utl_Http.Write_Text(o_Http_Req, v_Buffer_Text);
        n_Offset := n_Offset + n_Amount;
      End Loop;
    Else
      --发送HTTP请求
      v_Temp := c_Data;
      Utl_Http.Write_Text(o_Http_Req, v_Temp);
    End If;
    --接收响应
    o_Http_Resp := Utl_Http.Get_Response(o_Http_Req);
  
    Begin
      c_Temp := Null;
      v_Temp := Null;
      Loop
        Utl_Http.Read_Text(o_Http_Resp, v_Buffer_Text, 30000);
        If Length(Nvl(v_Temp, '') || v_Buffer_Text) >= 30000 Then
          c_Temp := c_Temp || To_Clob(Nvl(v_Temp, '') || v_Buffer_Text);
          v_Temp := Null;
        End If;
        v_Temp := Nvl(v_Temp, '') || v_Buffer_Text;
      End Loop;
      If v_Temp Is Not Null And c_Temp Is Not Null Then
        c_Temp := c_Temp || To_Clob(Nvl(v_Temp, ''));
        v_Temp := Null;
      End If;
      --关闭HTTP连接
      Utl_Http.End_Response(o_Http_Resp);
    Exception
      When Utl_Http.Request_Failed Then
        v_Err_Msg := 'HTTP请求失败：' || To_Char(SQLCode) || ':' || Substr(SQLErrM, 1, 128);
        Raise Err_Item;
      When Utl_Http.Transfer_Timeout Then
        v_Err_Msg := 'HTTP请求超时失败：' || To_Char(SQLCode) || ':' || Substr(SQLErrM, 1, 128);
        Raise Err_Item;
      When Utl_Http.End_Of_Body Then
        Utl_Http.End_Response(o_Http_Resp);
      When Others Then
        Errmsg_Out := 'HTTP请求错误：' || To_Char(SQLCode) || ':' || Substr(SQLErrM, 1, 128);
        Raise Err_Item;
    End;
  
    --解析数据
    If v_Temp Is Null Then
      If Getrequestdata_Decode(c_Temp, v_Appid, v_Key, Respdata_Out, Errmsg_Out, v_Version) = 1 Then
        Return 1;
      End If;
    Else
      If Getrequestdata_Decode(v_Temp, v_Appid, v_Key, Respdata_Out, Errmsg_Out, v_Version) = 1 Then
        Return 1;
      End If;
    End If;
    Return 0;
    --释放clob
    Dbms_Lob.Freetemporary(c_Temp);
    Return 1;
  Exception
    When Err_Item Then
      Utl_Http.End_Response(o_Http_Resp);
      Dbms_Lob.Freetemporary(c_Temp);
      Errmsg_Out := v_Err_Msg;
      Return 0;
    When Others Then
      Utl_Http.End_Response(o_Http_Resp);
      Dbms_Lob.Freetemporary(c_Temp);
      Errmsg_Out := 'JSON内容错误：' || To_Char(SQLCode) || ':' || Substr(SQLErrM, 1, 128);
      Return 0;
  End Request;

  Function Get_Einvoice_Node
  (
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    人员id_In     人员表.Id%Type := Null
  ) Return Varchar2 Is
    n_Count        Number(18);
    v_机器名       Varchar2(100);
    v_操作员开票点 电子票据开票点.编码%Type;
    v_机器名开票点 电子票据开票点.编码%Type;
    v_操作员编号   人员表.编号%Type;
    v_操作员姓名   人员表.姓名%Type;
    n_人员id       人员表.Id%Type;
  Begin
    --缺省以收费员编号为对照
    If Mv_开票点 Is Not Null Then
      Return Mv_开票点;
    End If;
  
    Select Max(1) Into n_Count From 票据开票点对照 Where Rownum < 2;
    If Nvl(n_Count, 0) = 0 Then
      If 操作员编号_In Is Null Then
        If Nvl(人员id_In, 0) <> 0 Then
          Select Max(编号) Into v_操作员编号 From 人员表 Where ID = 人员id_In;
          If v_操作员编号 Is Not Null Then
            Return v_操作员编号;
          End If;
        End If;
      
        If 操作员姓名_In Is Not Null Then
          Select Max(编号) Into v_操作员编号 From 人员表 Where 姓名 = 操作员姓名_In;
          If v_操作员编号 Is Not Null Then
            Return v_操作员编号;
          End If;
        End If;
        --以登陆的为准
        If Mv_操作员编号 Is Not Null Then
          Return Mv_操作员编号;
        End If;
      
        Get_Identity(n_人员id, v_操作员编号, v_操作员姓名);
        If v_操作员编号 Is Not Null Then
          Return v_操作员编号;
        End If;
      
      End If;
      Return 操作员编号_In;
    End If;
    n_人员id := 人员id_In;
    If 人员id_In = 0 Then
      If 操作员编号_In Is Not Null Or 操作员姓名_In Is Not Null Then
        Select Max(ID)
        Into n_人员id
        From 人员表
        Where 编号 = Nvl(操作员编号_In, 's-') Or 姓名 = Nvl(操作员姓名_In, '--');
      End If;
    End If;
  
    If Nvl(n_人员id, 0) = 0 Then
      Get_Identity(n_人员id, v_操作员编号, v_操作员姓名);
    End If;
    Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  
    For c_开票点 In (Select Nvl(a.人员id, 0) As 人员id, Nvl(a.客户端, '-') As 客户端, a.开票点id, b.编码 As 开票点编码, b.名称
                  From 票据开票点对照 A, 电子票据开票点 B
                  Where a.开票点id = b.Id And Nvl(b.撤档时间, Sysdate + 1) >= Sysdate And (a.人员id = n_人员id Or a.客户端 = v_机器名)
                  Order By 人员id, 客户端) Loop
      If Nvl(c_开票点.人员id, 0) = 0 And c_开票点.客户端 = v_机器名 Then
        v_机器名开票点 := c_开票点.开票点编码;
      End If;
    
      If Nvl(c_开票点.人员id, 0) = n_人员id And c_开票点.客户端 = '-' Then
        v_操作员开票点 := c_开票点.开票点编码;
      End If;
    
      If Nvl(c_开票点.人员id, 0) = n_人员id And c_开票点.客户端 = v_机器名 Then
        Return c_开票点.开票点编码;
      End If;
    End Loop;
    If v_操作员开票点 Is Not Null Then
      Return v_操作员开票点;
    End If;
    Return v_机器名开票点;
  End Get_Einvoice_Node;

  Function Einvoice_Start
  (
    业务场景_In Integer,
    险类_In     保险结算记录.险类%Type,
    类型_In     Integer := Null
  ) Return Number Is
    ------------------------------------------------------------------
    --功能:判断电子票据是否启用
    --入参:业务场景_In-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    --     类型_In：null-表示不区分;结帐和预交:1-门诊;2-住院 
    --出参:错误信息_Out-返回的错误信息
    --返回:1-启用;0-未启用
    -------------------------------------------------------------------
    n_Return Number(2);
  Begin
    n_Return := Zl_Fun_Isstarteinvoice(业务场景_In, 险类_In, 类型_In);
    Return n_Return;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Start;

  Procedure Get_Chargedata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  ) Is
    --
    ---------------------------------------------------------------------------
    --功能:获取收费开票数据
    --入参:
    --    Json_In,格式如下
    --  input
    --    occasion N 1  应用场合:1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    --    einvoice_id  N,1 当前电子票据ID
    --    balance_id N 1  结算ID  occasion=2时，预交id;occasion<>2表示结帐id
    --    writeoff_id  N 1  冲销ID  occasion=2时，冲销预交id;occasion<>2表示冲销id
    --出参:
    --  ReqData_Out-返回的业务请求数据
    --  Totalmoney_Out-票据总额
    --  Code_Out-获取是否成功：0-失败；1-成功
    --  Message_Out 错误信息
    ---------------------------------------------------------------------------
    j_Input PLJson;
    j_Json  PLJson;
  
    n_应用场合 Number(2);
    n_结算id   病人预交记录.结帐id%Type;
    --n_冲销id   病人预交记录.结帐id%Type;
  
    v_业务标识   Varchar2(20);
    v_业务流水号 Varchar2(50);
  
    v_开票点       Varchar2(100);
    v_缴费         Varchar2(32767);
    v_票据信息     Varchar2(32767);
    v_就诊信息     Varchar2(32767);
    v_通知         Varchar2(32767);
    v_缴费渠道     Varchar2(32767);
    v_费用         Varchar2(32767);
    v_其它扩展信息 Varchar2(32767);
    v_其它医保信息 Varchar2(32767);
    c_明细         Clob;
    v_明细         Varchar2(32767);
    c_分类明细     Clob;
    v_分类明细     Varchar2(32767);
    c_交易信息     Clob; --最终返回的交易信息集
  
    n_门诊号       病人信息.门诊号%Type;
    n_病人id       病人预交记录.病人id%Type;
    v_患者姓名     门诊费用记录.姓名%Type;
    v_患者性别     门诊费用记录.性别%Type;
    v_患者年龄     门诊费用记录.年龄%Type;
    d_业务发生时间 门诊费用记录.登记时间%Type;
    v_收费员       门诊费用记录.操作员姓名%Type;
  
    n_缺省卡类别id     Number(18);
    v_参数值           Varchar2(100);
    n_票据总金额       门诊费用记录.结帐金额%Type;
    n_误差总额         门诊费用记录.结帐金额%Type;
    n_用户id           人员表.Id%Type;
    v_操作员编号       人员表.编号%Type;
    v_操作员姓名       人员表.姓名%Type;
    v_Temp             Varchar2(32767);
    v_医疗付款方式编码 医疗付款方式.编码%Type;
    v_医疗付款方式名称 医疗付款方式.名称%Type;
    n_险类             保险结算记录.险类%Type;
    v_保险机构编码     保险类别.保险机构编码%Type;
    n_医嘱序号         门诊费用记录.医嘱序号%Type;
    n_挂号id           门诊费用记录.挂号id%Type;
    v_病种名称         保险病种.名称%Type;
    v_就诊日期         Varchar2(20);
    v_就诊科室编码     部门表.编码%Type;
    v_就诊科室名称     部门表.名称%Type;
    v_就诊编号         Varchar2(50);
    v_结帐ids          Varchar2(32767);
    v_医保号           保险帐户.医保号%Type;
    l_结帐id           t_NumList := t_NumList();
    v_版本号           Varchar2(30);
    n_电子票据id       电子票据使用记录.Id%Type;
  Begin
    j_Input := PLJson(Json_In);
    j_Json  := j_Input.Get_Pljson('input');
  
    n_应用场合 := Nvl(j_Json.Get_Number('occasion'), 0);
    n_结算id   := j_Json.Get_Number('balance_id');
    --n_冲销id     := Nvl(j_Json.Get_Number('writeoff_id'), 0);
    n_电子票据id := Nvl(j_Json.Get_Number('einvoice_id'), 0);
    If Nvl(n_应用场合, 0) = 0 Then
      Code_Out    := 0;
      Message_Out := '无效的应用场景';
      Return;
    End If;
  
    Select Nvl(Max(参数值), 'V2.0.3')
    Into v_版本号
    From 三方接口配置
    Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '支持版本';
  
    b_Einvoice_Request_Test.Get_Identity(n_用户id, v_操作员编号, v_操作员姓名);
  
    --v_业务标识:01  住院,02  门诊,03  急诊,04  门特,05  体检中心,06  挂号,07  住院预交金,08  体检预交金
    Select Decode(n_应用场合, 1, '02', 2, '07', 3, '01', 4, '06', 5, '02', '02') Into v_业务标识 From Dual;
  
    n_票据总金额   := 0;
    d_业务发生时间 := Null;
    v_结帐ids      := Null;
    c_明细         := Null;
    v_明细         := Null;
    For c_收费细目 In (Select Min(a.Id) As 费用id, a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号) As 序号, a.收费细目id, Max(a.计算单位) As 计算单位,
                          Sum(a.标准单价) As 价格, Avg(Nvl(a.付数, 1) * Nvl(a.数次, 0)) As 数量, Sum(a.应收金额) As 应收金额,
                          Sum(a.实收金额) As 实收金额, Sum(a.结帐金额) As 结帐金额, Sum(a.实收金额) - Sum(a.统筹金额) As 自费金额,
                          Max(t.编码) As 医保项目编码, Max(t.名称) As 医保项目名称, Max(t.统筹比额) As 医保报销比例, Max(a.摘要) As 备注,
                          Max(a.费用类型) As 费用类型, Max(a.操作员编号) As 操作员编号, Max(a.操作员姓名) As 操作员姓名, Max(a.姓名) As 姓名,
                          Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.病人id) As 病人id, Max(a.登记时间) As 登记时间,
                          Max(a.付款方式) As 付款方式编码, Max(a.收据费目) As 收据费目, Max(c.编码) As 收据费目编码, Max(a.医嘱序号) As 医嘱序号,
                          Max(a.挂号id) As 挂号id, Max(d.编码) As 类别编码, Max(d.类别) As 类别名称, Max(b.编码) As 项目编码, Max(b.名称) As 项目名称,
                          Max(b.规格) As 规格, Max(q.药品剂型) As 药品剂型
                   From 门诊费用记录 A, 收费项目目录 B, 收据费目 C, 收费类别 D, 药品规格 M, 药品特性 Q, 诊疗项目目录 J, 保险支付大类 T, 支付类别对照 S
                   Where a.No In (Select Distinct NO From 门诊费用记录 Where 结帐id = n_结算id) And Mod(a.记录性质, 10) = 1 And
                         a.收费类别 = d.编码(+) And a.收费细目id = b.Id And a.收据费目 = c.名称(+) And a.收费细目id = m.药品id(+) And
                         m.药名id = q.药名id(+) And q.药名id = j.Id(+) And a.保险大类id = t.Id(+) And t.性质(+) = 1 And
                         a.保险大类id = s.保险大类id(+)
                   Group By a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号), a.收费细目id, c.编码, c.名称, j.编码, j.名称
                   Order By NO, 序号) Loop
      If v_患者姓名 Is Null Then
        v_患者姓名 := c_收费细目.姓名;
        v_患者性别 := c_收费细目.性别;
        v_患者年龄 := c_收费细目.年龄;
        n_病人id   := c_收费细目.病人id;
      End If;
      If d_业务发生时间 Is Null And Nvl(c_收费细目.记录状态, 0) = 1 Then
        --取原始业务发生时间
        d_业务发生时间 := c_收费细目.登记时间;
        v_收费员       := c_收费细目.操作员姓名;
      End If;
      If v_医疗付款方式编码 Is Null Then
        v_医疗付款方式编码 := c_收费细目.付款方式编码;
      End If;
      If Nvl(n_医嘱序号, 0) = 0 Then
        n_医嘱序号 := c_收费细目.医嘱序号;
      End If;
      If Nvl(n_挂号id, 0) = 0 Then
        n_挂号id := c_收费细目.挂号id;
      End If;
    
      If Instr(Nvl(v_结帐ids, '') || ',', ',' || c_收费细目.结帐id || ',') = 0 Then
        l_结帐id.Extend;
        l_结帐id(l_结帐id.Count) := c_收费细目.结帐id;
      End If;
    
      --listDetailNo  明细流水号  String  60  否  明细流水号
      v_Temp := '{' || '"listDetailNo":"' || b_Einvoice_Request_Test.Zljsonstr(LPad(c_收费细目.费用id, 20, '0')) || '"';
      --chargeCode  收费项目代码  String  50  否  填写业务系统内部编码值，由医疗平台配置对照,如：床位费、检查费
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.收据费目编码) || '"';
      --chargeName  收费项目名称  String  100  否
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.收据费目) || '"';
      --prescribeCode  处方编码  String  60  否
      v_Temp := v_Temp || ',' || '"prescribeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.No) || '"';
      --listTypeCode  药品类别编码  String  50  否  如药品分类编码01，有则填写
      v_Temp := v_Temp || ',' || '"listTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.类别编码) || '"';
      --listTypeName  药品类别名称  String  50  否  如药品分类名称，抗生素类抗感染药物
      v_Temp := v_Temp || ',' || '"listTypeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.类别名称) || '"';
      --code  编码  String  50  否  如药品编码，有则填写
      v_Temp := v_Temp || ',' || '"code":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.项目编码) || '"';
      --name  药品名称  String  50  是  如药品名称，器材名称等
      v_Temp := v_Temp || ',' || '"name":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.项目名称) || '"';
      --form  剂型  String  50  否
      v_Temp := v_Temp || ',' || '"form":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.药品剂型) || '"';
      --specification  规格  String  50  否
      v_Temp := v_Temp || ',' || '"specification":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.规格) || '"';
      --unit  计量单位   String  20  否
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.计算单位) || '"';
      --std  单价  Number  14,6  是
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.价格, 1);
      --number  数量  Number  14,6  是
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.数量, 1);
      --amt  金额  Number  14,6  是
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.实收金额, 1);
      --selfAmt  自费金额  Number  14,6  是  如无金额，填写0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.自费金额, 1);
      --receivableAmt  应收费用  Number  14,6  否
      v_Temp := v_Temp || ',' || '"receivableAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.应收金额, 1);
      --medicalCareType  医保药品分类  String  1  否  1：无自负/甲
      --          2：有自负/乙
      --          3：全自负/丙
      v_Temp := v_Temp || ',' || '"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.医保项目编码) || '"';
      --medCareItemType  医保项目类型  String  100  否
      v_Temp := v_Temp || ',' || '"medCareItemType":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.医保项目名称) || '"';
      --medReimburseRate  医保报销比例  Number  3,2  否
      v_Temp := v_Temp || ',' || '"medReimburseRate":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.医保报销比例, 1);
      --remark  备注  String  200  否
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.备注) || '"';
      --sortNo  序号  Integer  不限  否  序号
      v_Temp := v_Temp || ',' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.序号, 1);
      --chrgtype  费用类型  String  50  否
      v_Temp := v_Temp || ',' || '"chrgtype":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.费用类型) || '"}';
    
      If Length(Nvl(v_明细, '') || v_Temp) > 32700 Then
        If c_明细 Is Null Then
          c_明细 := To_Clob(v_明细);
        Else
          c_明细 := c_明细 || To_Clob(',' || v_明细);
        End If;
        v_明细 := Null;
      End If;
    
      If v_明细 Is Null Then
        v_明细 := v_Temp;
      Else
        v_明细 := v_明细 || ',' || v_Temp;
      End If;
    End Loop;
  
    If v_明细 Is Not Null And c_明细 Is Not Null Then
      --listDetail  清单项目明细  String  不限  是  详见A-2,JSON格式列表
      c_明细 := c_明细 || ',' || To_Clob(v_明细);
      c_明细 := To_Clob(',"listDetail":[') || c_明细 || To_Clob(']');
    
      v_明细 := Null;
    Elsif v_明细 Is Not Null Then
      v_明细 := ',"listDetail":[' || v_明细 || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --分类明细
    v_分类明细 := Null;
    c_分类明细 := Null;
    For c_分类统计 In (Select Rownum As 序号, 收据费目编码, 收据费目名称, 数量, 计算单位, Round(单价, 2) As 单价, Round(结帐金额, 2) As 结帐金额,
                          Round(自费金额, 2) As 自费金额, 备注, 结帐金额 - Round(结帐金额, 2) As 误差费
                   From (Select /*+cardinality(b,10)*/
                           c.编码 As 收据费目编码, a.收据费目 As 收据费目名称, 1 As 数量, '' As 计算单位, Sum(a.结帐金额) As 单价, a.收据费目,
                           Sum(a.结帐金额) As 结帐金额, Sum(a.结帐金额) - Sum(a.统筹金额) As 自费金额, '' As 备注
                          From 门诊费用记录 A, Table(l_结帐id) B, 收据费目 C
                          Where a.结帐id = b.Column_Value And a.收据费目 = c.名称(+)
                          Group By c.编码, a.收据费目)) Loop
      --sortNo  序号  Integer  不限  是  默认从1开始，每个收费项目序号值递增1，本次不允许重复
      v_Temp := '{' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.序号, 1);
      --chargeCode  收费项目代码  String  50  是  填写业务系统内部编码值，由医疗平台配置对照
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.收据费目编码) || '"';
      --chargeName  收费项目名称  String  100  是  填写业务系统内部项目名称
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.收据费目名称) || '"';
      --unit  计量单位  String  20  否
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.计算单位) || '"';
      --std  收费标准  Number  14,2  是
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.单价, 1);
      --number  数量  Number  14,2  是
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.数量, 1);
      --amt  金额  Number  14,2  是
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.结帐金额, 1);
      --selfAmt  自费金额  Number  14,2  是  如无金额，填写0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.自费金额, 1);
      --remark  备注  String  200  否
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.备注) || '"}';
    
      If Length(Nvl(v_分类明细, '') || v_Temp) > 32700 Then
        If c_分类明细 Is Null Then
          c_分类明细 := To_Clob(v_分类明细);
        Else
          c_分类明细 := c_分类明细 || To_Clob(',' || v_分类明细);
        End If;
        v_分类明细 := Null;
      End If;
    
      If v_分类明细 Is Null Then
        v_分类明细 := v_Temp;
      Else
        v_分类明细 := v_分类明细 || ',' || v_Temp;
      End If;
    
      n_票据总金额 := Nvl(n_票据总金额, 0) + Nvl(c_分类统计.结帐金额, 0);
      n_误差总额   := Nvl(n_误差总额, 0) + Nvl(c_分类统计.误差费, 0);
    End Loop;
    Totalmoney_Out := n_票据总金额;
    If v_分类明细 Is Not Null And c_分类明细 Is Not Null Then
      c_分类明细 := c_分类明细 || ',' || To_Clob(v_分类明细);
      ----chargeDetail 收费项目明细  String  不限  是  详见A-1,JSON格式列表
      c_分类明细 := To_Clob(',"chargeDetail":[') || c_分类明细 || To_Clob(']');
      v_分类明细 := Null;
    Elsif v_分类明细 Is Not Null Then
      v_分类明细 := ',"chargeDetail":[' || v_分类明细 || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --票据信息
    --Select Count(1) Into n_作废次数 From 电子票据使用记录 Where 票种 = n_应用场合 And 结算id = n_结算id;
    --结算ID||_||电子票据ID
    v_业务流水号 := n_结算id || '_' || Nvl(n_电子票据id, 0);
    v_开票点     := b_Einvoice_Request_Test.Get_Einvoice_Node(v_操作员编号, v_操作员姓名, n_用户id);
  
    v_票据信息 := '"busNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_业务流水号) || '"'; --业务流水号
    v_票据信息 := v_票据信息 || ',"busType":"' || b_Einvoice_Request_Test.Zljsonstr(v_业务标识) || '"'; --业务标识
    v_票据信息 := v_票据信息 || ',"payer":"' || b_Einvoice_Request_Test.Zljsonstr(v_患者姓名) || '"'; --患者姓名
    v_票据信息 := v_票据信息 || ',"busDateTime":"' ||
              b_Einvoice_Request_Test.Zljsonstr(To_Char(d_业务发生时间, 'yyyymmddHH24miss') || '000') || '"'; --业务发生时间
    v_票据信息 := v_票据信息 || ',"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_开票点) || '"'; --开票点编码:直接填写业务系统内部编码值，由医疗平台配置对照
    v_票据信息 := v_票据信息 || ',"payee":"' || b_Einvoice_Request_Test.Zljsonstr(v_收费员) || '"'; --收费员
  
    v_票据信息 := v_票据信息 || ',"author":"' || b_Einvoice_Request_Test.Zljsonstr(v_操作员姓名) || '"'; --票据编制人
    v_票据信息 := v_票据信息 || ',"totalAmt":' || b_Einvoice_Request_Test.Zljsonstr(n_票据总金额, 1); --开票总金额
    v_票据信息 := v_票据信息 || ',"remark":"' || '' || '"'; --备注
    -------------------------------------------------------------------------------------------
  
    --取缴费信息
    v_缴费 := Null;
    For c_缴费 In (Select Max(Decode(信息名, '订单号', 信息值, '支付订单号', 信息值, '')) As 支付定单号,
                        Max(Decode(信息名, '医保支付订单号', 信息值, '医保订单号', 信息值, '')) As 医保支付定单号,
                        Max(Decode(Upper(信息名), '支付宝公众号USERID', 信息值, '')) As 支付宝公众号userid,
                        Max(Decode(Upper(信息名), '支付宝小程序USERID', 信息值, '')) As 支付宝小程序userid,
                        Max(Decode(Upper(信息名), '微信公众号OPENID', 信息值, '')) As 微信公众号openid,
                        Max(Decode(Upper(信息名), '微信小程序OPENID', 信息值, '')) As 微信小程序openid
                 From (Select 信息名, 信息值
                        From 病人信息从表
                        Where 病人id = n_病人id And 信息名 In ('支付宝公众号USERID', '支付宝小程序USERID', '微信公众号OPENID', '微信小程序OPENID')
                        Union All
                        Select 交易项目, 交易内容
                        From 三方结算交易
                        Where 交易id In (Select ID From 病人预交记录 Where 结帐id = n_结算id) And 交易项目 Like '%订单号')) Loop
      v_缴费 := ',"alipayCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.支付宝公众号userid) || '"'; --患者支付宝账户
      v_缴费 := v_缴费 || ',"weChatOrderNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.支付定单号) || '"'; --微信支付订单号
      If v_版本号 = 'V3.1.0' Then
        --该版本才有此接点
        v_缴费 := v_缴费 || ',"weChatMedTransNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.医保支付定单号) || '"'; --微信医保支付订单号
      End If;
    
      If c_缴费.微信公众号openid Is Not Null Then
        v_缴费 := v_缴费 || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.微信公众号openid) || '"'; --微信公众号或小程序用户ID
      Else
        v_缴费 := v_缴费 || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.微信小程序openid) || '"'; --微信公众号或小程序用户ID
      End If;
      Exit;
    End Loop;
  
    -------------------------------------------------------------------------------------------
    --取通知信息
    Select To_Number(Max(参数值))
    Into n_缺省卡类别id
    From 三方接口配置
    Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '缺省卡类别ID';
    v_通知 := Null;
    For c_通知 In (Select Max(a.病人id) As 病人id, Max(a.姓名) As 姓名, Max(a.手机号) As 手机号, Max(a.Email) As Email, Max(1) As 缴款类型,
                        Max(a.身份证号) As 身份证号, Max(m.名称) As 卡类别, Max(m.卡号) As 卡号, Max(a.门诊号) As 门诊号
                 From 病人信息 A,
                      (
                        
                        Select 病人id, 名称, 编码, 卡号
                        From (Select b.病人id, c.名称, c.编码, b.卡号, Decode(b.卡类别id, n_缺省卡类别id, 2, c.缺省标志) As 缺省标志
                                From 病人医疗卡信息 B, 医疗卡类别 C
                                Where b.卡类别id = c.Id And b.病人id = n_病人id
                                Order By 缺省标志)
                        Where Rownum < 2) M
                 Where a.病人id = m.病人id(+)) Loop
    
      v_通知 := ',"tel":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.手机号) || '"'; --患者手机号码
      v_通知 := v_通知 || ',"email":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.Email) || '"'; --患者邮箱地址
      If v_版本号 = 'V3.1.0' Then
        --该版本才有此接点
        v_通知 := v_通知 || ',"payerType":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.缴款类型) || '"'; --交款人类型
      End If;
      v_通知 := v_通知 || ',"idCardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.身份证号) || '"'; --统一社会信用代码
    
      If c_通知.卡类别 Is Not Null Then
        v_通知 := v_通知 || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.卡类别) || '"'; --卡类型
        v_通知 := v_通知 || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.卡号) || '"'; --卡号
      Elsif c_通知.身份证号 Is Not Null Then
        Select Nvl(Max(参数值), '99998')
        Into v_参数值
        From 三方接口配置
        Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '身份证作卡类型编号';
        v_通知 := v_通知 || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡类型
        v_通知 := v_通知 || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.身份证号) || '"'; --卡号
      Else
        --没有一张卡，固定一种卡类别
        Select Nvl(Max(参数值), '99999')
        Into v_参数值
        From 三方接口配置
        Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '病人无卡的卡类别编号';
        v_通知 := v_通知 || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡类型
        Select Nvl(Max(参数值), '-')
        Into v_参数值
        From 三方接口配置
        Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '病人无卡的卡号';
        v_通知 := v_通知 || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡号
      End If;
      If Nvl(n_门诊号, 0) = 0 Then
        n_门诊号 := c_通知.门诊号;
      
      End If;
    
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --就诊信息
    Select Max(内容) Into v_Temp From zlRegInfo Where 项目 = '医疗机构类型';
  
    --性质:1-收费;2-结算（包括住院结算、特殊门诊结算）；3-预交
    Select Max(a.险类), Max(b.保险机构编码), Max(Nvl(a.病种名称, c.名称))
    Into n_险类, v_保险机构编码, v_病种名称
    From 保险结算记录 A, 保险类别 B, 保险病种 C
    Where a.险类 = b.序号 And a.病种id = c.Id(+) And a.记录id = n_结算id And a.性质 = Decode(n_应用场合, 2, 3, 3, 2, 1);
  
    Select Max(名称) Into v_医疗付款方式名称 From 医疗付款方式 Where 编码 = v_医疗付款方式编码;
    If Nvl(n_险类, 0) <> 0 Then
      Select Max(医保号) Into v_医保号 From 保险帐户 Where 病人id = n_病人id And 险类 = n_险类;
    End If;
  
    v_就诊编号 := Null;
    If Nvl(n_医嘱序号, 0) <> 0 Then
      Select Max(To_Char(a.发生时间, 'yyyy-mm-dd')), Max(b.编码), Max(b.名称), Max(a.No)
      Into v_就诊日期, v_就诊科室编码, v_就诊科室名称, v_就诊编号
      From 病人挂号记录 A, 部门表 B
      Where a.执行部门id = b.Id And
            a.No = (Select Max(挂号单) From 病人医嘱记录 Where ID = n_医嘱序号 Or 相关id = n_医嘱序号);
    Elsif Nvl(n_挂号id, 0) <> 0 Then
      Select Max(To_Char(a.发生时间, 'yyyy-mm-dd')), Max(b.编码), Max(b.名称), Max(a.No)
      Into v_就诊日期, v_就诊科室编码, v_就诊科室名称, v_就诊编号
      From 病人挂号记录 A, 部门表 B
      Where a.执行部门id = b.Id And a.Id = n_挂号id;
    End If;
    If v_就诊编号 Is Null Then
      --取最近一次挂号ID
      Select Max(a.Id), Max(To_Char(a.发生时间, 'yyyy-mm-dd')), Max(b.编码), Max(b.名称), Max(a.No)
      Into n_挂号id, v_就诊日期, v_就诊科室编码, v_就诊科室名称, v_就诊编号
      From 病人挂号记录 A, 部门表 B
      Where a.执行部门id = b.Id And
            a.Id = (Select ID
                    From (Select ID, 发生时间 From 病人挂号记录 Where 病人id = n_病人id Order By 发生时间 Desc)
                    Where Rownum < 2);
    End If;
  
    If v_病种名称 Is Null And Nvl(n_险类, 0) <> 0 Then
    
      Select Max(病种名称)
      Into v_病种名称
      From (
             
             Select Distinct a.名称 As 病种名称
             From 保险病种 A, 保险特准项目 B
             Where a.险类 = n_险类 And a.Id = b.病种id And
                   b.收费细目id In (Select Distinct 收费细目id From 门诊费用记录 Where 结帐id = n_结算id)
             Union All
             Select Distinct a.名称 As 病种名称
             From 保险病种 A, 保险特准项目 B
             Where a.险类 = n_险类 And a.Id = b.病种id And
                   b.大类 In (Select Distinct 保险大类id From 门诊费用记录 Where 结帐id = n_结算id))
      Where Rownum < 2;
    End If;
    --medicalInstitution  医疗机构类型  String  60  否  按照《医疗机构管理条例实施细则》和《卫生部关于修订<医疗机构管理条例实施细则>第三条有关内容的通知》确定的医疗卫生机构类别
    v_就诊信息 := ',"medicalInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_Temp) || '"';
    --medCareInstitution  医保机构编码  String  60  否  医保机构的唯一编码
    v_就诊信息 := v_就诊信息 || ',"medCareInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_保险机构编码) || '"';
    --medCareTypeCode  医保类型编码  String  60  否  
    v_就诊信息 := v_就诊信息 || ',"medCareTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_医疗付款方式编码) || '"';
    --medicalCareType  医保类型名称  String  60  否  取值范围包括职工基本医疗保险、城乡居民基本医疗保险（城镇居民基本医疗保险、新型农村合作医疗保险）和其他医疗保险、非医保等
    v_就诊信息 := v_就诊信息 || ',"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(v_医疗付款方式名称) || '"';
    --medicalInsuranceID  患者医保编号  String  60  否  参保人在医保系统中的唯一标识(医保号)
    v_就诊信息 := v_就诊信息 || ',"medicalInsuranceID":"' || b_Einvoice_Request_Test.Zljsonstr(v_医保号) || '"';
    --consultationDate  就诊日期  String  10  否  患者就医时间
    --          格式:yyyy-MM-dd
    v_就诊信息 := v_就诊信息 || ',"consultationDate":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊日期) || '"';
    --category  就诊科室  String  200  否  
    v_就诊信息 := v_就诊信息 || ',"category":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊科室名称) || '"';
    --patientCategoryCode  就诊科室编码  String  60  否  
    v_就诊信息 := v_就诊信息 || ',"patientCategoryCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊科室编码) || '"';
    --patientNo  患者就诊编号  String  20  是  患者每次就诊一次就生成的一个新的编号。（患者登记号）
    v_就诊信息 := v_就诊信息 || ',"patientNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊编号) || '"';
    --patientId  患者唯一ID  String  50  否  患者在业务系统中的唯一标识ID，类似身份证号码。
    v_就诊信息 := v_就诊信息 || ',"patientId":"' || b_Einvoice_Request_Test.Zljsonstr(n_病人id) || '"';
    --sex  性别  String  2  是  
    v_就诊信息 := v_就诊信息 || ',"sex":"' || b_Einvoice_Request_Test.Zljsonstr(v_患者性别) || '"';
    --age  年龄  String  10  是  
    v_就诊信息 := v_就诊信息 || ',"age":"' || b_Einvoice_Request_Test.Zljsonstr(v_患者年龄) || '"';
    --caseNumber  病历号  String  50  否  
    v_就诊信息 := v_就诊信息 || ',"caseNumber":"' || b_Einvoice_Request_Test.Zljsonstr(n_门诊号) || '"';
    --specialDiseasesName  特殊病种名称  String  200  否  
    v_就诊信息 := v_就诊信息 || ',"specialDiseasesName":"' || b_Einvoice_Request_Test.Zljsonstr(v_病种名称) || '"';
    -------------------------------------------------------------------------------------------
    --结算信息
    v_费用 := Null;
    For c_结算 In (Select 现金预交, 支票预交, 转账预交, 个人帐户支付, 医保统筹基金支付, 其它医保支付, 个人现金支付, Decode(Sign(现金支付), -1, 现金支付, 0) As 现金退款,
                        Decode(Sign(现金支付), -1, 支票支付, 0) As 支票退款, Decode(Sign(现金支付), -1, 转帐支付, 0) As 转帐退款,
                        Decode(Sign(现金支付), -1, 0, 现金支付) As 现金支付, Decode(Sign(现金支付), -1, 0, 支票支付) As 支票支付,
                        Decode(Sign(现金支付), -1, 0, 转帐支付) As 转帐支付,
                        Nvl(个人帐户支付, 0) + Nvl(医保统筹基金支付, 0) + Nvl(其它医保支付, 0) As 报销总额,
                        Nvl(结算总额, 0) - Nvl(个人帐户支付, 0) - Nvl(医保统筹基金支付, 0) - Nvl(其它医保支付, 0) As 自费金额, 结算总额, 医保结算号码,
                        0 As 个人帐户余额
                 From (Select /*+cardinality(b,10)*/
                         Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '现金', 1, 0), 0) * a.冲预交) As 现金预交,
                         Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '支票', 1, 0), 0) * a.冲预交) As 支票预交,
                         Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '支票', 0, '现金', 0, 1), 0) * a.冲预交) As 转账预交,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '个人账户支付', 1, 0)) * a.冲预交) As 个人帐户支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '医保统筹基金支付', 1, 0)) * a.冲预交) As 医保统筹基金支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '其它医保支付', 1, 0)) * a.冲预交) As 其它医保支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '其它医保支付', 0, '个人账户支付', 0, '医保统筹基金支付', 0, 1)) *
                              a.冲预交) As 个人现金支付,
                         Max(Decode(Mod(a.记录性质, 10), 1, 0,
                                     Decode(c.开票结算方式, '其它医保支付', 结算号码, '个人账户支付', 结算号码, '医保统筹基金支付', 结算号码, ''))) As 医保结算号码,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(a.结算方式, '现金', 1, 0)) * a.冲预交) As 现金支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(a.结算方式, '支票', 1, 0)) * a.冲预交) As 支票支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(a.结算方式, '现金', 0, '支票', 0, 1) * a.冲预交)) As 转帐支付,
                         Sum(冲预交) As 结算总额
                        From 病人预交记录 A, Table(l_结帐id) B, 开票结算对照 C
                        Where a.结帐id = b.Column_Value And a.结算方式 = c.结算方式(+)))
    
     Loop
      --accountPay  个人账户支付  Number  14,2  是  按政策规定用个人账户支付参保人的医疗费用（含基本医疗保险目录范围内和目录范围外的费用）；
      --          如无金额，填写0
      v_费用 := ',"accountPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_结算.个人帐户支付, 0), 1);
      --fundPay  医保统筹基金支付  Number  14,2  是  患者本次就医所发生的医疗费用中按规定由基本医疗保险统筹基金支付的金额；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"fundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_结算.医保统筹基金支付, 0), 1);
      --otherfundPay  其它医保支付  Number  14,2  是  患者本次就医所发生的医疗费用中按规定由大病保险、医疗救助、公务员医疗补助、大额补充、企业补充等基金或资金支付的金额；
      v_费用 := v_费用 || ',"otherfundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_结算.其它医保支付, 0), 1);
      --ownPay  自费金额  Number  14,2  是  患者本次就医所发生的医疗费用中按照有关规定不属于基本医疗保险目录范围而全部由个人支付的费用；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"ownPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.自费金额, 1);
      --selfConceitedAmt  个人自负  Number  14,2  是  医保患者起付标准内个人支付费用；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"selfConceitedAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfPayAmt  个人自付  Number  14,2  是  患者本次就医所发生的医疗费用中由个人负担的属于基本医疗保险目录范围内自付部分的金额；开展按病种、病组、床日等打包付费方式且由患者定额付费的费用。该项为个人所得税大病医疗专项附加扣除信；息项如无金额，填写0
      v_费用 := v_费用 || ',"selfPayAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfCashPay  个人现金支付  Number  14,2  是  个人通过现金、银行卡、微信、支付宝等渠道支付的金额；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"selfCashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.个人现金支付, 1);
      --cashPay  现金预交款金额  Number  14,2  否
      v_费用 := v_费用 || ',"cashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.现金预交, 1);
      --chequePay  支票预交款金额  Number  14,2  否
      v_费用 := v_费用 || ',"chequePay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.支票预交, 1);
      --transferAccountPay  转账预交款金额  Number  14,2  否
      v_费用 := v_费用 || ',"transferAccountPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.转账预交, 1);
      --cashRecharge  补交金额(现金)  Number  14,2  否
      v_费用 := v_费用 || ',"cashRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.现金支付, 1);
      --chequeRecharge  补交金额(支票)  Number  14,2  否
      v_费用 := v_费用 || ',"chequeRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.支票支付, 1);
      --transferRecharge  补交金额（转账）  Number  14,2  否
      v_费用 := v_费用 || ',"transferRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.转帐支付, 1);
      --cashRefund  退还金额(现金)  Number  14,2  否
      v_费用 := v_费用 || ',"cashRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.现金退款, 1);
      --chequeRefund  退交金额(支票)  Number  14,2  否
      v_费用 := v_费用 || ',"chequeRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.支票退款, 1);
      --transferRefund  退交金额(转账)  Number  14,2  否
      v_费用 := v_费用 || ',"transferRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.转帐退款, 1);
      --ownAcBalance  个人账户余额  Number  14,2  否
      v_费用 := v_费用 || ',"ownAcBalance":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.个人帐户余额, 1);
      --reimbursementAmt  报销总金额  Number  14,2  否  医保结算后返回的总金额
      v_费用 := v_费用 || ',"reimbursementAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.报销总额, 1);
      --balancedNumber  结算号  String  100  否  医保结算后生成的号码/入账唯一值
      v_费用 := v_费用 || ',"balancedNumber":"' || b_Einvoice_Request_Test.Zljsonstr(c_结算.医保结算号码) || '"';
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --交费渠道
    v_缴费渠道 := Null;
    For c_渠道 In (Select /*+cardinality(b,10)*/
                  Nvl(c.渠道编码, Nvl(d.渠道编码, '-')) As 渠道编码, Sum(冲预交) As 结算总额
                 From 病人预交记录 A, Table(l_结帐id) B, 收费渠道对照 C,
                      (Select 结算方式, 渠道编码 From 收费渠道对照 D Where 卡类别id Is Null) D
                 Where a.结帐id = b.Column_Value And a.卡类别id = c.卡类别id(+) And a.结算方式 = c.结算方式(+) And a.结算方式 = d.结算方式(+)
                 Group By Nvl(c.渠道编码, Nvl(d.渠道编码, '-'))
                 Order By 渠道编码)
    
     Loop
      --payChannelCode  交费渠道编码  String  10  是
      If v_缴费渠道 Is Null Then
        v_缴费渠道 := Nvl(v_缴费渠道, '') || '{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_渠道.渠道编码, 0)) || '"';
      Else
        v_缴费渠道 := v_缴费渠道 || ',{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_渠道.渠道编码, 0)) || '"';
      End If;
      --payChannelValue  交费渠道金额  Number  14,2  是
      v_缴费渠道 := v_缴费渠道 || ',"payChannelValue":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_渠道.结算总额, 0), 1) || '}';
    End Loop;
  
    If v_缴费渠道 Is Not Null Then
      --payChannelDetail  交费渠道列表  String  不限  是  直接填写业务系统内部编码值，由医疗平台配置对照如：附录4交费渠道列表
      --        详见A-5,JSON格式列表
      v_缴费渠道 := ',"payChannelDetail":[' || v_缴费渠道 || ']';
    Else
      v_缴费渠道 := ',"payChannelDetail":[]';
    End If;
  
    -------------------------------------------------------------------------------------------
    --其他医保信息
    v_其它医保信息 := Null;
    --otherMedicalList  其它医保信息列表  String  不限  否  填写其它未知医保信息（在电子票据上以内容拼接方式显示）
    --            详见A-4,JSON格式列表
    --  infoNo  序号  Integer  不限  是  默认从1开始，每项数据序号值递增1，本次不允许重复
    --  infoName  医保信息名称  String  100  是  如费用报销类型编码，可参考附录7医保报销类型列表
    --  infoValue  医保信息值  String  100  是  如费用报销金额
    --  infoOther  医保其它信息  String  100  否  如医保报销比例。
  
    -------------------------------------------------------------------------------------------
    --其它扩展信息
    v_其它扩展信息 := Null;
    --otherInfo  其它扩展信息列表  String  不限  否  填写信息需要在电子票据上单独显示的其它扩展信息（未知信息）
    --          详见A-3,JSON格式列表
    --  infoNo  序号  Integer  不限  是  默认从1开始，每项数据序号值递增1，本次不允许重复
    --  infoName  扩展信息名称  String  100  是
    --  infoValue  扩展信息值  String  500  是
  
    c_交易信息 := To_Clob('{' || v_票据信息);
    c_交易信息 := c_交易信息 || To_Clob(v_缴费);
    c_交易信息 := c_交易信息 || To_Clob(v_通知);
    c_交易信息 := c_交易信息 || To_Clob(v_就诊信息);
    c_交易信息 := c_交易信息 || To_Clob(v_费用);
    c_交易信息 := c_交易信息 || To_Clob(v_缴费渠道);
  
    If v_其它扩展信息 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_其它扩展信息);
    End If;
    If v_其它医保信息 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_其它医保信息);
    End If;
    --  eBillRelateNo  业务票据关联号  String  32  否  如一笔业务数据需要开具N张电子票据，则N张电子票对应该值保持一致，用于后期关联查询
    c_交易信息 := c_交易信息 || To_Clob(',"eBillRelateNo":""');
    If v_分类明细 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_分类明细);
    Else
      c_交易信息 := c_交易信息 || c_分类明细;
    End If;
  
    If v_明细 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_明细);
    Else
      c_交易信息 := c_交易信息 || c_明细;
    End If;
    c_交易信息  := c_交易信息 || To_Clob('}');
    Reqdata_Out := c_交易信息;
    Code_Out    := 1;
  Exception
    When Others Then
      Message_Out := SQLCode || ':' || SQLErrM;
      Code_Out    := 0;
  End Get_Chargedata_Create;

  Procedure Get_Sendcarddata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  ) Is
    --
    ---------------------------------------------------------------------------
    --功能:获取发卡开票数据
    --入参:
    --    Json_In,格式如下
    --  input
    --    occasion N 1  应用场合:5-就诊卡
    --    einvoice_id  N,1 当前电子票据ID
    --    balance_id N 1  结算ID  occasion=2时，预交id;occasion<>2表示结帐id
    --    writeoff_id  N 1  冲销ID  occasion=2时，冲销预交id;occasion<>2表示冲销id
    --出参:
    --  ReqData_Out-返回的业务请求数据
    --  Totalmoney_Out-票据总额
    --  Code_Out-获取是否成功：0-失败；1-成功
    --  Message_Out 错误信息
    ---------------------------------------------------------------------------
    j_Input PLJson;
    j_Json  PLJson;
  
    n_应用场合 Number(2);
    n_结算id   病人预交记录.结帐id%Type;
    --n_冲销id   病人预交记录.结帐id%Type;
  
    v_业务标识   Varchar2(20);
    v_业务流水号 Varchar2(50);
  
    v_开票点       Varchar2(100);
    v_缴费         Varchar2(32767);
    v_票据信息     Varchar2(32767);
    v_就诊信息     Varchar2(32767);
    v_通知         Varchar2(32767);
    v_缴费渠道     Varchar2(32767);
    v_费用         Varchar2(32767);
    v_其它扩展信息 Varchar2(32767);
    v_其它医保信息 Varchar2(32767);
    c_明细         Clob;
    v_明细         Varchar2(32767);
    c_分类明细     Clob;
    v_分类明细     Varchar2(32767);
    c_交易信息     Clob; --最终返回的交易信息集
  
    n_门诊号       病人信息.门诊号%Type;
    n_病人id       病人预交记录.病人id%Type;
    v_患者姓名     门诊费用记录.姓名%Type;
    v_患者性别     门诊费用记录.性别%Type;
    v_患者年龄     门诊费用记录.年龄%Type;
    d_业务发生时间 门诊费用记录.登记时间%Type;
    v_收费员       门诊费用记录.操作员姓名%Type;
  
    n_缺省卡类别id     Number(18);
    v_参数值           Varchar2(100);
    n_票据总金额       门诊费用记录.结帐金额%Type;
    n_误差总额         门诊费用记录.结帐金额%Type;
    n_用户id           人员表.Id%Type;
    v_操作员编号       人员表.编号%Type;
    v_操作员姓名       人员表.姓名%Type;
    v_Temp             Varchar2(32767);
    v_医疗付款方式编码 医疗付款方式.编码%Type;
    v_医疗付款方式名称 医疗付款方式.名称%Type;
    n_险类             保险结算记录.险类%Type;
    v_保险机构编码     保险类别.保险机构编码%Type;
    n_医嘱序号         门诊费用记录.医嘱序号%Type;
    n_挂号id           门诊费用记录.挂号id%Type;
    v_病种名称         保险病种.名称%Type;
    v_就诊日期         Varchar2(20);
    v_就诊科室编码     部门表.编码%Type;
    v_就诊科室名称     部门表.名称%Type;
    v_就诊编号         Varchar2(50);
    v_结帐ids          Varchar2(32767);
    v_医保号           保险帐户.医保号%Type;
    l_结帐id           t_NumList := t_NumList();
    v_版本号           Varchar2(30);
    n_电子票据id       电子票据使用记录.Id%Type;
  
  Begin
    j_Input := PLJson(Json_In);
    j_Json  := j_Input.Get_Pljson('input');
  
    n_应用场合 := Nvl(j_Json.Get_Number('occasion'), 0);
    n_结算id   := j_Json.Get_Number('balance_id');
    --n_冲销id     := Nvl(j_Json.Get_Number('writeoff_id'), 0);
    n_电子票据id := Nvl(j_Json.Get_Number('einvoice_id'), 0);
  
    If Nvl(n_应用场合, 0) = 0 Then
      Code_Out    := 0;
      Message_Out := '无效的应用场景';
      Return;
    End If;
  
    Select Nvl(Max(参数值), 'V2.0.3')
    Into v_版本号
    From 三方接口配置
    Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '支持版本';
  
    b_Einvoice_Request_Test.Get_Identity(n_用户id, v_操作员编号, v_操作员姓名);
  
    --v_业务标识:01  住院,02  门诊,03  急诊,04  门特,05  体检中心,06  挂号,07  住院预交金,08  体检预交金
    Select Decode(n_应用场合, 1, '02', 2, '07', 3, '01', 4, '06', 5, '02', '02') Into v_业务标识 From Dual;
  
    n_票据总金额   := 0;
    d_业务发生时间 := Null;
    v_结帐ids      := Null;
    c_明细         := Null;
    v_明细         := Null;
    For c_收费细目 In (Select Min(a.Id) As 费用id, a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号) As 序号, a.收费细目id, Max(a.计算单位) As 计算单位,
                          Sum(a.标准单价) As 价格, Avg(Nvl(a.付数, 1) * Nvl(a.数次, 0)) As 数量, Sum(a.应收金额) As 应收金额,
                          Sum(a.实收金额) As 实收金额, Sum(a.结帐金额) As 结帐金额, Sum(a.实收金额) - Sum(a.统筹金额) As 自费金额,
                          Max(t.编码) As 医保项目编码, Max(t.名称) As 医保项目名称, Max(t.统筹比额) As 医保报销比例, Max(a.摘要) As 备注,
                          Max(a.费用类型) As 费用类型, Max(a.操作员编号) As 操作员编号, Max(a.操作员姓名) As 操作员姓名, Max(a.姓名) As 姓名,
                          Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.病人id) As 病人id, Max(a.登记时间) As 登记时间, Max('') As 付款方式编码,
                          Max(a.收据费目) As 收据费目, Max(c.编码) As 收据费目编码, Max(a.医嘱序号) As 医嘱序号, Max(0) As 挂号id,
                          Max(d.编码) As 类别编码, Max(d.类别) As 类别名称, Max(b.编码) As 项目编码, Max(b.名称) As 项目名称, Max(b.规格) As 规格,
                          Max(q.药品剂型) As 药品剂型
                   From 住院费用记录 A, 收费项目目录 B, 收据费目 C, 收费类别 D, 药品规格 M, 药品特性 Q, 诊疗项目目录 J, 保险支付大类 T, 支付类别对照 S
                   Where a.No In (Select Distinct NO From 住院费用记录 Where 结帐id = n_结算id) And a.记录性质 = 5 And
                         a.收费类别 = d.编码(+) And a.收费细目id = b.Id And a.收据费目 = c.名称(+) And a.收费细目id = m.药品id(+) And
                         m.药名id = q.药名id(+) And q.药名id = j.Id(+) And a.保险大类id = t.Id(+) And t.性质(+) = 1 And
                         a.保险大类id = s.保险大类id(+)
                   Group By a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号), a.收费细目id, c.编码, c.名称, j.编码, j.名称
                   Order By NO, 序号) Loop
      If v_患者姓名 Is Null Then
        v_患者姓名 := c_收费细目.姓名;
        v_患者性别 := c_收费细目.性别;
        v_患者年龄 := c_收费细目.年龄;
        n_病人id   := c_收费细目.病人id;
      End If;
      If d_业务发生时间 Is Null And Nvl(c_收费细目.记录状态, 0) = 1 Then
        --取原始业务发生时间
        d_业务发生时间 := c_收费细目.登记时间;
        v_收费员       := c_收费细目.操作员姓名;
      End If;
      If v_医疗付款方式编码 Is Null Then
        v_医疗付款方式编码 := c_收费细目.付款方式编码;
      End If;
      If Nvl(n_医嘱序号, 0) = 0 Then
        n_医嘱序号 := c_收费细目.医嘱序号;
      End If;
      If Nvl(n_挂号id, 0) = 0 Then
        n_挂号id := c_收费细目.挂号id;
      End If;
    
      If Instr(Nvl(v_结帐ids, '') || ',', ',' || c_收费细目.结帐id || ',') = 0 Then
        l_结帐id.Extend;
        l_结帐id(l_结帐id.Count) := c_收费细目.结帐id;
      End If;
    
      --listDetailNo  明细流水号  String  60  否  明细流水号
      v_Temp := '{' || '"listDetailNo":"' || b_Einvoice_Request_Test.Zljsonstr(LPad(c_收费细目.费用id, 20, '0')) || '"';
      --chargeCode  收费项目代码  String  50  否  填写业务系统内部编码值，由医疗平台配置对照,如：床位费、检查费
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.收据费目编码) || '"';
      --chargeName  收费项目名称  String  100  否
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.收据费目) || '"';
      --prescribeCode  处方编码  String  60  否
      v_Temp := v_Temp || ',' || '"prescribeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.No) || '"';
      --listTypeCode  药品类别编码  String  50  否  如药品分类编码01，有则填写
      v_Temp := v_Temp || ',' || '"listTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.类别编码) || '"';
      --listTypeName  药品类别名称  String  50  否  如药品分类名称，抗生素类抗感染药物
      v_Temp := v_Temp || ',' || '"listTypeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.类别名称) || '"';
      --code  编码  String  50  否  如药品编码，有则填写
      v_Temp := v_Temp || ',' || '"code":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.项目编码) || '"';
      --name  药品名称  String  50  是  如药品名称，器材名称等
      v_Temp := v_Temp || ',' || '"name":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.项目名称) || '"';
      --form  剂型  String  50  否
      v_Temp := v_Temp || ',' || '"form":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.药品剂型) || '"';
      --specification  规格  String  50  否
      v_Temp := v_Temp || ',' || '"specification":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.规格) || '"';
      --unit  计量单位   String  20  否
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.计算单位) || '"';
      --std  单价  Number  14,6  是
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.价格, 1);
      --number  数量  Number  14,6  是
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.数量, 1);
      --amt  金额  Number  14,6  是
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.实收金额, 1);
      --selfAmt  自费金额  Number  14,6  是  如无金额，填写0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.自费金额, 1);
      --receivableAmt  应收费用  Number  14,6  否
      v_Temp := v_Temp || ',' || '"receivableAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.应收金额, 1);
      --medicalCareType  医保药品分类  String  1  否  1：无自负/甲
      --          2：有自负/乙
      --          3：全自负/丙
      v_Temp := v_Temp || ',' || '"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.医保项目编码) || '"';
      --medCareItemType  医保项目类型  String  100  否
      v_Temp := v_Temp || ',' || '"medCareItemType":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.医保项目名称) || '"';
      --medReimburseRate  医保报销比例  Number  3,2  否
      v_Temp := v_Temp || ',' || '"medReimburseRate":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.医保报销比例, 1);
      --remark  备注  String  200  否
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.备注) || '"';
      --sortNo  序号  Integer  不限  否  序号
      v_Temp := v_Temp || ',' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.序号, 1);
      --chrgtype  费用类型  String  50  否
      v_Temp := v_Temp || ',' || '"chrgtype":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.费用类型) || '"}';
    
      If Length(Nvl(v_明细, '') || v_Temp) > 32700 Then
        If c_明细 Is Null Then
          c_明细 := To_Clob(v_明细);
        Else
          c_明细 := c_明细 || To_Clob(',' || v_明细);
        End If;
        v_明细 := Null;
      End If;
    
      If v_明细 Is Null Then
        v_明细 := v_Temp;
      Else
        v_明细 := v_明细 || ',' || v_Temp;
      End If;
    End Loop;
  
    If v_明细 Is Not Null And c_明细 Is Not Null Then
      --listDetail  清单项目明细  String  不限  是  详见A-2,JSON格式列表
      c_明细 := c_明细 || ',' || To_Clob(v_明细);
      c_明细 := To_Clob(',"listDetail":[') || c_明细 || To_Clob(']');
    
      v_明细 := Null;
    Elsif v_明细 Is Not Null Then
      v_明细 := ',"listDetail":[' || v_明细 || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --分类明细
    v_分类明细 := Null;
    c_分类明细 := Null;
    For c_分类统计 In (Select Rownum As 序号, 收据费目编码, 收据费目名称, 数量, 计算单位, Round(单价, 2) As 单价, Round(结帐金额, 2) As 结帐金额,
                          Round(自费金额, 2) As 自费金额, 备注, 结帐金额 - Round(结帐金额, 2) As 误差费
                   From (Select /*+cardinality(b,10)*/
                           c.编码 As 收据费目编码, a.收据费目 As 收据费目名称, 1 As 数量, '' As 计算单位, Sum(a.结帐金额) As 单价, a.收据费目,
                           Sum(a.结帐金额) As 结帐金额, Sum(a.结帐金额) - Sum(a.统筹金额) As 自费金额, '' As 备注
                          From 住院费用记录 A, Table(l_结帐id) B, 收据费目 C
                          Where a.结帐id = b.Column_Value And a.收据费目 = c.名称(+)
                          Group By c.编码, a.收据费目)) Loop
      --sortNo  序号  Integer  不限  是  默认从1开始，每个收费项目序号值递增1，本次不允许重复
      v_Temp := '{' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.序号, 1);
      --chargeCode  收费项目代码  String  50  是  填写业务系统内部编码值，由医疗平台配置对照
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.收据费目编码) || '"';
      --chargeName  收费项目名称  String  100  是  填写业务系统内部项目名称
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.收据费目名称) || '"';
      --unit  计量单位  String  20  否
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.计算单位) || '"';
      --std  收费标准  Number  14,2  是
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.单价, 1);
      --number  数量  Number  14,2  是
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.数量, 1);
      --amt  金额  Number  14,2  是
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.结帐金额, 1);
      --selfAmt  自费金额  Number  14,2  是  如无金额，填写0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.自费金额, 1);
      --remark  备注  String  200  否
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.备注) || '"}';
    
      If Length(Nvl(v_分类明细, '') || v_Temp) > 32700 Then
        If c_分类明细 Is Null Then
          c_分类明细 := To_Clob(v_分类明细);
        Else
          c_分类明细 := c_分类明细 || To_Clob(',' || v_分类明细);
        End If;
        v_分类明细 := Null;
      End If;
    
      If v_分类明细 Is Null Then
        v_分类明细 := v_Temp;
      Else
        v_分类明细 := v_分类明细 || ',' || v_Temp;
      End If;
    
      n_票据总金额 := Nvl(n_票据总金额, 0) + Nvl(c_分类统计.结帐金额, 0);
      n_误差总额   := Nvl(n_误差总额, 0) + Nvl(c_分类统计.误差费, 0);
    End Loop;
    Totalmoney_Out := n_票据总金额;
    If v_分类明细 Is Not Null And c_分类明细 Is Not Null Then
      c_分类明细 := c_分类明细 || ',' || To_Clob(v_分类明细);
      ----chargeDetail 收费项目明细  String  不限  是  详见A-1,JSON格式列表
      c_分类明细 := To_Clob(',"chargeDetail":[') || c_分类明细 || To_Clob(']');
      v_分类明细 := Null;
    Elsif v_分类明细 Is Not Null Then
      v_分类明细 := ',"chargeDetail":[' || v_分类明细 || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --票据信息
    --Select Count(1) Into n_作废次数 From 电子票据使用记录 Where 票种 = n_应用场合 And 结算id = n_结算id;
    --结算ID||_||电子票据ID
    v_业务流水号 := n_结算id || '_' || Nvl(n_电子票据id, 0);
  
    v_开票点 := b_Einvoice_Request_Test.Get_Einvoice_Node(v_操作员编号, v_操作员姓名, n_用户id);
  
    v_票据信息 := '"busNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_业务流水号) || '"'; --业务流水号
    v_票据信息 := v_票据信息 || ',"busType":"' || b_Einvoice_Request_Test.Zljsonstr(v_业务标识) || '"'; --业务标识
    v_票据信息 := v_票据信息 || ',"payer":"' || b_Einvoice_Request_Test.Zljsonstr(v_患者姓名) || '"'; --患者姓名
    v_票据信息 := v_票据信息 || ',"busDateTime":"' ||
              b_Einvoice_Request_Test.Zljsonstr(To_Char(d_业务发生时间, 'yyyymmddHH24miss') || '000') || '"'; --业务发生时间
    v_票据信息 := v_票据信息 || ',"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_开票点) || '"'; --开票点编码:直接填写业务系统内部编码值，由医疗平台配置对照
    v_票据信息 := v_票据信息 || ',"payee":"' || b_Einvoice_Request_Test.Zljsonstr(v_收费员) || '"'; --收费员
  
    v_票据信息 := v_票据信息 || ',"author":"' || b_Einvoice_Request_Test.Zljsonstr(v_操作员姓名) || '"'; --票据编制人
    v_票据信息 := v_票据信息 || ',"totalAmt":' || b_Einvoice_Request_Test.Zljsonstr(n_票据总金额, 1); --开票总金额
    v_票据信息 := v_票据信息 || ',"remark":"' || '' || '"'; --备注
    -------------------------------------------------------------------------------------------
  
    --取缴费信息
    v_缴费 := Null;
    For c_缴费 In (Select Max(Decode(信息名, '订单号', 信息值, '支付订单号', 信息值, '')) As 支付定单号,
                        Max(Decode(信息名, '医保支付订单号', 信息值, '医保订单号', 信息值, '')) As 医保支付定单号,
                        Max(Decode(Upper(信息名), '支付宝公众号USERID', 信息值, '')) As 支付宝公众号userid,
                        Max(Decode(Upper(信息名), '支付宝小程序USERID', 信息值, '')) As 支付宝小程序userid,
                        Max(Decode(Upper(信息名), '微信公众号OPENID', 信息值, '')) As 微信公众号openid,
                        Max(Decode(Upper(信息名), '微信小程序OPENID', 信息值, '')) As 微信小程序openid
                 From (Select 信息名, 信息值
                        From 病人信息从表
                        Where 病人id = n_病人id And 信息名 In ('支付宝公众号USERID', '支付宝小程序USERID', '微信公众号OPENID', '微信小程序OPENID')
                        Union All
                        Select 交易项目, 交易内容
                        From 三方结算交易
                        Where 交易id In (Select ID From 病人预交记录 Where 结帐id = n_结算id) And 交易项目 Like '%订单号')) Loop
      v_缴费 := ',"alipayCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.支付宝公众号userid) || '"'; --患者支付宝账户
      v_缴费 := v_缴费 || ',"weChatOrderNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.支付定单号) || '"'; --微信支付订单号
      If v_版本号 = 'V3.1.0' Then
        --该版本才有此接点
        v_缴费 := v_缴费 || ',"weChatMedTransNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.医保支付定单号) || '"'; --微信医保支付订单号
      End If;
    
      If c_缴费.微信公众号openid Is Not Null Then
        v_缴费 := v_缴费 || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.微信公众号openid) || '"'; --微信公众号或小程序用户ID
      Else
        v_缴费 := v_缴费 || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.微信小程序openid) || '"'; --微信公众号或小程序用户ID
      End If;
      Exit;
    End Loop;
  
    -------------------------------------------------------------------------------------------
    --取通知信息
    Select To_Number(Max(参数值))
    Into n_缺省卡类别id
    From 三方接口配置
    Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '缺省卡类别ID';
    v_通知 := Null;
    For c_通知 In (Select Max(a.病人id) As 病人id, Max(a.姓名) As 姓名, Max(a.手机号) As 手机号, Max(a.Email) As Email, Max(1) As 缴款类型,
                        Max(a.身份证号) As 身份证号, Max(m.名称) As 卡类别, Max(m.卡号) As 卡号, Max(a.门诊号) As 门诊号
                 From 病人信息 A,
                      (
                        
                        Select 病人id, 名称, 编码, 卡号
                        From (Select b.病人id, c.名称, c.编码, b.卡号, Decode(b.卡类别id, n_缺省卡类别id, 2, c.缺省标志) As 缺省标志
                                From 病人医疗卡信息 B, 医疗卡类别 C
                                Where b.卡类别id = c.Id And b.病人id = n_病人id
                                Order By 缺省标志)
                        Where Rownum < 2) M
                 Where a.病人id = m.病人id(+)) Loop
    
      v_通知 := ',"tel":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.手机号) || '"'; --患者手机号码
      v_通知 := v_通知 || ',"email":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.Email) || '"'; --患者邮箱地址
      If v_版本号 = 'V3.1.0' Then
        --该版本才有此接点
        v_通知 := v_通知 || ',"payerType":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.缴款类型) || '"'; --交款人类型
      End If;
      v_通知 := v_通知 || ',"idCardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.身份证号) || '"'; --统一社会信用代码
    
      If c_通知.卡类别 Is Not Null Then
        v_通知 := v_通知 || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.卡类别) || '"'; --卡类型
        v_通知 := v_通知 || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.卡号) || '"'; --卡号
      Elsif c_通知.身份证号 Is Not Null Then
        Select Nvl(Max(参数值), '99998')
        Into v_参数值
        From 三方接口配置
        Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '身份证作卡类型编号';
        v_通知 := v_通知 || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡类型
        v_通知 := v_通知 || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.身份证号) || '"'; --卡号
      Else
        --没有一张卡，固定一种卡类别
        Select Nvl(Max(参数值), '99999')
        Into v_参数值
        From 三方接口配置
        Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '病人无卡的卡类别编号';
        v_通知 := v_通知 || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡类型
        Select Nvl(Max(参数值), '-')
        Into v_参数值
        From 三方接口配置
        Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '病人无卡的卡号';
        v_通知 := v_通知 || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡号
      End If;
      If Nvl(n_门诊号, 0) = 0 Then
        n_门诊号 := c_通知.门诊号;
      
      End If;
    
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --就诊信息
    Select Max(内容) Into v_Temp From zlRegInfo Where 项目 = '医疗机构类型';
  
    --性质:1-收费;2-结算（包括住院结算、特殊门诊结算）；3-预交
    Select Max(a.险类), Max(b.保险机构编码), Max(Nvl(a.病种名称, c.名称))
    Into n_险类, v_保险机构编码, v_病种名称
    From 保险结算记录 A, 保险类别 B, 保险病种 C
    Where a.险类 = b.序号 And a.病种id = c.Id(+) And a.记录id = n_结算id And a.性质 = Decode(n_应用场合, 2, 3, 3, 2, 1);
  
    Select Max(名称) Into v_医疗付款方式名称 From 医疗付款方式 Where 编码 = v_医疗付款方式编码;
    If Nvl(n_险类, 0) <> 0 Then
      Select Max(医保号) Into v_医保号 From 保险帐户 Where 病人id = n_病人id And 险类 = n_险类;
    End If;
  
    v_就诊编号 := Null;
    If Nvl(n_医嘱序号, 0) <> 0 Then
      Select Max(To_Char(a.发生时间, 'yyyy-mm-dd')), Max(b.编码), Max(b.名称), Max(a.No)
      Into v_就诊日期, v_就诊科室编码, v_就诊科室名称, v_就诊编号
      From 病人挂号记录 A, 部门表 B
      Where a.执行部门id = b.Id And
            a.No = (Select Max(挂号单) From 病人医嘱记录 Where ID = n_医嘱序号 Or 相关id = n_医嘱序号);
    Elsif Nvl(n_挂号id, 0) <> 0 Then
      Select Max(To_Char(a.发生时间, 'yyyy-mm-dd')), Max(b.编码), Max(b.名称), Max(a.No)
      Into v_就诊日期, v_就诊科室编码, v_就诊科室名称, v_就诊编号
      From 病人挂号记录 A, 部门表 B
      Where a.执行部门id = b.Id And a.Id = n_挂号id;
    End If;
    If v_就诊编号 Is Null Then
      --取最近一次挂号ID
      Select Max(a.Id), Max(To_Char(a.发生时间, 'yyyy-mm-dd')), Max(b.编码), Max(b.名称), Max(a.No)
      Into n_挂号id, v_就诊日期, v_就诊科室编码, v_就诊科室名称, v_就诊编号
      From 病人挂号记录 A, 部门表 B
      Where a.执行部门id = b.Id And
            a.Id = (Select ID
                    From (Select ID, 发生时间 From 病人挂号记录 Where 病人id = n_病人id Order By 发生时间 Desc)
                    Where Rownum < 2);
    End If;
  
    If v_病种名称 Is Null And Nvl(n_险类, 0) <> 0 Then
    
      Select Max(病种名称)
      Into v_病种名称
      From (
             
             Select Distinct a.名称 As 病种名称
             From 保险病种 A, 保险特准项目 B
             Where a.险类 = n_险类 And a.Id = b.病种id And
                   b.收费细目id In (Select Distinct 收费细目id From 门诊费用记录 Where 结帐id = n_结算id)
             Union All
             Select Distinct a.名称 As 病种名称
             From 保险病种 A, 保险特准项目 B
             Where a.险类 = n_险类 And a.Id = b.病种id And
                   b.大类 In (Select Distinct 保险大类id From 门诊费用记录 Where 结帐id = n_结算id))
      Where Rownum < 2;
    End If;
    v_就诊信息 := ',"medicalInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_Temp) || '"'; --医疗机构类型
    v_就诊信息 := v_就诊信息 || ',"medCareInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_保险机构编码) || '"'; --医保机构的唯一编码
    v_就诊信息 := v_就诊信息 || ',"medCareTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_医疗付款方式编码) || '"'; --医保类型编码
    v_就诊信息 := v_就诊信息 || ',"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(v_医疗付款方式名称) || '"'; --取值范围包括职工基本医疗保险、城乡居民基本医疗保险（城镇居民基本医疗保险、新型农村合作医疗保险）和其他医疗保险、非医保等
    v_就诊信息 := v_就诊信息 || ',"medicalInsuranceID":"' || b_Einvoice_Request_Test.Zljsonstr(v_医保号) || '"';
    v_就诊信息 := v_就诊信息 || ',"consultationDate":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊日期) || '"'; --患者就医时间
    v_就诊信息 := v_就诊信息 || ',"category":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊科室名称) || '"'; --就诊科室
    v_就诊信息 := v_就诊信息 || ',"patientCategoryCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊科室编码) || '"'; --就诊科室编码
    --patientNo  患者就诊编号  String  20  是  患者每次就诊一次就生成的一个新的编号。（患者登记号）
    v_就诊信息 := v_就诊信息 || ',"patientNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊编号) || '"';
    v_就诊信息 := v_就诊信息 || ',"patientId":"' || b_Einvoice_Request_Test.Zljsonstr(n_病人id) || '"'; --患者在业务系统中的唯一标识ID，类似身份证号码。
    v_就诊信息 := v_就诊信息 || ',"sex":"' || b_Einvoice_Request_Test.Zljsonstr(v_患者性别) || '"'; --性别
    v_就诊信息 := v_就诊信息 || ',"age":"' || b_Einvoice_Request_Test.Zljsonstr(v_患者年龄) || '"'; --年龄
    v_就诊信息 := v_就诊信息 || ',"caseNumber":"' || b_Einvoice_Request_Test.Zljsonstr(n_门诊号) || '"'; --病历号
    v_就诊信息 := v_就诊信息 || ',"specialDiseasesName":"' || b_Einvoice_Request_Test.Zljsonstr(v_病种名称) || '"'; --特殊病种名称
  
    -------------------------------------------------------------------------------------------
    --结算信息
    v_费用 := Null;
    For c_结算 In (Select 现金预交, 支票预交, 转账预交, 个人帐户支付, 医保统筹基金支付, 其它医保支付, 个人现金支付, Decode(Sign(现金支付), -1, 现金支付, 0) As 现金退款,
                        Decode(Sign(现金支付), -1, 支票支付, 0) As 支票退款, Decode(Sign(现金支付), -1, 转帐支付, 0) As 转帐退款,
                        Decode(Sign(现金支付), -1, 0, 现金支付) As 现金支付, Decode(Sign(现金支付), -1, 0, 支票支付) As 支票支付,
                        Decode(Sign(现金支付), -1, 0, 转帐支付) As 转帐支付,
                        Nvl(个人帐户支付, 0) + Nvl(医保统筹基金支付, 0) + Nvl(其它医保支付, 0) As 报销总额,
                        Nvl(结算总额, 0) - Nvl(个人帐户支付, 0) - Nvl(医保统筹基金支付, 0) - Nvl(其它医保支付, 0) As 自费金额, 结算总额, 医保结算号码,
                        0 As 个人帐户余额
                 From (Select /*+cardinality(b,10)*/
                         Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '现金', 1, 0), 0) * a.冲预交) As 现金预交,
                         Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '支票', 1, 0), 0) * a.冲预交) As 支票预交,
                         Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '支票', 0, '现金', 0, 1), 0) * a.冲预交) As 转账预交,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '个人账户支付', 1, 0)) * a.冲预交) As 个人帐户支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '医保统筹基金支付', 1, 0)) * a.冲预交) As 医保统筹基金支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '其它医保支付', 1, 0)) * a.冲预交) As 其它医保支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '其它医保支付', 0, '个人账户支付', 0, '医保统筹基金支付', 0, 1)) *
                              a.冲预交) As 个人现金支付,
                         Max(Decode(Mod(a.记录性质, 10), 1, 0,
                                     Decode(c.开票结算方式, '其它医保支付', 结算号码, '个人账户支付', 结算号码, '医保统筹基金支付', 结算号码, ''))) As 医保结算号码,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(a.结算方式, '现金', 1, 0)) * a.冲预交) As 现金支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(a.结算方式, '支票', 1, 0)) * a.冲预交) As 支票支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(a.结算方式, '现金', 0, '支票', 0, 1) * a.冲预交)) As 转帐支付,
                         Sum(冲预交) As 结算总额
                        From 病人预交记录 A, Table(l_结帐id) B, 开票结算对照 C
                        Where a.结帐id = b.Column_Value And a.结算方式 = c.结算方式(+)))
    
     Loop
      --accountPay  个人账户支付  Number  14,2  是  按政策规定用个人账户支付参保人的医疗费用（含基本医疗保险目录范围内和目录范围外的费用）；
      --          如无金额，填写0
      v_费用 := ',"accountPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_结算.个人帐户支付, 0), 1);
      --fundPay  医保统筹基金支付  Number  14,2  是  患者本次就医所发生的医疗费用中按规定由基本医疗保险统筹基金支付的金额；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"fundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_结算.医保统筹基金支付, 0), 1);
      --otherfundPay  其它医保支付  Number  14,2  是  患者本次就医所发生的医疗费用中按规定由大病保险、医疗救助、公务员医疗补助、大额补充、企业补充等基金或资金支付的金额；
      v_费用 := v_费用 || ',"otherfundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_结算.其它医保支付, 0), 1);
      --ownPay  自费金额  Number  14,2  是  患者本次就医所发生的医疗费用中按照有关规定不属于基本医疗保险目录范围而全部由个人支付的费用；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"ownPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.自费金额, 1);
      --selfConceitedAmt  个人自负  Number  14,2  是  医保患者起付标准内个人支付费用；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"selfConceitedAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfPayAmt  个人自付  Number  14,2  是  患者本次就医所发生的医疗费用中由个人负担的属于基本医疗保险目录范围内自付部分的金额；开展按病种、病组、床日等打包付费方式且由患者定额付费的费用。该项为个人所得税大病医疗专项附加扣除信；息项如无金额，填写0
      v_费用 := v_费用 || ',"selfPayAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfCashPay  个人现金支付  Number  14,2  是  个人通过现金、银行卡、微信、支付宝等渠道支付的金额；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"selfCashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.个人现金支付, 1);
      --cashPay  现金预交款金额  Number  14,2  否
      v_费用 := v_费用 || ',"cashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.现金预交, 1);
      --chequePay  支票预交款金额  Number  14,2  否
      v_费用 := v_费用 || ',"chequePay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.支票预交, 1);
      --transferAccountPay  转账预交款金额  Number  14,2  否
      v_费用 := v_费用 || ',"transferAccountPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.转账预交, 1);
      --cashRecharge  补交金额(现金)  Number  14,2  否
      v_费用 := v_费用 || ',"cashRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.现金支付, 1);
      --chequeRecharge  补交金额(支票)  Number  14,2  否
      v_费用 := v_费用 || ',"chequeRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.支票支付, 1);
      --transferRecharge  补交金额（转账）  Number  14,2  否
      v_费用 := v_费用 || ',"transferRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.转帐支付, 1);
      --cashRefund  退还金额(现金)  Number  14,2  否
      v_费用 := v_费用 || ',"cashRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.现金退款, 1);
      --chequeRefund  退交金额(支票)  Number  14,2  否
      v_费用 := v_费用 || ',"chequeRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.支票退款, 1);
      --transferRefund  退交金额(转账)  Number  14,2  否
      v_费用 := v_费用 || ',"transferRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.转帐退款, 1);
      --ownAcBalance  个人账户余额  Number  14,2  否
      v_费用 := v_费用 || ',"ownAcBalance":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.个人帐户余额, 1);
      --reimbursementAmt  报销总金额  Number  14,2  否  医保结算后返回的总金额
      v_费用 := v_费用 || ',"reimbursementAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.报销总额, 1);
      --balancedNumber  结算号  String  100  否  医保结算后生成的号码/入账唯一值
      v_费用 := v_费用 || ',"balancedNumber":"' || b_Einvoice_Request_Test.Zljsonstr(c_结算.医保结算号码) || '"';
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --交费渠道
    v_缴费渠道 := Null;
    For c_渠道 In (Select /*+cardinality(b,10)*/
                  Nvl(c.渠道编码, Nvl(d.渠道编码, '-')) As 渠道编码, Sum(冲预交) As 结算总额
                 From 病人预交记录 A, Table(l_结帐id) B, 收费渠道对照 C,
                      (Select 结算方式, 渠道编码 From 收费渠道对照 D Where 卡类别id Is Null) D
                 Where a.结帐id = b.Column_Value And a.卡类别id = c.卡类别id(+) And a.结算方式 = c.结算方式(+) And a.结算方式 = d.结算方式(+)
                 Group By Nvl(c.渠道编码, Nvl(d.渠道编码, '-'))
                 Order By 渠道编码)
    
     Loop
      --payChannelCode  交费渠道编码  String  10  是
      If v_缴费渠道 Is Null Then
        v_缴费渠道 := Nvl(v_缴费渠道, '') || '{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_渠道.渠道编码, 0)) || '"';
      Else
        v_缴费渠道 := v_缴费渠道 || ',{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_渠道.渠道编码, 0)) || '"';
      End If;
      --payChannelValue  交费渠道金额  Number  14,2  是
      v_缴费渠道 := v_缴费渠道 || ',"payChannelValue":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_渠道.结算总额, 0), 1) || '}';
    End Loop;
  
    If v_缴费渠道 Is Not Null Then
      --payChannelDetail  交费渠道列表  String  不限  是  直接填写业务系统内部编码值，由医疗平台配置对照如：附录4交费渠道列表
      --        详见A-5,JSON格式列表
      v_缴费渠道 := ',"payChannelDetail":[' || v_缴费渠道 || ']';
    Else
      v_缴费渠道 := ',"payChannelDetail":[]';
    End If;
  
    -------------------------------------------------------------------------------------------
    --其他医保信息
    v_其它医保信息 := Null;
    --otherMedicalList  其它医保信息列表  String  不限  否  填写其它未知医保信息（在电子票据上以内容拼接方式显示）
    --            详见A-4,JSON格式列表
    --  infoNo  序号  Integer  不限  是  默认从1开始，每项数据序号值递增1，本次不允许重复
    --  infoName  医保信息名称  String  100  是  如费用报销类型编码，可参考附录7医保报销类型列表
    --  infoValue  医保信息值  String  100  是  如费用报销金额
    --  infoOther  医保其它信息  String  100  否  如医保报销比例。
  
    -------------------------------------------------------------------------------------------
    --其它扩展信息
    v_其它扩展信息 := Null;
    --otherInfo  其它扩展信息列表  String  不限  否  填写信息需要在电子票据上单独显示的其它扩展信息（未知信息）
    --          详见A-3,JSON格式列表
    --  infoNo  序号  Integer  不限  是  默认从1开始，每项数据序号值递增1，本次不允许重复
    --  infoName  扩展信息名称  String  100  是
    --  infoValue  扩展信息值  String  500  是
  
    c_交易信息 := To_Clob('{' || v_票据信息);
    c_交易信息 := c_交易信息 || To_Clob(v_缴费);
    c_交易信息 := c_交易信息 || To_Clob(v_通知);
    c_交易信息 := c_交易信息 || To_Clob(v_就诊信息);
    c_交易信息 := c_交易信息 || To_Clob(v_费用);
    c_交易信息 := c_交易信息 || To_Clob(v_缴费渠道);
  
    If v_其它扩展信息 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_其它扩展信息);
    End If;
    If v_其它医保信息 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_其它医保信息);
    End If;
    --  eBillRelateNo  业务票据关联号  String  32  否  如一笔业务数据需要开具N张电子票据，则N张电子票对应该值保持一致，用于后期关联查询
    c_交易信息 := c_交易信息 || To_Clob(',"eBillRelateNo":""');
    If v_分类明细 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_分类明细);
    Else
      c_交易信息 := c_交易信息 || c_分类明细;
    End If;
  
    If v_明细 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_明细);
    Else
      c_交易信息 := c_交易信息 || c_明细;
    End If;
    c_交易信息  := c_交易信息 || To_Clob('}');
    Reqdata_Out := c_交易信息;
    Code_Out    := 1;
  Exception
    When Others Then
      Message_Out := SQLCode || ':' || SQLErrM;
      Code_Out    := 0;
  End Get_Sendcarddata_Create;

  Procedure Get_Registerdata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  ) Is
    --
    ---------------------------------------------------------------------------
    --功能:获取挂号开票数据
    --入参:
    --    Json_In,格式如下
    --  input
    --    occasion N 1  应用场合: 4-挂号
    --    einvoice_id  N,1 当前电子票据ID
    --    balance_id N 1  结算ID  occasion=2时，预交id;occasion<>2表示结帐id
    --    writeoff_id  N 1  冲销ID  occasion=2时，冲销预交id;occasion<>2表示冲销id
    --出参:
    --  ReqData_Out-返回的业务请求数据
    --  Totalmoney_Out-票据总额
    --  Code_Out-获取是否成功：0-失败；1-成功
    --  Message_Out 错误信息
    ---------------------------------------------------------------------------
    j_Input PLJson;
    j_Json  PLJson;
  
    n_应用场合 Number(2);
    n_结算id   病人预交记录.结帐id%Type;
    --n_冲销id   病人预交记录.结帐id%Type;
  
    v_业务标识   Varchar2(20);
    v_业务流水号 Varchar2(50);
  
    v_开票点       Varchar2(100);
    v_缴费         Varchar2(32767);
    v_票据信息     Varchar2(32767);
    v_就诊信息     Varchar2(32767);
    v_通知         Varchar2(32767);
    v_缴费渠道     Varchar2(32767);
    v_费用         Varchar2(32767);
    v_其它扩展信息 Varchar2(32767);
    v_其它医保信息 Varchar2(32767);
    c_明细         Clob;
    v_明细         Varchar2(32767);
    c_分类明细     Clob;
    v_分类明细     Varchar2(32767);
    c_交易信息     Clob; --最终返回的交易信息集
  
    n_门诊号       病人信息.门诊号%Type;
    n_病人id       病人预交记录.病人id%Type;
    v_患者姓名     门诊费用记录.姓名%Type;
    v_患者性别     门诊费用记录.性别%Type;
    v_患者年龄     门诊费用记录.年龄%Type;
    d_业务发生时间 门诊费用记录.登记时间%Type;
    v_收费员       门诊费用记录.操作员姓名%Type;
  
    n_缺省卡类别id     Number(18);
    v_参数值           Varchar2(100);
    n_票据总金额       门诊费用记录.结帐金额%Type;
    n_误差总额         门诊费用记录.结帐金额%Type;
    n_用户id           人员表.Id%Type;
    v_操作员编号       人员表.编号%Type;
    v_操作员姓名       人员表.姓名%Type;
    v_Temp             Varchar2(32767);
    v_医疗付款方式编码 医疗付款方式.编码%Type;
    v_医疗付款方式名称 医疗付款方式.名称%Type;
    n_险类             保险结算记录.险类%Type;
    v_保险机构编码     保险类别.保险机构编码%Type;
    n_医嘱序号         门诊费用记录.医嘱序号%Type;
    n_挂号id           门诊费用记录.挂号id%Type;
    v_病种名称         保险病种.名称%Type;
    v_就诊日期         Varchar2(20);
    v_就诊科室编码     部门表.编码%Type;
    v_就诊科室名称     部门表.名称%Type;
    v_就诊编号         Varchar2(50);
    v_结帐ids          Varchar2(32767);
    v_医保号           保险帐户.医保号%Type;
    l_结帐id           t_NumList := t_NumList();
    v_版本号           Varchar2(30);
    n_电子票据id       电子票据使用记录.Id%Type;
  Begin
    j_Input := PLJson(Json_In);
    j_Json  := j_Input.Get_Pljson('input');
  
    n_应用场合 := Nvl(j_Json.Get_Number('occasion'), 0);
    n_结算id   := j_Json.Get_Number('balance_id');
    --n_冲销id     := Nvl(j_Json.Get_Number('writeoff_id'), 0);
    n_电子票据id := Nvl(j_Json.Get_Number('einvoice_id'), 0);
  
    If Nvl(n_应用场合, 0) = 0 Then
      Code_Out    := 0;
      Message_Out := '无效的应用场景';
      Return;
    End If;
  
    Select Nvl(Max(参数值), 'V2.0.3')
    Into v_版本号
    From 三方接口配置
    Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '支持版本';
  
    b_Einvoice_Request_Test.Get_Identity(n_用户id, v_操作员编号, v_操作员姓名);
  
    --v_业务标识:01  住院,02  门诊,03  急诊,04  门特,05  体检中心,06  挂号,07  住院预交金,08  体检预交金
    Select Decode(n_应用场合, 1, '02', 2, '07', 3, '01', 4, '06', 5, '02', '02') Into v_业务标识 From Dual;
  
    n_票据总金额   := 0;
    d_业务发生时间 := Null;
    v_结帐ids      := Null;
    c_明细         := Null;
    v_明细         := Null;
    For c_收费细目 In (Select Min(a.Id) As 费用id, a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号) As 序号, a.收费细目id, Max(a.计算单位) As 计算单位,
                          Sum(a.标准单价) As 价格, Avg(Nvl(a.付数, 1) * Nvl(a.数次, 0)) As 数量, Sum(a.应收金额) As 应收金额,
                          Sum(a.实收金额) As 实收金额, Sum(a.结帐金额) As 结帐金额, Sum(a.实收金额) - Sum(a.统筹金额) As 自费金额,
                          Max(t.编码) As 医保项目编码, Max(t.名称) As 医保项目名称, Max(t.统筹比额) As 医保报销比例, Max(a.摘要) As 备注,
                          Max(a.费用类型) As 费用类型, Max(a.操作员编号) As 操作员编号, Max(a.操作员姓名) As 操作员姓名, Max(a.姓名) As 姓名,
                          Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.病人id) As 病人id, Max(a.登记时间) As 登记时间,
                          Max(a.付款方式) As 付款方式编码, Max(a.收据费目) As 收据费目, Max(c.编码) As 收据费目编码, Max(a.医嘱序号) As 医嘱序号,
                          Max(B1.Id) As 挂号id, Max(d.编码) As 类别编码, Max(d.类别) As 类别名称, Max(b.编码) As 项目编码, Max(b.名称) As 项目名称,
                          Max(b.规格) As 规格, Max(q.药品剂型) As 药品剂型
                   From 门诊费用记录 A, 病人挂号记录 B1, 收费项目目录 B, 收据费目 C, 收费类别 D, 药品规格 M, 药品特性 Q, 诊疗项目目录 J, 保险支付大类 T, 支付类别对照 S
                   Where a.No = B1.No And a.No In (Select Distinct NO From 门诊费用记录 Where 结帐id = n_结算id) And a.记录性质 = 4 And
                         a.收费类别 = d.编码(+) And a.收费细目id = b.Id And a.收据费目 = c.名称(+) And a.收费细目id = m.药品id(+) And
                         m.药名id = q.药名id(+) And q.药名id = j.Id(+) And a.保险大类id = t.Id(+) And t.性质(+) = 1 And
                         a.保险大类id = s.保险大类id(+)
                   Group By a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号), a.收费细目id, c.编码, c.名称, j.编码, j.名称
                   Order By NO, 序号) Loop
      If v_患者姓名 Is Null Then
        v_患者姓名 := c_收费细目.姓名;
        v_患者性别 := c_收费细目.性别;
        v_患者年龄 := c_收费细目.年龄;
        n_病人id   := c_收费细目.病人id;
      End If;
      If d_业务发生时间 Is Null And Nvl(c_收费细目.记录状态, 0) = 1 Then
        --取原始业务发生时间
        d_业务发生时间 := c_收费细目.登记时间;
        v_收费员       := c_收费细目.操作员姓名;
      End If;
      If v_医疗付款方式编码 Is Null Then
        v_医疗付款方式编码 := c_收费细目.付款方式编码;
      End If;
      If Nvl(n_医嘱序号, 0) = 0 Then
        n_医嘱序号 := c_收费细目.医嘱序号;
      End If;
      If Nvl(n_挂号id, 0) = 0 Then
        n_挂号id := c_收费细目.挂号id;
      End If;
    
      If Instr(Nvl(v_结帐ids, '') || ',', ',' || c_收费细目.结帐id || ',') = 0 Then
        l_结帐id.Extend;
        l_结帐id(l_结帐id.Count) := c_收费细目.结帐id;
      End If;
    
      --listDetailNo  明细流水号  String  60  否  明细流水号
      v_Temp := '{' || '"listDetailNo":"' || b_Einvoice_Request_Test.Zljsonstr(LPad(c_收费细目.费用id, 20, '0')) || '"';
      --chargeCode  收费项目代码  String  50  否  填写业务系统内部编码值，由医疗平台配置对照,如：床位费、检查费
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.收据费目编码) || '"';
      --chargeName  收费项目名称  String  100  否
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.收据费目) || '"';
      --prescribeCode  处方编码  String  60  否
      v_Temp := v_Temp || ',' || '"prescribeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.No) || '"';
      --listTypeCode  药品类别编码  String  50  否  如药品分类编码01，有则填写
      v_Temp := v_Temp || ',' || '"listTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.类别编码) || '"';
      --listTypeName  药品类别名称  String  50  否  如药品分类名称，抗生素类抗感染药物
      v_Temp := v_Temp || ',' || '"listTypeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.类别名称) || '"';
      --code  编码  String  50  否  如药品编码，有则填写
      v_Temp := v_Temp || ',' || '"code":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.项目编码) || '"';
      --name  药品名称  String  50  是  如药品名称，器材名称等
      v_Temp := v_Temp || ',' || '"name":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.项目名称) || '"';
      --form  剂型  String  50  否
      v_Temp := v_Temp || ',' || '"form":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.药品剂型) || '"';
      --specification  规格  String  50  否
      v_Temp := v_Temp || ',' || '"specification":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.规格) || '"';
      --unit  计量单位   String  20  否
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.计算单位) || '"';
      --std  单价  Number  14,6  是
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.价格, 1);
      --number  数量  Number  14,6  是
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.数量, 1);
      --amt  金额  Number  14,6  是
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.实收金额, 1);
      --selfAmt  自费金额  Number  14,6  是  如无金额，填写0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.自费金额, 1);
      --receivableAmt  应收费用  Number  14,6  否
      v_Temp := v_Temp || ',' || '"receivableAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.应收金额, 1);
      --medicalCareType  医保药品分类  String  1  否  1：无自负/甲
      --          2：有自负/乙
      --          3：全自负/丙
      v_Temp := v_Temp || ',' || '"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.医保项目编码) || '"';
      --medCareItemType  医保项目类型  String  100  否
      v_Temp := v_Temp || ',' || '"medCareItemType":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.医保项目名称) || '"';
      --medReimburseRate  医保报销比例  Number  3,2  否
      v_Temp := v_Temp || ',' || '"medReimburseRate":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.医保报销比例, 1);
      --remark  备注  String  200  否
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.备注) || '"';
      --sortNo  序号  Integer  不限  否  序号
      v_Temp := v_Temp || ',' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.序号, 1);
      --chrgtype  费用类型  String  50  否
      v_Temp := v_Temp || ',' || '"chrgtype":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.费用类型) || '"}';
    
      If Length(Nvl(v_明细, '') || v_Temp) > 32700 Then
        If c_明细 Is Null Then
          c_明细 := To_Clob(v_明细);
        Else
          c_明细 := c_明细 || To_Clob(',' || v_明细);
        End If;
        v_明细 := Null;
      End If;
    
      If v_明细 Is Null Then
        v_明细 := v_Temp;
      Else
        v_明细 := v_明细 || ',' || v_Temp;
      End If;
    End Loop;
  
    If v_明细 Is Not Null And c_明细 Is Not Null Then
      --listDetail  清单项目明细  String  不限  是  详见A-2,JSON格式列表
      c_明细 := c_明细 || ',' || To_Clob(v_明细);
      c_明细 := To_Clob(',"listDetail":[') || c_明细 || To_Clob(']');
    
      v_明细 := Null;
    Elsif v_明细 Is Not Null Then
      v_明细 := ',"listDetail":[' || v_明细 || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --分类明细
    v_分类明细 := Null;
    c_分类明细 := Null;
    For c_分类统计 In (Select Rownum As 序号, 收据费目编码, 收据费目名称, 数量, 计算单位, Round(单价, 2) As 单价, Round(结帐金额, 2) As 结帐金额,
                          Round(自费金额, 2) As 自费金额, 备注, 结帐金额 - Round(结帐金额, 2) As 误差费
                   From (Select /*+cardinality(b,10)*/
                           c.编码 As 收据费目编码, a.收据费目 As 收据费目名称, 1 As 数量, '' As 计算单位, Sum(a.结帐金额) As 单价, a.收据费目,
                           Sum(a.结帐金额) As 结帐金额, Sum(a.结帐金额) - Sum(a.统筹金额) As 自费金额, '' As 备注
                          From 门诊费用记录 A, Table(l_结帐id) B, 收据费目 C
                          Where a.结帐id = b.Column_Value And a.收据费目 = c.名称(+)
                          Group By c.编码, a.收据费目)) Loop
      --sortNo  序号  Integer  不限  是  默认从1开始，每个收费项目序号值递增1，本次不允许重复
      v_Temp := '{' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.序号, 1);
      --chargeCode  收费项目代码  String  50  是  填写业务系统内部编码值，由医疗平台配置对照
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.收据费目编码) || '"';
      --chargeName  收费项目名称  String  100  是  填写业务系统内部项目名称
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.收据费目名称) || '"';
      --unit  计量单位  String  20  否
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.计算单位) || '"';
      --std  收费标准  Number  14,2  是
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.单价, 1);
      --number  数量  Number  14,2  是
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.数量, 1);
      --amt  金额  Number  14,2  是
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.结帐金额, 1);
      --selfAmt  自费金额  Number  14,2  是  如无金额，填写0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.自费金额, 1);
      --remark  备注  String  200  否
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.备注) || '"}';
    
      If Length(Nvl(v_分类明细, '') || v_Temp) > 32700 Then
        If c_分类明细 Is Null Then
          c_分类明细 := To_Clob(v_分类明细);
        Else
          c_分类明细 := c_分类明细 || To_Clob(',' || v_分类明细);
        End If;
        v_分类明细 := Null;
      End If;
    
      If v_分类明细 Is Null Then
        v_分类明细 := v_Temp;
      Else
        v_分类明细 := v_分类明细 || ',' || v_Temp;
      End If;
    
      n_票据总金额 := Nvl(n_票据总金额, 0) + Nvl(c_分类统计.结帐金额, 0);
      n_误差总额   := Nvl(n_误差总额, 0) + Nvl(c_分类统计.误差费, 0);
    End Loop;
    Totalmoney_Out := n_票据总金额;
    If v_分类明细 Is Not Null And c_分类明细 Is Not Null Then
      c_分类明细 := c_分类明细 || ',' || To_Clob(v_分类明细);
      --chargeDetail 收费项目明细  String  不限  是  详见A-1,JSON格式列表
      c_分类明细 := To_Clob(',"chargeDetail":[') || c_分类明细 || To_Clob(']');
      v_分类明细 := Null;
    Elsif v_分类明细 Is Not Null Then
      v_分类明细 := ',"chargeDetail":[' || v_分类明细 || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --票据信息
    --Select Count(1) Into n_作废次数 From 电子票据使用记录 Where 票种 = n_应用场合 And 结算id = n_结算id;
    --lpad(电子票据作废次数,5) & Lpad(原结帐ID,20)
    --结算ID||_||电子票据ID
    v_业务流水号 := n_结算id || '_' || Nvl(n_电子票据id, 0);
    v_开票点     := b_Einvoice_Request_Test.Get_Einvoice_Node(v_操作员编号, v_操作员姓名, n_用户id);
  
    v_票据信息 := '"busNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_业务流水号) || '"'; --业务流水号
    v_票据信息 := v_票据信息 || ',"busType":"' || b_Einvoice_Request_Test.Zljsonstr(v_业务标识) || '"'; --业务标识
    v_票据信息 := v_票据信息 || ',"payer":"' || b_Einvoice_Request_Test.Zljsonstr(v_患者姓名) || '"'; --患者姓名
    v_票据信息 := v_票据信息 || ',"busDateTime":"' ||
              b_Einvoice_Request_Test.Zljsonstr(To_Char(d_业务发生时间, 'yyyymmddHH24miss') || '000') || '"'; --业务发生时间
    v_票据信息 := v_票据信息 || ',"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_开票点) || '"'; --开票点编码:直接填写业务系统内部编码值，由医疗平台配置对照
    v_票据信息 := v_票据信息 || ',"payee":"' || b_Einvoice_Request_Test.Zljsonstr(v_收费员) || '"'; --收费员
  
    v_票据信息 := v_票据信息 || ',"author":"' || b_Einvoice_Request_Test.Zljsonstr(v_操作员姓名) || '"'; --票据编制人
    v_票据信息 := v_票据信息 || ',"totalAmt":' || b_Einvoice_Request_Test.Zljsonstr(n_票据总金额, 1); --开票总金额
    v_票据信息 := v_票据信息 || ',"remark":"' || '' || '"'; --备注
    -------------------------------------------------------------------------------------------
  
    --取缴费信息
    v_缴费 := Null;
    For c_缴费 In (Select Max(Decode(信息名, '订单号', 信息值, '支付订单号', 信息值, '')) As 支付定单号,
                        Max(Decode(信息名, '医保支付订单号', 信息值, '医保订单号', 信息值, '')) As 医保支付定单号,
                        Max(Decode(Upper(信息名), '支付宝公众号USERID', 信息值, '')) As 支付宝公众号userid,
                        Max(Decode(Upper(信息名), '支付宝小程序USERID', 信息值, '')) As 支付宝小程序userid,
                        Max(Decode(Upper(信息名), '微信公众号OPENID', 信息值, '')) As 微信公众号openid,
                        Max(Decode(Upper(信息名), '微信小程序OPENID', 信息值, '')) As 微信小程序openid
                 From (Select 信息名, 信息值
                        From 病人信息从表
                        Where 病人id = n_病人id And 信息名 In ('支付宝公众号USERID', '支付宝小程序USERID', '微信公众号OPENID', '微信小程序OPENID')
                        Union All
                        Select 交易项目, 交易内容
                        From 三方结算交易
                        Where 交易id In (Select ID From 病人预交记录 Where 结帐id = n_结算id) And 交易项目 Like '%订单号')) Loop
      v_缴费 := ',"alipayCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.支付宝公众号userid) || '"'; --患者支付宝账户
      v_缴费 := v_缴费 || ',"weChatOrderNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.支付定单号) || '"'; --微信支付订单号
      If v_版本号 = 'V3.1.0' Then
        --该版本才有此接点
        v_缴费 := v_缴费 || ',"weChatMedTransNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.医保支付定单号) || '"'; --微信医保支付订单号
      End If;
    
      If c_缴费.微信公众号openid Is Not Null Then
        v_缴费 := v_缴费 || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.微信公众号openid) || '"'; --微信公众号或小程序用户ID
      Else
        v_缴费 := v_缴费 || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.微信小程序openid) || '"'; --微信公众号或小程序用户ID
      End If;
      Exit;
    End Loop;
  
    -------------------------------------------------------------------------------------------
    --取通知信息
    Select To_Number(Max(参数值))
    Into n_缺省卡类别id
    From 三方接口配置
    Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '缺省卡类别ID';
    v_通知 := Null;
    For c_通知 In (Select Max(a.病人id) As 病人id, Max(a.姓名) As 姓名, Max(a.手机号) As 手机号, Max(a.Email) As Email, Max(1) As 缴款类型,
                        Max(a.身份证号) As 身份证号, Max(m.名称) As 卡类别, Max(m.卡号) As 卡号, Max(a.门诊号) As 门诊号
                 From 病人信息 A,
                      (
                        
                        Select 病人id, 名称, 编码, 卡号
                        From (Select b.病人id, c.名称, c.编码, b.卡号, Decode(b.卡类别id, n_缺省卡类别id, 2, c.缺省标志) As 缺省标志
                                From 病人医疗卡信息 B, 医疗卡类别 C
                                Where b.卡类别id = c.Id And b.病人id = n_病人id
                                Order By 缺省标志)
                        Where Rownum < 2) M
                 Where a.病人id = m.病人id(+)) Loop
    
      v_通知 := ',"tel":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.手机号) || '"'; --患者手机号码
      v_通知 := v_通知 || ',"email":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.Email) || '"'; --患者邮箱地址
      If v_版本号 = 'V3.1.0' Then
        --该版本才有此接点
        v_通知 := v_通知 || ',"payerType":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.缴款类型) || '"'; --交款人类型
      End If;
      v_通知 := v_通知 || ',"idCardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.身份证号) || '"'; --统一社会信用代码
    
      If c_通知.卡类别 Is Not Null Then
        v_通知 := v_通知 || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.卡类别) || '"'; --卡类型
        v_通知 := v_通知 || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.卡号) || '"'; --卡号
      Elsif c_通知.身份证号 Is Not Null Then
        Select Nvl(Max(参数值), '99998')
        Into v_参数值
        From 三方接口配置
        Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '身份证作卡类型编号';
        v_通知 := v_通知 || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡类型
        v_通知 := v_通知 || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.身份证号) || '"'; --卡号
      Else
        --没有一张卡，固定一种卡类别
        Select Nvl(Max(参数值), '99999')
        Into v_参数值
        From 三方接口配置
        Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '病人无卡的卡类别编号';
        v_通知 := v_通知 || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡类型
        Select Nvl(Max(参数值), '-')
        Into v_参数值
        From 三方接口配置
        Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '病人无卡的卡号';
        v_通知 := v_通知 || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡号
      End If;
      If Nvl(n_门诊号, 0) = 0 Then
        n_门诊号 := c_通知.门诊号;
      
      End If;
    
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --就诊信息
    Select Max(内容) Into v_Temp From zlRegInfo Where 项目 = '医疗机构类型';
  
    --性质:1-收费;2-结算（包括住院结算、特殊门诊结算）；3-预交
    Select Max(a.险类), Max(b.保险机构编码), Max(Nvl(a.病种名称, c.名称))
    Into n_险类, v_保险机构编码, v_病种名称
    From 保险结算记录 A, 保险类别 B, 保险病种 C
    Where a.险类 = b.序号 And a.病种id = c.Id(+) And a.记录id = n_结算id And a.性质 = Decode(n_应用场合, 2, 3, 3, 2, 1);
  
    Select Max(名称) Into v_医疗付款方式名称 From 医疗付款方式 Where 编码 = v_医疗付款方式编码;
    If Nvl(n_险类, 0) <> 0 Then
      Select Max(医保号) Into v_医保号 From 保险帐户 Where 病人id = n_病人id And 险类 = n_险类;
    End If;
  
    v_就诊编号 := Null;
    If Nvl(n_挂号id, 0) <> 0 Then
      Select Max(To_Char(a.发生时间, 'yyyy-mm-dd')), Max(b.编码), Max(b.名称), Max(a.No)
      Into v_就诊日期, v_就诊科室编码, v_就诊科室名称, v_就诊编号
      From 病人挂号记录 A, 部门表 B
      Where a.执行部门id = b.Id And a.Id = n_挂号id;
    End If;
  
    If v_病种名称 Is Null And Nvl(n_险类, 0) <> 0 Then
      Select Max(病种名称)
      Into v_病种名称
      From (Select Distinct a.名称 As 病种名称
             From 保险病种 A, 保险特准项目 B
             Where a.险类 = n_险类 And a.Id = b.病种id And
                   b.收费细目id In (Select Distinct 收费细目id From 门诊费用记录 Where 结帐id = n_结算id)
             Union All
             Select Distinct a.名称 As 病种名称
             From 保险病种 A, 保险特准项目 B
             Where a.险类 = n_险类 And a.Id = b.病种id And
                   b.大类 In (Select Distinct 保险大类id From 门诊费用记录 Where 结帐id = n_结算id))
      Where Rownum < 2;
    End If;
  
    v_就诊信息 := ',"medicalInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_Temp) || '"'; --医疗机构类型
    v_就诊信息 := v_就诊信息 || ',"medCareInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_保险机构编码) || '"'; --医保机构的唯一编码
    v_就诊信息 := v_就诊信息 || ',"medCareTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_医疗付款方式编码) || '"'; --医保类型编码
    v_就诊信息 := v_就诊信息 || ',"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(v_医疗付款方式名称) || '"'; --取值范围包括职工基本医疗保险、城乡居民基本医疗保险（城镇居民基本医疗保险、新型农村合作医疗保险）和其他医疗保险、非医保等
    v_就诊信息 := v_就诊信息 || ',"medicalInsuranceID":"' || b_Einvoice_Request_Test.Zljsonstr(v_医保号) || '"';
  
    v_就诊信息 := v_就诊信息 || ',"consultationDate":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊日期) || '"'; --患者就医时间
    v_就诊信息 := v_就诊信息 || ',"category":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊科室名称) || '"'; --就诊科室
    v_就诊信息 := v_就诊信息 || ',"patientCategory":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊科室编码) || '"'; --就诊科室编码
    --patientNo  患者就诊编号  String  20  是  患者每次就诊一次就生成的一个新的编号。（患者登记号）
    v_就诊信息 := v_就诊信息 || ',"patientNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊编号) || '"';
    v_就诊信息 := v_就诊信息 || ',"patientId":"' || b_Einvoice_Request_Test.Zljsonstr(n_病人id) || '"'; --患者在业务系统中的唯一标识ID，类似身份证号码。
    v_就诊信息 := v_就诊信息 || ',"sex":"' || b_Einvoice_Request_Test.Zljsonstr(v_患者性别) || '"'; --性别
    v_就诊信息 := v_就诊信息 || ',"age":"' || b_Einvoice_Request_Test.Zljsonstr(v_患者年龄) || '"'; --年龄
    -------------------------------------------------------------------------------------------
    --结算信息
    v_费用 := Null;
    For c_结算 In (Select 现金预交, 支票预交, 转账预交, 个人帐户支付, 医保统筹基金支付, 其它医保支付, 个人现金支付, Decode(Sign(现金支付), -1, 现金支付, 0) As 现金退款,
                        Decode(Sign(现金支付), -1, 支票支付, 0) As 支票退款, Decode(Sign(现金支付), -1, 转帐支付, 0) As 转帐退款,
                        Decode(Sign(现金支付), -1, 0, 现金支付) As 现金支付, Decode(Sign(现金支付), -1, 0, 支票支付) As 支票支付,
                        Decode(Sign(现金支付), -1, 0, 转帐支付) As 转帐支付,
                        Nvl(个人帐户支付, 0) + Nvl(医保统筹基金支付, 0) + Nvl(其它医保支付, 0) As 报销总额,
                        Nvl(结算总额, 0) - Nvl(个人帐户支付, 0) - Nvl(医保统筹基金支付, 0) - Nvl(其它医保支付, 0) As 自费金额, 结算总额, 医保结算号码,
                        0 As 个人帐户余额
                 From (Select /*+cardinality(b,10)*/
                         Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '现金', 1, 0), 0) * a.冲预交) As 现金预交,
                         Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '支票', 1, 0), 0) * a.冲预交) As 支票预交,
                         Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '支票', 0, '现金', 0, 1), 0) * a.冲预交) As 转账预交,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '个人账户支付', 1, 0)) * a.冲预交) As 个人帐户支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '医保统筹基金支付', 1, 0)) * a.冲预交) As 医保统筹基金支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '其它医保支付', 1, 0)) * a.冲预交) As 其它医保支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '其它医保支付', 0, '个人账户支付', 0, '医保统筹基金支付', 0, 1)) *
                              a.冲预交) As 个人现金支付,
                         Max(Decode(Mod(a.记录性质, 10), 1, 0,
                                     Decode(c.开票结算方式, '其它医保支付', 结算号码, '个人账户支付', 结算号码, '医保统筹基金支付', 结算号码, ''))) As 医保结算号码,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(a.结算方式, '现金', 1, 0)) * a.冲预交) As 现金支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(a.结算方式, '支票', 1, 0)) * a.冲预交) As 支票支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(a.结算方式, '现金', 0, '支票', 0, 1) * a.冲预交)) As 转帐支付,
                         Sum(冲预交) As 结算总额
                        From 病人预交记录 A, Table(l_结帐id) B, 开票结算对照 C
                        Where a.结帐id = b.Column_Value And a.结算方式 = c.结算方式(+)))
    
     Loop
      --accountPay  个人账户支付  Number  14,2  是  按政策规定用个人账户支付参保人的医疗费用（含基本医疗保险目录范围内和目录范围外的费用）；
      --          如无金额，填写0
      v_费用 := ',"accountPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_结算.个人帐户支付, 0), 1);
      --fundPay  医保统筹基金支付  Number  14,2  是  患者本次就医所发生的医疗费用中按规定由基本医疗保险统筹基金支付的金额；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"fundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_结算.医保统筹基金支付, 0), 1);
      --otherfundPay  其它医保支付  Number  14,2  是  患者本次就医所发生的医疗费用中按规定由大病保险、医疗救助、公务员医疗补助、大额补充、企业补充等基金或资金支付的金额；
      v_费用 := v_费用 || ',"otherfundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_结算.其它医保支付, 0), 1);
      --ownPay  自费金额  Number  14,2  是  患者本次就医所发生的医疗费用中按照有关规定不属于基本医疗保险目录范围而全部由个人支付的费用；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"ownPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.自费金额, 1);
      --selfConceitedAmt  个人自负  Number  14,2  是  医保患者起付标准内个人支付费用；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"selfConceitedAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfPayAmt  个人自付  Number  14,2  是  患者本次就医所发生的医疗费用中由个人负担的属于基本医疗保险目录范围内自付部分的金额；开展按病种、病组、床日等打包付费方式且由患者定额付费的费用。该项为个人所得税大病医疗专项附加扣除信；息项如无金额，填写0
      v_费用 := v_费用 || ',"selfPayAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfCashPay  个人现金支付  Number  14,2  是  个人通过现金、银行卡、微信、支付宝等渠道支付的金额；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"selfCashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.个人现金支付, 1);
      --以后可能涉及冲预交,暂保留
      --cashPay  现金预交款金额  Number  14,2  否
      --v_费用 := v_费用 || ',"cashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.现金预交, 1);
      --chequePay  支票预交款金额  Number  14,2  否
      --v_费用 := v_费用 || ',"chequePay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.支票预交, 1);
      --transferAccountPay  转账预交款金额  Number  14,2  否
      --v_费用 := v_费用 || ',"transferAccountPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.转账预交, 1);
      --cashRecharge  补交金额(现金)  Number  14,2  否
      --v_费用 := v_费用 || ',"cashRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.现金支付, 1);
      --chequeRecharge  补交金额(支票)  Number  14,2  否
      --v_费用 := v_费用 || ',"chequeRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.支票支付, 1);
      --transferRecharge  补交金额（转账）  Number  14,2  否
      --v_费用 := v_费用 || ',"transferRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.转帐支付, 1);
      --cashRefund  退还金额(现金)  Number  14,2  否
      --v_费用 := v_费用 || ',"cashRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.现金退款, 1);
      --chequeRefund  退交金额(支票)  Number  14,2  否
      --v_费用 := v_费用 || ',"chequeRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.支票退款, 1);
      --transferRefund  退交金额(转账)  Number  14,2  否
      --v_费用 := v_费用 || ',"transferRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.转帐退款, 1);
      --ownAcBalance  个人账户余额  Number  14,2  否
      v_费用 := v_费用 || ',"ownAcBalance":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.个人帐户余额, 1);
      --reimbursementAmt  报销总金额  Number  14,2  否  医保结算后返回的总金额
      v_费用 := v_费用 || ',"reimbursementAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.报销总额, 1);
      --balancedNumber  结算号  String  100  否  医保结算后生成的号码/入账唯一值
      v_费用 := v_费用 || ',"balancedNumber":"' || b_Einvoice_Request_Test.Zljsonstr(c_结算.医保结算号码) || '"';
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --交费渠道
    v_缴费渠道 := Null;
    For c_渠道 In (Select /*+cardinality(b,10)*/
                  Nvl(c.渠道编码, Nvl(d.渠道编码, '-')) As 渠道编码, Sum(冲预交) As 结算总额
                 From 病人预交记录 A, Table(l_结帐id) B, 收费渠道对照 C,
                      (Select 结算方式, 渠道编码 From 收费渠道对照 D Where 卡类别id Is Null) D
                 Where a.结帐id = b.Column_Value And a.卡类别id = c.卡类别id(+) And a.结算方式 = c.结算方式(+) And a.结算方式 = d.结算方式(+)
                 Group By Nvl(c.渠道编码, Nvl(d.渠道编码, '-'))
                 Order By 渠道编码)
    
     Loop
      --payChannelCode  交费渠道编码  String  10  是
      If v_缴费渠道 Is Null Then
        v_缴费渠道 := Nvl(v_缴费渠道, '') || '{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_渠道.渠道编码, 0)) || '"';
      Else
        v_缴费渠道 := v_缴费渠道 || ',{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_渠道.渠道编码, 0)) || '"';
      End If;
      --payChannelValue  交费渠道金额  Number  14,2  是
      v_缴费渠道 := v_缴费渠道 || ',"payChannelValue":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_渠道.结算总额, 0), 1) || '}';
    End Loop;
  
    If v_缴费渠道 Is Not Null Then
      --payChannelDetail  交费渠道列表  String  不限  是  直接填写业务系统内部编码值，由医疗平台配置对照如：附录4交费渠道列表
      --        详见A-5,JSON格式列表
      v_缴费渠道 := ',"payChannelDetail":[' || v_缴费渠道 || ']';
    Else
      v_缴费渠道 := ',"payChannelDetail":[]';
    End If;
  
    -------------------------------------------------------------------------------------------
    --其他医保信息
    v_其它医保信息 := Null;
    --otherMedicalList  其它医保信息列表  String  不限  否  填写其它未知医保信息（在电子票据上以内容拼接方式显示）
    --            详见A-4,JSON格式列表
    --  infoNo  序号  Integer  不限  是  默认从1开始，每项数据序号值递增1，本次不允许重复
    --  infoName  医保信息名称  String  100  是  如费用报销类型编码，可参考附录7医保报销类型列表
    --  infoValue  医保信息值  String  100  是  如费用报销金额
    --  infoOther  医保其它信息  String  100  否  如医保报销比例。
  
    -------------------------------------------------------------------------------------------
    --其它扩展信息
    v_其它扩展信息 := Null;
    --otherInfo  其它扩展信息列表  String  不限  否  填写信息需要在电子票据上单独显示的其它扩展信息（未知信息）
    --          详见A-3,JSON格式列表
    --  infoNo  序号  Integer  不限  是  默认从1开始，每项数据序号值递增1，本次不允许重复
    --  infoName  扩展信息名称  String  100  是
    --  infoValue  扩展信息值  String  500  是
  
    c_交易信息 := To_Clob('{' || v_票据信息);
    c_交易信息 := c_交易信息 || To_Clob(v_缴费);
    c_交易信息 := c_交易信息 || To_Clob(v_通知);
    c_交易信息 := c_交易信息 || To_Clob(v_就诊信息);
    c_交易信息 := c_交易信息 || To_Clob(v_费用);
    c_交易信息 := c_交易信息 || To_Clob(v_缴费渠道);
  
    If v_其它扩展信息 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_其它扩展信息);
    End If;
    If v_其它医保信息 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_其它医保信息);
    End If;
    --  eBillRelateNo  业务票据关联号  String  32  否  如一笔业务数据需要开具N张电子票据，则N张电子票对应该值保持一致，用于后期关联查询
    --isArrears  是否可流通  String  1  是  0-否、1-是（如欠费情况根据医院业务要求该票据是否可流通）
    --arrearsReason  不可流通原因  String  200  否  isArrears=0，填写不可流通的原因
  
    c_交易信息 := c_交易信息 || To_Clob(',"eBillRelateNo":"","isArrears":"1","arrearsReason":""');
    If v_分类明细 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_分类明细);
    Else
      c_交易信息 := c_交易信息 || c_分类明细;
    End If;
  
    If v_明细 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_明细);
    Else
      c_交易信息 := c_交易信息 || c_明细;
    End If;
    c_交易信息  := c_交易信息 || To_Clob('}');
    Reqdata_Out := c_交易信息;
    Code_Out    := 1;
  Exception
    When Others Then
      Message_Out := SQLCode || ':' || SQLErrM;
      Code_Out    := 0;
  End Get_Registerdata_Create;

  Procedure Get_Zybalancedata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  ) Is
    --
    ---------------------------------------------------------------------------
    --功能:获取住院结帐开票数据
    --入参:
    --    Json_In,格式如下
    --  input
    --    occasion N 1  应用场合:1-收费,2-预交,3-结帐,4-挂号;5-就诊卡，固定传3
    --    einvoice_id  N,1 当前电子票据ID
    --    balance_id N 1  结算ID  occasion=2时，预交id;occasion<>2表示结帐id
    --    writeoff_id  N 1  冲销ID  occasion=2时，冲销预交id;occasion<>2表示冲销id
    --出参:
    --  ReqData_Out-返回的业务请求数据
    --  Totalmoney_Out-票据总额
    --  Code_Out-获取是否成功：0-失败；1-成功
    --  Message_Out 错误信息
    ---------------------------------------------------------------------------
    j_Input PLJson;
    j_Json  PLJson;
  
    n_应用场合 Number(2);
    n_结算id   病人预交记录.结帐id%Type;
  
    v_业务流水号 Varchar2(50);
  
    v_开票点       Varchar2(100);
    v_缴费         Varchar2(32767);
    v_票据信息     Varchar2(32767);
    v_就诊信息     Varchar2(32767);
    v_通知         Varchar2(32767);
    v_缴费渠道     Varchar2(32767);
    v_费用         Varchar2(32767);
    v_其它扩展信息 Varchar2(32767);
    v_其它医保信息 Varchar2(32767);
    c_明细         Clob;
    v_明细         Varchar2(32767);
    c_分类明细     Clob;
    v_分类明细     Varchar2(32767);
    c_交易信息     Clob; --最终返回的交易信息集
  
    c_预交 Clob;
    v_预交 Varchar2(32767);
  
    n_病人id 病人预交记录.病人id%Type;
  
    n_缺省卡类别id     Number(18);
    v_参数值           Varchar2(100);
    n_票据总金额       门诊费用记录.结帐金额%Type;
    n_误差总额         门诊费用记录.结帐金额%Type;
    n_用户id           人员表.Id%Type;
    v_操作员编号       人员表.编号%Type;
    v_操作员姓名       人员表.姓名%Type;
    v_Temp             Varchar2(32767);
    n_险类             保险结算记录.险类%Type;
    v_保险机构编码     保险类别.保险机构编码%Type;
    v_医疗付款方式编码 医疗付款方式.编码%Type;
    v_病种名称         保险病种.名称%Type;
    v_医保号           保险帐户.医保号%Type;
    v_版本号           Varchar2(30);
    v_住院次数         Varchar2(4000);
    n_电子票据id       电子票据使用记录.Id%Type;
    Cursor c_Balance_Record Is
      Select a.No, a.收费时间, a.结帐类型, a.操作员编号, a.操作员姓名, a.病人id, a.主页id,
             Decode(Nvl(a.病人id, 0), 0, a.原因, Nvl(b.姓名, c.姓名)) As 姓名, Nvl(b.性别, c.性别) As 性别, Nvl(b.年龄, c.年龄) As 年龄, c.门诊号,
             Nvl(b.住院号, c.住院号) As 住院号, a.开始日期, a.结束日期, a.备注, a.结帐金额, Decode(Nvl(a.病人id, 0), 0, q.电子邮件, c.Email) As Email,
             q.联系人, Decode(Nvl(a.病人id, 0), 0, q.社会信用代码, c.身份证号) As 身份证号,
             Decode(Nvl(a.病人id, 0), 0, Nvl(q.电话, To_Char(j.移动电话)), c.手机号) As 手机号,
             Decode(Nvl(a.病人id, 0), 0, 2, 1) As 缴款类型, Decode(Nvl(a.结帐类型, 0), 1, '02', '01') As 业务标识, b.入院日期, b.出院日期,
             m.编码 As 入院科室编码, m.名称 As 入院科室名称, p.编码 As 出院科室编码, p.名称 As 出院科室名称, b.出院病床 As 床号, t.名称 As 病区名称,
             Nvl(b.病案号, b.住院号) As 病历号, Nvl(b.医疗付款方式, c.医疗付款方式) As 医疗付款方式, Nvl(b.出院日期, Sysdate) - b.入院日期 As 住院天数
      From 病人结帐记录 A, 病案主页 B, 病人信息 C, 合约单位 Q, 人员表 J, 部门表 M, 部门表 P, 部门表 T
      Where a.Id = n_结算id And a.病人id = b.病人id(+) And a.主页id = b.主页id(+) And a.病人id = c.病人id(+) And a.原因 = q.名称(+) And
            b.入院科室id = m.Id(+) And b.出院科室id = p.Id(+) And b.当前病区id = t.Id(+)
           
            And q.联系人 = j.姓名(+);
    r_Balance_Record c_Balance_Record%RowType;
  
  Begin
    j_Input := PLJson(Json_In);
    j_Json  := j_Input.Get_Pljson('input');
  
    n_应用场合 := Nvl(j_Json.Get_Number('occasion'), 0);
    n_结算id   := j_Json.Get_Number('balance_id');
    --n_冲销id     := Nvl(j_Json.Get_Number('writeoff_id'), 0);
    n_电子票据id := Nvl(j_Json.Get_Number('einvoice_id'), 0);
  
    If Nvl(n_应用场合, 0) = 0 Then
      Code_Out    := 0;
      Message_Out := '无效的应用场景';
      Return;
    End If;
  
    Select Nvl(Max(参数值), 'V2.0.3')
    Into v_版本号
    From 三方接口配置
    Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '支持版本';
  
    b_Einvoice_Request_Test.Get_Identity(n_用户id, v_操作员编号, v_操作员姓名);
  
    --v_业务标识:01  住院,02  门诊,03  急诊,04  门特,05  体检中心,06  挂号,07  住院预交金,08  体检预交金
  
    n_票据总金额 := 0;
    c_明细       := Null;
    v_明细       := Null;
  
    Open c_Balance_Record;
    Fetch c_Balance_Record
      Into r_Balance_Record;
    If c_Balance_Record%NotFound Then
      Code_Out    := 0;
      Message_Out := '未找到指定的结算数据';
      Return;
    End If;
  
    v_住院次数 := Null;
    For c_收费细目 In (Select Min(a.Id) As 费用id, a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号) As 序号, a.收费细目id, Max(a.计算单位) As 计算单位,
                          Sum(a.标准单价) As 价格, Avg(Nvl(a.付数, 1) * Nvl(a.数次, 0)) As 数量, Sum(a.应收金额) As 应收金额,
                          Sum(a.实收金额) As 实收金额, Sum(a.结帐金额) As 结帐金额, Sum(a.实收金额) - Sum(a.统筹金额) As 自费金额,
                          Max(t.编码) As 医保项目编码, Max(t.名称) As 医保项目名称, Max(t.统筹比额) As 医保报销比例, Max(a.摘要) As 备注,
                          Max(a.费用类型) As 费用类型, Max(a.操作员编号) As 操作员编号, Max(a.操作员姓名) As 操作员姓名, Max(a.病人id) As 病人id,
                          Max(a.登记时间) As 登记时间, Max(a.收据费目) As 收据费目, Max(c.编码) As 收据费目编码, Max(a.主页id) As 主页id,
                          Max(d.编码) As 类别编码, Max(d.类别) As 类别名称, Max(b.编码) As 项目编码, Max(b.名称) As 项目名称, Max(b.规格) As 规格,
                          Max(q.药品剂型) As 药品剂型
                   From 住院费用记录 A, 收费项目目录 B, 收据费目 C, 收费类别 D, 药品规格 M, 药品特性 Q, 诊疗项目目录 J, 保险支付大类 T, 支付类别对照 S
                   Where a.结帐id = n_结算id And a.记帐费用 = 1 And a.收费类别 = d.编码(+) And a.收费细目id = b.Id And a.收据费目 = c.名称(+) And
                         a.收费细目id = m.药品id(+) And m.药名id = q.药名id(+) And q.药名id = j.Id(+) And a.保险大类id = t.Id(+) And
                         t.性质(+) = 1 And a.保险大类id = s.保险大类id(+)
                   Group By a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号), a.收费细目id, c.编码, c.名称, j.编码, j.名称
                   Order By NO, 序号) Loop
    
      --listDetailNo  明细流水号  String  60  否  明细流水号
      v_Temp := '{' || '"listDetailNo":"' || b_Einvoice_Request_Test.Zljsonstr(LPad(c_收费细目.费用id, 20, '0')) || '"';
      --chargeCode  收费项目代码  String  50  否  填写业务系统内部编码值，由医疗平台配置对照,如：床位费、检查费
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.收据费目编码) || '"';
      --chargeName  收费项目名称  String  100  否
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.收据费目) || '"';
      --prescribeCode  处方编码  String  60  否
      v_Temp := v_Temp || ',' || '"prescribeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.No) || '"';
      --listTypeCode  药品类别编码  String  50  否  如药品分类编码01，有则填写
      v_Temp := v_Temp || ',' || '"listTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.类别编码) || '"';
      --listTypeName  药品类别名称  String  50  否  如药品分类名称，抗生素类抗感染药物
      v_Temp := v_Temp || ',' || '"listTypeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.类别名称) || '"';
      --code  编码  String  50  否  如药品编码，有则填写
      v_Temp := v_Temp || ',' || '"code":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.项目编码) || '"';
      --name  药品名称  String  50  是  如药品名称，器材名称等
      v_Temp := v_Temp || ',' || '"name":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.项目名称) || '"';
      --form  剂型  String  50  否
      v_Temp := v_Temp || ',' || '"form":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.药品剂型) || '"';
      --specification  规格  String  50  否
      v_Temp := v_Temp || ',' || '"specification":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.规格) || '"';
      --unit  计量单位   String  20  否
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.计算单位) || '"';
      --std  单价  Number  14,6  是
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.价格, 1);
      --number  数量  Number  14,6  是
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.数量, 1);
      --amt  金额  Number  14,6  是
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.实收金额, 1);
      --selfAmt  自费金额  Number  14,6  是  如无金额，填写0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.自费金额, 1);
      --receivableAmt  应收费用  Number  14,6  否
      v_Temp := v_Temp || ',' || '"receivableAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.应收金额, 1);
      --medicalCareType  医保药品分类  String  1  否  1：无自负/甲
      --          2：有自负/乙
      --          3：全自负/丙
      v_Temp := v_Temp || ',' || '"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.医保项目编码) || '"';
      --medCareItemType  医保项目类型  String  100  否
      v_Temp := v_Temp || ',' || '"medCareItemType":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.医保项目名称) || '"';
      --medReimburseRate  医保报销比例  Number  3,2  否
      v_Temp := v_Temp || ',' || '"medReimburseRate":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.医保报销比例, 1);
      --remark  备注  String  200  否
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.备注) || '"';
      --sortNo  序号  Integer  不限  否  序号
      v_Temp := v_Temp || ',' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.序号, 1);
      --chrgtype  费用类型  String  50  否
      v_Temp := v_Temp || ',' || '"chrgtype":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.费用类型) || '"}';
    
      If Instr(Nvl(v_住院次数, '') || ',', ',' || Nvl(c_收费细目.主页id, 0) || ',') = 0 Then
        v_住院次数 := Nvl(v_住院次数, '') || ',' || Nvl(c_收费细目.主页id, 0);
      End If;
      If Length(Nvl(v_明细, '') || v_Temp) > 32700 Then
        If c_明细 Is Null Then
          c_明细 := To_Clob(v_明细);
        Else
          c_明细 := c_明细 || To_Clob(',' || v_明细);
        End If;
        v_明细 := Null;
      End If;
    
      If v_明细 Is Null Then
        v_明细 := v_Temp;
      Else
        v_明细 := v_明细 || ',' || v_Temp;
      End If;
    End Loop;
    If v_住院次数 Is Not Null Then
      v_住院次数 := Substr(v_住院次数, 2);
    End If;
    If v_明细 Is Not Null And c_明细 Is Not Null Then
      --listDetail  清单项目明细  String  不限  是  详见A-2,JSON格式列表
      c_明细 := c_明细 || ',' || To_Clob(v_明细);
      c_明细 := To_Clob(',"listDetail":[') || c_明细 || To_Clob(']');
    
      v_明细 := Null;
    Elsif v_明细 Is Not Null Then
      v_明细 := ',"listDetail":[' || v_明细 || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --分类明细
    v_分类明细 := Null;
    c_分类明细 := Null;
    For c_分类统计 In (Select Rownum As 序号, 收据费目编码, 收据费目名称, 数量, 计算单位, Round(单价, 2) As 单价, Round(结帐金额, 2) As 结帐金额,
                          Round(自费金额, 2) As 自费金额, 备注, 结帐金额 - Round(结帐金额, 2) As 误差费
                   From (Select /*+cardinality(b,10)*/
                           c.编码 As 收据费目编码, a.收据费目 As 收据费目名称, 1 As 数量, '' As 计算单位, Sum(a.结帐金额) As 单价, a.收据费目,
                           Sum(a.结帐金额) As 结帐金额, Sum(a.结帐金额) - Sum(a.统筹金额) As 自费金额, '' As 备注
                          From 住院费用记录 A, 收据费目 C
                          Where a.结帐id = n_结算id And a.收据费目 = c.名称(+)
                          Group By c.编码, a.收据费目)) Loop
      --sortNo  序号  Integer  不限  是  默认从1开始，每个收费项目序号值递增1，本次不允许重复
      v_Temp := '{' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.序号, 1);
      --chargeCode  收费项目代码  String  50  是  填写业务系统内部编码值，由医疗平台配置对照
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.收据费目编码) || '"';
      --chargeName  收费项目名称  String  100  是  填写业务系统内部项目名称
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.收据费目名称) || '"';
      --unit  计量单位  String  20  否
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.计算单位) || '"';
      --std  收费标准  Number  14,2  是
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.单价, 1);
      --number  数量  Number  14,2  是
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.数量, 1);
      --amt  金额  Number  14,2  是
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.结帐金额, 1);
      --selfAmt  自费金额  Number  14,2  是  如无金额，填写0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.自费金额, 1);
      --remark  备注  String  200  否
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.备注) || '"}';
    
      If Length(Nvl(v_分类明细, '') || v_Temp) > 32700 Then
        If c_分类明细 Is Null Then
          c_分类明细 := To_Clob(v_分类明细);
        Else
          c_分类明细 := c_分类明细 || To_Clob(',' || v_分类明细);
        End If;
        v_分类明细 := Null;
      End If;
    
      If v_分类明细 Is Null Then
        v_分类明细 := v_Temp;
      Else
        v_分类明细 := v_分类明细 || ',' || v_Temp;
      End If;
    
      n_票据总金额 := Nvl(n_票据总金额, 0) + Nvl(c_分类统计.结帐金额, 0);
      n_误差总额   := Nvl(n_误差总额, 0) + Nvl(c_分类统计.误差费, 0);
    End Loop;
    Totalmoney_Out := n_票据总金额;
    If v_分类明细 Is Not Null And c_分类明细 Is Not Null Then
      c_分类明细 := c_分类明细 || ',' || To_Clob(v_分类明细);
      --chargeDetail  chargeDetail  收费项目明细  收费项目明细
      c_分类明细 := To_Clob(',"chargeDetail":[') || c_分类明细 || To_Clob(']');
      v_分类明细 := Null;
    Elsif v_分类明细 Is Not Null Then
      v_分类明细 := ',"chargeDetail":[' || v_分类明细 || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --票据信息
    --Select Count(1) Into n_作废次数 From 电子票据使用记录 Where 票种 = n_应用场合 And 结算id = n_结算id;
    --结算ID||_||电子票据ID
    v_业务流水号 := n_结算id || '_' || Nvl(n_电子票据id, 0);
    v_开票点     := b_Einvoice_Request_Test.Get_Einvoice_Node(v_操作员编号, v_操作员姓名, n_用户id);
  
    v_票据信息 := '"busNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_业务流水号) || '"'; --业务流水号
    v_票据信息 := v_票据信息 || ',"busType":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.业务标识) || '"'; --业务标识
    v_票据信息 := v_票据信息 || ',"payer":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.姓名) || '"'; --患者姓名
    v_票据信息 := v_票据信息 || ',"busDateTime":"' ||
              b_Einvoice_Request_Test.Zljsonstr(To_Char(r_Balance_Record.收费时间, 'yyyymmddHH24miss') || '000') || '"';
    v_票据信息 := v_票据信息 || ',"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_开票点) || '"'; --开票点编码:直接填写业务系统内部编码值，由医疗平台配置对照
    v_票据信息 := v_票据信息 || ',"payee":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.操作员姓名) || '"'; --收费员
  
    v_票据信息 := v_票据信息 || ',"author":"' || b_Einvoice_Request_Test.Zljsonstr(v_操作员姓名) || '"'; --票据编制人
    v_票据信息 := v_票据信息 || ',"totalAmt":' || b_Einvoice_Request_Test.Zljsonstr(n_票据总金额, 1); --开票总金额
    v_票据信息 := v_票据信息 || ',"remark":"' || Nvl(r_Balance_Record.备注, '') || '"'; --备注
    -------------------------------------------------------------------------------------------
  
    --取缴费信息
    v_缴费 := Null;
    For c_缴费 In (Select Max(Decode(信息名, '订单号', 信息值, '支付订单号', 信息值, '')) As 支付定单号,
                        Max(Decode(信息名, '医保支付订单号', 信息值, '医保订单号', 信息值, '')) As 医保支付定单号,
                        Max(Decode(Upper(信息名), '支付宝公众号USERID', 信息值, '')) As 支付宝公众号userid,
                        Max(Decode(Upper(信息名), '支付宝小程序USERID', 信息值, '')) As 支付宝小程序userid,
                        Max(Decode(Upper(信息名), '微信公众号OPENID', 信息值, '')) As 微信公众号openid,
                        Max(Decode(Upper(信息名), '微信小程序OPENID', 信息值, '')) As 微信小程序openid
                 From (Select 信息名, 信息值
                        From 病人信息从表
                        Where 病人id = n_病人id And 信息名 In ('支付宝公众号USERID', '支付宝小程序USERID', '微信公众号OPENID', '微信小程序OPENID')
                        Union All
                        Select 交易项目, 交易内容
                        From 三方结算交易
                        Where 交易id In (Select ID From 病人预交记录 Where 结帐id = n_结算id) And 交易项目 Like '%订单号')) Loop
      v_缴费 := ',"alipayCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.支付宝公众号userid) || '"'; --患者支付宝账户
      v_缴费 := v_缴费 || ',"weChatOrderNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.支付定单号) || '"'; --微信支付订单号
      If v_版本号 = 'V3.1.0' Then
        --该版本才有此接点
        v_缴费 := v_缴费 || ',"weChatMedTransNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.医保支付定单号) || '"'; --微信医保支付订单号
      End If;
    
      If c_缴费.微信公众号openid Is Not Null Then
        v_缴费 := v_缴费 || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.微信公众号openid) || '"'; --微信公众号或小程序用户ID
      Else
        v_缴费 := v_缴费 || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.微信小程序openid) || '"'; --微信公众号或小程序用户ID
      End If;
      Exit;
    End Loop;
  
    -------------------------------------------------------------------------------------------
    --取通知信息
    Select To_Number(Max(参数值))
    Into n_缺省卡类别id
    From 三方接口配置
    Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '缺省卡类别ID';
  
    v_通知 := ',"tel":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.手机号) || '"'; --患者手机号码
    v_通知 := v_通知 || ',"email":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.Email) || '"'; --患者邮箱地址
    If v_版本号 = 'V3.1.0' Then
      --该版本才有此接点
      v_通知 := v_通知 || ',"payerType":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.缴款类型) || '"'; --交款人类型
    End If;
    v_通知 := v_通知 || ',"idCardNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.身份证号) || '"'; --统一社会信用代码
  
    If Nvl(r_Balance_Record.病人id, 0) = 0 Then
      --没有一张卡，固定一种卡类别
      Select Nvl(Max(参数值), '99999')
      Into v_参数值
      From 三方接口配置
      Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '病人无卡的卡类别编号';
      v_通知 := v_通知 || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡类型
    
      Select Nvl(Max(参数值), '-')
      Into v_参数值
      From 三方接口配置
      Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '病人无卡的卡号';
      v_通知 := v_通知 || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡号
    
    Else
      v_Temp := Null;
    
      For c_通知 In (Select 名称, 编码, 卡号
                   From (Select b.病人id, c.名称, c.编码, b.卡号, Decode(b.卡类别id, n_缺省卡类别id, 2, c.缺省标志) As 缺省标志
                          From 病人医疗卡信息 B, 医疗卡类别 C
                          Where b.卡类别id = c.Id And b.病人id = r_Balance_Record.病人id
                          Order By 缺省标志)
                   Where Rownum < 2) Loop
      
        If c_通知.名称 Is Not Null Then
        
          v_Temp := ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.名称) || '"'; --卡类型
          v_Temp := v_Temp || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.卡号) || '"'; --卡号
        End If;
        Exit;
      End Loop;
      If v_Temp Is Not Null Then
        v_通知 := v_通知 || v_Temp;
      Else
        If r_Balance_Record.身份证号 Is Not Null Then
          Select Nvl(Max(参数值), '99998')
          Into v_参数值
          From 三方接口配置
          Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '身份证作卡类型编号';
          v_通知 := v_通知 || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡类型
          v_通知 := v_通知 || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.身份证号) || '"'; --卡号
        Else
          --没有一张卡，固定一种卡类别
          Select Nvl(Max(参数值), '99999')
          Into v_参数值
          From 三方接口配置
          Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '病人无卡的卡类别编号';
          v_通知 := v_通知 || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡类型
          Select Nvl(Max(参数值), '-')
          Into v_参数值
          From 三方接口配置
          Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '病人无卡的卡号';
          v_通知 := v_通知 || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡号
        End If;
      End If;
    End If;
  
    -------------------------------------------------------------------------------------------
    --就诊信息
    Select Max(内容) Into v_Temp From zlRegInfo Where 项目 = '医疗机构类型';
  
    --性质:1-收费;2-结算（包括住院结算、特殊门诊结算）；3-预交
    Select Max(a.险类), Max(b.保险机构编码), Max(Nvl(a.病种名称, c.名称))
    Into n_险类, v_保险机构编码, v_病种名称
    From 保险结算记录 A, 保险类别 B, 保险病种 C
    Where a.险类 = b.序号 And a.病种id = c.Id(+) And a.记录id = n_结算id And a.性质 = 2;
  
    If Nvl(n_险类, 0) <> 0 Then
      Select Max(医保号) Into v_医保号 From 保险帐户 Where 病人id = n_病人id And 险类 = n_险类;
    End If;
    Select Max(编码) Into v_医疗付款方式编码 From 医疗付款方式 Where 名称 = Nvl(r_Balance_Record.医疗付款方式, '-');
  
    --medicalInstitution  医疗机构类型  String  60  否  按照《医疗机构管理条例实施细则》和《卫生部关于修订<医疗机构管理条例实施细则>第三条有关内容的通知》确定的医疗卫生机构类别
    v_就诊信息 := ',"medicalInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_Temp) || '"';
    --medCareInstitution  医保机构编码  String  60  否  医保机构的唯一编码
    v_就诊信息 := v_就诊信息 || ',"medCareInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_保险机构编码) || '"';
    --medCareTypeCode  医保类型编码  String  60  否
    v_就诊信息 := v_就诊信息 || ',"medCareTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_医疗付款方式编码) || '"';
    --medicalCareType  医保类型名称  String  60  否  由城镇职工基本医疗保险、城镇居民基本医疗保险、新型农村合作医疗、其它医疗保险等构成
    v_就诊信息 := v_就诊信息 || ',"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.医疗付款方式) || '"';
    --medicalInsuranceID  患者医保编号  String  60  否  参保人在医保系统中的唯一标识(医保号)
    v_就诊信息 := v_就诊信息 || ',"medicalInsuranceID":"' || b_Einvoice_Request_Test.Zljsonstr(v_医保号) || '"';
    --category  入院科室名称  String  60  否
    v_就诊信息 := v_就诊信息 || ',"category":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.入院科室名称) || '"';
    --categoryCode  入院科室编码  String  60  否
    v_就诊信息 := v_就诊信息 || ',"categoryCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.入院科室编码) || '"';
    --leaveCategory  出院科室名称  String  60  否
    v_就诊信息 := v_就诊信息 || ',"leaveCategory":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.出院科室名称) || '"';
    --leaveCategoryCode  出院科室编码  String  60  否
    v_就诊信息 := v_就诊信息 || ',"leaveCategoryCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.出院科室编码) || '"';
    --hospitalNo  患者住院号  String  20  是  从入院到出院结束后，整个流程的唯一号
    v_就诊信息 := v_就诊信息 || ',"hospitalNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.住院号) || '"';
    --visitNo  住院就诊编号  String  20  是  住院期间，存在多次结算，结算后会重新生成一个住院就诊编号，如无就诊编号，可等于患者住院号
    v_就诊信息 := v_就诊信息 || ',"visitNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.住院号) || '"';
    --consultationDate  就诊日期  String  10  否  患者就医时间
    --          格式:yyyy-MM-dd
    v_就诊信息 := v_就诊信息 || ',"consultationDate":"' ||
              b_Einvoice_Request_Test.Zljsonstr(To_Char(r_Balance_Record.入院日期, 'yyyy-mm-dd')) || '"';
    --patientId  患者唯一ID  String  50  否  患者在业务系统中的唯一标识ID，类似身份证号码。
    v_就诊信息 := v_就诊信息 || ',"patientId":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.病人id) || '"';
    --patientNo  患者就诊编号  String  20  否  患者每次就诊一次就生成的一个新的编号。（患者登记号）
    v_就诊信息 := v_就诊信息 || ',"patientNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.主页id) || '"';
    --sex  性别  String  2  是
    v_就诊信息 := v_就诊信息 || ',"sex":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.性别) || '"';
    --age  年龄  String  10  是
    v_就诊信息 := v_就诊信息 || ',"age":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.年龄) || '"';
    --hospitalArea  病区  String  60  否
    v_就诊信息 := v_就诊信息 || ',"hospitalArea":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.病区名称) || '"';
    --bedNo  床号  String  20  否
    v_就诊信息 := v_就诊信息 || ',"bedNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.床号) || '"';
    --caseNumber  病历号  String  50  否
    v_就诊信息 := v_就诊信息 || ',"caseNumber":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.病历号) || '"';
  
    If Instr(v_住院次数, ',') > 0 Then
      --本次结算多次住院的
      For c_主页 In (Select Min(入院日期) As 入院日期, Max(出院日期) As 出院日期, Sum(Nvl(出院日期, Sysdate) - 入院日期) As 住院天数
                   From 病案主页
                   Where 病人id = r_Balance_Record.病人id And
                         主页id In (Select /*+cardinality(A,10)*/
                                   Column_Value
                                  From Table(f_Num2List(v_住院次数)) A)) Loop
      
        --inHospitalDate  住院日期  String  10  否  格式:yyyy-MM-dd
        v_就诊信息 := v_就诊信息 || ',"inHospitalDate":"' || b_Einvoice_Request_Test.Zljsonstr(To_Char(c_主页.入院日期, 'yyyy-mm-dd')) || '"';
        --outHospitalDate  出院日期  String  10  否  格式:yyyy-MM-dd
        v_就诊信息 := v_就诊信息 || ',"outHospitalDate":"' || b_Einvoice_Request_Test.Zljsonstr(To_Char(c_主页.出院日期, 'yyyy-mm-dd')) || '"';
        --hospitalDays  住院天数  Number  6,2  否
        v_就诊信息 := v_就诊信息 || ',"hospitalDays":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_主页.住院天数, 0), 1);
        Exit;
      
      End Loop;
    Else
      --inHospitalDate  住院日期  String  10  否  格式:yyyy-MM-dd
      v_就诊信息 := v_就诊信息 || ',"inHospitalDate":"' ||
                b_Einvoice_Request_Test.Zljsonstr(To_Char(r_Balance_Record.入院日期, 'yyyy-mm-dd')) || '"';
      --outHospitalDate  出院日期  String  10  否  格式:yyyy-MM-dd
      v_就诊信息 := v_就诊信息 || ',"outHospitalDate":"' ||
                b_Einvoice_Request_Test.Zljsonstr(To_Char(r_Balance_Record.出院日期, 'yyyy-mm-dd')) || '"';
      --hospitalDays  住院天数  Number  6,2  否
      v_就诊信息 := v_就诊信息 || ',"hospitalDays":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(r_Balance_Record.住院天数, 0), 1);
    End If;
  
    -------------------------------------------------------------------------------------------
    --结算信息
    v_费用 := Null;
  
    --预交列表 
    For c_冲预交 In (
                  
                  Select q.凭证代码, q.凭证号码, a.No, Max(a.冲预交) As 冲预交
                  From (Select NO, Sum(冲预交) As 冲预交
                          From 病人预交记录
                          Where 结帐id = n_结算id And Mod(记录性质, 10) = 1) A, 病人预交记录 B, 电子票据使用记录 Q
                  Where a.No = b.No And b.记录性质 = 1 And b.Id = q.结算id And q.票种 = 2 And q.退款id Is Null) Loop
      --    voucherBatchCode  预交金凭证代码  String  50  是
      v_Temp := '{voucherBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_冲预交.凭证代码) || '"';
      --    voucherNo  预交金凭证号码  String  20  是
      v_Temp := v_Temp || ',"voucherNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_冲预交.凭证号码) || '"';
      --    voucherAmt  预交金凭证金额  Number  14,2  是  参与结算的金额
      --          注:如预全额结算，传入总金额；如部分金额结算，传入实际参与结算金额
      v_Temp := v_Temp || ',"voucherAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_冲预交.冲预交, 1) || '}';
    
      If Length(Nvl(v_预交, '') || v_Temp) > 32700 Then
        If v_预交 Is Null Then
          c_预交 := To_Clob(v_分类明细);
        Else
          c_预交 := c_预交 || To_Clob(',' || v_预交);
        End If;
        v_预交 := Null;
      End If;
    
      If v_预交 Is Null Then
        v_预交 := v_Temp;
      Else
        v_预交 := v_预交 || ',' || v_Temp;
      End If;
    End Loop;
  
    If v_预交 Is Not Null And c_明细 Is Not Null Then
      --payMentVoucher  预交金凭证消费扣款列表  String  不限  否  结算开具住院电子票据时，传入消费扣款对应预交金凭证信息
      c_预交 := c_预交 || ',' || To_Clob(v_预交);
      c_预交 := To_Clob(',"payMentVoucher":[') || c_预交 || To_Clob(']');
    
      v_预交 := Null;
    Elsif v_预交 Is Not Null Then
      v_预交 := ',"payMentVoucher":[' || v_预交 || ']';
    End If;
  
    --    payMentVoucher  预交金凭证消费扣款列表  String  不限  否  结算开具住院电子票据时，传入消费扣款对应预交金凭证信息
    --          详见A-6,JSON格式列表
  
    For c_结算 In (Select 现金预交, 支票预交, 转账预交, 个人帐户支付, 医保统筹基金支付, 其它医保支付, 个人现金支付, Decode(Sign(现金支付), -1, 现金支付, 0) As 现金退款,
                        Decode(Sign(现金支付), -1, 支票支付, 0) As 支票退款, Decode(Sign(现金支付), -1, 转帐支付, 0) As 转帐退款,
                        Decode(Sign(现金支付), -1, 0, 现金支付) As 现金支付, Decode(Sign(现金支付), -1, 0, 支票支付) As 支票支付,
                        Decode(Sign(现金支付), -1, 0, 转帐支付) As 转帐支付,
                        Nvl(个人帐户支付, 0) + Nvl(医保统筹基金支付, 0) + Nvl(其它医保支付, 0) As 报销总额,
                        Nvl(结算总额, 0) - Nvl(个人帐户支付, 0) - Nvl(医保统筹基金支付, 0) - Nvl(其它医保支付, 0) As 自费金额, 结算总额, 医保结算号码,
                        0 As 个人帐户余额
                 From (Select /*+cardinality(b,10)*/
                         Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '现金', 1, 0), 0) * a.冲预交) As 现金预交,
                         Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '支票', 1, 0), 0) * a.冲预交) As 支票预交,
                         Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '支票', 0, '现金', 0, 1), 0) * a.冲预交) As 转账预交,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '个人账户支付', 1, 0)) * a.冲预交) As 个人帐户支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '医保统筹基金支付', 1, 0)) * a.冲预交) As 医保统筹基金支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '其它医保支付', 1, 0)) * a.冲预交) As 其它医保支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '其它医保支付', 0, '个人账户支付', 0, '医保统筹基金支付', 0, 1)) *
                              a.冲预交) As 个人现金支付,
                         Max(Decode(Mod(a.记录性质, 10), 1, 0,
                                     Decode(c.开票结算方式, '其它医保支付', 结算号码, '个人账户支付', 结算号码, '医保统筹基金支付', 结算号码, ''))) As 医保结算号码,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(a.结算方式, '现金', 1, 0)) * a.冲预交) As 现金支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(a.结算方式, '支票', 1, 0)) * a.冲预交) As 支票支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(a.结算方式, '现金', 0, '支票', 0, 1) * a.冲预交)) As 转帐支付,
                         Sum(冲预交) As 结算总额
                        From 病人预交记录 A, 开票结算对照 C
                        Where a.结帐id = n_结算id And a.结算方式 = c.结算方式(+))) Loop
      --accountPay  个人账户支付  Number  14,2  是  按政策规定用个人账户支付参保人的医疗费用（含基本医疗保险目录范围内和目录范围外的费用）；
      --          如无金额，填写0
      v_费用 := ',"accountPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_结算.个人帐户支付, 0), 1);
      --fundPay  医保统筹基金支付  Number  14,2  是  患者本次就医所发生的医疗费用中按规定由基本医疗保险统筹基金支付的金额；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"fundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_结算.医保统筹基金支付, 0), 1);
      --otherfundPay  其它医保支付  Number  14,2  是  患者本次就医所发生的医疗费用中按规定由大病保险、医疗救助、公务员医疗补助、大额补充、企业补充等基金或资金支付的金额；
      v_费用 := v_费用 || ',"otherfundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_结算.其它医保支付, 0), 1);
      --ownPay  自费金额  Number  14,2  是  患者本次就医所发生的医疗费用中按照有关规定不属于基本医疗保险目录范围而全部由个人支付的费用；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"ownPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.自费金额, 1);
      --selfConceitedAmt  个人自负  Number  14,2  是  医保患者起付标准内个人支付费用；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"selfConceitedAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfPayAmt  个人自付  Number  14,2  是  患者本次就医所发生的医疗费用中由个人负担的属于基本医疗保险目录范围内自付部分的金额；开展按病种、病组、床日等打包付费方式且由患者定额付费的费用。该项为个人所得税大病医疗专项附加扣除信；息项如无金额，填写0
      v_费用 := v_费用 || ',"selfPayAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfCashPay  个人现金支付  Number  14,2  是  个人通过现金、银行卡、微信、支付宝等渠道支付的金额；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"selfCashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.个人现金支付, 1);
      --cashPay  现金预交款金额  Number  14,2  否
      v_费用 := v_费用 || ',"cashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.现金预交, 1);
      --chequePay  支票预交款金额  Number  14,2  否
      v_费用 := v_费用 || ',"chequePay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.支票预交, 1);
      --transferAccountPay  转账预交款金额  Number  14,2  否
      v_费用 := v_费用 || ',"transferAccountPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.转账预交, 1);
      --cashRecharge  补交金额(现金)  Number  14,2  否
      v_费用 := v_费用 || ',"cashRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.现金支付, 1);
      --chequeRecharge  补交金额(支票)  Number  14,2  否
      v_费用 := v_费用 || ',"chequeRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.支票支付, 1);
      --transferRecharge  补交金额（转账）  Number  14,2  否
      v_费用 := v_费用 || ',"transferRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.转帐支付, 1);
      --cashRefund  退还金额(现金)  Number  14,2  否
      v_费用 := v_费用 || ',"cashRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.现金退款, 1);
      --chequeRefund  退交金额(支票)  Number  14,2  否
      v_费用 := v_费用 || ',"chequeRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.支票退款, 1);
      --transferRefund  退交金额(转账)  Number  14,2  否
      v_费用 := v_费用 || ',"transferRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.转帐退款, 1);
      --ownAcBalance  个人账户余额  Number  14,2  否
      v_费用 := v_费用 || ',"ownAcBalance":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.个人帐户余额, 1);
      --reimbursementAmt  报销总金额  Number  14,2  否  医保结算后返回的总金额
      v_费用 := v_费用 || ',"reimbursementAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.报销总额, 1);
      --balancedNumber  结算号  String  100  否  医保结算后生成的号码/入账唯一值
      v_费用 := v_费用 || ',"balancedNumber":"' || b_Einvoice_Request_Test.Zljsonstr(c_结算.医保结算号码) || '"';
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --交费渠道
    v_缴费渠道 := Null;
    For c_渠道 In (Select /*+cardinality(b,10)*/
                  Nvl(c.渠道编码, Nvl(d.渠道编码, '-')) As 渠道编码, Sum(冲预交) As 结算总额
                 From 病人预交记录 A, 收费渠道对照 C, (Select 结算方式, 渠道编码 From 收费渠道对照 D Where 卡类别id Is Null) D
                 Where a.结帐id = n_结算id And a.卡类别id = c.卡类别id(+) And a.结算方式 = c.结算方式(+) And a.结算方式 = d.结算方式(+)
                 Group By Nvl(c.渠道编码, Nvl(d.渠道编码, '-'))
                 Order By 渠道编码)
    
     Loop
      --payChannelCode  交费渠道编码  String  10  是
      If v_缴费渠道 Is Null Then
        v_缴费渠道 := Nvl(v_缴费渠道, '') || '{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_渠道.渠道编码, 0)) || '"';
      Else
        v_缴费渠道 := v_缴费渠道 || ',{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_渠道.渠道编码, 0)) || '"';
      End If;
      --payChannelValue  交费渠道金额  Number  14,2  是
      v_缴费渠道 := v_缴费渠道 || ',"payChannelValue":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_渠道.结算总额, 0), 1) || '}';
    End Loop;
  
    If v_缴费渠道 Is Not Null Then
      --payChannelDetail  交费渠道列表  String  不限  是  直接填写业务系统内部编码值，由医疗平台配置对照如：附录4交费渠道列表
      --        详见A-5,JSON格式列表
      v_缴费渠道 := ',"payChannelDetail":[' || v_缴费渠道 || ']';
    Else
      v_缴费渠道 := ',"payChannelDetail":[]';
    End If;
  
    -------------------------------------------------------------------------------------------
    --其他医保信息
    v_其它医保信息 := Null;
    --otherMedicalList  其它医保信息列表  String  不限  否  填写其它未知医保信息（在电子票据上以内容拼接方式显示）
    --            详见A-4,JSON格式列表
    --  infoNo  序号  Integer  不限  是  默认从1开始，每项数据序号值递增1，本次不允许重复
    --  infoName  医保信息名称  String  100  是  如费用报销类型编码，可参考附录7医保报销类型列表
    --  infoValue  医保信息值  String  100  是  如费用报销金额
    --  infoOther  医保其它信息  String  100  否  如医保报销比例。
  
    -------------------------------------------------------------------------------------------
    --其它扩展信息
    v_其它扩展信息 := Null;
    --otherInfo  其它扩展信息列表  String  不限  否  填写信息需要在电子票据上单独显示的其它扩展信息（未知信息）
    --          详见A-3,JSON格式列表
    --  infoNo  序号  Integer  不限  是  默认从1开始，每项数据序号值递增1，本次不允许重复
    --  infoName  扩展信息名称  String  100  是
    --  infoValue  扩展信息值  String  500  是
  
    c_交易信息 := To_Clob('{' || v_票据信息);
    c_交易信息 := c_交易信息 || To_Clob(v_缴费);
    c_交易信息 := c_交易信息 || To_Clob(v_通知);
    c_交易信息 := c_交易信息 || To_Clob(v_就诊信息);
  
    If v_预交 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_预交);
    Else
      c_交易信息 := c_交易信息 || c_预交;
    End If;
  
    c_交易信息 := c_交易信息 || To_Clob(v_费用);
    c_交易信息 := c_交易信息 || To_Clob(v_缴费渠道);
  
    If v_其它扩展信息 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_其它扩展信息);
    End If;
    If v_其它医保信息 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_其它医保信息);
    End If;
  
    --eBillRelateNo  业务票据关联号  String  32  否  如一笔业务数据需要开具N张电子票据，则N张电子票对应该值保持一致，用于后期关联查询
    --isArrears  是否可流通  String  1  是  0-否、1-是（如欠费情况根据医院业务要求该票据是否可流通）
    --arrearsReason  不可流通原因  String  200  否  isArrears=0，填写不可流通的原因
    c_交易信息 := c_交易信息 || To_Clob(',"eBillRelateNo":"","isArrears":"1","arrearsReason":""');
    If v_分类明细 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_分类明细);
    Else
      c_交易信息 := c_交易信息 || c_分类明细;
    End If;
  
    If v_明细 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_明细);
    Else
      c_交易信息 := c_交易信息 || c_明细;
    End If;
    c_交易信息  := c_交易信息 || To_Clob('}');
    Reqdata_Out := c_交易信息;
    Code_Out    := 1;
  Exception
    When Others Then
      Message_Out := SQLCode || ':' || SQLErrM;
      Code_Out    := 0;
  End Get_Zybalancedata_Create;

  Procedure Get_Mzbalancedata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  ) Is
    --
    ---------------------------------------------------------------------------
    --功能:获取门诊结帐开票数据
    --入参:
    --    Json_In,格式如下
    --  input
    --    occasion N 1  应用场合:1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    --    einvoice_id  N,1 当前电子票据ID
    --    balance_id N 1  结算ID  occasion=2时，预交id;occasion<>2表示结帐id
    --    writeoff_id  N 1  冲销ID  occasion=2时，冲销预交id;occasion<>2表示冲销id
    --出参:
    --  ReqData_Out-返回的业务请求数据
    --  Totalmoney_Out-票据总额
    --  Code_Out-获取是否成功：0-失败；1-成功
    --  Message_Out 错误信息
    ---------------------------------------------------------------------------
    j_Input PLJson;
    j_Json  PLJson;
  
    n_应用场合 Number(2);
    n_结算id   病人预交记录.结帐id%Type;
  
    v_业务流水号 Varchar2(50);
  
    v_开票点       Varchar2(100);
    v_缴费         Varchar2(32767);
    v_票据信息     Varchar2(32767);
    v_就诊信息     Varchar2(32767);
    v_通知         Varchar2(32767);
    v_缴费渠道     Varchar2(32767);
    v_费用         Varchar2(32767);
    v_其它扩展信息 Varchar2(32767);
    v_其它医保信息 Varchar2(32767);
    c_明细         Clob;
    v_明细         Varchar2(32767);
    c_分类明细     Clob;
    v_分类明细     Varchar2(32767);
    c_交易信息     Clob; --最终返回的交易信息集
  
    v_患者姓名 门诊费用记录.姓名%Type;
    v_患者性别 门诊费用记录.性别%Type;
    v_患者年龄 门诊费用记录.年龄%Type;
  
    n_缺省卡类别id     Number(18);
    v_参数值           Varchar2(100);
    n_票据总金额       门诊费用记录.结帐金额%Type;
    n_误差总额         门诊费用记录.结帐金额%Type;
    n_用户id           人员表.Id%Type;
    v_操作员编号       人员表.编号%Type;
    v_操作员姓名       人员表.姓名%Type;
    v_Temp             Varchar2(32767);
    v_医疗付款方式编码 医疗付款方式.编码%Type;
    v_医疗付款方式名称 医疗付款方式.名称%Type;
    n_险类             保险结算记录.险类%Type;
    v_保险机构编码     保险类别.保险机构编码%Type;
    n_医嘱序号         门诊费用记录.医嘱序号%Type;
    n_挂号id           门诊费用记录.挂号id%Type;
    v_病种名称         保险病种.名称%Type;
    v_就诊日期         Varchar2(20);
    v_就诊科室编码     部门表.编码%Type;
    v_就诊科室名称     部门表.名称%Type;
    v_就诊编号         Varchar2(50);
    v_医保号           保险帐户.医保号%Type;
    v_版本号           Varchar2(30);
    n_电子票据id       电子票据使用记录.Id%Type;
  
    Cursor c_Balance_Record Is
      Select a.No, a.收费时间, a.结帐类型, a.操作员编号, a.操作员姓名, a.病人id, a.主页id, Decode(Nvl(a.病人id, 0), 0, a.原因, c.姓名) As 姓名,
             '' As 性别, '' As 年龄, c.门诊号, a.备注, a.结帐金额, Decode(Nvl(a.病人id, 0), 0, q.电子邮件, c.Email) As Email, q.联系人,
             Decode(Nvl(a.病人id, 0), 0, q.社会信用代码, c.身份证号) As 身份证号,
             Decode(Nvl(a.病人id, 0), 0, Nvl(q.电话, To_Char(j.移动电话)), c.手机号) As 手机号,
             Decode(Nvl(a.病人id, 0), 0, 2, 1) As 缴款类型, Decode(Nvl(a.结帐类型, 0), 1, '02', '01') As 业务标识, c.门诊号 As 病历号
      From 病人结帐记录 A, 病人信息 C, 合约单位 Q, 人员表 J
      Where a.Id = n_结算id And a.病人id = c.病人id(+) And a.原因 = q.名称(+) And q.联系人 = j.姓名(+);
    r_Balance_Record c_Balance_Record%RowType;
  
  Begin
    j_Input := PLJson(Json_In);
    j_Json  := j_Input.Get_Pljson('input');
  
    n_应用场合 := Nvl(j_Json.Get_Number('occasion'), 0);
    n_结算id   := j_Json.Get_Number('balance_id');
    --n_冲销id   := Nvl(j_Json.Get_Number('writeoff_id'), 0);
  
    n_电子票据id := Nvl(j_Json.Get_Number('einvoice_id'), 0);
  
    If Nvl(n_应用场合, 0) = 0 Then
      Code_Out    := 0;
      Message_Out := '无效的应用场景';
      Return;
    End If;
  
    Select Nvl(Max(参数值), 'V2.0.3')
    Into v_版本号
    From 三方接口配置
    Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '支持版本';
  
    b_Einvoice_Request_Test.Get_Identity(n_用户id, v_操作员编号, v_操作员姓名);
  
    --v_业务标识:01  住院,02  门诊,03  急诊,04  门特,05  体检中心,06  挂号,07  住院预交金,08  体检预交金
  
    n_票据总金额 := 0;
    c_明细       := Null;
    v_明细       := Null;
  
    Open c_Balance_Record;
    Fetch c_Balance_Record
      Into r_Balance_Record;
    If c_Balance_Record%NotFound Then
      Code_Out    := 0;
      Message_Out := '未找到指定的结算数据';
      Return;
    End If;
  
    For c_收费细目 In (Select Min(a.Id) As 费用id, a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号) As 序号, a.收费细目id, Max(a.计算单位) As 计算单位,
                          Sum(a.标准单价) As 价格, Avg(Nvl(a.付数, 1) * Nvl(a.数次, 0)) As 数量, Sum(a.应收金额) As 应收金额,
                          Sum(a.实收金额) As 实收金额, Sum(a.结帐金额) As 结帐金额, Sum(a.实收金额) - Sum(a.统筹金额) As 自费金额,
                          Max(t.编码) As 医保项目编码, Max(t.名称) As 医保项目名称, Max(t.统筹比额) As 医保报销比例, Max(a.摘要) As 备注,
                          Max(a.费用类型) As 费用类型, Max(a.操作员编号) As 操作员编号, Max(a.操作员姓名) As 操作员姓名, Max(a.姓名) As 姓名,
                          Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.病人id) As 病人id, Max(a.登记时间) As 登记时间,
                          Max(a.付款方式) As 付款方式编码, Max(a.收据费目) As 收据费目, Max(c.编码) As 收据费目编码, Max(a.医嘱序号) As 医嘱序号,
                          Max(a.挂号id) As 挂号id, Max(d.编码) As 类别编码, Max(d.类别) As 类别名称, Max(b.编码) As 项目编码, Max(b.名称) As 项目名称,
                          Max(b.规格) As 规格, Max(q.药品剂型) As 药品剂型
                   From 门诊费用记录 A, 收费项目目录 B, 收据费目 C, 收费类别 D, 药品规格 M, 药品特性 Q, 诊疗项目目录 J, 保险支付大类 T, 支付类别对照 S
                   Where 结帐id = n_结算id And a.记帐费用 = 1 And a.收费类别 = d.编码(+) And a.收费细目id = b.Id And a.收据费目 = c.名称(+) And
                         a.收费细目id = m.药品id(+) And m.药名id = q.药名id(+) And q.药名id = j.Id(+) And a.保险大类id = t.Id(+) And
                         t.性质(+) = 1 And a.保险大类id = s.保险大类id(+)
                   Group By a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号), a.收费细目id, c.编码, c.名称, j.编码, j.名称
                   Order By NO, 序号) Loop
      If v_患者姓名 Is Null Then
        v_患者姓名 := c_收费细目.姓名;
        v_患者性别 := c_收费细目.性别;
        v_患者年龄 := c_收费细目.年龄;
      End If;
    
      If v_医疗付款方式编码 Is Null Then
        v_医疗付款方式编码 := c_收费细目.付款方式编码;
      End If;
      If Nvl(n_医嘱序号, 0) = 0 Then
        n_医嘱序号 := c_收费细目.医嘱序号;
      End If;
      If Nvl(n_挂号id, 0) = 0 Then
        n_挂号id := c_收费细目.挂号id;
      End If;
    
      --listDetailNo  明细流水号  String  60  否  明细流水号
      v_Temp := '{' || '"listDetailNo":"' || b_Einvoice_Request_Test.Zljsonstr(LPad(c_收费细目.费用id, 20, '0')) || '"';
      --chargeCode  收费项目代码  String  50  否  填写业务系统内部编码值，由医疗平台配置对照,如：床位费、检查费
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.收据费目编码) || '"';
      --chargeName  收费项目名称  String  100  否
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.收据费目) || '"';
      --prescribeCode  处方编码  String  60  否
      v_Temp := v_Temp || ',' || '"prescribeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.No) || '"';
      --listTypeCode  药品类别编码  String  50  否  如药品分类编码01，有则填写
      v_Temp := v_Temp || ',' || '"listTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.类别编码) || '"';
      --listTypeName  药品类别名称  String  50  否  如药品分类名称，抗生素类抗感染药物
      v_Temp := v_Temp || ',' || '"listTypeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.类别名称) || '"';
      --code  编码  String  50  否  如药品编码，有则填写
      v_Temp := v_Temp || ',' || '"code":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.项目编码) || '"';
      --name  药品名称  String  50  是  如药品名称，器材名称等
      v_Temp := v_Temp || ',' || '"name":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.项目名称) || '"';
      --form  剂型  String  50  否
      v_Temp := v_Temp || ',' || '"form":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.药品剂型) || '"';
      --specification  规格  String  50  否
      v_Temp := v_Temp || ',' || '"specification":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.规格) || '"';
      --unit  计量单位   String  20  否
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.计算单位) || '"';
      --std  单价  Number  14,6  是
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.价格, 1);
      --number  数量  Number  14,6  是
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.数量, 1);
      --amt  金额  Number  14,6  是
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.实收金额, 1);
      --selfAmt  自费金额  Number  14,6  是  如无金额，填写0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.自费金额, 1);
      --receivableAmt  应收费用  Number  14,6  否
      v_Temp := v_Temp || ',' || '"receivableAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.应收金额, 1);
      --medicalCareType  医保药品分类  String  1  否  1：无自负/甲
      --          2：有自负/乙
      --          3：全自负/丙
      v_Temp := v_Temp || ',' || '"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.医保项目编码) || '"';
      --medCareItemType  医保项目类型  String  100  否
      v_Temp := v_Temp || ',' || '"medCareItemType":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.医保项目名称) || '"';
      --medReimburseRate  医保报销比例  Number  3,2  否
      v_Temp := v_Temp || ',' || '"medReimburseRate":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.医保报销比例, 1);
      --remark  备注  String  200  否
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.备注) || '"';
      --sortNo  序号  Integer  不限  否  序号
      v_Temp := v_Temp || ',' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.序号, 1);
      --chrgtype  费用类型  String  50  否
      v_Temp := v_Temp || ',' || '"chrgtype":"' || b_Einvoice_Request_Test.Zljsonstr(c_收费细目.费用类型) || '"}';
    
      If Length(Nvl(v_明细, '') || v_Temp) > 32700 Then
        If c_明细 Is Null Then
          c_明细 := To_Clob(v_明细);
        Else
          c_明细 := c_明细 || To_Clob(',' || v_明细);
        End If;
        v_明细 := Null;
      End If;
    
      If v_明细 Is Null Then
        v_明细 := v_Temp;
      Else
        v_明细 := v_明细 || ',' || v_Temp;
      End If;
    End Loop;
  
    If v_明细 Is Not Null And c_明细 Is Not Null Then
      --listDetail  清单项目明细  String  不限  是  详见A-2,JSON格式列表
      c_明细 := c_明细 || ',' || To_Clob(v_明细);
      c_明细 := To_Clob(',"listDetail":[') || c_明细 || To_Clob(']');
    
      v_明细 := Null;
    Elsif v_明细 Is Not Null Then
      v_明细 := ',"listDetail":[' || v_明细 || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --分类明细
    v_分类明细 := Null;
    c_分类明细 := Null;
    For c_分类统计 In (Select Rownum As 序号, 收据费目编码, 收据费目名称, 数量, 计算单位, Round(单价, 2) As 单价, Round(结帐金额, 2) As 结帐金额,
                          Round(自费金额, 2) As 自费金额, 备注, 结帐金额 - Round(结帐金额, 2) As 误差费
                   From (Select /*+cardinality(b,10)*/
                           c.编码 As 收据费目编码, a.收据费目 As 收据费目名称, 1 As 数量, '' As 计算单位, Sum(a.结帐金额) As 单价, a.收据费目,
                           Sum(a.结帐金额) As 结帐金额, Sum(a.结帐金额) - Sum(a.统筹金额) As 自费金额, '' As 备注
                          From 门诊费用记录 A, 收据费目 C
                          Where a.结帐id = n_结算id And a.收据费目 = c.名称(+)
                          Group By c.编码, a.收据费目)) Loop
      --sortNo  序号  Integer  不限  是  默认从1开始，每个收费项目序号值递增1，本次不允许重复
      v_Temp := '{' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.序号, 1);
      --chargeCode  收费项目代码  String  50  是  填写业务系统内部编码值，由医疗平台配置对照
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.收据费目编码) || '"';
      --chargeName  收费项目名称  String  100  是  填写业务系统内部项目名称
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.收据费目名称) || '"';
      --unit  计量单位  String  20  否
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.计算单位) || '"';
      --std  收费标准  Number  14,2  是
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.单价, 1);
      --number  数量  Number  14,2  是
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.数量, 1);
      --amt  金额  Number  14,2  是
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.结帐金额, 1);
      --selfAmt  自费金额  Number  14,2  是  如无金额，填写0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.自费金额, 1);
      --remark  备注  String  200  否
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_分类统计.备注) || '"}';
    
      If Length(Nvl(v_分类明细, '') || v_Temp) > 32700 Then
        If c_分类明细 Is Null Then
          c_分类明细 := To_Clob(v_分类明细);
        Else
          c_分类明细 := c_分类明细 || To_Clob(',' || v_分类明细);
        End If;
        v_分类明细 := Null;
      End If;
    
      If v_分类明细 Is Null Then
        v_分类明细 := v_Temp;
      Else
        v_分类明细 := v_分类明细 || ',' || v_Temp;
      End If;
    
      n_票据总金额 := Nvl(n_票据总金额, 0) + Nvl(c_分类统计.结帐金额, 0);
      n_误差总额   := Nvl(n_误差总额, 0) + Nvl(c_分类统计.误差费, 0);
    End Loop;
    Totalmoney_Out := n_票据总金额;
    If v_分类明细 Is Not Null And c_分类明细 Is Not Null Then
      c_分类明细 := c_分类明细 || ',' || To_Clob(v_分类明细);
      --chargeDetail 收费项目明细  String  不限  是  详见A-1,JSON格式列表
      c_分类明细 := To_Clob(',"chargeDetail":[') || c_分类明细 || To_Clob(']');
      v_分类明细 := Null;
    Elsif v_分类明细 Is Not Null Then
      v_分类明细 := ',"chargeDetail":[' || v_分类明细 || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --票据信息
    --Select Count(1) Into n_作废次数 From 电子票据使用记录 Where 票种 = n_应用场合 And 结算id = n_结算id;
    --lpad(电子票据作废次数,5) & Lpad(原结帐ID,20)
    --结算ID||_||电子票据ID
    v_业务流水号 := n_结算id || '_' || Nvl(n_电子票据id, 0);
    v_开票点     := b_Einvoice_Request_Test.Get_Einvoice_Node(v_操作员编号, v_操作员姓名, n_用户id);
  
    v_票据信息 := '"busNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_业务流水号) || '"'; --业务流水号
    v_票据信息 := v_票据信息 || ',"busType":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.业务标识) || '"'; --业务标识
    If Nvl(r_Balance_Record.病人id, 0) = 0 Then
      v_票据信息 := v_票据信息 || ',"payer":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.姓名) || '"'; --患者姓名
    Else
      v_票据信息 := v_票据信息 || ',"payer":"' || b_Einvoice_Request_Test.Zljsonstr(v_患者姓名) || '"'; --患者姓名
    End If;
    v_票据信息 := v_票据信息 || ',"busDateTime":"' ||
              b_Einvoice_Request_Test.Zljsonstr(To_Char(r_Balance_Record.收费时间, 'yyyymmddHH24miss') || '000') || '"';
    v_票据信息 := v_票据信息 || ',"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_开票点) || '"'; --开票点编码:直接填写业务系统内部编码值，由医疗平台配置对照
    v_票据信息 := v_票据信息 || ',"payee":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.操作员姓名) || '"'; --收费员
  
    v_票据信息 := v_票据信息 || ',"author":"' || b_Einvoice_Request_Test.Zljsonstr(v_操作员姓名) || '"'; --票据编制人
    v_票据信息 := v_票据信息 || ',"totalAmt":' || b_Einvoice_Request_Test.Zljsonstr(n_票据总金额, 1); --开票总金额
    v_票据信息 := v_票据信息 || ',"remark":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.备注) || '"'; --备注
    -------------------------------------------------------------------------------------------
  
    --取缴费信息
    v_缴费 := Null;
    For c_缴费 In (Select Max(Decode(信息名, '订单号', 信息值, '支付订单号', 信息值, '')) As 支付定单号,
                        Max(Decode(信息名, '医保支付订单号', 信息值, '医保订单号', 信息值, '')) As 医保支付定单号,
                        Max(Decode(Upper(信息名), '支付宝公众号USERID', 信息值, '')) As 支付宝公众号userid,
                        Max(Decode(Upper(信息名), '支付宝小程序USERID', 信息值, '')) As 支付宝小程序userid,
                        Max(Decode(Upper(信息名), '微信公众号OPENID', 信息值, '')) As 微信公众号openid,
                        Max(Decode(Upper(信息名), '微信小程序OPENID', 信息值, '')) As 微信小程序openid
                 From (Select 信息名, 信息值
                        From 病人信息从表
                        Where 病人id = r_Balance_Record.病人id And
                              信息名 In ('支付宝公众号USERID', '支付宝小程序USERID', '微信公众号OPENID', '微信小程序OPENID')
                        Union All
                        Select 交易项目, 交易内容
                        From 三方结算交易
                        Where 交易id In (Select ID From 病人预交记录 Where 结帐id = n_结算id) And 交易项目 Like '%订单号')) Loop
      v_缴费 := ',"alipayCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.支付宝公众号userid) || '"'; --患者支付宝账户
      v_缴费 := v_缴费 || ',"weChatOrderNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.支付定单号) || '"'; --微信支付订单号
      If v_版本号 = 'V3.1.0' Then
        --该版本才有此接点
        v_缴费 := v_缴费 || ',"weChatMedTransNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.医保支付定单号) || '"'; --微信医保支付订单号
      End If;
    
      If c_缴费.微信公众号openid Is Not Null Then
        v_缴费 := v_缴费 || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.微信公众号openid) || '"'; --微信公众号或小程序用户ID
      Else
        v_缴费 := v_缴费 || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_缴费.微信小程序openid) || '"'; --微信公众号或小程序用户ID
      End If;
      Exit;
    End Loop;
  
    -------------------------------------------------------------------------------------------
    --取通知信息
    Select To_Number(Max(参数值))
    Into n_缺省卡类别id
    From 三方接口配置
    Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '缺省卡类别ID';
  
    v_通知 := ',"tel":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.手机号) || '"'; --患者手机号码
    v_通知 := v_通知 || ',"email":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.Email) || '"'; --患者邮箱地址
    If v_版本号 = 'V3.1.0' Then
      --该版本才有此接点
      v_通知 := v_通知 || ',"payerType":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.缴款类型) || '"'; --交款人类型
    End If;
    v_通知 := v_通知 || ',"idCardNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.身份证号) || '"'; --统一社会信用代码
  
    v_Temp := Null;
    If Nvl(r_Balance_Record.病人id, 0) <> 0 Then
    
      For c_通知 In (Select Max(名称) As 卡类别, Max(卡号) As 卡号
                   From (
                          
                          Select 病人id, 名称, 编码, 卡号
                          From (Select b.病人id, c.名称, c.编码, b.卡号, Decode(b.卡类别id, n_缺省卡类别id, 2, c.缺省标志) As 缺省标志
                                  From 病人医疗卡信息 B, 医疗卡类别 C
                                  Where b.卡类别id = c.Id And b.病人id = Nvl(r_Balance_Record.病人id, 0)
                                  Order By 缺省标志)
                          Where Rownum < 2)) Loop
      
        If c_通知.卡类别 Is Not Null Then
          v_Temp := ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.卡类别) || '"'; --卡类型
          v_Temp := v_Temp || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_通知.卡号) || '"'; --卡号
        End If;
        Exit;
      End Loop;
      If r_Balance_Record.身份证号 Is Not Null And v_Temp Is Null Then
        Select Nvl(Max(参数值), '99998')
        Into v_参数值
        From 三方接口配置
        Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '身份证作卡类型编号';
        v_Temp := ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡类型
        v_Temp := v_Temp || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.身份证号) || '"'; --卡号
      End If;
    End If;
    If v_Temp Is Null Then
      --没有一张卡，固定一种卡类别
      Select Nvl(Max(参数值), '99999')
      Into v_参数值
      From 三方接口配置
      Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '病人无卡的卡类别编号';
      v_Temp := ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡类型
      Select Nvl(Max(参数值), '-')
      Into v_参数值
      From 三方接口配置
      Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '病人无卡的卡号';
      v_Temp := v_Temp || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_参数值) || '"'; --卡号
    End If;
    v_通知 := v_通知 || v_Temp;
  
    -------------------------------------------------------------------------------------------
    --就诊信息
    Select Max(内容) Into v_Temp From zlRegInfo Where 项目 = '医疗机构类型';
  
    --性质:1-收费;2-结算（包括住院结算、特殊门诊结算）；3-预交
    Select Max(a.险类), Max(b.保险机构编码), Max(Nvl(a.病种名称, c.名称))
    Into n_险类, v_保险机构编码, v_病种名称
    From 保险结算记录 A, 保险类别 B, 保险病种 C
    Where a.险类 = b.序号 And a.病种id = c.Id(+) And a.记录id = n_结算id And a.性质 = 2;
  
    Select Max(名称) Into v_医疗付款方式名称 From 医疗付款方式 Where 编码 = v_医疗付款方式编码;
    If Nvl(n_险类, 0) <> 0 Then
      Select Max(医保号) Into v_医保号 From 保险帐户 Where 病人id = Nvl(r_Balance_Record.病人id, 0) And 险类 = n_险类;
    End If;
  
    v_就诊编号 := Null;
    If Nvl(n_医嘱序号, 0) <> 0 Then
      Select Max(To_Char(a.发生时间, 'yyyy-mm-dd')), Max(b.编码), Max(b.名称), Max(a.No)
      Into v_就诊日期, v_就诊科室编码, v_就诊科室名称, v_就诊编号
      From 病人挂号记录 A, 部门表 B
      Where a.执行部门id = b.Id And
            a.No = (Select Max(挂号单) From 病人医嘱记录 Where ID = n_医嘱序号 Or 相关id = n_医嘱序号);
    Elsif Nvl(n_挂号id, 0) <> 0 Then
      Select Max(To_Char(a.发生时间, 'yyyy-mm-dd')), Max(b.编码), Max(b.名称), Max(a.No)
      Into v_就诊日期, v_就诊科室编码, v_就诊科室名称, v_就诊编号
      From 病人挂号记录 A, 部门表 B
      Where a.执行部门id = b.Id And a.Id = n_挂号id;
    End If;
    If v_就诊编号 Is Null And Nvl(r_Balance_Record.病人id, 0) <> 0 Then
      --取最近一次挂号ID
      Select Max(a.Id), Max(To_Char(a.发生时间, 'yyyy-mm-dd')), Max(b.编码), Max(b.名称), Max(a.No)
      Into n_挂号id, v_就诊日期, v_就诊科室编码, v_就诊科室名称, v_就诊编号
      From 病人挂号记录 A, 部门表 B
      Where a.执行部门id = b.Id And a.Id = (Select ID
                                        From (Select ID, 发生时间
                                               From 病人挂号记录
                                               Where 病人id = Nvl(r_Balance_Record.病人id, 0)
                                               Order By 发生时间 Desc)
                                        Where Rownum < 2);
    End If;
  
    If v_病种名称 Is Null And Nvl(n_险类, 0) <> 0 Then
    
      Select Max(病种名称)
      Into v_病种名称
      From (
             
             Select Distinct a.名称 As 病种名称
             From 保险病种 A, 保险特准项目 B
             Where a.险类 = n_险类 And a.Id = b.病种id And
                   b.收费细目id In (Select Distinct 收费细目id From 门诊费用记录 Where 结帐id = n_结算id)
             Union All
             Select Distinct a.名称 As 病种名称
             From 保险病种 A, 保险特准项目 B
             Where a.险类 = n_险类 And a.Id = b.病种id And
                   b.大类 In (Select Distinct 保险大类id From 门诊费用记录 Where 结帐id = n_结算id))
      Where Rownum < 2;
    End If;
    v_就诊信息 := ',"medicalInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_Temp) || '"'; --医疗机构类型
    v_就诊信息 := v_就诊信息 || ',"medCareInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_保险机构编码) || '"'; --医保机构的唯一编码
    v_就诊信息 := v_就诊信息 || ',"medCareTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_医疗付款方式编码) || '"'; --医保类型编码
    v_就诊信息 := v_就诊信息 || ',"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(v_医疗付款方式名称) || '"'; --取值范围包括职工基本医疗保险、城乡居民基本医疗保险（城镇居民基本医疗保险、新型农村合作医疗保险）和其他医疗保险、非医保等
    v_就诊信息 := v_就诊信息 || ',"medicalInsuranceID":"' || b_Einvoice_Request_Test.Zljsonstr(v_医保号) || '"';
    v_就诊信息 := v_就诊信息 || ',"consultationDate":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊日期) || '"'; --患者就医时间
    v_就诊信息 := v_就诊信息 || ',"category":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊科室名称) || '"'; --就诊科室
    v_就诊信息 := v_就诊信息 || ',"patientCategoryCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊科室编码) || '"'; --就诊科室编码
    --patientNo  患者就诊编号  String  20  是  患者每次就诊一次就生成的一个新的编号。（患者登记号）
    v_就诊信息 := v_就诊信息 || ',"patientNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_就诊编号) || '"';
  
    v_就诊信息 := v_就诊信息 || ',"patientId":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(r_Balance_Record.病人id, 0)) || '"'; --患者在业务系统中的唯一标识ID，类似身份证号码。
    If Nvl(r_Balance_Record.病人id, 0) = 0 Then
      v_就诊信息 := v_就诊信息 || ',"sex":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.性别) || '"'; --性别
      v_就诊信息 := v_就诊信息 || ',"age":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.年龄) || '"'; --年龄
    Else
      v_就诊信息 := v_就诊信息 || ',"sex":"' || b_Einvoice_Request_Test.Zljsonstr(v_患者性别) || '"'; --性别
      v_就诊信息 := v_就诊信息 || ',"age":"' || b_Einvoice_Request_Test.Zljsonstr(v_患者年龄) || '"'; --年龄
    End If;
    v_就诊信息 := v_就诊信息 || ',"caseNumber":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.病历号) || '"'; --病历号
    v_就诊信息 := v_就诊信息 || ',"specialDiseasesName":"' || b_Einvoice_Request_Test.Zljsonstr(v_病种名称) || '"'; --特殊病种名称
    -------------------------------------------------------------------------------------------
  
    --结算信息
    v_费用 := Null;
    For c_结算 In (Select 现金预交, 支票预交, 转账预交, 个人帐户支付, 医保统筹基金支付, 其它医保支付, 个人现金支付, Decode(Sign(现金支付), -1, 现金支付, 0) As 现金退款,
                        Decode(Sign(现金支付), -1, 支票支付, 0) As 支票退款, Decode(Sign(现金支付), -1, 转帐支付, 0) As 转帐退款,
                        Decode(Sign(现金支付), -1, 0, 现金支付) As 现金支付, Decode(Sign(现金支付), -1, 0, 支票支付) As 支票支付,
                        Decode(Sign(现金支付), -1, 0, 转帐支付) As 转帐支付,
                        Nvl(个人帐户支付, 0) + Nvl(医保统筹基金支付, 0) + Nvl(其它医保支付, 0) As 报销总额,
                        Nvl(结算总额, 0) - Nvl(个人帐户支付, 0) - Nvl(医保统筹基金支付, 0) - Nvl(其它医保支付, 0) As 自费金额, 结算总额, 医保结算号码,
                        0 As 个人帐户余额
                 From (Select /*+cardinality(b,10)*/
                         Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '现金', 1, 0), 0) * a.冲预交) As 现金预交,
                         Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '支票', 1, 0), 0) * a.冲预交) As 支票预交,
                         Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '支票', 0, '现金', 0, 1), 0) * a.冲预交) As 转账预交,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '个人账户支付', 1, 0)) * a.冲预交) As 个人帐户支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '医保统筹基金支付', 1, 0)) * a.冲预交) As 医保统筹基金支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '其它医保支付', 1, 0)) * a.冲预交) As 其它医保支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '其它医保支付', 0, '个人账户支付', 0, '医保统筹基金支付', 0, 1)) *
                              a.冲预交) As 个人现金支付,
                         Max(Decode(Mod(a.记录性质, 10), 1, 0,
                                     Decode(c.开票结算方式, '其它医保支付', 结算号码, '个人账户支付', 结算号码, '医保统筹基金支付', 结算号码, ''))) As 医保结算号码,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(a.结算方式, '现金', 1, 0)) * a.冲预交) As 现金支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(a.结算方式, '支票', 1, 0)) * a.冲预交) As 支票支付,
                         Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(a.结算方式, '现金', 0, '支票', 0, 1) * a.冲预交)) As 转帐支付,
                         Sum(冲预交) As 结算总额
                        From 病人预交记录 A, 开票结算对照 C
                        Where a.结帐id = n_结算id And a.结算方式 = c.结算方式(+)))
    
     Loop
      --accountPay  个人账户支付  Number  14,2  是  按政策规定用个人账户支付参保人的医疗费用（含基本医疗保险目录范围内和目录范围外的费用）；
      --          如无金额，填写0
      v_费用 := ',"accountPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_结算.个人帐户支付, 0), 1);
      --fundPay  医保统筹基金支付  Number  14,2  是  患者本次就医所发生的医疗费用中按规定由基本医疗保险统筹基金支付的金额；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"fundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_结算.医保统筹基金支付, 0), 1);
      --otherfundPay  其它医保支付  Number  14,2  是  患者本次就医所发生的医疗费用中按规定由大病保险、医疗救助、公务员医疗补助、大额补充、企业补充等基金或资金支付的金额；
      v_费用 := v_费用 || ',"otherfundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_结算.其它医保支付, 0), 1);
      --ownPay  自费金额  Number  14,2  是  患者本次就医所发生的医疗费用中按照有关规定不属于基本医疗保险目录范围而全部由个人支付的费用；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"ownPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.自费金额, 1);
      --selfConceitedAmt  个人自负  Number  14,2  是  医保患者起付标准内个人支付费用；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"selfConceitedAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfPayAmt  个人自付  Number  14,2  是  患者本次就医所发生的医疗费用中由个人负担的属于基本医疗保险目录范围内自付部分的金额；开展按病种、病组、床日等打包付费方式且由患者定额付费的费用。该项为个人所得税大病医疗专项附加扣除信；息项如无金额，填写0
      v_费用 := v_费用 || ',"selfPayAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfCashPay  个人现金支付  Number  14,2  是  个人通过现金、银行卡、微信、支付宝等渠道支付的金额；
      --          如无金额，填写0
      v_费用 := v_费用 || ',"selfCashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.个人现金支付, 1);
      --cashPay  现金预交款金额  Number  14,2  否
      v_费用 := v_费用 || ',"cashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.现金预交, 1);
      --chequePay  支票预交款金额  Number  14,2  否
      v_费用 := v_费用 || ',"chequePay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.支票预交, 1);
      --transferAccountPay  转账预交款金额  Number  14,2  否
      v_费用 := v_费用 || ',"transferAccountPay":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.转账预交, 1);
      --cashRecharge  补交金额(现金)  Number  14,2  否
      v_费用 := v_费用 || ',"cashRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.现金支付, 1);
      --chequeRecharge  补交金额(支票)  Number  14,2  否
      v_费用 := v_费用 || ',"chequeRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.支票支付, 1);
      --transferRecharge  补交金额（转账）  Number  14,2  否
      v_费用 := v_费用 || ',"transferRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.转帐支付, 1);
      --cashRefund  退还金额(现金)  Number  14,2  否
      v_费用 := v_费用 || ',"cashRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.现金退款, 1);
      --chequeRefund  退交金额(支票)  Number  14,2  否
      v_费用 := v_费用 || ',"chequeRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.支票退款, 1);
      --transferRefund  退交金额(转账)  Number  14,2  否
      v_费用 := v_费用 || ',"transferRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.转帐退款, 1);
      --ownAcBalance  个人账户余额  Number  14,2  否
      v_费用 := v_费用 || ',"ownAcBalance":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.个人帐户余额, 1);
      --reimbursementAmt  报销总金额  Number  14,2  否  医保结算后返回的总金额
      v_费用 := v_费用 || ',"reimbursementAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_结算.报销总额, 1);
      --balancedNumber  结算号  String  100  否  医保结算后生成的号码/入账唯一值
      v_费用 := v_费用 || ',"balancedNumber":"' || b_Einvoice_Request_Test.Zljsonstr(c_结算.医保结算号码) || '"';
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --交费渠道
    v_缴费渠道 := Null;
    For c_渠道 In (Select /*+cardinality(b,10)*/
                  Nvl(c.渠道编码, Nvl(d.渠道编码, '-')) As 渠道编码, Sum(冲预交) As 结算总额
                 From 病人预交记录 A, 收费渠道对照 C, (Select 结算方式, 渠道编码 From 收费渠道对照 D Where 卡类别id Is Null) D
                 Where a.结帐id = n_结算id And a.卡类别id = c.卡类别id(+) And a.结算方式 = c.结算方式(+) And a.结算方式 = d.结算方式(+)
                 Group By Nvl(c.渠道编码, Nvl(d.渠道编码, '-'))
                 Order By 渠道编码)
    
     Loop
      --payChannelCode  交费渠道编码  String  10  是
      If v_缴费渠道 Is Null Then
        v_缴费渠道 := Nvl(v_缴费渠道, '') || '{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_渠道.渠道编码, 0)) || '"';
      Else
        v_缴费渠道 := v_缴费渠道 || ',{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_渠道.渠道编码, 0)) || '"';
      End If;
      --payChannelValue  交费渠道金额  Number  14,2  是
      v_缴费渠道 := v_缴费渠道 || ',"payChannelValue":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_渠道.结算总额, 0), 1) || '}';
    End Loop;
  
    If v_缴费渠道 Is Not Null Then
      --payChannelDetail  交费渠道列表  String  不限  是  直接填写业务系统内部编码值，由医疗平台配置对照如：附录4交费渠道列表
      --        详见A-5,JSON格式列表
      v_缴费渠道 := ',"payChannelDetail":[' || v_缴费渠道 || ']';
    Else
      v_缴费渠道 := ',"payChannelDetail":[]';
    End If;
  
    -------------------------------------------------------------------------------------------
    --其他医保信息
    v_其它医保信息 := Null;
    --otherMedicalList  其它医保信息列表  String  不限  否  填写其它未知医保信息（在电子票据上以内容拼接方式显示）
    --            详见A-4,JSON格式列表
    --  infoNo  序号  Integer  不限  是  默认从1开始，每项数据序号值递增1，本次不允许重复
    --  infoName  医保信息名称  String  100  是  如费用报销类型编码，可参考附录7医保报销类型列表
    --  infoValue  医保信息值  String  100  是  如费用报销金额
    --  infoOther  医保其它信息  String  100  否  如医保报销比例。
  
    -------------------------------------------------------------------------------------------
    --其它扩展信息
    v_其它扩展信息 := Null;
    --otherInfo  其它扩展信息列表  String  不限  否  填写信息需要在电子票据上单独显示的其它扩展信息（未知信息）
    --          详见A-3,JSON格式列表
    --  infoNo  序号  Integer  不限  是  默认从1开始，每项数据序号值递增1，本次不允许重复
    --  infoName  扩展信息名称  String  100  是
    --  infoValue  扩展信息值  String  500  是
  
    c_交易信息 := To_Clob('{' || v_票据信息);
    c_交易信息 := c_交易信息 || To_Clob(v_缴费);
    c_交易信息 := c_交易信息 || To_Clob(v_通知);
    c_交易信息 := c_交易信息 || To_Clob(v_就诊信息);
    c_交易信息 := c_交易信息 || To_Clob(v_费用);
    c_交易信息 := c_交易信息 || To_Clob(v_缴费渠道);
  
    If v_其它扩展信息 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_其它扩展信息);
    End If;
    If v_其它医保信息 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_其它医保信息);
    End If;
    --  eBillRelateNo  业务票据关联号  String  32  否  如一笔业务数据需要开具N张电子票据，则N张电子票对应该值保持一致，用于后期关联查询
    c_交易信息 := c_交易信息 || To_Clob(',"eBillRelateNo":""');
    If v_分类明细 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_分类明细);
    Else
      c_交易信息 := c_交易信息 || c_分类明细;
    End If;
  
    If v_明细 Is Not Null Then
      c_交易信息 := c_交易信息 || To_Clob(v_明细);
    Else
      c_交易信息 := c_交易信息 || c_明细;
    End If;
    c_交易信息  := c_交易信息 || To_Clob('}');
    Reqdata_Out := c_交易信息;
    Code_Out    := 1;
  Exception
    When Others Then
      Message_Out := SQLCode || ':' || SQLErrM;
      Code_Out    := 0;
  End Get_Mzbalancedata_Create;

  Procedure Get_Depositdata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  ) Is
    ---------------------------------------------------------------------------
    --功能:获取预交开票数据
    --入参:
    --    Json_In,格式如下
    --  input
    --    occasion N 1  应用场合:1-收费,2-预交,3-结帐,4-挂号;5-就诊卡，固定传2
    --    einvoice_id  N,1 当前电子票据ID
    --    deposit_id N 1  预交ID
    --    writeoff_id  N 1  冲销ID  occasion=2时，冲销预交id;occasion<>2表示冲销id
    --出参:
    --  ReqData_Out-返回的业务请求数据
    --  Totalmoney_Out-票据总额
    --  Code_Out-获取是否成功：0-失败；1-成功
    --  Message_Out 错误信息
    ---------------------------------------------------------------------------
    j_Input PLJson;
    j_Json  PLJson;
  
    n_应用场合 Number(2);
    n_结算id   病人预交记录.结帐id%Type;
    n_冲销id   病人预交记录.结帐id%Type;
  
    v_业务流水号 Varchar2(50);
  
    v_开票点   Varchar2(100);
    v_缴费渠道 Varchar2(32767);
    c_交易信息 Clob; --最终返回的交易信息集
  
    v_预交     Varchar2(32767);
    v_卡类别   医疗卡类别.名称%Type;
    v_卡号     病人医疗卡信息.卡号%Type;
    n_预交余额 病人余额.预交余额%Type;
  
    n_缺省卡类别id Number(18);
    v_参数值       Varchar2(100);
    n_票据总金额   门诊费用记录.结帐金额%Type;
    n_用户id       人员表.Id%Type;
    v_操作员编号   人员表.编号%Type;
    v_操作员姓名   人员表.姓名%Type;
    v_Temp         Varchar2(32767);
    n_作废次数     Number(2);
    v_版本号       Varchar2(30);
    n_电子票据id   电子票据使用记录.Id%Type;
    Cursor c_Deposit_Rec Is
      Select a.No, a.收款时间, a.预交类别, a.卡类别id, a.病人id, a.主页id, a.科室id, a.缴款单位, a.单位开户行, a.单位帐号, a.摘要, a.结算方式, a.结算号码, a.卡号,
             a.交易流水号, a.交易说明, a.合作单位, a.金额, a.操作员编号, a.操作员姓名, Nvl(b.姓名, c.姓名) As 姓名, Nvl(b.性别, c.性别) As 性别,
             Nvl(b.年龄, c.年龄) As 年龄, c.门诊号, Nvl(b.住院号, c.住院号) As 住院号, c.Email, c.身份证号, c.手机号, 1 As 缴款类型,
             Decode(Nvl(a.预交类别, 0), 1, '07', '07') As 业务标识, d.编码 As 入院科室编码, d.名称 As 入院科室名称, e.编码 As 出院科室编码,
             e.名称 As 出院科室名称, b.入院日期, b.出院日期, Nvl(b.病案号, b.住院号) As 病历号, j.名称 As 医疗卡名称
      From 病人预交记录 A, 病案主页 B, 病人信息 C, 部门表 D, 部门表 E, 医疗卡类别 J
      Where a.Id = n_结算id And a.病人id = b.病人id(+) And a.主页id = b.主页id(+) And a.病人id = c.病人id(+) And b.入院科室id = d.Id(+) And
            b.出院科室id = e.Id(+) And a.卡类别id = j.Id(+);
    r_Deposit_Rec c_Deposit_Rec%RowType;
  
  Begin
    j_Input := PLJson(Json_In);
    j_Json  := j_Input.Get_Pljson('input');
  
    n_应用场合   := Nvl(j_Json.Get_Number('occasion'), 0);
    n_结算id     := j_Json.Get_Number('deposit_id');
    n_冲销id     := Nvl(j_Json.Get_Number('writeoff_id'), 0);
    n_电子票据id := Nvl(j_Json.Get_Number('einvoice_id'), 0);
  
    If Nvl(n_应用场合, 0) = 0 Then
      Code_Out    := 0;
      Message_Out := '无效的应用场景';
      Return;
    End If;
  
    Select Nvl(Max(参数值), 'V2.0.3')
    Into v_版本号
    From 三方接口配置
    Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '支持版本';
  
    b_Einvoice_Request_Test.Get_Identity(n_用户id, v_操作员编号, v_操作员姓名);
  
    --v_业务标识:01  住院,02  门诊,03  急诊,04  门特,05  体检中心,06  挂号,07  住院预交金,08  体检预交金
  
    n_票据总金额 := 0;
    Open c_Deposit_Rec;
    Fetch c_Deposit_Rec
      Into r_Deposit_Rec;
  
    If c_Deposit_Rec%NotFound Then
      Code_Out    := 0;
      Message_Out := '未找到指定的预交结算数据';
      Return;
    End If;
    Select Count(1) Into n_作废次数 From 电子票据使用记录 Where 票种 = n_应用场合 And 结算id = n_结算id;
  
    Begin
      Select 名称, 卡号
      Into v_卡类别, v_卡号
      From (Select b.病人id, c.名称, c.编码, b.卡号, Decode(b.卡类别id, n_缺省卡类别id, 2, c.缺省标志) As 缺省标志
             From 病人医疗卡信息 B, 医疗卡类别 C
             Where b.卡类别id = c.Id And b.病人id = r_Deposit_Rec.病人id
             Order By 缺省标志)
      Where Rownum < 2;
    Exception
      When Others Then
        v_卡号 := Null;
    End;
  
    If v_卡类别 Is Null Then
      If r_Deposit_Rec.身份证号 Is Not Null Then
        Select Nvl(Max(参数值), '99998')
        Into v_参数值
        From 三方接口配置
        Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '身份证作卡类型编号';
      
        v_卡类别 := v_参数值;
        v_卡号   := r_Deposit_Rec.身份证号;
      Else
        --没有一张卡，固定一种卡类别
        Select Nvl(Max(参数值), '99999')
        Into v_参数值
        From 三方接口配置
        Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '病人无卡的卡类别编号';
        v_卡类别 := v_参数值;
      
        Select Nvl(Max(参数值), '-')
        Into v_参数值
        From 三方接口配置
        Where 接口名 = '中联模拟接口(V1.0.0)' And 参数名 = '病人无卡的卡号';
        v_卡号 := v_参数值;
      End If;
    End If;
  
    If Nvl(n_冲销id, 0) = 0 Then
      n_票据总金额 := r_Deposit_Rec.金额;
    Else
      Select -1 * Max(冲预交) Into n_票据总金额 From 病人预交记录 Where ID = n_冲销id;
    End If;
  
    Select Max(预交余额)
    Into n_预交余额
    From 病人余额
    Where 病人id = r_Deposit_Rec.病人id And 性质 = 1 And 类型 = r_Deposit_Rec.预交类别;
  
    Totalmoney_Out := n_票据总金额;
  
    If Nvl(n_冲销id, 0) <> 0 Then
      v_开票点 := b_Einvoice_Request_Test.Get_Einvoice_Node(v_操作员编号, v_操作员姓名, n_用户id);
      v_预交   := Null;
      For c_票据 In (Select 代码, 号码, 凭证代码, 凭证号码
                   From 电子票据使用记录
                   Where ID = n_结算id And 退款id Is Null And 记录状态 = 1) Loop
        --退款开票
        --busType  业务标识  String  20  是  直接填写业务系统内部编码值，由医疗平台配置对照，列如：附录5 业务标识列表
        --          07:标识住院预交金
        --          08:标识体检预交金
        v_预交 := '{"busType":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.业务标识) || '"';
        --billBatchCode  电子票据代码  String  50  是  
        v_预交 := v_预交 || ',"billBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_票据.代码) || '"';
        --billNo  电子票据号码  String  20  是  
        v_预交 := v_预交 || ',"billNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_票据.号码) || '"';
        --reason  冲红原因  String  200  是  
        v_预交 := v_预交 || ',"reason":"' || b_Einvoice_Request_Test.Zljsonstr('退款') || '"';
        --operator  经办人  String  60  是  
        v_预交 := v_预交 || ',"operator":"' || b_Einvoice_Request_Test.Zljsonstr(v_操作员姓名) || '"';
        --busDateTime  业务发生时间  String  17  是  yyyyMMddHHmmssSSS
        v_预交 := v_预交 || ',"busDateTime":"' || To_Char(r_Deposit_Rec.收款时间, 'yyyymmddhh24miss') || '000' || '"';
        --placeCode  开票点编码  String  50  是  直接填写业务系统内部编码值，由医疗平台配置对照
        v_预交 := v_预交 || ',"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_开票点) || '"';
        --voucherBatchCode  预交金凭证代码  String  50  是  
        v_预交 := v_预交 || ',"voucherBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_票据.凭证代码) || '"';
        --voucherNo  预交金凭证号码  String  20  是  
        v_预交 := v_预交 || ',"voucherNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_票据.凭证号码) || '"';
        --amt  预缴金退款金额  Number  14,2  是  
        v_预交 := v_预交 || ',"amt":' || b_Einvoice_Request_Test.Zljsonstr(-1 * n_票据总金额, 1);
        --ownAcBalance  预缴金账户余额  Number  14,2  是  本次退款之前的账户余额
        v_预交 := v_预交 || ',"amt":' || b_Einvoice_Request_Test.Zljsonstr(n_预交余额, 1);
        --remark  备注  String  600  否  
        v_预交 := v_预交 || ',"voucherBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.摘要) || '"}';
        Exit;
      End Loop;
      If v_预交 Is Null Then
        Message_Out := '原始预交款未开具电子票据凭证，不允行开具退款票据';
        Code_Out    := 0;
      End If;
      Reqdata_Out := To_Clob(v_预交);
      Code_Out    := 1;
      Return;
    End If;
  
    -------------------------------------------------------------------------------------------
    --交费渠道
    Select Max(c.渠道编码)
    Into v_Temp
    From 收费渠道对照 C
    Where c.卡类别id = r_Deposit_Rec.卡类别id And c.结算方式 = r_Deposit_Rec.结算方式;
  
    If v_Temp Is Null Then
      Select Max(渠道编码)
      Into v_Temp
      From 收费渠道对照 D
      Where 卡类别id Is Null And 结算方式 = r_Deposit_Rec.结算方式;
    End If;
  
    --payChannelCode  交费渠道编码  String  10  是
    v_缴费渠道 := ',"payChannelDetail":[{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_Temp) || '"';
    --payChannelValue  交费渠道金额  Number  14,2  是
    v_缴费渠道 := v_缴费渠道 || ',"payChannelValue":' || b_Einvoice_Request_Test.Zljsonstr(n_票据总金额, 1) || '}]';
  
    --结算ID||_||电子票据ID
    v_业务流水号 := n_结算id || '_' || Nvl(n_电子票据id, 0);
    v_开票点     := b_Einvoice_Request_Test.Get_Einvoice_Node(v_操作员编号, v_操作员姓名, n_用户id);
  
    --busType  业务标识  String  20  是  直接填写业务系统内部编码值，由医疗平台配置对照，列如：附录5 业务标识列表
    --          值：
    --          07:标识住院预交金
    --          08:标识体检预交金
    v_预交 := '"busType":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.业务标识) || '"';
    --busNo  预交金业务流水号  String  50  是  单位内部唯一
    v_预交 := v_预交 || ',"busNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_业务流水号) || '"';
    --payer  患者姓名  String  100  是
    v_预交 := v_预交 || ',"busNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.姓名) || '"';
    --busDateTime  业务发生时间  String  17  是  格式：yyyyMMddHHmmssSSS
    v_预交 := v_预交 || ',"busDateTime":"' ||
            b_Einvoice_Request_Test.Zljsonstr(To_Char(r_Deposit_Rec.收款时间, 'yyyymmddhh24miss') || '000') || '"';
    --placeCode  开票点编码  String  50  是  直接填写业务系统内部编码值，由医疗平台配置对照
    v_预交 := v_预交 || ',"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_开票点) || '"';
    --payee  收款人  String  50  是
    v_预交 := v_预交 || ',"payee":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.操作员姓名) || '"';
    --drawee  缴款人  String  50  否  缴费人名称
    v_预交 := v_预交 || ',"drawee":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.姓名) || '"';
    --author  编制人  String  100  是
    v_预交 := v_预交 || ',"author":"' || b_Einvoice_Request_Test.Zljsonstr(v_操作员姓名) || '"';
    --tel  患者手机号码  String  13  否  患者手机号（如需要用于预交金凭证归集、短信通知，必填）
    v_预交 := v_预交 || ',"tel":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.手机号) || '"';
    --email  患者邮箱地址  String  100  否  患者邮箱地址（如需预交金凭证归集、短信通知，必填）
    v_预交 := v_预交 || ',"email":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.Email) || '"';
    --idCardNo  患者身份证号码  String  20  否
    v_预交 := v_预交 || ',"idCardNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.身份证号) || '"';
    --cardType  卡类型  String  10  是  办理预交金缴缴存对应的卡类型，如就诊卡、社保卡等
    --          直接填写业务系统内部编码值，由医疗平台配置对照
    --          列如：附录3卡类型列表
    v_预交 := v_预交 || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_卡类别) || '"';
    --cardNo  卡号  String  50  是  根据卡类型填写
    v_预交 := v_预交 || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_卡号) || '"';
    --amt  预缴金金额  Number  14,2  是
    v_预交 := v_预交 || ',"amt":' || b_Einvoice_Request_Test.Zljsonstr(n_票据总金额, 1);
    --ownAcBalance  预缴金账户余额  Number  14,2  是  本次缴存之前的账户余额
    v_预交 := v_预交 || ',"ownAcBalance":' || b_Einvoice_Request_Test.Zljsonstr(n_预交余额, 1);
    --category  入院科室名称  String  200  是
    v_预交 := v_预交 || ',"category":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.入院科室名称) || '"';
    --categoryCode  入院科室编码  String  100  是
    v_预交 := v_预交 || ',"categoryCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.入院科室编码) || '"';
    --inHospitalDate  入院日期  String  10  是  格式:yyyy-MM-dd
    v_预交 := v_预交 || ',"inHospitalDate":"' || b_Einvoice_Request_Test.Zljsonstr(To_Char(r_Deposit_Rec.入院日期, 'yyyy-mm-dd')) || '"';
    --hospitalNo  患者住院号  String  20  是  从入院到出院结束后，整个流程的唯一号
    v_预交 := v_预交 || ',"hospitalNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.住院号) || '"';
    --visitNo  住院就诊编号  String  20  是  住院期间，存在多次结算，结算后会重新生成一个住院就诊编号，如无就诊编号，可等于患者住院号
    v_预交 := v_预交 || ',"visitNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.住院号) || '"';
    --patientId  患者唯一ID  String  50  否  患者在业务系统中的唯一标识ID，类似身份证号码。
    v_预交 := v_预交 || ',"patientId":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.病人id) || '"';
    --patientNo  患者就诊编号  String  20  否  患者每次就诊一次就生成的一个新的编号。（患者登记号）
    v_预交 := v_预交 || ',"patientNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.主页id) || '"';
    --caseNumber  病历号  String  50  否  病案编号
    v_预交 := v_预交 || ',"caseNumber":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.病历号) || '"';
    --payChannelDetail  交费渠道列表  String  不限  是  直接填写业务系统内部编码值，由医疗平台配置对照如：附录4交费渠道列表
    --          详见A-1,JSON格式列表
    --payChannelCode  交费渠道编码  String  10  是
    --payChannelValue  交费渠道金额  Number  14,2  是
    v_预交 := v_预交 || v_缴费渠道;
    --accountName  账户名称  String  200  否  按需填写，如缴费渠道含银行卡
    v_预交 := v_预交 || ',"accountName":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(r_Deposit_Rec.医疗卡名称, '')) || '"';
    --accountNo  账户号码  String  200  否  按需填写，如缴费渠道含银行卡
    v_预交 := v_预交 || ',"accountNo":"' ||
            b_Einvoice_Request_Test.Zljsonstr(Nvl(r_Deposit_Rec.卡号, Nvl(r_Deposit_Rec.单位帐号, ''))) || '"';
    --accountBank  账户开户行  String  200  否  按需填写，如缴费渠道含银行卡
    v_预交 := v_预交 || ',"accountBank":"' ||
            b_Einvoice_Request_Test.Zljsonstr(Nvl(r_Deposit_Rec.医疗卡名称, Nvl(r_Deposit_Rec.单位开户行, ''))) || '"';
    --remark  备注  String  600  否
    v_预交 := v_预交 || ',"remark":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.摘要) || '"';
    If v_版本号 = 'V3.1.0' Then
      --workUnit  工作单位或地址      String  200  否  缴款人的工作单位或地址
      v_预交 := v_预交 || ',"workUnit":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.缴款单位) || '"';
    
    End If;
    c_交易信息  := To_Clob('{' || v_预交 || '}');
    Reqdata_Out := c_交易信息;
    Code_Out    := 1;
  Exception
    When Others Then
      Message_Out := SQLCode || ':' || SQLErrM;
      Code_Out    := 0;
  End Get_Depositdata_Create;

  Function Einvoice_Create
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    冲销id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number Is
    Pragma Autonomous_Transaction;
    ---------------------------------------------------------------------------
    --功能:进行电子票据开具
    --入参:  
    --    业务场景_In- 1-收费,2-预交,3-结帐,4-挂号;5-就诊卡 
    --    结算id_In-业务场景_In=2,预交ID;业务场景_In<>2:结帐ID
    --    冲销ID_In- 冲销ID  业务场景_In=2时，冲销预交id;业务场景_In<>2表示冲销id
    --出参: 
    --  错误信息_Out-返回=0时：返回错误 
    --返回:
    --   1-开票成功;0-失败
    ---------------------------------------------------------------------------
  
    n_电子票据id 电子票据使用记录.Id%Type;
    v_姓名       Varchar2(100);
  
    n_病人id     病人信息.病人id%Type;
    v_性别       病人信息.性别%Type;
    v_年龄       病人信息.年龄 %Type;
    n_门诊号     病人信息.门诊号%Type;
    n_住院号     病人信息.住院号%Type;
    n_Find       Number(2);
    n_票据金额   电子票据使用记录.票据金额%Type;
    v_开票点     电子票据使用记录.开票点%Type;
    v_凭证代码   电子票据使用记录.凭证代码%Type;
    v_凭证号码   电子票据使用记录.凭证号码%Type;
    v_凭证校验码 电子票据使用记录.凭证检验码%Type;
    v_票据代码   电子票据使用记录.代码%Type;
    v_票据号码   电子票据使用记录.号码%Type;
    v_票据校验码 电子票据使用记录.检验码%Type;
    v_系统来源   电子票据使用记录.系统来源%Type;
    v_生成时间   Varchar2(20);
    c_二维码     Clob;
    v_Url        电子票据使用记录.Url内网%Type;
    v_外网url    电子票据使用记录.Url外网%Type;
  
    v_操作员编号   人员表.编号%Type;
    v_操作员姓名   人员表.姓名%Type;
    n_用户id       人员表.Id%Type;
    v_Req_Json     Varchar2(32767);
    c_Req_Data     Clob;
    v_Err_Msg      Varchar2(4000);
    n_Code         Number(2);
    n_是否门诊     Number(2);
    v_Service_Name Varchar2(100);
    v_Version      Varchar2(20);
    v_Respdata     Varchar2(32767); --响应数据
    j_Json         PLJson;
  Begin
  
    If Nvl(业务场景_In, 0) < 1 Or Nvl(业务场景_In, 0) > 5 Then
      错误信息_Out := '不能识别的业务!';
      Return 0;
    End If;
    n_Find := 1;
    If 业务场景_In = 1 Or 业务场景_In = 4 Then
      --收费及挂号
      Begin
        Select a.病人id, Nvl(a.姓名, b.姓名) As 姓名, Nvl(a.年龄, b.年龄) As 年龄, Nvl(a.性别, b.性别) As 性别, b.门诊号, b.住院号 As 住院号
        Into n_病人id, v_姓名, v_性别, v_年龄, n_门诊号, n_住院号
        From 门诊费用记录 A, 病人信息 B
        Where a.结帐id = 结算id_In And a.病人id = b.病人id(+) And Rownum < 2;
      Exception
        When Others Then
          n_Find := 0;
      End;
    End If;
    If 业务场景_In = 2 Then
      --预交
      Begin
        Select a.病人id, Nvl(c.姓名, b.姓名) As 姓名, Nvl(c.年龄, b.年龄) As 年龄, Nvl(c.性别, b.性别) As 性别, b.门诊号,
               Nvl(c.住院号, b.住院号) As 住院号, Nvl(预交类别, 2)
        Into n_病人id, v_姓名, v_性别, v_年龄, n_门诊号, n_住院号, n_是否门诊
        From 病人预交记录 A, 病人信息 B, 病案主页 C
        Where a.Id = 结算id_In And a.病人id = b.病人id And a.病人id = c.病人id(+) And a.主页id = c.主页id(+);
      Exception
        When Others Then
          n_Find := 0;
      End;
    End If;
    --暂不支门诊预交
    If Nvl(n_是否门诊, 0) = 1 Then
      Return 1;
    End If;
    If 业务场景_In = 3 Then
      --结帐
      Begin
        Select a.病人id, Decode(Nvl(a.病人id, 0), 0, a.原因, Nvl(c.姓名, b.姓名)) As 姓名, Nvl(c.年龄, b.年龄) As 年龄,
               Nvl(c.性别, b.性别) As 性别, b.门诊号, Nvl(c.住院号, b.住院号) As 住院号, Nvl(结帐类型, 2) As n_是否门诊
        Into n_病人id, v_姓名, v_性别, v_年龄, n_门诊号, n_住院号, n_是否门诊
        From 病人结帐记录 A, 病人信息 B, 病案主页 C
        Where a.Id = 结算id_In And a.病人id = b.病人id And a.病人id = c.病人id(+) And a.主页id = c.主页id(+);
      Exception
        When Others Then
          n_Find := 0;
      End;
    End If;
  
    If 业务场景_In = 5 Then
      --医疗卡
      Begin
        Select a.病人id, Nvl(a.姓名, b.姓名) As 姓名, Nvl(a.年龄, b.年龄) As 年龄, Nvl(a.性别, b.性别) As 性别, b.门诊号, b.住院号 As 住院号
        Into n_病人id, v_姓名, v_性别, v_年龄, n_门诊号, n_住院号
        From 住院费用记录 A, 病人信息 B
        Where a.结帐id = 结算id_In And a.病人id = b.病人id(+) And Rownum < 2;
      Exception
        When Others Then
          n_Find := 0;
      End;
    End If;
  
    n_票据金额 := 0;
    If Nvl(n_Find, 0) = 0 Then
      --未找到原始结算数据
      错误信息_Out := '未找到需要开具电子票据的结算数据!';
      Return 0;
    End If;
  
    Select 电子票据使用记录_Id.Nextval Into n_电子票据id From Dual;
    If Nvl(n_病人id, 0) = 0 Then
      n_病人id := Null;
    End If;
    b_Einvoice_Request_Test.Get_Identity(n_用户id, v_操作员编号, v_操作员姓名);
    v_开票点 := b_Einvoice_Request_Test.Get_Einvoice_Node(v_操作员编号, v_操作员姓名, n_用户id);
  
    --1.先处理电子票据
    Zl_电子票据使用记录_Insert(n_电子票据id, 业务场景_In, 结算id_In, n_病人id, v_姓名, v_性别, v_年龄, n_门诊号, n_住院号, n_票据金额, v_开票点, v_系统来源, Null,
                       '', v_操作员编号, v_操作员姓名, Sysdate);
  
    --    occasion N 1  应用场合:1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    v_Req_Json := '"occasion":' || b_Einvoice_Request_Test.Zljsonstr(业务场景_In, 1);
    --    einvoice_id  N,1 当前电子票据ID
    v_Req_Json := v_Req_Json || ',"einvoice_id":' || b_Einvoice_Request_Test.Zljsonstr(n_电子票据id, 1);
    If 业务场景_In = 2 Then
      --deposit_id N 1  预交ID
      v_Req_Json := v_Req_Json || ',"deposit_id":' || b_Einvoice_Request_Test.Zljsonstr(结算id_In, 1);
    Else
      --balance_id N 1  结算ID  occasion=2时，预交id;occasion<>2表示结帐id
      v_Req_Json := v_Req_Json || ',"balance_id":' || b_Einvoice_Request_Test.Zljsonstr(结算id_In, 1);
    End If;
    --    writeoff_id  N 1  冲销ID  occasion=2时，冲销预交id;occasion<>2表示冲销id
    v_Req_Json := v_Req_Json || ',"writeoff_id":' || b_Einvoice_Request_Test.Zljsonstr(冲销id_In, 1);
    v_Req_Json := '{"input":{' || v_Req_Json || '}}';
  
    --2.获取电子票据
    If 业务场景_In = 1 Then
      --收费
      b_Einvoice_Request_Test.Get_Chargedata_Create(v_Req_Json, c_Req_Data, n_票据金额, n_Code, v_Err_Msg);
      v_Service_Name := 'invoiceEBillOutpatient';
      v_Version      := '1.0';
    Elsif 业务场景_In = 2 Then
      --预交
    
      b_Einvoice_Request_Test.Get_Depositdata_Create(v_Req_Json, c_Req_Data, n_票据金额, n_Code, v_Err_Msg);
      If Nvl(冲销id_In, 0) <> 0 Then
        --退款
        v_Service_Name := 'writeOffPayMentVoucher';
      Else
        v_Service_Name := 'invoicePayMentVoucher';
      End If;
      v_Version := '1.0';
    Elsif 业务场景_In = 3 And Nvl(n_是否门诊, 0) = 1 Then
      --门诊结帐
      b_Einvoice_Request_Test.Get_Mzbalancedata_Create(v_Req_Json, c_Req_Data, n_票据金额, n_Code, v_Err_Msg);
      v_Service_Name := 'invoiceEBillOutpatient';
      v_Version      := '1.0';
    Elsif 业务场景_In = 3 And Nvl(n_是否门诊, 0) <> 1 Then
      --住院结帐
      b_Einvoice_Request_Test.Get_Zybalancedata_Create(v_Req_Json, c_Req_Data, n_票据金额, n_Code, v_Err_Msg);
      v_Service_Name := 'invEBillHospitalized';
      v_Version      := '1.0';
    Elsif 业务场景_In = 4 Then
      --挂号
      b_Einvoice_Request_Test.Get_Sendcarddata_Create(v_Req_Json, c_Req_Data, n_票据金额, n_Code, v_Err_Msg);
      v_Service_Name := 'invEBillRegistration';
      v_Version      := '1.0';
    Elsif 业务场景_In = 5 Then
      --发卡
      b_Einvoice_Request_Test.Get_Registerdata_Create(v_Req_Json, c_Req_Data, n_票据金额, n_Code, v_Err_Msg);
      v_Service_Name := 'invoiceEBillOutpatient';
      v_Version      := '1.0';
    End If;
  
    If n_Code = 0 Then
      Rollback;
      错误信息_Out := v_Err_Msg;
      Return 0;
    End If;
  
    --进行业务请求
    n_Code := b_Einvoice_Request_Test.Request(c_Req_Data, v_Service_Name, v_Respdata, v_Err_Msg, v_Version);
    If n_Code = 0 Then
      错误信息_Out := v_Err_Msg;
      Rollback;
      Return 0;
    End If;
    --解析数据
    j_Json := PLJson(v_Respdata);
    If Nvl(业务场景_In, 0) = 2 Then
      If Nvl(冲销id_In, 0) <> 0 Then
      
        --voucherBatchCode  预交金红票凭证代码  String  50  是  
        v_凭证代码 := j_Json.Get_String('voucherBatchCode');
        --voucherNo  预交金红票凭证号码  String  20  是  
        v_凭证号码 := j_Json.Get_String('voucherNo');
        --voucherRandom  预交金红票凭证校验码  String  20  是  
        v_凭证校验码 := j_Json.Get_String('voucherRandom');
        --eScarletBillBatchCode  电子红票票据代码  String  50  是  
        v_票据代码 := j_Json.Get_String('eScarletBillBatchCode');
        --eScarletBillNo  电子红票票据号码  String  20  是  
        v_票据号码 := j_Json.Get_String('eScarletBillNo');
        --eScarletRandom  电子红票票据校验码  String  20  是  
        v_票据校验码 := j_Json.Get_String('eScarletRandom');
        --createTime  电子红票生成时间  String  17  是  创建时间：时间格式精确到毫秒yyyyMMddHHmmssSSS
        v_生成时间 := j_Json.Get_String('createTime');
        --billQRCode  电子红票二维码图片数据  String  不限  是  该值已Base64编码，解析时需要Base64解码，图片格式为PNG
        c_二维码 := j_Json.Get_Clob('billQRCode');
        --pictureUrl  电子票据H5页面URL  String  不限  是  
        v_Url := j_Json.Get_String('pictureUrl');
      Else
        --预交
        --voucherBatchCode  预交金凭证代码  String  50  是  
        v_凭证代码 := j_Json.Get_String('voucherBatchCode');
        --voucherNo  预交金凭证号码  String  20  是  
        v_凭证号码 := j_Json.Get_String('voucherNo');
        --voucherRandom  预交金凭证校验码  String  20  是  
        v_凭证校验码 := j_Json.Get_String('voucherRandom');
        --billBatchCode  电子票据代码  String  50  是  
        v_票据代码 := j_Json.Get_String('billBatchCode');
        --billNo  电子票据号码  String  20  是  
        v_票据号码 := j_Json.Get_String('billNo');
        --random  电子校验码  String  20  是  
        v_票据校验码 := j_Json.Get_String('random');
        --createTime  电子票据生成时间  String  17  是  创建时间：时间格式精确到毫秒yyyyMMddHHmmssSSS
        v_生成时间 := j_Json.Get_String('createTime');
        --billQRCode  电子票据二维码图片数据  String  不限  是  该值已Base64编码，解析时需要Base64解码，图片格式为PNG
        c_二维码 := j_Json.Get_Clob('billQRCode');
        --pictureUrl  电子票据H5页面URL  String  不限  是  
        v_Url := j_Json.Get_String('pictureUrl');
      End If;
    Else
      --其他
      --billBatchCode  电子票据代码  String  50  是  
      v_票据代码 := j_Json.Get_String('billBatchCode');
      --billNo  电子票据号码  String  20  是  
      v_票据号码 := j_Json.Get_String('billNo');
      --random  电子校验码  String  20  是  
      v_票据校验码 := j_Json.Get_String('random');
      --createTime  电子票据生成时间  String  17  是  创建时间：时间格式精确到毫秒yyyyMMddHHmmssSSS
      v_生成时间 := j_Json.Get_String('createTime');
      --billQRCode  电子票据二维码图片数据  String  不限  是  该值已Base64编码，解析时需要Base64解码，图片格式为PNG
      c_二维码 := j_Json.Get_Clob('billQRCode');
      --pictureUrl  电子票据H5页面URL  String  不限  是  
      v_Url := j_Json.Get_String('pictureUrl');
      --pictureNetUrl  电子票据外网H5页面URL  String  不限  否  按需配置
      v_外网url := j_Json.Get_String('pictureNetUrl');
    End If;
  
    --更新电子票据信息
    Update 电子票据使用记录
    Set 代码 = v_票据代码, 号码 = v_票据号码, 检验码 = v_票据校验码, 生成时间 = v_生成时间, Url内网 = v_Url, Url外网 = v_外网url, 系统来源 = '', 票据金额 = n_票据金额,
        凭证代码 = v_凭证代码, 凭证号码 = v_凭证号码, 凭证检验码 = v_凭证校验码
    Where ID = n_电子票据id;
    --保存二维码
    Insert Into 电子票据二维码 (使用记录id, 二维码) Values (n_电子票据id, c_二维码);
    Commit;
    Return 1;
  Exception
    When Others Then
      错误信息_Out := SQLCode || ':' || SQLErrM;
      Rollback;
      Return 0;
  End Einvoice_Create;

  --电子票据作废检查
  Function Einvoice_Cancel_Check
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number Is
    ---------------------------------------------------------------------------
    --功能:进行电子票据冲红检查
    --入参:  
    --    业务场景_In- 1-收费,2-预交,3-结帐,4-挂号;5-就诊卡 
    --    结算id_In-业务场景_In=2,预交ID;业务场景_In<>2:结帐ID 
    --出参: 
    --  错误信息_Out-返回=0时：返回错误 
    --返回:
    --   1-退票合法;0-退票不合法
    ---------------------------------------------------------------------------
  
    v_Req_Data     Varchar2(32767);
    v_Err_Msg      Varchar2(4000);
    n_Code         Number(2);
    v_Service_Name Varchar2(100);
    v_Version      Varchar2(20);
    v_Respdata     Varchar2(32767); --响应数据
    j_Json         PLJson;
  
    Cursor c_Einvoice Is
      Select a.Id, Nvl(a.是否换开, 0) As 是否换开, a.纸质发票号, a.代码, a.号码, a.检验码, a.生成时间
      From 电子票据使用记录 A
      Where a.Id = 结算id_In And a.记录状态 = 1 And 票种 = 业务场景_In;
    r_Einvoice c_Einvoice%RowType;
  
  Begin
  
    If Nvl(业务场景_In, 0) < 1 Or Nvl(业务场景_In, 0) > 5 Then
      错误信息_Out := '不能识别的业务!';
      Return 0;
    End If;
    If 业务场景_In = 2 Then
      --预交金没有相关检查接口，直接返回1
      Return 1;
    End If;
    Open c_Einvoice;
    Fetch c_Einvoice
      Into r_Einvoice;
    If c_Einvoice%NotFound Then
      --无电子票据相关数据;允许退，直接返回1
    
      Return 1;
    End If;
  
    If r_Einvoice.是否换开 = 1 Then
      --已经换开纸质票据，不允许再作废
      错误信息_Out := '已经换开纸质发票' || Nvl(r_Einvoice.纸质发票号, '') || '，禁止对电子票据进行冲红操作!';
      Return 0;
    End If;
  
    v_Service_Name := 'getEBillStatesByBillInfo';
    v_Version      := '1.0';
  
    --billBatchCode  电子票据代码  String  50  是  
    v_Req_Data := '{' || '"billBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.代码) || '"';
    --billNo  电子票据号码  String  20  是  
    v_Req_Data := v_Req_Data || ',' || '"billNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.号码) || '"}';
  
    --进行业务请求
    n_Code := b_Einvoice_Request_Test.Request(v_Req_Data, v_Service_Name, v_Respdata, v_Err_Msg, v_Version);
    If n_Code = 0 Then
      错误信息_Out := v_Err_Msg;
      Return 0;
    End If;
    --解析数据
    j_Json := PLJson(v_Respdata);
    --state  状态  String  1  是  状态：1正常，2作废
    If j_Json.Get_String('state') = '2' Then
      --作废了的，可以重新开具
      Return 1;
    End If;
    --isScarlet  是否已开红票  String  1  是  0未开红票，1已开红票
    If j_Json.Get_String('isScarlet') = '1' Then
      --已经开具红票，可以再进行开具
      Return 1;
    End If;
    --isPrtPaper  是否打印纸质票据  String  1  是  0未打印，1已打印
    If j_Json.Get_String('state') = '1' Then
      错误信息_Out := '已经打印纸质票据，不允许作废操作!';
      Return 0;
    End If;
    If b_Einvoice_Request_Test.Get_Version <> '3.1.0' Then
      --无下帐接口
      Return 1;
    End If;
  
    --4.1.16  查询电子票据入账状态接口
    v_Service_Name := 'getEBillStatesByBillInfo';
    v_Version      := '1.0';
  
    --billBatchCode  电子票据代码  String  50  是  值为开具接口返回的电子票据代码(无需对照)
    v_Req_Data := '{' || '"billBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.代码) || '"';
  
    --billNo  电子票据号  String  20  是  
    v_Req_Data := v_Req_Data || ',' || '"billNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.号码) || '"';
    --random  电子校验码  String  20  是  
    v_Req_Data := v_Req_Data || ',' || '"random":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.号码) || '"';
    --createTime  电子票据生成时间  String  17  是  开具电子票据返回的生成时间：yyyyMMddHHmmssSSS
    v_Req_Data := v_Req_Data || ',' || '"createTime":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.生成时间) || '"}';
  
    --进行业务请求
    n_Code := b_Einvoice_Request_Test.Request(v_Req_Data, v_Service_Name, v_Respdata, v_Err_Msg, v_Version);
    If n_Code = 0 Then
      错误信息_Out := v_Err_Msg;
      Return 0;
    End If;
    --解析数据
    j_Json := PLJson(v_Respdata);
    --state  入账状态  String  1  是  0未入账，1已入账
  
    If j_Json.Get_String('state') = '1' Then
      错误信息_Out := '该电子票据已经入帐，不允许再作废操作';
      Return 0;
    End If;
    Return 1;
  Exception
    When Others Then
      错误信息_Out := SQLCode || ':' || SQLErrM;
      Return 0;
  End Einvoice_Cancel_Check;

  --电子票据作废
  Function Einvoice_Cancel
  (
    业务场景_In  Integer,
    结算id_In    病人预交记录.结帐id%Type,
    错误信息_Out Out Varchar2
  ) Return Number Is
    ---------------------------------------------------------------------------
    --功能:进行电子票据冲红
    --入参:  
    --    业务场景_In- 1-收费,2-预交,3-结帐,4-挂号;5-就诊卡 
    --    结算id_In-业务场景_In=2,预交ID;业务场景_In<>2:结帐ID 
    --出参: 
    --  错误信息_Out-返回=0时：返回错误 
    --返回:
    --   1-退票合法;0-退票不合法
    ---------------------------------------------------------------------------
  
    v_Req_Data     Varchar2(32767);
    v_Err_Msg      Varchar2(4000);
    n_Code         Number(2);
    v_Service_Name Varchar2(100);
    v_Version      Varchar2(20);
    v_Respdata     Varchar2(32767); --响应数据
    j_Json         PLJson;
    n_人员id       人员表.Id%Type;
    v_操作员编号   人员表.编号%Type;
    v_操作员姓名   人员表.姓名%Type;
    v_业务发生时间 Varchar2(30);
  
    v_红票代码     电子票据使用记录.代码%Type;
    v_红票号码     电子票据使用记录.号码%Type;
    v_红票校验码   电子票据使用记录.检验码%Type;
    v_系统来源     电子票据使用记录.系统来源%Type;
    c_红票二维码   Clob;
    v_红票url      电子票据使用记录.Url内网%Type;
    v_红票外网url  电子票据使用记录.Url外网%Type;
    v_开票点       电子票据使用记录.开票点%Type;
    v_红票生成时间 电子票据使用记录.生成时间%Type;
    v_原因         Varchar2(50);
    v_摘要         病人预交记录.摘要%Type;
    n_冲销id       电子票据使用记录.Id%Type;
    Cursor c_Einvoice Is
      Select a.Id, Nvl(a.是否换开, 0) As 是否换开, a.纸质发票号, a.代码, a.号码, a.检验码, a.生成时间, a.病人id, a.住院号
      From 电子票据使用记录 A
      Where a.Id = 结算id_In And a.记录状态 = 1 And 票种 = 业务场景_In;
    r_Einvoice c_Einvoice%RowType;
  
  Begin
  
    If Nvl(业务场景_In, 0) < 1 Or Nvl(业务场景_In, 0) > 5 Then
      错误信息_Out := '不能识别的业务!';
      Return 0;
    End If;
  
    Open c_Einvoice;
    Fetch c_Einvoice
      Into r_Einvoice;
    If c_Einvoice%NotFound Then
      --无电子票据相关数据;允许退，直接返回1
      Return 1;
    End If;
  
    n_Code := b_Einvoice_Request_Test.Einvoice_Cancel_Check(业务场景_In, 结算id_In, 错误信息_Out);
    If n_Code = 0 Then
      --失败，直接退出
      Return n_Code;
    End If;
    b_Einvoice_Request_Test.Get_Identity(n_人员id, v_操作员编号, v_操作员姓名);
    v_开票点 := b_Einvoice_Request_Test.Get_Einvoice_Node(v_操作员编号, v_操作员姓名, n_人员id);
    n_Code   := 1;
  
    If 业务场景_In = 1 Then
      v_原因 := '退费';
      Begin
        Select To_Char(登记时间, 'yyyymmddhh24miss') || '000'
        Into v_业务发生时间
        From 门诊费用记录
        Where 结帐id = 结算id_In And Rownum < 2;
      Exception
        When Others Then
          n_Code := 0;
      End;
    Elsif 业务场景_In = 2 Then
      v_原因 := '退预交';
      Begin
        Select To_Char(收款时间, 'yyyymmddhh24miss') || '000'
        Into v_业务发生时间
        From 病人预交记录
        Where ID = 结算id_In And Rownum < 2;
      Exception
        When Others Then
          n_Code := 0;
      End;
    Elsif 业务场景_In = 3 Then
      v_原因 := '结帐作废';
      Begin
        Select To_Char(收费时间, 'yyyymmddhh24miss') || '000'
        Into v_业务发生时间
        From 病人结帐记录
        Where ID = 结算id_In And Rownum < 2;
      Exception
        When Others Then
          n_Code := 0;
      End;
    
    Elsif 业务场景_In = 4 Then
      v_原因 := '退号';
      Begin
        Select To_Char(登记时间, 'yyyymmddhh24miss') || '000'
        Into v_业务发生时间
        From 门诊费用记录
        Where 结帐id = 结算id_In And Rownum < 2;
      Exception
        When Others Then
          n_Code := 0;
      End;
    Elsif 业务场景_In = 5 Then
      v_原因 := '退卡';
      Begin
        Select To_Char(登记时间, 'yyyymmddhh24miss') || '000'
        Into v_业务发生时间
        From 住院费用记录
        Where 结帐id = 结算id_In And Rownum < 2;
      Exception
        When Others Then
          n_Code := 0;
      End;
    End If;
  
    If n_Code = 0 Then
      错误信息_Out := '未找到原始结算数据!';
      Return n_Code;
    End If;
  
    If 业务场景_In = 2 Then
      --撤消预交
      v_Service_Name := 'cancelPayMentVoucherBalance';
      v_Version      := '1.0';
    Else
      v_Service_Name := 'writeOffEBill';
      v_Version      := '1.0';
    End If;
  
    If 业务场景_In = 2 Then
      Select Max(摘要) Into v_摘要 From 病人预交记录 Where ID = 结算id_In;
    
      --billBatchCode  电子票据代码  String  50  是  
      v_Req_Data := '{' || '"billBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.代码) || '"';
      --billNo  电子票据号码  String  20  是  
      v_Req_Data := v_Req_Data || ',' || '"billNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.号码) || '"';
      --reason  冲红原因  String  200  是  
      v_Req_Data := v_Req_Data || ',' || '"reason":"' || b_Einvoice_Request_Test.Zljsonstr(v_原因) || '"';
      --operator  经办人  String  60  是  
      v_Req_Data := v_Req_Data || ',' || '"operator":"' || b_Einvoice_Request_Test.Zljsonstr(v_操作员姓名) || '"';
      --busDateTime  业务发生时间  String  17  是  yyyyMMddHHmmssSSS
      v_Req_Data := v_Req_Data || ',' || '"busDateTime":"' || b_Einvoice_Request_Test.Zljsonstr(v_业务发生时间) || '"';
      --placeCode  开票点编码  String  50  是  直接填写业务系统内部编码值，由医疗平台配置对照
      v_Req_Data := v_Req_Data || ',' || '"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_开票点) || '"';
      --patientId  患者唯一ID  String  50  否  
      v_Req_Data := v_Req_Data || ',' || '"patientId":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.病人id) || '"';
      --hospitalNo  患者住院号  String  20  是  
      v_Req_Data := v_Req_Data || ',' || '"hospitalNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.住院号) || '"';
      --remark  备注  String  600  否  
      v_Req_Data := v_Req_Data || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(v_摘要) || '"}';
    Else
      --billBatchCode  电子票据代码  String  50  是  
      v_Req_Data := '{' || '"billBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.代码) || '"';
      --billNo  电子票据号码  String  20  是  
      v_Req_Data := v_Req_Data || ',' || '"billNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.号码) || '"';
      --reason  冲红原因  String  200  是  
      v_Req_Data := v_Req_Data || ',' || '"reason":"' || b_Einvoice_Request_Test.Zljsonstr(v_原因) || '"';
      --operator  经办人  String  60  是  
      v_Req_Data := v_Req_Data || ',' || '"operator":"' || b_Einvoice_Request_Test.Zljsonstr(v_操作员姓名) || '"';
    
      --busDateTime  业务发生时间  String  17  是  yyyyMMddHHmmssSSS
      v_Req_Data := v_Req_Data || ',' || '"busDateTime":"' || b_Einvoice_Request_Test.Zljsonstr(v_业务发生时间) || '"';
      --placeCode  开票点编码  String  50  是  直接填写业务系统内部编码值，由医疗平台配置对照
      v_Req_Data := v_Req_Data || ',' || '"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_开票点) || '"}';
    End If;
  
    Select 电子票据使用记录_Id.Nextval Into n_冲销id From Dual;
    --先冲销
    Zl_电子票据使用记录_Delete(n_冲销id, v_开票点, v_系统来源, Null, '', v_操作员编号, v_操作员姓名, Sysdate, r_Einvoice.Id);
  
    --进行业务请求
    n_Code := b_Einvoice_Request_Test.Request(v_Req_Data, v_Service_Name, v_Respdata, v_Err_Msg, v_Version);
    If n_Code = 0 Then
      Rollback;
      错误信息_Out := v_Err_Msg;
      Return 0;
    End If;
  
    --解析数据
    j_Json := PLJson(v_Respdata);
    --  eScarletBillBatchCode  电子红票票据代码  String  20  是  
    v_红票代码 := j_Json.Get_String('eScarletBillBatchCode');
    --  eScarletBillNo  电子红票票据号码  String  20  是  
    v_红票号码 := j_Json.Get_String('eScarletBillNo');
    --  eScarletRandom  电子红票校验码  String  20  是  
    v_红票校验码 := j_Json.Get_String('eScarletRandom');
    --  createTime  电子红票生成时间  String  17  是  yyyyMMddHHmmssSSS
    v_红票生成时间 := j_Json.Get_String('createTime');
    --  billQRCode  电子票据二维码图片数据  String  不限    该值已Base64编码，解析时需要Base64解码
    c_红票二维码 := j_Json.Get_String('billQRCode');
    --  pictureUrl  电子票据H5页面URL  String  不限    
    v_红票url := j_Json.Get_String('pictureUrl');
    --  pictureNetUrl  电子票据外网H5页面URL地址  String  不限    按需配置
    v_红票外网url := j_Json.Get_String('pictureNetUrl');
    --更新电子票据信息
    Update 电子票据使用记录
    Set 代码 = v_红票代码, 号码 = v_红票号码, 检验码 = v_红票校验码, 生成时间 = v_红票生成时间, Url内网 = v_红票url, Url外网 = v_红票外网url
    Where ID = n_冲销id;
  
    --保存二维码
    Insert Into 电子票据二维码 (使用记录id, 二维码) Values (n_冲销id, c_红票二维码);
    Commit;
    Return 1;
  Exception
    When Others Then
      错误信息_Out := SQLCode || ':' || SQLErrM;
      Return 0;
  End Einvoice_Cancel;
End b_Einvoice_Request_Test;
/

Create Or Replace Procedure Zl_电子票据核对记录_Update
(
  核对类型_In     电子票据核对记录.核对类型%Type,
  业务日期_In     电子票据核对记录.业务日期%Type,
  开票点_In       电子票据核对记录.开票点%Type,
  His开票数_In    电子票据核对记录.His开票数%Type,
  His开票金额_In  电子票据核对记录.His开票金额%Type,
  平台开票数_In   电子票据核对记录.平台开票数%Type,
  平台开票金额_In 电子票据核对记录.平台开票金额%Type,
  核对人_In       电子票据核对记录.核对人%Type,
  核对时间_In     电子票据核对记录.核对时间%Type,
  核对结果_In     电子票据核对记录.核对结果%Type,
  核对说明_In     电子票据核对记录.核对说明%Type
) As
  --功能：更新/插入电子票据核对记录
  --入参：
  --  核对类型_In 1-核对开票和退票，2-仅核对退票
  --  核对结果_In 1-核对成功，0-核对失败
Begin
  Update 电子票据核对记录
  Set His开票数 = His开票数_In, His开票金额 = His开票金额_In, 平台开票数 = 平台开票数_In, 平台开票金额 = 平台开票金额_In, 核对时间 = 核对时间_In, 核对结果 = 核对结果_In,
      核对说明 = 核对说明_In
  Where 业务日期 = 业务日期_In And 核对人 = 核对人_In And 核对类型 = 核对类型_In;

  If Sql%RowCount = 0 Then
    Insert Into 电子票据核对记录
      (业务日期, 开票点, His开票数, His开票金额, 平台开票数, 平台开票金额, 核对类型, 核对人, 核对时间, 核对结果, 核对说明)
    Values
      (业务日期_In, 开票点_In, His开票数_In, His开票金额_In, 平台开票数_In, 平台开票金额_In, 核对类型_In, 核对人_In, 核对时间_In, 核对结果_In, 核对说明_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据核对记录_Update;
/
Create Or Replace Procedure Zl_电子票据修正记录_Update
(
  业务日期_In     电子票据修正记录.业务日期%Type,
  电子票据id_In   电子票据修正记录.电子票据id%Type,
  业务流水号_In   电子票据修正记录.业务流水号%Type,
  His开票点_In    电子票据修正记录.His开票点%Type,
  His开票金额_In  电子票据修正记录.His开票金额%Type,
  His票据状态_In  电子票据修正记录.His票据状态%Type,
  平台开票点_In   电子票据修正记录.平台开票点%Type,
  平台开票金额_In 电子票据修正记录.平台开票金额%Type,
  平台票据状态_In 电子票据修正记录.平台票据状态%Type,
  修正方式_In     电子票据修正记录.修正方式%Type,
  修正人_In       电子票据修正记录.修正人%Type,
  修正时间_In     电子票据修正记录.修正时间%Type,
  修正结果_In     电子票据修正记录.修正结果%Type,
  修正说明_In     电子票据修正记录.修正说明%Type
) As
  --功能：更新/插入电子票据修正记录
  --入参：
  --  HIS票据状态_In\平台票据状态_In 1-正常，2-冲红，3-作废
  --  修正方式_In 1-作废HIS数据，2-作废平台数据，3-作废HIS和平台数据重开票据，4-不修正仅标记
  --  修正结果_In 1-修正成功，0-修正失败
Begin
  If 电子票据id_In Is Null Then
    Update 电子票据修正记录
    Set 修正方式 = 修正方式_In, 修正人 = 修正人_In, 修正时间 = 修正时间_In, 修正结果 = 修正结果_In, 修正说明 = 修正说明_In
    Where 业务日期_In = 业务日期_In And 电子票据id Is Null And 业务流水号 = 业务流水号_In;
  Elsif 业务流水号_In Is Null Then
    Update 电子票据修正记录
    Set 修正方式 = 修正方式_In, 修正人 = 修正人_In, 修正时间 = 修正时间_In, 修正结果 = 修正结果_In, 修正说明 = 修正说明_In
    Where 业务日期_In = 业务日期_In And 电子票据id = 电子票据id_In And 业务流水号 Is Null;
  Else
    Update 电子票据修正记录
    Set 修正方式 = 修正方式_In, 修正人 = 修正人_In, 修正时间 = 修正时间_In, 修正结果 = 修正结果_In, 修正说明 = 修正说明_In
    Where 业务日期_In = 业务日期_In And 电子票据id = 电子票据id_In And 业务流水号 = 业务流水号_In;
  End If;

  If Sql%RowCount = 0 Then
    Insert Into 电子票据修正记录
      (业务日期, 电子票据id, 业务流水号, His开票点, His开票金额, His票据状态, 平台开票点, 平台开票金额, 平台票据状态, 修正方式, 修正人, 修正时间, 修正结果, 修正说明)
    Values
      (业务日期_In, 电子票据id_In, 业务流水号_In, His开票点_In, His开票金额_In, His票据状态_In, 平台开票点_In, 平台开票金额_In, 平台票据状态_In, 修正方式_In, 修正人_In,
       修正时间_In, 修正结果_In, 修正说明_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据修正记录_Update;
/