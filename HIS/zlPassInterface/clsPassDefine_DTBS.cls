VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPassDefine_DTBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------------------------------------------------------------
' --大通BS--- 接口定义<CRMs-MR-ITF><CRMs/前台接口>版本 2.2【上海大通医药信息技术有限公司】
'--------------------------------------------------------------------------------------------------------------------------------------
'<-begin------------------------------------------------------------------------base_xml----------------------------------------------------->
'<base_xml>
'    <source>HIS</source>
'    <hosp_code>医院编码</hosp_code>
'    <dept_code>科室代码</dept_code>
'    <dept_name>科室名称</dept_name>
'    <doct>
'        <code>医生代码</code>
'        <name>医生名称</name>
'        <type>医生级别代码</type>
'        <type_name>医生级别名称</type_name >
'    </doct>
'</base_xml>
Public Type DTBS_BASE
    strHIS As String
    str医院编码 As String
    str科室代码 As String
    str科室名称 As String
    str医生代码 As String
    str医生名称 As String
    str医生级别代码 As String
    str医生级别名称 As String
End Type
'<-end------------------------------------------------------------------------base_xml----------------------------------------------------->

'<-begin------------------------------------------------------------------------details_xml 格式如下:----------------------------------------------------->
'HIS命令 ：1-登录
'details_xml 格式如下:
'<details_xml>
'<doct_pwd>密码</doct_pwd>
'</details_xml>

'HIS命令 ：2 退出
'details_xml 格式如下:
'<details_xml></details_xml>

'HIS命令 ：3 初始化UI
'details_xml 格式如下:
'<details_xml></details_xml>

'
'HIS命令 ：5 要点提示
'details_xml 格式如下:
'<details_xml>
'<hosp_flag>门诊住院标识</hosp_flag>
'<medicine>
'        <his_code>药品代码</his_code>
'        <his_name>药品名称</his_name>
'</medicine>
'</details_xml>

'HIS命令 ：6 处方分析
'details_xml 格式如下:
'<details_xml is_upload ="是否上传0/1">
'   <his_time>HIS系统时间（YYYY-MM-DD HH:mm:SS）</his_time>
'   <hosp_flag>门诊住院标识 op/ip</hosp_flag>
'   <treat_type>就诊类型</treat_type>
'   <treat_code>就诊号</treat_code>
'   <lis_adm_no>检验就诊号</lis_adm_no>
'   <bed_no>床位号</bed_no>
'   <area_code>病区号</area_code>
'   <patient>
'       <name>姓名</name>
'       <is_infant>是否婴幼儿</is_infant>
'       <birth>出生日期(YYYY-MM-DD)</birth>
'       <sex>性别（男/女/未知）</sex>
'       <weight>体重（单位：千克）</weight>
'       <height>身高（单位：厘米）</height>
'       <id_card>身份证号</id_card>
'       <card_type>卡类型</card_type>
'       <card_code>卡号</card_code>
'       <pregnant_unit>时间单位（天、周、月）</pregnant_unit>
'       <pregnant >怀孕时间</pregnant>
'       <allergic_data>
'           <!-过敏节点，可循环'
'           <allergic>
'               <type>过敏类型</type>
'               <name>过敏源名称</name>
'               <code>过敏源代码</code>
'           </allergic>
'       </allergic_data>
'       <diagnose_data>
'        -!--诊断节点，可循环-->
'           <diagnose>
'               <type>诊断类型</type>
'               <name>诊断名称</name>
'               <code>诊断代码</code>
'           </diagnose>
'       </diagnose_data >
'       <lis_data>
'       <!-检验检测单节点，可循环-->
'           <form>
'               <no>检验、检查单号</no>
'               <project_name>检验项目名称</project_name>
'               <lis_flag>检验、检查标记 0-检验, 1-检查</lis_flag>
'               <result_date>检验、检查结果出具时间</result_date>
'               <sample_code>检验样本编码</sample_code>
'               <sample_name>检验样本名称</sample_name>
'               <mac_flag>微生物送检标识 0-否, 1-是</mac_flag>
'               <!-检验检测项目节点，可循环-->
'               <item>
'                   <code>检验、检查编码</code>
'                   <name>检验、检查名称</name>
'                   <value>检验、检查结果</value>
'                   <uom>结果值的单位</uom>
'                   <upper>结果参考范围上限</upper>
'                   <lower>结果参考范围下限</lower>
'               </item>
'           </form>
'       </lis_data>
'   </patient>
'   <prescription_data>
'       -!--处方信息节点，可循环-->
'       <prescription>
'           <id>处方号</id>
'           <reason>处方理由</reason>
'           <is_urgent>是否紧急处方(0:否1:是)</ is_urgent>
'           <is_new>是否新开处方(0:否1:是)</ is_new>
'           <is_current>是否当前处方（0/1）</is_current>
'           <doct_code>开嘱医生代码</doct_code>
'           <doct_name>开嘱医生姓名</doct_name>
'           <dept_code>开嘱科室代码</dept_code>
'           <dept_name>开嘱科室姓名</dept_name>
'           <pres_type>长期医嘱L/临时医嘱T</pres_type>
'           <pres_time>处方时间（YYYY-MM-DD HH:mm:SS）</pres_time>
'           <medicine_data>
'               -!--药品信息节点，可循环-->
'               <medicine>
'                   <name>商品名</name>
'                   <his_code>医院药品代码</his_code>
'                   <insur_code>医保代码</insur_code>
'                   <pyd_code>配液单号</pyd_code>
'                   <link_group>配液单组号</ link_group>
'                   <spec>规格</spec>
'                   <group>组号</group>
'                   <reason>用药理由</reason>
'                   <dose_unit>单次量单位</dose_unit>
'                   <dose>单次量</dose>
'                   <freq>频次代码</freq>
'                   <administer>给药途径代码</administer>
'                   <begin_time>（住院）用药开始时间(YYYY-MM-DD HH:mm:SS)</begin_time>
'                   <end_time>（住院）用药结束时间(YYYY-MM-DD HH:mm:SS)</end_time>
'                   <days>服药天数（门诊）</days>
'                   < preventiveflag >（住院）是否预防用药（1是，0否）        </preventiveflag >
'                   <otno>（住院）手术单号</otno>
'                   <signer_code>签名医师工号</signer_code>
'                   <accredit_date>授权时间</accredit_date>
'                   <accredit_hours>允许用药时间（小时）</accredit_hours >
'                   <accredit_times>允许用药次数</accredit_times>
'               </medicine>
'           </medicine_data>
'       </prescription>
'   </prescription_data>
'</details_xml>

'过敏源对象
Public Type DTBS_ALLERGIC
    str过敏类型 As String
    str过敏源名称 As String
    str过敏源代码 As String
End Type

'诊断对象
Public Type DTBS_DIAGNOSE
    str诊断类型 As String
    str诊断名称 As String
    str诊断代码 As String
End Type
'药品信息
Public Type DTBS_MEDICINE
    str商品名 As String
    str医院药品代码 As String
    str医保代码 As String
    str配液单号 As String
    str配液单组号 As String
    str规格 As String
    str组号 As String
    str用药理由 As String
    str单次量单位 As String
    str单次量 As String
    str频次代码 As String
    str给药途径代码 As String
    str用药开始时间 As String '住院
    str用药结束时间 As String '住院
    str服药天数 As String      '门诊
    str是否预防用药 As String '住院
    str手术单号 As String     '住院
    str签名医师工号 As String
    str授权时间 As String
    str允许用药时间 As String
    str允许用药次数 As String
End Type
'病人信息
Public Type DTBS_PATIENT
    str姓名 As String
    str是否婴儿 As String
    str出生日期 As String '格式：YYYY - MM - DD
    str性别 As String   '值域范围:男 / 女 / 未知
    str体重 As String '单位 kg
    str身高 As String '单位 cm
    str身份证号 As String
    str卡类型 As String
    str卡号 As String
    str怀孕时间单位 As String '天,周,月
    str怀孕时间 As String
    col过敏源s  As Collection
    col诊断s As Collection
    col检验检查 As Collection
End Type
'处方信息
Public Type DTBS_PRESCRIPTION
    str处方号 As String
    str处方理由 As String
    str是否当前处方 As String   '（0/1）
    str是否紧急处方 As String
    str是否新开处方 As String
    str开嘱医生代码 As String
    str开嘱医生姓名 As String
    str开嘱科室代码 As String
    str开嘱科室名称 As String
    Str医嘱类型 As String    '长期医嘱L/临时医嘱T
    str处方时间 As String    '（YYYY-MM-DD HH:mm:SS）
    col药品信息 As Collection     '药品的
End Type
'DETAILS_XML
Public Type DTBS_DETAILS
    str是否上传 As String
    strHIS系统时间 As String  '（YYYY-MM-DD HH:mm:SS）
    str门诊住院标识 As String ' op/ip
    str就诊类型 As String
    str就诊号 As String
    str检验就诊号 As String
    str床位号 As String
    str病区号 As String
    udt病人信息 As DTBS_PATIENT      '病人信息
    col处方信息  As Collection
    str药品名称 As String
    str药品代码 As String
End Type
'FORM_ITEM
Public Type DTBS_LISITEM
    str编码 As String  '（YYYY-MM-DD HH:mm:SS）
    str名称 As String ' op/ip
    str结果 As String
    str结果值单位 As String
    str参考范围上限 As String
    str参考范围下限 As String
End Type

'LIS_FORM
Public Type DTBS_LISFORM
    str单号 As String
    str项目名称 As String
    str标记 As String  '检验、检查标记 0-检验, 1-检查
    str结果出具时间 As String
    str检验样本编码 As String
    str检验样本名称 As String
    str微生物送检标识 As String
    col项目节点 As Collection    'DTBS_LISITEM 的集合
End Type

Public Enum DTBS_Func_NUM
    DTBS_登录 = 1
    DTBS_退出 = 2
    DTBS_初始UI = 3
    DTBS_要点提示 = 5
    DTBS_处方分析 = 6
    DTBS_作废处方 = 13
    DTBS_获取审方状态 = 14
End Enum
'<-end------------------------------------------------------------------------details_xml----------------------------------------------------->

