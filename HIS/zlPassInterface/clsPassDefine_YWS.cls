VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPassDefine_YWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------------------------------------------------------------
' --药卫士--- 接口定义     【广州市保进网络科技有限公司】
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
Public Type YWS_BASE
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
'4)  功能号 5 要点提示
'HIS命令 ：5
'<details_xml>
'<hosp_flag>门诊住院标识</hosp_flag>
'<medicine>
'        <his_code>药品代码</his_code>
'        <his_name>药品名称</his_name>
'</medicine>
'</details_xml>

'5)  功能号 6 、8、9处方分析
'HIS命令 ：6、8、9
'<details_xml >
'    <his_time>HIS系统时间（YYYY-MM-DD HH:mm:SS）</his_time>
'    <hosp_flag>门诊住院标识 op/ip</hosp_flag>
'    <treat_type>就诊类型</treat_type>
'    <treat_code>就诊号</treat_code>
'    <bed_no>床位号</bed_no>
'    <patient>
'        <name>姓名</name>
'        <birth>出生日期(YYYY-MM-DD)</birth>
'        <sex>性别（男/女/未知）</sex>
'        <weight>体重（单位：千克）</weight>
'        <height>身高（单位：厘米）</height>
'        <id_card>身份证号</id_card>
'        <medical_record>病历卡号</medical_record>
'        <card_type>卡类型</card_type>
'        <card_code>卡号</card_code>
'        <pregnant_unit>时间单位（天、周、月）</pregnant_unit>
'        <pregnant >怀孕时间</pregnant>
'        <allergic_data>
'            <!--过敏节点，可循环-->
'            <allergic>
'                <type>过敏类型</type>
'                <name>过敏源名称</name>
'                <code>过敏源代码</code>
'            </allergic>
'        </allergic_data>
'        <diagnose_data>
'            <!--诊断节点，可循环-->
'            <diagnose>
'                <type>诊断类型</type>
'                <name>诊断名称</name>
'                <code>诊断代码</code>
'            </diagnose>
'       </diagnose_data >
'    </patient>
'    <prescription_data>
'        <prescription>
'            <id>处方号</id>
'            <reason>处方理由</reason>
'            <is_current>是否当前处方（0/1）</is_current>
'            <pres_type>长期医嘱L/临时医嘱T</pres_type>
'            <pres_time>处方时间（YYYY-MM-DD HH:mm:SS）</pres_time>
'            <medicine_data>
'           <!--药品信息节点，可循环-->
'                <medicine>
'                     <zxy_type>西药/中成药/草药</zxy_type>
'                     <oeridid>处方号-流水号(唯一)</oeridid>
'                     <pres_type>长期医嘱L/临时医嘱T</pres_type>
'                     <pres_time>处方时间（YYYY-MM-DD HH:mm:SS）</pres_time>
'                     <name>商品名</name>
'                     <his_code>医院药品代码</his_code>
'                     <insur_code>医保代码</insur_code>
'                     <approval>批准文号</approval>
'                     <spec>规格</spec>
'                     <group>组号</group>
'                     <reason>用药理由</reason>
'                     <dose_unit>单次量单位</dose_unit>
'                     <dose>单次量</dose>
'                     <freq>频次代码</freq>
'                     <administer>给药途径代码</administer>
'                     <begin_time>（住院）用药开始时间(YYYY-MM-DD HH:mm:SS)</begin_time>
'                     <end_time>（住院）用药结束时间(YYYY-MM-DD HH:mm:SS)</end_time>
'                     <days>服药天数（门诊）</days>
'                </medicine>
'           </medicine_data>
'       </prescription>
'   </prescription_data>
'</details_xml>
'
'过敏源对象
Public Type YWS_ALLERGIC
    str过敏类型 As String
    str过敏源名称 As String
    str过敏源代码 As String
End Type

'诊断对象
Public Type YWS_DIAGNOSE
    str诊断类型 As String
    str诊断名称 As String
    str诊断代码 As String
End Type
'药品信息
Public Type YWS_MEDICINE
    str药品类型 As String
    str处方号 As String
    Str医嘱类型 As String
    str处方时间 As String
    str商品名 As String
    str医院药品代码 As String
    str医保代码 As String
    str批准文号 As String
    str规格 As String
    str组号 As String
    str用药理由 As String
    str单次量单位 As String
    str单次量 As String
    str频次代码 As String
    str给药途径代码 As String
    str用药开始时间 As String
    str用药结束时间 As String
    str服药天数 As String      '门诊
End Type
'病人信息
Public Type YWS_PATIENT
    str姓名 As String
    str出生日期 As String '格式：YYYY - MM - DD
    str性别 As String   '值域范围:男 / 女 / 未知
    str体重 As String '单位 kg
    str身高 As String '单位 cm
    str身份证号 As String
    str病历卡号 As String
    str卡类型 As String
    str卡号 As String
    str怀孕时间单位 As String '天,周,月
    str怀孕时间 As String
    col过敏源s  As Collection
    col诊断s As Collection
End Type
'处方信息
Public Type YWS_PRESCRIPTION
    str处方号 As String
    str处方理由 As String
    str是否当前处方 As String   '（0/1）
    Str医嘱类型 As String
    str处方时间 As String    '（YYYY-MM-DD HH:mm:SS）
    col药品信息 As Collection     '药品的
End Type
'DETAILS_XML
Public Type YWS_DETAILS
    str门诊住院标识 As String ' op/ip
    str药品代码 As String
    str药品名称 As String
    strHIS系统时间 As String  '（YYYY-MM-DD HH:mm:SS）
    str就诊类型 As String
    str就诊号 As String
    str床位号 As String
    udt病人信息 As YWS_PATIENT      '病人信息
    udt处方信息 As YWS_PRESCRIPTION      '处方信息
End Type

Public Enum YWS_Func_NUM
    YWS_登录 = 1
    YWS_退出 = 3
    YWS_初始客户端 = 4
    YWS_要点提示 = 5
    YWS_处方分析 = 6
    YWS_上传处方 = 8
    YWS_处方分析仅亮灯 = 9
    YWS_调试调用 = 10
    YWS_处方分析并返回结果 = 11
End Enum
'<-end------------------------------------------------------------------------details_xml----------------------------------------------------->
