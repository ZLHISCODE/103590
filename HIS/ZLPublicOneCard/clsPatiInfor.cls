VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPatiInfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'保持属性值的局部变量
Private mvar姓名 As String '局部复制
Private mvar性别 As String '局部复制
Private mvar出生日期 As String '局部复制
Private mvar身份证号 As String '局部复制
Private mvar病人ID As Long '局部复制
Private mvar门诊号 As String '局部复制
Private mvar民族 As String '局部复制
Private mvar家庭地址 As String '局部复制
Private mvar联系电话 As String '局部复制
Private mvar密码 As String '局部复制
Private mvar卡余额  As Double   '局部复制
Private mvar费别 As String '局部复制
Private mvar医疗付款方式 As String '局部复制
Private mvar医疗付款方式编码 As String '局部复制
Private mvar年龄 As String '局部复制
Private mvar卡号 As String '局部复制
Private mvar出生地址 As String '局部复制
Private mvar其他证件 As String '局部复制
Private mvar职业 As String '局部复制
Private mvar国籍 As String '局部复制
Private mvar学历 As String '局部复制
Private mvar婚姻状况 As String '局部复制
Private mvar区域 As String '局部复制
Private mvar家庭电话 As String '局部复制
Private mvar家庭邮编 As String '局部复制
Private mvar联系人关系 As String '局部复制
Private mvar联系人 As String '局部复制
Private mvar联系人地址 As String '局部复制
Private mvar联系人电话 As String '局部复制
Private mvar工作单位 As String '局部复制
Private mvar工作单位地址 As String '局部复制
Private mvar工作单位电话 As String '局部复制
Private mvar工作单位邮编 As String '局部复制
Private mvar工作单位开户行 As String '局部复制
Private mvar工作单位开户行帐户 As String '局部复制
Private mvar户口地址 As String '局部复制
Private mvar户口地址邮编 As String '局部复制
Private mvar籍贯 As String '局部复制
Private mvar监护人 As String  '局部复制
Private mvar消费卡ID As Long    '局部复制
Private mvar险类 As Long    '局部复制
Private mvar病人类型 As String '局部复制
Private mvar就诊卡号 As String
Private mvar手机号 As String
Private mvar照片 As StdPicture
Private mvar照片文件 As String
Private mvar险类名称 As String '局部复制
Private mvarTag As String '局部复制
Private mvarPatiXML As String '局部复制

Public Property Let PatiXML(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.险类 = 5
    mvarPatiXML = vData
End Property
Public Property Get PatiXML() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.险类
    PatiXML = mvarPatiXML
End Property
Public Property Let Tag(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Tag = 5
    mvarTag = vData
End Property


Public Property Get Tag() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Tag
    Tag = mvarTag
End Property



Public Property Let 险类名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.险类名称 = 5
    mvar险类名称 = vData
End Property


Public Property Get 险类名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.险类名称
    险类名称 = mvar险类名称
End Property




Public Property Let 照片文件(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.照文件片 = 5
    mvar照片文件 = vData
End Property
Public Property Get 照片文件() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.照片文件
    照片文件 = mvar照片文件
End Property

Public Property Let 照片(ByVal vData As StdPicture)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.照片 = 5
    Set mvar照片 = vData
End Property
Public Property Get 照片() As StdPicture
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.照片
    Set 照片 = mvar照片
End Property

Public Property Let 手机号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.手机号 = 5
    mvar手机号 = vData
End Property
Public Property Get 手机号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.手机号
    手机号 = mvar手机号
End Property

Public Property Let 就诊卡号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.籍贯 = 5
    mvar就诊卡号 = vData
End Property
Public Property Get 就诊卡号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.籍贯
    就诊卡号 = mvar就诊卡号
End Property

Public Property Let 险类(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.险类 = 5
    mvar险类 = vData
End Property
Public Property Get 险类() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.险类
    险类 = mvar险类
End Property
Public Property Let 消费卡ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.消费卡ID = 5
    mvar消费卡ID = vData
End Property
Public Property Get 消费卡ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.消费卡ID
    消费卡ID = mvar消费卡ID
End Property

Public Property Let 卡余额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.卡余额 = 5
    mvar卡余额 = vData
End Property
Public Property Get 卡余额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.卡余额
    卡余额 = mvar卡余额
End Property
Public Property Let 籍贯(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.籍贯 = 5
    mvar籍贯 = vData
End Property
Public Property Get 籍贯() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.籍贯
    籍贯 = mvar籍贯
End Property
Public Property Let 户口地址邮编(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.户口地址邮编 = 5
    mvar户口地址邮编 = vData
End Property
Public Property Get 户口地址邮编() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.户口地址邮编
    户口地址邮编 = mvar户口地址邮编
End Property

Public Property Let 病人类型(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人类型 = 5
    mvar病人类型 = vData
End Property
Public Property Get 病人类型() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.病人类型
    病人类型 = mvar病人类型
End Property
Public Property Let 户口地址(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.户口地址 = 5
    mvar户口地址 = vData
End Property
Public Property Get 户口地址() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.户口地址
    户口地址 = mvar户口地址
End Property

Public Property Let 工作单位开户行帐户(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.工作单位开户行帐户 = 5
    mvar工作单位开户行帐户 = vData
End Property
Public Property Get 工作单位开户行帐户() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.工作单位开户行帐户
    工作单位开户行帐户 = mvar工作单位开户行帐户
End Property
Public Property Let 工作单位开户行(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.工作单位开户行 = 5
    mvar工作单位开户行 = vData
End Property
Public Property Get 工作单位开户行() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.工作单位开户行
    工作单位开户行 = mvar工作单位开户行
End Property

Public Property Let 工作单位邮编(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.工作单位邮编 = 5
    mvar工作单位邮编 = vData
End Property
Public Property Get 工作单位邮编() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.工作单位邮编
    工作单位邮编 = mvar工作单位邮编
End Property

Public Property Let 工作单位电话(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.工作单位电话 = 5
    mvar工作单位电话 = vData
End Property
Public Property Get 工作单位电话() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.工作单位电话
    工作单位电话 = mvar工作单位电话
End Property

Public Property Let 工作单位地址(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.工作单位地址 = 5
    mvar工作单位地址 = vData
End Property
Public Property Get 工作单位地址() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.工作单位地址
    工作单位地址 = mvar工作单位地址
End Property

Public Property Let 工作单位(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.工作单位 = 5
    mvar工作单位 = vData
End Property
Public Property Get 工作单位() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.工作单位
    工作单位 = mvar工作单位
End Property

Public Property Let 联系人电话(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.联系人电话 = 5
    mvar联系人电话 = vData
End Property
Public Property Get 联系人电话() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.联系人电话
    联系人电话 = mvar联系人电话
End Property

Public Property Let 联系人地址(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.联系人地址 = 5
    mvar联系人地址 = vData
End Property
Public Property Get 联系人地址() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.联系人地址
    联系人地址 = mvar联系人地址
End Property

Public Property Let 联系人(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.联系人 = 5
    mvar联系人 = vData
End Property
Public Property Get 联系人() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.联系人
    联系人 = mvar联系人
End Property


Public Property Let 联系人关系(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.联系人关系 = 5
    mvar联系人关系 = vData
End Property
Public Property Get 联系人关系() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.联系人关系
    联系人关系 = mvar联系人关系
End Property

Public Property Let 监护人(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.监护人 = 5
    mvar监护人 = vData
End Property
Public Property Get 监护人() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.监护人
    监护人 = mvar监护人
End Property
Public Property Let 家庭邮编(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.家庭邮编 = 5
    mvar家庭邮编 = vData
End Property
Public Property Get 家庭邮编() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.家庭邮编
    家庭邮编 = mvar家庭邮编
End Property

Public Property Let 家庭电话(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.家庭电话 = 5
    mvar家庭电话 = vData
End Property
Public Property Get 家庭电话() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.家庭电话
    家庭电话 = mvar家庭电话
End Property

Public Property Let 区域(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.区域 = 5
    mvar区域 = vData
End Property
Public Property Get 区域() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.区域
    区域 = mvar区域
End Property


Public Property Let 婚姻状况(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.婚姻状况 = 5
    mvar婚姻状况 = vData
End Property
Public Property Get 婚姻状况() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.婚姻状况
    婚姻状况 = mvar婚姻状况
End Property

Public Property Let 学历(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.学历 = 5
    mvar学历 = vData
End Property
Public Property Get 学历() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.学历
    学历 = mvar学历
End Property
Public Property Let 国籍(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.国籍 = 5
    mvar国籍 = vData
End Property
Public Property Get 国籍() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.国籍
    国籍 = mvar国籍
End Property


Public Property Let 职业(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.职业 = 5
    mvar职业 = vData
End Property
Public Property Get 职业() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.职业
    职业 = mvar职业
End Property


Public Property Let 其他证件(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.其他证件 = 5
    mvar其他证件 = vData
End Property
Public Property Get 其他证件() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.其他证件
    其他证件 = mvar其他证件
End Property

Public Property Let 出生地址(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.出生地址 = 5
    mvar出生地址 = vData
End Property
Public Property Get 出生地址() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.出生地址
    出生地址 = mvar出生地址
End Property

Public Property Let 卡号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.卡号 = 5
    mvar卡号 = vData
End Property
Public Property Get 卡号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.卡号
    卡号 = mvar卡号
End Property

Public Property Let 年龄(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.联系电话 = 5
    mvar年龄 = vData
End Property
Public Property Get 年龄() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.联系电话
    年龄 = mvar年龄
End Property

Public Property Let 医疗付款方式(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.联系电话 = 5
    mvar医疗付款方式 = vData
End Property
Public Property Get 医疗付款方式() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.联系电话
    医疗付款方式 = mvar医疗付款方式
End Property


Public Property Let 医疗付款方式编码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.联系电话 = 5
    mvar医疗付款方式编码 = vData
End Property
Public Property Get 医疗付款方式编码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.联系电话
    医疗付款方式编码 = mvar医疗付款方式编码
End Property

Public Property Let 费别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.联系电话 = 5
    mvar费别 = vData
End Property
Public Property Get 费别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.联系电话
    费别 = mvar费别
End Property

Public Property Let 密码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.联系电话 = 5
    mvar密码 = vData
End Property
Public Property Get 密码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.联系电话
    密码 = mvar密码
End Property

Public Property Let 联系电话(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.联系电话 = 5
    mvar联系电话 = vData
End Property
Public Property Get 联系电话() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.联系电话
    联系电话 = mvar联系电话
End Property
Public Property Let 门诊号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊号 = 5
    mvar门诊号 = vData
End Property


Public Property Get 门诊号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.门诊号
    门诊号 = mvar门诊号
End Property



Public Property Let 病人ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mvar病人ID = vData
End Property


Public Property Get 病人ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.病人ID
    病人ID = mvar病人ID
End Property



Public Property Let 身份证号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.身份证号 = 5
    mvar身份证号 = vData
End Property


Public Property Get 身份证号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.身份证号
    身份证号 = mvar身份证号
End Property



Public Property Let 出生日期(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.出生日期 = 5
    mvar出生日期 = vData
End Property


Public Property Get 出生日期() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.出生日期
    出生日期 = mvar出生日期
End Property



Public Property Let 性别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.性别 = 5
    mvar性别 = vData
End Property


Public Property Get 性别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.性别
    性别 = mvar性别
End Property



Public Property Let 姓名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.姓名 = 5
    mvar姓名 = vData
End Property


Public Property Get 姓名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.姓名
    姓名 = mvar姓名
End Property

Public Property Let 民族(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.民族 = 5
    mvar民族 = vData
End Property
Public Property Get 民族() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.民族
    民族 = mvar民族
End Property
Public Property Let 家庭地址(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.家庭地址 = 5
    mvar家庭地址 = vData
End Property
Public Property Get 家庭地址() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.家庭地址
    家庭地址 = mvar家庭地址
End Property


