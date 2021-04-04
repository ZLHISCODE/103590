VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTendSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'###############################################################################################
'##模 块 名：cTendSign.cls
'##创 建 人：吴庆伟
'##日    期：2005年11月24日
'##修 改 人：
'##日    期：
'##描    述：单个签名对象（对象类型＝8）
'##版    本：
'###############################################################################################

Option Explicit
'签名状态
Public Enum EPRSignLevelEnum
    cprSL_空白 = 0              '未签名
    cprSL_经治 = 1              '经治医师签名
    cprSL_主治 = 2              '主治医师签名
    cprSL_主任 = 3              '主任医师签名
    cprSL_正高 = 4              '正高：签名级别不包含，只表示人员居右正高职称，以便区别副主任医师
End Enum

Private mvarKey As Long                     '关键字，存储在“对象标记”中
Private mvarID As Long                      '用于保存的ID值
Private mvar文件ID As Long                  '文件ID
Private mvar父ID As Long                    '父ID
Private mvar开始版 As Long                  '开始版
Private mvar终止版 As Long                  '终止版
Private mvar对象序号 As Long                '每个签名按照文本顺序有个唯一序号（全文唯一）
Private mvar姓名 As String                  '签名用户的姓名，存储在“内容文本”中
Private mvar签名信息 As String              '数字签名字符串，存储在“要素值域”中
Private mvar前置文字 As String              '称呼或者顿号等前置文字，存储在“要素名称”中
Private mvar签名级别 As EPRSignLevelEnum    '0,1,2,3，存储在“要素表示”中
Private mvar签名人ID As Long                '记录签名人ID
Private mbln签名图片 As Boolean             '是否使用签名图片,存储在"输入形态"中

'下面5个属性组成“对象属性”，存储在“对象属性”字段中
Private mvar签名方式 As Long                '1－密码签名；2－数字签名
Private mvar签名规则 As Long                '从1开始编号，由程序自定义其涵义（表示不同版本的签名方法）
                                            '目前我们使用的签名规则＝1
Private mvar证书ID As Long                  '数字证书ID号
Private mvar显示手签 As Boolean             '是否显示手签标志
Private mvar签名时间 As String              '签名时间，用于回退时恢复完成时间
Private mvar显示时间 As String              '是否显示签名时间及显示的格式串,空串表示不显示
Private mvar签名要素 As String             '是否为签名要素Key值
Private mvar时间戳 As String                '时间戳，存储在“要素单位”中
Private mvar时间戳信息 As String            '电子签名返回的时间戳的认证编码信息(Base64)

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let ID(ByVal vData As Long)
    mvarID = vData
End Property

Public Property Get ID() As Long
    ID = mvarID
End Property

Public Property Let 文件ID(ByVal vData As Long)
    mvar文件ID = vData
End Property

Public Property Get 文件ID() As Long
    文件ID = mvar文件ID
End Property

Public Property Let 父ID(ByVal vData As Long)
    mvar父ID = vData
End Property

Public Property Get 父ID() As Long
    父ID = mvar父ID
End Property

Public Property Let 开始版(ByVal vData As Long)
    mvar开始版 = vData
End Property

Public Property Get 开始版() As Long
    开始版 = mvar开始版
End Property

Public Property Let 终止版(ByVal vData As Long)
    mvar终止版 = vData
End Property

Public Property Get 终止版() As Long
    终止版 = mvar终止版
End Property

Public Property Let 对象序号(ByVal vData As Long)
    mvar对象序号 = vData
End Property

Public Property Get 对象序号() As Long
    对象序号 = mvar对象序号
End Property

Public Property Let 姓名(ByVal vData As String)
    mvar姓名 = vData
End Property

Public Property Get 姓名() As String
Attribute 姓名.VB_UserMemId = 0
    姓名 = mvar姓名
End Property
Public Property Let 签名人ID(ByVal vData As Long)
    mvar签名人ID = vData
End Property

Public Property Get 签名人ID() As Long
    签名人ID = mvar签名人ID
End Property

Public Property Let 签名信息(ByVal vData As String)
    mvar签名信息 = vData
End Property

Public Property Get 签名信息() As String
    签名信息 = mvar签名信息
End Property

Public Property Let 签名要素(ByVal vData As String)
    mvar签名要素 = vData
End Property

Public Property Get 签名要素() As String
    签名要素 = mvar签名要素
End Property

Public Property Let 前置文字(ByVal vData As String)
    mvar前置文字 = vData
End Property

Public Property Get 前置文字() As String
    前置文字 = mvar前置文字
End Property

Public Property Let 签名级别(ByVal vData As EPRSignLevelEnum)
    mvar签名级别 = vData
End Property

Public Property Get 签名级别() As EPRSignLevelEnum
    签名级别 = mvar签名级别
End Property

Public Property Let 签名方式(ByVal vData As Long)
    mvar签名方式 = vData
End Property

Public Property Get 签名方式() As Long
    签名方式 = mvar签名方式
End Property

Public Property Let 签名规则(ByVal vData As Long)
    mvar签名规则 = vData
End Property

Public Property Get 签名规则() As Long
    签名规则 = mvar签名规则
End Property

Public Property Let 证书ID(ByVal vData As Long)
    mvar证书ID = vData
End Property

Public Property Get 证书ID() As Long
    证书ID = mvar证书ID
End Property

Public Property Let 签名时间(ByVal vData As String)
    mvar签名时间 = vData
End Property

Public Property Get 签名时间() As String
    签名时间 = mvar签名时间
End Property

Public Property Let 显示时间(ByVal vData As String)
    mvar显示时间 = vData
End Property

Public Property Get 显示时间() As String
    显示时间 = mvar显示时间
End Property

Public Property Let 时间戳(ByVal vData As String)
    mvar时间戳 = vData
End Property

Public Property Get 时间戳() As String
    时间戳 = mvar时间戳
End Property

Public Property Let 时间戳信息(ByVal vData As String)
    mvar时间戳信息 = vData
End Property

Public Property Get 时间戳信息() As String
    时间戳信息 = mvar时间戳信息
End Property

Public Property Let 显示手签(ByVal vData As Boolean)
    mvar显示手签 = vData
End Property

Public Property Get 显示手签() As Boolean
    显示手签 = mvar显示手签
End Property

Public Property Let 对象属性(ByVal vData As String)
    Dim T As Variant
    T = Split(vData, ";")
    On Error Resume Next
    mvar签名方式 = Val(T(0))
    mvar签名规则 = Val(T(1))
    mvar证书ID = Val(T(2))
    mvar显示手签 = (Val(T(3)) = 1)
    mvar签名时间 = Format(T(4), "yyyy-mm-dd hh:mm:ss")
    mvar显示时间 = T(5)
    If UBound(T) >= 6 Then
        mvar签名要素 = CStr(T(6))
    Else
        mvar签名要素 = ""
    End If
End Property

Public Property Get 对象属性() As String
    对象属性 = CStr(mvar签名方式) & ";" & CStr(mvar签名规则) & ";" & _
        CStr(mvar证书ID) & ";" & IIf(mvar显示手签, 1, 0) & ";" & _
        Format(mvar签名时间, "yyyy-mm-dd hh:mm:ss") & ";" & mvar显示时间 & ";" & mvar签名要素
End Property
