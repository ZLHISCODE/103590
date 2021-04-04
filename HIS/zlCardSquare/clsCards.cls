VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsCard"
Attribute VB_Ext_KEY = "Member0" ,"clsCard"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'局部变量，保存集合
Private mCol As Collection
Public Function AddItem(卡种类 As gCard_Type, 接口序号 As Long, 接口编码 As String, 名称 As String, 短名 As String, _
    启用 As Boolean, 消费卡 As Boolean, 接口程序名 As String, 自制卡 As Boolean, 可否设置 As Integer, _
    是否自动读取 As Integer, 自动读取间隔 As Integer, 系统 As Boolean, 结算方式 As String, 前缀文本 As String, _
    卡号长度 As Long, 是否严格控制 As Boolean, 是否刷卡 As Boolean, 是否存在帐户 As Boolean, 是否全退 As Boolean, _
    特定项目 As String, 备注 As String, 卡号重复使用 As Boolean, 卡号密文规则 As String, 是否退现 As Boolean, _
    密码长度 As Integer, 密码长度限制 As Integer, 密码规则 As Integer, _
    Optional sKey As String, Optional 图像标识 As String = "", _
    Optional 支付启用 As Boolean = False, Optional 支付图像标识 As String = "", _
    Optional 是否制卡 As Boolean = False, Optional 是否发卡 As Boolean = False, Optional 是否写卡 As Boolean = False, _
    Optional 险类 As Integer = 0, Optional 险类名称 As String = "", Optional 键盘类型 As gKeyBoard_Type = EM_KeyBoard_NotKey, _
    Optional 结算性质 As Integer, Optional 是否模糊查找 As Boolean = False, Optional 是否转帐及代扣 As Boolean = False, _
    Optional 是否扫描 As Boolean, Optional 是否接触式读卡 As Boolean, Optional 是否非接触式读卡 As Boolean, _
    Optional 是否持卡消费 As Boolean = False, Optional 发送调用接口 As Boolean = False, _
    Optional 是否退款验卡 As Boolean = False, Optional 是否证件 As Boolean = False, _
    Optional 设备是否启用回车 As Boolean = False, Optional 是否缺省退现 As Boolean, Optional 是否缺省 As Boolean) As clsCard
   '--卡种类:1-消费卡;2-结算卡;3-现金;4-帐户;5-险类
    '创建新对象
    Dim objNewMember As clsCard
    Set objNewMember = New clsCard
    
    '设置传入方法的属性
    objNewMember.接口序号 = 接口序号
    objNewMember.接口编码 = 接口编码
    objNewMember.短名 = 短名
    objNewMember.名称 = 名称
    objNewMember.启用 = 启用
    objNewMember.接口程序名 = 接口程序名
    objNewMember.可否设置 = 可否设置
    objNewMember.是否自动读取 = 是否自动读取
    objNewMember.自动读取间隔 = 自动读取间隔
    objNewMember.系统 = 系统
    objNewMember.结算方式 = 结算方式
    objNewMember.前缀文本 = 前缀文本
    objNewMember.卡号长度 = 卡号长度
    objNewMember.自制卡 = 自制卡
    objNewMember.消费卡 = 消费卡
    objNewMember.是否严格控制 = 是否严格控制
    objNewMember.特定项目 = 特定项目
    objNewMember.备注 = 备注
    objNewMember.是否存在帐户 = 是否存在帐户
    objNewMember.是否全退 = 是否全退
    objNewMember.卡号重复使用 = 卡号重复使用
    objNewMember.卡号密文规则 = 卡号密文规则
    objNewMember.是否退现 = 是否退现
    objNewMember.密码长度 = 密码长度
    objNewMember.密码长度限制 = 密码长度限制
    objNewMember.密码规则 = 密码规则
    objNewMember.图像标识 = 图像标识
    objNewMember.支付启用 = 支付启用
    objNewMember.支付图像标识 = 支付图像标识
    objNewMember.是否制卡 = 是否制卡    '56615
    objNewMember.是否发卡 = 是否发卡
    objNewMember.是否写卡 = 是否写卡
    objNewMember.险类 = 险类
    objNewMember.险类名称 = 险类名称
    objNewMember.键盘类型 = 键盘类型    '60572:李南春,2014-6-20,为费用查询增加手工录入的身份识别方式
    objNewMember.结算性质 = 结算性质
    objNewMember.是否模糊查找 = 是否模糊查找
    objNewMember.是否转帐及代扣 = 是否转帐及代扣    '77872,李南春,2014/10/28:是否支持转帐及代扣
    objNewMember.是否刷卡 = 是否刷卡    '85565,李南春,2015/7/10:读卡性质
    objNewMember.是否扫描 = 是否扫描
    objNewMember.是否接触式读卡 = 是否接触式读卡
    objNewMember.是否非接触式读卡 = 是否非接触式读卡
    objNewMember.是否持卡消费 = 是否持卡消费
    objNewMember.发送调用接口 = 发送调用接口
    objNewMember.是否退款验卡 = 是否退款验卡
    objNewMember.是否证件 = 是否证件
    objNewMember.设备是否启用回车 = 设备是否启用回车
    objNewMember.是否缺省退现 = 是否缺省退现
    objNewMember.缺省标志 = 是否缺省
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
   '返回已创建的对象
    Set AddItem = objNewMember
    Set objNewMember = Nothing
End Function
Public Function Add(objCard As clsCard, Optional sKey As String) As clsCard
    If Len(sKey) = 0 Then
       mCol.Add objCard
    Else
        mCol.Add objCard, sKey
    End If
    Set Add = objCard
End Function

Public Property Get Item(vntIndexKey As Variant) As clsCard
Attribute Item.VB_UserMemId = 0
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get count() As Long
    '检索集合中的元素数时使用。语法：Debug.Print x.Count
    count = mCol.count
End Property


Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '创建类后创建集合
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
End Sub

