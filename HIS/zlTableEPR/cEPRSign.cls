VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarKey As Long                     '关键字，存储在“对象标记”中
Private mvar姓名 As String                  '签名用户的姓名，存储在“内容文本”中
Private mvar签名信息 As String              '数字签名字符串，存储在“要素值域”中
Private mvar前置文字 As String              '称呼或者顿号等前置文字，存储在“要素名称”中
Private mvar签名级别 As EPRSignLevel    '0,1,2,3，存储在“要素表示”中
Private mvar时间戳 As String                '时间戳，存储在“要素单位”中
Private mvar时间戳信息 As String            '电子签名返回的时间戳的认证编码信息(Base64)

'下面5个属性组成“对象属性”，存储在“对象属性”字段中
Private mvar签名方式 As Long                '1－密码签名；2－数字签名
Private mvar签名规则 As Long                '从1开始编号，由程序自定义其涵义（表示不同版本的签名方法）
                                            '目前我们使用的签名规则＝1
Private mvar证书ID As Long                  '数字证书ID号
Private mvar显示手签 As Boolean             '是否显示手签标志
Private mvar签名时间 As String              '签名时间，用于回退时恢复完成时间
Private mvar显示时间 As String              '是否显示签名时间及显示的格式串,空串表示不显示

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let 姓名(ByVal vData As String)
    mvar姓名 = vData
End Property

Public Property Get 姓名() As String
Attribute 姓名.VB_UserMemId = 0
    姓名 = mvar姓名
End Property

Public Property Let 签名信息(ByVal vData As String)
    mvar签名信息 = vData
End Property

Public Property Get 签名信息() As String
    签名信息 = mvar签名信息
End Property
Public Property Let 前置文字(ByVal vData As String)
    mvar前置文字 = vData
End Property

Public Property Get 前置文字() As String
    前置文字 = mvar前置文字
End Property

Public Property Let 签名级别(ByVal vData As EPRSignLevel)
    mvar签名级别 = vData
End Property

Public Property Get 签名级别() As EPRSignLevel
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
    
    '第6位预留,全文病历中用于存放签名要素Key值
    
    If UBound(T) >= 7 Then
        mvar时间戳信息 = CStr(T(7))
    Else
        mvar时间戳信息 = ""
    End If
    Err.Clear
End Property

Public Property Get 对象属性() As String
'签名方式|签名规则|证书ID|显示手签|签名时间|显示时间
    对象属性 = CStr(mvar签名方式) & ";" & CStr(mvar签名规则) & ";" & _
        CStr(mvar证书ID) & ";" & IIf(mvar显示手签, 1, 0) & ";" & _
        Format(mvar签名时间, "yyyy-mm-dd hh:mm:ss") & ";" & mvar显示时间 & ";;" & mvar时间戳信息
End Property

'################################################################################################################
'## 功能：  复制签名对象
'##
'## 参数：  blnWithKeys :是否连同Key值一起复制
'################################################################################################################
Public Sub Clone(cNode As cTabSign)
'功能：追加时复制子成员调用
    With cNode
         .姓名 = ""
         .签名信息 = ""
         .前置文字 = ""
         .签名级别 = 0
         .对象属性 = Me.签名方式 & ";1;;0;;;;"
         .时间戳 = ""
    End With
End Sub
Public Function GetSaveSQL(ByVal editType As TabET, ByVal CellTmp As cTabCell, arrSQL As Variant) As Boolean
'功能:生成签名保存SQL
'说明:对象属性 行,列|签名方式|签名规则|证书ID|显示手签|签名时间|显示时间
'    签名用户的姓名，存储在“内容文本”中
'    称呼或者顿号等前置文字，存储在“要素名称”中
'    时间戳，存储在“要素单位”中
'    签名级别，存储在“要素表示”中
'    数字签名字符串，存储在“要素值域”中
    On Error GoTo errHand
    With CellTmp
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        Select Case editType
        Case TabET_病历文件定义
            arrSQL(UBound(arrSQL)) = "Zl_病历文件结构_Update(" & .ID & "," & .文件ID & _
                                        ",Null," & .对象序号 & "," & .对象类型 & ",Null," & IIf(.保留对象, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|" & Me.对象属性 & "',Null,'" & mvar姓名 & "',0,Null,Null,Null" & _
                                        ",Null, Null,'" & mvar前置文字 & "',1,50,Null,'" & mvar时间戳 & "'," & mvar签名级别 & ",0,'" & mvar签名信息 & "')"
        Case TabET_全文示范编辑
            arrSQL(UBound(arrSQL)) = "Zl_病历范文内容_Update(" & .ID & "," & .文件ID & _
                                        ",Null," & .对象序号 & "," & .对象类型 & ",Null," & IIf(.保留对象, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|" & Me.对象属性 & "',Null,'" & mvar姓名 & "',0,Null,Null,Null,Null" & _
                                        ",Null, Null,'" & mvar前置文字 & "',1,50,Null,'" & mvar时间戳 & "'," & mvar签名级别 & ",0,'" & mvar签名信息 & "')"
                                        
        Case TabET_单病历编辑, TabET_单病历审核
            arrSQL(UBound(arrSQL)) = "Zl_电子病历内容_Update(" & .ID & "," & .文件ID & "," & .开始版 & "," & .终止版 & _
                                        ",Null," & .对象序号 & "," & .对象类型 & ",Null," & IIf(.保留对象, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|" & Me.对象属性 & "',Null,'" & mvar姓名 & "',0,Null,Null,Null,Null" & _
                                        ",Null, Null,'" & mvar前置文字 & "',1,50,Null,'" & mvar时间戳 & "'," & mvar签名级别 & ",0,'" & mvar签名信息 & "')"
        End Select
    End With
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
