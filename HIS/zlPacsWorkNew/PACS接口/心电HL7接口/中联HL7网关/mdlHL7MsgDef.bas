Attribute VB_Name = "mdlHL7MsgDef"
Option Explicit

'本模块主要是HL7消息结构的定义

'HL7字段定义
Public Type THL7Field
    intNo As Integer            '序号，按照序号排序
    strDataType As String       '数据类型，HL7中定义的类型
    strRecDataDef As String     '接收数据定义
    strSendDataDef As String    '发送数据定义
    strRecDataValue As String   '接收数据值
    strSendDataValue As String  '发送数据值
    strElementName As String    '元素名称
    blnEnable As Boolean        '可用性
End Type

'HL7消息段定义
Public Type THL7Segment
    intNo As Integer            '消息段的序号，按照序号排序
    arrFields() As THL7Field    '段中字段定义
    strName As String           '消息段名称
    strText As String           '消息段文本，接收或者发送的文本
    blnEnable As Boolean        '可用性
End Type

'HL7消息
Public Type THL7Message
    lngID As Long                   '消息ID
    arrSegments() As THL7Segment    '消息段定义
    strMsgName As String            '消息名称
    lngServiceID As Long            '服务ID
    strActionType As String         '动作类型
    strMsgType As String            '消息类型
    strMsgSegmentDef As String      '消息段组合定义
    strText As String               '消息文本
    strIP As String                 '接收消息的IP地址
    lngPort As Long                 '接收消息的端口号
    blnSendOK As Boolean            '发送消息成功，接收到AA的响应
End Type

Public Type THl7Messages
    arrMsgs() As THL7Message        '消息数组
    lngActionID As Long             'HL7待发消息记录的ID
End Type

'常量定义
Public Const HL7_MSG_SEND_NEW_ORDER = "发送新医嘱"
Public Const HL7_MSG_SEND_CANCEL_ORDER = "发送取消医嘱"
Public Const HL7_MSG_SEND_DEL_ORDER = "发送删除医嘱"

