VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRPatiRecordInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'记录信息
Public ID As Long                          '记录ID
Public 病人来源 As PatiFromEnum            '1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
Public 病人ID As Long                      '
Public 主页ID As Long                      '住院病人必须填写
Public 婴儿 As Long                        '如果是婴儿，显示时将姓名处理为**之婴
Public 科室ID As Long                      '就诊科室,记录书写本次病历的医生或护士所在科室
Public 病历种类 As EPRDocTypeEnum          '对应于病历文件的种类
Public 文件ID As Long                      '如果对应的病历文件已经删除,则格式按缺省处理。
Public 病历名称 As String                  '
Public 创建人 As String
Public 创建时间 As Date
Public 保存人 As String
Public 保存时间 As Date
Public 归档人 As String
Public 归档日期 As Date
Public 完成时间 As Date                    '在签名时填写,平常产生新的版本保存时不填写,表示正在进行中
Public 最后版本 As Integer
Public 签名级别 As EPRSignLevelEnum        '1-书写;2-主治医师审阅;3-主任医师审阅。住院病历以外的病历只有书写和审阅状态
Public 编辑方式 As Integer                  ' 0-使用全文式电子病历编辑器编辑；1-使用表格式病历编辑器编辑；2-门诊病历快捷输入(2兼容0)

'补充：
Public 医嘱id As Long                      '医嘱ID，用于特殊病历！

Private mvar书写签名 As String
Private mvar医生签名 As String
Private mvar主治签名 As String
Private mvar主任签名 As String

Public Property Get 书写部门() As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String

    书写部门 = ""
    Err = 0: On Error GoTo errHand
    strSQL = "Select 名称 From 部门表 Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "cEPRPatiRecordInfo", Me.科室ID)
    If rsTemp.RecordCount > 0 Then 书写部门 = Trim(rsTemp.Fields(0).Value)
    rsTemp.Close
    Set rsTemp = Nothing
    
    Exit Property
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err = 0
End Property
Public Property Let 医生签名(ByVal vData As String)
    mvar医生签名 = vData
End Property

Public Property Get 医生签名() As String
    医生签名 = mvar医生签名
End Property
Public Property Let 主治签名(ByVal vData As String)
    mvar主治签名 = vData
End Property

Public Property Get 主治签名() As String
    主治签名 = mvar主治签名
End Property
Public Property Let 主任签名(ByVal vData As String)
    mvar主任签名 = vData
End Property

Public Property Get 主任签名() As String
    主任签名 = mvar主任签名
End Property
Public Property Let 书写签名(ByVal vData As String)
    mvar书写签名 = vData
End Property

Public Property Get 书写签名() As String

    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String

    If mvar书写签名 = "" Then
    
        Err = 0
        On Error GoTo errHand
        
        strSQL = "Select 内容文本 From 电子病历内容 Where 文件id = [1] And 对象类型 = 8 And 开始版 = 1"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "cEPRPatiRecordInfo", Me.ID)
        If rsTemp.RecordCount > 0 Then mvar书写签名 = Trim(rsTemp.Fields(0).Value)
        rsTemp.Close
        Set rsTemp = Nothing

    End If
    
    书写签名 = mvar书写签名
    

    Exit Property
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err = 0

End Property
