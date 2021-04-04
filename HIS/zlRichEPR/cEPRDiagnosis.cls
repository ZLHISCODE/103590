VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRDiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarKey As Long         '对象标记
Private mvarID As Double          'ID
Private mvar文件ID As Long      '文件ID
Private mvar父ID As Double       '父ID：诊断所属的提纲
Private mvar对象序号 As Double    '对象序号

Private mvar类型 As Long        '诊断类型：11-门诊诊断;21-入院初步诊断;22-入院确诊诊断;23-入院修正诊断;24-入院补充诊断；31-出院诊断;41-术前诊断;42-术后诊断
Private mvar中医 As Byte        '是否中医：0-西医;1-中医
Private mvar疾病id As Long      '对应疾病编码目录(ICD或中医疾病)的ID
Private mvar诊断id As Long      '对应疾病诊断目录的ID
Private mvar证候id As Long      '对应疾病诊断目录的ID
Private mvar描述 As String      '疾病诊断的描述，中医诊断格式为“病(证)”，如：中风(痰热腑实，风痰上扰)
Private mvar疑诊 As Byte        '是否疑诊：0-确诊;1-疑诊
Private mvar日期 As String
Private mvar入院病情 As String  '出院诊断同步首页时使用
Private mvar出院情况 As String  '出院诊断同步首页时使用

Private mvar开始版 As Long      '开始版，专用于“电子病历记录”
Private mvar终止版 As Long      '终止版，专用于“电子病历记录”

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let ID(ByVal vData As Double)
    mvarID = vData
End Property

Public Property Get ID() As Double
    ID = mvarID
End Property

Public Property Let 文件ID(ByVal vData As Long)
    mvar文件ID = vData
End Property

Public Property Get 文件ID() As Long
    文件ID = mvar文件ID
End Property

Public Property Let 父ID(ByVal vData As Double)
    mvar父ID = vData
End Property

Public Property Get 父ID() As Double
    父ID = mvar父ID
End Property

Public Property Let 对象序号(ByVal vData As Double)
    mvar对象序号 = vData
End Property

Public Property Get 对象序号() As Double
    对象序号 = mvar对象序号
End Property

Public Property Let 类型(ByVal vData As Long)
    mvar类型 = vData
End Property

Public Property Get 类型() As Long
    类型 = mvar类型
End Property

Public Property Let 中医(ByVal vData As Byte)
    mvar中医 = vData
End Property

Public Property Get 中医() As Byte
    中医 = mvar中医
End Property

Public Property Let 疾病id(ByVal vData As Long)
    mvar疾病id = vData
End Property

Public Property Get 疾病id() As Long
    疾病id = mvar疾病id
End Property

Public Property Let 诊断id(ByVal vData As Long)
    mvar诊断id = vData
End Property

Public Property Get 诊断id() As Long
    诊断id = mvar诊断id
End Property

Public Property Let 证候id(ByVal vData As Long)
    mvar证候id = vData
End Property

Public Property Get 证候id() As Long
    证候id = mvar证候id
End Property

Public Property Let 描述(ByVal vData As String)
    mvar描述 = vData
End Property

Public Property Get 描述() As String
    描述 = mvar描述
End Property

Public Property Let 疑诊(ByVal vData As Byte)
    mvar疑诊 = vData
End Property

Public Property Get 疑诊() As Byte
    疑诊 = mvar疑诊
End Property

Public Property Let 日期(ByVal vData As String)
    mvar日期 = vData
End Property

Public Property Get 日期() As String
    日期 = mvar日期
End Property

Public Property Let 入院病情(ByVal vData As String)
    mvar入院病情 = vData
End Property

Public Property Get 入院病情() As String
    入院病情 = mvar入院病情
End Property
Public Property Let 出院情况(ByVal vData As String)
    mvar出院情况 = vData
End Property

Public Property Get 出院情况() As String
    出院情况 = mvar出院情况
End Property

Public Property Let 开始版(ByVal vData As Long)
    mvar开始版 = vData
End Property

Public Property Get 开始版() As Long
    开始版 = IIf(mvar开始版 = 0, 1, mvar开始版)
End Property

Public Property Let 终止版(ByVal vData As Long)
    mvar终止版 = vData
End Property

Public Property Get 终止版() As Long
    终止版 = mvar终止版
End Property


'################################################################################################################
'## 功能：  复制对象
'##
'## 参数：  blnWithKeys :是否连同Key值一起复制
'################################################################################################################
Public Function Clone(Optional blnWithKey As Boolean = False) As cEPRDiagnosis
    '用于拷贝数据，产生一个新类对象
    Dim cNode As New cEPRDiagnosis
    If blnWithKey Then cNode.Key = Me.Key
    cNode.ID = Me.ID
    cNode.文件ID = Me.文件ID
    cNode.开始版 = Me.开始版
    cNode.终止版 = Me.终止版
    cNode.父ID = Me.父ID
    cNode.对象序号 = Me.对象序号
    
    cNode.类型 = Me.类型
    cNode.中医 = Me.中医
    cNode.疾病id = Me.疾病id
    cNode.诊断id = Me.诊断id
    cNode.证候id = Me.证候id
    cNode.描述 = Me.描述
    cNode.疑诊 = Me.疑诊
    cNode.日期 = Me.日期
    cNode.入院病情 = Me.入院病情
    cNode.出院情况 = Me.出院情况

    Set Clone = cNode
End Function
Public Function FillDiagnosisMember(ByVal rs As ADODB.Recordset, ByVal strTableName As String) As Boolean
'功能：用数据记录填充类成员
Dim aryProperty() As String

    If rs Is Nothing Then Exit Function
    If rs.EOF Then Exit Function
    
    If NVL(rs!对象标记, 0) <> 0 Then Me.Key = Val("" & rs!对象标记)
    Me.ID = rs!ID
    Me.文件ID = Val("" & rs!文件ID)
    Me.父ID = Val("" & rs!父ID)
    Me.对象序号 = Val("" & rs!对象序号)
    Me.描述 = "" & rs!内容文本
    aryProperty = Split("" & rs!对象属性, ";")
    Me.类型 = Val(aryProperty(0))
    Me.中医 = Val(aryProperty(1))
    Me.疾病id = Val(aryProperty(2))
    Me.诊断id = Val(aryProperty(3))
    Me.证候id = Val(aryProperty(4))
    Me.疑诊 = Val(aryProperty(5))
    Me.日期 = Format(aryProperty(6), "yyyy-mm-dd hh:mm:ss")
    If UBound(aryProperty) > 6 Then
        Me.入院病情 = aryProperty(7)
        Me.出院情况 = aryProperty(8)
    End If

    If InStr(strTableName, "电子病历内容") > 0 Then '病历文件结构和病历范文内容没有这两个字段,但都共用这个函数
        Me.开始版 = NVL(rs!开始版, 1)
        Me.终止版 = NVL(rs!终止版, 0)
    End If

    FillDiagnosisMember = True
End Function

'################################################################################################################
'## 功能：  检查对象的ID是否有效
'##
'## 参数：  eEditType :当前的编辑模式
'##
'## 说明：  如果没有ID，则到数据库中提取一个唯一ID号。
'################################################################################################################
Public Sub CheckValidID(eEdtType As EditTypeEnum)
    Select Case eEdtType
    Case cprET_病历文件定义
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("病历文件结构")
    Case cprET_全文示范编辑
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("病历范文内容")
    Case cprET_单病历编辑, cprET_单病历审核
        If Me.ID = 0 Then Me.ID = GetEPRContentNextId
    End Select
End Sub

'################################################################################################################
'## 功能：  插入指定对象到Editor中
'##
'## 参数：  edtThis         :当前的编辑器控件
'##         lngStartPos     :当前位置
'##         blnForceInsert  :是否强制插入，在修改要素时如果前后都保护，这时需要强制插入
'################################################################################################################
Public Function InsertIntoEditor(ByRef edtThis As Object, _
    Optional ByVal lngStartPos As Long = -1, _
    Optional ByVal blnForceInsert As Boolean = False) As Boolean
    
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim blnForce As Boolean, lngLen As Long, strTmp As String, strText As String
    
    If lngStartPos = -1 Then lngStartPos = edtThis.Selection.StartPos
    bInKeys = IsBetweenAnyKeys(edtThis, lngStartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys Then InsertIntoEditor = False: Exit Function    '保证不能插入关键字内部
    If blnForceInsert = False And edtThis.Range(lngStartPos, lngStartPos).Font.ForeColor = PROTECT_FORECOLOR _
        And edtThis.Range(lngStartPos, lngStartPos).Font.Protected Then Exit Function
    If Trim(Me.描述) = "" Then Exit Function  '如果签名内容为空，也退出
    
    With edtThis
        .Freeze
        blnForce = .ForceEdit
        .ForceEdit = True
        
        strTmp = Format(Me.Key, "00000000") & ",0,0)"

        .Range(lngStartPos, lngStartPos).Font.Protected = False
        .Range(lngStartPos, lngStartPos).Font.Hidden = False
        
        strText = Me.描述
        lngLen = Len(strText)
        .Range(lngStartPos, lngStartPos) = "DS(" & strTmp & strText & "DE(" & strTmp
            
        .Range(lngStartPos, lngStartPos + 32 + lngLen).Font.Protected = True
        .Range(lngStartPos, lngStartPos + 16).Font.Hidden = True
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Hidden = False
        '删除线
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Strikethrough = (Me.终止版 > 0)
        '下划线
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Underline = cprNone
        '前景色
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.ForeColor = GetCharColor(Me.开始版, Me.终止版 + 1)
        '背景色
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.BackColor = ELE_BACKCOLOR
        .Range(lngStartPos + 16 + lngLen, lngStartPos + 32 + lngLen).Font.Hidden = True
        lngStartPos = lngStartPos + 32 + lngLen
        .Range(lngStartPos, lngStartPos).Selected

        .ForceEdit = blnForce
        .UnFreeze
    End With
    InsertIntoEditor = True
End Function

'################################################################################################################
'## 功能：  刷新对象的显示
'##
'## 参数：  edtThis         :当前的编辑器控件
'################################################################################################################
Public Function Refresh(ByRef edtThis As Object) As Boolean
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "D", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Freeze
            blnForce = .ForceEdit
            .ForceEdit = True
            Dim ParaFmt As New cParaFormat
            Set ParaFmt = edtThis.Range(lKSE, lKES).Para.GetParaFmt

            .Range(lKSS, lKEE) = ""
            Me.InsertIntoEditor edtThis, lKSS, True
            .Range(lKSE, lKSE + Len(Me.描述)).Selected
            .Range(lKSE, lKES).Para.SetParaFmt ParaFmt
            .ForceEdit = blnForce
            .UnFreeze
        End With
        Refresh = True
    Else
        Refresh = False
    End If
End Function

'################################################################################################################
'## 功能：  从编辑器中删除对象
'##
'## 参数：  edtThis         :当前的编辑器控件
'################################################################################################################
Public Function DeleteFromEditor(ByRef edtThis As Object) As Boolean
    '保存诊治要素编辑结果
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "D", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Freeze
            blnForce = .ForceEdit
            .ForceEdit = True
            .Range(lKSS, lKEE) = ""
            .ForceEdit = blnForce
            .UnFreeze
        End With
        DeleteFromEditor = True
    Else
        DeleteFromEditor = False
    End If
End Function

'################################################################################################################
'## 功能：  获取保存到数据库的SQL语句
'##
'## 参数：  eEdtType    :当前编辑方式；
'## 返回值：返回 存储过程调用的SQL语句。
'################################################################################################################
Public Function GetSaveSQL(ByVal eEdtType As EditTypeEnum, ByRef ArraySQL() As String) As Boolean
    
    Dim lngLBound As Long, lngUBound As Long    '传入数组的最小最大下标
    Err = 0: On Error Resume Next
    lngLBound = LBound(ArraySQL): lngUBound = UBound(ArraySQL)
    If Err <> 0 Then lngLBound = 0: lngUBound = -1
    ReDim Preserve ArraySQL(lngLBound To lngUBound + 1)

    Dim strProperty As String
    strProperty = Me.类型 & ";" & Me.中医 & ";" & Me.疾病id & ";" & Me.诊断id & ";" & Me.证候id & ";" & Me.疑诊 & ";" & Format(Me.日期, "yyyy-mm-dd hh:mm:ss") & ";" & Me.入院病情 & ";" & Me.出院情况
    
    Call CheckValidID(eEdtType)         '检查ID有效性！
    Select Case eEdtType
    Case cprET_病历文件定义
        ArraySQL(lngUBound + 1) = "Zl_病历文件结构_Update(" & Me.ID & "," & Me.文件ID & "," & _
            IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & Me.对象序号 & ",7," & Me.Key & ",0,'" & strProperty & "',NULL,'" & Me.描述 & "')"
    Case cprET_全文示范编辑
        ArraySQL(lngUBound + 1) = "Zl_病历范文内容_Update(" & Me.ID & "," & Me.文件ID & "," & _
            IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & Me.对象序号 & ",7," & Me.Key & ",0,'" & strProperty & "',NULL,'" & Me.描述 & "')"
    Case cprET_单病历编辑, cprET_单病历审核
        ArraySQL(lngUBound + 1) = "ZL_电子病历内容_Update(" & Me.ID & "," & Me.文件ID & "," & Me.开始版 & "," & Me.终止版 & "," & _
            IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & Me.对象序号 & ",7," & Me.Key & ",0,'" & strProperty & "',NULL,'" & Me.描述 & "')"
    End Select
    GetSaveSQL = True
End Function

'################################################################################################################
'## 功能：  选中该对象
'##
'## 参数：  edtThis         :当前的编辑器控件
'################################################################################################################
Public Function Selected(ByRef edtThis As Object) As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    On Error Resume Next
    bFinded = FindKey(edtThis, "D", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Range(lKSE, lKES).Selected
        End With
        Selected = True
    Else
        Selected = False
    End If
End Function

