VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarKey As Long                 '对象标记
Private mvarID As Double                  'ID
Private mvar文件ID As Long              '文件ID
Private mvar父ID As Double               '父ID
Private mvar对象序号 As Double            '序号
Private mvar保留对象 As Boolean         '保留对象
Private mvar内容文本 As String          '结果值，内容文本
Private mvar是否换行 As Boolean         '是否换行
Private mvar诊治要素ID As Long          '诊治要素ID
Private mvar签名要素 As Boolean         '可用的签名要素
Private mvar替换域 As Integer           '是否是替换域
Private mvar要素名称 As String          '要素名称
Private mvar要素类型 As Integer         '要素类型       0:数值 1:文字 2:日期
Private mvar要素长度 As Long            '要素长度
Private mvar要素小数 As Long            '要素小数位数
Private mvar要素单位 As String          '要素单位
Private mvar要素表示 As Integer         '要素表示       0-文本 1-上下 2-单选 3-复选
Private mvar输入形态 As Integer         '输入形态       0-弹出;1-展开(表示单选,复选类诊治要素的输入形态是展开放置选择项目)
Private mvar要素值域 As String          '要素值域       数值项目可以输入的最小值、最大值；对下拉、单选项目的选择项目列表
Private mvar开始版 As Long              '开始版，专用于“电子病历记录”
Private mvar终止版 As Long              '终止版，专用于“电子病历记录”
Private mvar自动转文本 As Boolean       '对于替换要素而言，是否在完成替换后自动转换为文本
Private mvar必填   As Byte            '是否必填要素　1-必填 0-非必填
Private mvar动态域 As Byte              '是否允许自定义 1-允许 0-不允许
Private mvar提示 As String              '在要素编辑窗口中提示

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

Public Property Let 保留对象(ByVal vData As Boolean)
    mvar保留对象 = vData
End Property

Public Property Get 保留对象() As Boolean
    保留对象 = mvar保留对象
End Property

Public Property Let 签名要素(ByVal vData As Boolean)
    mvar签名要素 = vData
End Property

Public Property Get 签名要素() As Boolean
    签名要素 = mvar签名要素
End Property

Public Property Let 内容文本(ByVal vData As String)
    '文本有效性检测！
    mvar内容文本 = vData
End Property

Public Property Get 内容文本() As String
    内容文本 = mvar内容文本
End Property

Public Property Let 是否换行(ByVal vData As Boolean)
    mvar是否换行 = vData
End Property

Public Property Get 是否换行() As Boolean
    是否换行 = mvar是否换行
End Property

Public Property Let 诊治要素ID(ByVal vData As Long)
    mvar诊治要素ID = vData
End Property

Public Property Get 诊治要素ID() As Long
    诊治要素ID = mvar诊治要素ID
End Property

Public Property Let 替换域(ByVal vData As Integer)
    mvar替换域 = vData
End Property

Public Property Get 替换域() As Integer
    替换域 = mvar替换域
End Property

Public Property Let 要素名称(ByVal vData As String)
    mvar要素名称 = vData
End Property

Public Property Get 要素名称() As String
    要素名称 = mvar要素名称
End Property

Public Property Let 要素类型(ByVal vData As Integer)
    mvar要素类型 = vData
End Property

Public Property Get 要素类型() As Integer
    要素类型 = mvar要素类型
End Property

Public Property Let 要素长度(ByVal vData As Long)
    mvar要素长度 = vData
End Property

Public Property Get 要素长度() As Long
    要素长度 = mvar要素长度
End Property

Public Property Let 要素小数(ByVal vData As Long)
    mvar要素小数 = vData
End Property

Public Property Get 要素小数() As Long
    要素小数 = mvar要素小数
End Property

Public Property Let 要素单位(ByVal vData As String)
    mvar要素单位 = vData
End Property

Public Property Get 要素单位() As String
    要素单位 = mvar要素单位
End Property

Public Property Let 要素表示(ByVal vData As Integer)
    mvar要素表示 = vData
End Property

Public Property Get 要素表示() As Integer
    要素表示 = mvar要素表示
End Property

Public Property Let 输入形态(ByVal vData As Integer)
    mvar输入形态 = vData
End Property

Public Property Get 输入形态() As Integer
    输入形态 = mvar输入形态
End Property

Public Property Let 要素值域(ByVal vData As String)
    mvar要素值域 = vData
End Property

Public Property Get 要素值域() As String
    要素值域 = mvar要素值域
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
    终止版 = IIf(mvar终止版 = 0, 0, mvar终止版)
End Property

Public Property Let 自动转文本(ByVal vData As Boolean)
    mvar自动转文本 = vData
End Property

Public Property Get 自动转文本() As Boolean
    自动转文本 = mvar自动转文本
End Property
Public Property Let 必填(ByVal vData As Byte)
    mvar必填 = vData
End Property
Public Property Get 必填() As Byte
    必填 = mvar必填
End Property
Public Property Let 动态域(ByVal vData As Byte)
    mvar动态域 = vData
End Property
Public Property Get 动态域() As Byte
    动态域 = mvar动态域
End Property
Public Property Let 提示(ByVal vDate As String)
    mvar提示 = vDate
End Property
Public Property Get 提示() As String
    提示 = mvar提示
End Property
'################################################################################################################
'## 功能：  是否自动转文本
'################################################################################################################
Public Property Get 对象属性() As String
    对象属性 = IIf(mvar自动转文本, 1, 0) & "|" & mvar必填 & "|" & mvar动态域 & "|" & mvar提示
End Property

'################################################################################################################
'## 功能：  是否自动转文本
'################################################################################################################
Public Property Let 对象属性(ByRef vData As String)
    Select Case UBound(Split(vData, "|"))
        Case Is <= 0
            mvar自动转文本 = IIf(Val(vData) = 0, False, True)
            mvar必填 = 0
            mvar动态域 = 0
            mvar提示 = ""
        Case 1
            mvar自动转文本 = IIf(Val(Split(vData, "|")(0)) = 0, False, True)
            mvar必填 = Split(vData, "|")(1)
            mvar动态域 = 0
            mvar提示 = ""
        Case 2
            mvar自动转文本 = IIf(Val(Split(vData, "|")(0)) = 0, False, True)
            mvar必填 = Split(vData, "|")(1)
            mvar动态域 = Split(vData, "|")(2)
            mvar提示 = ""
        Case 3
            mvar自动转文本 = IIf(Val(Split(vData, "|")(0)) = 0, False, True)
            mvar必填 = Split(vData, "|")(1)
            mvar动态域 = Split(vData, "|")(2)
            mvar提示 = Split(vData, "|")(3)
    End Select
End Property

'################################################################################################################
'## 功能：  复制诊治要素对象
'##
'## 参数：  blnWithKeys :是否连同Key值一起复制
'################################################################################################################
Public Function Clone(Optional blnWithKey As Boolean = False) As cEPRElement
    '用于拷贝数据，产生一个新类对象
    Dim cNode As New cEPRElement
    If blnWithKey Then cNode.Key = Me.Key
    cNode.ID = Me.ID
    cNode.文件ID = Me.文件ID
    cNode.开始版 = Me.开始版
    cNode.终止版 = Me.终止版
    cNode.父ID = Me.父ID
    cNode.对象序号 = Me.对象序号
    cNode.保留对象 = Me.保留对象
    cNode.内容文本 = Me.内容文本
    cNode.是否换行 = Me.是否换行
    cNode.诊治要素ID = Me.诊治要素ID
    cNode.替换域 = Me.替换域
    cNode.要素名称 = Me.要素名称
    cNode.要素类型 = Me.要素类型
    cNode.要素长度 = Me.要素长度
    cNode.要素小数 = Me.要素小数
    cNode.要素单位 = Me.要素单位
    cNode.要素表示 = Me.要素表示
    cNode.输入形态 = Me.输入形态
    cNode.要素值域 = Me.要素值域
    cNode.对象属性 = Me.对象属性
    cNode.必填 = Me.必填
    cNode.动态域 = Me.动态域
    cNode.提示 = Me.提示
    Set Clone = cNode
End Function
Public Function FillElementMember(ByVal rs As ADODB.Recordset, ByVal strTableName As String) As Boolean
'功能:用数据记录填充类成员
    If rs Is Nothing Then Exit Function
    If rs.EOF Then Exit Function
    
    If NVL(rs!对象标记, 0) <> 0 Then Me.Key = NVL(rs!对象标记, 0)
    Me.ID = rs!ID
    Me.文件ID = NVL(rs!文件ID, 0)
    Me.父ID = NVL(rs!父ID, 0)
    Me.对象序号 = NVL(rs!对象序号, 0)
    Me.保留对象 = IIf(NVL(rs!保留对象, 0) = 0, False, True)
    Me.内容文本 = NVL(rs!内容文本)
    Me.是否换行 = IIf(NVL(rs!是否换行, 0) = 0, False, True)
    Me.诊治要素ID = NVL(rs!诊治要素ID, 0)
    Me.替换域 = NVL(rs!替换域, 0)
    Me.要素名称 = NVL(rs!要素名称)
    Me.要素类型 = NVL(rs!要素类型, 0)
    Me.要素长度 = NVL(rs!要素长度, 0)
    Me.要素小数 = NVL(rs!要素小数, 0)
    Me.要素单位 = NVL(rs!要素单位)
    Me.要素表示 = NVL(rs!要素表示, 0)
    Me.输入形态 = NVL(rs!输入形态, 0)
    Me.要素值域 = NVL(rs!要素值域)
    Me.对象属性 = NVL(rs!对象属性)
        
    If InStr(strTableName, "电子病历内容") > 0 Then '病历文件结构和病历范文内容没有这两个字段,但都共用这个函数
        Me.开始版 = NVL(rs!开始版, 1)
        Me.终止版 = NVL(rs!终止版, 0)
    End If
    
    FillElementMember = True
End Function
'################################################################################################################
'## 功能：  从数据库直接获取诊治要素内容
'##
'## 参数：  eEdtType    :当前编辑方式；
'##         lngID       :数据库中的对象ID号
'##         blnWithKey  :是否包含Key值一起恢复
'##         strTableName:表名，如果省略，则根据当前的eEdtType来判断表名
'################################################################################################################
Public Function GetElementFromDB(ByVal eEdtType As EditTypeEnum, ByVal lngID As Double, _
    Optional blnWithKey As Boolean = False, _
    Optional ByVal strTableName As String = "") As Boolean
    
    Dim rs As New ADODB.Recordset
    If strTableName = "" Then
        Select Case eEdtType
        Case cprET_病历文件定义
            strTableName = "病历文件结构"
        Case cprET_全文示范编辑
            strTableName = "病历范文内容"
        Case cprET_单病历编辑, cprET_单病历审核
            strTableName = "电子病历内容"
        End Select
    End If
    
    gstrSQL = "Select ID, 文件id," & IIf(InStr(strTableName, "电子病历内容") > 0, "开始版, 终止版,", "") & " 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 诊治要素id," & vbNewLine & _
                "       替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域 " & vbNewLine & _
                "From " & strTableName & vbNewLine & _
                "Where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRElement", lngID)
    If Not rs.EOF Then
        If blnWithKey And NVL(rs("对象标记"), 0) <> 0 Then Me.Key = NVL(rs("对象标记"), 0)
        Me.ID = rs("ID")
        Me.文件ID = NVL(rs("文件ID"), 0)
        Me.父ID = NVL(rs("父ID"), 0)
        Me.对象序号 = NVL(rs("对象序号"), 0)
        Me.保留对象 = IIf(NVL(rs("保留对象"), 0) = 0, False, True)
        Me.内容文本 = NVL(rs("内容文本"))
        Me.是否换行 = IIf(NVL(rs("是否换行"), 0) = 0, False, True)
        Me.诊治要素ID = NVL(rs("诊治要素ID"), 0)
        Me.替换域 = NVL(rs("替换域"), 0)
        Me.要素名称 = NVL(rs("要素名称"))
        Me.要素类型 = NVL(rs("要素类型"), 0)
        Me.要素长度 = NVL(rs("要素长度"), 0)
        Me.要素小数 = NVL(rs("要素小数"), 0)
        Me.要素单位 = NVL(rs("要素单位"))
        Me.要素表示 = NVL(rs("要素表示"), 0)
        Me.输入形态 = NVL(rs("输入形态"), 0)
        Me.要素值域 = NVL(rs("要素值域"))
        Me.对象属性 = NVL(rs("对象属性"))
        
        If eEdtType = cprET_单病历审核 Or eEdtType = cprET_单病历编辑 Then
            Me.开始版 = NVL(rs("开始版"), 1)
            Me.终止版 = NVL(rs("终止版"), 0)
        End If
        GetElementFromDB = True
    Else
        GetElementFromDB = False
    End If
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
'## 功能：  检查诊治要素的原始定义是否存在（用于XML导入时的验证）
'################################################################################################################
Public Function CheckValid() As Boolean
    Dim rs As New Recordset
    gstrSQL = "Select ID From 诊治所见项目 Where ID = [1] And 中文名 = [2]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRElement", Me.ID, Me.要素名称)
    If rs.EOF Then
        CheckValid = False
    Else
        CheckValid = (rs!ID > 0)
    End If
End Function

'################################################################################################################
'## 功能：  获取可供显示的文本长度
'##
'## 说明：  注意不是内容文本的长度，比如内容文本为空时要显示[要素名称]
'################################################################################################################
Public Function GetValidTextLength() As Long
    Dim lngLen As Long, lngLen2 As Long
    If Me.替换域 = 0 Then
        lngLen = Len(IIf(Me.内容文本 = "", "[" & Me.要素名称 & "]", Me.内容文本))
    Else
        lngLen = Len(IIf(Me.内容文本 = "", "{" & Me.要素名称 & "}", Me.内容文本))
    End If
    lngLen2 = Len(Me.要素单位)
    GetValidTextLength = lngLen + lngLen2
End Function


Public Function InsertIntoEditor(ByRef edtThis As Object, Optional ByVal lngStartPos As Long = -1, Optional ByVal blnFirstIns As Boolean = False, Optional ByVal blnFromPhrase As Boolean = False, Optional ByVal blnForceInsert As Boolean = False) As Boolean
    '******************************************************************************************************************
    ' 功能：  插入诊治要素到Editor中
    '
    ' 参数：  edtThis         :当前的编辑器控件
    '         lngStartPos     :当前位置
    '         blnFirstIns     :是否是第一次插入（默认第一次插入会在末尾加逗号）
    '         blnFromPhrase   :如果是从词句示范中调入的诊治要素为True，表示还未进行编辑（有下划线）。
    '         blnForceInsert  :是否强制插入，在修改要素时如果前后都保护，这时需要强制插入
    '******************************************************************************************************************
    
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    
    If lngStartPos = -1 Then lngStartPos = edtThis.Selection.StartPos
    bInKeys = IsBetweenAnyKeys(edtThis, lngStartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys Then InsertIntoEditor = False: Exit Function    '保证不能插入关键字内部
    If blnForceInsert = False And edtThis.Range(lngStartPos, lngStartPos).Font.ForeColor = PROTECT_FORECOLOR _
        And edtThis.Range(lngStartPos, lngStartPos).Font.Protected Then Exit Function

    Dim p As Long, lngKey As Long, blnForce As Boolean, lngLen As Long, lngLen2 As Long, strTmp As String
    p = lngStartPos
    blnForce = edtThis.ForceEdit

    If blnFirstIns Then Me.是否换行 = False
    With edtThis
        .Freeze
        .ForceEdit = True
        lngKey = Me.Key
        
        strTmp = Format(lngKey, "00000000") & "," & IIf(Me.保留对象, 1, 0) & ",0)"

        .Range(lngStartPos, lngStartPos).Font.Protected = False
        .Range(lngStartPos, lngStartPos).Font.Hidden = False
        
        If Me.输入形态 = 0 Then
            If Me.替换域 = 0 Then
                lngLen = Len(IIf(Me.内容文本 = "", "[" & Me.要素名称 & "]", Me.内容文本))
            Else
                lngLen = Len(IIf(Me.内容文本 = "", "{" & Me.要素名称 & "}", Me.内容文本))
            End If
            lngLen2 = Len(Me.要素单位)
            If Me.替换域 = 0 Then
                .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & _
                    IIf(Me.内容文本 = "", "[" & Me.要素名称 & "]", Me.内容文本) & _
                    Me.要素单位 & "EE(" & strTmp & IIf(Me.是否换行, vbCrLf, "")
            ElseIf Me.替换域 = 2 Then
                '字典项目
                .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & _
                    IIf(Me.内容文本 = "", "<" & Me.要素名称 & ">", Me.内容文本) & _
                    Me.要素单位 & "EE(" & strTmp & IIf(Me.是否换行, vbCrLf, "")
            Else
                .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & _
                    IIf(Me.内容文本 = "", "{" & Me.要素名称 & "}", Me.内容文本) & _
                    Me.要素单位 & "EE(" & strTmp & IIf(Me.是否换行, vbCrLf, "")
            End If
        Else
            If blnFirstIns Then
                Dim strContent As String, i As Long, T As Variant
                T = Split(Me.要素值域, ";")
                For i = 0 To UBound(T)
                    strContent = strContent & IIf(Me.要素表示 = 3, "□", "○") & T(i) & IIf(i = UBound(T), "", "  ")   '○●□■
                Next
                lngLen = Len(strContent)
                lngLen2 = Len(Me.要素单位)
                Me.内容文本 = strContent
                .Range(lngStartPos, lngStartPos) = "ES(" & strTmp & _
                    strContent & _
                    Me.要素单位 & "EE(" & strTmp & IIf(Me.是否换行, vbCrLf, "")
            Else
                lngLen = Len(Me.内容文本)
                lngLen2 = Len(Me.要素单位)
                .Range(lngStartPos, lngStartPos) = "ES(" & strTmp & _
                    Me.内容文本 & _
                    Me.要素单位 & "EE(" & strTmp & IIf(Me.是否换行, vbCrLf, "")
            End If
        End If
        .Range(lngStartPos, lngStartPos + 32 + lngLen + lngLen2).Font.Protected = True
        .Range(lngStartPos, lngStartPos + 16).Font.Hidden = True
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen + lngLen2).Font.Hidden = False
        '下划线
        If Me.输入形态 = 0 Then
            .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Underline = IIf(Me.内容文本 = "" Or blnFromPhrase, ELE_UNDERLINE, cprNone)
        Else
            If Me.要素表示 = 2 And InStr(Me.内容文本, "●") = 0 Then
                .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Underline = cprwave
            ElseIf Me.要素表示 = 3 And InStr(Me.内容文本, "■") = 0 Then
                .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Underline = cprwave
            Else
                .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Underline = IIf(Me.内容文本 = "" Or blnFromPhrase, ELE_UNDERLINE, cprNone)
            End If
        End If
        '删除线的设置
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen + lngLen2).Font.Strikethrough = (Me.终止版 > 0)
        '前景色
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen + lngLen2).Font.ForeColor = GetCharColor(Me.开始版, Me.终止版 + 1)
        '背景色
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen + lngLen2).Font.BackColor = ELE_BACKCOLOR
        .Range(lngStartPos + 16 + lngLen + lngLen2, lngStartPos + 32 + lngLen + lngLen2).Font.Hidden = True
        lngStartPos = lngStartPos + 32 + lngLen + lngLen2

        If Me.是否换行 Then
            .Range(lngStartPos, lngStartPos + 2).Font.Protected = False
            .Range(lngStartPos, lngStartPos + 2).Font.Hidden = False
            lngStartPos = lngStartPos + 2
        End If
        
        Me.签名要素 = False
        
        If Me.替换域 = 1 Then
            Select Case Me.要素名称
            Case "经治医师签名", "主治医师签名", "主任医师签名"
                If Me.内容文本 <> "" Then
                    Dim lESS As Long
                    Dim lESE As Long
                    Dim lEES As Long
                    Dim lEEE As Long
                    
                    Dim blnNeeded As Boolean
                    Dim blnFinded As Boolean
        
                    blnFinded = FindKey(edtThis, "E", Me.Key, lESS, lESE, lEES, lEEE, blnNeeded)
                    
                    If blnFinded Then
                        .Range(lESS, lEEE).Font.Protected = True
                        .Range(lESS, lEEE).Font.Hidden = True
                    End If
                Else
                    '有可用的签名要素
                    Me.签名要素 = True
                End If
            End Select
        End If
        
        .Range(lngStartPos, lngStartPos).Selected

        .ForceEdit = blnForce
        .UnFreeze
    End With
    InsertIntoEditor = True
End Function

'################################################################################################################
'## 功能：  刷新显示
'##
'## 参数：  edtThis         :当前的编辑器控件
'################################################################################################################
Public Function Refresh(ByRef edtThis As Object) As Boolean
    '保存诊治要素编辑结果
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "E", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Freeze
            blnForce = .ForceEdit
            .ForceEdit = True
            Dim ParaFmt As New cParaFormat, FontFmt As New cFontFormat
            Set ParaFmt = edtThis.Range(lKSE, lKES).Para.GetParaFmt
            Set FontFmt = edtThis.Range(lKSE, lKES).Font.GetFontFmt
            CheckFonts FontFmt
            If FontFmt.Underline = cprwave Then FontFmt.Underline = cprNone
            
            .Range(lKSS, lKEE) = ""
            Me.是否换行 = False
            Me.InsertIntoEditor edtThis, lKSS, , , True
            If InStr(",经治医师签名,主治医师签名,主任医师签名", Me.要素名称) = 0 Then '非签名要素的编辑刷新
                .Range(lKSE, lKSE + Len(Me.内容文本) + Len(Me.要素单位)).Selected
                .Range(lKSE, lKSE + Len(Me.内容文本) + Len(Me.要素单位)).Font.SetFontFmt FontFmt
                .Range(lKSE, lKSE + Len(Me.内容文本) + Len(Me.要素单位)).Para.SetParaFmt ParaFmt
            End If
            .ForceEdit = blnForce
            .UnFreeze
        End With
        Refresh = True
    Else
        Refresh = False
    End If
End Function

'################################################################################################################
'## 功能：  从编辑器中删除
'##
'## 参数：  edtThis         :当前的编辑器控件
'################################################################################################################
Public Function DeleteFromEditor(ByRef edtThis As Object) As Boolean
    '保存诊治要素编辑结果
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "E", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
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
'## 功能：  获取提纲保存到数据库的SQL语句
'##
'## 参数：  eEdtType    :当前编辑方式；
'## 返回值：返回 存储过程调用的SQL语句。
'################################################################################################################
Public Function GetSaveSQL(ByVal eEdtType As EditTypeEnum, ByRef ArraySQL() As String) As Boolean
    Dim lngCount As Long
    lngCount = UBound(ArraySQL) + 1
    ReDim Preserve ArraySQL(1 To lngCount) As String

    Call CheckValidID(eEdtType)         '检查ID有效性！
    Select Case eEdtType
    Case cprET_病历文件定义
        gstrSQL = "Zl_病历文件结构_Update(" & Me.ID & "," & Me.文件ID & "," & IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & _
            Me.对象序号 & ",4," & Me.Key & "," & IIf(Me.保留对象, 1, 0) & ",'" & Me.对象属性 & "',NULL,'" & Replace(Me.内容文本, "'", "' || chr(39) || '") & "'," & IIf(Me.是否换行, 1, 0) & _
            ",NULL,NULL,NULL," & _
            IIf(Me.诊治要素ID = 0, "NULL", Me.诊治要素ID) & "," & Me.替换域 & ",'" & Me.要素名称 & "'," & Me.要素类型 & "," & Me.要素长度 & "," & _
            Me.要素小数 & ",'" & Me.要素单位 & "'," & Me.要素表示 & "," & Me.输入形态 & ",'" & Me.要素值域 & "')"
    Case cprET_全文示范编辑
        gstrSQL = "Zl_病历范文内容_Update(" & Me.ID & "," & Me.文件ID & "," & IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & _
            Me.对象序号 & ",4," & Me.Key & "," & IIf(Me.保留对象, 1, 0) & ",'" & Me.对象属性 & "',NULL,'" & Replace(Me.内容文本, "'", "' || chr(39) || '") & "'," & IIf(Me.是否换行, 1, 0) & _
            ",NULL,NULL,NULL,NULL," & _
            IIf(Me.诊治要素ID = 0, "NULL", Me.诊治要素ID) & "," & Me.替换域 & ",'" & Me.要素名称 & "'," & Me.要素类型 & "," & Me.要素长度 & "," & _
            Me.要素小数 & ",'" & Me.要素单位 & "'," & Me.要素表示 & "," & Me.输入形态 & ",'" & Me.要素值域 & "')"
    Case cprET_单病历编辑, cprET_单病历审核
        '加了 开始版，终止版 字段，其他相同
        gstrSQL = "ZL_电子病历内容_Update(" & Me.ID & "," & Me.文件ID & "," & Me.开始版 & "," & Me.终止版 & "," & IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & _
            Me.对象序号 & ",4," & Me.Key & "," & IIf(Me.保留对象, 1, 0) & ",'" & Me.对象属性 & "',NULL,'" & Replace(Me.内容文本, "'", "' || chr(39) || '") & "'," & IIf(Me.是否换行, 1, 0) & _
            ",NULL,NULL,NULL,NULL," & _
            IIf(Me.诊治要素ID = 0, "NULL", Me.诊治要素ID) & "," & Me.替换域 & ",'" & Me.要素名称 & "'," & Me.要素类型 & "," & Me.要素长度 & "," & _
            Me.要素小数 & ",'" & Me.要素单位 & "'," & Me.要素表示 & "," & Me.输入形态 & ",'" & Me.要素值域 & "')"
    End Select
    ArraySQL(lngCount) = gstrSQL
    
    GetSaveSQL = True
End Function

'################################################################################################################
'## 功能：  选中该诊治要素
'##
'## 参数：  edtThis         :当前的编辑器控件
'################################################################################################################
Public Function Selected(ByRef edtThis As Object) As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    On Error Resume Next
    bFinded = FindKey(edtThis, "E", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Range(lKSE, lKES).Selected
        End With
        Selected = True
    Else
        Selected = False
    End If
End Function

Public Sub CheckFonts(fmtFont As cFontFormat)
Dim i As Integer
    For i = 0 To Screen.FontCount - 1
        If Screen.Fonts(i) = fmtFont.Name Then
            Exit Sub
        End If
    Next i
    fmtFont.Name = "宋体"
End Sub
