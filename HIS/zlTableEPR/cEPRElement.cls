VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarKey As Long                 '对象标记
Private mvarID As Double                  '混合编辑区域中要素的ID,单要素的ID
Private mvar父ID As Double                '混合编辑区域中要素父ID
Private mvar区域 As String              '要素所在单元格 行|列
Private mvar内容文本 As String          '内容文本
Private mvar内容行次 As Long            '内容行次
Private mvar诊治要素ID As Long          '诊治要素ID
Private mvar替换域 As Integer           '是否是替换域
Private mvar要素名称 As String          '要素名称
Private mvar要素类型 As Integer         '要素类型       0:数值 1:文字 2:日期 3-逻辑
Private mvar要素长度 As Long            '要素长度
Private mvar要素小数 As Long            '要素小数位数
Private mvar要素单位 As String          '要素单位
Private mvar要素表示 As Integer         '要素表示       0-文本 1-上下 2-单选 3-复选
Private mvar输入形态 As Integer         '输入形态       0-弹出;1-展开(表示单选,复选类诊治要素的输入形态是展开放置选择项目)
Private mvar要素值域 As String          '要素值域       数值项目可以输入的最小值、最大值；对下拉、单选项目的选择项目列表
Private mvar保留对象 As Boolean         '
Private mvar自动转文本 As Boolean       '对于替换要素而言，是否在完成替换后自动转换为文本
Private mvar必填   As Integer            '是否必填要素　1-必填 0-非必填
Private mvar动态域 As Byte              '是否允许自定义 1-允许 0-不允许
Private mvar开始版 As Long
Private mvar终止版 As Long
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
Public Property Let 父ID(ByVal vData As Double)
    mvar父ID = vData
End Property

Public Property Get 父ID() As Double
    父ID = mvar父ID
End Property
Public Property Let 区域(ByVal vData As String)
    mvar区域 = vData
End Property
Public Property Get 区域() As String
    区域 = mvar区域
End Property

Public Property Let 内容文本(ByVal vData As String)
    mvar内容文本 = vData
End Property

Public Property Get 内容文本() As String
    内容文本 = mvar内容文本
End Property
Public Property Let 内容行次(ByVal vData As Long)
    mvar内容行次 = vData
End Property

Public Property Get 内容行次() As Long
    内容行次 = mvar内容行次
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
Public Property Let 保留对象(ByVal vData As Boolean)
    mvar保留对象 = vData
End Property

Public Property Get 保留对象() As Boolean
    保留对象 = mvar保留对象
End Property
Public Property Let 自动转文本(ByVal vData As Boolean)
    mvar自动转文本 = vData
End Property

Public Property Get 自动转文本() As Boolean
    自动转文本 = mvar自动转文本
End Property
Public Property Let 必填(ByVal vData As Integer)
    mvar必填 = vData
End Property
Public Property Get 必填() As Integer
    必填 = mvar必填
End Property
Public Property Let 动态域(ByVal vData As Byte)
    mvar动态域 = vData
End Property
Public Property Get 动态域() As Byte
    动态域 = mvar动态域
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
Public Property Get 对象属性() As String
    对象属性 = IIf(mvar自动转文本, 1, 0) & "|" & mvar必填 & "|" & mvar动态域
End Property
Public Property Let 对象属性(ByRef vData As String)
Dim strFirst As String, strData As String
    strFirst = Split(vData, "|")(0)
    If InStr(strFirst, ",") > 0 Then
        strData = Mid(vData, InStr(vData, "|") + 1)
    Else
        strData = vData
    End If
        
    Select Case UBound(Split(strData, "|"))
        Case Is <= 0
            mvar自动转文本 = IIf(Val(strData) = 0, False, True)
            mvar必填 = 0
            mvar动态域 = 0
        Case 1
            mvar自动转文本 = IIf(Val(Split(strData, "|")(0)) = 0, False, True)
            mvar必填 = Val(Split(strData, "|")(1))
            mvar动态域 = 0
        Case 2
            mvar自动转文本 = IIf(Val(Split(strData, "|")(0)) = 0, False, True)
            mvar必填 = Val(Split(strData, "|")(1))
            mvar动态域 = Val(Split(strData, "|")(2))
    End Select
End Property

'################################################################################################################
'## 功能：  复制诊治要素对象
'##
'## 参数：  blnWithKeys :是否连同Key值一起复制
'################################################################################################################
Public Sub Clone(cNode As cTabElement)
'复制
    With cNode
        .内容文本 = Me.内容文本
        .内容行次 = Me.内容行次
        .诊治要素ID = Me.诊治要素ID
        .替换域 = Me.替换域
        .要素名称 = Me.要素名称
        .要素类型 = Me.要素类型
        .要素长度 = Me.要素长度
        .要素小数 = Me.要素小数
        .要素单位 = Me.要素单位
        .要素表示 = Me.要素表示
        .输入形态 = Me.输入形态
        .要素值域 = Me.要素值域
        .保留对象 = Me.保留对象
        .对象属性 = Me.对象属性
        .必填 = Me.必填
        .动态域 = Me.动态域
        .开始版 = Me.开始版
        .终止版 = Me.终止版
    End With
End Sub
'################################################################################################################
'## 功能：  检查诊治要素的原始定义是否存在（用于XML导入时的验证）
'################################################################################################################
Public Function CheckValid() As Boolean
    Dim rs As New Recordset
    gstrSQL = "Select Count(*) From 诊治所见项目 Where ID = [1] And 中文名 = [2]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cTabElement", Me.诊治要素ID, Me.要素名称)
    If rs.EOF Then
        CheckValid = False
    Else
        CheckValid = (rs(0) > 0)
    End If
End Function
Public Function GetSaveSQL(ByVal editType As TabET, ByVal CellTmp As cTabCell, arrSQL As Variant) As Boolean
'功能:生成单要素和混合区域中要素保存SQL,CellTmp中
    On Error GoTo errHand
    With CellTmp
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        Select Case editType
        Case TabET_病历文件定义
            arrSQL(UBound(arrSQL)) = "Zl_病历文件结构_Update(" & IIf(mvar父ID = 0, .ID, mvarID) & "," & .文件ID & "," & _
                                        IIf(.对象类型 = cprCTElement, "Null", .ID) & "," & .对象序号 & "," & cprCTElement & ",Null," & _
                                        IIf(.保留对象, 1, 0) & ",'" & .Row & "," & .Col & "|" & Me.对象属性 & "'," & mvar内容行次 & ",'" & _
                                        IIf(.对象类型 = cprCTElement, .内容文本, mvar内容文本) & "',0,Null,Null,Null," & _
                                        IIf(mvar诊治要素ID = 0, "Null", mvar诊治要素ID) & "," & mvar替换域 & ",'" & mvar要素名称 & "'," & mvar要素类型 & "," & _
                                        mvar要素长度 & "," & mvar要素小数 & ",'" & mvar要素单位 & "'," & mvar要素表示 & "," & mvar输入形态 & ",'" & mvar要素值域 & "')"
        Case TabET_全文示范编辑
            arrSQL(UBound(arrSQL)) = "zl_病历范文内容_update(" & IIf(mvar父ID = 0, .ID, mvarID) & "," & .文件ID & "," & _
                                        IIf(.对象类型 = cprCTElement, "Null", .ID) & "," & .对象序号 & "," & cprCTElement & ",Null," & _
                                        IIf(.保留对象, 1, 0) & ",'" & .Row & "," & .Col & "|" & Me.对象属性 & "'," & mvar内容行次 & ",'" & _
                                        IIf(.对象类型 = cprCTElement, .内容文本, mvar内容文本) & "',0,Null,Null,Null,Null," & _
                                        IIf(mvar诊治要素ID = 0, "Null", mvar诊治要素ID) & "," & mvar替换域 & ",'" & mvar要素名称 & "'," & mvar要素类型 & "," & _
                                        mvar要素长度 & "," & mvar要素小数 & ",'" & mvar要素单位 & "'," & mvar要素表示 & "," & mvar输入形态 & ",'" & mvar要素值域 & "')"
        Case TabET_单病历编辑, TabET_单病历审核
            arrSQL(UBound(arrSQL)) = "Zl_电子病历内容_Update(" & IIf(mvar父ID = 0, .ID, mvarID) & "," & .文件ID & "," & mvar开始版 & "," & mvar终止版 & "," & _
                                        IIf(.对象类型 = cprCTElement, "Null", .ID) & "," & .对象序号 & "," & cprCTElement & ",Null," & _
                                        IIf(.保留对象, 1, 0) & ",'" & .Row & "," & .Col & "|" & Me.对象属性 & "'," & mvar内容行次 & ",'" & _
                                        IIf(.对象类型 = cprCTElement, .内容文本, mvar内容文本) & "',0,Null,Null,Null,Null," & _
                                        IIf(mvar诊治要素ID = 0, "Null", mvar诊治要素ID) & "," & mvar替换域 & ",'" & mvar要素名称 & "'," & mvar要素类型 & "," & _
                                        mvar要素长度 & "," & mvar要素小数 & ",'" & mvar要素单位 & "'," & mvar要素表示 & "," & mvar输入形态 & ",'" & mvar要素值域 & "')"
        End Select
    End With
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function InsertIntoEditor(ByRef edtThis As Object, ByVal editType As Integer, Optional ByVal lngStartPos As Long = -1) As Boolean
'******************************************************************************************************************
' 功能：  插入诊治要素到Editor中
'
' 参数：  edtThis         :当前的编辑器控件
'         lngStartPos     :当前位置
'         blnFirstIns     :是否是第一次插入（默认第一次插入会在末尾加逗号）
'******************************************************************************************************************
Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    
    If lngStartPos = -1 Then lngStartPos = edtThis.Selection.StartPos
    bInKeys = IsBetweenAnyKeys(edtThis, lngStartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys Then InsertIntoEditor = False: Exit Function    '保证不能插入关键字内部
    
    If edtThis.Range(lngStartPos, lngStartPos).Font.ForeColor = PROTECT_FORECOLOR And edtThis.Range(lngStartPos, lngStartPos).Font.Protected Then Exit Function

    Dim p As Long, lngKey As Long, blnForce As Boolean, lngLen As Long, lngLen2 As Long, strTmp As String
    p = lngStartPos
    blnForce = edtThis.ForceEdit

    With edtThis
        .Freeze
        .ForceEdit = True
        lngKey = Me.Key
        
        strTmp = Format(lngKey, "00000000") & "," & IIf(Me.保留对象, 1, 0) & ",0)"

        .Range(lngStartPos, lngStartPos).Font.Protected = False
        .Range(lngStartPos, lngStartPos).Font.Hidden = False
        
        If Me.替换域 = 1 And (editType = TabET_单病历编辑 Or editType = TabET_单病历审核) Then '自动替换要素有"自动转文本"功能
            If Trim(Me.内容文本) = "" Then
                If Me.自动转文本 Then
                    lngLen = Len(" ")
                    lngLen2 = Len(Me.要素单位)
                    .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & " " & Me.要素单位 & "EE(" & strTmp
                Else
                    lngLen = Len("[" & Me.要素名称 & "]")
                    lngLen2 = Len(Me.要素单位)
                    .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & "[" & Me.要素名称 & "]" & Me.要素单位 & "EE(" & strTmp
                End If
            Else
                lngLen = Len(Me.内容文本)
                lngLen2 = Len(Me.要素单位)
                .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & Me.内容文本 & Me.要素单位 & "EE(" & strTmp
            End If
        Else
            If Me.输入形态 = 0 Then
                lngLen = Len(IIf(Me.内容文本 = "", "[" & Me.要素名称 & "]", Me.内容文本))
                lngLen2 = Len(Me.要素单位)
                .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & IIf(Me.内容文本 = "", "[" & Me.要素名称 & "]", Me.内容文本) & Me.要素单位 & "EE(" & strTmp
            Else
                lngLen = Len(Me.内容文本)
                lngLen2 = Len(Me.要素单位)
                .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & Me.内容文本 & Me.要素单位 & "EE(" & strTmp
            End If
        End If
        .Range(lngStartPos, lngStartPos + 32 + lngLen + lngLen2).Font.Protected = True '保护要素及隐藏区域
        .Range(lngStartPos, lngStartPos + 16).Font.Hidden = True                       '隐藏前导关键字
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen + lngLen2).Font.Hidden = False   '正文显示
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen + lngLen2).Font.BackColor = ELE_BACKCOLOR
        .Range(lngStartPos + 16 + lngLen + lngLen2, lngStartPos + 32 + lngLen + lngLen2).Font.Hidden = True '隐藏尾部关键字
        lngStartPos = lngStartPos + 32 + lngLen + lngLen2                                   '位置后移
        .Range(lngStartPos, lngStartPos).Selected                                           '光标定位
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
            Dim ParaFmt As New cTabParaFormat, FontFmt As New cTabFontFormat
            Set ParaFmt = edtThis.Range(lKSE, lKSE).Para.GetParaFmt
            Set FontFmt = edtThis.Range(lKSE, lKSE).Font.GetFontFmt
            If FontFmt.Underline = cprwave Then FontFmt.Underline = cprnone
            
            .Range(lKSS, lKEE).Text = ""
            Me.InsertIntoEditor edtThis, lKSS
            .Range(lKSE, lKSE + Len(Me.内容文本)).Selected
            .Range(lKSE, lKES).Font.SetFontFmt FontFmt
            .Range(lKSE, lKSE).Para.SetParaFmt ParaFmt
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
            .Range(lKSS, lKSS).Selected
            .Range(lKSS, lKEE).Text = ""
            .ForceEdit = blnForce
            .UnFreeze
        End With
        DeleteFromEditor = True
    Else
        DeleteFromEditor = False
    End If
End Function
'################################################################################################################
'## 功能：  检查所有对象的ID是否有效
'##
'## 参数：  eEditType :当前的编辑模式
'##
'## 说明：  如果没有ID，则到数据库中提取一个唯一ID号。父ID=0的情况是单元素由Cell取ID
'################################################################################################################
Public Sub CheckValidID(eEditType As TabET)
    Dim i As Long, strTable As String
    Select Case eEditType
    Case TabET_病历文件定义
        strTable = "病历文件结构"
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId(strTable)
    Case TabET_全文示范编辑
        strTable = "病历范文内容"
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId(strTable)
    Case TabET_单病历编辑, TabET_单病历审核
        strTable = "电子病历内容"
        If Me.ID = 0 Then Me.ID = GetEPRContentNextId
    End Select
    
End Sub