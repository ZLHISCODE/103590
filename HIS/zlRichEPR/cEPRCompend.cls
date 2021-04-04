VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRCompend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Const COM_BACKCOLOR = &HD5FEFF              '特殊提纲的背景颜色 '&HDCDCDC
Private Const COM_FORECOLOR = vbBlack               '特殊提纲的前景颜色
Private Const COM_UNDERLINE = cprwave               '特殊提纲的下划线

Private mvarKey As Long             '对应于集合的Key
Private mvar父Key As Long           '父Key
Private mvarID As Double              'ID
Private mvar文件ID As Long          '文件ID
Private mvar父ID As Double           '父ID
Private mvar对象序号 As Double        '对象序号
Private mvar保留对象 As Boolean     '保留对象与否
Private mvar名称 As String          '名称，内容文本
Private mvar说明 As String          '说明，对象属性
Private mvar定义提纲ID As Long      '定义提纲ID
Private mvar预制提纲ID As Long      '预制提纲ID
Private mvar复用提纲 As Boolean     '复用提纲与否   '对于示范，没有“复用提纲”的概念
Private mvar使用时机 As String      '提纲使用使用时机   '对于示范，没有“使用时机”的概念
Private mvar开始版 As Long          '开始版，专用于“电子病历记录”
Private mvar终止版 As Long          '终止版，专用于“电子病历记录”
Private mvarLevel As Long           'Level，Level 1，2，3，...

Private mvar内部序号 As Long        '内部序号，用于提纲快速排序，与对象序号不同（对象序号在文件中统一排序）

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let 父Key(ByVal vData As Long)
    mvar父Key = vData
End Property

Public Property Get 父Key() As Long
    父Key = mvar父Key
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

Public Property Let 名称(ByVal vData As String)
    mvar名称 = vData
End Property

Public Property Get 名称() As String
    名称 = Replace(mvar名称, "'", "' || chr(39) || '")
End Property

Public Property Let 说明(ByVal vData As String)
    mvar说明 = vData
End Property

Public Property Get 说明() As String
    说明 = Replace(mvar说明, "'", "' || chr(39) || '")
End Property

Public Property Let 定义提纲ID(ByVal vData As Long)
    mvar定义提纲ID = vData
End Property

Public Property Get 定义提纲ID() As Long
    定义提纲ID = mvar定义提纲ID
End Property

Public Property Let 预制提纲ID(ByVal vData As Long)
    mvar预制提纲ID = vData
End Property

Public Property Get 预制提纲ID() As Long
    预制提纲ID = mvar预制提纲ID
End Property

Public Property Let 复用提纲(ByVal vData As Boolean)
    mvar复用提纲 = vData
End Property

Public Property Get 复用提纲() As Boolean
    复用提纲 = mvar复用提纲
End Property

Public Property Let 使用时机(ByVal vData As String)
    mvar使用时机 = vData
End Property

Public Property Get 使用时机() As String
    使用时机 = mvar使用时机
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
'    终止版 = IIf(mvar终止版 = 0, 0, mvar终止版)
    终止版 = 0
End Property

Public Property Let Level(ByVal vData As Long)
    mvarLevel = vData
End Property

Public Property Get Level() As Long
    Level = mvarLevel
End Property

Public Property Let 内部序号(ByVal vData As Long)
    mvar内部序号 = vData
End Property

Public Property Get 内部序号() As Long
    内部序号 = mvar内部序号
End Property

'################################################################################################################
'## 功能：  复制提纲对象
'##
'## 参数：  blnWithKey  :是否连同Key值一起复制！
'################################################################################################################
Public Function Clone(Optional blnWithKey As Boolean = False) As cEPRCompend
    Dim cNode As New cEPRCompend
    If blnWithKey Then
        cNode.Key = Me.Key
    End If
    cNode.ID = Me.ID
    cNode.文件ID = Me.文件ID
    cNode.开始版 = Me.开始版
    cNode.终止版 = Me.终止版
    cNode.父ID = Me.父ID
    cNode.父Key = Me.父Key
    cNode.对象序号 = Me.对象序号
    cNode.保留对象 = Me.保留对象
    cNode.名称 = Me.名称
    cNode.说明 = Me.说明
    cNode.定义提纲ID = Me.定义提纲ID
    cNode.预制提纲ID = Me.预制提纲ID
    cNode.复用提纲 = Me.复用提纲
    cNode.使用时机 = Me.使用时机
    cNode.Level = Me.Level
    cNode.内部序号 = Me.内部序号
    Set Clone = cNode
End Function

Public Function FillCompendMember(ByVal rs As ADODB.Recordset, ByVal strTableName As String) As Boolean
'功能：用数据记录填充类成员
    If rs Is Nothing Then Exit Function
    If rs.EOF Then Exit Function
    
    If NVL(rs!对象标记, 0) <> 0 Then Me.Key = NVL(rs!对象标记, 0)
    Me.Level = NVL(rs!Level, 0)
    Me.ID = rs!ID
    Me.文件ID = NVL(rs!文件ID, 0)
    Me.父ID = NVL(rs!父ID, 0)
    Me.对象序号 = NVL(rs!对象序号, 0)
    Me.保留对象 = IIf(NVL(rs!保留对象, 0) = 0, False, True)
    Me.名称 = NVL(rs!内容文本)
    Me.说明 = NVL(rs!对象属性)
    Me.定义提纲ID = NVL(rs!定义提纲ID, 0)
    Me.预制提纲ID = NVL(rs!预制提纲ID, 0)
    Me.复用提纲 = IIf(NVL(rs!复用提纲, 0) = 0, False, True)
    Me.使用时机 = NVL(rs!使用时机)
    Me.内部序号 = Me.对象序号
    
    If InStr(strTableName, "电子病历内容") > 0 Then '病历文件结构和病历范文内容没有这两个字段,但都共用这个函数
        Me.开始版 = NVL(rs!开始版, 1)
        Me.终止版 = NVL(rs!终止版, 0)
    End If
    
    FillCompendMember = True
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
'## 功能：  插入提纲到Editor中
'##
'## 参数：  edtThis         :当前的编辑器控件
'##         lngStartPos     :当前位置
'##         blnFirstInsert  :是不是第一次插入（此时，末尾自动新增独立一行带段落格式的文本，内容与提纲名称相同）
'##         objDocument     :Document对象，用于处理特殊预制提纲时新增表格对象等的特殊处理。
'################################################################################################################
Public Function InsertIntoEditor(ByRef edtThis As Object, _
    Optional ByVal lngStartPos As Long = -1, _
    Optional ByVal blnFirstInsert As Boolean = False, _
    Optional ByRef objDocument As cEPRDocument) As Boolean
        
    If edtThis.ViewMode <> cprNormal Then InsertIntoEditor = False: Exit Function
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    If lngStartPos = -1 Then lngStartPos = edtThis.Selection.StartPos
    bInKeys = IsBetweenAnyKeys(edtThis, lngStartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys Then InsertIntoEditor = False: Exit Function    '保证不能插入关键字内部

    Dim p As Long, lngKey As Long, blnForce As Boolean, strTmp As String, lngLen As Long
    blnForce = edtThis.ForceEdit

    With edtThis
        .Freeze
        .ForceEdit = True
        
        strTmp = Format(Me.Key, "00000000") & "," & IIf(Me.保留对象, 1, 0) & ",0)"
        
        .Range(lngStartPos, lngStartPos).Font.Protected = False
        .Range(lngStartPos, lngStartPos).Font.Hidden = False
        If lngStartPos = 0 Then
            .Range(lngStartPos, lngStartPos) = "OS(" & strTmp & "OE(" & strTmp
            .Range(lngStartPos, lngStartPos + 32).Font.Hidden = True
            .Range(lngStartPos, lngStartPos + 32).Font.Protected = True
        Else
            If .Range(lngStartPos - 2, lngStartPos) = vbCrLf Then
                .Range(lngStartPos, lngStartPos) = "OS(" & strTmp & "OE(" & strTmp
                .Range(lngStartPos, lngStartPos + 32).Font.Hidden = True
                .Range(lngStartPos, lngStartPos + 32).Font.Protected = True
            Else
                .Range(lngStartPos, lngStartPos) = vbCrLf & "OS(" & strTmp & "OE(" & strTmp
                .Range(lngStartPos, lngStartPos + 2).Font.Protected = False
                .Range(lngStartPos, lngStartPos + 2).Font.Hidden = False
                lngStartPos = lngStartPos + 2
                .Range(lngStartPos, lngStartPos + 32).Font.Hidden = True
                .Range(lngStartPos, lngStartPos + 32).Font.Protected = True
            End If
        End If
        If Me.预制提纲ID < 0 Then
            '预制提纲的特殊处理
            If blnFirstInsert Then
'               SetCommonStyle edtThis, Me.Level, lngStartPos, lngStartPos + 32, True
                lngStartPos = lngStartPos + 32
                .Range(lngStartPos, lngStartPos) = Me.名称 & vbCrLf
                lEE = lngStartPos + Len(Me.名称)
'               SetCommonStyle edtThis, Me.Level, lngStartPos, lEE, True
                .Range(lngStartPos, lEE + 2).Font.Protected = False
                .Range(lngStartPos, lEE + 2).Font.Hidden = False
                lngStartPos = lEE + 2
                
            End If
        Else
            lngStartPos = lngStartPos + 32
            If blnFirstInsert Then
'                SetCommonStyle edtThis, Me.Level, lngStartPos - 32, lngStartPos, True
                .Range(lngStartPos, lngStartPos) = Me.名称 & vbCrLf
                lEE = lngStartPos + Len(Me.名称)
'                SetCommonStyle edtThis, Me.Level, lngStartPos, lEE, True
                .Range(lngStartPos, lEE + 2).Font.Protected = False
                .Range(lngStartPos, lEE + 2).Font.Hidden = False
                lngStartPos = lEE + 2
'                SetCommonStyle edtThis, "正文", lngStartPos, lngStartPos, True
            End If
        End If
        .Range(lngStartPos, lngStartPos).Font.Protected = False
        .Range(lngStartPos, lngStartPos).Font.Hidden = False
        .Range(lngStartPos, lngStartPos).Selected
                
        .ForceEdit = blnForce
        .UnFreeze
    End With
    InsertIntoEditor = True
End Function

'################################################################################################################
'## 功能：  获取提纲保存到数据库的SQL语句
'##
'## 参数：  eEdtType        :IN，       当前编辑模式
'##         ArraySQL()      :IN/OUT，   SQL数组
'################################################################################################################
Public Function GetSaveSQL(ByVal eEdtType As EditTypeEnum, ByRef ArraySQL() As String) As Boolean
    Dim lngCount As Long
    lngCount = UBound(ArraySQL) + 1
    ReDim Preserve ArraySQL(1 To lngCount) As String
    
    Call CheckValidID(eEdtType)         '检查ID有效性！
    Select Case eEdtType
    Case cprET_病历文件定义
        gstrSQL = "Zl_病历文件结构_Update(" & Me.ID & "," & Me.文件ID & "," & IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & _
            Me.对象序号 & ",1," & Me.Key & "," & IIf(Me.保留对象, 1, 0) & ",'" & Me.说明 & "',NULL,'" & Me.名称 & "',NULL," & _
            IIf(Me.预制提纲ID = 0, "NULL", Me.预制提纲ID) & "," & IIf(Me.复用提纲, 1, 0) & ",'" & Me.使用时机 & "')"
    Case cprET_全文示范编辑
        gstrSQL = "Zl_病历范文内容_Update(" & Me.ID & "," & Me.文件ID & "," & IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & _
            Me.对象序号 & ",1," & Me.Key & "," & IIf(Me.保留对象, 1, 0) & ",'" & Me.说明 & "',NULL,'" & Me.名称 & "',NULL," & _
            IIf(Me.定义提纲ID = 0, "NULL", Me.定义提纲ID) & "," & IIf(Me.预制提纲ID = 0, "NULL", Me.预制提纲ID) & "," & IIf(Me.复用提纲, 1, 0) & ",'" & Me.使用时机 & "')"
    Case cprET_单病历编辑, cprET_单病历审核
        '加了 开始版，终止版 字段，其他相同
        gstrSQL = "ZL_电子病历内容_Update(" & Me.ID & "," & Me.文件ID & "," & Me.开始版 & "," & Me.终止版 & "," & IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & _
            Me.对象序号 & ",1," & Me.Key & "," & IIf(Me.保留对象, 1, 0) & ",'" & Me.说明 & "',NULL,'" & Me.名称 & "',NULL," & _
            IIf(Me.定义提纲ID = 0, "NULL", Me.定义提纲ID) & "," & IIf(Me.预制提纲ID = 0, "NULL", Me.预制提纲ID) & "," & IIf(Me.复用提纲, 1, 0) & ",'" & Me.使用时机 & "')"
    End Select
    ArraySQL(lngCount) = gstrSQL
    GetSaveSQL = True
End Function

'################################################################################################################
'## 功能：  定位到文本中的提纲开头位置
'##
'## 参数：  edtThis     :Editor编辑器
'################################################################################################################
Public Sub GotoStartPos(ByRef edtThis As Object)
    Dim lSS As Long, lSE As Long, lES As Long, lEE As Long, bNeeded As Boolean, bFinded As Boolean
    bFinded = FindKey(edtThis, "O", Me.Key, lSS, lSE, lES, lEE, bNeeded)
    If bFinded Then
        edtThis.TOM.TextDocument.Range(lEE, lEE).Select
'        edtThis.Range(lEE, lEE).ScrollIntoView cprSPStart
        If edtThis.Visible And edtThis.Enabled Then edtThis.SetFocus
    End If
End Sub

'################################################################################################################
'## 功能：  获取提纲内容的起始和终止位置
'##
'## 参数：  edtThis     :Editor编辑器
'##         lngStartPos :返回起始位置
'##         lngEndPos   :返回结束位置
'##
'## 说明：  提纲末尾有一个隐藏并保护的 vbCRLF。
'################################################################################################################
Public Function GetPosition(ByRef edtThis As Object, ByRef lngStartPos As Long, ByRef lngEndPos As Long) As Boolean
    Dim lSS As Long, lSE As Long, lES As Long, lEE As Long, bNeeded As Boolean, bFinded As Boolean, lKey As Long
    bFinded = FindKey(edtThis, "O", Me.Key, lSS, lSE, lES, lEE, bNeeded)
    If bFinded Then
        lngStartPos = lEE - 32   '加上一个vbCrLf的位置
    Else
        GetPosition = False
        Exit Function
    End If
    
    bFinded = FindNextKey(edtThis, lEE, "O", lKey, lSS, lSE, lES, lEE, bNeeded)
    If bFinded Then
        lngEndPos = lSS
    Else
        lngEndPos = Len(edtThis.Text)
    End If
    GetPosition = True
End Function
