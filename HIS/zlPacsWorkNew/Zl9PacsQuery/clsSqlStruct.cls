VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSqlStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

 

Private mobjParent As clsSqlStruct      '父查询
Private mobjRoot As clsSqlStruct        '根查询

Private mstrAnalyzeSql As String
Private mlngSqlFromType As Long  'sql来源 0-来自With,1-来自Union, 2-来自Minus, 3-来自select, 4-来自From, 5-来自Where, 6-来自func
Private mlngDepth As Long

Private mobjBracket As clsSqlBracket

Private mobjWiths() As clsSqlStruct
Private mobjUnion() As clsSqlStruct
Private mobjMinus() As clsSqlStruct

Private mobjSelect() As clsSqlItem
Private mobjFrom() As clsSqlItem
Private mobjWhere() As clsSqlItem
Private mobjFunc() As clsSqlItem

Private maryAllParameters() As String


Public Others As String
Public BracketTag As String


Private aryStrNull(0) As String
Private aryBoolNull(0) As Boolean
    

'with
Property Get WithCount() As Long
    WithCount = UBound(mobjWiths)
End Property

Property Get Withs(ByVal lngIndex As Long) As clsSqlStruct
    Set Withs = mobjWiths(lngIndex)
End Property

Property Set Withs(ByVal lngIndex As Long, ByRef objSqlStruct As clsSqlStruct)
    Set mobjWiths(lngIndex) = objSqlStruct
End Property

Public Sub AddWiths(ByRef objWith As clsSqlStruct)
    Dim lngUbound As Long
    
    lngUbound = UBound(mobjWiths) + 1
    
    ReDim Preserve mobjWiths(lngUbound)
    Set mobjWiths(lngUbound) = objWith
End Sub

'union
Property Get UnionCount() As Long
    UnionCount = UBound(mobjUnion)
End Property

Property Get Union(ByVal lngIndex As Long) As clsSqlStruct
    Set Union = mobjUnion(lngIndex)
End Property

Property Set Union(ByVal lngIndex As Long, ByRef objSqlStruct As clsSqlStruct)
    Set mobjUnion(lngIndex) = objSqlStruct
End Property

Public Sub AddUnion(ByRef objUnion As clsSqlStruct)
    Dim lngUbound As Long
    
    lngUbound = UBound(mobjUnion) + 1
    
    ReDim Preserve mobjUnion(lngUbound)
    Set mobjUnion(lngUbound) = objUnion
End Sub

'minus
Property Get MinusCount() As Long
    MinusCount = UBound(mobjMinus)
End Property

Property Get Minus(ByVal lngIndex As Long) As clsSqlStruct
    Set Minus = mobjMinus(lngIndex)
End Property

Property Set Minus(ByVal lngIndex As Long, ByRef objSqlStruct As clsSqlStruct)
    Set mobjMinus(lngIndex) = objSqlStruct
End Property

Public Sub AddMinus(ByRef objMinus As clsSqlStruct)
    Dim lngUbound As Long
    
    lngUbound = UBound(mobjMinus) + 1
    
    ReDim Preserve mobjMinus(lngUbound)
    Set mobjMinus(lngUbound) = objMinus
End Sub

'select
Property Get SelectCount() As Long
    SelectCount = UBound(mobjSelect)
End Property

Property Get Selects(ByVal lngIndex As Long) As clsSqlItem
    Set Selects = mobjSelect(lngIndex)
End Property

Property Set Selects(ByVal lngIndex As Long, ByRef objSqlItem As clsSqlItem)
    Set mobjSelect(lngIndex) = objSqlItem
End Property

Public Sub AddSelect(ByRef objSelect As clsSqlItem)
    Dim lngUbound As Long
    
    lngUbound = UBound(mobjSelect) + 1
    
    ReDim Preserve mobjSelect(lngUbound)
    Set mobjSelect(lngUbound) = objSelect
End Sub

'from
Property Get FromCount() As Long
    FromCount = UBound(mobjFrom)
End Property

Property Get From(ByVal lngIndex As Long) As clsSqlItem
    Set From = mobjFrom(lngIndex)
End Property

Property Set From(ByVal lngIndex As Long, ByRef objSqlItem As clsSqlItem)
    Set mobjFrom(lngIndex) = objSqlItem
End Property

Public Sub AddFrom(ByRef objFrom As clsSqlItem)
    Dim lngUbound As Long
    
    lngUbound = UBound(mobjFrom) + 1
    
    ReDim Preserve mobjFrom(lngUbound)
    Set mobjFrom(lngUbound) = objFrom
End Sub

'where
Property Get WhereCount() As Long
    WhereCount = UBound(mobjWhere)
End Property

Property Get Where(ByVal lngIndex As Long) As clsSqlItem
    Set Where = mobjWhere(lngIndex)
End Property

Property Set Where(ByVal lngIndex As Long, ByRef objSqlItem As clsSqlItem)
    Set mobjWhere(lngIndex) = objSqlItem
End Property

Public Sub AddWhere(ByRef objWhere As clsSqlItem)
    Dim lngUbound As Long
    
    lngUbound = UBound(mobjWhere) + 1
    
    ReDim Preserve mobjWhere(lngUbound)
    Set mobjWhere(lngUbound) = objWhere
End Sub

'func
Property Get FuncCount() As Long
    FuncCount = UBound(mobjFunc)
End Property

Property Get Func(ByVal lngIndex As Long) As clsSqlItem
    Set Func = mobjFunc(lngIndex)
End Property

Property Set Func(ByVal lngIndex As Long, ByRef objSqlItem As clsSqlItem)
    Set mobjFunc(lngIndex) = objSqlItem
End Property

Public Sub AddFunc(ByRef objFunc As clsSqlItem)
    Dim lngUbound As Long
    
    lngUbound = UBound(mobjFunc) + 1
    
    ReDim Preserve mobjFunc(lngUbound)
    Set mobjFunc(lngUbound) = objFunc
End Sub


'parent
Property Get Parent() As clsSqlStruct
    Set Parent = mobjParent
End Property

'root
Property Get Root() As clsSqlStruct
    Set Root = mobjRoot
End Property

'bracket
Property Get Bracket() As clsSqlBracket
    Set Bracket = mobjBracket
End Property

Property Set Bracket(ByRef objBrack As clsSqlBracket)
    Set mobjBracket = objBrack
End Property

'sql
Property Get Sql() As String
    Sql = mstrAnalyzeSql
End Property

Property Get FromType() As Long
    FromType = mlngSqlFromType
End Property

Property Get Depth() As Long
    Depth = mlngDepth
End Property


'parameter
Property Get ParCount() As Long
    ParCount = UBound(maryAllParameters)
End Property

Property Get AllParameter(ByVal lngIndex As Long) As String
    AllParameter = maryAllParameters(lngIndex)
End Property

Property Get AllParameterAry() As String()
    AllParameterAry = maryAllParameters
End Property



'*********************************************************************************************
'*********************************************************************************************

Private Sub init()
    mstrAnalyzeSql = ""
    mlngSqlFromType = -1
    mlngDepth = 0
    
    Set mobjParent = Nothing
    Set mobjRoot = Nothing
    
    Others = ""
    BracketTag = ""
    
    ReDim mobjWiths(0)
    ReDim mobjUnion(0)
    ReDim mobjMinus(0)

    ReDim mobjSelect(0)
    ReDim mobjFrom(0)
    ReDim mobjWhere(0)
    ReDim mobjFunc(0)

    ReDim maryAllParameters(0)
End Sub

Public Function HasParName(ByVal strPar As String) As Boolean
'判断是否存在指定par
    Dim i As Long
    Dim strCurPar As String
    
    HasParName = False
    For i = 1 To UBound(maryAllParameters)
        strCurPar = maryAllParameters(i)
        
        If (InStr(strCurPar, "[@") > 0) Or (InStr(strCurPar, "[*") > 0) Then
            strCurPar = Mid$(strCurPar, 3, InStr(strCurPar, ",") - 3)
        Else
            strCurPar = Mid$(strCurPar, 2, Len(strCurPar) - 2)
        End If
        
        If strCurPar = UCase(strPar) Then
            HasParName = True
            Exit Function
        End If
    Next i
End Function

Public Sub Load(ByVal strFormatSql As String)
On Error GoTo errHandle:
    Dim strSourceSql As String
    
    Call init

    strSourceSql = ResolveBracket(strFormatSql, mobjBracket, maryAllParameters)
    
    Call Analyze(strSourceSql)
Exit Sub
errHandle:
    Err.Raise -1, "clsSqlStruct.Analyze", "[Analyze]处理错误>>" + vbCrLf + "  解析语句为:" + strFormatSql + vbCrLf + Err.Description
    Resume
End Sub



Public Sub Analyze(ByVal strSource As String, _
    Optional ByVal lngFromType As Long = -1, _
    Optional ByVal lngDepth As Long = 0, _
    Optional ByRef objParent As clsSqlStruct = Nothing, _
    Optional ByRef objRoot As clsSqlStruct = Nothing)
    
On Error GoTo errHandle
    Dim strSourceSql As String
    
    Dim strWithPart As String
    
    Dim aryWith() As String
    Dim aryUnion() As String
    Dim aryMinus() As String
    
    mstrAnalyzeSql = strSource
    mlngSqlFromType = lngFromType
    mlngDepth = lngDepth
    
    Set mobjParent = objParent
    Set mobjRoot = objRoot
    
    '获取with部分的sql并解析
    strSourceSql = GetWithContext(strSource, strWithPart)
    If Len(strWithPart) > 0 Then
        aryWith = GetWithPart(strWithPart)
        mobjWiths = AnalyzeSql(aryWith, aryWith, 0, objParent, objRoot)
    End If
    
    '获取union部分的sql并解析
    aryUnion = GetParallelPart(strSourceSql, pstUnion)
    mobjUnion = AnalyzeSql(aryUnion, aryUnion, 1, objParent, objRoot)
    
    '获取minus部分的sql并解析
    aryMinus = GetParallelPart(strSourceSql, pstMinus)
    mobjMinus = AnalyzeSql(aryMinus, aryMinus, 2, objParent, objRoot)
    
    If UBound(aryUnion) <= 0 And UBound(aryMinus) <= 0 Then
        '判断是否存在where,
        
        '如果没有union和minus，说明可能是单个的select或者过程函数
        Call AnalyzeSqlPart(strSourceSql, lngFromType, objParent, objRoot)
    End If

Exit Sub
errHandle:
    Err.Raise -1, "clsSqlStruct.Analyze", "[Analyze]处理错误>>" + vbCrLf + "  解析语句为:" + strSource + vbCrLf + Err.Description
    Resume
End Sub


Private Function AnalyzeSql(arySql() As String, _
                            ByRef strBracketTag() As String, _
                            ByRef lngFromType As Long, _
                            ByRef objParent As clsSqlStruct, _
                            ByRef objRoot As clsSqlStruct) As clsSqlStruct()
'解析sql语句

    Dim objCurSqlStruct As clsSqlStruct
    Dim aryResult() As clsSqlStruct
    Dim lngAryLen As Long
    Dim lngUbound As Long
    Dim blnHasPar As Boolean
    Dim blnHasSelect As Boolean
    Dim lngTagCount As Long
     
    Dim i As Long
    Dim strTmp As String
    
On Error GoTo errHandle

    ReDim aryResult(0)
    
    lngTagCount = UBound(strBracketTag)
    lngAryLen = UBound(arySql)
    
    For i = 1 To lngAryLen
        strTmp = arySql(i)
        
        '恢复当前sql中包含的{%0#1}形式的内容,with或者from部分中不一定是select语句，可能存在table(cast(形式
        '如果本身就是sql语句，则不需要恢复select语句中包含的{%0#1}形式内容
        If Not HasSelect(strTmp) Then
            '当前不是select语句
            strTmp = RestoreBracketContext(arySql(i), mobjBracket)
        End If
        
        Set objCurSqlStruct = New clsSqlStruct
        
        Set objCurSqlStruct.Bracket = mobjBracket
        
        objCurSqlStruct.BracketTag = ""
        If lngTagCount > 0 Then
            objCurSqlStruct.BracketTag = strBracketTag(i)
        End If
    
        Call objCurSqlStruct.Analyze(strTmp, lngFromType, mlngDepth + 1, objParent, objRoot)
        
        lngUbound = UBound(aryResult) + 1
        ReDim Preserve aryResult(lngUbound)
        
        Set aryResult(lngUbound) = objCurSqlStruct
    Next i

    AnalyzeSql = aryResult
Exit Function
errHandle:
    Err.Raise -1, "clsSqlStruct.AnalyzeSql", "[AnalyzeSql]处理错误>>" + vbCrLf + "  解析语句为:" + strTmp + vbCrLf + Err.Description
    Resume
End Function

Private Sub AnalyzeSqlPart(strSql As String, _
                            ByRef lngFromType As Long, _
                            ByRef objParent As clsSqlStruct, _
                            ByRef objRoot As clsSqlStruct)
On Error GoTo errHandle
    Dim strWithPart As String
    Dim strSelectPart As String
    Dim strFromPart As String
    Dim strWherePart As String
    Dim strOtherPart As String
    Dim strFuncPart As String
    
    Dim aryItem() As String
    
    
    Call Parse(strSql, strWithPart, strSelectPart, strFromPart, strWherePart, strOtherPart, strFuncPart)

    'with部分，存在如下查询形式
'        with
'        a as (select 1 from dual),
'        b as (select 2 from dual)
'        select * from (with
'                       c as (select 3 from dual),
'                       d as (select 4 from dual)
'                       select * from c, d
'        ), a, b

    If Trim(strWithPart) <> "" Then
        aryItem = GetWithPart(strWithPart)
        mobjWiths = AnalyzeSql(aryItem(), aryItem(), 0, objParent, objRoot)
    End If

    'select部分
    If Trim(strSelectPart) <> "" Then
        aryItem = GetSelectFromPart(strSelectPart)
        mobjSelect = AnalyzeSqlItem(aryItem(), aryStrNull, aryBoolNull, 3, objParent, objRoot)
    End If

    'from部分
    If Trim(strFromPart) <> "" Then
        aryItem = GetSelectFromPart(strFromPart)
        mobjFrom = AnalyzeSqlItem(aryItem(), aryStrNull, aryBoolNull, 4, objParent, objRoot)
    End If

    'where部分
    If Trim(strWherePart) <> "" Then
        aryItem = GetWherePart(strWherePart)
        mobjWhere = AnalyzeSqlItem(aryItem(), aryStrNull, aryBoolNull, 5, objParent, objRoot)
    End If

    'func部分
    If Trim(strFuncPart) <> "" Then
        ReDim aryItem(1)
        aryItem(1) = strFuncPart
        mobjFunc = AnalyzeSqlItem(aryItem(), aryStrNull, aryBoolNull, 6, objParent, objRoot)
    End If
     
    'other部分
    Others = strOtherPart
Exit Sub
errHandle:
    Err.Raise -1, "clsSqlStruct.AnalyzeSqlPart", "[AnalyzeSqlPart]处理错误>>" + vbCrLf + "  解析语句为:" + strSql + vbCrLf + Err.Description
    Resume
End Sub


Private Function AnalyzeSqlItem(arySqlPart() As String, _
                            ByRef aryBracketTag() As String, _
                            ByRef aryLinkWhere() As Boolean, _
                            ByVal lngFromType As Long, _
                            ByRef objParent As clsSqlStruct, _
                            ByRef objRoot As clsSqlStruct, _
                            Optional ByRef objitem As clsSqlItem = Nothing) As clsSqlItem()
    Dim i As Long
    Dim lngAryLen As Long
    Dim strTmp As String
    Dim aryResult() As clsSqlItem
    Dim arySubSql() As String
    Dim arySubFunc() As String
    Dim arySqlTag() As String
    Dim aryFuncTag() As String
    Dim aryParNo() As Long
    
    Dim aryFuncLink() As Boolean
    Dim aryWhereSubItem() As String
    
    Dim blnHasSelect As Boolean
    Dim objCurSqlItem As clsSqlItem
    Dim objParentItem As clsSqlItem
    Dim lngUbound As Long
    Dim lngTagCount As Long
    Dim lngLinkCount As Long
    
    Dim blnHasItemPar As Boolean
    Dim blnHasSubPar As Boolean
    
    Dim blnIsParse As Boolean
    
On Error GoTo errHandle
     
    ReDim aryResult(0)
    
    lngTagCount = UBound(aryBracketTag)
    lngLinkCount = UBound(aryLinkWhere)
    lngAryLen = UBound(arySqlPart)
     
    For i = 1 To lngAryLen
        strTmp = arySqlPart(i)
        
        Set objCurSqlItem = New clsSqlItem
        Set objCurSqlItem.Parent = objitem
        
        objCurSqlItem.BracketTag = ""
        If lngTagCount > 0 Then
            objCurSqlItem.BracketTag = aryBracketTag(i)
        End If
        
        
        objCurSqlItem.Context = strTmp
        '先判断是否有参数存在, 如 [1] + Id In(...), 或 Num+[1]=([2]+[3])
        blnHasItemPar = GetParNos(strTmp, aryParNo)
        

        blnIsParse = False
        If lngLinkCount > 0 Then
            blnIsParse = aryLinkWhere(i)
        End If
        
        If blnIsParse Then
            objCurSqlItem.HasLinkWhere = True
            
            '如果有括弧，需要先取消括弧
            strTmp = Mid$(strTmp, 2, Len(strTmp) - 2)
            '如 id=[1] or id={%0#1#1} or id=[3] 需要先分解
            aryWhereSubItem = GetWherePart(strTmp)
            
            '分析item中包含的子项
            Call objCurSqlItem.AddSubItemAry(AnalyzeSqlItem(aryWhereSubItem, aryStrNull, aryBoolNull, _
                                            lngFromType, objParent, objRoot, objCurSqlItem))
        Else
            Call GetBracketDetail(arySqlPart(i), mobjBracket, blnHasSubPar, blnHasSelect, _
                                    arySubSql, arySubFunc, arySqlTag, aryFuncTag, aryFuncLink)
             
            
            '分析item中包含的sql
            Call objCurSqlItem.AddSubSqlAry(AnalyzeSql(arySubSql, arySqlTag, lngFromType, objParent, objRoot))
        
            '分析item中包含的子项
            Call objCurSqlItem.AddSubItemAry(AnalyzeSqlItem(arySubFunc, aryFuncTag, aryFuncLink, _
                                            lngFromType, objParent, objRoot, objCurSqlItem))

        End If
        
 
        If blnHasSubPar Or blnHasItemPar Then
            objCurSqlItem.HasPar = blnHasItemPar
            If blnHasItemPar Then
                Call objCurSqlItem.AddParNoAry(aryParNo)
            End If
            
            Set objParentItem = objCurSqlItem.Parent
            While Not objParentItem Is Nothing
                objParentItem.HasSubPar = True
                Set objParentItem = objParentItem.Parent
            Wend
        End If
        
        If blnHasSelect Then
            objCurSqlItem.HasSubSql = True
            
            Set objParentItem = objCurSqlItem.Parent
            While Not objParentItem Is Nothing
                objParentItem.HasSubSql = True
                Set objParentItem = objParentItem.Parent
            Wend
        End If
        
        lngUbound = UBound(aryResult) + 1
        ReDim Preserve aryResult(lngUbound)
        
        Set aryResult(lngUbound) = objCurSqlItem
    Next i

    AnalyzeSqlItem = aryResult
Exit Function
errHandle:
    Err.Raise -1, "clsSqlStruct.AnalyzeSqlItem", "[AnalyzeSqlItem]处理错误>>" + vbCrLf + "  解析语句为:" + strTmp + vbCrLf + Err.Description
    Resume
End Function

Private Sub Class_Initialize()
    Call init
End Sub

Private Sub Class_Terminate()
    Set mobjParent = Nothing
    Set mobjRoot = Nothing
    Set mobjBracket = Nothing
    
    ReDim mobjWiths(0)
    Set mobjWiths(0) = Nothing
    
    ReDim mobjUnion(0)
    Set mobjUnion(0) = Nothing
    
    ReDim mobjMinus(0)
    Set mobjMinus(0) = Nothing
    
    ReDim mobjSelect(0)
    Set mobjSelect(0) = Nothing
    
    ReDim mobjFrom(0)
    Set mobjFrom(0) = Nothing
    
    ReDim mobjWhere(0)
    Set mobjWhere(0) = Nothing
    
    ReDim mobjFunc(0)
    Set mobjFunc(0) = Nothing
End Sub
