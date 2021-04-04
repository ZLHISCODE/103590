VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSqlParse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 

Private mstrSourceSql As String         '原始查询
Private mstrFormatSql As String         '格式化后的查询，去掉了回车，双引号，开始括弧
Private mstrQuerySql As String          '对参数进行处理后的查询，将参数替换成了[1] [2]的形式

Private mvarParValues() As Variant      '参数值

Private mstrQuotes() As String
Private mobjSqlStruct As New clsSqlStruct

Private mblnNullParameterHint As Boolean
Private mblnNullWhereHint As Boolean


Private mstrSqlProcessing As String


Private mstrAttachItem As String
Private mstrAttachData As Variant
Private mblnIsParSymbolTransfer As Boolean
Private mlngResetWhereIndex() As Long       '需要重置检索条件的索引，参数中使用了“*”号的，表示查询的检索条件方式可以根据选择进行配置

Public IsInit As Boolean

 

'参数值获取事件,
'strParName:参数名称
'lngParType:参数类型 0-系统参数，1-条件参数
'value:参数返回值
Public Event OnGetParameterValue(ByVal strParName As String, ByRef Value As Variant)

'***********************************************************************************************
'***********************************************************************************************

'原始SQL
Property Get Sql() As String
    Sql = mstrSourceSql
End Property

'参数值
Property Get ParValue(ByVal lngIndex As Long) As Variant
    ParValue = mvarParValues(lngIndex)
End Property

Property Get ParValues() As Variant()
    ParValues = mvarParValues
End Property

'sql结构对象
Property Get SqlStruct() As clsSqlStruct
    Set SqlStruct = mobjSqlStruct
End Property

'空参数提醒
Property Get NullParameterHint() As Boolean
    NullParameterHint = mblnNullParameterHint
End Property

Property Let NullParameterHint(Value As Boolean)
    mblnNullParameterHint = Value
End Property

'空条件提醒
Property Get NullWhereHint() As Boolean
    NullWhereHint = mblnNullWhereHint
End Property

Property Let NullWhereHint(Value As Boolean)
    mblnNullWhereHint = Value
End Property


'是否针对参数符号"[]"转义
Property Get IsParSymbolTransfer() As Boolean
    IsParSymbolTransfer = mblnIsParSymbolTransfer
End Property

Property Let IsParSymbolTransfer(Value As Boolean)
    mblnIsParSymbolTransfer = Value
End Property

'***********************************************************************************************
'***********************************************************************************************

Public Sub init(ByVal strSql As String)
    mblnNullParameterHint = False
    mblnNullWhereHint = False
    
    mstrSourceSql = strSql
    mstrFormatSql = FormatSql(mstrSourceSql, mstrQuotes)
    
    '解析sql语句
    'mstrFormatSql调用load后，参数已经被[1],[2],...的形式替换
    Call mobjSqlStruct.Load(mstrFormatSql)
    
    IsInit = True
End Sub


Public Function GetTestSql(Optional ByVal strSql As String = "") As String
'获取测试的SQL语句
    Dim strTestSql As String
    Dim lngCount As Long
    Dim i As Long
    Dim objSelf As clsSqlParse
    Dim strPar As String
    Dim strTmp As String
    Dim arrTmp() As String
    Dim blnIsParSymbolTransfer As Boolean
    
    strTestSql = strSql
    If Trim(strTestSql) = "" Then
        strTestSql = mstrFormatSql
        '测试当前初始化所用的Sql
        lngCount = mobjSqlStruct.ParCount
       
        For i = 1 To lngCount
            strPar = mobjSqlStruct.AllParameter(i)
            
            '判断是否可选参数
            If (InStr(strPar, "[@") > 0) Or (InStr(strPar, "[*") > 0) Then
                strTmp = Mid$(strPar, 2, Len(strPar) - 2)
                arrTmp = Split(strTmp, ",")
                If UBound(arrTmp) > 0 Then
                    strTmp = arrTmp(1)
                End If
                
                strTestSql = Replace(strTestSql, strPar, strTmp)
            Else
                strTestSql = Replace(strTestSql, strPar, " NULL ")
            End If
            
        Next i
        
        '将恢复双引号之间的内容
        strTestSql = RestoreQuote(strTestSql, mstrQuotes, blnIsParSymbolTransfer)
        
        '为真时表示对单引号中含有的“[]”符号进行了转义，因此在实际查询时，需要替换回来
        If blnIsParSymbolTransfer Then
            strTestSql = Replace(Replace(strTestSql, "&lbkt;", "["), "&rbkt;", "]")
        End If
    Else
        Set objSelf = New clsSqlParse
        
        Call objSelf.init(strSql)
        
        strTestSql = objSelf.GetTestSql()
    End If
    
    GetTestSql = strTestSql
End Function



Public Function GetQuerySql(Optional ByVal blnIsRefresh As Boolean = False) As String
    mstrAttachItem = ""
    mstrAttachData = ""
    
    GetQuerySql = RecombineSql
End Function



Public Function GetAttachSql(ByVal strAttachItem As String, ByVal strAttachData As Variant) As String
'读取医嘱
    mstrAttachItem = strAttachItem
    mstrAttachData = strAttachData
    
    GetAttachSql = RecombineSql
End Function


'***********************************************************************************************************************************


Private Function RecombineSql(Optional ByVal blnIsRefresh As Boolean = False)
'重组sql语句
    Dim i As Long
    Dim lngParIndex As Long
    Dim objCurSqlStruct As clsSqlStruct
      
    mstrSqlProcessing = mstrFormatSql ' mstrQuerySql
    mblnIsParSymbolTransfer = False
    
    '读取参数
    '如果是刷新，则不需要重复读取参数值
    If Not blnIsRefresh Then
        Call ReadParameterValue
    End If
    
    '分析参数为null的条件
    If AnalyzeSqls(mobjSqlStruct) = False Then
        RecombineSql = ""
        Exit Function
    End If
    
    '处理实现查询的条件判断方式，
    For i = 1 To UBound(mlngResetWhereIndex)
        lngParIndex = mlngResetWhereIndex(i)
        mstrSqlProcessing = Replace(mstrSqlProcessing, mobjSqlStruct.AllParameter(lngParIndex), mvarParValues(lngParIndex))
    Next
    
    
    '格式化null参数条件的sql语句
    mstrSqlProcessing = FormatNull(mstrSqlProcessing)
    
    '恢复双引号之间的内容
    mstrSqlProcessing = RestoreQuote(mstrSqlProcessing, mstrQuotes, mblnIsParSymbolTransfer)
    
    '恢复参数中的参数索引
    mstrSqlProcessing = RestoreParIndex(mstrSqlProcessing, mobjSqlStruct.AllParameterAry())
        
    RecombineSql = mstrSqlProcessing
End Function


Private Sub ReadParameterValue()
'读取参数值
    Dim i As Long
    Dim Value As Variant
    Dim lngUbound As Long
    Dim strPar As String
    Dim blnIsNeedRange As Boolean
    Dim lngResetUbound As Long
    Dim blnIsResetWhere As Boolean
    Dim strTmp As String
    Dim arrTmp() As String
    
    '读取系统参数
    lngUbound = mobjSqlStruct.ParCount
    
    ReDim mvarParValues(lngUbound)
    ReDim mlngResetWhereIndex(0)
    
    For i = 1 To lngUbound
        strPar = mobjSqlStruct.AllParameter(i)
        Value = Null
        
        If strPar = mstrAttachItem Or strPar = "[" & mstrAttachItem & "]" Then
            If Len(mstrAttachData) > 0 Then
                Value = mstrAttachData
            End If
        Else
            blnIsResetWhere = InStr(strPar, "[*") > 0
            If blnIsResetWhere Then
                lngResetUbound = UBound(mlngResetWhereIndex)
                
                ReDim Preserve mlngResetWhereIndex(lngResetUbound + 1)
                mlngResetWhereIndex(lngResetUbound + 1) = i
            End If
            
            If (InStr(strPar, "[@") > 0) Or blnIsResetWhere Then
                strPar = Mid$(strPar, 2, Len(strPar) - 2)
                strPar = Split(strPar, ",")(0)
            Else
                strPar = Mid$(strPar, 2, Len(strPar) - 2)
            End If
            
            '判断是否可选参数
            RaiseEvent OnGetParameterValue(strPar, Value)
            
            If (IsEmpty(Value) Or IsNull(Value)) And blnIsResetWhere Then
                '如果为NULL则获取默认条件方式
                strTmp = mobjSqlStruct.AllParameter(i)
                strTmp = Mid$(strTmp, 2, Len(strTmp) - 2)
                arrTmp = Split(strTmp, ",")
                
                If UBound(arrTmp) > 0 Then
                    Value = arrTmp(1)
                End If
            End If
        End If
        
        mvarParValues(i) = Value
        
    Next i
End Sub

Private Function RestoreSqlPar(ByVal strSql As String) As String
'恢复sql中的参数信息
    Dim i As Long
    Dim strResult As String
    
    strResult = strSql
    
    '恢复系统参数名称
    For i = 1 To mobjSqlStruct.ParCount
        strResult = Replace(strResult, "[" & i & "]", mobjSqlStruct.AllParameter(i))
    Next i

    RestoreSqlPar = strResult
End Function



Private Function FormatSelectFromItem(ByRef objCurItem As clsSqlItem) As Boolean
'格式化当前item录入
'如Select [1] || (select Name from t1 where id=[2] and num=[3]) as title1 from dual
'此时如果[1]没有录入，则整个item项表示没有参数录入
    Dim i As Long
    Dim lngCount As Long
    Dim lngCurParNo As Long
    Dim strNewItem As String
    Dim blnIsNullPar As Boolean
    Dim strSourceItem As String
    Dim objSubSqlStrcut() As clsSqlStruct
    Dim strPar As String
    Dim strTmp As String
    
    FormatSelectFromItem = True

'    方式1
'    ReDim objSubSqlStrcut(0)
'
'    strNewItem = GetSelectFromNewItem(objCurItem, objSubSqlStrcut, blnIsNullPar)
'    strSourceItem = objCurItem.Context
'
'    If blnIsNullPar Then
'        For i = 1 To objCurItem.ParsCount
'            lngCurParNo = objCurItem.ParNo(i)
'            strSourceItem = Replace(strSourceItem, "[" & lngCurParNo & "]", mobjSqlStruct.AllParameter(lngCurParNo))
'        Next i
'
'        strNewItem = RestoreBracketContext(strNewItem, mobjSqlStruct.Bracket, True, True)
'        strSourceItem = RestoreBracketContext(strSourceItem, mobjSqlStruct.Bracket, True, True)
'
'        strNewItem = Mid$(strNewItem, 2, Len(strNewItem) - 2) '删除item的前后逗号,如",table(xxx) as f1,"
'        strSourceItem = Mid$(strSourceItem, 2, Len(strSourceItem) - 2)
'
'        mstrSqlProcessing = Replace(mstrSqlProcessing, strSourceItem, strNewItem, 1, 1)
'    End If
'
'    '处理item中包含的sql子查询
'    lngCount = UBound(objSubSqlStrcut)
'    For i = 1 To lngCount
'        FormatSelectFromItem = AnalyzeSqls(objSubSqlStrcut(i))
'        If FormatSelectFromItem = False Then Exit Function
'    Next i



'   方式2
    '判断子sql中是否录入参数
    lngCount = objCurItem.SubSqlsCount
    For i = 1 To lngCount
        FormatSelectFromItem = AnalyzeSqls(objCurItem.SubSqls(i))
        If FormatSelectFromItem = False Then Exit Function
    Next i


    '判断子项目中是否录入参数
    lngCount = objCurItem.SubItemsCount
    For i = 1 To lngCount
        FormatSelectFromItem = FormatSelectFromItem(objCurItem.SubItems(i))
        If FormatSelectFromItem = False Then Exit Function
    Next i


    '判断当前item中的参数录入
    lngCount = objCurItem.ParsCount
    For i = 1 To lngCount
        lngCurParNo = objCurItem.ParNo(i)
        
        If mobjSqlStruct.ParCount >= lngCurParNo Then
            strPar = mobjSqlStruct.AllParameter(lngCurParNo)
            
            If IsEmpty(mvarParValues(lngCurParNo)) Or IsNull(mvarParValues(lngCurParNo)) Or Len(mvarParValues(lngCurParNo)) <= 0 Then
'                mstrSqlProcessing = Replace(mstrSqlProcessing, mobjSqlStruct.AllParameter(lngCurParNo), " NULL ")
                mstrSqlProcessing = Replace(mstrSqlProcessing, strPar, " NULL ", 1, 1)
            Else
                '包含@符号说明是可选参数
                If InStr(strPar, "[@") > 0 Then
                    If mvarParValues(lngCurParNo) = True Then
                        strTmp = Mid$(strPar, 2, Len(strPar) - 2)
                        strTmp = Split(strTmp, ",")(1)
                        
                        mstrSqlProcessing = Replace(mstrSqlProcessing, strPar, strTmp, 1, 1)
                    Else
                        mstrSqlProcessing = Replace(mstrSqlProcessing, strPar, " NULL ", 1, 1)
                    End If
                End If
                
'                '包含*符号说明是可条件方式选择，如果为NULL，则使用默认条件
'                If InStr(strPar, "[*") > 0 Then
'                    strTmp = Mid$(strPar, 2, Len(strPar) - 2)
'                    strTmp = Split(strTmp, ",")(1)
'
'                    mstrSqlProcessing = Replace(mstrSqlProcessing, strPar, strTmp, 1, 1)
'                End If
            End If
        End If
    Next i


End Function


Private Function GetSelectFromNewItem(ByRef objRootItem As clsSqlItem, ByRef objSubSqlStructs() As clsSqlStruct, ByRef blnHasNullPar As Boolean) As String
 
    Dim strTmp As String
    Dim strResult As String
    Dim i As Long
    Dim lngCurParNo As Long
    Dim blnCurNullPar As Boolean

    GetSelectFromNewItem = ""
    blnHasNullPar = False

    strResult = objRootItem.Context
     
    If mobjSqlStruct.ParCount > 0 Then
        For i = 1 To objRootItem.ParsCount
            lngCurParNo = objRootItem.ParNo(i)
            
            If IsEmpty(mvarParValues(lngCurParNo)) Or IsNull(mvarParValues(lngCurParNo)) Or Len(mvarParValues(lngCurParNo)) <= 0 Then
                strResult = Replace(strResult, "[" & lngCurParNo & "]", " NULL ")
                blnHasNullPar = True
            End If
        Next i
    End If
    
    '保存item中包含的子查询
    For i = 1 To objRootItem.SubSqlsCount
        ReDim Preserve objSubSqlStructs(UBound(objSubSqlStructs) + 1)
        Set objSubSqlStructs(UBound(objSubSqlStructs)) = objRootItem.SubSqls(i)
    Next i
    
    '处理子item项
    For i = 1 To objRootItem.SubItemsCount
        strTmp = GetSelectFromNewItem(objRootItem.SubItems(i), objSubSqlStructs, blnCurNullPar)
        
        strResult = Replace(strResult, objRootItem.SubItems(i).BracketTag, strTmp)
        
        If blnCurNullPar Then blnHasNullPar = True
    Next i
    
    '返回新的内容
    GetSelectFromNewItem = strResult

End Function


Private Function GetWhereSource(ByRef objCurItem As clsSqlItem, ByRef strNewWhere As String, ByRef blnIsNullPar As Boolean) As String
'获取where原始条件
    Dim objParent As clsSqlItem
    Dim strTmp As String
    Dim aryParNo() As Long
    Dim strResult As String
    Dim i As Long
    Dim lngCount As Long
    Dim lngCurParNo As Long
    Dim strProcParRecord As String
    Dim strPar As String
    
    
    Set objParent = objCurItem
    
    ReDim aryParNo(0)
    
    strResult = ""
    strTmp = ""
    strNewWhere = ""
    
    Do While Not objParent Is Nothing
        
        '判断是否存在or或者and
        If objParent.HasLinkWhere = True Then Exit Do
        strTmp = objParent.Context
        
        Call objParent.CopyParNoTo(aryParNo)
        
        Set objParent = objParent.Parent
    Loop
    
    If Len(strTmp) <= 0 Then
        strTmp = objCurItem.Context
        Call objParent.CopyParNoTo(aryParNo)
    End If
 
    '恢复Item
    strResult = RestoreBracketContext(strTmp, mobjSqlStruct.Bracket, True, True)
    strNewWhere = strResult
    
    blnIsNullPar = False
    strProcParRecord = ""
    lngCount = UBound(aryParNo)
    
    For i = 1 To lngCount
        lngCurParNo = aryParNo(i)
        
        If lngCurParNo <= mobjSqlStruct.ParCount Then
            strTmp = "[" & lngCurParNo & "]"
            
            '判断给参数是否已经被替换
            If InStr(strProcParRecord, strTmp) <= 0 Then
                strProcParRecord = strProcParRecord & strTmp
                strPar = mobjSqlStruct.AllParameter(lngCurParNo)
                
                strResult = Replace(strResult, strTmp, strPar)
                If blnIsNullPar = False Then strNewWhere = Replace(strNewWhere, strTmp, strPar)
                
                If IsEmpty(mvarParValues(lngCurParNo)) Or IsNull(mvarParValues(lngCurParNo)) Or Len(mvarParValues(lngCurParNo)) <= 0 Then
                    '当item中包含参数时，才进行替换
                    blnIsNullPar = True
                    strNewWhere = " <#TMP> "
                Else
                    '判断是否可选参数
                    If blnIsNullPar = False Then
                        If InStr(strPar, "[@") > 0 Then
                            If mvarParValues(lngCurParNo) = True Then
                                strTmp = Mid$(strPar, 2, Len(strPar) - 2)
                                strTmp = Split(strTmp, ",")(1)
                                strNewWhere = Replace(strNewWhere, strPar, strTmp)
                            Else
                                blnIsNullPar = True
                                strNewWhere = " <#TMP> "
                            End If
                        End If
                        
                    End If
                End If
            End If
        End If
    Next i
    
    GetWhereSource = strResult
End Function

Private Function FormatWhereItem(ByRef objCurItem As clsSqlItem, ByRef blnIsInput As Boolean) As Boolean
'如果item中，有一个参数没有录入，则将该item替换为<#TMP>
'如a between [1] and [2]
'如([1] + [2] = (decode([3], 0, (select f1 from t1), 1, (select f2 from t2), 'x')) and (...))
    Dim i As Long
    Dim lngCount As Long
    Dim lngCurParNo As Long
    Dim strTmp As String
    Dim strNewWhere As String
    Dim blnIsNullPar As Boolean
    
    
    FormatWhereItem = True
    blnIsInput = True
    
    '先判断当前where item中包含的参数
    If objCurItem.HasLinkWhere = False Then
'        lngCount = objCurItem.ParsCount
        
        '当item中包含参数，且参数没有被录入，blnIsReplace将返回true
        strTmp = GetWhereSource(objCurItem, strNewWhere, blnIsNullPar)
        blnIsInput = Not blnIsNullPar
        
        If Len(strNewWhere) > 0 Then
'            mstrSqlProcessing = Replace(mstrSqlProcessing, strTmp, " <#TMP> ", 1, 1)
            mstrSqlProcessing = Replace(mstrSqlProcessing, strTmp, strNewWhere, 1, 1)
        End If
    End If
    
    '判断子item项目
    lngCount = objCurItem.SubItemsCount
    For i = 1 To lngCount
        FormatWhereItem = FormatWhereItem(objCurItem.SubItems(i), blnIsInput)
        If FormatWhereItem = False Then Exit Function
    Next i

    '判断子sql项目
    lngCount = objCurItem.SubSqlsCount
    For i = 1 To lngCount
        FormatWhereItem = AnalyzeSqls(objCurItem.SubSqls(i))
        If FormatWhereItem = False Then Exit Function
    Next i
End Function

Private Function AnalyzeSqls(objCurSqlStruct As clsSqlStruct) As Boolean
    Dim blnParIsInput As Boolean
    Dim i As Long
    Dim lngCount As Long
    Dim objCurItem As clsSqlItem
    Dim lngParItemCount As Long
    Dim lngParInputCount As Long
    Dim strTmp As String
    Dim lngHintResult As Long
    
    AnalyzeSqls = True
    
    'with部分处理
    lngCount = objCurSqlStruct.WithCount
    For i = 1 To lngCount
        AnalyzeSqls = AnalyzeSqls(objCurSqlStruct.Withs(i))
        If AnalyzeSqls = False Then Exit Function
    Next i

    
    'union部分处理
    lngCount = objCurSqlStruct.UnionCount
    For i = 1 To lngCount
        AnalyzeSqls = AnalyzeSqls(objCurSqlStruct.Union(i))
        If AnalyzeSqls = False Then Exit Function
    Next i
    
    
    'minus部分处理
    lngCount = objCurSqlStruct.MinusCount
    For i = 1 To lngCount
        AnalyzeSqls = AnalyzeSqls(objCurSqlStruct.Minus(i))
        If AnalyzeSqls = False Then Exit Function
    Next i
    
    'select ... from ... where .. other 部分处理
    
    'select 部分处理，如果参数没有录入，则直接使用null替换
    lngCount = objCurSqlStruct.SelectCount
    For i = 1 To lngCount
        Set objCurItem = objCurSqlStruct.Selects(i)
        
        AnalyzeSqls = FormatSelectFromItem(objCurItem)
        If AnalyzeSqls = False Then Exit Function
    Next i
    
    'from 部分处理
    lngCount = objCurSqlStruct.FromCount
    For i = 1 To lngCount
        Set objCurItem = objCurSqlStruct.From(i)
        
        AnalyzeSqls = FormatSelectFromItem(objCurItem)
        If AnalyzeSqls = False Then Exit Function
    Next i
    
    'where 部分处理
    lngParItemCount = 0
    lngParInputCount = 0
    
    lngCount = objCurSqlStruct.WhereCount
    For i = 1 To lngCount
        Set objCurItem = objCurSqlStruct.Where(i)
        AnalyzeSqls = FormatWhereItem(objCurItem, blnParIsInput)
        If AnalyzeSqls = False Then Exit Function
        
        If mblnNullParameterHint Then
            If objCurItem.HasPar Or objCurItem.HasSubPar Then lngParItemCount = lngParItemCount + 1
            If blnParIsInput Then lngParInputCount = lngParInputCount + 1
        End If
    Next i
    
    If objCurSqlStruct.FromType = -1 Then
        '如果是union 或者minus则不进行条件录入数量检查
        If objCurSqlStruct.UnionCount > 0 Or objCurSqlStruct.MinusCount > 0 Then
            lngParItemCount = 0
            lngCount = 1
        End If
    End If
    
    
    If (lngParItemCount > 0 And lngParInputCount = 0) Or (lngCount <= 0 And mblnNullWhereHint) Then
        '没有录入任何参数
        If Len(objCurSqlStruct.BracketTag) > 0 Then
            strTmp = objCurSqlStruct.BracketTag
            strTmp = RestoreBracketContext(strTmp, mobjSqlStruct.Bracket, True, True)
        Else
            strTmp = objCurSqlStruct.Sql
            strTmp = RestoreBracketContext(strTmp, mobjSqlStruct.Bracket, True, True)
            
            strTmp = RestoreSqlPar(strTmp)
        End If
        
        If lngCount <= 0 Then
            lngHintResult = MsgBox("检测到查询中没有条件语句，是否继续？" & vbCrLf & "    查询语句如下:" & vbCrLf & "    " & strTmp, vbYesNo)
        Else
            lngHintResult = MsgBox("检测到查询中未录入任何参数，是否继续？" & vbCrLf & "    查询语句如下:" & vbCrLf & "    " & strTmp, vbYesNo)
        End If
        
        If lngHintResult = vbNo Then
            AnalyzeSqls = False
            Exit Function
        End If
    End If
   
    
    'func 部分处理
    lngCount = objCurSqlStruct.FuncCount
    For i = 1 To lngCount
        Set objCurItem = objCurSqlStruct.Func(i)
        
        AnalyzeSqls = FormatSelectFromItem(objCurItem)
        If AnalyzeSqls = False Then Exit Function
    Next i
    
End Function


Private Function Infer(ByRef strSource As String, _
                        ByVal blnIsImmediate As Boolean, _
                        ByVal strInfer As String, _
                        ByVal strTag As String, _
                        ByVal strNew As String) As Boolean
'进行条件推断

    Dim lngIndex As Long
    
    Infer = blnIsImmediate
    
    If blnIsImmediate Then
        strSource = Replace(strSource, strInfer, strNew)
    Else
        lngIndex = InStr(1, strSource, strInfer)
        If lngIndex > 0 Then
            strSource = Replace(strSource, strInfer, strNew)
            Infer = True
        End If
    End If
    
End Function

'格式化参数为空的条件
Private Function FormatNull(ByVal strSql As String) As String
'包含空的条件=> [@EMP]
'多个[@EMP] =>单个[@EMP]
    
    Dim blnIsReplace As Boolean
    
    If Trim(strSql) = "" Then Exit Function
    
    FormatNull = "(" & strSql & ")"
        
    blnIsReplace = True
    
    While blnIsReplace
        blnIsReplace = False
        
        While InStr(FormatNull, "  ") > 0
            FormatNull = Replace(FormatNull, "  ", " ")
        Wend
        
        FormatNull = Replace(FormatNull, "( ", "(")
        FormatNull = Replace(FormatNull, " )", ")")
        
        FormatNull = Replace(FormatNull, " (", "(")
        FormatNull = Replace(FormatNull, ") ", ")")
        
        While InStr(FormatNull, "<#TMP><#TMP>") > 0
            FormatNull = Replace(FormatNull, "<#TMP><#TMP>", "<#TMP>")
        Wend
        
        While InStr(FormatNull, "<#TMP> <#TMP>") > 0
            FormatNull = Replace(FormatNull, "<#TMP> <#TMP>", "<#TMP>")
        Wend
    
        
        blnIsReplace = Infer(FormatNull, blnIsReplace, "AND <#TMP>", "=>", "AND")
        blnIsReplace = Infer(FormatNull, blnIsReplace, "OR <#TMP>", "=>", "OR")
        
        blnIsReplace = Infer(FormatNull, blnIsReplace, "<#TMP> ORDER", "=>", "ORDER")
        blnIsReplace = Infer(FormatNull, blnIsReplace, "<#TMP> AND", "=>", "AND")
        blnIsReplace = Infer(FormatNull, blnIsReplace, "<#TMP> OR", "=>", "OR")
        
        blnIsReplace = Infer(FormatNull, blnIsReplace, "AND AND", "=>", "AND")
        blnIsReplace = Infer(FormatNull, blnIsReplace, "AND OR", "=>", "OR")
        blnIsReplace = Infer(FormatNull, blnIsReplace, "OR AND", "=>", "OR")
        blnIsReplace = Infer(FormatNull, blnIsReplace, "OR OR", "=>", "OR")
    
        blnIsReplace = Infer(FormatNull, blnIsReplace, "(AND", "=>", "(")
        blnIsReplace = Infer(FormatNull, blnIsReplace, "(OR", "=>", "(")
        
        blnIsReplace = Infer(FormatNull, blnIsReplace, "AND)", "=>", ")")
        blnIsReplace = Infer(FormatNull, blnIsReplace, "OR)", "=>", ")")
        
        blnIsReplace = Infer(FormatNull, blnIsReplace, "(AND)", "=>", "<#TMP>")
        blnIsReplace = Infer(FormatNull, blnIsReplace, "(OR)", "=>", "<#TMP>")
        
        blnIsReplace = Infer(FormatNull, blnIsReplace, "(<#TMP>", "=>", "(")
        blnIsReplace = Infer(FormatNull, blnIsReplace, "<#TMP>)", "=>", ")")
        
        blnIsReplace = Infer(FormatNull, blnIsReplace, "(<#TMP>)", "=>", "<#TMP>")
        
        blnIsReplace = Infer(FormatNull, blnIsReplace, "<#TMP>(", "=>", "(")
        blnIsReplace = Infer(FormatNull, blnIsReplace, ")<#TMP>", "=>", ")")

        blnIsReplace = Infer(FormatNull, blnIsReplace, "()", "=>", "<#TMP>")
        
        blnIsReplace = Infer(FormatNull, blnIsReplace, "WHERE ORDER", "=>", "ORDER")
        blnIsReplace = Infer(FormatNull, blnIsReplace, "WHERE OR", "=>", "WHERE")
        blnIsReplace = Infer(FormatNull, blnIsReplace, "WHERE AND", "=>", "WHERE")
        
        blnIsReplace = Infer(FormatNull, blnIsReplace, "WHERE)", "=>", ")")
    Wend
    
    If InStr(FormatNull, "UNION") > 0 Then
        Call Infer(FormatNull, True, "WHERE UNION", "=>", "UNION")
        Call Infer(FormatNull, True, "OR UNION", "=>", "UNION")
        Call Infer(FormatNull, True, "AND UNION", "=>", "UNION")
    End If
    
    If InStr(FormatNull, "MINUS") > 0 Then
        Call Infer(FormatNull, True, "WHERE MINUS", "=>", "MINUS")
        Call Infer(FormatNull, True, "OR MINUS", "=>", "MINUS")
        Call Infer(FormatNull, True, "AND MINUS", "=>", "MINUS")
    End If
    
    'Where [@EMP] >= 空
    If InStr(FormatNull, "<#TMP>") > 0 Then
        '需要判断where <#Tmp> 之后是否存在右括弧
        Call Infer(FormatNull, True, "WHERE <#TMP>", "=>", "")
        Call Infer(FormatNull, True, "<#TMP>", "=>", "")
    End If
    
    FormatNull = Mid$(FormatNull, 2, Len(FormatNull) - 2)
End Function

Private Sub Class_Initialize()
    IsInit = False
End Sub

Private Sub Class_Terminate()
    Set mobjSqlStruct = Nothing
End Sub
