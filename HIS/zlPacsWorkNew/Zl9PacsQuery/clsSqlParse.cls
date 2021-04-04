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
 

Private mstrSourceSql As String         'ԭʼ��ѯ
Private mstrFormatSql As String         '��ʽ����Ĳ�ѯ��ȥ���˻س���˫���ţ���ʼ����
Private mstrQuerySql As String          '�Բ������д�����Ĳ�ѯ���������滻����[1] [2]����ʽ

Private mvarParValues() As Variant      '����ֵ

Private mstrQuotes() As String
Private mobjSqlStruct As New clsSqlStruct

Private mblnNullParameterHint As Boolean
Private mblnNullWhereHint As Boolean


Private mstrSqlProcessing As String


Private mstrAttachItem As String
Private mstrAttachData As Variant
Private mblnIsParSymbolTransfer As Boolean
Private mlngResetWhereIndex() As Long       '��Ҫ���ü���������������������ʹ���ˡ�*���ŵģ���ʾ��ѯ�ļ���������ʽ���Ը���ѡ���������

Public IsInit As Boolean

 

'����ֵ��ȡ�¼�,
'strParName:��������
'lngParType:�������� 0-ϵͳ������1-��������
'value:��������ֵ
Public Event OnGetParameterValue(ByVal strParName As String, ByRef Value As Variant)

'***********************************************************************************************
'***********************************************************************************************

'ԭʼSQL
Property Get Sql() As String
    Sql = mstrSourceSql
End Property

'����ֵ
Property Get ParValue(ByVal lngIndex As Long) As Variant
    ParValue = mvarParValues(lngIndex)
End Property

Property Get ParValues() As Variant()
    ParValues = mvarParValues
End Property

'sql�ṹ����
Property Get SqlStruct() As clsSqlStruct
    Set SqlStruct = mobjSqlStruct
End Property

'�ղ�������
Property Get NullParameterHint() As Boolean
    NullParameterHint = mblnNullParameterHint
End Property

Property Let NullParameterHint(Value As Boolean)
    mblnNullParameterHint = Value
End Property

'����������
Property Get NullWhereHint() As Boolean
    NullWhereHint = mblnNullWhereHint
End Property

Property Let NullWhereHint(Value As Boolean)
    mblnNullWhereHint = Value
End Property


'�Ƿ���Բ�������"[]"ת��
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
    
    '����sql���
    'mstrFormatSql����load�󣬲����Ѿ���[1],[2],...����ʽ�滻
    Call mobjSqlStruct.Load(mstrFormatSql)
    
    IsInit = True
End Sub


Public Function GetTestSql(Optional ByVal strSql As String = "") As String
'��ȡ���Ե�SQL���
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
        '���Ե�ǰ��ʼ�����õ�Sql
        lngCount = mobjSqlStruct.ParCount
       
        For i = 1 To lngCount
            strPar = mobjSqlStruct.AllParameter(i)
            
            '�ж��Ƿ��ѡ����
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
        
        '���ָ�˫����֮�������
        strTestSql = RestoreQuote(strTestSql, mstrQuotes, blnIsParSymbolTransfer)
        
        'Ϊ��ʱ��ʾ�Ե������к��еġ�[]�����Ž�����ת�壬�����ʵ�ʲ�ѯʱ����Ҫ�滻����
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
'��ȡҽ��
    mstrAttachItem = strAttachItem
    mstrAttachData = strAttachData
    
    GetAttachSql = RecombineSql
End Function


'***********************************************************************************************************************************


Private Function RecombineSql(Optional ByVal blnIsRefresh As Boolean = False)
'����sql���
    Dim i As Long
    Dim lngParIndex As Long
    Dim objCurSqlStruct As clsSqlStruct
      
    mstrSqlProcessing = mstrFormatSql ' mstrQuerySql
    mblnIsParSymbolTransfer = False
    
    '��ȡ����
    '�����ˢ�£�����Ҫ�ظ���ȡ����ֵ
    If Not blnIsRefresh Then
        Call ReadParameterValue
    End If
    
    '��������Ϊnull������
    If AnalyzeSqls(mobjSqlStruct) = False Then
        RecombineSql = ""
        Exit Function
    End If
    
    '����ʵ�ֲ�ѯ�������жϷ�ʽ��
    For i = 1 To UBound(mlngResetWhereIndex)
        lngParIndex = mlngResetWhereIndex(i)
        mstrSqlProcessing = Replace(mstrSqlProcessing, mobjSqlStruct.AllParameter(lngParIndex), mvarParValues(lngParIndex))
    Next
    
    
    '��ʽ��null����������sql���
    mstrSqlProcessing = FormatNull(mstrSqlProcessing)
    
    '�ָ�˫����֮�������
    mstrSqlProcessing = RestoreQuote(mstrSqlProcessing, mstrQuotes, mblnIsParSymbolTransfer)
    
    '�ָ������еĲ�������
    mstrSqlProcessing = RestoreParIndex(mstrSqlProcessing, mobjSqlStruct.AllParameterAry())
        
    RecombineSql = mstrSqlProcessing
End Function


Private Sub ReadParameterValue()
'��ȡ����ֵ
    Dim i As Long
    Dim Value As Variant
    Dim lngUbound As Long
    Dim strPar As String
    Dim blnIsNeedRange As Boolean
    Dim lngResetUbound As Long
    Dim blnIsResetWhere As Boolean
    Dim strTmp As String
    Dim arrTmp() As String
    
    '��ȡϵͳ����
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
            
            '�ж��Ƿ��ѡ����
            RaiseEvent OnGetParameterValue(strPar, Value)
            
            If (IsEmpty(Value) Or IsNull(Value)) And blnIsResetWhere Then
                '���ΪNULL���ȡĬ��������ʽ
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
'�ָ�sql�еĲ�����Ϣ
    Dim i As Long
    Dim strResult As String
    
    strResult = strSql
    
    '�ָ�ϵͳ��������
    For i = 1 To mobjSqlStruct.ParCount
        strResult = Replace(strResult, "[" & i & "]", mobjSqlStruct.AllParameter(i))
    Next i

    RestoreSqlPar = strResult
End Function



Private Function FormatSelectFromItem(ByRef objCurItem As clsSqlItem) As Boolean
'��ʽ����ǰitem¼��
'��Select [1] || (select Name from t1 where id=[2] and num=[3]) as title1 from dual
'��ʱ���[1]û��¼�룬������item���ʾû�в���¼��
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

'    ��ʽ1
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
'        strNewItem = Mid$(strNewItem, 2, Len(strNewItem) - 2) 'ɾ��item��ǰ�󶺺�,��",table(xxx) as f1,"
'        strSourceItem = Mid$(strSourceItem, 2, Len(strSourceItem) - 2)
'
'        mstrSqlProcessing = Replace(mstrSqlProcessing, strSourceItem, strNewItem, 1, 1)
'    End If
'
'    '����item�а�����sql�Ӳ�ѯ
'    lngCount = UBound(objSubSqlStrcut)
'    For i = 1 To lngCount
'        FormatSelectFromItem = AnalyzeSqls(objSubSqlStrcut(i))
'        If FormatSelectFromItem = False Then Exit Function
'    Next i



'   ��ʽ2
    '�ж���sql���Ƿ�¼�����
    lngCount = objCurItem.SubSqlsCount
    For i = 1 To lngCount
        FormatSelectFromItem = AnalyzeSqls(objCurItem.SubSqls(i))
        If FormatSelectFromItem = False Then Exit Function
    Next i


    '�ж�����Ŀ���Ƿ�¼�����
    lngCount = objCurItem.SubItemsCount
    For i = 1 To lngCount
        FormatSelectFromItem = FormatSelectFromItem(objCurItem.SubItems(i))
        If FormatSelectFromItem = False Then Exit Function
    Next i


    '�жϵ�ǰitem�еĲ���¼��
    lngCount = objCurItem.ParsCount
    For i = 1 To lngCount
        lngCurParNo = objCurItem.ParNo(i)
        
        If mobjSqlStruct.ParCount >= lngCurParNo Then
            strPar = mobjSqlStruct.AllParameter(lngCurParNo)
            
            If IsEmpty(mvarParValues(lngCurParNo)) Or IsNull(mvarParValues(lngCurParNo)) Or Len(mvarParValues(lngCurParNo)) <= 0 Then
'                mstrSqlProcessing = Replace(mstrSqlProcessing, mobjSqlStruct.AllParameter(lngCurParNo), " NULL ")
                mstrSqlProcessing = Replace(mstrSqlProcessing, strPar, " NULL ", 1, 1)
            Else
                '����@����˵���ǿ�ѡ����
                If InStr(strPar, "[@") > 0 Then
                    If mvarParValues(lngCurParNo) = True Then
                        strTmp = Mid$(strPar, 2, Len(strPar) - 2)
                        strTmp = Split(strTmp, ",")(1)
                        
                        mstrSqlProcessing = Replace(mstrSqlProcessing, strPar, strTmp, 1, 1)
                    Else
                        mstrSqlProcessing = Replace(mstrSqlProcessing, strPar, " NULL ", 1, 1)
                    End If
                End If
                
'                '����*����˵���ǿ�������ʽѡ�����ΪNULL����ʹ��Ĭ������
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
    
    '����item�а������Ӳ�ѯ
    For i = 1 To objRootItem.SubSqlsCount
        ReDim Preserve objSubSqlStructs(UBound(objSubSqlStructs) + 1)
        Set objSubSqlStructs(UBound(objSubSqlStructs)) = objRootItem.SubSqls(i)
    Next i
    
    '������item��
    For i = 1 To objRootItem.SubItemsCount
        strTmp = GetSelectFromNewItem(objRootItem.SubItems(i), objSubSqlStructs, blnCurNullPar)
        
        strResult = Replace(strResult, objRootItem.SubItems(i).BracketTag, strTmp)
        
        If blnCurNullPar Then blnHasNullPar = True
    Next i
    
    '�����µ�����
    GetSelectFromNewItem = strResult

End Function


Private Function GetWhereSource(ByRef objCurItem As clsSqlItem, ByRef strNewWhere As String, ByRef blnIsNullPar As Boolean) As String
'��ȡwhereԭʼ����
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
        
        '�ж��Ƿ����or����and
        If objParent.HasLinkWhere = True Then Exit Do
        strTmp = objParent.Context
        
        Call objParent.CopyParNoTo(aryParNo)
        
        Set objParent = objParent.Parent
    Loop
    
    If Len(strTmp) <= 0 Then
        strTmp = objCurItem.Context
        Call objParent.CopyParNoTo(aryParNo)
    End If
 
    '�ָ�Item
    strResult = RestoreBracketContext(strTmp, mobjSqlStruct.Bracket, True, True)
    strNewWhere = strResult
    
    blnIsNullPar = False
    strProcParRecord = ""
    lngCount = UBound(aryParNo)
    
    For i = 1 To lngCount
        lngCurParNo = aryParNo(i)
        
        If lngCurParNo <= mobjSqlStruct.ParCount Then
            strTmp = "[" & lngCurParNo & "]"
            
            '�жϸ������Ƿ��Ѿ����滻
            If InStr(strProcParRecord, strTmp) <= 0 Then
                strProcParRecord = strProcParRecord & strTmp
                strPar = mobjSqlStruct.AllParameter(lngCurParNo)
                
                strResult = Replace(strResult, strTmp, strPar)
                If blnIsNullPar = False Then strNewWhere = Replace(strNewWhere, strTmp, strPar)
                
                If IsEmpty(mvarParValues(lngCurParNo)) Or IsNull(mvarParValues(lngCurParNo)) Or Len(mvarParValues(lngCurParNo)) <= 0 Then
                    '��item�а�������ʱ���Ž����滻
                    blnIsNullPar = True
                    strNewWhere = " <#TMP> "
                Else
                    '�ж��Ƿ��ѡ����
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
'���item�У���һ������û��¼�룬�򽫸�item�滻Ϊ<#TMP>
'��a between [1] and [2]
'��([1] + [2] = (decode([3], 0, (select f1 from t1), 1, (select f2 from t2), 'x')) and (...))
    Dim i As Long
    Dim lngCount As Long
    Dim lngCurParNo As Long
    Dim strTmp As String
    Dim strNewWhere As String
    Dim blnIsNullPar As Boolean
    
    
    FormatWhereItem = True
    blnIsInput = True
    
    '���жϵ�ǰwhere item�а����Ĳ���
    If objCurItem.HasLinkWhere = False Then
'        lngCount = objCurItem.ParsCount
        
        '��item�а����������Ҳ���û�б�¼�룬blnIsReplace������true
        strTmp = GetWhereSource(objCurItem, strNewWhere, blnIsNullPar)
        blnIsInput = Not blnIsNullPar
        
        If Len(strNewWhere) > 0 Then
'            mstrSqlProcessing = Replace(mstrSqlProcessing, strTmp, " <#TMP> ", 1, 1)
            mstrSqlProcessing = Replace(mstrSqlProcessing, strTmp, strNewWhere, 1, 1)
        End If
    End If
    
    '�ж���item��Ŀ
    lngCount = objCurItem.SubItemsCount
    For i = 1 To lngCount
        FormatWhereItem = FormatWhereItem(objCurItem.SubItems(i), blnIsInput)
        If FormatWhereItem = False Then Exit Function
    Next i

    '�ж���sql��Ŀ
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
    
    'with���ִ���
    lngCount = objCurSqlStruct.WithCount
    For i = 1 To lngCount
        AnalyzeSqls = AnalyzeSqls(objCurSqlStruct.Withs(i))
        If AnalyzeSqls = False Then Exit Function
    Next i

    
    'union���ִ���
    lngCount = objCurSqlStruct.UnionCount
    For i = 1 To lngCount
        AnalyzeSqls = AnalyzeSqls(objCurSqlStruct.Union(i))
        If AnalyzeSqls = False Then Exit Function
    Next i
    
    
    'minus���ִ���
    lngCount = objCurSqlStruct.MinusCount
    For i = 1 To lngCount
        AnalyzeSqls = AnalyzeSqls(objCurSqlStruct.Minus(i))
        If AnalyzeSqls = False Then Exit Function
    Next i
    
    'select ... from ... where .. other ���ִ���
    
    'select ���ִ������������û��¼�룬��ֱ��ʹ��null�滻
    lngCount = objCurSqlStruct.SelectCount
    For i = 1 To lngCount
        Set objCurItem = objCurSqlStruct.Selects(i)
        
        AnalyzeSqls = FormatSelectFromItem(objCurItem)
        If AnalyzeSqls = False Then Exit Function
    Next i
    
    'from ���ִ���
    lngCount = objCurSqlStruct.FromCount
    For i = 1 To lngCount
        Set objCurItem = objCurSqlStruct.From(i)
        
        AnalyzeSqls = FormatSelectFromItem(objCurItem)
        If AnalyzeSqls = False Then Exit Function
    Next i
    
    'where ���ִ���
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
        '�����union ����minus�򲻽�������¼���������
        If objCurSqlStruct.UnionCount > 0 Or objCurSqlStruct.MinusCount > 0 Then
            lngParItemCount = 0
            lngCount = 1
        End If
    End If
    
    
    If (lngParItemCount > 0 And lngParInputCount = 0) Or (lngCount <= 0 And mblnNullWhereHint) Then
        'û��¼���κβ���
        If Len(objCurSqlStruct.BracketTag) > 0 Then
            strTmp = objCurSqlStruct.BracketTag
            strTmp = RestoreBracketContext(strTmp, mobjSqlStruct.Bracket, True, True)
        Else
            strTmp = objCurSqlStruct.Sql
            strTmp = RestoreBracketContext(strTmp, mobjSqlStruct.Bracket, True, True)
            
            strTmp = RestoreSqlPar(strTmp)
        End If
        
        If lngCount <= 0 Then
            lngHintResult = MsgBox("��⵽��ѯ��û��������䣬�Ƿ������" & vbCrLf & "    ��ѯ�������:" & vbCrLf & "    " & strTmp, vbYesNo)
        Else
            lngHintResult = MsgBox("��⵽��ѯ��δ¼���κβ������Ƿ������" & vbCrLf & "    ��ѯ�������:" & vbCrLf & "    " & strTmp, vbYesNo)
        End If
        
        If lngHintResult = vbNo Then
            AnalyzeSqls = False
            Exit Function
        End If
    End If
   
    
    'func ���ִ���
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
'���������ƶ�

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

'��ʽ������Ϊ�յ�����
Private Function FormatNull(ByVal strSql As String) As String
'�����յ�����=> [@EMP]
'���[@EMP] =>����[@EMP]
    
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
    
    'Where [@EMP] >= ��
    If InStr(FormatNull, "<#TMP>") > 0 Then
        '��Ҫ�ж�where <#Tmp> ֮���Ƿ����������
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