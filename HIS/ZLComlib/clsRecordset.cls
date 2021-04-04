VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'##############################################################################
'#功能：记录集类型相关的处理、判断等公共函数
'命名规则：函数过程名称为通俗易懂的英文即可，不用加Rec关键字
'上级类：clsDatabase
'##############################################################################
Public Function InsertCurRow(ByVal rsSource As ADODB.Recordset, ByRef rsNewData As ADODB.Recordset, Optional strPKFields As String = "ID", Optional blnPkIsNumder As Boolean = True) As Boolean
'功能:将当前记录的数据,插入到目标记录集中
'入参:rsSouce-原记录集
'       strPKFields-主键是哪个(如果主键存在相同值，则不能插入.),只能有一个字段
'       blnPkIsNumder-主键类型
'出参:rsNewData-添加数据后的新记录集
 '返回:插入成功,返回true,否则返回False
    InsertCurRow = gobjComLib.zlDatabase.zlInsertCurrRowData(rsSource, rsNewData, strPKFields, blnPkIsNumder)
End Function

Public Function CopyStructure(ByVal rsSource As ADODB.Recordset) As ADODB.Recordset
'功能:    复制记录集的数据结构
'参数： rsSource-原记录集
'返回：具有与原记录集相同结构的本地记录集
    Set CopyStructure = gobjComLib.zlDatabase.zlCopyDataStructure(rsSource)
End Function

Public Function Update(ByRef rsInput As Recordset, ByVal strFilter As String, ParamArray arrInput() As Variant) As Boolean
'功能：更新指定条件的记录集的记录
'参数：rsInput=记录集
'      strFilter=条件
'      arrInput=输入的字段名以及值，格式：字段名1,值1, 字段名2,值2,....
'返回：是否成功
'      rsInput=经过更新后的记录集
'说明：arrInput的字段值可以用记录集中的其他字段来更新该字段，此时格式为：!字段名
    Dim strFiledName As String, strFileValue As String
    Dim blnFiled As Boolean, i As Long

    On Error GoTo errH
    With rsInput
        .Filter = strFilter
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            For i = LBound(arrInput) To UBound(arrInput) Step 2
                strFiledName = arrInput(i)
                If IsNull(arrInput(i + 1)) Then
                    rsInput(strFiledName).value = Null
                Else
                    If arrInput(i + 1) Like "!?*" Then
                        blnFiled = True
                        On Error Resume Next
                        strFileValue = rsInput(Mid(arrInput(i + 1), 2)).value & ""
                        If Err.Number <> 0 Then Err.Clear: blnFiled = False
                        On Error GoTo errH
                    End If
                    If Not blnFiled Then
                        rsInput(strFiledName).value = arrInput(i + 1)
                    Else
                        rsInput(strFiledName).value = rsInput(Mid(arrInput(i + 1), 2)).value
                    End If
                End If
                blnFiled = False
                Call rsInput.Update
            Next
            .MoveNext
        Loop
    End With
    Update = True
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function ToString(ByVal rsObj As ADODB.Recordset) As String
'功能：将记录集转化成字符串，是记录集GetString函数的简单使用，用来调试记录集
    ToString = rsObj.GetString(, , "[Col]", vbNewLine, "NULL")
    If rsObj.RecordCount <> 0 Then rsObj.MoveFirst
End Function

Public Function ToSQL(ByVal rsTemp As ADODB.Recordset, Optional ByVal strHiddenCol As String) As String
'功能：将记录集转成SQL串
'参数:
'    strHiddenCol -不显示的列 列名之间用分号分隔
    Dim strSQL As String
    Dim i As Long
    
    With rsTemp
        strSQL = ""
        Do While Not .EOF
            strSQL = strSQL & " Union ALL Select "
            For i = 0 To .Fields.count - 1
                If InStr(";" & strHiddenCol & ";", ";" & .Fields(i).Name & ";") = 0 Then '屏蔽掉隐藏列
                    If Not IsNull(.Fields(i).value) Then
                        If IsType(.Fields(i).type, adVarChar) Then
                            strSQL = strSQL & "'" & Replace(Replace(.Fields(i).value, "[", "("), "]", ")") & "'"
                        ElseIf IsType(.Fields(i).type, adDBTimeStamp) Then    '日期型
                            strSQL = strSQL & "to_date('" & .Fields(i).value & "','" & "YYYY-MM-DD')"
                        Else
                            strSQL = strSQL & .Fields(i).value '没有日期型
                        End If
                    Else
                        If IsType(.Fields(i).type, adBigInt) Or IsType(.Fields(i).type, adSmallInt) Or IsType(.Fields(i).type, adSingle) Then
                            strSQL = strSQL & "-Null"
                        Else
                            strSQL = strSQL & "Null"
                        End If
                    End If
                    strSQL = strSQL & " As " & .Fields(i).Name & ","
                End If
            Next
            strSQL = Left(strSQL, Len(strSQL) - 1) & " From Dual"
            .MoveNext
        Loop
        .Filter = ""
        strSQL = Mid(strSQL, 12)
    End With
    
    ToSQL = strSQL
End Function

Public Function ToComboList(ByVal rsInput As ADODB.Recordset, ByVal strFormat As String, ParamArray arrInput() As Variant) As String
'功能：将记录集的指定字段，组装成特定格式字符串此种字符串可适用余VsFlexGrid的下拉列表
'参数：rsInput=数据源
'           strFormat=格式字符串，[0]，[1]...等进行占位
'           arrInput=占位符对应的字段名
'返回：返回下拉列表值
    Dim i As Long
    Dim strTmp As String, strReturn As String
    
    Do While Not rsInput.EOF
        strTmp = strFormat
        For i = LBound(arrInput) To UBound(arrInput)
              strTmp = Replace(strTmp, "[" & i & "]", rsInput.Fields(arrInput(i)).value & "")
        Next
        strReturn = strReturn & strTmp
        rsInput.MoveNext
    Loop
    ToComboList = strReturn
End Function

Public Function FilterNew(ByVal rsInput As ADODB.Recordset, Optional ByVal strFilter As String, Optional ByVal strFiledsName As String) As ADODB.Recordset
'功能：返回指定过滤条件的记录集结果，该记录集是新的记录集，与原来的记录集没有关系
'参数：rsInput=目标记录集
'      strFilter=过滤条件，二次过滤条件，传入的记录集已经过滤，在过滤后的过滤条件
'      strFiledsName=需要复制的字段名
'返回：过滤后的记录集的副本
'说明：可能过滤条件太多且复杂，因此需要二次过滤，该函数主要解决这种情况
    Dim rsClone As ADODB.Recordset
    
    On Error GoTo errH
    '没有过滤条件直接复制
    If strFilter = "" Then
        Set FilterNew = CopyNew(rsInput, , strFiledsName)
        Exit Function
    End If
    Set rsClone = CopyNew(rsInput)
    rsClone.Filter = strFilter
    Set FilterNew = CopyNew(rsClone, , strFiledsName)
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function ExitsField(ByVal rsData As ADODB.Recordset, ByVal strFieldName As String) As Boolean
    '******************************************************************************************************************
    '功能：判断记录集是否存在指定名称的字段
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTmp As String
    
    On Error Resume Next
    strTmp = ""
    strTmp = rsData.Fields(strFieldName).Name
    ExitsField = (strTmp = strFieldName)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Public Function Distinct(ByVal rsSource As ADODB.Recordset, Optional ByVal strDisFieldsName As String, Optional ByVal strFieldsName As String) As ADODB.Recordset
'功能：记录集去重复
'参数：rsSource=要去重复的记录集
'strDisFieldsName=去重复的字段,为空，则对所有字段去重
'strFieldsName=返回结果集字段，为空，则返回去重复的字段
'返回：操作后的记录集
    Dim rsReturn As ADODB.Recordset
    Dim arrFilds As Variant, arrValues As Variant
    Dim i As Long, j As Long
    Dim strTmp As String, strOldRow As String
    
    On Error GoTo errH
    '读取默认字段名
    If strDisFieldsName = "" Then
        For i = 0 To rsSource.Fields.count - 1
            strTmp = strTmp & "," & rsSource.Fields(i).Name
        Next
        strTmp = Mid(strTmp, 2)
        If strDisFieldsName = "" Then strDisFieldsName = strTmp
    End If
    If strFieldsName = "" Then strFieldsName = strDisFieldsName
    
    Set rsReturn = CopyNew(rsSource, , strFieldsName)
    If rsSource.RecordCount = 0 Then Set Distinct = rsReturn: Exit Function
    
    rsReturn.Sort = strDisFieldsName '排序，自动将光标移动到开头
    Do While Not rsReturn.EOF
        strTmp = rsReturn.GetString(, 1, "[ColumnSpliter]", , "[NULLEXP]") '自动移动光标
        rsReturn.MovePrevious
        If strTmp = strOldRow Then  '删除重复行
            Call rsReturn.Delete: Call rsReturn.Update
        Else
            strOldRow = strTmp
        End If
        rsReturn.MoveNext
    Loop
    rsReturn.Sort = strDisFieldsName
    Set Distinct = rsReturn
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function Delete(ByRef rsInput As ADODB.Recordset, Optional ByVal strFilter As String) As Boolean
'功能：删除指定条件的记录集的记录
'参数：rsInput=记录集
'      strFilter=条件
'返回：是否成功
'      rsInput=经过删除后的记录集
    On Error GoTo errH
    rsInput.Filter = strFilter
    If rsInput.RecordCount > 0 Then
        rsInput.MoveFirst
        Do While Not rsInput.EOF
            Call rsInput.Delete: Call rsInput.Update
            rsInput.MoveNext
        Loop
    End If
    Delete = True
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function Compare(ByVal rsSouce As ADODB.Recordset, ByVal rsCompare As ADODB.Recordset) As Boolean
'比较两个记录集是否相同
'说明：两记录集结构不同返回False,结构相同数据不同返回False,数据以及结构完全相同返回True
'      结构仅是比较字段名字以及顺序，具体类型不比较（暂时用不到，需要时可以扩展）
    Dim strTmp1 As String, strTmp2 As String
    Dim i As Long
    
    On Error GoTo errH
    If rsSouce.RecordCount <> rsCompare.RecordCount Then Compare = False: Exit Function
    If rsSouce.Fields.count <> rsCompare.Fields.count Then Compare = False: Exit Function
    
    For i = 0 To rsSouce.Fields.count - 1
        strTmp1 = strTmp1 & "," & rsSouce.Fields(i).Name
    Next
    
    For i = 0 To rsCompare.Fields.count - 1
        strTmp2 = strTmp2 & "," & rsCompare.Fields(i).Name
    Next
    
    If strTmp1 <> strTmp2 Then Compare = False: Exit Function
    strTmp1 = Mid(strTmp1, 2)
    rsCompare.Sort = strTmp1: rsSouce.Sort = strTmp1
    strTmp1 = rsCompare.GetString(, , "[ColumnSplit]", "[RowSplit]", "NULL")
    strTmp2 = rsSouce.GetString(, , "[ColumnSplit]", "[RowSplit]", "NULL")
    Compare = (strTmp1 = strTmp2)
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function Append(ByRef rsSource As ADODB.Recordset, ByVal rsAppend As ADODB.Recordset, ParamArray arrInput() As Variant) As Boolean
'功能：将指定记录集的数据添加到另一个记录集上
'参数：rsSource=目标记录集
'      rsAppend=数据记录集
'      arrInput=字段对应规则，该参数不传时，默认两记录集结构相同，格式：arrInput(0):[记录集1].字段1,字段2...；arrInput(1)：[记录集2].字段1,字段2...
'返回：是否成功
'      rsSource=添加数据后的记录集
    Dim arrSource As Variant, arrAppend As Variant
    Dim i As Long, arrValues() As Variant
    Dim strTmp As String
    
    If rsAppend Is Nothing Then Append = True: Exit Function
    If rsAppend.RecordCount = 0 Then Append = True: Exit Function
    If rsSource Is Nothing Then Exit Function
    On Error GoTo errH
    If LBound(arrInput) = 2 Then
        arrSource = Split(arrInput(LBound(arrInput)), ",")
        arrAppend = Split(arrInput(UBound(arrInput)), ",")
        If UBound(arrSource) <> UBound(arrAppend) Then Exit Function
        ReDim arrValues(UBound(arrAppend)): rsAppend.MoveFirst
        Do While Not rsAppend.EOF
            For i = LBound(arrAppend) To UBound(arrAppend)
                arrValues(i) = rsAppend(arrAppend(i)).value
            Next
            rsSource.AddNew arrSource, arrValues
            Erase arrValues
            rsAppend.MoveNext
        Loop
    ElseIf LBound(arrInput) = 0 Then
        Do While Not rsAppend.EOF
            rsSource.AddNew
            For i = 0 To rsSource.Fields.count - 1
                rsSource.Fields(i).value = rsAppend.Fields(i).value
            Next
            rsSource.Update
            rsAppend.MoveNext
        Loop
    End If
    
    Append = True
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Sub AddNew(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    Dim arrFields, arrValues, intField As Integer
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub
    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function CopyNew(ByVal rsSource As ADODB.Recordset, Optional blnOnlyStructure As Boolean, Optional ByVal strFields As String, Optional arrAppFields As Variant) As ADODB.Recordset
'复制记录集
'参数：strFields=需要复制的记录集的字段的列顺序或字段名组成的字符串
'          如：1 别名1,3 别名2,7 别名3...表示复制记录集的第1,3,7..字段组成记录集并返回
'              ID 别名1,姓名 别名2,....表示复制记录集的ID,姓名...字段组成记录集返回
'              别名*为新的记录集的列名
'              两中类型混搭容易出现列名相同的问题，请注意
'           arrAppFields=追加的字段信息：列名,类型,长度,默认值,没有默认值传Empty,没有指定长度传Empty
'      blnOnlyStructure=是否只复制结构
'在程序中，经常会涉及到相互传递记录集，而使用ADO的Clone复制产生的记录集，当其中一个记录集的数据发生变化的时候，所有副本都将发生相同的变化（通常指修改或删除），而我们往往希望这些记录集相互间保持独立
    Set CopyNew = gobjComLib.zlDatabase.CopyNewRec(rsSource, blnOnlyStructure, strFields, arrAppFields)
End Function

Public Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
'功能：判断某个ADO字段数据类型是否与指定字段类型是同一类(如数字,日期,字符,二进制)
    Dim intA As Integer, intB As Integer
    
    Select Case varBase
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intA = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intA = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intA = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intA = -4
        Case Else
            intA = varBase
    End Select
    Select Case varType
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intB = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intB = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intB = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intB = -4
        Case Else
            intB = varType
    End Select
    IsType = intA = intB
End Function

Public Function DownloadPicture(rsTable As ADODB.Recordset, strField As String, Optional strFile As String) As String
'功能：将指定的记录集图形字段下载到本地生成为图形临时文件
'参数：
'       rsTable   图形存储记录集
'       strField  图形字段
'       strFile   用户定义的文件名（可选项）
'返回：
    DownloadPicture = gobjComLib.zlDatabase.ReadPicture(rsTable, strField, strFile)
End Function

Public Function UploadPicture(strFile As String, rsTable As ADODB.Recordset, strField As String) As Boolean
'功能：将指定的图形文件上传到指定的记录集字段中
'参数：
'       strFile：图形文件名称
'       rsTable：图形存储记录集
'       strField：图形字段
'返回：
    UploadPicture = gobjComLib.zlDatabase.SavePicture(strFile, rsTable, strField)
End Function
