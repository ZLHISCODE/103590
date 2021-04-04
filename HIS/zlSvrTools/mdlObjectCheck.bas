Attribute VB_Name = "mdlObjectCheck"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Const GSTR_APPNAME As String = "对象检查修复"    '程序名

'变量定义
Private mrsSequenceFromFile As ADODB.Recordset
Private mrsViewFromFile As ADODB.Recordset
Private mrsPackageFromFile As ADODB.Recordset
Private mrsFildFromFile As ADODB.Recordset
Private mrsConstraintFromFile As ADODB.Recordset
Private mrsIndexFromFile As ADODB.Recordset
Private mrsProcedureFromFile As ADODB.Recordset

Private mrsSequenceFromDB As ADODB.Recordset
Private mrsViewFromDB As ADODB.Recordset
Private mrsPackageFromDB As ADODB.Recordset
Private mrsFildFromDB As ADODB.Recordset
Private mrsConstraintFromDB As ADODB.Recordset
Private mrsIndexFromDB As ADODB.Recordset
Private mrsProcedureFromDB As ADODB.Recordset

Private mrsDataFormFile As New ADODB.Recordset
Private mrsDataFormDB As New ADODB.Recordset
Private mrsProData As New ADODB.Recordset
Private mstrSysName As String
Private mlngNum As Long
Private mlngProgress As Long
Private mblnIndex As Boolean
Private mblnReport As Boolean
Private mblnzlTables As Boolean
Private mblnProcedure As Boolean
Private mblnParameter As Boolean

Public Function IniFilePathRecordset() As ADODB.Recordset
'初始化本地需检查路径的记录集

    Set IniFilePathRecordset = New ADODB.Recordset
    With IniFilePathRecordset
        .Fields.Append "FilePath", adVarChar, 1000, adFldIsNullable
        .Fields.Append "SystemNum", adDouble, 20, adFldIsNullable
        .Fields.Append "FileName", adVarChar, 50, adFldIsNullable
        .Fields.Append "FileType", adVarChar, 50, adFldIsNullable
        .Fields.Append "FullVer", adVarChar, 50, adFldIsNullable
        .Fields.Append "共享号", adDouble, 10, adFldIsNullable
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Function

Public Function InitDataRecordset() As ADODB.Recordset
'功能：初始化解析本地SQL文件数据保存的记录集
    
    Set InitDataRecordset = New ADODB.Recordset
    With InitDataRecordset
        .Fields.Append "类别", adVarChar, 50, adFldIsNullable
        .Fields.Append "SQL", adVarChar, 2000, adFldIsNullable
        .Fields.Append "系统编号", adVarChar, 50, adFldIsNullable
        .Fields.Append "序号", adVarChar, 100, adFldIsNullable
        .Fields.Append "对象", adVarChar, 100, adFldIsNullable
        .Fields.Append "参数号", adVarChar, 100, adFldIsNullable
        .Fields.Append "参数名", adVarChar, 2000, adFldIsNullable
        .Fields.Append "名称", adVarChar, 100, adFldIsNullable
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Function

Public Function InitProDataRecordset() As ADODB.Recordset
'功能：初始化解析本地SQL文件数据保存的记录集
    
    Set InitProDataRecordset = New ADODB.Recordset
    With InitProDataRecordset
        .Fields.Append "修正SQL", adVarChar, 2000, adFldIsNullable
        .Fields.Append "系统名称", adVarChar, 20, adFldIsNullable
        .Fields.Append "类别", adVarChar, 20, adFldIsNullable
        .Fields.Append "对象名", adVarChar, 1000, adFldIsNullable
        .Fields.Append "严重程度", adVarChar, 50, adFldIsNullable
        .Fields.Append "问题描述", adVarChar, 1000, adFldIsNullable
        .Fields.Append "修正说明", adVarChar, 1000, adFldIsNullable
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Function

Public Function GainData(ByRef rsSequenceFromFile As ADODB.Recordset, ByRef rsViewFromFile As ADODB.Recordset, ByRef rsPackageFromFile As ADODB.Recordset, ByRef rsFildFromFile As ADODB.Recordset, _
                        ByRef rsConstraintFromFile As ADODB.Recordset, ByRef rsIndexFromFile As ADODB.Recordset, ByRef rsProcedureFromFile As ADODB.Recordset, ByRef rsDataFormFile As ADODB.Recordset, _
                        ByVal rsSequenceFromDB As ADODB.Recordset, ByVal rsViewFromDB As ADODB.Recordset, ByVal rsPackageFromDB As ADODB.Recordset, ByVal rsFildFromDB As ADODB.Recordset, _
                        ByVal rsConstraintFromDB As ADODB.Recordset, ByVal rsIndexFromDB As ADODB.Recordset, ByVal rsProcedureFromDB As ADODB.Recordset, ByVal rsDataFormDB As ADODB.Recordset, _
                        ByVal blnIndex As Boolean, ByVal blnReport As Boolean, ByVal blnzlTables As Boolean, ByVal blnProcedure As Boolean, ByVal blnParameter As Boolean)
'初始化所需检查的所有数据
    Set mrsSequenceFromFile = rsSequenceFromFile
    Set mrsViewFromFile = rsViewFromFile
    Set mrsPackageFromFile = rsPackageFromFile
    Set mrsFildFromFile = rsFildFromFile
    Set mrsConstraintFromFile = rsConstraintFromFile
    Set mrsIndexFromFile = rsIndexFromFile
    Set mrsProcedureFromFile = rsProcedureFromFile
    Set mrsDataFormFile = rsDataFormFile
    
    Set mrsSequenceFromDB = rsSequenceFromDB
    Set mrsViewFromDB = rsViewFromDB
    Set mrsPackageFromDB = rsPackageFromDB
    Set mrsFildFromDB = rsFildFromDB
    Set mrsConstraintFromDB = rsConstraintFromDB
    Set mrsIndexFromDB = rsIndexFromDB
    Set mrsProcedureFromDB = rsProcedureFromDB
    Set mrsDataFormDB = rsDataFormDB
    
    mblnIndex = blnIndex
    mblnReport = blnReport
    mblnzlTables = blnzlTables
    mblnProcedure = blnProcedure
    mblnParameter = blnParameter
    
End Function

Public Sub CompareCheck(ByRef lngNum As Long, ByRef strSysName As String, ByRef rsPro As ADODB.Recordset, ByRef lngProgress As Long)
'功能：对比本地脚本和数据库进行比较
'参数：rsLocalObject-本地脚本解析的数据，rsOraObject-数据库查询的数据
    
    Set mrsProData = rsPro
    mstrSysName = strSysName
    mlngNum = lngNum
    mlngProgress = lngProgress

    Call CheckSequence
    Call CheckView
    Call CheckPackage
    Call CheckTable
    Call CheckConstraint
    Call CheckIndex
    Call CheckProcedure
    Call CheckBasicData
    lngProgress = mlngProgress
End Sub

Private Sub CheckBasicData()
'功能：检查基础数据
    Dim strSQL As String
    Dim strFild As String
    Dim strLevel As String
    
    '模块数据：系统，序号
    mrsDataFormFile.Filter = "类别='模块' and 系统编号=" & mlngNum
    If mrsDataFormFile.RecordCount > 0 Then mrsDataFormFile.MoveFirst
    Do While Not mrsDataFormFile.EOF
        Call frmAppCheck.ShowProgress(mstrSysName, mrsDataFormFile.RecordCount, mrsDataFormFile.AbsolutePosition, "模块数据")
        mrsDataFormDB.Filter = "类别='模块' and 系统编号=" & mrsDataFormFile!系统编号 & " and 序号=" & mrsDataFormFile!序号
        If mrsDataFormDB.RecordCount = 0 Then
            mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, mrsDataFormFile!SQL, "模块", "严重", _
                "数据库中该模块数据缺失，可能影响产品相关功能的正常使用", "添加该条模块数据", "序号,标题：" & mrsDataFormFile!序号 & "," & mrsDataFormFile!对象)
        End If
        DoEvents
        mrsDataFormFile.MoveNext
    Loop
    '功能数据：系统，序号，功能
    mrsDataFormFile.Filter = "类别='功能' and 系统编号=" & mlngNum
    If mrsDataFormFile.RecordCount > 0 Then mrsDataFormFile.MoveFirst
    Do While Not mrsDataFormFile.EOF
        Call frmAppCheck.ShowProgress(mstrSysName, mrsDataFormFile.RecordCount, mrsDataFormFile.AbsolutePosition, "功能数据")
        mrsDataFormDB.Filter = "类别='功能' and 系统编号=" & mrsDataFormFile!系统编号 & " and 序号=" & mrsDataFormFile!序号 & " and 对象='" & mrsDataFormFile!对象 & "'"
        If mrsDataFormDB.RecordCount = 0 Then
            mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, mrsDataFormFile!SQL, "功能", "严重", _
                "数据库中该功能数据缺失，可能影响产品相关功能的正常使用", "添加该条功能数据", "序号,功能：" & mrsDataFormFile!序号 & "," & mrsDataFormFile!对象)
        End If
        DoEvents
        mrsDataFormFile.MoveNext
    Loop
    '参数数据：模块，参数号，系统
    mrsDataFormFile.Filter = "类别='参数' and 系统编号=" & mlngNum
    If mrsDataFormFile.RecordCount > 0 Then mrsDataFormFile.MoveFirst
    Do While Not mrsDataFormFile.EOF
        Call frmAppCheck.ShowProgress(mstrSysName, mrsDataFormFile.RecordCount, mrsDataFormFile.AbsolutePosition, "参数数据")
        If mrsDataFormFile!对象 = "NULL" Then
            mrsDataFormDB.Filter = "类别='参数' and 系统编号=" & mrsDataFormFile!系统编号 & " and 参数号=" & mrsDataFormFile!参数号 & " and 对象='" & mrsDataFormFile!对象 & "'"
        Else
            mrsDataFormDB.Filter = "类别='参数' and 系统编号=" & mrsDataFormFile!系统编号 & " and 参数名='" & mrsDataFormFile!参数名 & "' and 对象='" & mrsDataFormFile!对象 & "'"
        End If
        If mrsDataFormDB.RecordCount = 0 Then
            mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, mrsDataFormFile!SQL, "参数", "严重", _
                "数据库中该参数数据缺失，可能影响产品相关功能的正常使用", "添加该条参数数据", "模块,参数号,参数名：" & mrsDataFormFile!对象 & "," & mrsDataFormFile!参数号 & "," & mrsDataFormFile!参数名)
        Else
            If mrsDataFormFile!对象 = "NULL" Then
                strLevel = "严重"
            Else
                strLevel = "轻微"
            End If
            If Not (strLevel = "轻微" And mblnParameter = False) Then
                If mrsDataFormDB!对象 = mrsDataFormFile!对象 Then
                    If Val(mrsDataFormDB!参数号) = Val(mrsDataFormFile!参数号) Then
                        If mrsDataFormDB!参数名 <> mrsDataFormFile!参数名 Then
                            If strLevel = "严重" Then
                                strSQL = "Update Zlparameters Set 参数名 ='" & mrsDataFormFile!参数名 & "' Where 系统 =" & mlngNum & " And 模块 is null And 参数号 =" & mrsDataFormFile!参数号
                            Else
                                strSQL = "Update Zlparameters Set 参数名 ='" & mrsDataFormFile!参数名 & "' Where 系统 =" & mlngNum & " And 模块 =" & mrsDataFormFile!对象 & " And 参数号 =" & mrsDataFormFile!参数号
                            End If
                            mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, strSQL, "参数", strLevel, _
                                "参数号相同,参数名(" & mrsDataFormDB!参数名 & ")不同，可能影响产品相关功能的正常使用", "调整参数名", "模块,参数号,参数名：" & mrsDataFormFile!对象 & "," & mrsDataFormFile!参数号 & "," & mrsDataFormFile!参数名)
                        End If
                    End If
                    If mrsDataFormDB!参数名 = mrsDataFormFile!参数名 Then
                        If Val(mrsDataFormDB!参数号) <> Val(mrsDataFormFile!参数号) Then
                            If strLevel = "严重" Then
                                strSQL = "Update Zlparameters Set 参数号 ='" & mrsDataFormFile!参数号 & "' Where 系统 =" & mlngNum & " And 模块 is null And 参数名 ='" & mrsDataFormFile!参数名 & "'"
                            Else
                                strSQL = "Update Zlparameters Set 参数号 ='" & mrsDataFormFile!参数号 & "' Where 系统 =" & mlngNum & " And 模块 =" & mrsDataFormFile!对象 & " And 参数名 ='" & mrsDataFormFile!参数名 & "'"
                            End If
                            mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, strSQL, "参数", strLevel, _
                                "参数名相同,参数号(" & mrsDataFormDB!参数号 & ")不同，可能影响产品相关功能的正常使用", "调整参数号", "模块,参数号,参数名：" & mrsDataFormFile!对象 & "," & mrsDataFormFile!参数号 & "," & mrsDataFormFile!参数名)
                        End If
                    End If
                End If
            End If
        End If
        DoEvents
        mrsDataFormFile.MoveNext
    Loop
    '报表数据：编号，系统
    If mblnReport Then
        mrsDataFormFile.Filter = "类别='报表' and 系统编号=" & mlngNum
        If mrsDataFormFile.RecordCount > 0 Then mrsDataFormFile.MoveFirst
        Do While Not mrsDataFormFile.EOF
            Call frmAppCheck.ShowProgress(mstrSysName, mrsDataFormFile.RecordCount, mrsDataFormFile.AbsolutePosition, "报表数据")
            mrsDataFormDB.Filter = "类别='报表' and 系统编号=" & mrsDataFormFile!系统编号 & " and 对象='" & mrsDataFormFile!对象 & "'"
            If mrsDataFormDB.RecordCount = 0 Then
                mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, "", "报表", "较重", _
                    "数据库中该报表数据缺失，可能影响产品相关功能的正常使用", "人工导入该报表", "编号、名称：" & mrsDataFormFile!对象 & "," & mrsDataFormFile!名称)
            End If
            DoEvents
            mrsDataFormFile.MoveNext
        Loop
    End If
    If mblnzlTables Then
        '表目录数据：表名，系统
        mrsDataFormFile.Filter = "类别='表目录' and 系统编号=" & mlngNum
        If mrsDataFormFile.RecordCount > 0 Then mrsDataFormFile.MoveFirst
        Do While Not mrsDataFormFile.EOF
            Call frmAppCheck.ShowProgress(mstrSysName, mrsDataFormFile.RecordCount, mrsDataFormFile.AbsolutePosition, "zlTables数据")
            mrsDataFormDB.Filter = "类别='表目录' and 系统编号=" & mrsDataFormFile!系统编号 & " and 对象='" & mrsDataFormFile!对象 & "'"
            If mrsDataFormDB.RecordCount = 0 Then
                mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, mrsDataFormFile!SQL, "表目录", "轻微", _
                    "数据库中该表目录数据缺失，可能影响产品相关功能的正常使用", "添加该条数据", "表名为：" & mrsDataFormFile!对象)
            End If
            DoEvents
            mrsDataFormFile.MoveNext
        Loop
    End If
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Private Sub CheckSequence()
'检查序列
    Dim i As Long
    Dim strName As String
    
    mrsSequenceFromFile.Filter = "系统编号=" & mlngNum
    For i = 1 To mrsSequenceFromFile.RecordCount
        strName = mrsSequenceFromFile!名称
        Call frmAppCheck.ShowProgress(mstrSysName, mrsSequenceFromFile.RecordCount, i, "序列", strName)
        mrsSequenceFromDB.Filter = "名称='" & strName & "'"
        If mrsSequenceFromDB.RecordCount = 0 Then
            mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, mrsSequenceFromFile!SQL, "序列", "严重", _
                                "数据库中该序列不存在，可能影响产品相关功能的正常使用", "添加该序列", strName)
        End If
        DoEvents
        mrsSequenceFromFile.MoveNext
    Next
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Private Sub CheckView()
'检查视图
    Dim i As Long
    Dim strName As String
    
    mrsViewFromFile.Filter = "系统编号=" & mlngNum
    For i = 1 To mrsViewFromFile.RecordCount
        strName = mrsViewFromFile!名称
        Call frmAppCheck.ShowProgress(mstrSysName, mrsViewFromFile.RecordCount, i, "视图", strName)
        mrsViewFromDB.Filter = "名称='" & strName & "'"
        If mrsViewFromDB.RecordCount = 0 Then
            mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, mrsViewFromFile!SQL, "视图", "严重", _
                                "数据库中该视图不存在，可能影响产品相关功能的正常使用", "添加该视图", strName)
        End If
        DoEvents
        mrsViewFromFile.MoveNext
    Next
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Private Sub CheckPackage()
'检查包
    Dim i As Long
    Dim strName As String
    Dim strReName As String
    
    mrsPackageFromFile.Filter = "系统编号=" & mlngNum
    For i = 1 To mrsPackageFromFile.RecordCount
        strName = mrsPackageFromFile!名称
        Call frmAppCheck.ShowProgress(mstrSysName, mrsPackageFromFile.RecordCount, i, "包", strName)
        mrsPackageFromDB.Filter = "名称='" & strName & "'"
        If mrsPackageFromDB.RecordCount = 0 Then
            mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, mrsPackageFromFile!SQL, "包", "严重", _
                "数据库中该包不存在，可能影响产品相关功能的正常使用", "添加该包", strName)
        Else
            If mrsPackageFromDB!Status <> "VALID" Then
                mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, mrsPackageFromFile!SQL, "包", "严重", _
                    "数据库中包处于无效状态，可能影响产品相关功能的正常使用", "重建包", strName)
            End If
        End If
        DoEvents
        mrsPackageFromFile.MoveNext
    Next
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Private Sub CheckTable()
'检查表以及字段
    Dim lngProgress As Long
    Dim strTableName As String

    Dim strFild As String
    Dim strFildLength As Long
    Dim strTemp As String
    Dim varTemp As Variant
    Dim strSQL As String
    
    lngProgress = 1
    mrsFildFromFile.Filter = "系统编号=" & mlngNum
    While Not mrsFildFromFile.EOF
        strTableName = mrsFildFromFile!表名
        
        Call frmAppCheck.ShowProgress(mstrSysName, mrsFildFromFile.RecordCount, lngProgress, "表及其字段", strTableName)
        mrsFildFromDB.Filter = "表名='" & strTableName & "'"
        
        If mrsFildFromDB.RecordCount = 0 Then
            mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, mrsFildFromFile!SQL, "表", "严重", _
                    "数据库中表不存在，可能影响产品相关功能的正常使用", "添加该表", strTableName)
            
            While strTableName = mrsFildFromFile!表名
                DoEvents
                mrsFildFromFile.MoveNext
                lngProgress = lngProgress + 1
            Wend
        Else
            Do While strTableName = mrsFildFromFile!表名
                strFild = mrsFildFromFile!字段
                mrsFildFromDB.Filter = "表名='" & strTableName & "' and 字段='" & strFild & "'"
                '判断表中该字段是否存在
                If mrsFildFromDB.RecordCount = 0 Then
                    If mrsFildFromFile!字段长度 <> "" Then
                        strSQL = "Alter Table " & strTableName & " Add " & mrsFildFromFile!字段 & " " & mrsFildFromFile!字段类型 & "(" & mrsFildFromFile!字段长度 & ")"
                    Else
                        strSQL = "Alter Table " & strTableName & " Add " & mrsFildFromFile!字段 & " " & mrsFildFromFile!字段类型
                    End If
                    mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, strSQL, "字段", "严重", _
                                    "数据库中表的该字段不存在，可能影响产品相关功能的正常使用", "添加该字段", strTableName & "：" & strFild)
                Else
                    '判断字段类型与数据库是否一致
                    If mrsFildFromFile!字段类型 <> mrsFildFromDB!字段类型 And mrsFildFromFile!字段类型 <> "VARCHAR" Then
                        If IsNull(mrsFildFromFile!字段长度) = False And IsNull(mrsFildFromDB!字段类型) = False Then
                            strSQL = "Alter Table " & strTableName & " Modify " & mrsFildFromFile!字段 & " " & mrsFildFromFile!字段类型 & "(" & mrsFildFromFile!字段长度 & ")"
                        Else
                            strSQL = "Alter Table " & strTableName & " Modify " & mrsFildFromFile!字段 & " " & mrsFildFromFile!字段类型
                        End If
                        mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, strSQL, "字段", "较重", _
                                        "数据库中表的该字段类型(" & mrsFildFromDB!字段类型 & ")与标准产品(" & mrsFildFromFile!字段类型 & ")中不一致，插入数据可能没有成功", "将数据库字段类型调整为与产品标准脚本一致", strTableName & "：" & strFild)
                    Else
                        '只检查NUMBER和VARCHAR2两种类型的字段长度
                        If mrsFildFromFile!字段类型 = "NUMBER" Then
                            If IsNull(mrsFildFromFile!字段长度) = False And IsNull(mrsFildFromDB!字段实际长度) = False Then
                                varTemp = Split(mrsFildFromFile!字段长度, ",")
                                strTemp = varTemp(0)
                                strFildLength = mrsFildFromDB!字段实际长度
                                If Val(strTemp) > Val(strFildLength) And Val(strFildLength) <> 0 Then
                                    strSQL = "Alter Table " & strTableName & " Modify " & mrsFildFromFile!字段 & " " & mrsFildFromFile!字段类型 & "(" & mrsFildFromFile!字段长度 & ")"
                                    mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, strSQL, "字段", "较重", _
                                                "数据库中该字段长度(" & strFildLength & ")比标准产品(" & strTemp & ")短，可能导致数据的不完整。", "将数据库字段长度调整为与产品标准脚本一致", strTableName & "：" & strFild)
                                ElseIf Val(strTemp) < Val(strFildLength) Then
                                    mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, "", "字段", "轻微", _
                                                "数据库中该字段长度(" & strFildLength & ")比标准产品(" & strTemp & ")长，一般不会影响产品正常使用", "人工调整脚本", strTableName & "：" & strFild)
                                End If
                            End If
                        ElseIf mrsFildFromFile!字段类型 = "VARCHAR2" Then
                            If IsNull(mrsFildFromFile!字段长度) = False And IsNull(mrsFildFromDB!字段实际长度) = False Then
                                strTemp = mrsFildFromFile!字段长度
                                strFildLength = mrsFildFromDB!字段长度
                                If Val(strTemp) > Val(strFildLength) And Val(strFildLength) <> 0 Then
                                    strSQL = "Alter Table " & strTableName & " Modify " & mrsFildFromFile!字段 & " " & mrsFildFromFile!字段类型 & "(" & mrsFildFromFile!字段长度 & ")"
                                    mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, strSQL, "字段", "较重", _
                                                "数据库中该字段长度(" & strFildLength & ")比标准产品(" & strTemp & ")短，可能导致数据的不完整。", "将数据库字段长度调整为与产品标准脚本一致", strTableName & "：" & strFild)
                                ElseIf Val(strTemp) < Val(strFildLength) Then
                                    mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, "", "字段", "轻微", _
                                                "数据库中该字段长度(" & strFildLength & ")比标准产品(" & strTemp & ")长，一般不会影响产品正常使用", "人工调整脚本", strTableName & "：" & strFild)
                                End If
                            End If
                        End If
                    End If
                End If
                DoEvents
                lngProgress = lngProgress + 1
                mrsFildFromFile.MoveNext
                If mrsFildFromFile.EOF Then Exit Do
            Loop
        End If
    Wend
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Private Sub CheckConstraint()
'检查约束
    Dim i As Long
    Dim strName As String
    Dim strSQL As String
    Dim varTemp As Variant
    Dim lngOra As Long
    Dim lngLocal As Long
    
    If mblnzlTables Then
        mrsConstraintFromDB.Filter = ""
        lngOra = mrsConstraintFromDB.RecordCount
    Else
        lngOra = 0
    End If
    
    mrsConstraintFromFile.Filter = "系统编号=" & mlngNum
    lngLocal = mrsConstraintFromFile.RecordCount
    For i = 1 To mrsConstraintFromFile.RecordCount
        strName = mrsConstraintFromFile!名称
        Call frmAppCheck.ShowProgress(mstrSysName, lngOra + lngLocal, i, "约束", strName)
        mrsConstraintFromDB.Filter = "名称='" & strName & "'"
        If mrsConstraintFromDB.RecordCount = 0 Then
            mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, mrsConstraintFromFile!SQL, "约束", "较重", _
                "数据库中该约束不存在，可能影响产品相关功能的正常使用", "添加该约束", strName)
        Else
            If mrsConstraintFromDB!Status <> "ENABLED" Then
                strSQL = "Alter Table " & mrsConstraintFromFile!表名 & " Enable Novalidate Constraint " & strName
                mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, strSQL, "约束", "较重", _
                    "约束当前处于禁止状态，可能影响产品相关功能的正常使用", "恢复约束", strName)
            End If
            If mrsConstraintFromDB!字段 <> mrsConstraintFromFile!字段 Then
                If strName = "体检任务细菌_FK_任务ID" And mrsConstraintFromFile!字段 = "任务ID,清单ID,病人ID" Then
                
                Else
                    strSQL = "Alter Table " & mrsConstraintFromFile!表名 & " Drop Constraint " & strName & " Cascade Drop Index"
                    strSQL = strSQL & "{JM|SQL分隔符}" & vbNewLine & mrsConstraintFromFile!SQL
                    mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, strSQL, "约束", "较重", _
                        "数据库中该约束列(" & mrsConstraintFromDB!字段 & ")与标准产品(" & mrsConstraintFromFile!字段 & ")不一致，可能影响产品相关查询性能", "删除该约束后再重建约束", strName)
                End If
            End If
        End If
        DoEvents
        mrsConstraintFromFile.MoveNext
    Next
    If mblnzlTables Then
        mrsConstraintFromDB.Filter = ""
        For i = 1 To mrsConstraintFromDB.RecordCount
            strName = mrsConstraintFromDB!名称
            mrsDataFormDB.Filter = "类别='表目录' and 对象='" & mrsConstraintFromDB!表名 & "'"
            If mrsDataFormDB.RecordCount = 1 Then
                If mrsDataFormDB!系统编号 = mlngNum Then
                    Call frmAppCheck.ShowProgress(mstrSysName, lngOra + lngLocal, lngLocal + i, "约束", strName)
                    mrsConstraintFromFile.Filter = "名称='" & strName & "'"
                    If mrsConstraintFromFile.RecordCount = 0 Then
                        strSQL = "Alter Table " & mrsConstraintFromDB!表名 & " Drop Constraint " & strName & " Cascade Drop Index"
                        mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, strSQL, "约束", "较重", _
                            "数据库中存在，但产品标准脚本没有，可能影响产品相关查询性能", "删除该约束", strName)
                    End If
                End If
            End If
            DoEvents
            mrsConstraintFromDB.MoveNext
        Next
    End If
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Private Sub CheckIndex()
'功能：检查索引
'说明：
    Dim i As Long
    Dim strName As String
    Dim strSQL As String
    Dim varFild As Variant
    Dim lngLocal As Long
    Dim lngOra As Long
    Dim blnSpace As Boolean
    Dim strSpace As String
    
    If mblnzlTables Then
        mrsIndexFromDB.Filter = ""
        lngOra = mrsIndexFromDB.RecordCount
    Else
        lngOra = 0
    End If
    
    mrsIndexFromFile.Filter = "系统编号=" & mlngNum
    lngLocal = mrsIndexFromFile.RecordCount
    For i = 1 To mrsIndexFromFile.RecordCount
        strName = mrsIndexFromFile!名称
        Call frmAppCheck.ShowProgress(mstrSysName, lngOra + lngLocal, i, "索引", strName)
        mrsIndexFromDB.Filter = "名称='" & strName & "'"
        If mrsIndexFromDB.RecordCount = 0 Then
            If Not (strName Like "*PK" Or strName Like "*UQ*") Then
                mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, Split(mrsIndexFromFile!SQL, "||")(0), "索引", "较重", _
                    "数据库中该索引不存在，可能影响产品运行速度", "添加该索引", strName)
            End If
        Else
            If mrsIndexFromDB!Status <> "VALID" Then
                strSQL = "Alter Index " & strName & " rebulid nologging"
                mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, strSQL, "索引", "较重", _
                                    "数据库中索引处于无效状态，可能影响产品运行速度", "重建索引", strName)
            End If
            If strName Like "*PK" Or strName Like "*UQ*" Then
                If mrsIndexFromDB!UNIQUENESS <> "UNIQUE" Then
                    If InStr(Split(mrsIndexFromFile!SQL, "||")(1), "INITIALLY DEFERRED") = 0 Then
                        strSQL = "Alter Table " & mrsIndexFromFile!表名 & " Drop Constraint " & strName & " Cascade Drop Index"
                        strSQL = strSQL & "{JM|SQL分隔符}" & vbNewLine & Split(mrsIndexFromFile!SQL, "||")(1)
                        mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, strSQL, "索引", "较重", _
                            "主键或唯一键对应的索引不是唯一索引，可能影响产品性能", "删除对应的约束后重建约束", strName)
                    End If
                End If
            Else
                If mrsIndexFromDB!字段 <> mrsIndexFromFile!字段 Then
                    strSQL = "drop index " & strName
                    strSQL = strSQL & "{JM|SQL分隔符}" & mrsIndexFromFile!SQL
                    mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, strSQL, "索引", "较重", _
                        "数据库中该索引列字段(" & mrsIndexFromDB!字段 & ")与产品标准脚本(" & mrsIndexFromFile!字段 & ")不一致，可能影响系统运行速度", "删除该索引后再重建索引", strName)
                End If
            End If
        End If
        DoEvents
        mrsIndexFromFile.MoveNext
    Next
    
    If mblnzlTables Then
        strName = ""
        mrsIndexFromDB.Filter = ""
        For i = 1 To mrsIndexFromDB.RecordCount
            If strName <> mrsIndexFromDB!名称 Then
                strName = mrsIndexFromDB!名称
                If Not (strName Like "*PK*" Or strName Like "*UQ*") Then
                    mrsDataFormDB.Filter = "类别='表目录' and 对象='" & mrsIndexFromDB!表名 & "'"
                    If mrsDataFormDB.RecordCount = 1 Then
                        If mrsDataFormDB!系统编号 = mlngNum Then
                            Call frmAppCheck.ShowProgress(mstrSysName, lngOra + lngLocal, lngLocal + i, "索引", strName)
                            mrsIndexFromFile.Filter = "名称='" & strName & "'"
                            If mrsIndexFromFile.RecordCount = 0 Then
                                strSQL = "drop index " & strName
                                mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, strSQL, "索引", "较重", _
                                    "数据库中存在，但产品标准脚本没有，可能影响产品相关功能的写入性能", "删除该索引", strName)
                            Else
                                If mblnIndex Then
                                    blnSpace = False
                                    strSpace = ""
                                    Do While Not mrsIndexFromFile.EOF
                                        If mrsIndexFromFile!表空间 = mrsIndexFromDB!表空间 And mrsIndexFromFile!表空间 <> "" Then
                                            blnSpace = True
                                        Else
                                            strSpace = mrsIndexFromFile!表空间
                                        End If
                                        mrsIndexFromFile.MoveNext
                                    Loop
                                    If blnSpace = False And IsNull(mrsIndexFromDB!表空间) = False And strSpace <> "" And mrsIndexFromDB!表名 <> "输血性质" Then
                                        strSQL = "Alter Index " & strName & " rebulid tablespace " & strSpace & " nologging"
                                        mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, strSQL, "索引", "轻微", _
                                            "索引表空间(" & mrsIndexFromDB!表空间 & ")与产品标准脚本(" & strSpace & ")不一致，可维护性降低", "重建索引", strName)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            DoEvents
            mrsIndexFromDB.MoveNext
        Next
    End If
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Private Sub CheckProcedure()
'功能：检查过程/函数
'说明：过程中的表空间值放的参数位置
    Dim i As Long
    Dim strName As String
    Dim strTemp As String
    Dim strFild As String
    Dim varDBFild As Variant
    Dim varFileFild As Variant
    Dim strSQL As String
    
    mrsProcedureFromFile.Filter = "系统编号=" & mlngNum
    For i = 1 To mrsProcedureFromFile.RecordCount
        strName = mrsProcedureFromFile!名称
        strFild = mrsProcedureFromFile!字段
        Call frmAppCheck.ShowProgress(mstrSysName, mrsProcedureFromFile.RecordCount, i, "过程/函数", strName)
        mrsProcedureFromDB.Filter = "名称='" & strName & "'"
        If mrsProcedureFromDB.RecordCount = 0 Then
            mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, mrsProcedureFromFile!SQL, "过程/函数", "严重", _
                "数据库中该过程或函数不存在，可能影响产品相关功能的正常使用", "添加该过程/函数", strName)
        Else
            If mrsProcedureFromDB!Status = "VALID" Then
                varFileFild = Split(mrsProcedureFromFile!字段, ",")
                varDBFild = Split(Nvl(mrsProcedureFromDB!字段, ""), ",")
                If UBound(varFileFild) < UBound(varDBFild) Then
                    mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, "", "过程/函数", "轻微", _
                        "数据库中该过程或函数参数个数(" & UBound(varDBFild) + 1 & "个)比标准产品(" & UBound(varFileFild) + 1 & "个)多，可能影响产品相关功能的正常使用", "人工调整脚本", strName)
                ElseIf UBound(varFileFild) > UBound(varDBFild) Then
                    mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, "", "过程/函数", "严重", _
                        "数据库中该过程或函数参数个数(" & UBound(varDBFild) + 1 & "个)比标准产品(" & UBound(varFileFild) + 1 & "个)少，可能影响产品相关功能的正常使用", "人工调整脚本", strName)
                Else
                    If mrsProcedureFromFile!字段 <> mrsProcedureFromDB!字段 Then
                        mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, "", "过程/函数", "严重", _
                            "数据库中该过程或函数参数顺序或名称(" & mrsProcedureFromDB!字段 & ")与标准产品(" & mrsProcedureFromFile!字段 & ")不一致，可能影响产品相关功能的正常使用", "人工调整脚本", strName)
                    End If
                End If
            Else
                If mblnProcedure Then
                    strTemp = UCase(Mid(mrsProcedureFromFile!SQL, 1, 50))
                    If InStr(strTemp, "PROCEDURE") > 0 Then
                        strSQL = "Alter procedure " & strName & " Compile"
                    Else
                        strSQL = "Alter Function " & strName & " Compile"
                    End If
                    mrsProData.AddNew Array("系统名称", "修正SQL", "类别", "严重程度", "问题描述", "修正说明", "对象名"), Array(mstrSysName, strSQL, "过程/函数", "严重", _
                        "数据库中该过程或函数处于无效状态，可能影响产品相关功能的正常使用", "重新编译该过程/函数", strName)
                End If
            End If
        End If
        DoEvents
        mrsProcedureFromFile.MoveNext
    Next
    mlngProgress = mlngProgress + 1
    Call frmAppCheck.ShowFinalPro(mlngProgress)
End Sub

Public Function SetSelectRecordset(ByRef strSelect As String, ByRef strFilds As String, ByVal arrFields As Variant, _
     ByRef strTableName As String) As ADODB.Recordset
'功能：将Insert Into语句的字段名和字段值转换成记录集对象
'参数：
'  strSelect：Insert Into语句
'  strFilds：字段值
'  arrFields：Insert Into的字段名数组
'  strTableName：表对象名
'返回：记录集对象
    Dim rsSelect As New ADODB.Recordset
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Integer, j As Integer
    Dim arrSelect As Variant, arrValue As Variant
    Dim strTmp As String, strSQL As String, strHead As String
    Dim bytLevel As Byte
    Dim strTeam As String, strParentLevel As String
    Dim varTemp As Variant
    Dim strModifySQL As String

    strSelect = UCase(strSelect)
    
    '生成记录集
    With rsSelect
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Fields.Append "修正SQL", adVarChar, 2000
        For i = LBound(arrFields) To UBound(arrFields)
            If (arrFields(i) Like "*号" Or arrFields(i) = "菜单ID") And arrFields(i) <> "编号" Then
                .Fields.Append Trim(arrFields(i)), adBigInt
            Else
                .Fields.Append Trim(arrFields(i)), adVarChar, 300
            End If
        Next
        .Open
    End With
    
    '从strScript取出字段值
    If strSelect Like "*SELECT *" Then
        'select格式
        lngBegin = InStr(strSelect, "SELECT ")
        arrSelect = Split(Mid(strSelect, lngBegin), "UNION ALL")
    Else
        'values格式
        If strSelect Like "*VALUES*" Then
            lngBegin = InStr(strSelect, "VALUES") + 6
            lngBegin = InStr(Mid(strSelect, lngBegin), "(") + lngBegin
        Else
            Exit Function
        End If
        lngEnd = InStr(Mid(strSelect, lngBegin), ")") - 1
        arrSelect = Array()
        ReDim Preserve arrSelect(UBound(arrSelect) + 1)
        strTeam = Mid(strSelect, lngBegin, lngEnd)
        If InStr(strTeam, "(") > 0 Then strTeam = strTeam & ")"
        arrSelect(UBound(arrSelect)) = "SELECT " & strTeam & " FROM DUAL "
    End If
    '向记录集写值
    strHead = ""
    On Error GoTo errHandle
    For i = LBound(arrSelect) To UBound(arrSelect)
        strSQL = Trim(ClearSpace(arrSelect(i), True))
        strTmp = strSQL
        If Trim(strTmp) <> "" Then
            If strTmp Like "SELECT *[,| ]A.[*] FROM *" Then
                '特殊头
                lngBegin = InStr(strTmp, "SELECT ") + 7
                lngEnd = InStr(Mid(strTmp, lngBegin), ",A.*") - 1
                If lngEnd < 0 Then
                    lngEnd = InStr(Mid(strTmp, lngBegin), ", A.*") - 1
                End If
                strHead = Mid(strTmp, lngBegin, lngEnd)
            ElseIf strTmp Like "*) A;*" Then
                '特殊尾
            Else
                '数据
                If strTmp Like "* DUAL*" Then
                    lngBegin = InStr(strTmp, "SELECT ") + 7
                    lngEnd = InStr(Mid(strTmp, lngBegin), " FROM ")
                    If strHead <> "" Then
                        strSQL = strHead & "," & Mid(strSQL, lngBegin, lngEnd)
                    Else
                        strSQL = Mid(strSQL, lngBegin, lngEnd)
                    End If
                    arrValue = Split(strSQL, ",")
                    strModifySQL = "INSERT INTO " & strTableName & "(" & strFilds & ") VALUES (" & strSQL & ")"
                    
                    rsSelect.AddNew
                    rsSelect!修正SQL = strModifySQL
                    For j = LBound(arrValue) + 1 To UBound(arrValue) + 1
                        If rsSelect.Fields(j).Type = adVarChar Then
                            rsSelect.Fields(j).value = Trim(Replace(Trim(arrValue(j - 1)), "'", ""))
                        Else
                            If Not Trim(LCase(arrValue(j - 1))) Like "*NULL" Then
                                rsSelect.Fields(j).value = Trim(Replace(Trim(arrValue(j - 1)), "'", ""))
                            Else
                                rsSelect.Fields(j).value = 0
                            End If
                        End If
                        rsSelect.Fields(j).value = Trim(rsSelect.Fields(j).value)
                    Next
                    rsSelect.Update
                End If
            End If
        End If
    Next
    
    Set SetSelectRecordset = rsSelect
    Exit Function
    
errHandle:

End Function

Public Function ReplaceNoteMark(ByVal strScript As String, ByVal strSymbol As String, ByVal strSymbolNew As String) As String
'功能：对脚本中引号内的符号替换
'参数：
'  strScript：要处理的SQL脚本
'  strSymbol：指定原字符
'  strSymbolNew：替换的字符
'返回：

    Const STR_SQM  As String = "'"

    Dim l As Long
    Dim blnStart As Boolean
    Dim strTmp As String
    
    If strSymbol = "" Then Exit Function
    If Len(strSymbol) > 1 Then Exit Function
    For l = 1 To Len(strScript)
        If Mid(strScript, l, 1) = STR_SQM Then
            blnStart = Not blnStart
            strTmp = strTmp & Mid(strScript, l, 1)
        Else
            If Mid(strScript, l, 1) = strSymbol And blnStart Then
                strTmp = strTmp & strSymbolNew
            Else
                strTmp = strTmp & Mid(strScript, l, 1)
            End If
        End If
    Next
    
    ReplaceNoteMark = strTmp
End Function

Public Function ClearSpace(ByVal strVal As String, Optional ByVal blnSpace As Boolean = False) As String
'功能：清理多余的空格
'参数：
'  strVal：需要清理的字串
'  blnSpace：换行符转空格符
'返回：已清理后的字串

    Dim strResult As String
    Dim l As Long
    Dim blnStart As Boolean
    
    If strVal = "" Then Exit Function
    
    '保留单引号内的回车换行符
    For l = 1 To Len(strVal)
        If Mid(strVal, l, 1) = "'" Then
            blnStart = Not blnStart
        End If
        If blnStart Then
            If Asc(Mid(strVal, l, 1)) = Asc(vbCrLf) Or Asc(Mid(strVal, l, 1)) = Asc(vbCr) Then
                strVal = Left(strVal, l - 1) & "[[ENTER]]" & Mid(strVal, l + 1)
            End If
        End If
    Next
    
    '
    strResult = Replace(Replace(strVal, vbTab, " "), vbCrLf, IIf(blnSpace, " ", ""))
    If blnSpace Then strResult = Replace(strResult, vbCr, " ")
    
    '还原回车换行符
    strResult = Replace(strResult, "[[ENTER]]", vbCr)
    
    Do While InStr(strResult, "  ") > 0
        strResult = Replace(strResult, "  ", " ")
    Loop
    
    ClearSpace = strResult
    
End Function

Public Function GetSystemList() As ADODB.Recordset
'功能：获取Zlsystems中各系统的信息
    Dim rsSys As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 编号 系统编号, 名称 系统名称, 版本号 系统版本号, 所有者 系统所有者, 共享号, 正常安装 From Zlsystems where Upper(所有者)=[1] Order by Nvl(共享号,0),编号"
    Set rsSys = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "读取系统清单信息", gstrUserName)
    
    Set GetSystemList = rsSys
End Function

Public Function GetSystemSetupIni() As ADODB.Recordset
'功能：获取Zlsystems中各系统的安装脚本文件位置
    Dim rsSys As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select A.系统 系统编号, A.操作, upper(A.文件名) 文件名 From Zlsysfiles a Where Upper(操作人)=[1] and  A.操作 in(1,2) Order By 系统,操作"
    Set rsSys = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "读取系统安装脚本文件", gstrUserName)
    
    Set GetSystemSetupIni = rsSys
End Function

Public Sub ReleaseMe()
'关闭检查结果窗体时释放模块窗体

    Set mrsSequenceFromFile = Nothing
    Set mrsViewFromFile = Nothing
    Set mrsPackageFromFile = Nothing
    Set mrsFildFromFile = Nothing
    Set mrsConstraintFromFile = Nothing
    Set mrsIndexFromFile = Nothing
    Set mrsProcedureFromFile = Nothing
    Set mrsDataFormFile = Nothing
    
    Set mrsSequenceFromDB = Nothing
    Set mrsViewFromDB = Nothing
    Set mrsPackageFromDB = Nothing
    Set mrsFildFromDB = Nothing
    Set mrsConstraintFromDB = Nothing
    Set mrsIndexFromDB = Nothing
    Set mrsProcedureFromDB = Nothing
    Set mrsDataFormDB = Nothing
    
    Set mrsProData = Nothing
    mstrSysName = ""
End Sub

