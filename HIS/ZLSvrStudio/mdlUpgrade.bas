Attribute VB_Name = "mdlUpgrade"
Option Explicit
'文件类型,该类型顺序与文件执行顺序相同
Public Enum FileType
    FT_BefUp = 1 '提前执行脚本：ZLUPgradeX.X.X_Before.sql.sql(管理工具）,ZL*_X.X.X_History_Before.sql (应用系统历史库)ZL*_X.X.X_Before.sql(应用系统在线库) *代表系统号\100
    FT_DBAUp = 2 '需要DBA用户执行的脚本(System用户):ZLUPgradeX.X.X_DBA.sql,ZL*_X.X.X_DBA.sql
    FT_StUp = 3 '普通升级脚本：ZLUPgradeX.X.X.sql,ZLUPgradeX.X.X(补充).sql,ZL*_X.X.X.sql ,ZL*_X.X.X(补充).sql,ZL*_X.X.X_History.sql
    FT_OptUp = 4 '可选执行脚本:ZLUPgradeX.X.X_Optional.sql,ZL*_X.X.X_Optional.sql，ZL*_X.X.X__HISTORY_Optional.sql
    FT_DefUp = 5 '延迟执行脚本:ZL*_X.X.X_Deferred.sql,ZL*_X.X.X__HISTORY_DEFERRED
End Enum
'文件所属系统
Public Enum SysType
    ST_Tools = 1 '管理工具脚本,具有文件类型：FT_BefUp,FT_DBAUp,FT_StUp,FT_OptUp
    ST_App = 2 '应用系统在线库,具有文件类型：FT_BefUp,FT_DBAUp,FT_StUp,FT_OptUp，FT_DefUp
    ST_AppHis = 3 '应用系统历史库，具有文件类型：FT_BefUp,FT_StUp,FT_DefUp，FT_OptUp
End Enum
'版本类型
Public Enum VersionType
    VT_Normal = 1 '正常版本
    VT_Supple = 2 '补充发布版本，下一个大版本发布后，前一个版本新发布的SP就是补充版本
End Enum

'Public Enum FileTypeSys
'    FT_toolsUp = 1 '管理工具脚本中 格式如 ZLUPgradeX.X.X.sql类的文件
'    FT_toolsUpDbA = 2 '管理工具脚本中 格式如 ZLUPgradeX.X.X_DBA.sql类的文件
'    FT_toolsUpOpt = 3 '管理工具脚本中 格式如 ZLUPgradeX.X.X_Optional.sql类的文件
'    FT_toolsUpBef = 4 '管理工具脚本中 格式如 ZLUPgradeX.X.X_Before.sql.sql类的文件
'
'    FT_SysUp = 1 '系统升级脚本中 格式如ZL*_X.X.X.sql 类的文件  *代表系统号\100
'    FT_SysUpDBA = 2 '系统升级脚本中 格式如ZL*_X.X.X_DBA.sql 类的文件 *代表系统号\100
'    FT_SysUpOpt = 3 '系统升级脚本中 格式如ZL*_X.X.X_Optional.sql 类的文件 *代表系统号\100
'    FT_SysUpHis = 4 '系统升级脚本中 格式如ZL*_X.X.X_History.sql 类的文件 *代表系统号\100
'    FT_SysUpBef = 5 '系统升级脚本中 格式如ZL*_X.X.X_Before.sql 类的文件  *代表系统号\100
'    FT_SysUpHisBef = 6 '系统升级脚本中 格式如ZL*_X.X.X_History_Before.sql 类的文件  *代表系统号\100
'    FT_SysUpDef = 7 '系统升级脚本中 格式如ZL*_X.X.X_Deferred.sql 类的文件 *代表系统号\100
'    FT_SysUpHisDef = 8 '系统升级脚本中 格式如ZL*_X.X.X__HISTORY_DEFERRED 类的文件 *代表系统号\100
'    FT_SysUpOther = 9 '系统升级脚本中 格式如ZL*_X.X.X(补充).sql 类的文件 *代表系统号\100
'End Enum

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional blnOnlyStructure As Boolean, Optional ByVal strFields As String, Optional arrAppFields As Variant) As ADODB.Recordset
'编制人:朱玉宝
'修改人：刘硕
'修改日期：2014-1-6
'修改点：增加复制记录集的部分字段功能
'编制日期:2000-11-02
'复制记录集
'参数：strFields=需要复制的记录集的字段的列顺序或字段名组成的字符串
'          如：1 别名1,3 别名2,7 别名3...表示复制记录集的第1,3,7..字段组成记录集并返回
'              ID 别名1,姓名 别名2,....表示复制记录集的ID,姓名...字段组成记录集返回
'              别名*为新的记录集的列名
'              两中类型混搭容易出现列名相同的问题，请注意
'           arrAppFields=追加的字段信息：列名,类型,长度,默认值,没有默认值传Empty,没有指定长度传Empty
'      blnOnlyStructure=是否只复制结构
'在程序中，经常会涉及到相互传递记录集，而使用ADO的Clone复制产生的记录集，当其中一个记录集的数据发生变化的时候，所有副本都将发生相同的变化（通常指修改或删除），而我们往往希望这些记录集相互间保持独立
  
    Dim rsClone As New ADODB.Recordset
    Dim rsTarget As New ADODB.Recordset
    Dim intFields As Integer
    Dim arrFieldsName As Variant, strFieldName As String, strFieldNameAlias As String
    Dim arrTmp As Variant
    Dim i As Long
    
    If Not rsSource Is Nothing Then
        Set rsClone = rsSource.Clone
        rsClone.Filter = rsSource.Filter
    End If
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        '产生记录集结构
        If strFields = "" Then '记录集全复制模式
            arrFieldsName = Array()
            If Not rsClone Is Nothing Then
                ReDim arrFieldsName(rsClone.Fields.Count - 1)
                For intFields = 0 To rsClone.Fields.Count - 1
                    arrFieldsName(intFields) = rsClone.Fields(intFields).name & ""
                    .Fields.Append rsClone.Fields(intFields).name, IIf(rsClone.Fields(intFields).Type = adNumeric, adDouble, rsClone.Fields(intFields).Type), rsClone.Fields(intFields).DefinedSize, adFldIsNullable    '0:表示新增
                Next
            End If
        Else '记录集部分复制模式
            arrFieldsName = Split(strFields, ",")
            For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                '列包含别名
                arrTmp = Split(arrFieldsName(intFields) & " ", " ")
                strFieldName = Trim(arrTmp(0)): strFieldNameAlias = Trim(arrTmp(1))
                If IsNumeric(strFieldName) Then strFieldName = rsClone.Fields(Val(strFieldName)).name & ""
                '获取字段原名，存入数组
                arrFieldsName(intFields) = strFieldName
                '添加字段,若果存在别名，则新增列的列名为别名
                .Fields.Append IIf(strFieldNameAlias = "", strFieldName, strFieldNameAlias), IIf(rsClone.Fields(strFieldName).Type = adNumeric, adDouble, rsClone.Fields(strFieldName).Type), rsClone.Fields(strFieldName).DefinedSize, adFldIsNullable '0:表示新增
            Next
        End If
        '追加字段添加
        If TypeName(arrAppFields) = "Variant()" Then
            For i = LBound(arrAppFields) To UBound(arrAppFields) Step 4
                If arrAppFields(i + 2) = Empty Then
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable, arrAppFields(i + 3)
                    End If
                Else
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable, arrAppFields(i + 3)
                    End If
                End If
            Next
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        '复制数据
        If Not blnOnlyStructure Then
            If rsClone Is Nothing Then Exit Function
            If rsClone.RecordCount <> 0 Then rsClone.MoveFirst
            Do While Not rsClone.EOF
                .AddNew
                For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                    '新记录集的列按顺序添加，因此可以这样
                    .Fields(intFields).value = rsClone.Fields(arrFieldsName(intFields)).value
                Next
                .Update
                rsClone.MoveNext
            Loop
            If rsClone.RecordCount <> 0 Then .Filter = "": .MoveFirst
        End If
    End With
    
    Set CopyNewRec = rsTarget
End Function

Public Function RecDelete(ByRef rsInput As ADODB.Recordset, Optional ByVal strFilter As String) As Boolean
'功能：删除指定条件的记录集的记录
'参数：rsInput=记录集
'      strFilter=条件
'返回：是否成功
'      rsInput=经过删除后的记录集
    rsInput.Filter = strFilter
    If rsInput.RecordCount > 0 Then
        rsInput.MoveFirst
        Do While Not rsInput.EOF
            Call rsInput.Delete
            rsInput.MoveNext
        Loop
        Call rsInput.UpdateBatch
    End If
    RecDelete = True
End Function

Public Function RecUpdate(ByRef rsInput As Recordset, ByVal strFilter As String, ParamArray arrInput() As Variant) As Boolean
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
                        If err.Number <> 0 Then err.Clear: blnFiled = False
                        On Error GoTo errH
                    End If
                    If Not blnFiled Then
                        rsInput(strFiledName).value = arrInput(i + 1)
                    Else
                        rsInput(strFiledName).value = rsInput(Mid(arrInput(i + 1), 2)).value
                    End If
                End If
                blnFiled = False
            Next
            .MoveNext
        Loop
        Call rsInput.UpdateBatch
    End With
    RecUpdate = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function RecDataAppend(ByRef rsSource As ADODB.Recordset, ByVal rsAppend As ADODB.Recordset, ParamArray arrInput() As Variant) As Boolean
'功能：将指定记录集的数据添加到另一个记录集上
'参数：rsSource=目标记录集
'      rsAppend=数据记录集
'      arrInput=字段对应规则，该参数不传时，默认两记录集结构相同，格式：arrInput(0):[记录集1].字段1,字段2...；arrInput(1)：[记录集2].字段1,字段2...
'返回：是否成功
'      rsSource=添加数据后的记录集
    Dim arrSource As Variant, arrAppend As Variant
    Dim i As Long, arrValues() As Variant
    Dim strTmp As String
    
    If rsAppend Is Nothing Then RecDataAppend = True: Exit Function
    If rsAppend.RecordCount = 0 Then RecDataAppend = True: Exit Function
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
            For i = 0 To rsSource.Fields.Count - 1
                rsSource.Fields(i).value = rsAppend.Fields(i).value
            Next
            rsSource.Update
            rsAppend.MoveNext
        Loop
    End If
    
    RecDataAppend = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
    
End Function

Public Function ReadINIToRec(ByVal strFile As String) As ADODB.Recordset
'功能：将指定INI配置文件的内容读取到记录集中
'返回：Nothing或包含"项目,内容"的记录集,其中同一项目可能有多行内容
    Dim rsTmp As New ADODB.Recordset
    Dim objINI As Scripting.TextStream
    
    Dim strItem As String, strText As String
    Dim strLine As String
            
    rsTmp.Fields.Append "项目", adVarChar, 100
    rsTmp.Fields.Append "内容", adVarChar, 4000, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set objINI = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objINI.AtEndOfStream
        strLine = Replace(objINI.ReadLine, vbTab, " ")
        If Left(Trim(strLine), 1) = "[" And InStr(strLine, "]") > InStr(strLine, "[") Then
            If strItem <> "" And strText = "" Then
                rsTmp.AddNew
                rsTmp!项目 = strItem
                rsTmp!内容 = Null
                rsTmp.Update
            End If
            strItem = Trim(Mid(strLine, InStr(strLine, "[") + 1, InStr(strLine, "]") - InStr(strLine, "[") - 1))
            strText = Trim(Mid(strLine, InStr(strLine, "]") + 1))

            If strItem <> "" And strText <> "" Then
                rsTmp.AddNew
                rsTmp!项目 = strItem
                rsTmp!内容 = strText
                rsTmp.Update
            End If
        ElseIf Trim(strLine) <> "" And strItem <> "" Then
            strText = Trim(strLine)
            rsTmp.AddNew
            rsTmp!项目 = strItem
            rsTmp!内容 = strText
            rsTmp.Update
        End If
    Loop
    
    If strItem <> "" And strText = "" Then
        rsTmp.AddNew
        rsTmp!项目 = strItem
        rsTmp!内容 = Null
        rsTmp.Update
    End If
    
    objINI.Close
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    
    Set ReadINIToRec = rsTmp
End Function

Public Function CheckINIValid(rsINI As ADODB.Recordset, ByVal strItem As String) As Boolean
'功能：检查对应的配置文件格式是否正确
'参数：rsINI=存放配置文件内容的记录集，包含"项目,内容"字段
'      strItem=配置文件中必须要求有内容的项目串,如"项目1|项目2|..."
    Dim arrItem As Variant, i As Long
    
    arrItem = Split(strItem, "|")
    For i = 0 To UBound(arrItem)
        rsINI.Filter = "项目='" & arrItem(i) & "'"
        If rsINI.EOF Then Exit Function
        If IsNull(rsINI!内容) Then Exit Function
        If arrItem(i) Like "*版本号" Then
            If Not IsVerSion(rsINI!内容) Then Exit Function
        End If
    Next
    CheckINIValid = True
End Function

Public Function VerCompare(ByVal strVer1 As String, ByVal strVer2 As String, Optional ByVal blnPrimary As Boolean) As Integer
'功能：比较两个字符串表示的版本号的大小
'参数：blnPrimary=是否只比较"主版本.次版本",不管附版本
'返回：1=strVer1>strVer1,-1=strVer1<strVer1,0=strVer1=strVer1
'说明：VB最大支持的版本号为9999.9999.9999
    Dim arrVer As Variant
    
    If strVer1 Like "*.*.*" And strVer2 Like "*.*.*" Then
        arrVer = Split(strVer1, ".")
        strVer1 = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & IIf(blnPrimary, "", "." & Format(arrVer(2), "0000"))
        
        arrVer = Split(strVer2, ".")
        strVer2 = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & IIf(blnPrimary, "", "." & Format(arrVer(2), "0000"))
    End If
    If strVer1 > strVer2 Then
        VerCompare = 1
    ElseIf strVer1 < strVer2 Then
        VerCompare = -1
    End If
End Function

Public Function GetCurPriVersion(ByVal strVer As String) As String
'功能：获取当前版本的大版本
'参数：strVer 当前版本
'返回： GetNextVersion 当前版本的大版本
    Dim arrVer As Variant
    
    If IsVerSion(strVer) Then
        If Not strVer Like "*.*.0" Then
            arrVer = Split(strVer, ".")
            strVer = arrVer(0) & "." & arrVer(1) & ".0"
        End If
    Else
        Exit Function
    End If
    
    GetCurPriVersion = strVer
End Function

Public Function GetNextVersion(ByVal strVer As String, Optional ByVal blnPrimary As Boolean) As String
'功能：获取当前SP版本的下一个版本
'参数：strVer 当前版本
'     blnPrimary 是否获取下一个大版本
'返回： GetNextVersion blnPrimary=true:下一个大版本,blnPrimary=false :下一个SP版本
    Dim arrVer As Variant
    
    If IsVerSion(strVer) Then
        arrVer = Split(strVer, ".")
        If blnPrimary Then
            strVer = Val(arrVer(0)) & "." & Val(arrVer(1)) + 1 & ".0"
        Else
            strVer = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & Val(arrVer(2)) + 10
        End If
    Else
        Exit Function
    End If
    
    GetNextVersion = strVer
    
End Function

Public Function GetPreVersion(ByVal strVer As String, Optional ByVal blnPrimary As Boolean) As String
'功能：获取当前SP版本的上一个版本
'参数：strVer 当前版本
'     blnPrimary 是否获取上一个大版本
'返回： GetNextVersion blnPrimary=true:上一个大版本,blnPrimary=false :上一个SP版本
    Dim arrVer As Variant
    
    If IsVerSion(strVer) Then
        arrVer = Split(strVer, ".")
        If blnPrimary Then
            If Val(arrVer(1)) - 1 < 0 Then
               arrVer(1) = "*"
               arrVer(0) = Val(arrVer(0)) - 1
            Else
                arrVer(1) = Val(arrVer(1)) - 1
            End If
            strVer = Val(arrVer(0)) & "." & arrVer(1) - 1 & ".0"
        Else
            If Val(arrVer(2)) - 10 < 0 Then
                arrVer(2) = "*"
                If Val(arrVer(1)) - 1 < 0 Then
                   arrVer(1) = "*"
                   arrVer(0) = Val(arrVer(0)) - 1
                Else
                    arrVer(1) = Val(arrVer(1)) - 1
                End If
            Else
                arrVer(2) = Val(arrVer(2)) - 10
            End If
            strVer = Val(arrVer(0)) & "." & arrVer(1) & "." & arrVer(2)
        End If
    Else
        Exit Function
    End If
    
    GetPreVersion = strVer
End Function

Public Function GetFileInfo(ByVal strFile As String, ByVal lngSys As Long, Optional ByRef strVerReturn As String, Optional ByRef ftReturn As FileType, _
                                    Optional ByRef stReturn As SysType, Optional ByRef vtReturn As VersionType) As Boolean
'功能:获取文件信息
'参数：
'   strFile=不包含路径的文件名,带扩展名
'   lngSys=系统号
'返回:
'       True=成功获取，False=获取失败（文件不是系统升级脚本）
'       strVerReturn=文件版本
'       ftReturn=文件类型
'       stReturn=系统类型
'       vtReturn=版本类型
    Dim strSysString, strSuffix As String
    Dim arrVer As Variant, strVerTmp As String
    '初始化变量
    strVerReturn = "": ftReturn = 0: stReturn = 0: vtReturn = VT_Normal
    If Not UCase(strFile) Like "*.SQL" Then Exit Function
    strFile = UCase(Left(strFile, Len(strFile) - 4))
    '获取脚本系统前缀
    If strFile Like "ZLUPGRADE*.*.*" Then
        strSysString = "ZLUPGRADE"
        stReturn = ST_Tools
    ElseIf strFile Like "ZL" & lngSys \ 100 & "_*.*.*" Then
        strSysString = "ZL" & lngSys \ 100 & "_"
        stReturn = ST_App
    Else
        Exit Function '没有系统标识前缀，不是系统脚本
    End If
    '系统标识后面紧跟的是版本
    strSuffix = Mid(strFile, Len(strSysString) + 1)
    arrVer = Split(strSuffix, ".")
    If UBound(arrVer) <> 2 Then Exit Function '不含版本的脚本，不是系统脚本
    strVerTmp = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & Val(arrVer(2))
    If Not IsVerSion(strVerTmp) Then Exit Function '不是版本信息
    If Not strSuffix Like strVerTmp & "*" Then Exit Function
    '版本后是文件类型信息
    strSuffix = Mid(strSuffix, Len(strVerTmp) + 1)
    If InStr(strSuffix, "(补充)") > 0 Then
        vtReturn = VT_Supple
        strSuffix = Replace(strSuffix, "(补充)", "") '防止补充信息位置不固定
    End If
    If stReturn = ST_App And strSuffix Like "_HISTORY*" Then
        stReturn = ST_AppHis
        strSuffix = Mid(strSuffix, Len("_HISTORY") + 1)
    End If
    Select Case strSuffix
        Case ""
            ftReturn = FT_StUp
        Case "_DBA"
            If stReturn <> ST_AppHis Then ftReturn = FT_DBAUp
        Case "_OPTIONAL"
            ftReturn = FT_OptUp
        Case "_BEFORE"
            ftReturn = FT_BefUp
        Case "_DEFERRED"
            If stReturn <> ST_Tools Then ftReturn = FT_DefUp
    End Select
    If ftReturn = 0 Then Exit Function
    strVerReturn = strVerTmp
    GetFileInfo = True
End Function

Public Function VerFull(ByVal strVer As String, Optional ByVal blnMax As Boolean = True) As String
'功能：返回VB最大支持的版本号形式:9999.9999.9999,最小版本号0000.0000.0000
'参数：strVer=当前版本号
'           blnMax=True,若果为空，则返回最大支持版本，False=若果为空，则返回最小支持版本
    Dim arrVer As Variant
    If strVer = "" Then
        VerFull = IIf(blnMax, "9999.9999.9999", "0000.0000.0000")
        Exit Function
    End If
    arrVer = Split(strVer, ".")
    VerFull = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & "." & Format(arrVer(2), "0000")
End Function

Public Function VerNormal(ByVal strVer As String) As String
'功能：将VB最大支持的版本号形式:9999.9999.9999转换为常见版本虚形式，如0010.0034.0000，转换为10.34.0
    Dim arrVer As Variant
    If strVer = "" Then
        VerNormal = "0.0.0"
        Exit Function
    End If
    arrVer = Split(strVer, ".")
    VerNormal = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & Val(arrVer(2))
End Function

Public Function IsVerSion(ByVal strVer As String) As Boolean
'功能：判断字符串是否是版本号
    Dim arrVer As Variant
    Dim i As Integer
    If strVer = "" Then Exit Function
    arrVer = Split(strVer, ".")
    If UBound(arrVer) <> 2 Then Exit Function
    
    For i = LBound(arrVer) To UBound(arrVer)
        If Val(arrVer(i)) < 0 Or Val(arrVer(i)) > 9999 Then Exit Function
        If Val(arrVer(i)) & "" <> arrVer(i) Then Exit Function
    Next
    
    IsVerSion = True
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
'功能：取指定字符串按字节算的长度
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function ActualStr(ByVal strAsk As String, ByVal lngLen As Long) As String
'功能：取指定字符串左边指定字节长度的内容
    Dim strTemp As String, i As Long
    
    strTemp = StrConv(LeftB(StrConv(strAsk, vbFromUnicode), lngLen), vbUnicode)
    If InStr(strTemp, Chr(0)) > 0 Then
        strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
    End If
    ActualStr = strTemp
End Function

Public Function TrimEx(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'说明：主要是RunSQLFile的子函数
    Dim i As Long
    
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    i = 5
    Do While i > 1
        strText = Replace(strText, String(i, " "), " ")
        If InStr(strText, String(i, " ")) = 0 Then i = i - 1
    Loop
    TrimEx = strText
End Function

Public Function TrimComment(ByVal strSQL As String) As String
'功能：去掉写在单行strSQL语句后面的"--"注释
'说明：主要是RunSQLFile的子函数
    Dim blnStr As Boolean
    Dim i As Long, k As Long
    
    If Left(strSQL, 2) <> "--" And InStr(strSQL, "--") > 0 Then
        For i = 1 To Len(strSQL)
            If Mid(strSQL, i, 1) = "'" Then blnStr = Not blnStr
            If Mid(strSQL, i, 2) = "--" And Not blnStr Then
                k = i: Exit For
            End If
        Next
        If k > 0 Then strSQL = RTrim(Left(strSQL, k - 1))
    End If
    TrimComment = strSQL
End Function

Public Function SplitSQL(ByVal strSQL As String) As String
'功能：取";"结尾前面的的SQL语句,可能";"号后有"--"注释。
'说明：主要是RunSQLFile的子函数
    Dim i As Long, k As Long
    
    '先去掉注释部份
    strSQL = TrimComment(strSQL)
    
    For i = Len(strSQL) To 1 Step -1
        If Mid(strSQL, i, 1) = ";" Then
            k = i: Exit For
        End If
    Next
    If k > 0 Then strSQL = Left(strSQL, k - 1)
    
    SplitSQL = strSQL
End Function

Public Function RemoveMark(ByVal strText As String) As String
'功能：去除一段文字中的前导"--"注释标记
    Dim arrText As Variant, strTemp As String, i As Long
    
    arrText = Split(strText, vbCrLf)
    
    strText = ""
    For i = 0 To UBound(arrText)
        strTemp = arrText(i)
        If Left(strTemp, 2) = "--" And Replace(strTemp, "-", "") <> "" Then
            strText = strText & vbCrLf & Mid(strTemp, 3)
        End If
    Next
    RemoveMark = Mid(strText, 3)
End Function

Public Function GetLogSQL(objSQL As clsSQLInfo) As String
'功能：获取简要SQL语句，用于填写日志
    Dim strSQL As String
    
    If objSQL.Block Then
        If objSQL.BlockName <> "" Then
            strSQL = Trim(Split(objSQL.SQL, vbCrLf)(0))
            If InStr(strSQL, "(") > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(strSQL, "(") - 1))
            End If
            If InStr(1, strSQL, " as", vbTextCompare) > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " as", vbTextCompare) - 1))
            End If
            If InStr(1, strSQL, " is", vbTextCompare) > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " is", vbTextCompare) - 1))
            End If
            If InStr(1, strSQL, " Return", vbTextCompare) > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " Return", vbTextCompare) - 1))
            End If
        Else '匿名块
            strSQL = ActualStr(TrimEx(objSQL.SQL, True), 150)
        End If
    ElseIf UCase(LTrim(objSQL.SQL)) Like "CREATE * VIEW" Then
        '视图特殊处理
        strSQL = Split(objSQL.SQL, vbCrLf)(0)
        If InStr(1, strSQL, " as", vbTextCompare) > 0 Then '视图只能用as
            strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " as", vbTextCompare) - 1))
        End If
    Else
        If InStr(objSQL.SQL, vbCrLf) > 0 Then
            '多行SQL
            strSQL = ActualStr(TrimEx(objSQL.SQL, True), 150)
        Else
            strSQL = ActualStr(objSQL.SQL, 150)
        End If
    End If
    GetLogSQL = strSQL
End Function


Public Function CheckHavHistory(ByVal lngSys As Long) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:检查是否需要创建历史空间
    '参数:lngSys-系统号
    '返回:需要创建,返true,否则False
    '--------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 1 from zltools.zlbakTables where 系统=" & lngSys & " and rownum<=1"
    OpenRecordset rsTemp, gstrSQL, "获取bak数据", , , gcnOracle
    If rsTemp.EOF Then
       '返回False,表示该系统没有历史数据空间,没有要处理历史数据空间
       Exit Function
    End If
    CheckHavHistory = True
End Function

Public Function GrantBakToUser(ByVal cnOracle As ADODB.Connection, ByVal strToOwner As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------------------
    '功能:检查表是否存在
    '参数:strTableName-表名
    '     cnoracle-数据库连接名
    '     strOwNer-所有者
    '返回:存在该表返回true,否则False
    '-----------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    err = 0: On Error GoTo errHand:
    strSQL = "Select TABLE_NAME from user_all_tables" & _
            " Union All Select View_Name From User_Views"
    Call OpenRecordset(rsTemp, strSQL, "重新授权", , , cnOracle)
    With rsTemp
        Do While Not .EOF
            strSQL = "Grant ALL on " & Nvl(!Table_Name) & " to " & strToOwner & " With Grant Option"
            cnOracle.Execute strSQL
            .MoveNext
        Loop
    End With
    GrantBakToUser = True
    Exit Function
errHand:
    If MsgBox("在授权时出现如下错误,请检查!" & vbCrLf & " (" & err.Number & ") " & err.Description, vbRetryCancel + vbDefaultButton1 + vbQuestion, gstrSysName) = vbRetry Then
        Resume
    End If
    GrantBakToUser = False
End Function


Public Function IsNetServer(ByVal strPath As String, ByVal strUser As String, ByVal strPassword As String) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '--功能:检查服务器是否正常并连接
    '--参数:strPath -访问路径
    '       strUser-用户名
    '       strPassWord -访问密码
    '返回:连接顺畅,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/06
    '----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
      
    '刘兴洪:可能存在windows资源管理器已经有访问的了
    '
    If objFile.FolderExists(strPath) Then
        IsNetServer = True: Exit Function
    End If
    
    Dim NetR As NETRESOURCE
    With NetR
        .dwScope = RESOURCE_GLOBALNET
        .dwType = RESOURCETYPE_DISK
        .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
        .dwUsage = RESOURCEUSAGE_CONNECTABLE
        .lpLocalName = "" '映射的驱动器
        .lpRemoteName = strPath  '服务器路径
    End With
    
    err = 0
    On Error GoTo errHand:
    If WNetAddConnection2(NetR, strPassword, strUser, CONNECT_UPDATE_PROFILE) = NO_ERROR Then
       IsNetServer = True
    Else
       IsNetServer = False
    End If
    Exit Function
errHand:
       IsNetServer = False
End Function
Public Function CancelNetServer(ByVal strPath As String) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '功能:断开服务器连接
    '参数:
    '返回:断找成功,返回true,否则返回False
    '----------------------------------------------------------------------------------------------------------
    err = 0
    On Error Resume Next
    If WNetCancelConnection2(strPath, CONNECT_UPDATE_PROFILE, True) = 0 Then
        CancelNetServer = True
    Else
        CancelNetServer = False
    End If
    err = 0
End Function

Public Sub ReGrantForTools(ByVal cnTools As ADODB.Connection, Optional ByVal strSysOwner As String, Optional ByVal rsToolsObjs As ADODB.Recordset, Optional ByVal blnSysGrant As Boolean, Optional ByVal blnALLSysGrant As Boolean)
    '----------------------------------------------------------------------------------------------------------
    '功能:对管理工具的对象进行重新授权并创建同义词
    '参数:cnTools：管理工具连接。strSysOwner为空时，可以传应用系统连接，此时为应用系统转授权限。
    '     strSysOwner:应用系统所有者。为空是服务器创建调用，只创建公共同义词以及对Public授权
    '     rsToolsObjs：管理工具对象记录集，不传时读取数据库
    '     blnSysGrant:系统所有者转授管理工具权限，需加前缀ZLTOOLS.
    '     blnALLSysGrant:对所有系统进行授权,此时strSysOwner入参无效
    '返回:
    '----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset, rsSys As New ADODB.Recordset
    Dim arrObjects As Variant
    Dim i As Long
    Dim strObjectName As String
    Dim arrOwners() As Variant
    
    arrOwners = Array()
    '获取所有者
    If blnALLSysGrant Then
        '管理工具对象重新授权并创建公共同义词
        gstrSQL = "Select Distinct 所有者 FROM zlsystems"
        OpenRecordset rsSys, gstrSQL, "创建公共同义词", , , cnTools
        Do While Not rsSys.EOF
            ReDim Preserve arrOwners(UBound(arrOwners) + 1)
            arrOwners(UBound(arrOwners)) = rsSys!所有者 & ""
            rsSys.MoveNext
        Loop
    ElseIf strSysOwner <> "" Then
        arrOwners = Array(strSysOwner)
    End If
    
    If rsToolsObjs Is Nothing Then
        strSQL = "Select Object_Name, Object_Type" & vbNewLine & _
                "From User_Objects" & vbNewLine & _
                "Where Object_Type In ('FUNCTION', 'PROCEDURE', 'TYPE', 'PACKAGE', 'SEQUENCE', 'TABLE', 'VIEW') And" & vbNewLine & _
                "      Instr(Object_Name, 'BIN$') <= 0"
        Call OpenRecordset(rsTemp, strSQL, "重新授权", , , cnTools)
    Else
        Set rsTemp = rsToolsObjs
    End If
    
    On Error Resume Next
    With rsTemp
        
        '管理工具对象170个左右，通过循环，执行SQL大约900次左右，耗时在2-3秒
        Do While Not .EOF
            '对应用系统所有者授予管理工具权限
            If blnSysGrant Then
                strObjectName = "ZLTOOLS." & !OBJECT_NAME
            Else
                strObjectName = !OBJECT_NAME & ""
            End If
            
            For i = 0 To UBound(arrOwners)
                Select Case !OBJECT_TYPE
                    Case "FUNCTION", "PROCEDURE", "TYPE", "PACKAGE"
                        strSQL = "grant execute on " & strObjectName & " to " & arrOwners(i) & " With GRANT Option"
                    Case "VIEW"
                        strSQL = "grant select on " & strObjectName & " to " & arrOwners(i) & " With GRANT Option"
                    Case "SEQUENCE"
                        strSQL = "grant select,alter on " & strObjectName & " to " & arrOwners(i) & " With GRANT Option"
                    Case "TABLE"
                        strSQL = "grant select,insert,update,delete on " & strObjectName & " to " & arrOwners(i) & " With GRANT Option"
                End Select
                cnTools.Execute strSQL
            Next
            '同义词修正，先删除同义词，再重新创建
            cnTools.Execute "drop synonym " & !OBJECT_NAME: err.Clear
            cnTools.Execute "drop public synonym " & !OBJECT_NAME
            cnTools.Execute "create public synonym " & !OBJECT_NAME & " for " & strObjectName
            '将对象权限授予PUBLIC
            Select Case !OBJECT_TYPE
                Case "FUNCTION", "PROCEDURE", "TYPE", "PACKAGE"
                    If Not ",B_ROLEGROUPMGR,ZL_ZLROLEGRANT_BATCHDELETE,ZL_ZLROLEGRANT_BATCHINSERT," _
                         Like "*," & UCase(!OBJECT_NAME & "") & ",*" Then
                        strSQL = "grant execute on " & strObjectName & " to Public"
                    End If
                Case "SEQUENCE", "TABLE", "VIEW"
                    strSQL = "grant select on " & strObjectName & " to Public"
            End Select
            cnTools.Execute strSQL
            err.Clear
            .MoveNext
        Loop
    End With
End Sub

Public Function GrantSpecialToRole(ByVal cnOracle As ADODB.Connection, ByVal strRoleNames As String, ByVal blnGrantBase As Boolean, strOwners() As String) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '功能:对管理工具的对象或应用程序一些对象进行授权（特殊的对象）
    '参数:cnOracle：应用系统连接
    '     strRoleNames:被授权的角色，多个角色以逗号分割，一般不超过15个角色
    '     blnGrantBase:是否对应用系统基础表进行授权
    '     strOwners：应用系统所遇者
    '返回:
    '----------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim blnsysSt As String
    Dim strStSysOwner As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select 所有者 From zlSystems Where Floor(编号 / 100) = 1"
    If UBound(strOwners) <> -1 Then
        OpenRecordset rsTmp, strSQL, "获取标准版系统所有者", , , cnOracle
        Do While Not rsTmp.EOF
            strStSysOwner = strStSysOwner & "," & rsTmp!所有者
            rsTmp.MoveNext
        Loop
        If strStSysOwner <> "" Then strStSysOwner = strStSysOwner & ","
    End If
    
    For i = LBound(strOwners) To UBound(strOwners)
        If strOwners(i) <> "" Then
            cnOracle.Execute "grant select on " & strOwners(i) & ".部门表 to " & strRoleNames
            cnOracle.Execute "grant select on " & strOwners(i) & ".人员表 to " & strRoleNames
            cnOracle.Execute "grant select on " & strOwners(i) & ".部门人员 to " & strRoleNames
            cnOracle.Execute "grant select on " & strOwners(i) & ".上机人员表 to " & strRoleNames
            cnOracle.Execute "grant select on " & strOwners(i) & ".人员性质说明 to " & strRoleNames
            cnOracle.Execute "grant select on " & strOwners(i) & ".人员性质分类 to " & strRoleNames
            If InStr(strStSysOwner, "," & strOwners(i) & ",") > 0 Then
                '消息平台对象
                cnOracle.Execute "grant select on " & strOwners(i) & ".业务消息类型 to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".业务消息清单 to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".业务消息提醒部门 to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".业务消息提醒人员 to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".业务消息状态 to " & strRoleNames
                cnOracle.Execute "grant execute on " & strOwners(i) & ".Zlpub_业务消息清单_insert to " & strRoleNames
                cnOracle.Execute "grant execute on " & strOwners(i) & ".Zl_业务消息清单_insert to " & strRoleNames
                cnOracle.Execute "grant execute on " & strOwners(i) & ".Zl_业务消息清单_read to " & strRoleNames
            End If
            If blnGrantBase Then
                cnOracle.Execute "grant execute on " & strOwners(i) & ".zl_字典管理_execute to " & strRoleNames
            End If
        End If
    Next
    '对服务器的几个表进行特殊授权
    '------------------------------------------------------------------------------------------------------------------
    cnOracle.Execute "grant insert,update         on ZLTOOLS.zlDiaryLog to " & strRoleNames
    cnOracle.Execute "grant insert                on ZLTOOLS.zlErrorLog to " & strRoleNames
    cnOracle.Execute "grant update,delete         on ZLTOOLS.zlMessages to " & strRoleNames
    cnOracle.Execute "grant update,delete         on ZLTOOLS.zlMsgState to " & strRoleNames
    cnOracle.Execute "grant insert,update,delete  on ZLTOOLS.zlClientScheme to " & strRoleNames
    cnOracle.Execute "grant insert,update,delete  on ZLTOOLS.zlClientParaSet to " & strRoleNames
    cnOracle.Execute "grant insert,update,delete  on ZLTOOLS.zlClientparaList to " & strRoleNames
    cnOracle.Execute "grant Select on sys.dba_role_privs to " & strRoleNames
    GrantSpecialToRole = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function
