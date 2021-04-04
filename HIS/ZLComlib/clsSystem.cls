VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'##############################################################################
'#功能：ZLHIS系统基础函数功能
'命名规则：函数过程名称为通俗易懂的英文即可
'上级类：clsDatabase
'##############################################################################
Private mobjFIle As New FileSystemObject

Public Function IsMovedByNO(ByVal strTableName As String, ByVal strNO As String, Optional ByVal strIF As String = "记录性质 IN", Optional ByVal strIFValue As String = "", Optional ByVal strFormCaption As String = "ZL9ComLib.NOMoved", Optional ByVal blnOnline As Boolean = False) As Boolean
'功能:根据指定的表名、单据号、记录性质,判断指定的单据号是否已转出到后备数据表中
'参数:NO可以是:F0000007，或:Between 'F0000001' And 'F0000007',或:In('F0000001','F0000002')
'       strIFValue在费用系统中指记录性质,可以用1,2这种形式
'       strIF,strIFValue:除单据号外的附加条件及条件值,默认为费用系统的条件,两个条件必须配对使用
'返回:
'     注意:函数名NOMoved中的NO表示单据号NO,而不是no,不要理解为没有被转出,正确的含义是:指定的单据被转出,则为真,否则为假,出错也返回假.
    IsMovedByNO = gobjComLib.zlDatabase.NOMoved(strTableName, strNO, strIF, strIFValue, strFormCaption, blnOnline)
End Function

Public Function IsMovedByDate(ByVal strDate As String, Optional ByVal int组号 As Integer = 1, Optional ByVal int系统 As Integer = 100, Optional ByVal strFormCaption As String = "ZL9ComLib.DateMoved") As Boolean
'功能:根据指定的日期、组号、系统判断指定日期的数据是否已转出到后备数据表中
     IsMovedByDate = gobjComLib.zlDatabase.DateMoved(strDate, int组号, int系统, strFormCaption)
End Function

Public Function IsMovedByTable(ByVal strTableName As String, ByVal strIF As String, ByVal strIFValue As String, Optional ByVal strFormCaption As String = "ZL9ComLib.NOMoved", Optional ByVal blnOnline As Boolean = False) As Boolean
'功能:根据指定的表名、判断转出到后备数据表中
'       strIF,strIFValue:除单据号外的附加条件及条件值,默认为费用系统的条件,两个条件必须配对使用
'返回:
    IsMovedByTable = gobjComLib.zlDatabase.TableDataMoved(strTableName, strIF, strIFValue, strFormCaption, blnOnline)
End Function

Public Function ReturnMovedClinic(ByVal intPatientID As Integer, ByVal strTimes As String, Optional ByVal bytFlag As Byte = 1, Optional ByVal strFormCaption As String = "ZL9ComLib.ReturnMovedClinic") As Boolean
'功能:抽选返回病人某次门诊住院医疗数据(Oracle过程:Zl_Retu_Clinic)
'参数:  Times,挂号单号或住院主页id
'       Flag,门诊或住院标志:0-门诊,1-住院
    ReturnMovedClinic = gobjComLib.zlDatabase.ReturnMovedClinic(intPatientID, strTimes, bytFlag, strFormCaption)
End Function

Public Function ReturnMovedExes(ByVal strNO As String, Optional ByVal bytType As Byte = 1, Optional ByVal strFormCaption As String = "ZL9ComLib.ReturnMovedExes") As Boolean
'功能:抽选返回指定费用单据号的数据(Oracle过程:Zl_Retu_Exes)
'参数:NO单据号,bytType表示单据类型,值::1-收费,2-记帐,3-自动记帐,4-挂号,5-就诊卡,6-预交,7-结帐；
    ReturnMovedExes = gobjComLib.zlDatabase.ReturnMovedExes(strNO, bytType, strFormCaption)
End Function

Public Function ReadLob(ByVal lngSys As Long, ByVal Action As Long, ByVal KeyWord As String, _
                        Optional ByVal strFile As String, Optional ByVal bytFunc As Byte = 0, _
                        Optional bytMoved As Byte = 0) As String
'功能：将指定的LOB字段复制为临时文件
'参数：
'lngSys:系统编号
'Action:操作类型（用以区别是操作哪个表）
'---系统100,Zl_Lob_Read
'0-病历标记图形;1-病历文件格式;2-病历文件图形;3-病历范文格式;4-病历范文图形;
'5-电子病历格式;6-电子病历图形;7-病历页面格式(图形)；8-电子病历附件;9-体温重叠标记
'10-临床路径文件,11-临床路径图标;14-人员证书记录;15-人员表;16-人员照片;
'17-药品规格(使用说明);18-药品规格(图片);
'19-部门扩展信息;20-人员扩展信息;22-医嘱报告内容;23-供应商照片;24-自定义申请单文件;25-医嘱申请单文件
'26-门诊路径文件,27-病人照片,28-咨询图片元素,29-咨询段落目录
'--系统600，ZL6_Lob_Read
'0-设备照片
'---系统2400,Zl24_Lob_Read
'手麻常用图形,无Action
'---系统2100,Zl21_Lob_Read
'1-体质类型调养;2-体检体辨结论(该图片只有读取，没有保存);3-体检申报记录;4-体检任务人员,5-体检任务结果
'---系统2500,ZL25_Lob_Read
'0-微生物涂片报告
'---系统2600,Zl26_Lob_Read
'14-导诊控件目录,15-导诊资源目录
'      KeyWord:确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'      strFile:用户指定存放的文件名；不指定时，自动取临时文件名
'bytFunc-0-BLOB,1-CLOB
'bytMoved=0正常记录,1读取转储后备表记录
'返回：存放内容的文件名，失败则返回零长度""
    Const conChunkSize As Long = 10240
    Dim rsLOB       As ADODB.Recordset
    Dim lngFileNum  As Long, lngCount       As Long, lngBound       As Long
    Dim aryChunk()  As Byte, strText        As String
    Dim strSQL      As String
    Dim objFile     As New FileSystemObject
    Dim lngCurSize  As Long
    
    Err = 0: On Error GoTo Errhand
    Select Case lngSys \ 100
        Case 1
            strSQL = "Select Zl_Lob_Read([1],[2],[3],[4],[5]) as 片段 From Dual"
        Case 6
            strSQL = "Select Zl6_Lob_Read([1],[2],[3],[4],[5]) as 片段 From Dual"
        Case 24
            strSQL = "Select Zl24_Lob_Read([2],[3]) as 片段 From Dual"
        Case 21
            strSQL = "Select Zl21_Lob_Read([1],[2],[3]) as 片段 From Dual"
        Case 25
            strSQL = "Select Zl25_Lob_Read([1],[2],[3],[4],[5]) as 片段 From Dual"
        Case 26
            strSQL = "Select Zl26_Lob_Read([1],[2],[3]) as 片段 From Dual"
    End Select
    If strSQL = "" Then strFile = "": Exit Function
    If bytFunc = 0 Then 'BLOB
        If strFile = "" Then
            strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\" & objFile.GetTempName
        End If
        lngFileNum = FreeFile
        Open strFile For Binary As lngFileNum
        lngCount = 0
        lngCurSize = 0
        Do
            Set rsLOB = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "zllobRead", Action, KeyWord, lngCount, bytMoved, bytFunc)
            If rsLOB.EOF Then Exit Do
            If IsNull(rsLOB.Fields(0).value) Then Exit Do
            strText = rsLOB.Fields(0).value
            If lngCurSize = 0 Then
                lngCurSize = Len(strText) / 2
                If lngCurSize = 0 Then Exit Do
                ReDim aryChunk(lngCurSize - 1) As Byte
            ElseIf lngCurSize <> Len(strText) / 2 Then '防止重复分配内存
                lngCurSize = Len(strText) / 2
                If lngCurSize = 0 Then Exit Do
                ReDim aryChunk(lngCurSize - 1) As Byte
            End If
            For lngBound = LBound(aryChunk) To UBound(aryChunk)
                aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
            Next
            Put lngFileNum, , aryChunk()
            lngCount = lngCount + 1
        Loop
        Close lngFileNum
        If lngCount = 0 Then Kill strFile: strFile = ""
    Else  'CLOB
        lngCount = 0
        strFile = ""
        Do
            Set rsLOB = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "zllobRead", Action, KeyWord, lngCount, bytMoved, bytFunc)
            If rsLOB.EOF Then Exit Do
            If IsNull(rsLOB.Fields(0).value) Then Exit Do
            strText = rsLOB.Fields(0).value
            strFile = strFile & strText
            lngCount = lngCount + 1
        Loop
    End If
    ReadLob = strFile
    Exit Function
Errhand:
    If bytFunc = 0 Then
        Close lngFileNum
        Kill strFile: ReadLob = ""
    End If
    Err.Clear
End Function

Public Function SaveLob(ByVal lngSys As Long, ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String, Optional ByVal bytFunc As Byte = 0) As Boolean
'功能：产生保存指定的文件到指定表记录BLOB/CLOB字段的SQL语句
'参数：
'lngSys:系统编号
'Action:操作类型（用以区别是操作哪个表）
'---系统100,Zl_Lob_Append
'0-病历标记图形;1-病历文件格式;2-病历文件图形;3-病历范文格式;4-病历范文图形;
'5-电子病历格式;6-电子病历图形;7-病历页面格式(图形)；8-电子病历附件;9-体温重叠标记
'10-临床路径文件,11-临床路径图标;14-人员证书记录;15-人员表;16-人员照片;
'17-药品规格(使用说明);18-药品规格(图片);
'19-部门扩展信息;20-人员扩展信息;22-医嘱报告内容;23-供应商照片;24-自定义申请单文件;25-医嘱申请单文件
'26-门诊路径文件,27-病人照片,28-咨询图片元素,29-咨询段落目录
'--系统600，ZL6_Lob_Append
'0-设备照片
'---系统2400,Zl24_Lob_Append
'手麻常用图形,无Action
'---系统2100,Zl21_Lob_Append
'1-体质类型调养;2-体检体辨结论(该图片只有读取，没有保存);3-体检申报记录;4-体检任务人员,5-体检任务结果
'---系统2500,ZL25_Lob_Append
'0-微生物涂片报告
'---系统2600,Zl26_Lob_Append
'14-导诊控件目录,15-导诊资源目录
'      KeyWord:确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'      strFile:用户指定存放的文件名；不指定时，自动取临时文件名
'      KeyWord:确定数据记录的关键字，复合关键字以逗号分隔
'      strFile:BLOB时,用户指定要存放的文件名;CLOB时,需要存储的文本内容
'      arySql():在该数据的基础上扩展增加保存的SQL语句
'      bytFunc=0-BLOB,1-CLOB
'返回：成功返回True，失败返回False
    Dim arrSQL() As String
    Dim i As Long
    
    If GetLobSql(lngSys, Action, KeyWord, strFile, arrSQL, bytFunc) Then
        Call gobjComLib.zlDatabase.ExecuteProcedureBeach(arrSQL, "zllobSave", False, False)
    Else
        SaveLob = False
    End If
    SaveLob = True
    Exit Function
Errhand:
    Err.Clear
    SaveLob = False
End Function

Public Function GetLobSql(ByVal lngSys As Long, ByVal Action As Long, ByVal KeyWord As String, _
                ByVal strFile As String, ByRef arySql() As String, Optional ByVal bytFunc As Byte = 0) As Boolean
'功能：产生保存指定的文件到指定表记录BLOB/CLOB字段的SQL语句
'参数：
'lngSys:系统编号
'Action:操作类型（用以区别是操作哪个表）
'---系统100,Zl_Lob_Append
'0-病历标记图形;1-病历文件格式;2-病历文件图形;3-病历范文格式;4-病历范文图形;
'5-电子病历格式;6-电子病历图形;7-病历页面格式(图形)；8-电子病历附件;9-体温重叠标记
'10-临床路径文件,11-临床路径图标;14-人员证书记录;15-人员表;16-人员照片;
'17-药品规格(使用说明);18-药品规格(图片);
'19-部门扩展信息;20-人员扩展信息;22-医嘱报告内容;23-供应商照片;24-自定义申请单文件;25-医嘱申请单文件
'26-门诊路径文件,27-病人照片,28-咨询图片元素,29-咨询段落目录
'--系统600，ZL6_Lob_Append
'0-设备照片
'---系统2400,Zl24_Lob_Append
'手麻常用图形,无Action
'---系统2100,Zl21_Lob_Append
'1-体质类型调养;2-体检体辨结论(该图片只有读取，没有保存);3-体检申报记录;4-体检任务人员,5-体检任务结果
'---系统2500,ZL25_Lob_Append
'0-微生物涂片报告
'---系统2600,Zl26_Lob_Append
'14-导诊控件目录,15-导诊资源目录
'      KeyWord:确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'      strFile:用户指定存放的文件名；不指定时，自动取临时文件名
'      KeyWord:确定数据记录的关键字，复合关键字以逗号分隔
'      strFile:BLOB时,用户指定要存放的文件名;CLOB时,需要存储的文本内容
'      arySql():在该数据的基础上扩展增加保存的SQL语句
'      bytFunc=0-BLOB,1-CLOB
'返回：成功返回True，失败返回False
    Const conChunkSize      As Long = 2000
    Dim lngFileSize         As Long, lngCurSize         As Long, lngModSize     As Long
    Dim lngBlocks           As Long, lngFileNum         As Long
    Dim lngCount            As Long, lngBound           As Long
    Dim aryChunk()          As Byte, aryHex()           As String, strText      As String
    Dim strSQLRoot          As String
    Dim lngLBound           As Long, lngUbound          As Long    '传入数组的最小最大下标
    Dim strSubTxt           As String
    
    Err = 0: On Error Resume Next
    lngLBound = LBound(arySql): lngUbound = UBound(arySql)
    If Err <> 0 Then lngLBound = 0: lngUbound = -1
    Err = 0: On Error GoTo 0
    Select Case lngSys \ 100
        Case 1
            strSQLRoot = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "'"
        Case 6
            strSQLRoot = "Zl6_Lob_Append(" & Action & ",'" & KeyWord & "'"
        Case 24
            strSQLRoot = "Zl24_Lob_Append('" & KeyWord & "'"
        Case 21
            strSQLRoot = "Zl21_Lob_Append(" & Action & ",'" & KeyWord & "'"
        Case 25
            strSQLRoot = "Zl25_Lob_Append(" & Action & ",'" & KeyWord & "'"
        Case 26
            strSQLRoot = "Zl26_Lob_Append(" & Action & ",'" & KeyWord & "'"
    End Select
    If strSQLRoot = "" Then GetLobSql = False: Exit Function
    If bytFunc = 0 Then  'BLOB
        lngFileNum = FreeFile
        Open strFile For Binary Access Read As lngFileNum
        lngFileSize = LOF(lngFileNum)
        If lngFileSize <> 0 Then
            Err = 0: On Error GoTo Errhand
            lngModSize = lngFileSize Mod conChunkSize
            lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
            ReDim Preserve arySql(lngLBound To lngUbound + lngBlocks + 1)
            For lngCount = 0 To lngBlocks
                If lngCount = lngFileSize \ conChunkSize Then
                    lngCurSize = lngModSize
                    ReDim aryChunk(lngCurSize - 1) As Byte
                    ReDim aryHex(lngCurSize - 1) As String
                Else
                    lngCurSize = conChunkSize
                    If lngCount = 0 Then '防止不停分配内存
                        ReDim aryChunk(lngCurSize - 1) As Byte
                        ReDim aryHex(lngCurSize - 1) As String
                    End If
                End If
                Get lngFileNum, , aryChunk()
                For lngBound = LBound(aryChunk) To UBound(aryChunk)
                    aryHex(lngBound) = Hex(aryChunk(lngBound))
                    If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
                Next
                strText = Join(aryHex, "")
                arySql(lngUbound + lngCount + 1) = strSQLRoot & ",'" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
            Next
        Else
            ReDim Preserve arySql(lngLBound To lngUbound + 1)
            arySql(lngUbound + 1) = strSQLRoot & ",NULL,1)"
        End If
        Close lngFileNum
    Else  'CLOB
        strText = strFile
        lngCount = 0
        Do
            strSubTxt = Left(strText, conChunkSize)
            strText = Mid(strText, conChunkSize + 1)
            ReDim Preserve arySql(lngCount)
            arySql(lngCount) = strSQLRoot & ",'" & strSubTxt & "'," & IIf(lngCount = 0, 1, 0) & ",1)"
            lngCount = lngCount + 1
        Loop While Len(strText) > 0
    End If
    GetLobSql = True
    Exit Function
Errhand:
    Err.Clear
    Close lngFileNum
    GetLobSql = False
End Function

Public Function ReadLobV2(ByVal strTable As String, ByVal strLOBField As String, ByVal strWhere As String, ByVal strFile As String, ParamArray arrInput() As Variant) As String
'功能：将指定的LOB字段复制为临时文件(BLOB）或文本内容（CLOB）
'警告：该函数不能读取转储后备表记录，因为读取转储后备表可能会用到抽选返回，条件的开放性导致抽选返回的难以处理。
'参数：
'    strTable:LOB读取的表 自动添加别名A, 以供条件中使用
'    strLOBField:LOB读取的字段，只能是单个字段
'    strWhere,Where条件，可以使用绑定变量，和OpenSQLRecord方法的绑定变量一致
'    strFile:用户指定存放的文件名；不指定时，自动取临时文件名。CLOB时存储文本内容
'    arrInput=绑定变量值，和OpenSQLRecord方法的绑定变量值一致
'返回：存放内容的文件名，失败则返回零长度""，对于CLOB失败和NULL都返回零长度""
    Const conChunkSize      As Integer = 10240
    Dim rsLOB               As ADODB.Recordset
    Dim lngFileNum          As Long, lngBound       As Long
    Dim lngFileSize         As Long, lngModSize     As Long, lngBlocks      As Long, lngCount As Long, lngCurSize As Long
    Dim aryChunk()          As Byte, strText        As String
    Dim strSQL              As String
    Dim objFile             As New FileSystemObject
    Dim arrPars()           As Variant
    Err = 0: On Error GoTo Errhand
    arrPars = arrInput
    strSQL = "Select A." & Trim(strLOBField) & " From " & strTable & " A Where " & strWhere
    Set rsLOB = gobjComLib.zlDatabase.OpenSQLRecordByArray(strSQL, "ReadLobV2", arrPars, 1)
    If Not rsLOB.EOF Then
        If rsLOB.Fields(strLOBField).type = adLongVarWChar Then   'CLOB
            strFile = rsLOB.Fields(strLOBField).value & ""
        ElseIf rsLOB.Fields(strLOBField).type = adLongVarBinary Then 'BLOB
            If strFile = "" Then
                strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\" & objFile.GetTempName
            End If
            lngFileNum = FreeFile
            lngFileSize = 0
            Open strFile For Binary As lngFileNum
            If Not IsNull(rsLOB.Fields(strLOBField).value) Then
                lngFileSize = rsLOB.Fields(strLOBField).ActualSize
                lngModSize = lngFileSize Mod conChunkSize
                lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
                For lngCount = 0 To lngBlocks
                    
                    If lngCount = lngFileSize \ conChunkSize Then
                        lngCurSize = lngModSize
                        ReDim aryChunk(lngCurSize - 1) As Byte
                    Else
                        lngCurSize = conChunkSize
                        If lngCount = 0 Then '防止不停分配内存
                            ReDim aryChunk(lngCurSize - 1) As Byte
                        End If
                    End If
                    aryChunk() = rsLOB.Fields(strLOBField).GetChunk(lngCurSize)
                    Put lngFileNum, , aryChunk()
                Next
            End If
            Close lngFileNum
            If lngFileSize = 0 Then Kill strFile: strFile = ""
            lngFileNum = 0
        End If
    Else
        strFile = ""
    End If
    ReadLobV2 = strFile
    Exit Function
Errhand:
    If lngFileNum <> 0 Then
        Close lngFileNum
        Kill strFile
    End If
    Err.Clear
End Function

Public Function SaveLobV2(ByVal strTable As String, ByVal strLOBField As String, ByVal strWhere As String, ByVal strFile As String, ParamArray arrInput() As Variant) As Boolean
'功能：保存文件或文本到指定表记录BLOB/CLOB
'警告：该保存请放在事务之外处理。在调用该方法时必须保证记录存在，没有记录添加一条LOB字段为空的数据（因为字段数目与数据规则的未知，无法自动处理）
'参数：
'strTable:LOB读取的表 自动添加别名A,以供条件中使用
'strLOBField:LOB读取的字段，只能是单个字段
'strWhere,Where条件，可以使用绑定变量，和OpenSQLRecord方法的绑定变量一致
'strFile:用户指定存放的文件名；不指定时，自动取临时文件名。CLOB为存储的文本
'arrInput=绑定变量值，和OpenSQLRecord方法的绑定变量值一致
'返回：成功返回True，失败返回False
    Const conChunkSize  As Integer = 10240
    Dim lngFileSize     As Long, lngCurSize     As Long, lngModSize     As Long
    Dim lngBolcks       As Long, lngFileNum     As Long, lngCount       As Long
    Dim aryChunk()      As Byte
    Dim arrPars()       As Variant
    Dim strSQL          As String, rsLOB        As ADODB.Recordset
    
    On Error GoTo errh
    arrPars = arrInput
    strSQL = "Select A." & Trim(strLOBField) & " From " & strTable & " A Where " & strWhere
    Set rsLOB = gobjComLib.zlDatabase.OpenSQLRecordByArray(strSQL, "ReadLobV2", arrPars, 2)
    If Not rsLOB.EOF Then
        rsLOB.Fields(strLOBField).value = Null
        If rsLOB.Fields(strLOBField).type = adLongVarWChar Then   'CLOB
            rsLOB.Fields(strLOBField).value = strFile
        ElseIf rsLOB.Fields(strLOBField).type = adLongVarBinary Then 'BLOB
            lngFileNum = FreeFile
            Open strFile For Binary Access Read As lngFileNum
            lngFileSize = LOF(lngFileNum)
            If lngFileSize <> 0 Then
                lngModSize = lngFileSize Mod conChunkSize
                lngBolcks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
                For lngCount = 0 To lngBolcks
                    If lngCount = lngFileSize \ conChunkSize Then
                        lngCurSize = lngModSize
                        ReDim aryChunk(lngCurSize - 1) As Byte
                    Else
                        lngCurSize = conChunkSize
                        If lngCount = 0 Then
                            ReDim aryChunk(lngCurSize - 1) As Byte
                        End If
                    End If
                    Get lngFileNum, , aryChunk()
                    rsLOB.Fields(strLOBField).AppendChunk aryChunk()
                Next
            End If
            Close lngFileNum
        End If
        rsLOB.Update
    End If
    SaveLobV2 = True
    Exit Function
errh:
    If Not rsLOB Is Nothing Then rsLOB.CancelUpdate
    If lngFileNum <> 0 Then
        Close lngFileNum
    End If
    Err.Clear
End Function


Public Function ReadXML(ByVal strTable As String, ByVal strXMLColumn As String, ByVal strWhereFilter As String, ByRef strErrInfo As String, ParamArray arrInput() As Variant) As String
'功能：读取XML字段并返回字符串
'参数：
'strTable:XML字段所在的表
'strXMLColumn:XML字段名
'strWhereFilter:Where条件，条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'               因为使用绑定变量,对带"'"的字符参数,不需要使用"''"形式。
'strErrInfo=返回的错误信息，
'           多行记录时，返回错误。
'返回：获取的XML字段
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim arrPars() As Variant, intLoop As Long
    Dim lngLength As Long
    
    On Error GoTo errh
    ReadXML = ""
    If strWhereFilter <> "" Then strWhereFilter = " Where " & strWhereFilter
    arrPars = arrInput
    strSQL = "Select length(" & strXMLColumn & ".getclobval()) As ValueLength" & vbNewLine & _
            "  From " & strTable & strWhereFilter
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecordByArray(strSQL, "获取XML字段值", arrPars)
    If rsTmp.RecordCount = 1 Then
        lngLength = Val(rsTmp!ValueLength & "")
        If lngLength > 2000 Then
            For intLoop = 1 To lngLength / 2000 + 1
                strSQL = "select To_Char(substr(" & strXMLColumn & ".getclobval(), " & CDbl(intLoop) * 2000 - 1999 & ", 2000)) as PartValue" & vbNewLine & _
                        "  from " & strTable & strWhereFilter
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecordByArray(strSQL, "获取XML字段值", arrPars)
                If Not rsTmp.EOF Then
                    ReadXML = ReadXML & rsTmp!PartValue
                End If
            Next
        Else
            strSQL = "select " & strXMLColumn & ".getstringval() as PartValue" & vbNewLine & _
                    "  from " & strTable & strWhereFilter
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecordByArray(strSQL, "获取XML字段值", arrPars)
            If Not rsTmp.EOF Then
                ReadXML = ReadXML & rsTmp!PartValue
            End If
        End If
    ElseIf rsTmp.RecordCount > 1 Then '传出错误，XML字段只能单条返回
        strErrInfo = "不允许返回多条数据"
    End If
    Exit Function
errh:
    strErrInfo = Err.Description
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function


Public Function RowValue(ByVal strTable As String, Optional ByVal arrValues As Variant, Optional ByVal strGetFields As String, Optional ByVal strWhereField As String = "ID") As Variant
'功能：获取指定表指定字段信息
'参数：strTable=读取数据的表
'          arrValues=过滤值，可以传数组，也可以传单个值，也可以不传，不传读取全表
'          strGetField=获取的字段,多个字段以逗号分割，同SQL书写获取字段一致
'          strWhereField=过滤的字段，该字段为简单的数值或字符类型或日期类型，其余类型无法支持
'返回：
'只有以下一种情况返回特定的一个值（未处理NULL值）：
'      strGetField=单个字段
'      arrValues=为单个值，或不高于一个元素的数组
'其余情况均返回记录集

    Dim rsTmp As New ADODB.Recordset, blnReturnRec As Boolean
    Dim strSQL As String
    Dim strWhere As String
    Dim arrPars As Variant
    Dim i As Long, strPars As String
    
    On Error GoTo errh
    blnReturnRec = True
    If TypeName(arrValues) = "Variant()" Then
        arrPars = arrValues
        For i = LBound(arrValues) To UBound(arrValues)
            strPars = strPars & ",[" & i + 1 & "]"
        Next
        If strGetFields <> "" Then '数组顾虑元素不超过一个,且获取单个元素，则不返回记录集
            If UBound(arrValues) - LBound(arrValues) + 1 <= 1 And Not strGetFields Like "*,*" Then blnReturnRec = False
        End If
        If strPars <> "" Then
            strWhere = " Where " & strWhereField & " In (" & strPars & ")"
        End If
    ElseIf TypeName(arrValues) <> "Error" Then
        '单个值时，若获取单个字段，则不反悔记录集
         If strGetFields <> "" And Not strGetFields Like "*,*" Then blnReturnRec = False
        arrPars = Array(arrValues)
        strWhere = " Where " & strWhereField & "=[1]"
    Else
        strWhere = ""
    End If
    
    If strGetFields = "" Then strGetFields = "*"
    strSQL = "Select " & strGetFields & " From " & strTable & strWhere
    If strWhere <> "" Then
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "RowValue", arrPars)
    Else
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "RowValue")
    End If
    If blnReturnRec Then
        Set RowValue = rsTmp
    Else
        If Not rsTmp.EOF Then
            RowValue = rsTmp.Fields(strGetFields).value
        Else '获取单值时，没有获取到数值，则返回默认值
            If IsType(rsTmp.Fields(strGetFields).type, adVarChar) Then
                RowValue = ""
            ElseIf IsType(rsTmp.Fields(strGetFields).type, adInteger) Then
                RowValue = 0
            Else
                RowValue = Null
            End If
        End If
    End If
    Exit Function
errh:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function IsZLHIS10() As Boolean
'功能：判断标准版是否是10版本的系统
    IsZLHIS10 = gobjComLib.zlDatabase.IsZLHIS10
End Function

Public Function GetUserInfo() As ADODB.Recordset
'功能：获取当前用户的基本信息
'返回：返回Ado记录集
    Set GetUserInfo = gobjComLib.zlDatabase.GetUserInfo
End Function

Public Function Currentdate() As Date
'功能：提取服务器上当前日期
'参数：
'返回：由于Oracle日期格式的问题，所以
    Currentdate = gobjComLib.zlDatabase.Currentdate
End Function

Public Function IP(Optional ByVal strErr As String) As String
    '******************************************************************************************************************
    '功能:通过oracle获取的计算机的IP地址
    '入参:strDefaultIp_Address-缺省IP地址
    '出参:
    '返回:返回IP地址
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim strIp_Address As String
    Dim strSQL As String
        
    On Error GoTo Errhand
    
    strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "获取IP地址")
    If rsTmp.EOF = False Then
        strIp_Address = gobjComLib.NVL(rsTmp!Ip_Address)
    End If
    If strIp_Address = "" Then strIp_Address = gobjComLib.OS.IP(strErr)
    If Replace(strIp_Address, " ", "") = "0.0.0.0" Then strIp_Address = ""
    IP = strIp_Address
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    strErr = strErr & IIf(strErr = "", "", "|") & Err.Description
    Err.Clear
End Function

Public Function GetNextNo(ByVal int序号 As Integer, Optional ByVal lng科室ID As Long, Optional ByVal strTag As String, Optional ByVal intStep As Integer = 1) As Variant
    GetNextNo = gobjComLib.zlDatabase.GetNextNo(int序号, lng科室ID, strTag, intStep)
End Function

Public Function NextId(strTable As String) As Long
'功能：读取指定表名对应的序列(按规范，其序列名称为“表名称_id”)的下一数值
'参数：
'   strTable：表名称
'返回：
    NextId = gobjComLib.zlDatabase.GetNextId(strTable)
End Function

Public Function MaxCode(ByVal strTable As String, ByVal strField As String, ByVal intLength As Integer, Optional ByVal strWhere As String) As String
'功能：读取指定表的本级编码的最大值
'参数：strTable  表名;
'      strField  字段名;
'      intLength 字段长度
'返回：成功返回 下级最大编码; 否者返回 0
    MaxCode = gobjComLib.zlDatabase.GetMax(strTable, strField, intLength, strWhere)
End Function

Public Function DeptHaveProperty(ByVal lng科室ID As Long, ByVal str性质 As String, Optional ByRef blnOutDept As Boolean, Optional ByVal blnClearCache As Boolean) As Boolean
'功能：检查指定科室是否具有指定工作性质
'说明：因为部门性质一般不变动，又大量使用，利用缓存读取
'        blnClearCache=是否清除缓存
'返回：blnOutDept=是否为仅服务于门诊的部门
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, blnNew As Boolean
    Dim rsDept As ADODB.Recordset
    blnOutDept = False
    
    If blnClearCache Then Set rsTmp = Nothing '清空缓存
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    strSQL = "Select 部门ID,工作性质,服务对象 From 部门性质说明 where 部门ID=[1]"
    On Error GoTo errh
    
    '获取数据源
    If blnNew Then
        Set rsTmp = gobjComLib.Rec.CopyNew(gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "DeptHaveProperty", lng科室ID))
    Else
        rsTmp.Filter = "部门ID=" & lng科室ID
        If rsTmp.EOF Then
            Call gobjComLib.Rec.Append(rsTmp, gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "DeptHaveProperty", lng科室ID))
        End If
    End If
    
    '性质判断
    rsTmp.Filter = "部门ID=" & lng科室ID & " And 工作性质='" & str性质 & "'"
    DeptHaveProperty = Not rsTmp.EOF
    If rsTmp.RecordCount > 0 Then
        rsTmp.Filter = "部门ID=" & lng科室ID & " And 工作性质='" & str性质 & "' And 服务对象<>1"
        blnOutDept = rsTmp.RecordCount = 0
    End If
    Exit Function
errh:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function SystemVersion(Optional ByVal lngSys As Long = 100) As String
'功能：提取指定系统版本号，默认为医院系统标准版
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo Errhand
    
    strSQL = " select 版本号 from zlsystems where 编号=[1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "提取系统版本号", lngSys)
    SystemVersion = rsTemp!版本号
    Exit Function
Errhand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function SystemOwner(Optional ByVal lngSys As Long = 100) As String
'功能：获取指定系统的所有者
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errh
    strSQL = "Select 所有者 From zlSystems Where 编号=[1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "SystemOwner", lngSys)
    If Not rsTmp.EOF Then
        SystemOwner = rsTmp!所有者
    End If
    Exit Function
errh:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function SystemShareWith(ByVal lngSysShareNO As Long, Optional ByVal lngSysMainNO As Long = 100) As Boolean
'功能：标准系统和其他系统是否是共享安装
'参数：lngSysShareNO= 共享安装的系统
'           lngSysMainNO=主系统
    Dim lngShareNum As Long
    Dim strSQL As String
    Dim rsTmp As Recordset
'Select * From (Select * From zlSystems Start With 编号 = 100 Connect By Prior 编号 = 共享号) Where 编号 = 300
'Select * From (Select * From zlSystems Start With 编号 = 300 Connect By Prior 编号 = 共享号) Where 编号 = 100
    strSQL = "Select s.编号" & vbNewLine & _
            "From zlSystems S" & vbNewLine & _
            "Where s.正常安装 = 1 And s.编号  = [1] And s.共享号 = [2]"
    On Error GoTo errh
    '由于存在多帐套情况若标准版多帐套编号100，101，。。。。，199，因此若此判断
    '多张套不能共享安装
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, gstrSysName, (lngSysShareNO \ 100) * 100, (lngSysMainNO \ 100) * 100)
    SystemShareWith = rsTmp.RecordCount > 0
    Exit Function
errh:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function IsSysSetUp(ByVal lngSys As Long) As Boolean
'功能：是否安装某个系统
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    
    strSQL = "Select 1 From zlSystems Where 编号=[1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "IsSysSetUp", lngSys)
    IsSysSetUp = Not rsTmp.EOF
    Exit Function
errh:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function FieldsLength(strTable As String, strColumns As String) As Variant
'功能：获取一个或多个字段定义长度
'参数：strTable=表名
'          strColumns=列名，多列之间以逗号分割
'返回：单列直接返回列长度，多列返回长度以逗号分割,若某一列未获取到，则返回-1
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim arrTmp As Variant, i As Long
    Dim strReturn As String
    strSQL = "Select Column_Name,Nvl(Data_Precision, Data_Length) Collen ,Decode(Owner, User, 1, 2) Sort" & vbNewLine & _
                "From All_Tab_Columns" & vbNewLine & _
                "Where Table_Name = [1] And Column_Name In(Select Column_Value From Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)))"
    On Error GoTo errh
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "FieldsLength", strTable, UCase(strColumns))
    arrTmp = Split(UCase(strColumns), ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        rsTmp.Filter = "Column_Name='" & arrTmp(i) & "'": rsTmp.Sort = "Sort"
        If Not rsTmp.EOF Then
            strReturn = strReturn & "," & Val("" & rsTmp!collen)
        Else
            strReturn = strReturn & ",-1"
        End If
    Next
    If strReturn <> "" Then strReturn = Mid(strReturn, 2)
    If strReturn Like "*,*" Then
        FieldsLength = strReturn
    Else
        FieldsLength = Val(strReturn)
    End If
    Exit Function
errh:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function GetPrivFuncByUser(lngSys As Long, lngProgId As Long, ByVal strDBUser As String) As String
'功能：返回用户具有的指定程序的功能串,传登录用户
'参数：lngSys     如果是固定模块，则为0
'      lngProgId  程序序号
'      strDBUser=数据库用户
'返回：分号间隔的功能串,为空表示没有权限

    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strTmp As String
    
    On Error GoTo errh
    '当前用户特殊处理，直接调用gobjRegister.zlRegFunc
    If UCase(strDBUser) = UCase(gstrDBUser) Then
        GetPrivFuncByUser = gobjRegister.zlRegFunc(lngSys, lngProgId)
        Exit Function
    End If
    '非当前用户，则读取两张表，简单返回
    '角色授权时只可以授权ZlregFunc中存在的功能权限，因此可以不关联。
    '报表发布到模块内作为模块功能，发布到导航台作为单独模块，他们的授权均在Zlrolegrant中。
    strSQL = "Select Distinct b.功能" & vbNewLine & _
            "From Zluserroles a, Zlrolegrant b" & vbNewLine & _
            "Where a.用户 = [3] And a.角色 = b.角色 And b.系统 = [1] And b.序号 = [2]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "获取其他人员功能权限", lngSys, lngProgId, strDBUser)
    Do While Not rsTmp.EOF
        strTmp = strTmp & ";" & rsTmp!功能
        rsTmp.MoveNext
    Loop
    strTmp = Mid(strTmp, 2)
    
    GetPrivFuncByUser = strTmp
    Exit Function
errh:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function CheckRushHours(ByVal lngSysNo As Long, ByVal lngModuleNo As Long, ByVal strFuncName As String) As Boolean
'功能：检查当前时间是否是业务高峰期
'参数：
'      lngSysNo=系统号
'      lngModuleNo=模块号
'      strFuncName=功能名称
'返回：是否可以进行操作

    Dim strSQL As String
    Dim strTime As String
    Dim rsTemp As ADODB.Recordset
    Dim blnLimit As Boolean
    Dim dateNow As Date
    Dim strNote As String
    
    On Error GoTo errh
    CheckRushHours = True
    dateNow = CDate(Format(gobjComLib.zlDatabase.Currentdate(), "HH:MM:SS"))
    strSQL = "Select a.操作选项, a.限时原因, To_Char(b.开始时间, 'HH24:MI:SS') 开始时间, To_Char(b.结束时间, 'HH24:MI:SS') 结束时间" & vbNewLine & _
            "From Zlrunlimitset A, Zlrunlimittime B" & vbNewLine & _
            "Where a.方案序号 = b.方案 And a.系统 = [1] And a.模块 = [2] And a.功能 = [3] And b.星期 = To_Char(Sysdate, 'd') - 1" & vbNewLine & _
            "Order By b.开始时间"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(gcnOracle, strSQL, "获取指定功能当天的限时时间数据", lngSysNo, lngModuleNo, strFuncName)
    With rsTemp
        If .RecordCount = 0 Then Exit Function
        If gobjComLib.NVL(!限时原因) = "" Then
            strNote = "当前时段处于业务高峰期，使用此功能可能会对系统使用造成一定影响"
        Else
            strNote = !限时原因
        End If
        Do While Not .EOF
            strTime = strTime & "，" & !开始时间 & "-" & !结束时间
            If dateNow > !开始时间 And dateNow < !结束时间 Then
                '说明当前时间在限制时间的范围里
                blnLimit = True
            End If
            .MoveNext
        Loop
        If blnLimit = True Then
            .MoveFirst
            If !操作选项 = 0 Then  '弹出提示，并禁止用户进行后续操作
                MsgBox strNote & "，" & vbNewLine & "故在时间范围：" & vbNewLine & Mid(strTime, 2) & vbNewLine & "内禁止使用此功能！", vbInformation, gstrSysName
            Else   '弹出提示，但不禁止用户进行后续操作
                If MsgBox(strNote & "，" & vbNewLine & "确定要继续吗？", vbInformation + vbOKCancel, gstrSysName) = vbOK Then
                    blnLimit = False
                Else
                    blnLimit = True
                End If
            End If
        End If
    End With
    CheckRushHours = Not blnLimit
    Exit Function
errh:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function CheckAuditStatus(ByVal objParent As Object, ByVal lngSysNo As Long, ByVal lngModuleNo As Long, ByVal strFuncName As String, ByRef strRemarks As String) As Boolean
    '功能：检查传入的功能是否需要进行审核
    'lngSysNo=系统编号
    'lngModuleNo = 模块编号
    'strFuncName = 功能名称
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errh
    strSQL = "Select 是否需审核 From Zlauditlogconfig Where 系统 = [1] And 模块 = [2] And 功能 = [3]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "获取对应模块功能是否需要审核", lngSysNo, lngModuleNo, strFuncName)
    With rsTemp
        If .RecordCount > 0 Then
            If !是否需审核 = 1 Then
                If Not frmCheckUserIdentify.ShowMe(objParent, lngSysNo, strRemarks) Then Exit Function
            Else
                strRemarks = ""
            End If
        Else
            Exit Function
        End If
    End With
    CheckAuditStatus = True
    Exit Function
errh:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Sub SaveAuditLog(ByVal lngType As Long, ByVal lngSysNo As Long, ByVal lngModuleNo As Long, ByVal strFuncName As String, ByVal strContent As String, Optional ByVal strDescription As String)
'功能：插入重要工作日志
'参数：lngType = 操作类型，1-新增，2-修改，3-删除
'      lngSysNo = 系统编号
'      lngModuleNo = 模块编号
'      strFuncName = 功能名称
'      strContent = 操作内容
'      strDescription = 操作说明
    Dim strSQL As String
    
    On Error GoTo errh:
    If LenB(StrConv(strContent, vbFromUnicode)) > 1024 Then
        strContent = Mid(strContent, 1, 500)
    End If
    strSQL = "zltools.Zl_Zlauditlog_Insert('" & gstrDBUser & "','" & _
                                                gstrComputerName & "'," & _
                                                lngType & "," & _
                                                lngSysNo & ",'" & _
                                                lngModuleNo & "','" & _
                                                strFuncName & "','" & _
                                                strContent & "','" & _
                                                strDescription & "')"
    Call gobjComLib.zlDatabase.ExecuteProcedure(strSQL, "保存重要工作日志")
    Exit Sub
errh:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Public Function WebAPIByBasic(ByVal strURL As String, ByVal strJsIn As String, ByRef strJsOut As String, Optional ByRef strErr As String) As Boolean
'功能：通用WebAPI服务调用，使用的Basic认证方式
'参数：strURL 服务地址，
'      strJsIn 服务参数 JSON格式
'      strJsOut 调用服务后返回的出参 JSON格式
'      strErr  出错后的误信息
    Dim objHTTP As Object ' WinHttp.WinHttpRequest
    
    On Error GoTo errH
    
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    objHTTP.Option(4) = 13056
    objHTTP.Open "POST", strURL, False    '默认是POST方式
    objHTTP.SetRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    objHTTP.SetRequestHeader "AUthorization", "Basic " & gobjComLib.zlStr.Base64Encode("xxx:xxx")
    objHTTP.Send strJsIn
    strJsOut = objHTTP.ResponseText
    WebAPIByBasic = True
    Exit Function
errH:
    strErr = "调用服务失败【" & Err.Number & "--" & Err.Description & "】" & vbCrLf & "地址：" & strURL
    Err.Clear
End Function

Public Function NewSystemSvr(ByVal strSys As String, ByVal strName As String, ByVal strJsIn As String, ByRef strJsOut As String, Optional ByRef strErr As String) As Boolean
'功能：三方服务调用
'参数：strSys 系统标识， 三方服务配置目录.系统标识
'      strName 服务名称，三方服务配置目录.服务名称
'      strJsIn 服务参数 JSON格式
'      strJsOut 调用服务后返回的出参 JSON格式
'      strErr  出错后的误信息
    
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = "select a.服务地址 from 三方服务配置目录 a where a.系统标识=[1] and a.服务名称=[2] and a.服务地址 is not null"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "NewSystemSvr", strSys, strName)
    If Not rsTmp.EOF Then
        NewSystemSvr = WebAPIByBasic(rsTmp!服务地址 & "", strJsIn, strJsOut, strErr)
    End If
    
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function