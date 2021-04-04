Attribute VB_Name = "mdlSysOperate"
Option Explicit
'文件类型,该类型顺序与文件执行顺序相同
'X.X.X可能为4位版本号X.X.X.X,此时为特殊SP脚本。
Public Enum FileType
    'FT_Before 脚本与FT_DBA脚本执行执行顺序可以互换
    FT_DBA = 0 '需要DBA用户执行的脚本(System用户):ZLUPgradeX.X.X_DBA.sql,ZL*_X.X.X_DBA.sql
    FT_Before = 1 '提前执行脚本：ZLUPgradeX.X.X_Before.sql.sql(管理工具）,ZL*_X.X.X_History_Before.sql (应用系统历史库)ZL*_X.X.X_Before.sql(应用系统在线库) *代表系统号\100
    FT_Standard = 2 '普通升级脚本：ZLUPgradeX.X.X.sql,ZLUPgradeX.X.X(补充).sql,ZL*_X.X.X.sql ,ZL*_X.X.X(补充).sql,ZL*_X.X.X_History.sql
    FT_Optional = 3 '可选执行脚本:ZLUPgradeX.X.X_Optional.sql,ZL*_X.X.X_Optional.sql，ZL*_X.X.X__HISTORY_Optional.sql
    FT_Deferred = 4 '延迟执行脚本:ZL*_X.X.X_Deferred.sql,ZL*_X.X.X__HISTORY_DEFERRED
End Enum
'文件所属系统
Public Enum SysType
    ST_Tools = 0 '管理工具脚本,具有文件类型：FT_Before,FT_DBA,FT_Standard,FT_Optional
    ST_App = 1 '应用系统在线库,具有文件类型：FT_Before,FT_DBA,FT_Standard,FT_Optional，FT_Deferred
    ST_History = 2 '应用系统历史库，具有文件类型：FT_Before,FT_Standard,FT_Deferred，FT_Optional
End Enum
'版本类型
Public Enum VersionType
    VT_Normal = 0 '正常版本
    VT_Supple = 1 '补充发布版本，下一个大版本发布后，前一个版本新发布的SP就是补充版本
End Enum

Public Enum UserCheckType
    UCT_ZLTOOLS = 0 '管理工具用户验证
    UCT_DBAUser = 1 'DBA用户验证
    '以前该类的序号为1，现在调整为2，主要后面连续这几种类型都是通过直接调用窗体来使用的
    UCT_CurZLBAK = 2 '当前历史库验证
    UCT_NormalUser = 3 '普通用户验证
    UCT_SysOwner = 4 '管理员登录验证
    UCT_RACInsUser = 5 'RAC实例用户验证
End Enum

Public gcllMustObj As Collection '必要对象检查
Public gobjLog As TextStream
Private mstrStSysOwner As String '标准版所有者
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
        If rsINI!内容 & "" = "" Then Exit Function
        If arrItem(i) Like "*版本号" Then
            If Not IsVerSion(rsINI!内容) Then Exit Function
        End If
    Next
    CheckINIValid = True
End Function

Public Function SplitLine(ByVal strSQL As String) As Variant
'功能：对SQL进行换行拆分，同时记录换行符
    Dim arrLine As Variant, arrReturn() As Variant
    Dim i As Long, j As Long, lngStart As Long, lngEx As Long, lngCur As Long
    Dim strTmp As String
    arrReturn = Array()
    If strSQL = "" Then SplitLine = arrReturn: Exit Function
    arrLine = Split(Replace(Replace(strSQL, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    ReDim Preserve arrReturn(UBound(arrLine) * 2)
    lngStart = 1
    For i = LBound(arrLine) To UBound(arrLine)
        If i <> 0 Then
            strTmp = Mid(strSQL, lngStart, 2)
            If strTmp = vbCrLf Then
                arrReturn(i * 2 - 1) = vbCrLf
                lngStart = lngStart + 2
            Else
                arrReturn(i * 2 - 1) = Mid(strSQL, lngStart, 1)
                lngStart = lngStart + 1
            End If
        End If
        arrReturn(i * 2) = arrLine(i)
        lngStart = lngStart + Len(arrLine(i))
    Next
    SplitLine = arrReturn
End Function

Public Function TrimCommentLossless(ByVal strSQL As String) As String
'功能：无损去掉注释，与TrimComment比较，该算法不会损害真实数据。
    Dim arrLine As Variant, arrTmp As Variant
    Dim i As Long, j As Long
    Dim blnStr As Boolean, blnMultiCom As Boolean
    Dim lngPos1 As Long, lngPos2 As Long, lngPos3 As Long
    Dim blnAddLine As Boolean
    Dim strTmp As String, strFMT As String
    
    On Error GoTo errH
    '去除多行注释。
    arrTmp = Split(strSQL, "'")
    strFMT = "": blnStr = False: blnMultiCom = False
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Not blnStr Then
            arrLine = SplitLine(arrTmp(i))
            blnAddLine = True
            For j = LBound(arrLine) To UBound(arrLine) Step 2
                strTmp = arrLine(j)
                blnAddLine = j <> UBound(arrLine)
                If blnMultiCom Then '已经处于多行注释范围，则优先查找结束符
                    lngPos2 = InStr(strTmp, "*/")
                    If lngPos2 > 0 Then
                        strTmp = Mid(strTmp, lngPos2 + 2)
                        blnMultiCom = False
                    Else
                        strTmp = "": blnAddLine = False
                    End If
                End If
                If Not blnMultiCom Then '针对/* -- */ 与/*   */--处理
                    lngPos2 = InStr(strTmp, "/*")
                    lngPos1 = InStr(strTmp, "--")
                    '去掉有效的多行注释内容'/* --*/ ,/* */ 代码段 --/* */
                    '1、存在--,但是--在多行开始符之后
                    '2、不存在--，存在多行开始符
                    Do While Not blnMultiCom And (lngPos2 > 0 And lngPos2 < lngPos1 Or lngPos1 = 0 And lngPos2 > 0)
                        lngPos3 = InStr(lngPos2, strTmp, "*/")
                        If lngPos3 > 0 Then
                            strTmp = Left(strTmp, lngPos2 - 1) & Mid(strTmp, lngPos3 + 2)
                        Else
                            strTmp = Left(strTmp, lngPos2 - 1)
                            blnMultiCom = True
                        End If
                        lngPos2 = InStr(strTmp, "/*")
                        lngPos1 = InStr(strTmp, "--")
                    Loop
                End If
                '注释中的空行，则不做处理
                If blnAddLine Then
                    strFMT = strFMT & strTmp & arrLine(j + 1)
                Else
                    strFMT = strFMT & strTmp
                End If
            Next
        Else
            strTmp = ""
            '针对 "'B''C''D'"该类字符串进行识别
            For j = i To UBound(arrTmp) Step 2
                strTmp = strTmp & arrTmp(j)
                If j + 1 <= UBound(arrTmp) Then
                    If arrTmp(j + 1) = "" Then '存在空串，则为单引号字符
                        strTmp = strTmp & "''"
                    Else '不存在，则该处为字符的最后一段
                        i = j: Exit For
                    End If
                Else
                    i = j: Exit For
                End If
            Next
            strFMT = strFMT & "'" & strTmp & "'"
        End If
        If Not blnMultiCom Then '非多行注释，则调整字符串边界
            blnStr = Not blnStr '开始进入字符串边界
        End If
    Next
    
    '去除单行注释
    arrTmp = Split(strFMT, "'")
    strFMT = "": blnStr = False: blnMultiCom = False
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Not blnStr Then
            arrLine = SplitLine(arrTmp(i))
'            blnMultiCom = False
            For j = LBound(arrLine) To UBound(arrLine) Step 2
                strTmp = arrLine(j)
                If j = LBound(arrLine) And blnMultiCom Then
                    blnMultiCom = UBound(arrLine) = 0
                Else
                    blnAddLine = j <> UBound(arrLine)
                    lngPos1 = InStr(strTmp, "--")
                    If lngPos1 > 0 Then
                        strTmp = Left(strTmp, lngPos1 - 1)
                        blnMultiCom = UBound(arrLine) = j
                    End If
                    If blnAddLine Then
                        strFMT = strFMT & strTmp & arrLine(j + 1)
                    Else
                        strFMT = strFMT & strTmp
                    End If
                End If
            Next
        Else
            strTmp = ""
            '针对 "'B''C''D'"该类字符串进行识别
            For j = i To UBound(arrTmp) Step 2
                strTmp = strTmp & arrTmp(j)
                If j + 1 <= UBound(arrTmp) Then
                    If arrTmp(j + 1) = "" Then '存在空串，则为单引号字符
                        strTmp = strTmp & "''"
                    Else '不存在，则该处为字符的最后一段
                        i = j: Exit For
                    End If
                Else
                    i = j: Exit For
                End If
            Next
            strFMT = strFMT & "'" & strTmp & "'"
        End If
        If Not blnMultiCom Then
            blnStr = Not blnStr '开始进入字符串边界
        End If
    Next
    TrimCommentLossless = strFMT
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function GetFMTSQLStr(ByVal strSQL As String, ByRef cllStrs As Collection) As String
'功能：获取SQL中的字符串，并用占位符占位，返回格式化的SQL
    Dim arrTmp As Variant
    Dim i As Long, j As Long, intIndex As Integer
    Dim strFMT As String, strTmp As String
    Dim blnStr As Boolean
    
    Set cllStrs = New Collection
    arrTmp = Split(strSQL, "'")
    strFMT = "": blnStr = False
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Not blnStr Then
            strFMT = strFMT & arrTmp(i)
        Else
            strTmp = ""
            '针对 "'B''C''D'"该类字符串进行识别
            For j = i To UBound(arrTmp) Step 2
                strTmp = strTmp & arrTmp(j)
                If j + 1 <= UBound(arrTmp) Then
                    If arrTmp(j + 1) = "" Then '存在空串，则为单引号字符
                        strTmp = strTmp & "''"
                    Else '不存在，则该处为字符的最后一段
                        i = j: Exit For
                    End If
                Else
                    i = j: Exit For
                End If
            Next
            intIndex = intIndex + 1
            '标记字符串
            strFMT = strFMT & "[S" & intIndex & "]"
            cllStrs.Add strTmp, "S" & intIndex
        End If
        blnStr = Not blnStr '开始进入字符串边界
    Next
    arrTmp = SplitLine(strFMT)
    strFMT = "": blnStr = False
    For i = LBound(arrTmp) To UBound(arrTmp) Step 2
        strTmp = TrimEx(arrTmp(i))
        If strTmp <> "" Then
            If Right(strTmp, 1) = ";" And i <> UBound(arrTmp) Then
                strFMT = strFMT & " " & strTmp & vbCrLf
            Else
                strFMT = strFMT & " " & strTmp
            End If
        End If
    Next
    '去掉操作符中的空格
    arrTmp = SplitLine(strFMT)
    strFMT = ""
    For i = LBound(arrTmp) To UBound(arrTmp) Step 2
        strTmp = TrimEx(TrimBesideOperator(arrTmp(i)))
        If strTmp <> "" Then
            If Right(strTmp, 1) = ";" And i <> UBound(arrTmp) Then
                strFMT = strFMT & " " & strTmp & vbCrLf
            Else
                strFMT = strFMT & " " & strTmp
            End If
        End If
    Next
    GetFMTSQLStr = UCase(strFMT)
End Function

Public Function TrimBesideOperator(ByVal strText As String) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'说明：主要是RunSQLFile的子函数
    Dim i As Long
    
    strText = Replace(Replace(strText, " :", ":"), ": ", ":")
    strText = Replace(Replace(strText, " =", "="), "= ", "=")
    strText = Replace(Replace(strText, " .", "."), ". ", ".")
    strText = Replace(Replace(strText, " )", ")"), ") ", ")")
    strText = Replace(Replace(strText, " (", "("), "( ", "(")
    strText = Replace(Replace(strText, " %", "("), "% ", "%")
    strText = Replace(Replace(strText, " \", "\"), "\ ", "\")
    TrimBesideOperator = strText
End Function

Public Function GetInfoInsideBracket(ByVal strInfo As String, Optional ByVal strLeftChar As String, Optional ByVal strRightChar As String) As String
'从括号里面取内容
'返回括号里面的内容，只取最外层
    Dim lngSart As Long, lngEnd As Long
    If strRightChar = "" Then strRightChar = ")"
    If strLeftChar = "" Then strLeftChar = "("
    lngEnd = InStrRev(strInfo, strRightChar) - Len(strRightChar) + 1 '算头不算尾，所以不减一
    lngSart = InStr(strInfo, strLeftChar) + Len(strLeftChar)
    If lngEnd < lngSart Then
        GetInfoInsideBracket = ""
    Else
        GetInfoInsideBracket = Mid(strInfo, lngSart, lngEnd - lngSart)
    End If
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
    Dim i As Long, K As Long
    
    If Left(strSQL, 2) <> "--" And InStr(strSQL, "--") > 0 Then
        For i = 1 To Len(strSQL)
            If Mid(strSQL, i, 1) = "'" Then blnStr = Not blnStr
            If Mid(strSQL, i, 2) = "--" And Not blnStr Then
                K = i: Exit For
            End If
        Next
        If K > 0 Then strSQL = RTrim(Left(strSQL, K - 1))
    End If
    TrimComment = strSQL
End Function

Public Function SplitSQL(ByVal strSQL As String) As String
'功能：取";"结尾前面的的SQL语句,可能";"号后有"--"注释。
'说明：主要是RunSQLFile的子函数
    Dim i As Long, K As Long
    
    '先去掉注释部份
    strSQL = TrimComment(strSQL)
    
    For i = Len(strSQL) To 1 Step -1
        If Mid(strSQL, i, 1) = ";" Then
            K = i: Exit For
        End If
    Next
    If K > 0 Then strSQL = Left(strSQL, K - 1)
    
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


Public Function CheckInitFile(ByVal lngSys As Long, ByVal strFile As String, Optional ByVal blnOnlyCheck As Boolean, Optional ByRef rsReturnINI As ADODB.Recordset, Optional ByVal blnUpgradeCheck As Boolean = True) As Boolean
'参数：blnUpgradeCheck=检查升迁检查文件
   Dim strSysPath As String, strTmp As String
   Dim rsINI As ADODB.Recordset
   If Not gobjFile.FileExists(strFile) Then
        If Not blnOnlyCheck Then MsgBox "安装配置文件""" & strFile & """不存在。", vbExclamation, gstrSysName
        Exit Function
    End If
    If UCase(gobjFile.GetFileName(strFile)) <> IIf(lngSys = 0, "ZLSERVER.SQL", "ZLSETUP.INI") Then
        If Not blnOnlyCheck Then MsgBox "安装配置文件名不正确。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If lngSys = 0 Then '管理工具
        '检查管理工具升级检查函数文件是否存在。
        If blnUpgradeCheck Then
            strSysPath = gobjFile.GetParentFolderName(strFile)
            strTmp = strSysPath & "\zlUpgradeCheck.sql"
            If Not gobjFile.FileExists(strTmp) Then
                If Not blnOnlyCheck Then MsgBox "管理工具升级检查文件""" & strTmp & """不存在。", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
    Else '应用系统
        Set rsINI = ReadINIToRec(strFile)
        If Not CheckINIValid(rsINI, "系统号|版本号|表空间|管理工具版本号") Then
            If Not blnOnlyCheck Then MsgBox "安装配置文件格式不正确。", vbExclamation, gstrSysName
            Exit Function
        End If
        '配置文件系统号不匹配
        rsINI.Filter = "项目='系统号'"
        If Val(rsINI!内容) <> lngSys \ 100 Then
            If Not blnOnlyCheck Then MsgBox "所选配置文件不是本系统的安装配置文件。", vbExclamation, gstrSysName
            Exit Function
        End If
        strSysPath = gobjFile.GetParentFolderName(gobjFile.GetParentFolderName(strFile))
        '系统升迁目录检查
        If Not gobjFile.FolderExists(strSysPath & "\升级脚本") Then
            If Not blnOnlyCheck Then MsgBox "系统升迁目录""" & strSysPath & "\升级脚本""不存在。", vbExclamation, gstrSysName
            Exit Function
        End If
        If blnUpgradeCheck Then
            '检查应用系统升级检查函数文件是否存在。
            strTmp = strSysPath & "\升级脚本\zl" & lngSys \ 100 & "_UpgradeCheck.sql"
            If Not gobjFile.FileExists(strTmp) Then
                If Not blnOnlyCheck Then MsgBox "系统升级检查文件""" & strTmp & """不存在。", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
        '对应的安装脚本文件是否存在,不需要检查，因为已经取消了可选脚本执行
    End If
    Set rsReturnINI = rsINI
    CheckInitFile = True
End Function

Public Function GetUpgradeFiles(ByVal rsUpgradeFiles As ADODB.Recordset, ByVal lngSys As Long, ByVal strCurVer As String, ByVal strIniPath As String, _
                                                        Optional ByVal strNoramlBreak As String, Optional ByVal strBeforeBreak As String, _
                                                        Optional ByRef strMaxVer As String, Optional ByRef strCurMaxVer As String, Optional ByVal strBakDB As String, _
                                                        Optional ByVal blnReadByMax As Boolean, Optional ByVal blnDeleteSpfile As Boolean = True) As ADODB.Recordset
'功能：获取升级要执行的文件
'参数：rsUpgradeFiles=升级文件记录集，可能是多个系统的升级文件记录集
'          lngSys=系统号,=-1表示只初始化记录集
'          strIniPath=安装配置文件
'          strBreakVers=升迁配置文件的断点版本
'          strBakDB=历史库用名
'          strMaxVer=最大的版本
'          strCurMaxVer=本次升迁的目标版本
'          blnReadByMax=根据最大版本strMaxVer读取脚本（主要用于系统安装时管理工具版本较低管理工具单独升级时使用）
'                                   该参数为True时，不会进行断点处理，其余和正常应用系统脚本处理一致
'          blnDeleteSpfile=是否删除特殊SP文件,Ture-只获取升级的目标版本的特殊SP脚本 False-获取所有安装过的特殊SP脚本
'返回:升级文件记录
'        strMaxVer=最终目标版本,即当前脚本所能升迁到的最大打版本
'        strCurMaxVer=本次升迁的目标版本，系统升迁可能由于某些版本不能连续升迁，可能需要分多次升迁在能到最终目标版本。
'                               没有不能连续升迁的版本时,该版本与strMaxVer相同
'说明：
'        strBakDB="":读取所有脚本。此时如下参数含义
'                            strNoramlBreak：在线库（lngSys=0是为管理工具）常规升级中止信息
'                            strBeforeBreak:在线库（lngSys=0是为管理工具）提前升级中止信息
'                            strMaxVer:仅用来返回升迁的最终目标版本
'                            strCurMaxVer:仅用来返回升迁的本次目标版本
'                            返回的文件记录集中高于本次升迁目标版本的脚本全部剔除。
'        strBakDB<>"":读取大于strCurVer并且不大于strMaxVer的脚本。并于生成历史库的脚本文件记录集。
'                             在历史库非单独升迁时，生成的脚本文件记录集，要包括大于应用系统当前版本与应用系统本次目标版本之间的历史库脚本
'                             此时如下参数含义：
'                            strNoramlBreak：历史库常规升级中止信息
'                            strBeforeBreak:历史库提前升级中止信息
'                            strMaxVer:在线库的当前版本
    Dim rsCurFiles As ADODB.Recordset, arrFields As Variant, blnNew As Boolean
    Dim strCurPriFull As String, strCurFull As String, strMaxFull As String, strMaxPriFull As String
    Dim cllFolder As New Collection, objFolder As Folder, objFile As File
    Dim strBreak As String, strTmp As String, arrTmp As Variant, strFilter As String
    Dim strFileVer As String, stFile As SysType, ftFile As FileType, vtFile As VersionType, strSetupVer As String, blnSpecial As Boolean
    Dim strFileNameRule As String, stJudge As SysType
    Dim cllSuppleVers As New Collection, Item As Variant
    Dim i As Long
    Dim strFirstBreak As String, strSecdBreak As String
    Dim strBaseSupple As String
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strBanner As String, intSpVer As Integer
    
    On Error GoTo errH
    
    strCurPriFull = VerFull(GetPrimaryVer(strCurVer))
    strCurFull = VerFull(strCurVer)
    strMaxFull = VerFull(strMaxVer, True) '空串会生成9999.9999.9999.9999
    strMaxPriFull = VerFull(GetPrimaryVer(strMaxFull)) '防止空串生成失败，因此不用strMaxVer生成
    If rsUpgradeFiles Is Nothing Then
        blnNew = True
    ElseIf rsUpgradeFiles.State = adStateClosed Then
        blnNew = True
    End If
    
    If blnNew Or lngSys = -1 Then
        '配置版本:对提前执行脚本为最低要求版本，对应应用系统在线库普通升级脚本为对应管理工具脚本
        Set rsUpgradeFiles = CopyNewRec(Nothing, True, , _
                                                                Array("系统编号", adInteger, 5, Empty, "所有者", adVarChar, 100, Empty, "SysType", adInteger, 1, Empty, _
                                                                        "FileName", adVarChar, 50, Empty, "FilePath", adVarChar, 1000, Empty, "FileType", adInteger, 1, Empty, _
                                                                        "SPVer", adVarChar, 20, Empty, "FullSPVer", adVarChar, 20, Empty, "VerType", adInteger, 1, Empty, _
                                                                        "Optional", adVarChar, 2000, Empty, "AbortLine", adInteger, 10, Empty, "Special", adInteger, 1, Empty, _
                                                                        "配置版本", adVarChar, 20, Empty, "断点", adInteger, 1, Empty))
    End If
    If lngSys = -1 Then Set GetUpgradeFiles = rsUpgradeFiles: Exit Function
    '读取当前系统的脚本
    rsUpgradeFiles.Filter = "系统编号=" & lngSys & IIf(strBakDB <> "", " And 所有者='" & UCase(strBakDB) & "'", "")
    '脚本已经存在，则不用重新读取。
    '历史库读取，必须最大版本不为空。因为历史库单独升迁的目标版本为在线库当前版本。非单独升级时，在线库当前版本之上的历史脚本已经读取
    If Not rsUpgradeFiles.EOF Or strBakDB <> "" And strMaxVer = "" Then Set GetUpgradeFiles = rsUpgradeFiles: Exit Function
    Set rsCurFiles = CopyNewRec(rsUpgradeFiles, strBakDB = "")
    '////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////          1、升迁文件读取            ///////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////
    '获取需要搜集脚本的文件夹
    If lngSys = 0 Then
        cllFolder.Add gobjFile.GetFile(strIniPath).ParentFolder
        strFileNameRule = "ZLUPGRADE*.*.*.SQL"
    Else
        strFileNameRule = "ZL" & lngSys \ 100 & "_*.*.*.SQL"
        For Each objFolder In gobjFile.GetFolder(gobjFile.GetParentFolderName(gobjFile.GetParentFolderName(strIniPath)) & "\升级脚本\").SubFolders
            If IsVerSion(objFolder.Name) And objFolder.Name Like "*.*.0" Then
                If VerFull(objFolder.Name) >= strCurPriFull And VerFull(objFolder.Name) <= strMaxPriFull Then
                    cllFolder.Add objFolder
                End If
            End If
        Next
    End If
    arrFields = Array("系统编号", "SysType", "FileName", "FilePath", "FileType", "SPVer", "FullSPVer", "VerType", "Special", "配置版本")
    '遍历,提取文件
    For Each objFolder In cllFolder
        If lngSys <> 0 And strBakDB = "" Then '获取zlUpgrade.ini
            '获取有效的断点版本
            strTmp = GetUpgradeIniBreak(objFolder.Path & "\zlUpgrade.ini", IIf(VerFull(objFolder.Name) >= strCurPriFull, strCurVer, objFolder.Name), GetPrimaryVer(objFolder.Name, True))
            If strTmp <> "" Then
                strBreak = strBreak & "," & strTmp
            End If
        End If
        '获取文件
        For Each objFile In objFolder.Files
            If UCase(objFile.Name) Like strFileNameRule Then '符合文件的规则的才进行名称解析
                If AnalysisFileName(objFile.Name, lngSys, strFileVer, ftFile, stFile, vtFile, blnSpecial) Then
                    If VerFull(strFileVer) > strCurFull And VerFull(strFileVer) <= strMaxFull Then
                        If vtFile = VT_Supple Then
                            On Error Resume Next
                            '确认该大版本已经标记的补充版本
                            strBaseSupple = cllSuppleVers("K_" & GetPrimaryVer(strFileVer))
                            If Err.Number <> 0 Then
                                Err.Clear
                                cllSuppleVers.Add strFileVer, "K_" & GetPrimaryVer(strFileVer)
                            '已经标记的补充版本小于当前版本，则讲标记修改为当前版本
                            ElseIf VerFull(strBaseSupple) > VerFull(strFileVer) Then
                                cllSuppleVers.Remove "K_" & GetPrimaryVer(strFileVer)
                                cllSuppleVers.Add strFileVer, "K_" & GetPrimaryVer(strFileVer)
                            End If
                            On Error GoTo errH
                        End If
                        '获取配置版本
                        If ftFile = FT_Before Or ftFile = FT_Standard And stFile = ST_App And VerFull(strFileVer) > VerFull("10.32.0") Then
                            arrTmp = Split(GetUpgradeCtrolInfo(objFile.Path, ftFile = FT_Before) & "|", "|")
                            strSetupVer = VerFull(arrTmp(IIf(ftFile = FT_Before, 0, 1))) '扩充为标准版本，方便比较;    提前执行返回：最低要求版本，常规升级脚本返回：连续升级|对应管理工具版本
                            '10.34.0之后，管理工具，应用系统版本已经一一对应，且没有脚本的版本用空文件放置
                            If ftFile = FT_Standard Then
                                 If VerFull(strFileVer) >= VerFull("10.34.0") Then
                                    strSetupVer = VerFull(strFileVer) '扩充为标准版本，方便比较
                                ElseIf strSetupVer = VerFull("0") Then  '读取应用对应工具版本失败，则自动生成一个
                                    strSetupVer = VerFull(GetContractVersion(strFileVer, True))
                                End If
                            End If
                            If Val(arrTmp(0)) <> 1 And ftFile = FT_Standard And strBakDB = "" Then strBreak = strBreak & "," & strFileVer
                        Else
                            strSetupVer = ""
                        End If
                        
                        rsCurFiles.AddNew arrFields, Array(lngSys, stFile, objFile.Name, objFile.Path, ftFile, strFileVer, VerFull(strFileVer), vtFile, IIf(blnSpecial, 1, 0), strSetupVer)
                    End If
                End If
            End If
        Next
    Next
    '////////////////////////////////////////////////////////////////////////////////////
    '////////////////////   2.上次升迁信息的剔除，补充版本断点标记  ///////////////////
    '///////////////////////////////////////////////////////////////////////////////////
    '标记补充版本
    For Each Item In cllSuppleVers
        '大于该大版本的最小的补充版本，且小余下一个版本
        Call RecUpdate(rsCurFiles, "FullSPVer>='" & VerFull(Item) & "' And FullSPVer<'" & VerFull(GetPrimaryVer(Item, True)) & "'", "VerType", VT_Supple)
    Next
    stJudge = IIf(lngSys = 0, ST_Tools, IIf(strBakDB = "", ST_App, ST_History))
    strFilter = "SysType=" & stJudge & " And FileType<>" & FT_Deferred
    '剔除提前中止语句之前的文件
    arrTmp = Split(strBeforeBreak & "||", "|")
    '没有中止文件，则小于等于中止版本的提前执行脚本都要删除，否则，只删除小于中止版本的提前脚本
    Call RecDelete(rsCurFiles, strFilter & " And FileType=" & FT_Before & " And FullSPVer<" & IIf(arrTmp(1) = "", "=", "") & "'" & VerFull(arrTmp(0)) & "'")
    If arrTmp(1) <> "" Then '有中止文件，记录中止点
        Call RecUpdate(rsCurFiles, strFilter & "And FileType=" & FT_Before & " And SPVer='" & arrTmp(0) & "'", "AbortLine", Val(arrTmp(2)))
    End If
    arrTmp = Split(strNoramlBreak & "||", "|")
    '剔除正常中止语句之前的文件
    Call RecDelete(rsCurFiles, strFilter & " And FullSPVer<" & IIf(arrTmp(1) = "", "=", "") & "'" & VerFull(arrTmp(0)) & "'")
    If arrTmp(1) <> "" Then '有中止文件
        '删除中止中止版本中执行顺序在中止文件之前的文件
        Call RecDelete(rsCurFiles, strFilter & " And SPVer='" & arrTmp(0) & "' And FileType<" & Val(arrTmp(1)))
        '记录中止点
        Call RecUpdate(rsCurFiles, strFilter & " And SPVer='" & arrTmp(0) & "' And FileType=" & Val(arrTmp(1)), "AbortLine", Val(arrTmp(2)))
    End If
    '不能连续升迁版本的标记
    strBreak = Mid(strBreak, 2): arrTmp = Split(strBreak, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        Call RecUpdate(rsCurFiles, "SPVer='" & arrTmp(i) & "'", "断点", 1)
    Next
    
    '剔除补充版本。按版本排序，第一个非补充版本之前的所有补充版本全部删掉。
    rsCurFiles.Filter = "VerType=" & VT_Normal: rsCurFiles.Sort = "FullSPVer Desc"
    If Not rsCurFiles.EOF Then Call RecDelete(rsCurFiles, "VerType=" & VT_Supple & " And FullSPVer<'" & rsCurFiles!fullspver & "'")
    
    rsCurFiles.Filter = "": rsCurFiles.Sort = "FullSPVer Desc"
    If blnDeleteSpfile Then
        '剔除特殊SP脚本。按版本排序，第一个非特殊SP版本之前的所有特殊SP全部删掉
        '这种判断有个问题，可能一个版本没有正式脚本，但是有特殊SP脚本，因此不按这种处理。
        If Not rsCurFiles.EOF Then
            strTmp = VerFull(VerSpecialNormal(rsCurFiles!SPVer))
            Call RecDelete(rsCurFiles, "Special=1 And FullSPVer<'" & strTmp & "'")
        End If
    Else
        '如果要保留特殊SP脚本,就需要判断该特殊SP脚本是否安装
        rsCurFiles.Filter = "Special  =1"
        Do While Not rsCurFiles.EOF
            strTmp = Mid(rsCurFiles!SPVer, 1, InStrRev(rsCurFiles!SPVer, ".") - 1) '记录当前小版本
            If strBanner <> strTmp Then '排序后的记录,小版本不同,需要重新获取该小版本的最大特殊SP版本
                strBanner = strTmp
                strSQL = "Select Nvl(Max(Substr(结果版本, Instr(结果版本, '.', 1, 3) + 1)), 0) 最大sp" & vbNewLine & _
                            "From zlUpGrade A" & vbNewLine & _
                            "Where 系统 = [1] And 结果版本 Like '" & strTmp & ".%'"
                Set rsTmp = OpenSQLRecord(strSQL, "获取最大SP版本", lngSys)
                intSpVer = Val(rsTmp!最大sp)
            End If
            
            If Val(Split(rsCurFiles!fullspver, ".")(3)) > intSpVer Then
                rsCurFiles.Delete adAffectCurrent
            End If
            
            rsCurFiles.MoveNext
        Loop
    End If
    '////////////////////////////////////////////////////////////////////////////////////
    '/////////////// 3、最终目标版本、本次目标版本、以及历史库脚本的读取 ////////////
    '///////////////////////////////////////////////////////////////////////////////////
    If strBakDB = "" Then
        If blnReadByMax Then '根据最大版本读取
            '获取实际可以升级到的最大版本
            rsCurFiles.Filter = "": rsCurFiles.Sort = "FullSPVer Desc"
            strCurMaxVer = ""
            If Not rsCurFiles.EOF Then
                strCurMaxVer = rsCurFiles!SPVer & ""
            End If
        Else
            '获取最终目标版本以及本次目标版本
            rsCurFiles.Filter = "": rsCurFiles.Sort = "FullSPVer Desc"
            strMaxVer = "": strCurMaxVer = ""
            If Not rsCurFiles.EOF Then
                strMaxVer = rsCurFiles!SPVer & ""
                rsCurFiles.Filter = "断点=1": rsCurFiles.Sort = "FullSPVer"
                If Not rsCurFiles.EOF Then
                    strFirstBreak = rsCurFiles!SPVer
                    If rsCurFiles.RecordCount > 1 Then
                        rsCurFiles.MoveNext: strSecdBreak = rsCurFiles!SPVer
                    End If
                    rsCurFiles.Filter = "FullSPVer<'" & VerFull(strFirstBreak) & "'"
                    strCurMaxVer = IIf(rsCurFiles.EOF, strSecdBreak, strFirstBreak)
                End If
            End If
            If strCurMaxVer = "" Then
                strCurMaxVer = strMaxVer
            Else '删除不需要本次升迁不需要执行的脚本
                Call RecDelete(rsCurFiles, "FullSPVer>'" & VerFull(strCurMaxVer) & "'")
            End If
        End If
    Else
    '获取历史库升迁记录
        '删除小于历史库当前版本的脚本（历史库版本可能高于在线库，因此需要这样处理）
        Call RecDelete(rsCurFiles, "FullSPVer<='" & VerFull(strCurVer) & "'")
        '删除在线库脚本
        Call RecDelete(rsCurFiles, "SysType<>" & ST_History)
        '更新文件记录集的所有者
        Call RecUpdate(rsCurFiles, "", "所有者", UCase(strBakDB))
    End If
    '合并记录集，将本次读取的文件合并到所有记录集中
    rsCurFiles.Filter = ""
    Call RecDataAppend(rsUpgradeFiles, rsCurFiles)
    Set GetUpgradeFiles = rsUpgradeFiles
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox Err.Description, vbInformation, gstrSysName
End Function

Public Function FormatUpgradeBreak(ByVal lngSys As Long, ByVal strResultVer As String, Optional ByVal strUpgradeBreak As String) As String
'功能：解析中止信息，将中止语句标准化 格式：文件版本|文件类型|出错行号
'参数：
'     strResultVer:ZLUpgrade中的结果版本
'     strUpgradeBreak=升迁中止语句
'返回：文件的不带路径的文件名
    Dim arrTmp As Variant
    Dim lngPos As Long
    Dim strTmp As String
    Dim strFileName As String
    Dim lngAbort As Long
    Dim strFileVer As String '从文件名上读取的版本信息
    Dim ftReturn As FileType
    Dim strReturn As String
    
    strReturn = strResultVer & "||"
    If strUpgradeBreak <> "" Then
        '历史库的中止语句可能为版本号
        If Not IsVerSion(strUpgradeBreak) Then
            strUpgradeBreak = strUpgradeBreak & "||"
            arrTmp = Split(strUpgradeBreak, "|")
            If gobjFile.FileExists(arrTmp(0)) Then
                strFileName = gobjFile.GetFileName(arrTmp(0))
            Else '可能是补充版本已经删掉了
                strTmp = StrReverse(arrTmp(0))
                lngPos = InStr(strTmp, "\")
                '截取最后一个\后的内容
                If lngPos <> 0 Then
                    strFileName = StrReverse(Mid(strTmp, lngPos - 1))
                Else
                    strFileName = ""
                End If
            End If
            lngAbort = Val(arrTmp(1))
            If strFileName <> "" Then
                If AnalysisFileName(strFileName, lngSys, strFileVer, ftReturn) Then
                    strReturn = strFileVer & "|" & ftReturn & "|" & lngAbort
                End If
            End If
        Else '历史库提前升级存放的是版本号
            strReturn = strUpgradeBreak & "||"
        End If
    End If
    FormatUpgradeBreak = strReturn
End Function

Public Function GetUpgradeIniBreak(ByVal strFile As String, Optional ByVal strMinVer As String, Optional ByVal strMaxVer As String)
'功能：获取升迁配置文件的断点
'参数：strFile=升迁配置文件路径
'          strMinVer=升迁配置文件目标版本的最小值
'          strMaxVer=升迁配置文件目标版本的最大值
    Dim rsSub As ADODB.Recordset
    Dim strBreakVer As String
    
    If Not gobjFile.FileExists(strFile) Then Exit Function
    Set rsSub = ReadINIToRec(strFile)
    If rsSub Is Nothing Then Exit Function
    rsSub.Filter = "项目='连续升级'" '升级配置文件的目标版本是否能连续升级
    If rsSub.EOF Then Exit Function
    If Val(rsSub!内容 & "") = 1 Then Exit Function '连续升级不用处理
    rsSub.Filter = "项目='目标版本'" '升级配置文件的目标版本
    If rsSub.EOF Then Exit Function
    strBreakVer = Trim(rsSub!内容 & "")
    If Not IsVerSion(strBreakVer) Then Exit Function
    If strMinVer <> "" Then '小于最小版本，则该断点无效
        If VerFull(strBreakVer) <= VerFull(strMinVer) Then Exit Function
    End If
    If strMaxVer <> "" Then '大于最小版本，则该断点无效
        If VerFull(strBreakVer) > VerFull(strMaxVer) Then Exit Function
    End If
    GetUpgradeIniBreak = strBreakVer
End Function

Public Function GetUpgradeCtrolInfo(ByVal strFile As String, Optional ByVal blnBefore As Boolean) As String
'功能：获取文件中的控制信息
'      strFile=进行判断的脚本文件路径
'      blnBefore=文件是否是提起执行脚本
'返回: blnBefore=false: 连续升级|管理工具版本号
'        blnBefore=True: 最低版本号

    Dim objStream As Scripting.TextStream
    Dim strLine As String, arrFind() As Variant, i As Long, strTmp As String, arrTmp As Variant
    Dim strContinue As String, strToolVer As String, strBreakVer As String, strReqVer As String
    Dim rsSub As ADODB.Recordset
    
    On Error GoTo errH
    
    Set objStream = gobjFile.OpenTextFile(strFile, ForReading)
    If blnBefore Then
        arrFind = Array("[[]最低版本号[]]")
    Else
        arrFind = Array("[[]连续升级[]]", "[[]管理工具版本号[]]")
    End If
    Do While Not objStream.AtEndOfStream
        strLine = TrimEx(objStream.ReadLine, True)
        If strLine Like "--" & arrFind(i) & "*" Then
            strTmp = Trim(Mid(strLine, Len("--" & arrFind(i)) - 4 + 1))
            If Not blnBefore Then
                If i = 0 Then
                    strContinue = strTmp
                Else
                    strToolVer = strTmp
                End If
            Else
                strReqVer = strTmp
            End If
        End If
        If i = UBound(arrFind) Then Exit Do
        i = i + 1
    Loop
    objStream.Close
    
    If blnBefore Then
        GetUpgradeCtrolInfo = Trim(strReqVer)
    Else
        If Trim(strContinue) = "" Then strContinue = "1"
        GetUpgradeCtrolInfo = Trim(strContinue) & "|" & Trim(strToolVer)
    End If
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
'    Debug.Print err.Source & "\" & Me.name & "\GetCtrolInfo:" & err.Description
End Function


Public Function AnalysisFileName(ByVal strFileName As String, ByVal lngSys As Long, Optional ByRef strVersion As String, Optional ByRef ftReturn As FileType, _
                                                        Optional ByRef stReturn As SysType, Optional ByRef vtReturn As VersionType = VT_Normal, Optional ByRef blnSpecial As Boolean) As Boolean
'功能:t通过文件名获取文件信息
'参数：
'   strFile=不包含路径的文件名,带扩展名
'   lngSys=系统号
'返回:
'       True=成功获取，False=获取失败（文件不是系统升级脚本）
'       strVerReturn=文件版本
'       ftReturn=文件类型
'       stReturn=系统类型
'       vtReturn=版本类型
    Dim strSysString As String, strSuffix As String
    Dim arrVer As Variant
    vtReturn = VT_Normal
    blnSpecial = False
    strVersion = ""
    ftReturn = FT_Before
    stReturn = ST_Tools
    If Not UCase(strFileName) Like "*.SQL" Then Exit Function
    strFileName = UCase(Left(strFileName, Len(strFileName) - 4))
    arrVer = Split(strFileName, ".")
    '版本文件的文件名仅有2个句点号(特殊SP包含3个）
    If UBound(arrVer) < 2 Or UBound(arrVer) > 3 Then Exit Function
    '获取脚本系统前缀
    If arrVer(0) Like "ZLUPGRADE*" Then
        strSysString = "ZLUPGRADE"
        stReturn = ST_Tools
    ElseIf arrVer(0) Like "ZL" & lngSys \ 100 & "_*" Then
        strSysString = "ZL" & lngSys \ 100 & "_"
        stReturn = ST_App
    Else
        Exit Function '没有系统标识前缀，不是系统脚本
    End If
    '系统标识后面紧跟的是版本
    arrVer(0) = Mid(arrVer(0), Len(strSysString) + 1) '获取主板本
    arrVer(UBound(arrVer)) = GetPrefixNumber(arrVer(UBound(arrVer)), strSuffix) '获取次级版本
    '获取的主板本，大版本以及次级版本若不为数字，则退出
    If Not IsNumeric(arrVer(0)) Or Not IsNumeric(arrVer(1)) Or Not IsNumeric(arrVer(2)) Or Not IsNumeric(arrVer(UBound(arrVer))) Then Exit Function
    strVersion = arrVer(0) & "." & arrVer(1) & "." & arrVer(2) & IIf(UBound(arrVer) = 2, "", "." & arrVer(UBound(arrVer)))
    If Not IsVerSion(strVersion) Then Exit Function
    '四位版本号就是特殊SP
    blnSpecial = strVersion Like "*.*.*.*"
    '版本后是文件类型信息
    If stReturn = ST_App And strSuffix Like "_HISTORY*" Then
        stReturn = ST_History
        strSuffix = Mid(strSuffix, Len("_HISTORY") + 1)
    End If
    If strSuffix Like "*(补充)" Then
        vtReturn = VT_Supple
        strSuffix = Replace(strSuffix, "(补充)", "") '防止补充信息位置不固定
    End If
    Select Case strSuffix
        Case ""
            ftReturn = FT_Standard
        Case "_DBA"
            If stReturn = ST_History Then Exit Function '历史库不支持DBA脚本
            ftReturn = FT_DBA
        Case "_OPTIONAL"
            ftReturn = FT_Optional
        Case "_BEFORE"
            ftReturn = FT_Before
        Case "_DEFERRED"
            If stReturn = ST_Tools Then Exit Function '管理工具不支持延迟执行脚本
            ftReturn = FT_Deferred
        Case Else '不再命名规则范围内的，则获取失败
            Exit Function
    End Select
    AnalysisFileName = True
End Function

Public Function GetPrefixNumber(ByVal strInput As String, Optional ByRef strOther As String) As String
'功能：获取一个字符串的数字前缀，以及剩余部分
'参数：strInput=输入的字符串
'          strOther =去掉数字前缀的剩余部分
    Dim i As Long
    
    For i = 1 To Len(strInput)
        If Not IsNumeric(Mid(strInput, i, 1)) Then
            Exit For
        End If
    Next
    strOther = Mid(strInput, i)
    GetPrefixNumber = Mid(strInput, 1, i - 1)
End Function

Public Function VerFull(ByVal strVer As String, Optional ByVal blnMax As Boolean) As String
'功能：返回VB最大支持的版本号形式:9999.9999.9999.9999,最小版本号0000.0000.0000.0000
'参数：strVer=当前版本号
'           blnMax=True,若果为空，则返回最大支持版本，False=若果为空，则返回最小支持版本
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then
        VerFull = IIf(blnMax, "9999.9999.9999.9999", "0000.0000.0000.0000")
        Exit Function
    End If
    '增加一段，以兼容特殊SP版本号
    arrVer = Split(strVer & ".0", ".")
    VerFull = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & "." & Format(arrVer(2), "0000") & "." & Format(arrVer(3), "0000")
End Function

Public Function VerPAD(ByVal strVer As String) As String
'功能：使版本号的主版本号左填充为4位，保证主版本后原点可以与其他版本号对齐
'参数：strVer=当前版本号
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then
        Exit Function
    End If
    arrVer = Split(strVer & ".", ".")
    VerPAD = RPAD(Lpad(arrVer(0), 2) & "." & arrVer(1) & "." & arrVer(2) & IIf(Val(arrVer(3)) = 0, "", "." & Format(Val(arrVer(3)), "0000")), 20)
End Function

Public Function GetPrimaryVer(ByVal strVer As String, Optional ByVal blnNext As Boolean)
'功能：获取一个版本的主版本
'参数：strVer=当前版本
'          blnNext=是否获取下一个主版本
'返回：主版本
    Dim arrVer As Variant
    
    arrVer = Split(strVer & "..", ".")
    If blnNext Then
        GetPrimaryVer = Val(arrVer(0)) & "." & (Val(arrVer(1)) + 1) & "." & 0
        '管理工具没有9.45.0，直接和应用系统同一编号，为10.34.0
        If GetPrimaryVer = "9.45.0" Then GetPrimaryVer = "10.34.0"
    Else
        GetPrimaryVer = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & 0
    End If
End Function

Public Function GetContractVersion(ByVal strVer As String, Optional ByVal blnGetTools As Boolean = True)
'功能：获取应用系统对应管理工具的主版本，或者管理工具对应应用系统版本，需要推算
'参数：strVer=当前应用系统版本
'          blnGetTools=True-获取对应的管理工具版本,False-获取对应的应用系统版本
'返回：对应版本，应用系统10.34.0之前，只求对应大版本，不具体到SP版本
'                          管理工具10.34.0之前，只求对应大版本，不具体到SP版本
    Dim arrVer As Variant
    Dim lngDistance As Long
    If strVer = "" Then strVer = "9.1.0"
    If blnGetTools Then
        If VerFull(strVer) >= VerFull("10.34.0") Then '10.34.0  以后管理工具和应用系统版本统一
            GetContractVersion = strVer
        Else
            arrVer = Split(strVer & "...", ".")
            lngDistance = 33 - Val(arrVer(1)) '获取应用系统与10.33.0版本的大版本间隔
            '管理工具9.44.0减去相应大版本间隔就为对应管理工具版本
            GetContractVersion = "9." & (44 - lngDistance) & ".0"
        End If
    Else
        If VerFull(strVer) >= VerFull("10.34.0") Then  '  以后管理工具和应用系统版本统一
            GetContractVersion = strVer
        Else
            arrVer = Split(strVer & "...", ".")
            lngDistance = 44 - Val(arrVer(1)) '获取管理工具与9.44.0版本的大版本间隔
            '应用系统10.33.0减去相应大版本间隔就为对应应用系统的版本
            GetContractVersion = "10." & (33 - lngDistance) & ".0"
        End If
    End If
End Function

Public Function VerNormal(ByVal strVer As String) As String
'功能：将VB最大支持的版本号形式:9999.9999.9999转换为常见版本虚形式，如0010.0034.0000.0000，转换为10.34.0
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then Exit Function
    arrVer = Split(strVer & ".", ".")
    VerNormal = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & Val(arrVer(2)) & IIf(Val(arrVer(3)) = 0, "", "." & Format(Val(arrVer(3)), "0000"))
End Function

Public Function VerSpecialNormal(ByVal strVer As String) As String
'获取一个特殊sp对应的正式版本，如果是一个正式版本，则返回其自身
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then Exit Function
    arrVer = Split(strVer & ".", ".")
    VerSpecialNormal = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & Val(arrVer(2))
End Function

Public Function IsVerSion(ByVal strVer As String) As Boolean
'功能：判断字符串是否是版本号
    Dim arrVer As Variant
    Dim i As Integer
    If Not strVer Like "*.*.*" Then Exit Function
    arrVer = Split(strVer, ".")
    If UBound(arrVer) < 2 Or UBound(arrVer) > 3 Then Exit Function
    
    For i = LBound(arrVer) To UBound(arrVer)
        If Not IsNumeric(arrVer(i)) Then Exit Function
        If Val(arrVer(i)) < 0 Or Val(arrVer(i)) > 9999 Then Exit Function
        If i = 3 Then
            If Format(Val(arrVer(i)), "0000") <> Format(Trim(arrVer(i)), "0000") Then Exit Function
        Else
            If Val(arrVer(i)) & "" <> Trim(arrVer(i)) Then Exit Function
        End If
    Next
    
    IsVerSion = True
End Function

