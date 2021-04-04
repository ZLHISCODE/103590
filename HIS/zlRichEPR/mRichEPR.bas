Attribute VB_Name = "mRichEPR"
'#########################################################################
'##模 块 名：mRichEPR.bas
'##创 建 人：吴庆伟
'##日    期：2005年8月11日
'##修 改 人：
'##日    期：
'##描    述：全局变量、类型的定义
'##版    本：
'#########################################################################

Option Explicit

'##########################################################################################
'## 全局类型
'##########################################################################################

Public Type PreDefinedKeyInfo   '保留关键字
    KeyStart As String
    KeyEnd As String
End Type

'##########################################################################################
'## 全局变量
'##########################################################################################
Public gfrmPublic As frmPublic

Public gblnShowInTaskBar As Boolean         '是否显示窗体在任务条上
Public gcnOracle As New ADODB.Connection    '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                  '当前用户具有的当前模块的功能
Public gstrProductName As String            '产品简称，例如：中联
Public gstrSysName As String                '系统名称，例如：中联软件
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录
Public glngModul As Long                    '模块编号
Public glngSys As Long                      '系统编号，例如：100

Public gstrDbOwner As String                '当前数据库所有者（不同模块可能不一样）
Public gstrDBUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码
Public gstrSignName As String               '签名姓名
Public gstrPrivsEpr As String               '病历编辑模块1070权限
Public gstrCopyPID As String                '复制源病人ID

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public gstrMatch As String                  '根据本地参数“匹配模式”确定的左匹配符号
Public gstrSQL As String

Public gobjFSO As New Scripting.FileSystemObject    'FSO对象
Public gfrmParent As Object                         '全局的父窗体
Public gobjPacsCore As Object                       '由其它工作调用是传入了观片站对象
Public gobjESign As Object                  '电子签名接口部件
Public gobjTendESign As Object           '电子签名接口部件(护理)
Public gstrESign As String                  '是否启用电子签名
Public gobjEmr As Object                    '新版电子病历
Public gobjInfection As Object              '传染病报告卡
Public gobjPlugIn As Object                 '插件
Public gobjRegister As Object               'ZLHIS密码验证组件

Public gKeyWords(1 To 6) As PreDefinedKeyInfo       '预定义关键字

Public Const ELE_BACKCOLOR = &HD5FEFF               '要素的背景颜色 '&HDCDCDC
Public Const ELE_UNDERLINE = cprwave                '要素的下划线
Public Const PROTECT_BGCOLOR = &HE0E0E0             '自定义保护文本的背景色
Public Const PROTECT_FORECOLOR = &H662200           '自定义保护文本的前景色
Public Const TABLEELE_FORECOLOR = &H100080          '表格要素的前景色
Public Const ELE_JUMP_LIMIT = 32                    '回车键后自动跳到下一要素的距离限制

'刷新数据时传入的病人状态
Public Enum TYPE_PATI_State
    ps在院 = 0
    ps预出 = 1
    ps出院 = 2
    ps待诊 = 3          '医生站:待会诊病人(在院)
    ps已诊 = 4          '医生站:已会诊病人
    ps最近转出 = 5      '医护站:最近转科或转病区的病人(在院)
    ps待转入 = 6        '医护站:入科待入住或转病区待入往病人
End Enum

'##########################################################################################
'## 压缩与解压
'##########################################################################################
Private mclsZip As New cZip
Private mclsUnzip As New cUnzip

'################################################################################################################
'## 功能：  将指定的LOB字段复制为临时文件
'##
'## 参数：  Action      :操作类型（用以区别是操作哪个表）
'##         KeyWord     :确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'##         strFile     :用户指定存放的文件名；不指定时，取当前路径产生文件名
'##
'## 返回：  存放内容的文件名，失败则返回零长度""
'##
'## 说明：  Action取值说明：
'##         0-病历标记图形；1-病历文件格式；2-病历文件图形；3-病历范文格式；4-病历范文图形；5-电子病历格式；6-电子病历图形；
'################################################################################################################
Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String, Optional ByVal blnMoved As Boolean) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, StrText As String
    Dim rsLob As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand
    
    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".tmp"
            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum
    
    gstrSQL = "Select Zl_Lob_Read([1],[2],[3],[4]) as 片段 From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSQL, "zlBlobRead", Action, KeyWord, lngCount, IIf(blnMoved, 1, 0))
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        StrText = rsLob.Fields(0).Value
        
        ReDim aryChunk(Len(StrText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(StrText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

errHand:
    Close lngFileNum
    Kill strFile
    If ErrCenter = 1 Then
        Resume
    End If
    zlBlobRead = ""
End Function


'Writed by zyb 20110907
Public Function zlClobRead(ByVal Action As Long, ByVal KeyWord As String) As String
    'KeyWord:ID,'/ITEM/XH'或者'/ITEM/MC'
    Dim lngCount As Long
    Dim StrText As String, strReturn As String
    Dim rsLob As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand
    
    gstrSQL = "Select Zl_Lob_Read([1],[2],[3],[4]) as 片段 From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSQL, "zlClobRead", Action, KeyWord, lngCount, 0)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        
        StrText = rsLob.Fields(0).Value
        strReturn = strReturn & StrText
        lngCount = lngCount + 1
    Loop
    zlClobRead = strReturn
errHand:
End Function

Public Function zlClobSql(ByVal KeyWord As String, ByVal strFileContent As String, ByRef arySql() As String) As Boolean
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, StrText As String
    Dim lngLBound As Long, lngUBound As Long    '传入数组的最小最大下标
    
    Err = 0: On Error Resume Next
    lngLBound = LBound(arySql): lngUBound = UBound(arySql)
    If Err <> 0 Then lngLBound = 0: lngUBound = -1
    Err = 0: On Error GoTo errHand
    
    lngFileSize = Len(strFileContent)
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    
    ReDim Preserve arySql(lngLBound To lngUBound + lngBlocks + 1) As String
    For lngCount = 0 To lngBlocks
        StrText = Mid(strFileContent, conChunkSize * lngCount + 1, conChunkSize)
        arySql(lngUBound + lngCount + 1) = "Zl_Lob_Append(21,'" & KeyWord & "','" & StrText & "'," & IIf(lngCount = 0, 1, 0) & ",1)"
    Next
    zlClobSql = True
    Exit Function

errHand:
    zlClobSql = False
End Function


'################################################################################################################
'## 功能：  将指定的文件保存到指定记录的LOB字段中
'##
'## 参数：  Action      :操作类型（用以区别是操作哪个表）
'##         KeyWord     :确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'##         strFile     :用户指定存放的文件名；不指定时，取当前路径产生文件名
'##
'## 返回：  成功返回True，失败返回False
'##
'## 说明：  Action取值说明：
'##         0-病历标记图形；1-病历文件格式；2-病历文件图形；3-病历范文格式；4-病历范文图形；5-电子病历格式；6-电子病历图形；
'################################################################################################################
Public Function zlBlobSave(ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String) As Boolean
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, StrText As String
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    Err = 0: On Error GoTo errHand
    
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        StrText = Join(aryHex, "")
        gstrSQL = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "','" & StrText & "'," & IIf(lngCount = 0, 1, 0) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "zlBlobSave")
    Next
    Close lngFileNum
    zlBlobSave = True
    Exit Function

errHand:
    Close lngFileNum
    zlBlobSave = False
End Function

'################################################################################################################
'## 功能：  产生保存指定的文件到指定表记录BLOB字段的SQL语句
'##
'## 参数：  Action      :操作类型（用以区别是操作哪个表）
'##         KeyWord     :确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'##         strFile     :用户指定存放的文件名；不指定时，取当前路径产生文件名
'##         arySql()    :在该数据的基础上扩展增加保存的SQL语句；不指定时，取当前路径产生文件名
'##
'## 返回：  成功返回True，失败返回False
'##
'## 说明：  Action取值说明：
'##         0-病历标记图形；1-病历文件格式；2-病历文件图形；3-病历范文格式；4-病历范文图形；5-电子病历格式；6-电子病历图形；
'################################################################################################################
Public Function zlBlobSql(ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String, ByRef arySql() As String) As Boolean
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, StrText As String
    
    Dim lngLBound As Long, lngUBound As Long    '传入数组的最小最大下标
    Err = 0: On Error Resume Next
    lngLBound = LBound(arySql): lngUBound = UBound(arySql)
    If Err <> 0 Then lngLBound = 0: lngUBound = -1
    Err = 0: On Error GoTo errHand
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    
    ReDim Preserve arySql(lngLBound To lngUBound + lngBlocks + 1) As String
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        StrText = Join(aryHex, "")
        arySql(lngUBound + lngCount + 1) = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "','" & StrText & "'," & IIf(lngCount = 0, 1, 0) & ")"
    Next
    Close lngFileNum
    zlBlobSql = True
    Exit Function

errHand:
    Close lngFileNum
    zlBlobSql = False
End Function

'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String
    
    On Error GoTo errHand
    
    If Not gobjFSO.FileExists(strZipFile) Then zlFileUnzip = "": Exit Function
    
    strZipPath = gobjFSO.GetSpecialFolder(2) '取临时目录
    strZipPathTmp = strZipPath & "\" & Format(Now, "yyMMdd") & CStr(100 * Timer)
    If Not gobjFSO.FolderExists(strZipPathTmp) Then Call gobjFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp & "\TMP.RTF"
    If gobjFSO.FileExists(strZipFileTmp) Then gobjFSO.DeleteFile strZipFileTmp
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    If gobjFSO.FileExists(strZipFileTmp) Then
        
        strZipFileName = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer) & ".RTF"
        If gobjFSO.FileExists(strZipFileName) Then gobjFSO.DeleteFile strZipFileName
                
        Call gobjFSO.CopyFile(strZipFileTmp, strZipFileName)
        
        If gobjFSO.FileExists(strZipFileTmp) Then gobjFSO.DeleteFile strZipFileTmp, True
        
        On Error Resume Next
        If gobjFSO.FolderExists(strZipPathTmp) Then gobjFSO.DeleteFolder strZipPathTmp, True
        
        zlFileUnzip = strZipFileName
    Else
        zlFileUnzip = ""
    End If
    
    Exit Function
    
errHand:
    Call SaveErrLog
End Function

'################################################################################################################
'## 功能：  将文件压缩为新文件放到相同目录中
'## 参数：  strFile     :原始文件
'## 返回：  压缩文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    On Error GoTo errHand
    If Not gobjFSO.FileExists(strFile) Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = gobjFSO.GetParentFolderName(strFile) & "\ZLZIP" & lngCount & ".ZIP"
        If Not gobjFSO.FileExists(strZipFile) Then Exit Do
        lngCount = lngCount + 1
    Loop

    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'################################################################################################################
'## 功能：  替换诊治要素的处理
'##
'## 参数：  ElementName     :替换项目的名称
'##         sPatientID      :病人ID
'##         sPageID         :主页ID或挂号id
'##         iPatientType    :0=门诊、1=住院
'##         lng医嘱ID       :医嘱ID
'##
'## 返回：  返回替换结果
'################################################################################################################
Public Function GetReplaceEleValue(ByVal ElementName As String, _
    ByVal sPatientID As String, _
    ByVal sPageID As String, _
    ByVal iPatientType As PatiFromEnum, _
    ByVal lng医嘱id As Long, Optional lngBabyNum As Long) As String

    Dim rsTmp As New ADODB.Recordset
    
    If ElementName = "单位名称" Then
        GetReplaceEleValue = zl9ComLib.zlRegInfo("单位名称")
        Exit Function
    End If
    
    gstrSQL = "Select Zl_Replace_Element_Value([1],[2],[3],[4],[5],[6]) From Dual"
    Err = 0: On Error GoTo DBError
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "读取替换项", ElementName, CLng(sPatientID), _
        CLng(sPageID), CLng(iPatientType), lng医嘱id, lngBabyNum)
    If rsTmp.EOF Or rsTmp.BOF Then
        GetReplaceEleValue = ""
    Else
        GetReplaceEleValue = Trim(IIf(IsNull(rsTmp.Fields(0).Value), "", rsTmp.Fields(0).Value))
    End If
    Exit Function

DBError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function

'################################################################################################################
'## 功能：  判断指定用户是否是主任医师
'##
'## 参数：  lngUserID       ：用户ID
'##         strUserName     ：用户名
'##         lngPatiID       ：病人ID
'##         lngPatiPageID   ：主页ID
'##
'## 说明：  根据“人员表”中的“聘任技术职务”字段确定医生技术级别（住院医师、主治医师、主任医师）
'##         ＋病人变动记录中的医生级别，从而确定审核级别
'################################################################################################################
Public Function GetUserSignLevel(lngUserID As Long, Optional strUserName As String, _
    Optional lngPatiID As Long, Optional lngPatiPageID As Long) As EPRSignLevelEnum
    Dim rs As New ADODB.Recordset, lngR As Long, lngLevel1 As Long, lngLevel2 As Long
    
    Err = 0: On Error GoTo errHand
    If InStr(gstrPrivsEpr, "签名权") = 0 Then
        GetUserSignLevel = cprSL_空白
        Exit Function
    End If
    
    gstrSQL = "select 聘任技术职务 from 人员表 p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngUserID)
    If Not rs.EOF Then
        lngR = NVL(rs("聘任技术职务"), 0)
    End If
    Select Case lngR    '1 正高  2 副高  3 中级  4 助理/师级  5 员/士  9 待聘
    Case 1: lngLevel1 = cprSL_正高
    Case 2: lngLevel1 = cprSL_主任
    Case 3: lngLevel1 = cprSL_主治
    Case Else: lngLevel1 = cprSL_经治
    End Select
    rs.Close
    
    If lngPatiID > 0 Then
        gstrSQL = "Select 经治医师, 主治医师, 主任医师 " & _
            " From 病人变动记录 " & _
            " Where 病人ID = [1] And 主页ID = [2] And (终止时间 Is Null Or 终止原因 = 1) " & _
            "       And 开始时间 Is Not Null And Nvl(附加床位, 0) = 0"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", lngPatiID, lngPatiPageID)
        If rs.EOF Then
            lngLevel2 = cprSL_经治
        Else
            If rs.Fields("主任医师") = IIf(strUserName = "", gstrUserName, strUserName) Then
                lngLevel2 = cprSL_主任
            ElseIf rs.Fields("主治医师") = IIf(strUserName = "", gstrUserName, strUserName) Then
                lngLevel2 = cprSL_主治
            Else
                lngLevel2 = cprSL_经治
            End If
        End If
    End If
    GetUserSignLevel = IIf(lngLevel1 >= lngLevel2, lngLevel1, lngLevel2)
    Exit Function

errHand:
    GetUserSignLevel = cprSL_空白
End Function

Public Function GetColorVectorG(ByVal lngVersion As Long) As Long
    '根据版本获取RGB颜色中的G颜色分量值
    Select Case lngVersion
    Case 0
        GetColorVectorG = 0     '未开始
    Case 1
        GetColorVectorG = 0     '第一版还不能修订！
    Case 2
        GetColorVectorG = 10
    Case 3
        GetColorVectorG = 90
    Case 4
        GetColorVectorG = 140
    Case 5
        GetColorVectorG = 145
    Case 6
        GetColorVectorG = 150
    Case 7
        GetColorVectorG = 155
    Case 8
        GetColorVectorG = 160
    Case 9
        GetColorVectorG = 165
    Case 10
        GetColorVectorG = 170
    Case 11
        GetColorVectorG = 175
    Case 12
        GetColorVectorG = 180
    Case 13
        GetColorVectorG = 185
    Case 14
        GetColorVectorG = 190
    Case 15
        GetColorVectorG = 195
    Case 16
        GetColorVectorG = 200
    Case 17
        GetColorVectorG = 205
    End Select
End Function

Public Function GetColorVectorB(ByVal lngVersion As Long) As Long
    '根据版本获取RGB颜色中的B颜色分量值
    Select Case lngVersion
    Case 0
        GetColorVectorB = 0     '未终止
    Case 1
        GetColorVectorB = 0     '第一版还不能修订！
    Case 2
        GetColorVectorB = 10
    Case 3
        GetColorVectorB = 15
    Case 4
        GetColorVectorB = 20
    Case 5
        GetColorVectorB = 25
    Case 6
        GetColorVectorB = 30
    Case 7
        GetColorVectorB = 35
    Case 8
        GetColorVectorB = 40
    Case 9
        GetColorVectorB = 45
    Case 10
        GetColorVectorB = 50
    Case 11
        GetColorVectorB = 55
    Case 12
        GetColorVectorB = 60
    Case 13
        GetColorVectorB = 65
    Case 14
        GetColorVectorB = 70
    Case 15
        GetColorVectorB = 75
    Case 16
        GetColorVectorB = 80
    Case 17
        GetColorVectorB = 85
    End Select
End Function

Public Function Get开始版(ByVal COLOR As OLE_COLOR) As Long
    '获取指定颜色的开始版本号，为0表示原始文本颜色
    Dim i As Long
    If COLOR = tomAutoColor Or COLOR = vbBlack Then COLOR = vbBlack: Get开始版 = 1: Exit Function
    For i = 1 To 17
        If GetColorVectorG(i) = rgbGreen(COLOR) Then
            Get开始版 = i
            Exit Function
        End If
    Next
    Get开始版 = 1
End Function

Public Function Get终止版(ByVal COLOR As OLE_COLOR) As Long
    '获取指定颜色的终止版本号，为0表示未结束（保留上次的颜色值）
    Dim i As Long
    If COLOR = tomAutoColor Or COLOR = vbBlack Then COLOR = vbBlack: Get终止版 = 0: Exit Function
    For i = 1 To 17
        If GetColorVectorB(i) = rgbBlue(COLOR) Then
            Get终止版 = i - 1
            Exit Function
        End If
    Next
    Get终止版 = 0
End Function

Public Function GetCharColor(ByVal lng开始版 As Long, ByVal lng终止版 As Long) As OLE_COLOR
    '根据开始版、终止版获取最终字符颜色
    Dim r As Long, g As Long, b As Long
    r = 255
    g = GetColorVectorG(lng开始版)
    b = GetColorVectorB(lng终止版)
    If g = 0 And b = 0 Then
        GetCharColor = vbBlack
    Else
        GetCharColor = RGB(r, g, b)
    End If
End Function


Public Function GetMax(ByVal strTable As String, ByVal strField As String, ByVal intLength As Integer, Optional ByVal strWhere As String) As String
'功能：读取指定表的本级编码的最大值
'参数：strTable  表名;
'      strField  字段名;
'      intLength 字段长度
'返回：成功返回 下级最大编码; 否者返回 0
    Dim rsTemp As New ADODB.Recordset
    Dim varTemp As Variant
    Dim lngLengh As Long
    
    On Error GoTo errHand
    gstrSQL = "SELECT MAX(LPAD(" & strField & "," & intLength & ",' ')) as ""最大值"",max(length(" & _
         strField & ")) as ""最长值"" FROM " & strTable & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR")
    With rsTemp
        If rsTemp.EOF Then
            GetMax = Format(1, String(intLength, "0"))
            Exit Function
        End If
        varTemp = IIf(IsNull(.Fields("最大值").Value), "0", .Fields("最大值").Value)
        lngLengh = IIf(IsNull(.Fields("最长值").Value), intLength, .Fields("最长值").Value)
        If IsNumeric(varTemp) Then
            GetMax = CStr(Val(varTemp) + 1)
            GetMax = Format(GetMax, String(lngLengh, "0"))
        Else
            gstrSQL = "Select ZL_INCSTR([1]) As MAXVALUE From Dual"
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", CStr(varTemp))
            If rsTemp.BOF = False Then
                GetMax = Trim(rsTemp("MAXVALUE").Value)
            End If
        End If
        .Close
    End With
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function CommandBarInit(ByRef cbsMain As CommandBars) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    CommandBarInit = True
    
End Function

Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    DockPannelInit = True
    
End Function

Public Function CommandBarExecutePublic(Control As Object, frmMain As Object, Optional ByVal objPrnVsf As Object, Optional ByVal strPrintTitle As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    Dim bytMode As Byte
        
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet              '打印设置
    
        Call zlPrintSet
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel               '打印数据,预览数据,输出到Excel
        
'        If objPrnVsf Is Nothing Then Exit Function
'
'        Call SearchPrintData(objPrnVsf, frmPubResource.msfPrint)
'
'        '调用打印部件处理
'        Set objPrint.Body = frmPubResource.msfPrint
'        objPrint.Title.Text = strPrintTitle
'        Set objAppRow = New zlTabAppRow
'        Call objAppRow.Add("")
'        Call objAppRow.Add("打印时间:" & Now())
'        Call objPrint.BelowAppRows.Add(objAppRow)
'
'        Select Case Control.ID
'        Case conMenu_File_Print
'            bytMode = zlPrintAsk(objPrint)
'            If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
'        Case conMenu_File_Preview
'            zlPrintOrView1Grd objPrint, 2
'        Case conMenu_File_Excel
'            zlPrintOrView1Grd objPrint, 3
'        End Select
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '工具栏
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text      '按钮文字
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.STYLE = IIf(objControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size      '大图标
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar         '状态栏
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
    
    Case conMenu_Help_Help              '帮助主题
    
        Call ShowHelp(App.ProductName, frmMain.hWnd, frmMain.Name, Int((glngSys) / 100))
        
    Case conMenu_Help_Web_Home          'Web上的中联
        
        Call zlHomePage(frmMain.hWnd)
        
    Case conMenu_Help_Web_Forum         'Web上的论坛
    
        Call zlWebForum(frmMain.hWnd)
        
    Case conMenu_Help_Web_Mail          '发送反馈
        
        Call zlMailTo(frmMain.hWnd)
            
    Case conMenu_Help_About             '关于
        
        Call ShowAbout(frmMain, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    
    Case conMenu_File_Exit              '退出
    
        Unload frmMain
            
    End Select
    
    CommandBarExecutePublic = True
End Function

Public Function CommandBarUpdatePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Select Case Control.ID
    Case conMenu_View_ToolBar_Button            '工具栏
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = frmMain.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              '图标文字
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = Not (frmMain.cbsMain(2).Controls(1).STYLE = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '大图标
        Control.Checked = frmMain.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '状态栏
        Control.Checked = frmMain.stbThis.Visible
    End Select
    
    CommandBarUpdatePublic = True
End Function

Public Function NewCommandBar(objMenu As CommandBarControl, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngId As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal strParameter As String) As CommandBarControl
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpType, lngId, strCaption)
        
        objControl.IconId = IIf(lngIcon = -1, lngId, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        objControl.Parameter = strParameter
        
    End With
    
    Set NewCommandBar = objControl
    
End Function

Public Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngId As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer) As CommandBarControl
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngId, strCaption, intBefore)
        objControl.ID = lngId
        objControl.IconId = IIf(lngIcon = -1, lngId, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.Type = xtpControlButton Or objControl.Type = xtpControlPopup Then
            objControl.STYLE = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '功能：生成字符串的简码
    '入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
    '出参：正确返回字符串；错误返回"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    
    If bytIsWB Then
        gstrSQL = "Select zlWBcode('" & strInput & "') from dual"
    Else
        gstrSQL = "Select zlSpellcode('" & strInput & "') from dual"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR")
    zlGetSymbol = NVL(rsTmp.Fields(0).Value)
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlGetSymbol = "-"
End Function

Public Function IsPrivs(ByVal strPrivs As String, ByVal strPriv As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If InStr(";" & strPrivs & ";", ";" & strPriv & ";") > 0 Then
        IsPrivs = True
    Else
        IsPrivs = False
    End If
End Function
Public Function Get会诊文件ID(ByVal lngRecordId As Long, ByVal lngAdviceID As Long) As String
Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    If lngAdviceID <> 0 Then
        gstrSQL = "Select 病历id From 病人医嘱报告 A Where 医嘱id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取相同医嘱ID的会诊记录", lngAdviceID)
    Else
        gstrSQL = "Select b.病历id From 病人医嘱报告 A, 病人医嘱报告 B Where a.病历id = [1] And a.医嘱id = b.医嘱id"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取相同医嘱ID的会诊记录", lngRecordId)
    End If
    
    Do Until rsTemp.EOF
        Get会诊文件ID = Get会诊文件ID & "," & rsTemp!病历ID
        rsTemp.MoveNext
    Loop
    If Len(Get会诊文件ID) > 0 Then
        Get会诊文件ID = Mid(Get会诊文件ID, 2)
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function GetFileRange(ByVal lFileId As Long, ByVal lngRecordId As Long, ByVal strCreateTime As String, _
                            ByVal eDocType As EPRDocTypeEnum, ByVal lngPatId As Long, ByVal lngPageId As Long, _
                            Optional ByVal blnMoved As Boolean, Optional ByVal lngAdviceID As Long) As String
    '******************************************************************************************************************
    '功能：读取当前病历(可能当前病历未保存)前所有共享病历
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTemp As New ADODB.Recordset, strTime As String, dStar As Date, dEnd As Date
    Dim strSQL As String, strIDs As String, blnNewPage As Boolean, n_Num As Integer, n_S As Integer, n_E As Integer

    On Error GoTo errHand
    strIDs = Get会诊文件ID(lngRecordId, lngAdviceID)
    If strIDs <> "" Then GetFileRange = strIDs: Exit Function
    
    strTime = Format(strCreateTime, "yyyy-MM-dd HH:mm:ss")
    If strTime = "" Then strTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If strTime = "00:00:00" Then strTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    blnNewPage = zlDatabase.GetPara("转科后要求书写的共享病历另起一页打印", glngSys, 1251, 1) = 1 '=0表示转科后共享病历连续打印 =1 表示另起一页打印

    strSQL = "Select m.Id" & vbNewLine & _
            "From 病历文件列表 L, 病历文件列表 M" & vbNewLine & _
            "Where l.Id = [1] And l.页面 = m.编号 And l.页面 = m.页面 And " & vbNewLine & _
            "      m.种类 In (" & Decode(eDocType, 4, "4", 1, "1", "2, 5, 6") & ") And l.种类 = m.种类"
    If blnNewPage And eDocType = 2 Then '住院病历转科后另起一页
    strSQL = strSQL & vbNewLine & _
            "Union" & vbNewLine & _
            "Select b.Id" & vbNewLine & _
            "From 病历文件列表 A, 病历文件列表 B, 病历时限要求 C" & vbNewLine & _
            "Where a.Id = [1] And a.页面 = b.页面 And b.种类 In (" & Decode(eDocType, 4, "4", 1, "1", "2, 5, 6") & ") And " & vbNewLine & _
            "      a.种类 = b.种类 And c.文件id = b.Id And c.事件 = '转科' And c.书写时限 >= 0"
    End If
    
    If lngRecordId <> 0 Then
        gstrSQL = "Select nvl(序号,0) 序号 From 电子病历记录 Where ID=[1]"
        If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取序号", lngRecordId)
        n_Num = rsTemp!序号
    Else
        n_Num = 9999
    End If
    
    '提取页面文件当前时间之前最近一次书写记录
    gstrSQL = "Select 创建时间,序号" & vbNewLine & _
                "From (Select a.创建时间,a.序号" & vbNewLine & _
                "       From 电子病历记录 A, (" & strSQL & ") B" & vbNewLine & _
                "       Where a.文件id = b.Id And a.病人id = [2] And a.主页id = [3] And " & IIf(n_Num = 0, "a.创建时间 <= [4]", "((a.创建时间 <= [4] and Nvl(a.序号,0)=0) Or  a.序号<=[5])") & vbNewLine & _
                "       Order By " & IIf(n_Num <> 0, "a.序号", "a.创建时间") & " Desc)" & vbNewLine & _
                "Where Rownum = 1"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取页面文件之前一次书写", lFileId, lngPatId, lngPageId, CDate(strTime), n_Num)
    If rsTemp.EOF Then
        dStar = CDate("2000-01-01 00:00:00"): n_S = -1 '表明之前没有写过页面文件
    Else
        dStar = CDate(rsTemp!创建时间)
        n_S = IIf(n_Num = 0, 0, NVL(rsTemp!序号, -1))
    End If
    
    '提取页面文件当前时间之后最近一次书写记录
    gstrSQL = "Select 创建时间,序号" & vbNewLine & _
                "From (Select a.创建时间,a.序号" & vbNewLine & _
                "       From 电子病历记录 A, (" & strSQL & ") B" & vbNewLine & _
                "       Where a.文件id = b.Id And a.病人id = [2] And a.主页id = [3] And " & IIf(n_Num = 0, "a.创建时间 > [4]", "((a.创建时间 > [4] and Nvl(a.序号,0)=0) Or a.序号>[5])") & vbNewLine & _
                "       Order By a.创建时间)" & vbNewLine & _
                "Where Rownum = 1"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取页面文件之后一次书写", lFileId, lngPatId, lngPageId, CDate(strTime), n_Num)
    If rsTemp.EOF Then
        dEnd = CDate("3001-01-01") '表明之后没有写过页面文件
        n_E = 9999
    Else
        dEnd = CDate(rsTemp!创建时间) - 1 / 24 / 60 / 60 '表明之后写过，取近记录的时间减一秒,即不包含此记录
        n_E = IIf(n_Num = 0, 0, NVL(rsTemp!序号, 9999) - 1)
    End If
    
    '具有相同共享属性病历的书写记录
    strSQL = "Select m.Id" & vbNewLine & _
        "From 病历文件列表 L, 病历文件列表 M" & vbNewLine & _
        "Where l.Id = [1] And l.页面 = m.页面 And " & vbNewLine & _
        "      m.种类 In (" & Decode(eDocType, 4, "4", 1, "1", "2, 5, 6") & ") And l.种类 = m.种类"
    gstrSQL = "Select a.ID" & vbNewLine & _
                "   From 电子病历记录 A, (" & strSQL & ") B" & vbNewLine & _
                "   Where a.文件id = b.Id And a.病人id = [2] And a.主页id = [3] And " & IIf(n_Num = 0, "a.创建时间 Between [4] And [5]", "((a.创建时间 Between [4] And [5] and Nvl(a.序号,0)=0 ) or a.序号 Between [6] And [7])") & vbNewLine & _
                "   Order By a.创建时间"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取页面文件之后一次书写", lFileId, lngPatId, lngPageId, dStar, dEnd, n_S, n_E)
    Do Until rsTemp.EOF
        strIDs = strIDs & "," & rsTemp!ID
        rsTemp.MoveNext
    Loop
    strIDs = Mid(strIDs, 2)
    
    GetFileRange = strIDs
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function SetPaneRange(dkpMain As Object, ByVal intPane As Integer, ByVal lngMinW As Long, lngMinH As Long, lngMaxW As Long, lngMaxH As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objPan As Pane
    
    On Error Resume Next
    
    Set objPan = dkpMain.FindPane(intPane)
    
    If objPan Is Nothing Then Exit Function
    With objPan
        .MaxTrackSize.SetSize lngMaxW, lngMaxH
        .MinTrackSize.SetSize lngMinW, lngMinH
    End With
    
    SetPaneRange = True
End Function

Public Function CreateTmpFile(Optional ByVal strFileType As String = "tmp", Optional ByVal strName As String, Optional ByVal blnTime As Boolean = True) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strFile As String
    Dim strFileTemp As String
    Dim lngTemp As Long
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    
    strFileTemp = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
    
    If blnTime Then
        strFileTemp = strFileTemp & strName & Format(Now, "yyyymmdd") & Format(Timer, "0") & "." & strFileType
    Else
        strFileTemp = strFileTemp & strName & "." & strFileType
    End If
    
    CreateTmpFile = strFileTemp
    
End Function

Public Function RemoveSign(ByRef edtThis As Editor, ByRef objDocument As cEPRDocument) As Boolean
'******************************************************************************************************************
'功能：从打印/预览文档中移出签名内容及其前缀
'******************************************************************************************************************
Dim intLoop As Integer, strFoot As String, strHead As String
Dim lESS As Long, lESE As Long, lEES As Long, lEEE As Long, blnNeeded As Boolean, blnFinded As Boolean
Dim strAllSign As String, strFSign As String, strSSign As String, strTSign As String
    
    On Error GoTo errHand
    strFoot = edtThis.FootFileText: strHead = edtThis.HeadFileText
    edtThis.ForceEdit = True
    If InStr(strFoot, "{书写签名}") > 0 Or InStr(strFoot, "{医生签名}") > 0 Or InStr(strFoot, "{主治签名}") > 0 Or InStr(strFoot, "{主任签名}") > 0 Or _
        InStr(strHead, "{书写签名}") > 0 Or InStr(strHead, "{医生签名}") > 0 Or InStr(strHead, "{主治签名}") > 0 Or InStr(strHead, "{主任签名}") > 0 Then
        '查找并隐藏原来的签名
        For intLoop = 1 To objDocument.Signs.Count
            blnFinded = False
            blnFinded = FindKey(edtThis, "S", objDocument.Signs(intLoop).Key, lESS, lESE, lEES, lEEE, blnNeeded)
            If blnFinded Then
                Select Case objDocument.Signs(intLoop).签名级别
                    Case Is <= cprSL_经治
                        strFSign = strFSign & " " & edtThis.Range(lESS + 16, lEES).Text
                    Case cprSL_主治
                        strSSign = strSSign & " " & edtThis.Range(lESS + 16, lEES).Text
                    Case Is >= cprSL_主任
                        strTSign = strTSign & " " & edtThis.Range(lESS + 16, lEES).Text
                End Select
                edtThis.Range(lESE, lEES).Text = ""
            End If
        Next
        
'        For intLoop = 1 To objDocument.Elements.Count
'            If objDocument.Elements(intLoop).替换域 = 1 Then
'                Select Case objDocument.Elements(intLoop).要素名称
'                Case "经治医师签名"
'                    strFSign = strFSign & " " & objDocument.Elements(intLoop).内容文本
'                Case "主治医师签名"
'                    strSSign = strSSign & " " & objDocument.Elements(intLoop).内容文本
'                Case "主任医师签名"
'                    strTSign = strTSign & " " & objDocument.Elements(intLoop).内容文本
'                End Select
'            End If
'        Next
    End If
    edtThis.ForceEdit = False
    strFSign = Mid(strFSign, 2): strSSign = Mid(strSSign, 2): strTSign = Mid(strTSign, 2)
    strAllSign = strFSign & " " & strSSign & " " & strTSign
    objDocument.EPRPatiRecInfo.医生签名 = strFSign
    objDocument.EPRPatiRecInfo.主治签名 = strSSign
    objDocument.EPRPatiRecInfo.主任签名 = strTSign
    objDocument.EPRPatiRecInfo.书写签名 = strAllSign
    RemoveSign = True
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SetHeadFoot(edtThis As Editor, ByVal lngFileID As Long)
'从数据库读取记录刷新页眉页脚
'格式=PaperKind;PaperOrient;PaperHeight;PaperWidth;MarginLeft;MarginRight;MarginTop;MarginBottom;BackColor;PaperColor;ShowPageNumber;页眉格式;页脚格式

Dim strFile As String, lngType As Long, strPage As String, rsTemp As New ADODB.Recordset
    gstrSQL = "Select a.种类, a.编号, a.格式, a.页眉, a.页脚" & vbNewLine & _
                "From 病历页面格式 A, 病历文件列表 B" & vbNewLine & _
                "Where b.Id = [1] And a.种类 = b.种类 And b.页面 = a.编号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取页眉页脚", lngFileID)
    If rsTemp.EOF Then Exit Sub
    If NVL(rsTemp!格式) = "" Then Exit Sub
    
    With edtThis
        .PaperKind = Split(rsTemp!格式, ";")(0)
        .PaperOrient = Split(rsTemp!格式, ";")(1)
        If UBound(Split(rsTemp!格式, ";")) > 10 Then
        .HeadFontFormat = Split(rsTemp!格式, ";")(11)
        .FootFontFormat = Split(rsTemp!格式, ";")(12)
        End If
        .PaperHeight = Split(rsTemp!格式, ";")(2)
        .PaperWidth = Split(rsTemp!格式, ";")(3)
        .MarginLeft = Split(rsTemp!格式, ";")(4)
        .MarginRight = Split(rsTemp!格式, ";")(5)
        .MarginTop = Split(rsTemp!格式, ";")(6)
        .MarginBottom = Split(rsTemp!格式, ";")(7)
    
        strFile = zlBlobRead(7, rsTemp!种类 & "-" & rsTemp!编号) '读取页眉图片
        If gobjFSO.FileExists(strFile) Then
            Set .Picture = LoadPicture(strFile)
            gobjFSO.DeleteFile strFile, True      '删除临时文件
        End If
        
        strFile = zlBlobRead(12, rsTemp!种类 & "-" & rsTemp!编号, App.Path & "\Head.rtf") '读取页眉文件
        If gobjFSO.FileExists(strFile) Then
            edtThis.HeadFile = strFile           '读取文件
            gobjFSO.DeleteFile strFile, True      '删除临时文件
            If Trim(edtThis.HeadFileText) = "" Then GoTo Headtxt
        Else
Headtxt:
            If NVL(rsTemp!页眉) <> "" Then
                edtThis.Head = rsTemp!页眉
                edtThis.HeadTextToFile '将文字读入Rtf控件中
            End If
        End If
        
        strFile = zlBlobRead(13, rsTemp!种类 & "-" & rsTemp!编号, App.Path & "\Foot.rtf") '读取页脚文件
        If gobjFSO.FileExists(strFile) Then
            edtThis.FootFile = strFile            '读取文件
            gobjFSO.DeleteFile strFile, True      '删除临时文件
            If Trim(edtThis.FootFileText) = "" Then GoTo Foottxt
        Else
Foottxt:
            If NVL(rsTemp!页脚) <> "" Then
                edtThis.Foot = rsTemp!页脚
                edtThis.FootTextToFile '将文字读入Rtf控件中
            End If
        End If
    End With
End Sub
Public Sub ReadRTF(edtThis As Editor, ByVal lngFileID As Long, ByVal blnClearMode As Boolean, ByVal blnMoved As Boolean, Optional ByVal blnClearBColor As Boolean = True)
'读取RTF文件
Dim ParaFmt As New cParaFormat, FontFmt As New cFontFormat, i As Long, lngLen As Long
Dim oEles As New cEPRElements, oTabs As New cEPRTables, oPics As New cEPRPictures
Dim strZipFile As String, strRtfFile As String, j As Long, rs As New ADODB.Recordset
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    strZipFile = zlBlobRead(5, lngFileID, , blnMoved)
    If gobjFSO.FileExists(strZipFile) Then
        strRtfFile = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strRtfFile) Then
                edtThis.OpenDoc strRtfFile
                gobjFSO.DeleteFile strRtfFile, True
        End If
        gobjFSO.DeleteFile strZipFile, True
    End If
    If Trim(edtThis.Text) = "" Then Exit Sub

    '读取图片,表格,要素
    gstrSQL = "Select Level,ID, 文件id,开始版, 终止版," & _
                "   父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 定义提纲id, 复用提纲," & vbNewLine & _
                "       使用时机, 诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
                "From (Select ID, 文件id,开始版, 终止版," & _
                "               父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id,定义提纲id," & vbNewLine & _
                "              复用提纲, 使用时机, 诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
                "       From 电子病历内容" & vbNewLine & _
                "       Where 文件id = [1] And 对象序号<>ID)" & vbNewLine & _
                "Start With 父id Is Null" & vbNewLine & _
                "Connect By Prior ID = 父id" & vbNewLine & _
                "Order By 对象序号, 内容行次"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历内容", "H电子病历内容")
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "还原表图", lngFileID)
    Do Until rs.EOF
        Select Case rs!对象类型
            Case 3  '表格
                lKey = oTabs.Add(NVL(rs!对象标记, 0))                  '恢复Key值！
                Call oTabs("K" & lKey).FillTableMember(rs, IIf(blnMoved, "H电子病历内容", "电子病历内容"))
            Case 4  '要素
                lKey = oEles.Add(NVL(rs!对象标记, 0))
                Call oEles("K" & lKey).FillElementMember(rs, IIf(blnMoved, "H电子病历内容", "电子病历内容"))
            Case 5  '图片
                lKey = oPics.Add(NVL(rs("对象标记"), 0))
                Call oPics("K" & lKey).FillPictureMember(rs, IIf(blnMoved, "H电子病历内容", "电子病历内容"))
        End Select
        rs.MoveNext
    Loop
    
    For j = 1 To oPics.Count '还原图片
        bFinded = FindKey(edtThis, "P", oPics(j).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            oPics(j).DeleteFromEditor edtThis
            oPics(j).InsertIntoEditor edtThis, -1, True
        End If
    Next
    
    For j = 1 To oTabs.Count '还原表格
        bFinded = FindKey(edtThis, "T", oTabs(j).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
                Set ParaFmt = edtThis.Range(lKSE, lKES).Para.GetParaFmt
                Set FontFmt = edtThis.Range(lKSE, lKES).Font.GetFontFmt
                
                If oTabs(j).是否换行 Then
                    edtThis.Range(lKSS, lKEE + 2).Text = ""
                Else
                    edtThis.Range(lKSS, lKEE).Text = ""
                End If
                oTabs(j).InsertIntoEditor edtThis, lKSS, , , True
                
                edtThis.Range(lKSE, lKES).Para.SetParaFmt ParaFmt
                edtThis.Range(lKSE, lKES).Font.SetFontFmt FontFmt
                edtThis.Range(lKSS, lKEE).Font.Protected = True
        End If
    Next

    ' 对最终文档的处理
    edtThis.SelectAll
    If blnClearMode Then
        edtThis.AuditMode = True
        edtThis.AcceptAuditText    '清洁模式
    End If
    
    If blnClearBColor Then
        For j = 1 To oEles.Count '删除空的要素，展开型要素重新刷新以去掉下波浪线
            If oEles(j).内容文本 = "" Then
                oEles(j).DeleteFromEditor edtThis
            ElseIf oEles(j).输入形态 = 1 Then
                oEles(j).Refresh edtThis
            End If
        Next
        
        lngLen = Len(edtThis.Text)
        For i = 0 To lngLen - 1 '只将背景色为要素背景色颜色去掉
            If edtThis.Range(i, i + 1).Font.BackColor = ELE_BACKCOLOR Then
                edtThis.Range(i, i + 1).Font.BackColor = tomAutoColor
            End If
            If edtThis.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR Then
                edtThis.Range(i, i + 1).Font.ForeColor = tomAutoColor
            End If
        Next
    End If
End Sub
Public Sub BuildRTF(edtThis As Editor, ByVal lngFileID As Long, ByVal blnMoved As Boolean)
'无RTF文件时，读取电子病历内容仅用显示，无法编辑
Dim strContent As String, rs As New ADODB.Recordset
    gstrSQL = "Select ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行" & vbNewLine & _
            "From (Select ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行" & vbNewLine & _
            "       From 电子病历内容" & vbNewLine & _
            "       Where 文件id = [1] And 对象序号 <> ID And 终止版 = 0)" & vbNewLine & _
            "Start With 父id Is Null" & vbNewLine & _
            "Connect By Prior ID = 父id" & vbNewLine & _
            "Order By 对象序号, 内容行次"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历内容", "H电子病历内容")
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "读取内容", lngFileID)
        
    Do Until rs.EOF
        If rs!对象类型 <> 1 Then '提纲不加载显示
            strContent = strContent & IIf(rs!是否换行 = 1, vbCrLf, "") & rs!内容文本
        End If
        rs.MoveNext
    Loop
    
    edtThis.Text = strContent
End Sub
Public Sub ReplacedHeadFootString(ByRef edtThis As Object, ByVal lngRecId As Long, ByVal blnMoved As Boolean)
'功能： 页眉/页脚中的替换要素内容
'参数： objDoc对象，如果类型是Long,那么传入的是电子病历记录ID
Dim strElements As String, j As Long, aryEle() As String, strEleValue As String
Dim lngStartPos As Long, lngEndPos As Long
Dim strHead As String, strFoot As String
Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select a.病历名称, a.完成时间, a.病人id, a.主页id, a.病人来源, b.名称 书写部门, c.内容文本 As 书写签名, d.医嘱id" & vbNewLine & _
            "From 电子病历记录 A, 部门表 B, 电子病历内容 C, 病人医嘱报告 D" & vbNewLine & _
            "Where a.Id =[1] And a.Id = c.文件id(+) And a.科室id = b.Id And c.对象类型(+) = 8 And c.开始版(+) = 1 And a.Id = d.病历id(+)"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历内容", "H电子病历内容")
    If blnMoved Then gstrSQL = Replace(gstrSQL, "病人医嘱报告", "H病人医嘱报告")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取信息", lngRecId)
    
    '从strHead中分析出替换要素，并放到aryEle数组中
    '------------------------------------------------------------------------------------------------------------------
    strHead = edtThis.HeadFileText
    lngStartPos = 0
    lngEndPos = 0
    strElements = ""
    For j = 1 To Len(strHead)
        If Mid(strHead, j, 1) = "{" Then lngStartPos = j
        If Mid(strHead, j, 1) = "}" Then lngEndPos = j

        If lngStartPos > 0 And lngEndPos > 0 Then
            If lngEndPos > lngStartPos + 1 Then
                strElements = strElements & ";" & Mid(strHead, lngStartPos + 1, lngEndPos - lngStartPos - 1)
            End If

            lngStartPos = 0
            lngEndPos = 0
        End If
    Next

    '从strFoot中分析出替换要素，并放到aryEle数组中
    '------------------------------------------------------------------------------------------------------------------
    strFoot = edtThis.FootFileText
    lngStartPos = 0
    lngEndPos = 0
    For j = 1 To Len(strFoot)
        If Mid(strFoot, j, 1) = "{" Then lngStartPos = j
        If Mid(strFoot, j, 1) = "}" Then lngEndPos = j

        If lngStartPos > 0 And lngEndPos > 0 Then
            If lngEndPos > lngStartPos + 1 Then
                strElements = strElements & ";" & Mid(strFoot, lngStartPos + 1, lngEndPos - lngStartPos - 1)
            End If

            lngStartPos = 0
            lngEndPos = 0
        End If
    Next
    If strElements <> "" Then
        strElements = Mid(strElements, 2)
    Else
        Exit Sub
    End If
    aryEle = Split(strElements, ";")
    
    '对于医生签名、主治签名、主任签名，只会出现在诊疗报告中，需要结合电子病历内容清除Edit中的签名信息并转移到页眉页脚中
    '目前因本函数只适用于共享病历，暂不处理
    
    For j = 0 To UBound(aryEle)
        Select Case aryEle(j)
            Case "病历名称"
                strEleValue = rsTemp!病历名称
            Case "书写部门"
                strEleValue = rsTemp!书写部门
            Case "完成时间"
                strEleValue = Format(NVL(rsTemp!完成时间), "yyyy-MM-dd hh:mm")
            Case "书写签名"
                strEleValue = NVL(rsTemp!书写签名)
            Case "页码", "总页数", "标题", "文件名", "路径", "打印日期", "打印时间"
                strEleValue = "" '本部份关键字由控件内部替换
            Case Else
                strEleValue = GetReplaceEleValue(aryEle(j), rsTemp!病人ID, rsTemp!主页ID, NVL(rsTemp!病人来源, 0), NVL(rsTemp!医嘱id, 0))
        End Select
        
        If strEleValue <> "" Then '取到值,替换值
            Call edtThis.DocHeadReplaceKey("{" & aryEle(j) & "}", strEleValue)
            Call edtThis.DocFootReplaceKey("{" & aryEle(j) & "}", strEleValue)
        End If
    Next
End Sub

'################################################################################################################
'## 功能：  将文件压缩为新文件放到相同目录中并删除XML文件
'## 参数：  strFiles     :原始文件路径字符串，（多个以“，”分隔开）
'## 参数：  strZipPath   :压缩后的文件路径
'## 返回：  压缩文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFilesZip(ByVal strFiles As String, ByVal strZipPath As String) As String
    Dim strZipFile As String, strFile As Variant
    Dim lngFileNum As Long, lngFile As Long, i As Long, j As Long
    Dim aryChunk() As Byte, bytTmp As Byte
    On Error GoTo errHand:
    strFile = Split(strFiles, ",")
        With mclsZip
            .Encrypt = False: .AddComment = False
            .ZipFile = strZipPath
            .StoreFolderNames = False
            .RecurseSubDirs = False
            .ClearFileSpecs
            For i = 0 To UBound(strFile)
              .AddFileSpec strFile(i)
            Next i
            .Zip
            If (.Success) Then
                zlFilesZip = .ZipFile
            Else
                zlFilesZip = ""
            End If
            '删除XML文件
            For i = 0 To UBound(strFile)
                gobjFSO.DeleteFile (strFile(i))
            Next i
        End With
        Exit Function
errHand:
        zlFilesZip = ""
End Function
'################################################################################################################
'## 功能：  在XML压缩文件相同目录释放产生解压XML文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFilesUnZip(ByVal strZipFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String, strZipFileTmp2 As String
    Dim strZipFileName As String, strUnZipFile As File, strUnZipFileName As String
    Dim lngFileNum As Long   ' 声明变量。
    Dim aryChunk() As Byte, lngFile As Long, bytTmp As Byte
    Dim i As Long
    On Error GoTo errHand
    
    If Not gobjFSO.FileExists(strZipFile) Then zlFilesUnZip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    strZipPath = zl9ComLib.OS.TempPath
    strZipPathTmp = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer)
    Call gobjFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp & "\TMP.XML"
    If gobjFSO.FileExists(strZipFileTmp) Then gobjFSO.DeleteFile strZipFileTmp
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    For Each strUnZipFile In gobjFSO.GetFolder(strZipPathTmp).Files
        If InStr(1, strUnZipFile.Name, "范文列表") > 0 Then
            strUnZipFileName = strZipPathTmp & "\" & strUnZipFile.Name
        End If
        If InStr(1, strUnZipFile.Name, ".xml") > 0 And InStr(1, strUnZipFile.Name, "范文列表") < 1 Then
            strUnZipFileName = strZipPathTmp & "\" & strUnZipFile.Name
        End If
        If InStr(1, strUnZipFile.Name, "TMP") > 0 Then
            '删除解压后的.ZIP文件
            gobjFSO.DeleteFile (strZipPathTmp & "\" & strUnZipFile.Name)
        End If
    Next
    zlFilesUnZip = strUnZipFileName
    Exit Function
errHand:
    Call SaveErrLog
End Function
Public Sub VerifyPatiSign(ByVal frmParent As Object, ByVal lFileId As Long, ByVal blnMoved As Boolean)
    '功能:根据传入的病历文件ID提取患者签名相关信息，调用验证接口
Dim rsTemp As New ADODB.Recordset, lngSignID As Double
Dim strSource As String, strName As String, strIdentifyNo As String, strOtherParms As String, strSignInfo As String, strPenSignBase64 As String
    On Error GoTo errHand
    gstrSQL = "Select ID,对象标记,对象属性,内容文本" & vbNewLine & _
                "From 电子病历内容" & vbNewLine & _
                "Where 文件id = [1] And 对象类型 = 5"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历内容", "H电子病历内容")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取患者签名记录", lFileId)
    If rsTemp.EOF Then MsgBox "当前所选病历没有有效的数字签名记录！", vbInformation, gstrSysName: Exit Sub
    If Split(rsTemp!对象属性, ";")(0) <> 5 Then MsgBox "当前病历患者签名，无需验证！", vbInformation, gstrSysName: Exit Sub
    
    If UBound(Split(rsTemp!内容文本, "|")) > 2 Then
        strName = Split(rsTemp!内容文本, "|")(0)
        strIdentifyNo = Split(rsTemp!内容文本, "|")(1)
        strSignInfo = rsTemp!内容文本
    End If
    strSource = GetPatiSignSource(lFileId, blnMoved)
    If strSource = "" Then Exit Sub
    
    If gobjESign Is Nothing Then
        Set gobjESign = CreateObject("zl9ESign.clsESign")
        Call gobjESign.Initialize(gcnOracle, glngSys)
    End If
    If gobjESign.EnabledVerifyPatiSign() = False Then
        MsgBox "当前接口不支持患者签名验证。", vbInformation, gstrSysName
    End If
    Call gobjESign.ValidatePenSignature(strSource, strName, strIdentifyNo, strOtherParms, strSignInfo, strPenSignBase64)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetPatiSignSource(ByVal lFileId As Long, ByVal blnMoved As Boolean) As String
'功能:根据文件ID,提取签名源内容,用于在不打开编辑器情况下进行签名验证
'步骤:1 去掉最后一次签名相邻的签名图片 (如果有) 图片的组成字为 " PS(0000000X,0,1) PE(0000000X,0,1) "
'     2 去掉所有S关键字的签名对象
'     3 将所有图片,表格关键字中间的"口"字换成空格(因为签名时是空格表示)
'     4 将签名要素还原，签名时是"{经治医师签名}""主治医师签名""主任医师签名"
    Dim lSS As Long, lSE As Long, lES As Long, lEE As Long, bNeeded As Boolean, bFinded As Boolean, lKey As Long, lPos As Long
    Dim strZipFile As String, strRtfFile As String, lSEKey As Long
    Dim rsTemp As New ADODB.Recordset, strSource As String
    
    strZipFile = zlBlobRead(5, lFileId, , blnMoved)
    If gobjFSO.FileExists(strZipFile) Then
        strRtfFile = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strRtfFile) Then
                gfrmPublic.edtBuff.OpenDoc strRtfFile
                gobjFSO.DeleteFile strRtfFile, True
        End If
        gobjFSO.DeleteFile strZipFile, True
    End If
    
    gfrmPublic.edtBuff.Freeze
    gfrmPublic.edtBuff.ForceEdit = True
    
    '去掉所有S关键字的签名对象
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "S", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then
            If gfrmPublic.edtBuff.Range(lEE, lEE + 4).Text = "PS(" Then
                gfrmPublic.edtBuff.Range(lEE, lEE + 35).Text = "" '签名相邻的签名图片,签名图片紧贴签名关键字
            End If
            gfrmPublic.edtBuff.Range(lSS, lEE).Text = ""
        End If
    Loop Until bFinded = False
    
    '将所有表格关键字中间的"口"字换成空格(因为签名时是空格表示)
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "T", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSE, lES).Text = " ": lPos = lEE + 1
    Loop Until bFinded = False
    
    '将所有图片关键字中间的"口"字换成空格(因为签名时是空格表示)
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "P", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSE, lES).Text = " ": lPos = lEE + 1
    Loop Until bFinded = False
    
    '处理签名要素,因为签名时使用的源文中签名要素是"{经治医师签名}"形式，但签名后被更改为具体的姓名
    gstrSQL = "Select 对象标记,要素名称" & vbNewLine & _
            "From 电子病历内容" & vbNewLine & _
            "Where 文件id = [1] And 对象类型 = 4 And 要素名称 In ('经治医师签名', '主治医师签名', '主任医师签名')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取签名要素名称", lFileId)
    Do Until rsTemp.EOF  '有则使用签名要素名称还原
        lPos = 0
        lSEKey = rsTemp!对象标记
        bFinded = FindKey(gfrmPublic.edtBuff, "E", lSEKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then
            If gfrmPublic.edtBuff.Range(lEE, lEE + 4).Text = " PS(" Then '使用了签名图片
                gfrmPublic.edtBuff.Range(lEE, lEE + 35).Text = ""
            End If
            gfrmPublic.edtBuff.Range(lSE, lES).Text = "{" & rsTemp!要素名称 & "}"
        End If
        rsTemp.MoveNext
    Loop
    
    gfrmPublic.edtBuff.ForceEdit = False
    gfrmPublic.edtBuff.UnFreeze
    strSource = gfrmPublic.edtBuff.Text
    strSource = Replace(strSource, Chr(32), "")
    strSource = Replace(strSource, vbCr, "")
    strSource = Replace(strSource, vbLf, "")
    GetPatiSignSource = strSource
End Function

Public Sub VerifySignature(ByVal frmParent As Object, ByVal lFileId As Long, ByVal blnMoved As Boolean)
'功能:根据传入的病历文件ID提取签名记录原始信息，调用验证接口
Dim rsTemp As New ADODB.Recordset, lngSignID As Double, strSource As String
    On Error GoTo errHand
    gstrSQL = "Select ID,对象标记,对象属性" & vbNewLine & _
                "From (Select ID, 对象标记,对象属性 From 电子病历内容 Where 文件id = [1] And 对象类型 = 8 Order By 对象标记 Desc)" & vbNewLine & _
                "Where Rownum = 1"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历内容", "H电子病历内容")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取最后签名记录", lFileId)
    If rsTemp.EOF Then MsgBox "当前所选病历没有有效的数字签名记录！", vbInformation, gstrSysName: Exit Sub
    If Split(rsTemp!对象属性, ";")(0) <> 2 Then MsgBox "当前病历签名不是数字签名，无需验证！", vbInformation, gstrSysName: Exit Sub
    
    lngSignID = rsTemp!ID
    Select Case Split(rsTemp!对象属性, ";")(1)
        Case 1
            strSource = GetSignSourceFromRTF1(lFileId, lngSignID, rsTemp!对象标记, blnMoved)
        Case 2
            strSource = GetSignSourceFromRTF2(lFileId, lngSignID, rsTemp!对象标记, blnMoved)
        Case 3 '表示源文组成方式由数据库跟据版次生成文本串
            Dim frmSVerify As New frmEPRSignVerify '使用新的源文组成方式
            Call frmSVerify.ShowMe(frmParent, lFileId)
            Unload frmSVerify: Set frmSVerify = Nothing
    End Select
    
    If strSource = "" Then Exit Sub
    If gobjESign Is Nothing Then
        Set gobjESign = CreateObject("zl9ESign.clsESign")
        Call gobjESign.Initialize(gcnOracle, glngSys)
    End If
    Call gobjESign.VerifySignature(strSource, lngSignID, 2)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Public Function GetSignSourceFromRTF1(ByVal lFileId As Long, ByVal lSignID As Double, ByVal lSignKey As Long, ByVal blnMoved As Boolean) As String
'功能:根据文件ID,提取签名源内容,用于在不打开编辑器情况下进行签名验证
'步骤:1 去掉最后一次签名相邻的签名图片 (如果有) 图片的组成字为 " PS(0000000X,0,1) PE(0000000X,0,1) "
'     2 去掉所有S关键字的签名对象
'     3 将所有图片,表格关键字中间的"口"字换成空格(因为签名时是空格表示)
'     4 将签名要素还原，签名时是"{经治医师签名}""主治医师签名""主任医师签名"
    Dim lSS As Long, lSE As Long, lES As Long, lEE As Long, bNeeded As Boolean, bFinded As Boolean, lKey As Long, lPos As Long
    Dim strZipFile As String, strRtfFile As String, lSigns As Long, lSEKey As Long
    Dim rsTemp As New ADODB.Recordset, strSource As String
    
    gstrSQL = "Select a.Id" & vbNewLine & _
                "From 电子病历内容 A, 电子病历内容 B" & vbNewLine & _
                "Where a.文件id = [1] And a.文件id = b.文件id And b.Id = [2] And a.开始版 > b.开始版"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取签名后变更记录", lFileId, lSignID)
    If Not rsTemp.EOF Then MsgBox "当前病历签名后发生修改，如需验证上次签名请回退修改！", vbInformation, gstrSysName
    
    strZipFile = zlBlobRead(5, lFileId, , blnMoved)
    If gobjFSO.FileExists(strZipFile) Then
        strRtfFile = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strRtfFile) Then
                gfrmPublic.edtBuff.OpenDoc strRtfFile
                gobjFSO.DeleteFile strRtfFile, True
        End If
        gobjFSO.DeleteFile strZipFile, True
    End If
    
    gfrmPublic.edtBuff.Freeze
    gfrmPublic.edtBuff.ForceEdit = True
    '查找最后一次签名相邻的签名图片,签名图片紧贴签名关键字
    bFinded = FindKey(gfrmPublic.edtBuff, "S", lSignKey, lSS, lSE, lES, lEE, bNeeded)
    If bFinded Then
        If gfrmPublic.edtBuff.Range(lEE, lEE + 4).Text = " PS(" Then
            gfrmPublic.edtBuff.Range(lEE, lEE + 35).Text = ""
        End If
    End If
    
    '去掉所有S关键字的签名对象
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "S", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSS, lEE).Text = "": lSigns = lSigns + 1
    Loop Until bFinded = False
    
    '将所有表格关键字中间的"口"字换成空格(因为签名时是空格表示,多次签名时，首次为空格，其它次为问号)
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "T", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSE, lES).Text = IIf(lSigns = 1, " ", "?"): lPos = lEE + 1
    Loop Until bFinded = False
    
    '将所有图片关键字中间的"口"字换成空格(因为签名时是空格表示)
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "P", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSE, lES).Text = IIf(lSigns = 1, " ", "?"): lPos = lEE + 1
    Loop Until bFinded = False
    
    '处理当次签名要素,因为签名当次时使用的源文中签名要素是"{经治医师签名}"形式，但签名后被更改为具体的姓名
    gstrSQL = "Select 对象属性 From 电子病历内容 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取签名要素", lSignID)
    If Not rsTemp.EOF Then
    If UBound(Split(rsTemp!对象属性, ";")) > 5 Then '历史版本可能没有第6个
    lSEKey = Val(Split(rsTemp!对象属性, ";")(6))
    If lSEKey <> 0 Then '签名可能没有使用签名要素
        gstrSQL = "Select 要素名称" & vbNewLine & _
                "From 电子病历内容" & vbNewLine & _
                "Where 文件id = [1] And 对象类型 = 4 And 对象标记 = [2] And 要素名称 In ('经治医师签名', '主治医师签名', '主任医师签名')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取签名要素名称", lFileId, lSEKey)
        If Not rsTemp.EOF Then '有则使用签名要素名称还原
            lPos = 0
            bFinded = FindKey(gfrmPublic.edtBuff, "E", lSEKey, lSS, lSE, lES, lEE, bNeeded)
            If bFinded Then
                If gfrmPublic.edtBuff.Range(lEE, lEE + 4).Text = " PS(" Then '使用了签名图片
                    gfrmPublic.edtBuff.Range(lEE, lEE + 35).Text = ""
                End If
                gfrmPublic.edtBuff.Range(lSE, lES).Text = "{" & rsTemp!要素名称 & "}"
            End If
        End If
    End If
    End If
    End If
    
    gfrmPublic.edtBuff.ForceEdit = False
    gfrmPublic.edtBuff.UnFreeze
    strSource = gfrmPublic.edtBuff.Text
    GetSignSourceFromRTF1 = strSource
End Function
Public Function GetSignSourceFromRTF2(ByVal lFileId As Long, ByVal lSignID As Double, ByVal lSignKey As Long, ByVal blnMoved As Boolean) As String
'功能:根据文件ID,提取签名源内容,用于在不打开编辑器情况下进行签名验证
'步骤:1 去掉最后一次签名相邻的签名图片 (如果有) 图片的组成字为 " PS(0000000X,0,1) PE(0000000X,0,1) "
'     2 去掉所有S关键字的签名对象
'     3 将所有图片,表格关键字中间的"口"字换成空格(因为签名时是空格表示)
'     4 将签名要素还原，签名时是"{经治医师签名}""主治医师签名""主任医师签名"
    Dim lSS As Long, lSE As Long, lES As Long, lEE As Long, bNeeded As Boolean, bFinded As Boolean, lKey As Long, lPos As Long
    Dim strZipFile As String, strRtfFile As String, lSEKey As Long
    Dim rsTemp As New ADODB.Recordset, strSource As String
        
    gstrSQL = "Select a.Id" & vbNewLine & _
                "From 电子病历内容 A, 电子病历内容 B" & vbNewLine & _
                "Where a.文件id = [1] And a.文件id = b.文件id And b.Id = [2] And a.开始版 > b.开始版"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取签名后变更记录", lFileId, lSignID)
    If Not rsTemp.EOF Then MsgBox "当前病历签名后发生修改，如需验证上次签名请回退修改！", vbInformation, gstrSysName
    
    strZipFile = zlBlobRead(5, lFileId, , blnMoved)
    If gobjFSO.FileExists(strZipFile) Then
        strRtfFile = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strRtfFile) Then
                gfrmPublic.edtBuff.OpenDoc strRtfFile
                gobjFSO.DeleteFile strRtfFile, True
        End If
        gobjFSO.DeleteFile strZipFile, True
    End If
    
    gfrmPublic.edtBuff.Freeze
    gfrmPublic.edtBuff.ForceEdit = True
    '查找最后一次签名相邻的签名图片,签名图片紧贴签名关键字
    bFinded = FindKey(gfrmPublic.edtBuff, "S", lSignKey, lSS, lSE, lES, lEE, bNeeded)
    If bFinded Then
        If gfrmPublic.edtBuff.Range(lEE, lEE + 4).Text = " PS(" Then
            gfrmPublic.edtBuff.Range(lEE, lEE + 35).Text = ""
        End If
    End If
    
    '去掉所有S关键字的签名对象
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "S", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSS, lEE).Text = ""
    Loop Until bFinded = False
    
    '将所有表格关键字中间的"口"字换成空格(因为签名时是空格表示,多次签名时，首次为空格，其它次为问号)
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "T", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSE, lES).Text = " ": lPos = lEE + 1
    Loop Until bFinded = False
    
    '将所有图片关键字中间的"口"字换成空格(因为签名时是空格表示)
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "P", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSE, lES).Text = " ": lPos = lEE + 1
    Loop Until bFinded = False
    
    '处理当次签名要素,因为签名当次时使用的源文中签名要素是"{经治医师签名}"形式，但签名后被更改为具体的姓名
    gstrSQL = "Select 对象属性 From 电子病历内容 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取签名要素", lSignID)
    If Not rsTemp.EOF Then
    If UBound(Split(rsTemp!对象属性, ";")) > 5 Then '历史版本可能没有第6个
    lSEKey = Val(Split(rsTemp!对象属性, ";")(6))
    If lSEKey <> 0 Then '签名可能没有使用签名要素
        gstrSQL = "Select 要素名称" & vbNewLine & _
                "From 电子病历内容" & vbNewLine & _
                "Where 文件id = [1] And 对象类型 = 4 And 对象标记 = [2] And 要素名称 In ('经治医师签名', '主治医师签名', '主任医师签名')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取签名要素名称", lFileId, lSEKey)
        If Not rsTemp.EOF Then '有则使用签名要素名称还原
            lPos = 0
            bFinded = FindKey(gfrmPublic.edtBuff, "E", lSEKey, lSS, lSE, lES, lEE, bNeeded)
            If bFinded Then
                If gfrmPublic.edtBuff.Range(lEE, lEE + 4).Text = " PS(" Then '使用了签名图片
                    gfrmPublic.edtBuff.Range(lEE, lEE + 35).Text = ""
                End If
                gfrmPublic.edtBuff.Range(lSE, lES).Text = "{" & rsTemp!要素名称 & "}"
            End If
        End If
    End If
    End If
    End If
    
    gfrmPublic.edtBuff.ForceEdit = False
    gfrmPublic.edtBuff.UnFreeze
    strSource = gfrmPublic.edtBuff.Text
    strSource = Replace(strSource, Chr(32), "")
    strSource = Replace(strSource, vbCr, "")
    strSource = Replace(strSource, vbLf, "")
    GetSignSourceFromRTF2 = strSource
End Function
Public Function GetSignSourceFromDB(ByVal lFileId As Long, ByVal lSignKey As String)
'功能：使用保存数据库后的内容文本（不含提纲 签名要素，签名对象,图片、表格及子对象）为数字签名原文
'说明：ID <> 对象序号 表示非表格的子对象，表格的子对象终止版=当前签名版次，以此为区别,而其它对象终止版=当前签名版次-1
    Dim rsTemp As New ADODB.Recordset, strSource As String
    gstrSQL = "Select ID, 父id, 开始版, 终止版, 对象类型,对象属性, 内容文本, 对象序号, 内容行次, 要素名称" & vbNewLine & _
                "From (Select ID, 父id, 开始版, 终止版, 对象类型,对象属性, 内容文本, 对象序号, 内容行次, 要素名称" & vbNewLine & _
                "       From 电子病历内容 A, (Select 开始版 版次 From 电子病历内容 Where 文件id = [1] And 对象类型 = 8 And 对象标记 = [2]) B" & vbNewLine & _
                "       Where 文件id = [1] And Instr(',经治医师签名,主治医师签名,主任医师签名,',',' || 要素名称 || ',') = 0 And" & vbNewLine & _
                "             (开始版 <= b.版次 And 终止版 = 0 Or 开始版 <= b.版次 And 终止版 = b.版次 And ID <> 对象序号 Or" & vbNewLine & _
                "             开始版 <= b.版次 And 终止版 = b.版次 + 1 And ID = 对象序号))" & vbNewLine & _
                "Start With 父id Is Null" & vbNewLine & _
                "Connect By Prior ID = 父id" & vbNewLine & _
                "Order Siblings By Decode(ID, 对象序号, 1, 对象序号), 内容行次"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取原文", lFileId, lSignKey)
    Do Until rsTemp.EOF
        Select Case rsTemp!对象类型
            Case 1 '不含提纲，因为提纲不显示，提纲的文字做为文本会另存，但SQL需要提纲做树形查询
            Case 5
                strSource = strSource & rsTemp!ID
            Case 8
                strSource = strSource & Split(rsTemp!内容文本, ";")(0) & Split(rsTemp!对象属性, ";")(4)
            Case Else
                strSource = strSource & rsTemp!内容文本
        End Select
        rsTemp.MoveNext
    Loop
    strSource = Replace(strSource, Chr(32), "")
    strSource = Replace(strSource, vbCr, "") '去掉文中的回车换行符，因为第一版时单独的换行只有回车符，再次修改保存或签名后后，被修改成回车换符
    strSource = Replace(strSource, vbLf, "")
    GetSignSourceFromDB = strSource
End Function

Public Function getPassESign(ByVal lngKind As Long, ByVal lngDeptId As Long) As Long
Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '0-门诊医嘱和病历；1-住院医生医嘱和病历；2-住院护士医嘱；3-医技医嘱和报告；4-护理记录和护理病历；5-药品发药；6-LIS;7-PACS
    
    gstrSQL = "Select Zl_Fun_Getsignpar([1],[2]) as 启用 From Dual "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取电子签名控制部门", lngKind, lngDeptId)
    If rsTemp.EOF Then
        getPassESign = 1
    Else
        getPassESign = rsTemp!启用
    End If

    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function LongIDsTable(ByVal strIDs As String, ByRef idPar() As String, Optional ByVal idParStart As Long = 1, Optional ByVal Alias As String = "B") As String
Dim strSQL As String, lngS As String, N As Integer, strReturn As String, strThis As String
    
    ReDim idPar(10) As String
    strSQL = "Select Column_Value ID From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))"
    N = 0
    Do While True
        If Len(strIDs) <= 4000 Then
            strThis = strIDs
            strIDs = ""
        Else
            strThis = Mid(strIDs, 1, InStrRev(Mid(strIDs, 1, 4000), ",") - 1)
            strIDs = Mid(strIDs, InStrRev(Mid(strIDs, 1, 4000), ",") + 1)
        End If
        
        If N > 9 Then
            strReturn = strReturn & vbNewLine & " Union " & Replace(strSQL, "[1]", "'" & strThis & "'")
        Else
            idPar(N) = strThis
            strReturn = IIf(strReturn = "", "", strReturn & vbNewLine & " Union ") & Replace(strSQL, "[1]", "[" & (N + idParStart) & "]")
        End If
        
        N = N + 1
        If strIDs = "" Then Exit Do
    Loop
    
    LongIDsTable = " (" & strReturn & ") " & Alias & " "
End Function
Public Function GetEPRContentNextId() As Double
'功能：因使用第三方PACS引起电子病历内容序列严重浪费超过LONG型最大值，临时处理为  单独提取电子病内容序列ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    '不能用错误错处理,原因是序列失效和没有序列时,应该返回错误,不然返回零,就有问题!

    
    strSQL = "Select 电子病历内容_ID.Nextval From Dual"
    
    Call zl9ComLib.SQLTest(App.ProductName, "mRichEPR", strSQL)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "单独提取电子病内容序列ID")
    Call zl9ComLib.SQLTest
    GetEPRContentNextId = rsTmp.Fields(0).Value
End Function
