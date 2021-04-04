Attribute VB_Name = "mRichEPR"
Option Explicit
Public Const ELE_BACKCOLOR = &HFFEBD7               '要素的背景颜色 '&HDCDCDC
Public Const ELE_UNDERLINE = cprWave                '要素的下划线
Public Const PROTECT_BGCOLOR = &HE0E0E0             '自定义保护文本的背景色
Public Const PROTECT_FORECOLOR = &H662200           '自定义保护文本的前景色
Public gobjRegister As Object                       '密码验证组件
Public gobjESign As Object                          '电子签名接口部件
Public Type PreDefinedKeyInfo   '保留关键字
    KeyStart As String
    KeyEnd As String
End Type
Public gKeyWords(1 To 6) As PreDefinedKeyInfo       '预定义关键字
'##########################################################################################
'## 压缩与解压
'##########################################################################################
Private mclsZip As New cTabZip
Private mclsUnzip As New cTabUnzip

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
    Dim aryChunk() As Byte, strText As String
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
    
    gstrSQL = "Select Zl_Lob_Read(" & Action & ",'" & KeyWord & "'," & "[1]" & IIf(blnMoved, ",1", "") & ") as 片段 From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSQL, "zlBlobRead", lngCount)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.fields(0).Value) Then Exit Do
        strText = rsLob.fields(0).Value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
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
    Kill strFile: zlBlobRead = ""
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
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
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
        strText = Join(aryHex, "")
        gstrSQL = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "','" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
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
Public Function zlBlobSql(ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String, arrSQL As Variant) As Boolean
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
    Err = 0: On Error GoTo errHand
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    Err = 0: On Error GoTo errHand
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize '2000字节余数
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0) '余数=0表示整除
    
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
        strText = Join(aryHex, "")
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "','" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
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
Public Function zlFileUnzip(ByVal strZipFile As String, Optional ByVal strExtenName As String = "XML") As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String
    
    On Error GoTo errHand
    
    If Not gobjFSO.FileExists(strZipFile) Then zlFileUnzip = "": Exit Function '原文件不存在直接退出
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))       '提取原文件路径
    
    strZipPathTmp = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer)    '原文件目录下生成临时目录
    Call gobjFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp & "\TMP." & strExtenName                          '指定临时目录下的解压文件全路径
    If gobjFSO.FileExists(strZipFileTmp) Then gobjFSO.DeleteFile strZipFileTmp      '如果全路径文件存先删除
    
    With mclsUnzip                                                                  '解压
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    If gobjFSO.FileExists(strZipFileTmp) Then                                       '解压后临时文件存在
        
        strZipFileName = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer) & "." & strExtenName  '生成原文件目录下临时文件
        If gobjFSO.FileExists(strZipFileName) Then gobjFSO.DeleteFile strZipFileName
                
        Call gobjFSO.CopyFile(strZipFileTmp, strZipFileName)                        '将解压文件COPY到原文件目录下
        
        If gobjFSO.FileExists(strZipFileTmp) Then gobjFSO.DeleteFile strZipFileTmp, True    '删除解压文件
        
        zlFileUnzip = strZipFileName
    Else
        zlFileUnzip = ""
    End If
    On Error Resume Next
    If gobjFSO.FolderExists(strZipPathTmp) Then gobjFSO.DeleteFolder strZipPathTmp, True '删除解压文件目录
    Err.Clear
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
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
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
    ByVal iPatientType As PatiFrom, _
    ByVal lng医嘱id As Long, Optional lngBabyNum As Long) As String
    
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = "Select Zl_Replace_Element_Value([1],[2],[3],[4],[5],[6]) From Dual"
    Err = 0: On Error GoTo DBError
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "读取替换项", ElementName, CLng(sPatientID), _
        CLng(sPageID), CLng(iPatientType), lng医嘱id, lngBabyNum)
    If rsTmp.EOF Or rsTmp.BOF Then
        GetReplaceEleValue = ""
    Else
        GetReplaceEleValue = Trim(IIf(IsNull(rsTmp.fields(0).Value), "", rsTmp.fields(0).Value))
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
    Optional lngPatiID As Long, Optional lngPatiPageID As Long) As EPRSignLevel
    Dim rs As New ADODB.Recordset, lngR As Long, lngLevel1 As Long, lngLevel2 As Long
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select g.功能" & vbNewLine & _
            "From zlRoleGrant g, Sys.Dba_Role_Privs r, 上机人员表 p" & vbNewLine & _
            "Where r.Grantee = Upper(p.用户名) And g.角色 = r.Granted_Role And g.系统 = [2] And g.序号 = [3] And g.功能 = [4] And" & vbNewLine & _
            "      p.人员id = [1]" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select [4] As 功能 From 上机人员表 p Where 用户名 = '" & UCase(gstrDbOwner) & "' And p.人员id = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngUserID, glngSys, 1070, "签名权")
    If rs.RecordCount <= 0 Then GetUserSignLevel = TabSL_空白: Exit Function
    
    gstrSQL = "select 聘任技术职务 from 人员表 p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngUserID)
    If Not rs.EOF Then
        lngR = Nvl(rs("聘任技术职务"), 0)
    End If
    Select Case lngR    '1 正高  2 副高  3 中级  4 助理/师级  5 员/士  9 待聘
    Case 1: lngLevel1 = TabSL_正高
    Case 2: lngLevel1 = TabSL_主任
    Case 3: lngLevel1 = TabSL_主治
    Case Else: lngLevel1 = TabSL_经治
    End Select
    rs.Close
    
    If lngPatiID > 0 Then
        gstrSQL = "Select 经治医师, 主治医师, 主任医师 " & _
            " From 病人变动记录 " & _
            " Where 病人ID = [1] And 主页ID = [2] And (终止时间 Is Null Or 终止原因 = 1) " & _
            "       And 开始时间 Is Not Null And Nvl(附加床位, 0) = 0"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", lngPatiID, lngPatiPageID)
        If rs.EOF Then
            lngLevel2 = TabSL_经治
        Else
            If rs.fields("主任医师") = IIf(strUserName = "", UserInfo.姓名, strUserName) Then
                lngLevel2 = TabSL_主任
            ElseIf rs.fields("主治医师") = IIf(strUserName = "", UserInfo.姓名, strUserName) Then
                lngLevel2 = TabSL_主治
            Else
                lngLevel2 = TabSL_经治
            End If
        End If
    End If
    GetUserSignLevel = IIf(lngLevel1 >= lngLevel2, lngLevel1, lngLevel2)
    Exit Function

errHand:
    GetUserSignLevel = TabSL_空白
End Function
Public Function GetCharColor(ByVal lng开始版 As Long, ByVal lng终止版 As Long) As OLE_COLOR
    '根据开始版、终止版获取最终字符颜色
    Dim R As Long, G As Long, b As Long
    R = 255
    G = GetColorVectorG(lng开始版)
    b = GetColorVectorB(lng终止版)
    If G = 0 And b = 0 Then
        GetCharColor = vbBlack
    Else
        GetCharColor = RGB(R, G, b)
    End If
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
Public Function GetEPRContentNextId() As Double
'功能：因使用第三方PACS引起电子病历内容序列严重浪费超过LONG型最大值，临时处理为  单独提取电子病内容序列ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    '不能用错误错处理,原因是序列失效和没有序列时,应该返回错误,不然返回零,就有问题!

    
    strSQL = "Select 电子病历内容_ID.Nextval From Dual"
    
    Call zl9ComLib.SQLTest(App.ProductName, "mRichEPR", strSQL)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "单独提取电子病内容序列ID")
    Call zl9ComLib.SQLTest
    GetEPRContentNextId = rsTmp.fields(0).Value
End Function