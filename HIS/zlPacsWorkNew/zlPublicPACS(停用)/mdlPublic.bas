Attribute VB_Name = "mdlPublic"
Option Explicit

'文件种类  1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-诊断文书;6-知情文件;7-诊疗报告;8-诊疗申请
Public Enum EPRDocTypeEnum
    cpr门诊病历 = 1
    cpr住院病历 = 2
    cpr护理记录 = 3
    cpr护理病历 = 4
    cpr诊断文书 = 5
    cpr知情文件 = 6
    cpr诊疗报告 = 7             '诊疗单据：报告
    cpr诊疗申请 = 8             '诊疗单据：申请
End Enum

Public Const ELE_BACKCOLOR = &HD5FEFF               '要素的背景颜色 '&HDCDCDC
Public Const ELE_UNDERLINE = cprWave                '要素的下划线
Public Const PROTECT_FORECOLOR = &H662200           '自定义保护文本的前景色

Public gobjComLib As Object    'zl9ComLib.clsComLib
Public gcnOracle As ADODB.Connection
Public gstrSysName  As String
Public glngSys As Long

Public gstrSQL As String
Private mclsUnzip As Object

Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long

Public Type NETRESOURCE ' 网络资源
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Public Const RESOURCETYPE_ANY = &H0


Public Sub MkLocalDir(ByVal strDir As String)
'功能：创建本地目录
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Sub ClearCacheFolder(ByVal strCacheFolder As String)
'功能：当指定目录的大小达到一定百分比时，清空该目录
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String
    
    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        objCurFolder.Delete True
    End If
End Sub

Public Function funcConnectShardDir(strShareRemoteDir As String, strUserName As String, strPassWord As String) As Long
    '创建网络资源
    Dim NetR As NETRESOURCE
    Dim lngResult As Long
    
    NetR.dwType = RESOURCETYPE_ANY
    NetR.lpLocalName = vbNullString
    NetR.lpRemoteName = strShareRemoteDir
    NetR.lpProvider = vbNullString
    lngResult = WNetAddConnection2(NetR, strPassWord, strUserName, 0)
    
    If lngResult <> 0 Then
        MsgBox "网络连接失败，请检查网络设置是否正确！", vbInformation, gstrSysName
    End If
    funcConnectShardDir = lngResult
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer)
'-----------------------------------------------------------------------------
'功能：根据输入的图像数量，图像区域的宽度和高度，返回最佳的图像排列行数和列数
'参数： ImageCount－－图像数量
'       RegionWidth--图像显示区域的宽度
'       RegionHeight--图像显示区域的高度
'       Rows－－[返回]最佳行数
'       Cols－－[返回]最佳列数
'返回：返回最佳行数Rows，最佳列数Cols
'-----------------------------------------------------------------------------
    Dim iCols As Integer, iRows As Integer
    Dim iBase As Integer, blnDoLoop As Integer
    Dim lngFreeCount As Long
    
    If RegionHeight = 0 Then RegionHeight = 1
    If RegionWidth = 0 Then RegionWidth = 1
    
    On Error GoTo err
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))

    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    
    '当图像格式为如下等形式时，需要对行列进行修正
    
    '格式1：
    '图1  图2  图3  图4
    '图5  图6  图7  图8
    '空1  空2  空3  空4
    
    '格式2：
    '图1  图2  图3  图4
    '图5  图6  图7  图8
    '图9  空1  空2  空3
    
    lngFreeCount = iRows * iCols - ImageCount
    Do While lngFreeCount >= iCols Or lngFreeCount >= iRows
        If lngFreeCount >= iCols Then
            iRows = iRows - 1
        Else
            iCols = iCols - 1
        End If
        
        lngFreeCount = iRows * iCols - ImageCount
    Loop
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / iCols > RegionHeight > iRows Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    
    '再次修正行列数
    lngFreeCount = iRows * iCols - ImageCount
    Do While lngFreeCount >= iCols Or lngFreeCount >= iRows
        If lngFreeCount >= iCols Then
            iRows = iRows - 1
        Else
            iCols = iCols - 1
        End If
        
        lngFreeCount = iRows * iCols - ImageCount
    Loop
    
    Rows = iRows: Cols = iCols
err:
End Sub

Public Function FindNextKey(ByRef edtThis As Object, _
    ByVal lngCurPosition As Long, _
    ByVal strKeyType As String, _
    ByRef lngKey As Long, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTmp As String
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    
    sTmp = strKeyType & "S("
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTmp)
        If i <> 0 Then
            '看是否是关键字
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            '已找到起始关键字
            
            '查找结束关键字
            j = i + 16
LL2:
            sTmp = strKeyType & "E("
            j = InStr(j, sText, sTmp)
            If j <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '找到结束关键字
                strKeyType = strKeyType
                lngKSS = i - 1 '转换为0开始的坐标位置。
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                lngKey = Val(.Range(i + 2, i + 10))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindNextKey = True
            End If
        End If
    End With
End Function

Public Sub ReadRTF(edtThis As Editor, ByVal lngFileID As Long, ByVal blnClearMode As Boolean, ByVal blnMoved As Boolean)
'读取RTF文件
On Error GoTo errH
Dim ParaFmt As New cParaFormat, FontFmt As New cFontFormat, i As Long, lngLen As Long
Dim oEles As Object, oTabs As Object, oPics As Object
Dim strZipFile As String, strRtfFile As String, j As Long, rs As New ADODB.Recordset
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
Dim objFSO As New FileSystemObject
Dim strSQL As String

    Set oEles = CreateObject("zlRichEPR.cEPRElements")
    Set oTabs = CreateObject("zlRichEPR.cEPRTables")
    Set oPics = CreateObject("zlRichEPR.cEPRPictures")
    
    strZipFile = zlBlobRead(5, lngFileID, , blnMoved)
    If objFSO.FileExists(strZipFile) Then
        strRtfFile = zlFileUnzip(strZipFile)
        If objFSO.FileExists(strRtfFile) Then
                edtThis.OpenDoc strRtfFile
                objFSO.DeleteFile strRtfFile, True
        End If
        objFSO.DeleteFile strZipFile, True
    End If
    If Trim(edtThis.Text) = "" Then Exit Sub

    '读取图片,表格,要素
    strSQL = "Select Level,ID, 文件id,开始版, 终止版," & _
                "   父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 定义提纲id, 复用提纲," & vbNewLine & _
                "       使用时机, 诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
                "From (Select ID, 文件id,开始版, 终止版," & _
                "               父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id,定义提纲id," & vbNewLine & _
                "              复用提纲, 使用时机, 诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
                "       From 电子病历内容" & vbNewLine & _
                "       Where 文件id = [1] And 对象序号>0 and 对象序号<ID)" & vbNewLine & _
                "Start With 父id Is Null" & vbNewLine & _
                "Connect By Prior ID = 父id" & vbNewLine & _
                "Order By 对象序号, 内容行次"
    If blnMoved Then strSQL = Replace(strSQL, "电子病历内容", "H电子病历内容")
    Set rs = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "还原表图", lngFileID)
    Do Until rs.EOF
        Select Case rs!对象类型
            Case 3  '表格
                lKey = oTabs.Add(Nvl(rs!对象标记, 0))                  '恢复Key值！
                Call oTabs("K" & lKey).FillTableMember(rs, "电子病历内容")
            Case 4  '要素
                lKey = oEles.Add(Nvl(rs!对象标记, 0))
                Call oEles("K" & lKey).FillElementMember(rs, "电子病历内容")
            Case 5  '图片
                lKey = oPics.Add(Nvl(rs("对象标记"), 0))
                Call oPics("K" & lKey).FillPictureMember(rs, "电子病历内容")
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
    
    For j = 1 To oEles.Count '删除空的要素，展开型要素重新刷新以去掉下波浪线
        If oEles(j).内容文本 = "" Then
            oEles(j).DeleteFromEditor edtThis
        ElseIf oEles(j).输入形态 = 1 Then
            oEles(j).Refresh edtThis
        End If
    Next
    

    ' 对最终文档的处理
    edtThis.SelectAll
    If blnClearMode Then
        edtThis.AuditMode = True
        edtThis.AcceptAuditText    '清洁模式
    End If
    lngLen = Len(edtThis.Text)
    For i = 0 To lngLen - 1 '只将背景色为要素背景色颜色去掉
        If edtThis.Range(i, i + 1).Font.BackColor = ELE_BACKCOLOR Then
            edtThis.Range(i, i + 1).Font.BackColor = tomAutoColor
        End If
        If edtThis.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR Then
            edtThis.Range(i, i + 1).Font.ForeColor = tomAutoColor
        End If
    Next
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Public Function FindKey(ByRef edtThis As Object, _
        ByRef strKeyType As String, _
        ByRef lngKey As Long, _
        ByRef lngKSS As Long, _
        ByRef lngKSE As Long, _
        ByRef lngKES As Long, _
        ByRef lngKEE As Long, _
        ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTmp As String
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    
    sTmp = strKeyType & "S(" & Format(lngKey, "00000000")
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = 1
LL1:
        i = InStr(i, sText, sTmp)
        If i <> 0 Then
            '看是否是关键字
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            '已找到起始关键字
            
            '查找结束关键字
            j = i + 16
LL2:
            sTmp = strKeyType & "E(" & Format(lngKey, "00000000")
            j = InStr(j, sText, sTmp)
            If j <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '找到结束关键字
                strKeyType = strKeyType
                lngKSS = i - 1 '转换为0开始的坐标位置。
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindKey = True
            End If
        End If
    End With
End Function

Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String, Optional ByVal blnMoved As Boolean) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, StrText As String
    Dim rsLob As New ADODB.Recordset
    Dim strSQL As String
    
    err = 0: On Error GoTo errHand
    
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
    
    strSQL = "Select Zl_Lob_Read([1],[2],[3],[4]) as 片段 From Dual"
    lngCount = 0
    Do
        Set rsLob = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "zlBlobRead", Action, KeyWord, lngCount, IIf(blnMoved, 1, 0))
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
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
    zlBlobRead = ""
End Function

Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String
    Dim objFSO As New FileSystemObject
    
    On Error GoTo errHand
    
    If Not objFSO.FileExists(strZipFile) Then zlFileUnzip = "": Exit Function
    
    strZipPath = objFSO.GetSpecialFolder(2) '取临时目录
    strZipPathTmp = strZipPath & "\" & Format(Now, "yyMMdd") & CStr(100 * Timer)
    If Not objFSO.FolderExists(strZipPathTmp) Then Call objFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp & "\TMP.RTF"
    If objFSO.FileExists(strZipFileTmp) Then objFSO.DeleteFile strZipFileTmp
    
    If mclsUnzip Is Nothing Then Set mclsUnzip = CreateObject("zlRichEPR.cUnzip")

    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    
    If objFSO.FileExists(strZipFileTmp) Then
        
        strZipFileName = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer) & ".RTF"
        If objFSO.FileExists(strZipFileName) Then objFSO.DeleteFile strZipFileName
                
        Call objFSO.CopyFile(strZipFileTmp, strZipFileName)
        
        If objFSO.FileExists(strZipFileTmp) Then objFSO.DeleteFile strZipFileTmp, True
        
        On Error Resume Next
        If objFSO.FolderExists(strZipPathTmp) Then objFSO.DeleteFolder strZipPathTmp, True
        
        zlFileUnzip = strZipFileName
    Else
        zlFileUnzip = ""
    End If
    
    Exit Function
    
errHand:
    Call gobjComLib.SaveErrLog
End Function

Public Function GetFileRange(ByVal lFileId As Long, ByVal lngRecordId As Long, ByVal strCreateTime As String, _
                            ByVal eDocType As EPRDocTypeEnum, ByVal lngPatId As Long, ByVal lngPageId As Long, _
                            Optional ByVal blnMoved As Boolean) As String
    '******************************************************************************************************************
    '功能：读取当前病历(可能当前病历未保存)前所有共享病历
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTemp As New ADODB.Recordset, strTime As String, dStar As Date, dEnd As Date
    Dim strSQL As String, strIDs As String, blnNewPage As Boolean, n_Num As Integer, n_S As Integer, n_E As Integer

    On Error GoTo errHand
    strTime = Format(strCreateTime, "yyyy-MM-dd HH:mm:ss")
    If strTime = "" Then strTime = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If strTime = "00:00:00" Then strTime = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    blnNewPage = gobjComLib.zlDatabase.GetPara("转科后要求书写的共享病历另起一页打印", glngSys, 1251, 1) = 1 '=0表示转科后共享病历连续打印 =1 表示另起一页打印

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
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "提取序号", lngRecordId)
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
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "提取页面文件之前一次书写", lFileId, lngPatId, lngPageId, CDate(strTime), n_Num)
    If rsTemp.EOF Then
        dStar = CDate("2000-01-01 00:00:00"): n_S = -1 '表明之前没有写过页面文件
    Else
        dStar = CDate(rsTemp!创建时间)
        n_S = IIf(n_Num = 0, 0, Nvl(rsTemp!序号, -1))
    End If
    
    '提取页面文件当前时间之后最近一次书写记录
    gstrSQL = "Select 创建时间,序号" & vbNewLine & _
                "From (Select a.创建时间,a.序号" & vbNewLine & _
                "       From 电子病历记录 A, (" & strSQL & ") B" & vbNewLine & _
                "       Where a.文件id = b.Id And a.病人id = [2] And a.主页id = [3] And " & IIf(n_Num = 0, "a.创建时间 > [4]", "((a.创建时间 > [4] and Nvl(a.序号,0)=0) Or a.序号>[5])") & vbNewLine & _
                "       Order By a.创建时间)" & vbNewLine & _
                "Where Rownum = 1"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "提取页面文件之后一次书写", lFileId, lngPatId, lngPageId, CDate(strTime), n_Num)
    If rsTemp.EOF Then
        dEnd = CDate("3001-01-01") '表明之后没有写过页面文件
        n_E = 9999
    Else
        dEnd = CDate(rsTemp!创建时间) - 1 / 24 / 60 / 60 '表明之后写过，取近记录的时间减一秒,即不包含此记录
        n_E = IIf(n_Num = 0, 0, Nvl(rsTemp!序号, 9999) - 1)
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
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "提取页面文件之后一次书写", lFileId, lngPatId, lngPageId, dStar, dEnd, n_S, n_E)
    Do Until rsTemp.EOF
        strIDs = strIDs & "," & rsTemp!Id
        rsTemp.MoveNext
    Loop
    strIDs = Mid(strIDs, 2)
    
    GetFileRange = strIDs
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

