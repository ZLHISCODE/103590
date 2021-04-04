Attribute VB_Name = "mPartogram"

Option Explicit
Public mclsUnzip As New cUnzip
Public mclsZip As New cZip

Public glngHours As Long
Public gobjBodyEditor As Object
Public gfrmPublic As Object
Public gobjESign As Object                  '电子签名接口部件
Public gobjFSO As New FileSystemObject

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
Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const GWL_STYLE = (-16)              'Set the window style
Public Const WS_CAPTION = &HC00000
Public Const WS_THICKFRAME = &H40000        '厚边框
Public Const WS_SYSMENU = &H80000           '在标题栏是否具备系统菜单
Public Const WS_MINIMIZEBOX = &H20000       '具备最小化按钮
Public Const WS_MAXIMIZEBOX = &H10000       '具备最大化按钮
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

'窗体样式:
Public Const WS_CHILD = &H40000000          '子窗体
Public Const WS_HSCROLL = &H100000          '具备水平滚动条
Public Const WS_VSCROLL = &H200000          '具备垂直滚动条
Public Const WS_VISIBLE = &H10000000        '可视
Public Const WS_CLIPCHILDREN = &H2000000    '出去子窗体的父窗体绘图区域
Public Const WS_CLIPSIBLINGS = &H4000000    '更新子窗体时，排除重叠的其他子窗体
Public Const WS_BORDER = &H800000           '具备边框
Public Const WS_TABSTOP = &H10000           'Tab停留
Public Const WS_POPUP = &H80000000          '弹出窗体
Public Const WS_DLGFRAME = &H400000         '双边框但是无标题栏的窗体
Public Const WS_EX_TOPMOST = &H8&           '最前面
Public Const WS_EX_CLIENTEDGE = &H200&      '3D效果
Public Const WS_EX_Transparent = &H20&      '窗体透明
Public Const WS_DISABLED = &H8000000        '不可用

Public Const GWL_EXSTYLE = (-20)            'Set the extended window style
Public Const GWL_USERDATA = (-21)           'Sets the 32-bit value associated with the window.
Public Const GWL_WNDPROC = -4               'Sets a new address for the window procedure.
Public Const GWL_HWNDPARENT = (-8)          'Sets a new application instance handle.

Public Const HWND_TOPMOST = -1              '最前面
Public Const SW_SHOW = 5                    '激活窗体并显示
Public Const SW_HIDE = 0                    '隐藏
Public Const SW_SHOWNORMAL = 1              '还原
Public Const GW_CHILD = 5                   '获取主窗体句柄
Public Const GW_HWNDNEXT = 2                '获取指定窗体Z-Order下的下一窗体的句柄
Public Const CW_USEDEFAULT  As Long = &H80000000        '默认值
Public Const GDI_ERROR = &HFFFF             '出现GDI错误！


'#########################################################################
' 鼠标激活响应
Public Const MA_ACTIVATE = 1                '激活CWnd
Public Const MA_ACTIVATEANDEAT = 2          '激活CWnd，屏蔽鼠标事件
Public Const MA_NOACTIVATE = 3              '不激活CWnd
Public Const MA_NOACTIVATEANDEAT = 4        '不激活CWnd，屏蔽鼠标事件

Public Const H_MAX As Long = &HFFFF + 1     '最大值

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'获取指定窗体的边界矩形尺寸
Public Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
'获取指定窗体的属性
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
'改变窗体位置、Zorder、尺寸等
Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'改变指定窗体的属性
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
'改变指定窗体的父窗体
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
' 发送指定消息到窗体，等待处理完才返回；而 PostMessage() 函数发送消息，立即返回！HWND hWnd 目标窗体的句柄。wMsg待发送的消息。wParam消息第一参数。lParam消息第二参数。
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowFocus Lib "user32" Alias "SetFocus" (ByVal Hwnd As Long) As Long

''去掉TextBox的默认右键菜单
'Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
'    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
'    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hwnd, msg, wp, lp)
'End Function

Public Function ReDimArray(ByRef strArray() As String) As Long
    '----------------------------------------------------------------------
    '功能：重新定义数组
    '----------------------------------------------------------------------
    Dim lngCount As Long
    Dim strTmp As String
    
    On Error GoTo InitHand
    
    strTmp = strArray(1)
    
    lngCount = UBound(strArray) + 1
    
    GoTo OkHand
    
InitHand:
    
    lngCount = 1
    
OkHand:
    
    ReDim Preserve strArray(1 To lngCount)
            
    ReDimArray = lngCount
End Function

Public Function ZVal(ByVal varValue As Variant) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function

Public Function ArchiveChart(ByVal lngFileID As Long) As Boolean
'功能：检查文件是否归档
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    gstrSQL = "select 1 From 病人护理文件 where ID=[1] And 归档人 IS NOT NULL"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取体温单文件是否归档", lngFileID)
    ArchiveChart = (rsTemp.RecordCount <> 0)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ToStandDate(ByVal strDate As String) As String
    Dim arrData
    Dim strMonth As String, strDay As String
    
    arrData = Split(strDate, "/")
    strMonth = arrData(1)
    strDay = arrData(0)
    If Len(strMonth) = 1 Then strMonth = "0" & strMonth
    If Len(strDay) = 1 Then strDay = "0" & strDay
    ToStandDate = strMonth & "-" & strDay
End Function

Public Sub GetUserInfo()
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHand
    Set rsTemp = zlDatabase.GetUserInfo
    With rsTemp
        If rsTemp.RecordCount <> 0 Then
            gstrDBUser = .Fields("用户名").Value
            glngUserId = .Fields("ID").Value                '当前用户id
            gstrUserCode = .Fields("编号").Value            '当前用户编码
            gstrUserName = .Fields("姓名").Value            '当前用户姓名
            gstrUserAbbr = NVL(.Fields("简码").Value, "")  '当前用户简码
            glngDeptId = .Fields("部门id").Value            '当前用户部门id
            gstrDeptCode = .Fields("部门码").Value        '当前用户
            gstrDeptName = .Fields("部门名").Value        '当前用户
        Else
            gstrDBUser = ""
            glngUserId = 0
            gstrUserCode = ""
            gstrUserName = ""
            gstrUserAbbr = ""
            glngDeptId = 0
            gstrDeptCode = ""
            gstrDeptName = ""
        End If
    End With
    
    gstrSQL = "Select 签名 From 人员表 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取签名名字", glngUserId)
    If Not rsTemp.EOF Then
        gstrSignName = NVL(rsTemp!签名, gstrUserName)
    End If
   
   
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err = 0
End Sub

Public Function GetDbOwner(ByVal lngSys As Long) As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String

    GetDbOwner = ""
    Err = 0: On Error GoTo errHand
    strSQL = "Select 所有者 From Zlsystems Where 编号 = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetDbOwner", lngSys)
    If rsTemp.RecordCount <> 0 Then GetDbOwner = "" & rsTemp!所有者
    rsTemp.Close
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "SQL", adVarChar, 300
        .Fields.Append "Trans", adTinyInt                   '1表示开始;2表示结束
        .Fields.Append "Custom", adTinyInt
        .Fields.Append "Parameter", adVarChar, 500
        
        .Open
    End With
    
    SQLRecord = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String, Optional ByVal intTrans As Integer = 0, Optional ByVal intCustom As Integer = 0, Optional ByVal strParameter As String = "") As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.AddNew
    rs("SQL").Value = strSQL
    rs("Trans").Value = intTrans
    rs("Custom").Value = intCustom
    rs("Parameter").Value = strParameter
    SQLRecordAdd = True
    
    Exit Function
    
errHand:
End Function

Public Function SQLRecordExecute(ByVal rs As ADODB.Recordset, Optional ByVal strTitle As String, Optional ByVal blnHaveTrans As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim strSQL As String
    
    On Error GoTo errHand
    
    If rs.RecordCount > 0 Then
        If Len(strTitle) = 0 Then strTitle = gstrSysName
        blnTran = True
        
        If blnHaveTrans Then gcnOracle.BeginTrans
        
        rs.MoveFirst
    
        For intLoop = 1 To rs.RecordCount
            
            If Val(rs("Custom").Value) = 0 Then
                strSQL = CStr(rs("SQL").Value)
                Call zlDatabase.ExecuteProcedure(strSQL, strTitle)
            End If
            
            rs.MoveNext
        Next
    
        If blnHaveTrans Then gcnOracle.CommitTrans
        blnTran = False
    End If
    
    SQLRecordExecute = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran And blnHaveTrans Then gcnOracle.RollbackTrans
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
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    strZipPath = gobjFSO.GetSpecialFolder(2)
    strZipPathTmp = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer)
    Call gobjFSO.CreateFolder(strZipPathTmp)
    
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

Public Function GetTmpPath() As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strFileTemp As String
    Dim lngTemp As Long
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    
    GetTmpPath = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
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

Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'功能：显示或隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.Hwnd, vRect)
    lngStyle = GetWindowLong(objForm.Hwnd, GWL_STYLE)
    If blnCaption Then
        lngStyle = lngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then lngStyle = lngStyle Or WS_SYSMENU
        If objForm.MaxButton Then lngStyle = lngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then lngStyle = lngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong objForm.Hwnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.Hwnd, 0, 0, 0, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Public Function IsAllowInput(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strTime As String, ByVal strCurTime As String) As Boolean
    '取出指定病人在指定时间之后关键点的时间
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    IsAllowInput = True
    gstrSQL = "" & _
              " SELECT DECODE(终止原因,1,'出院',3,'转科',10,'预出院',15,'转病区',DECODE(开始原因,10,'出院','未定义')) AS 类型,终止时间 AS 时间" & _
              " From 病人变动记录" & _
              " WHERE (终止原因 IN (1,3,10,15) OR 开始原因=10) And 病人ID=[1] And 主页ID=[2] And [3] <= 终止时间" & _
              " ORDER BY 终止时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取出指定病人在指定时间之后关键点的时间", lng病人ID, lng主页ID, CDate(strTime))
    If rsTemp.RecordCount = 0 Then Exit Function
    
    '只取第一条符合的记录
    strTime = Format(DateAdd("H", glngHours, rsTemp!时间), "yyyy-MM-dd HH:mm")
    strCurTime = Format(strCurTime, "yyyy-MM-dd HH:mm")
    
    If strTime < strCurTime Then IsAllowInput = False
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SQLDIY(strSQL As String)
    If gblnMoved Then
        strSQL = Replace(strSQL, "病人护理文件", "H病人护理文件")
        strSQL = Replace(strSQL, "病人护理数据", "H病人护理数据")
        strSQL = Replace(strSQL, "病人护理明细", "H病人护理明细")
        strSQL = Replace(strSQL, "病人护理打印", "H病人护理打印")
    End If
End Sub

