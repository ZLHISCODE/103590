Attribute VB_Name = "mTend"

Option Explicit
Public mclsUnzip As New cUnzip
Public mclsZip As New cZip

Public glngHours As Long
Public gobjBodyEditor As Object
Public gobjPartogram As Object
Public gfrmPublic As Object
Public gobjFSO As New FileSystemObject
Public gobjESign As Object  '电子签名接口部件

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
Public gstrMecState As String                '当前病人病案状态(结合EprIsCommit函数使用)

'矩形
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

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
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'获取指定窗体的边界矩形尺寸
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'获取指定窗体的属性
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'改变窗体位置、Zorder、尺寸等
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'改变指定窗体的属性
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
'改变指定窗体的父窗体
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
' 发送指定消息到窗体，等待处理完才返回；而 PostMessage() 函数发送消息，立即返回！HWND hWnd 目标窗体的句柄。wMsg待发送的消息。wParam消息第一参数。lParam消息第二参数。
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

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

Public Function CreateBodyEditor() As Boolean
    Dim strDLL As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    If gobjBodyEditor Is Nothing Then
        gstrSQL = " Select 新部件 From 体温部件 Where Nvl(启用,0)=1"
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "提取体温部件")
        If Err <> 0 Then
            strDLL = "zl9TemperatureChart"
        Else
            If rsTemp.RecordCount = 0 Then
                strDLL = "zl9TemperatureChart"
            Else
                strDLL = NVL(rsTemp!新部件, "zl9TemperatureChart")
            End If
        End If
        
        Err = 0
        strDLL = strDLL & ".clsBodyEditor"
        Set gobjBodyEditor = CreateObject(strDLL)
        If Err <> 0 Then
            MsgBox "    创建体温部件失败！" & vbCrLf & "    程序将创建标准的体温部件进行数据展现，请检查指定的体温部件是否存在或已损坏！" & vbCrLf & "    详细错误：" & Err.Description, vbInformation, gstrSysName
            
            '如果创建指定的体温部件出错则创建标准的体温部件，因为这里不处理的话，后面可能存在直接使用体温部件中的对象，从而导致程序崩溃
            strDLL = "zl9TemperatureChart.clsBodyEditor"
            Set gobjBodyEditor = CreateObject(strDLL)
        End If
        
        Call gobjBodyEditor.InitBodyEditor(glngSys, gcnOracle)
    End If
    
    CreateBodyEditor = True
    Exit Function
End Function

Public Function CreatePartogram() As Boolean
    Dim strDLL As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    If gobjPartogram Is Nothing Then
        gstrSQL = " Select 部件 From 产程部件 Where Nvl(启用,0)=1"
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "提取产程部件")
        If Err <> 0 Then
            strDLL = "zl9Partogram"
        Else
            If rsTemp.RecordCount = 0 Then
                strDLL = "zl9Partogram"
            Else
                strDLL = NVL(rsTemp!部件, "zl9Partogram")
            End If
        End If
        
        Err = 0
        strDLL = strDLL & ".clsPartogram"
        Set gobjPartogram = CreateObject(strDLL)
        If Err <> 0 Then
            MsgBox "    创建产程部件失败！" & vbCrLf & "    程序将创建标准的产程部件进行数据展现，请检查指定的产程部件是否存在或已损坏！" & vbCrLf & "    详细错误：" & Err.Description, vbInformation, gstrSysName
            
            '如果创建指定的产程部件出错则创建标准的产程部件，因为这里不处理的话，后面可能存在直接使用产程部件中的对象，从而导致程序崩溃
            strDLL = "zl9Partogram.clsPartogram"
            Err = 0
            Set gobjPartogram = CreateObject(strDLL)
        End If
        If Err <> 0 Then Err.Clear: Set gobjPartogram = Nothing: Exit Function
        
        Call gobjPartogram.InitPartogram(gcnOracle, glngSys)
    End If
    
    CreatePartogram = True
    Exit Function
End Function

Public Function ArchiveChart(ByVal lngFileID As Long) As Boolean
'功能：检查文件是否归档
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    gstrSQL = "select 1 From 病人护理文件 where ID=[1] And 归档人 IS NOT NULL"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取体温单文件是否归档", lngFileID)
    ArchiveChart = (rsTemp.RecordCount <> 0)
    Exit Function
ErrHand:
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

    On Error GoTo ErrHand
        
    Set rsTemp = zlDatabase.GetUserInfo
    With rsTemp
        If .RecordCount <> 0 Then
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
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Err = 0
End Sub

Public Function GetDbOwner(ByVal lngSys As Long) As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String

    GetDbOwner = ""
    Err = 0: On Error GoTo ErrHand
    strSQL = "Select 所有者 From Zlsystems Where 编号 = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetDbOwner", lngSys)
    If rsTemp.RecordCount <> 0 Then GetDbOwner = "" & rsTemp!所有者
    rsTemp.Close
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
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
    
ErrHand:
    
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String, Optional ByVal intTrans As Integer = 0, Optional ByVal intCustom As Integer = 0, Optional ByVal strParameter As String = "") As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
    rs.AddNew
    rs("SQL").Value = strSQL
    rs("Trans").Value = intTrans
    rs("Custom").Value = intCustom
    rs("Parameter").Value = strParameter
    SQLRecordAdd = True
    
    Exit Function
    
ErrHand:
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
    
    On Error GoTo ErrHand
    
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
ErrHand:
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
    
    On Error GoTo ErrHand
    
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
    
ErrHand:
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

'################################################################################################################
'## 功能：  将数据从一个XtremeReportControl控件复制到VSFlexGrid，以便进行打印
'################################################################################################################
Public Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '将全部组强制展开,复制数据表格
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    
    Dim lngCol As Long, lngRow As Long
    
    On Error GoTo ErrHand:
    For Each rptRow In rptList.Rows
        If rptRow.GroupRow Then rptRow.Expanded = True
    Next
    
    With vfgList
        .Clear
        .Rows = rptList.Records.Count + 1
        .Cols = 0: .Cols = rptList.Columns.Count
        .FixedCols = rptList.GroupsOrder.Count
        
        '标题行复制
        .ROW = 0
        lngCol = 0
        For Each rptCol In rptList.GroupsOrder
            .TextMatrix(0, lngCol) = rptCol.Caption
            .ColData(lngCol) = rptCol.ItemIndex
            Select Case rptCol.Alignment
            Case xtpAlignmentLeft: .FixedAlignment(lngCol) = flexAlignLeftCenter
            Case xtpAlignmentCenter: .FixedAlignment(lngCol) = flexAlignCenterCenter
            Case xtpAlignmentRight:  .FixedAlignment(lngCol) = flexAlignRightCenter
            End Select
            .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .FixedAlignment(lngCol)
            .ColWidth(lngCol) = rptCol.Width * 15
            .MergeCol(lngCol) = True
            lngCol = lngCol + 1
        Next
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .TextMatrix(0, lngCol) = rptCol.Caption
                .ColData(lngCol) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .ColAlignment(lngCol) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .ColAlignment(lngCol) = flexAlignCenterCenter
                Case xtpAlignmentRight: .ColAlignment(lngCol) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
                .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .ColAlignment(lngCol)
                If rptCol.Width < 20 Then
                    .ColWidth(lngCol) = 0
                Else
                    .ColWidth(lngCol) = rptCol.Width * 15
                End If
                lngCol = lngCol + 1
            End If
        Next
        vfgList.Cols = lngCol
        
        '数据行复制
        lngRow = 0
        For Each rptRow In rptList.Rows
            If rptRow.GroupRow = False Then
                lngRow = lngRow + 1
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(lngRow, lngCol) = rptRow.Record(.ColData(lngCol)).Value
                Next
            End If
        Next
    End With
    zlReportToVSFlexGrid = True
    Exit Function

ErrHand:
    zlReportToVSFlexGrid = False
End Function

'################################################################################################################
'## 功能：  将数据从一个展示VSFlexGrid控件复制到打印VSFlexGrid，以便进行打印
'################################################################################################################
Public Function zlDataToPrint(vsfPrint As VSFlexGrid, VsfData As VSFlexGrid) As Boolean
    '-------------------------------------------------
    '将全部组强制展开,复制数据表格
    
    Dim lngCol As Long, lngRow As Long
    Dim lngPrintCol As Long, lngPrintRow As Long
    On Error GoTo ErrHand:
    
    With vsfPrint
        .Clear
        .MergeCells = flexMergeFixedOnly ' = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFree
        .Rows = VsfData.Rows
        .Cols = 0: .Cols = VsfData.Cols
        .FixedCols = VsfData.FixedCols
        .FixedRows = VsfData.FixedRows
        .GridColor = vbBlack
        
        '标题行复制
        .ROW = 0
        .Rows = VsfData.Rows
        .Cols = VsfData.Cols
        lngPrintRow = 0
        For lngRow = 0 To VsfData.Rows - 1
            If Not VsfData.RowHidden(lngRow) Then
                lngPrintCol = 0
                For lngCol = 0 To VsfData.Cols - 1
                    If Not VsfData.ColHidden(lngCol) Then
                        .TextMatrix(lngPrintRow, lngPrintCol) = VsfData.TextMatrix(lngRow, lngCol)
                        .ColWidth(lngPrintCol) = VsfData.ColWidth(lngCol)
                         .ColAlignment(lngPrintCol) = VsfData.ColAlignment(lngCol)
                        lngPrintCol = lngPrintCol + 1
                    End If
                Next
                .RowHeight(lngPrintRow) = VsfData.RowHeight(lngRow)
                lngPrintRow = lngPrintRow + 1
            End If
        Next
        
        lngPrintCol = 0
        For lngCol = 0 To .Cols - 1
           If VsfData.ColHidden(lngCol) Then lngPrintCol = lngPrintCol + 1
        Next
        .Cols = VsfData.Cols - lngPrintCol
         lngPrintRow = 0
        For lngRow = 0 To .Rows - 1
           If VsfData.RowHidden(lngRow) Then lngPrintRow = lngPrintRow + 1
        Next
        .Rows = VsfData.Rows - lngPrintRow
        
        lngPrintRow = 0
        For lngRow = 0 To .FixedRows - 1
           If VsfData.RowHidden(lngRow) Then lngPrintRow = lngPrintRow + 1
        Next
        .FixedRows = .FixedRows - lngPrintRow
        If .FixedRows = 0 Then .FixedRows = 1
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
       
        '再按行合并
        For lngRow = 0 To .FixedRows - 1
            .MergeRow(lngRow) = True
        Next
        
        
    End With
    zlDataToPrint = True
    Exit Function

ErrHand:
    zlDataToPrint = False
End Function


Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'功能：显示或隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.hWnd, vRect)
    lngStyle = GetWindowLong(objForm.hWnd, GWL_STYLE)
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
    SetWindowLong objForm.hWnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.hWnd, 0, 0, 0, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Public Function IsAllowInput(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strTime As String, ByVal strCurTime As String) As Boolean
    '取出指定病人在指定时间之后关键点的时间
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
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
ErrHand:
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

Public Function GetAdviceOutTime(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer) As String
'功能:获取病人或婴儿的医嘱出院时间
    Dim rsTemp As New ADODB.Recordset
    Dim strTmp As String, strTime As String
    On Error GoTo ErrHand
    If int婴儿 = 0 Then
        strTmp = ",5,11,"
    Else
        strTmp = ",3,5,11,"
    End If
    gstrSQL = "Select 开始执行时间" & vbNewLine & _
        " From 病人医嘱记录 b, 诊疗项目目录 c" & vbNewLine & _
        " Where b.诊疗项目id + 0 = c.Id And b.医嘱状态 = 8 And Nvl(b.婴儿, 0) <> 0 And c.类别 = 'Z' And instr([4],',' || c.操作类型 || ',',1)>0 And" & vbNewLine & _
        "      b.病人id = [1] And b.主页id = [2] And b.婴儿 = [3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人医嘱出院时间", lng病人ID, lng主页ID, int婴儿, strTmp)
    If rsTemp.RecordCount > 0 Then strTime = Format(rsTemp!开始执行时间, "YYYY-MM-DD HH:mm:ss")
    GetAdviceOutTime = strTime
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function EprIsCommit(ByVal lngPatiID As Long, ByVal lngPageId As Long) As String
'以|分隔方式返回,状态为0 不允许 1 允许，分别控制 新增|删除|撤档

    Dim rsTemp As ADODB.Recordset
    Dim intNew As Integer, intDel As Integer, intMod As Integer
    Dim strState As String
    
    EprIsCommit = "1|1|1"
    strState = "未审查": gstrMecState = strState
    On Error GoTo ErrHand
    gstrSQL = "Select 病案状态 From 病案主页 Where 病人id = [1] And 主页id = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "病案状态", lngPatiID, lngPageId)
    If Not rsTemp.EOF Then
        Select Case NVL(rsTemp!病案状态, 0)
            Case 0
                intNew = 1: intDel = 1: intMod = 1
                strState = "未审查"
            Case 1 '等待审查
                intNew = 0: intDel = 0: intMod = 0
                strState = "等待审查"
            Case 2 '拒绝审查
                intNew = 1: intDel = 1: intMod = 1
                strState = "拒绝审查"
            Case 3 '正在审查
                intNew = 0: intDel = 0: intMod = 0
                strState = "正在审查"
            Case 4 '审查反馈
                intNew = 0: intDel = 0: intMod = 1
                strState = "审查反馈"
            Case 5 '审查归档
                intNew = 0: intDel = 0: intMod = 0
                strState = "审查归档"
            Case 6 '审查整改
                intNew = 0: intDel = 0: intMod = 1
                strState = "审查整改"
            Case 13 '正在抽查
                intNew = 1: intDel = 1: intMod = 1
                strState = "正在抽查"
            Case 14 '抽查反馈
                intNew = 1: intDel = 1: intMod = 1
                strState = "抽查反馈"
            Case 16 '抽查整改
                intNew = 1: intDel = 1: intMod = 1
                strState = "抽查整改"
            Case 10 '接收待审
                intNew = 0: intDel = 0: intMod = 0
                strState = "接收待审"
            Case Else
                intNew = 0: intDel = 0: intMod = 0
                strState = "审查"
        End Select
    End If
    gstrMecState = strState
    EprIsCommit = CStr(intNew) & "|" & CStr(intDel) & "|" & CStr(intMod)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ISCollectSigned(ByVal lngFileID As Long, ByVal strDate As String, ByVal strTime As String) As Boolean
    Dim blnDetail As Boolean
    Dim str发生时间 As String, strStart As String, strEnd As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    str发生时间 = Format(strDate, "yyyy-MM-dd")
    
    strStart = str发生时间 & " 00:00:00"
    strEnd = Format(DateAdd("d", 2, strStart), "yyyy-MM-dd HH:mm:ss")
    str发生时间 = str发生时间 & " " & strTime & ":00"
    
    gstrSQL = " Select A.发生时间,A.开始时点,A.结束时点,A.汇总类别,B.类别,B.开始,B.结束,A.签名人" & vbNewLine & _
              " From 病人护理数据 A,护理汇总时段 B" & vbNewLine & _
              " Where B.单据(+)=2 And abs(A.汇总类别)=B.类别(+) And A.汇总类别<0 And A.签名人 Is Not NULL And A.文件ID=[1] And A.发生时间 between [2] and [3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取当前记录发生时间当天及之后一天的汇总数据", lngFileID, CDate(strStart), CDate(strEnd))
    With rsTemp
        '循环检查，发现退出
        Do While Not .EOF
            '拼开始结束时间串
            If IsNull(!类别) Then
                strEnd = Format(!发生时间, "YYYY-MM-DD") & " " & !结束时点 & ":59"
                If !结束时点 < !开始时点 Then
                    strStart = Format(DateAdd("d", -1, !发生时间), "YYYY-MM-DD") & " " & !开始时点 & ":00"
                Else
                    strStart = Format(!发生时间, "YYYY-MM-DD") & " " & !开始时点 & ":00"
                End If
            Else
                strEnd = Format(!发生时间, "YYYY-MM-DD") & " " & !结束 & ":59"
                If !结束 < !开始 Then
                    strStart = Format(DateAdd("d", -1, !发生时间), "YYYY-MM-DD") & " " & !开始 & ":00"
                Else
                    strStart = Format(!发生时间, "YYYY-MM-DD") & " " & !开始 & ":00"
                End If
            End If
            
            If str发生时间 >= strStart And str发生时间 <= strEnd Then
                ISCollectSigned = True
                Exit Function
            End If
            .MoveNext
        Loop
    End With
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

