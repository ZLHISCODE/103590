Attribute VB_Name = "mdlPublic"
Option Explicit

Public gstrUnitName As String       '当前用户单位名称
Public gfrmMain As Object           '导航台窗体
Public gobjEmr As Object                    '新版电子病历
Public gcnOracle As ADODB.Connection  '数据库连接
Public gstrSysName As String                '系统名称，例如：中联软件
Public gstrProductName As String            '产品简称，例如：中联
Public glngModul As Long                    '模块编号
Public glngSys As Long                      '系统编号，例如：100
Public gstrDBUser As String
Public gstrPrivs As String                     '用户在该模块下面的权限
Public gblnShowInTaskBar As Boolean         '是否显示窗体在任务条上
Public UserInfo As TYPE_USER_INFO            '用户信息
Public gobjFSO As New Scripting.FileSystemObject    'FSO对象
Public gMainPrivs As String
Public gstrNodeNo As String          '当前站点编号；如果未设置启用站点，则为"-"
Private mclsZip As New cZip
Private mclsUnzip As New cUnzip
Public gclsMipModule As zl9ComLib.clsMipModule
Public gstrLike As String  '项目匹配方法,%或空
Public gbytCode As Byte '简码输入方式
Public gstrDBOwer As String
Public gobjComlib As Object
Public gobjLIS As Object
Public glngPreHWnd As Long '用于支持鼠标滚轮功能
Public glngOpenedID As Long '医生站处理时打开的反馈单ID
Public gObjRichEPR As zlRichEPR.cRichEPR

'改变窗体位置、Zorder、尺寸等
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = -4
Public Const WM_MOUSEWHEEL = &H20A


Public Type TYPE_USER_INFO
    ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    性质 As String
    部门ID As Long
    部门码 As String
    部门名 As String
    专业技术职务 As String
    用药级别 As Long
End Type

Public Function InitSysPar() As Boolean
'功能：初始化系统参数
'返回：真-处理成功
    Dim strTmp As String
    gstrLike = IIf(gobjComlib.zlDatabase.GetPara("输入匹配") = "0", "%", "")
    gbytCode = Val(gobjComlib.zlDatabase.GetPara("简码方式"))
    InitSysPar = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    Set rsTmp = gobjComlib.zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.用户名 = rsTmp!User
            UserInfo.编号 = rsTmp!编号
            UserInfo.简码 = NVL(rsTmp!简码)
            UserInfo.姓名 = NVL(rsTmp!姓名)
            UserInfo.部门ID = NVL(rsTmp!部门ID, 0)
            UserInfo.部门码 = NVL(rsTmp!部门码)
            UserInfo.部门名 = NVL(rsTmp!部门名)
            UserInfo.性质 = Get人员性质
            UserInfo.专业技术职务 = NVL(rsTmp!专业技术职务)
            GetUserInfo = True
        End If
    End If
    Exit Function
errH:
   If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get人员性质(Optional ByVal str姓名 As String) As String
'功能：读取当前登录人员或指定人员的人员性质
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    If str姓名 <> "" Then
        strSQL = "Select B.人员性质 From 人员表 A,人员性质说明 B Where A.ID=B.人员ID And A.姓名=[1]"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str姓名)
    Else
        strSQL = "Select 人员性质 From 人员性质说明 Where 人员ID = [1]"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.ID)
    End If
    Do While Not rsTmp.EOF
        Get人员性质 = Get人员性质 & "," & rsTmp!人员性质
        rsTmp.MoveNext
    Loop
    Get人员性质 = Mid(Get人员性质, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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

    On Error GoTo errHand:
    For Each rptRow In rptList.Rows
        If rptRow.GroupRow Then rptRow.Expanded = True
    Next
    With vfgList
        .Clear
        .Rows = rptList.Records.Count + 1
        .Cols = 0: .Cols = rptList.Columns.Count
        .FixedCols = rptList.GroupsOrder.Count
        '标题行复制
        .Row = 0
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
errHand:
    zlReportToVSFlexGrid = False
End Function
'
Public Function DynamicCreate(ByVal strClass As String, ByVal strCaption As String, Optional ByVal blnMsg As Boolean) As Object
'动态创建对象
    On Error Resume Next
    Set DynamicCreate = CreateObject(strClass)
    If Err <> 0 Then
        If blnMsg Then MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, gstrSysName
        Set DynamicCreate = Nothing
    End If
    Err.Clear
End Function

Public Function MovedByDate(ByVal vDate As Date) As Boolean
'功能：判断指定日期之前的是否可能已经执行了数据转出
'参数：vDate=时间点或时间段的开始时间
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    strSQL = "Select 上次日期 From zlDataMove Where 系统=[1] And 组号=1 And 上次日期 is Not Null"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", glngSys)
    If Not rsTmp.EOF Then
        '上次日期没有时点,"<"判断与转出过程中一致
        If vDate < rsTmp!上次日期 Then
            MovedByDate = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetDbOwner(ByVal lngSys As Long) As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String

    GetDbOwner = ""
    On Error GoTo errHand
    strSQL = "Select 所有者 From Zlsystems Where 编号 = [1]"
    Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "GetDbOwner", lngSys)
    If rsTemp.RecordCount <> 0 Then GetDbOwner = "" & rsTemp!所有者
    rsTemp.Close
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


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
    Dim strSQL As String

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

    strSQL = "Select Zl_Lob_Read([1],[2],[3],[4]) as 片段 From Dual"
    lngCount = 0
    Do
        Set rsLob = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "zlBlobRead", Action, KeyWord, lngCount, IIf(blnMoved, 1, 0))
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
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset

    strSQL = "Select Zl_Replace_Element_Value([1],[2],[3],[4],[5],[6]) From Dual"
    Err = 0: On Error GoTo DBError
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "读取替换项", ElementName, CLng(sPatientID), _
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
'## 功能：  搜索整个文本给出指定关键字区域的定位信息
'##
'## 参数：  edtThis         :   IN  ，编辑控件
'##         strKeyType      :   IN  ，给定关键字名称。取值为："O"、"P"、"T"、"E"、"U"
'##         lngKey           :   IN  ，给定欲查找的关键字ID号。
'##         lngKSS、lngKSE  :   OUT ，分别表示起始关键字的开始位置和结束位置；
'##         lngKES、lngKEE  :   OUT ，分别表示终止关键字的开始位置和结束位置；
'##         blnNeeded:      :   OUT ，是否是保留对象
'##
'## 返回：  如果找到该关键字具体位置，则返回True，否则返回False
'################################################################################################################
Public Function FindKey(ByRef edtThis As Object, _
        ByRef strKeyType As String, _
        ByRef lngKey As Long, _
        ByRef lngKSS As Long, _
        ByRef lngKSE As Long, _
        ByRef lngKES As Long, _
        ByRef lngKEE As Long, _
        ByRef blnNeeded As Boolean) As Boolean

    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！

    sTMP = strKeyType & "S(" & Format(lngKey, "00000000")
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = 1
LL1:
        i = InStr(i, sText, sTMP)
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
            sTMP = strKeyType & "E(" & Format(lngKey, "00000000")
            j = InStr(j, sText, sTMP)
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

Public Function GetDeptID(ByVal strDeptCode As String) As Long
'功能：根据部门编码获取部门ID
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
On Error GoTo errH
    
    strSQL = "Select a.Id, a.名称 From 部门表 A Where A.编码 = [1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "部门查询", strDeptCode)

    If rsTmp.RecordCount > 0 Then
        GetDeptID = rsTmp!ID
    Else
        GetDeptID = 0
        MsgBox "没有查询到编码为“" & strDeptCode & "”的部门科室，请联系管理员对码！", vbInformation, gstrSysName
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlGetComLib() As Boolean
    If Not gobjComlib Is Nothing Then
        Call gobjComlib.InitCommon(gcnOracle)
        zlGetComLib = True
        Exit Function
    End If
    On Error Resume Next
    Set gobjComlib = CreateObject("zl9Comlib.clsComlib")
    Call gobjComlib.InitCommon(gcnOracle)
    zlGetComLib = True
End Function
 
Public Function FlexScroll(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'功能：支持滚轮的滚动
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
        Case -7864320  '向下滚
            gobjComlib.ZLCommFun.PressKey vbKeyPageDown
        Case 7864320   '向上滚
            gobjComlib.ZLCommFun.PressKey vbKeyPageUp
        End Select
    End Select
    FlexScroll = CallWindowProc(glngPreHWnd, hwnd, wMsg, wParam, lParam)
End Function

Public Function CheckOperateState(ByVal lngID As Long, ByRef intCode As Integer) As Boolean
'功能: 查询是否能够处理该反馈单（删除或者修改）
'参数: lngID - 反馈单ID ；intCode - 不能操作的原因 ；1-未查找到；2-他人的反馈单；3-医生已经处理
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    '读取反馈单相关信息
    On Error GoTo errH
    strSQL = "Select a.Id, a.记录状态, a.登记人 From 疾病阳性记录 A  Where a.Id = [1] "

    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "查询阳性结果反馈单", lngID)
    
    If rsTmp.RecordCount > 0 Then
        If UserInfo.姓名 <> NVL(rsTmp!登记人) Then
            intCode = 2
            Exit Function
        ElseIf rsTmp!记录状态 > 1 Then
            intCode = 3
            Exit Function
        End If
    Else
        intCode = 1
        Exit Function
    End If
    CheckOperateState = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub PrintDiseaseRegist(ByVal intType As Integer, ByVal lngID As Long, ByRef frmParent As Object)
'功能: 打印阳性结果反馈单
'参数：lngID : 反馈单ID；intType:1-预览，2-打印
    Dim objReport As clsReport
    
    On Error GoTo errH
  
    If objReport Is Nothing Then Set objReport = New clsReport
    Call objReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1278_1", frmParent, "反馈单ID=" & lngID, intType)
    If Not objReport Is Nothing Then Set objReport = Nothing
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CheckDisNum(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long, ByRef rsDisease As ADODB.Recordset, Optional ByVal lngID As Long) As Boolean
'功能: 检查该病人有多少没有填写报告卡的反馈单
'lngPatFrom: 2-住院, 1-门诊
    Dim strSQL As String
    On Error GoTo errH
    If lngID <> 0 Then
        strSQL = " and a.ID = " & lngID
    End If
    If lngPatFrom = 1 Then
        strSQL = "select rowNum as NO,a.ID,c.名称 as 科室, a.登记时间, a.记录状态, a.处理情况说明 from  疾病阳性记录 A ,病人挂号记录 B ,部门表 C where A.文件ID is NULL  and A.挂号单 = B.NO and A.病人ID = B.病人ID and A.记录状态 <> 3 and A.登记科室ID = C.ID  and A.病人ID = [1] and B.ID = [2]" & strSQL
    ElseIf lngPatFrom = 2 Then
        strSQL = "select rowNum as NO,a.ID ,c.名称 as 科室,a.登记时间, a.记录状态, a.处理情况说明 from  疾病阳性记录 A ,部门表 C  where A.文件ID is NULL  and A.记录状态 <> 3  and A.登记科室ID = C.ID and A.病人ID = [1] and A.主页ID = [2] " & strSQL
    End If
    Set rsDisease = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "查询阳性结果反馈单", lngPatiID, lngPageId)
    
    If rsDisease.RecordCount > 0 Then
        CheckDisNum = True
    Else
        CheckDisNum = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SaveReason(ByVal strReason As String, ByVal lngID As Long, ByVal lngState As Long) As Boolean
'功能: 存储不填写报告卡的原因
'参数：strReason-原因；lngID-反馈单ID ；lngState-反馈单当前的记录状态
    Dim strSQL As String
    Dim str处理时间 As String
    Dim str处理医生 As String
    Dim str处理情况 As String, strTmp As String

    On Error GoTo errH
    str处理时间 = "to_date('" & Format(gobjComlib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
    str处理医生 = "'" & UserInfo.姓名 & "'"
    str处理情况 = "'" & strReason & "'"
    lngState = IIf(lngState = 1, 2, lngState)

    strSQL = "Zl_疾病阳性检测记录_update(1," & lngID & "," & "NULL" & "," & CStr(lngState) & "," & str处理医生 & "," & str处理时间 & "," & str处理情况 & ")"
    Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, "保存反馈单的处理情况")
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function QueryRegistByPati(ByRef frmParent As Object, ByVal intType As Integer, ByVal lng病人ID As Long, _
                            Optional ByVal lng主页ID As Long, Optional ByVal str挂号No As String, Optional ByVal var登记科室 As Variant) As Boolean
    Dim strSQL As String
    Dim lng登记科室ID As Long
    Dim rsDisease As ADODB.Recordset
    Dim lngID As Long

    On Error GoTo errH
    
    If TypeName(var登记科室) = "String" Then         '传的编码
        lng登记科室ID = GetDeptID(var登记科室)
    ElseIf IsNumeric(var登记科室) Then
        lng登记科室ID = Val(var登记科室)
    Else
        lng登记科室ID = 0
    End If
    
    If lng主页ID <> 0 Then
        strSQL = " Select a.Id, '住院' As 来源, c.病人id, c.姓名, c.性别, c.年龄, e.名称 As 科室, c.住院号 As 标识号, a.送检时间, a.送检医生, a.记录状态, f.名称 As 登记科室," & vbNewLine & _
                 "       a.标本名称, a.反馈结果, a.传染病名称 As 疑似疾病, a.登记人, a.登记时间, a.处理人, a.处理时间, a.处理情况说明" & vbNewLine & _
                 " From 疾病阳性记录 A, 病案主页 C, 部门表 E, 部门表 F" & vbNewLine & _
                 " Where a.病人id = c.病人id And a.主页id = c.主页id And c.病人id = [1] And C.主页id = [2]  And a.登记科室ID = f.Id(+) And" & vbNewLine & _
                 "      c.出院科室id = e.Id(+)" & IIf(lng登记科室ID <> 0, " and a.登记科室id =[3] ", "") & vbNewLine & _
                 " Order By a.登记时间 Desc"
        Set rsDisease = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "查询阳性结果反馈单", lng病人ID, lng主页ID, lng登记科室ID)
    ElseIf str挂号No <> "" Then
        strSQL = " Select a.Id, '门诊' As 来源, b.病人id, b.姓名, b.性别, b.年龄, e.名称 As 科室, b.门诊号 As 标识号, a.送检时间, a.送检医生, a.记录状态, f.名称 As 登记科室," & vbNewLine & _
                 "       a.标本名称, a.反馈结果, a.传染病名称 As 疑似疾病, a.登记人, a.登记时间, a.处理人, a.处理时间, a.处理情况说明" & vbNewLine & _
                 " From 疾病阳性记录 A, 病人挂号记录 B, 部门表 E, 部门表 F" & vbNewLine & _
                 " Where a.病人id = b.病人id And a.挂号单 = b.No And b.病人id = [1] And b.No = [2] And a.登记科室ID = f.Id(+) And" & vbNewLine & _
                 "      b.执行部门id = e.Id(+)" & IIf(lng登记科室ID <> 0, " and a.登记科室id =[3] ", "") & vbNewLine & _
                 " Order By a.登记时间 Desc"
        Set rsDisease = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "查询阳性结果反馈单", lng病人ID, str挂号No, lng登记科室ID)
    End If
    If rsDisease.RecordCount = 0 Then
        MsgBox "该病人没有查询到反馈单记录。", vbInformation, gstrSysName
    ElseIf rsDisease.RecordCount = 1 Then
        lngID = Val(rsDisease!ID)
        Call frmDiseaseRegist.ShowDiseaseRegist(frmParent, intType, lngID)
    Else
        lngID = frmDiseaseQuery.ShowPatiDis(rsDisease, frmParent)
        If lngID <> 0 Then
            Call frmDiseaseRegist.ShowDiseaseRegist(frmParent, intType, lngID)
        End If
    End If
    
    QueryRegistByPati = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

