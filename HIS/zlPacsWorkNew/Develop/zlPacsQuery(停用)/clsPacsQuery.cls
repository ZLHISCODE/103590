VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacsQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const LOCAL_BAK_TABLES_FILE_NAME As String = "zlBakTables.xml"

Private WithEvents mobjSqlParseEvent As clsSqlParse
Attribute mobjSqlParseEvent.VB_VarHelpID = -1
Private mobjCurFilterValue As clsSqlFilterValue '用于参数取值

Private maryScheme() As TSchemeItem
Private maryBakTables() As String   '保存转储表的名称

Private mstrStationName As String
Private mdblFontSize As Double

Private mlngCurSchemeId As Long         '存储当前使用的方案ID
Private mlngUserId As Long

Private mblnIsMoved As Boolean

Public Event OnGetParameterValue(ByVal strParName As String, ByRef Value As Variant)

Public Event OnQueryState(ByVal lngCurState As Long, ByVal strStateMsg As String)

'方案数量
Property Get SchemeCount() As Long
    SchemeCount = UBound(maryScheme)
End Property

'是否转储
Property Get IsMoved() As Boolean
    IsMoved = mblnIsMoved
End Property

''获取方案结构
'Property Get SchemeItem(ByVal lngIndex As Long) As TSchemeItem
'    SchemeItem = maryScheme(lngIndex)
'
'    If SchemeItem.IsLoaded = False Then
'        Call ParseScheme(SchemeItem, True)
'        maryScheme(lngIndex) = SchemeItem
'    End If
'End Property



'获取方案基本信息
Property Get SchemeInfo(ByVal lngIndex As Long) As TSchemeBaseInfo
    SchemeInfo = maryScheme(lngIndex).BaseInfo
End Property

'字体大小
Property Get FontSize() As Double
    FontSize = mdblFontSize
End Property

Property Let FontSize(ByVal Value As Double)
    mdblFontSize = Value
End Property


'***************************************************************************************


Public Sub init(cnOracle As ADODB.Connection, ByVal lngUserId As Long, ByVal strCachePath As String)
    Set gcnOracle = cnOracle
    
    mlngUserId = lngUserId
    gstrCachePath = strCachePath
    
    Call LoadBakTabsInfo(gstrCachePath)
End Sub

Private Sub LoadBakTabsInfo(ByVal strCachePath As String)
'读取转存表信息
    Dim rsData As New ADODB.Recordset
    Dim strSql As String
    Dim strCacheFile As String
    Dim i As Long
        
    '判断本地文件是否缓存
    strCacheFile = Replace(gstrCachePath & "\" & LOCAL_BAK_TABLES_FILE_NAME, "\\", "\")

    
    If Dir(strCacheFile) <> "" Then
        Call rsData.Open(strCacheFile)
    Else
        strSql = "Select 表名 From zlBakTables"
        Set rsData = ExecuteSql(strSql, "查询转储表信息")
        
        If rsData Is Nothing Then Exit Sub
        If rsData.RecordCount <= 0 Then Exit Sub

        Call rsData.Save(strCacheFile)
    End If
    
    rsData.MoveFirst
    
    ReDim maryBakTables(rsData.RecordCount)
    
    i = 1
    While Not rsData.EOF
        maryBakTables(i) = UCase(NVL(rsData!表名))
        i = i + 1
        
        Call rsData.MoveNext
    Wend
End Sub

Private Sub UpdateBakTablesCache()
'更新转储表信息
    Dim strCacheFile As String
    Dim fso As New FileSystemObject
    
    strCacheFile = Replace(gstrCachePath & "\" & LOCAL_BAK_TABLES_FILE_NAME, "\\", "\")
    
    If Format(fso.GetFile(strCacheFile).DateCreated, "yyyymmdd") <> Format(Now, "yyyymmdd") Then
        Call LoadBakTabsInfo(gstrCachePath)
    End If
End Sub

Public Function ShowSchemeCfg(ByVal lngModuleNo As Long, ByVal strSysParDefines As String, owner As Object) As Boolean
'显示方案配置...
'lngModuleNo:模块号
'strSysParDefines:系统参数定义，多个参数使用“逗号分割”，如“[系统.系统号],[系统.模块号]...”
'如果有调整返回true，没有调整返回false
On Error GoTo errHandle
    Dim objSchemeCfg As New frmSchemeCfg
    
    objSchemeCfg.ShowMe lngModuleNo, strSysParDefines, owner
Exit Function
errHandle:
    Set objSchemeCfg = Nothing
    Err.Raise -1, "clsPacsQuery.ShowSchemeCfg", "方案配置调用失败:" & Err.Description
End Function

Public Function ShowUserScheme(ByVal lngModuleNo As Long, owner As Object) As Boolean
'显示用户常用查询方案配置...
'如果有调整返回true，没有调整返回false
On Error GoTo errHandle
    Dim objQueryCfg As New frmUserQueryReleation
    
    ShowUserScheme = objQueryCfg.ShowUserScheme(owner, lngModuleNo, mlngUserId)
Exit Function
errHandle:
    Set objQueryCfg = Nothing
    Err.Raise -1, "clsPacsQuery.ShowUserScheme", "用户查询方案关联调用失败:" & Err.Description
End Function

Public Function GetIconRes(ByVal strResName As String) As StdPicture
'获取图标资源
On Error GoTo errHandle
    Dim strFile As String
    
    Set GetIconRes = Nothing
    
    strFile = zlIconResRead(strResName)
    If Len(Dir(strFile)) <> 0 Then
        Set GetIconRes = LoadPicture(strFile)
    End If
Exit Function
errHandle:
    Set GetIconRes = Nothing
    Debug.Print "clsPacsQuery.GetIconRes Err:" & Err.Description
End Function

Public Sub LoadQueryScheme(ByVal lngModuleNo As Long)
'载入查询方案到sqlScheme和sqlParse对象中
On Error GoTo errHandle
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    ReDim maryScheme(0)
    
    strSql = "Select A.ID, A.方案名称, A.方案说明, A.版本, Nvl(B.是否默认,A.是否默认) As 是否默认, Nvl(B.是否常用,A.是否常用) As 是否常用, B.所属站点,  B.用户ID " & _
                " From 影像查询方案 A, 影像查询关联 B " & _
                " Where A.ID=B.查询方案ID(+) And A.使用状态=1 And A.所属模块=[1] And B.用户ID(+)=[2] order by 是否默认 Desc, 是否常用, 方案序号"
    

    Set rsData = ExecuteSql(strSql, "载入所有方案", lngModuleNo, mlngUserId)
    
    If rsData Is Nothing Then Exit Sub
    If rsData.RecordCount <= 0 Then Exit Sub
    
    rsData.Filter = "用户ID=" & mlngUserId
    
    If rsData.RecordCount <= 0 Then
        rsData.Filter = ""
    End If
    
    Call ChangeCurScheme(Val(NVL(rsData!Id)))
    
    '对方案进行筛选
    Call WriteScheme(rsData, True)
    
    If rsData.Filter <> "" Then
        rsData.Filter = "用户ID = NULL"
        If rsData.RecordCount <= 0 Then Exit Sub
        
        Call WriteScheme(rsData, False)
        
    End If
    
Exit Sub
errHandle:
    Err.Raise -1, "clsPacsQuery.LoadQueryScheme", "[LoadQueryScheme]处理错误>>" & vbCrLf & Err.Description
    Resume
End Sub

Private Sub WriteScheme(rsData As ADODB.Recordset, Optional ByVal blnIsReadDefaultSet As Boolean = False)
    Dim objSchemeItem As TSchemeItem
    Dim strScName As String
    Dim strSchemeStationName As String
    
    While Not rsData.EOF
        strScName = NVL(rsData!方案名称)
        strSchemeStationName = NVL(rsData!所属站点)
        
        If strSchemeStationName = "" Or UCase(strSchemeStationName) = mstrStationName Then
            
            objSchemeItem.IdKey = Val(NVL(rsData!Id))
            objSchemeItem.BaseInfo.SchemeId = objSchemeItem.IdKey
            objSchemeItem.BaseInfo.Ver = Val(NVL(rsData!版本))
            objSchemeItem.BaseInfo.Name = NVL(rsData!方案名称)
            objSchemeItem.BaseInfo.Memo = NVL(rsData!方案说明)
            
            If blnIsReadDefaultSet Then
                objSchemeItem.BaseInfo.IsDefault = IIf(Val(NVL(rsData!是否默认)) <> 0, True, False)
                objSchemeItem.BaseInfo.IsOften = IIf(Val(NVL(rsData!是否常用)) <> 0, True, False)
            Else
                objSchemeItem.BaseInfo.IsDefault = False
                objSchemeItem.BaseInfo.IsOften = False
            End If
            
            Call WriteSchemeItemArray(objSchemeItem)
        End If
    
        Call rsData.MoveNext
    Wend
End Sub


Public Function GetQueryField(Optional objSqlParse As clsSqlParse) As ADODB.Recordset
'获取返回的查询字段
    Dim strSql As String
    
    Set GetQueryField = Nothing
    
    If objSqlParse Is Nothing Then Exit Function
    
    objSqlParse.NullParameterHint = False
    objSqlParse.NullWhereHint = False
    
    strSql = objSqlParse.GetTestSql
    
    If Len(strSql) <= 0 Then
        Set GetQueryField = Nothing
        Exit Function
    End If
    
    Set GetQueryField = ExecuteSql(strSql, "获取查询字段")
End Function

Public Function GetFilterValue(ByVal strFilterName As String) As Variant
'获取方案对应的参数值
    Dim i As Long
'    Dim lngCurSchemeId As Long
'
'    lngCurSchemeId = lngSchemeId
'    If lngSchemeId <= 0 Then lngCurSchemeId = mlngCurSchemeId
    
    For i = 1 To UBound(maryScheme)
        If maryScheme(i).IdKey = mlngCurSchemeId Then
            GetFilterValue = maryScheme(i).FilterValues.ParData(strFilterName)
            Exit Function
        End If
    Next i
End Function

Public Sub SetFilterValue(ByVal strFilterName As String, Value As Variant)
'设置方案对应的参数值
    Dim i As Long
'    Dim lngCurSchemeId As Long
'
'    lngCurSchemeId = lngSchemeId
'    If lngSchemeId <= 0 Then lngCurSchemeId = mlngCurSchemeId
    
    For i = 1 To UBound(maryScheme)
        If maryScheme(i).IdKey = mlngCurSchemeId Then
            If Not maryScheme(i).FilterValues Is Nothing Then
                maryScheme(i).FilterValues.ParData(strFilterName) = Value
                
                If Not maryScheme(i).FilterWindow Is Nothing Then
                    Call maryScheme(i).FilterWindow.UpdateInputData(strFilterName, Value)
                End If
                
                Exit Sub
            End If
        End If
    Next i
End Sub


Public Sub ChangeCurScheme(ByVal lngSchemeId As Long)
'设置当前使用的方案ID
    mlngCurSchemeId = lngSchemeId
End Sub


Private Sub LoadSchemeWhereConfig(ByRef curSchemeItem As TSchemeItem)
'载入方案对应的条件配置
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim objFilterValue As clsSqlFilterValue
    
    Set objFilterValue = New clsSqlFilterValue
    Set curSchemeItem.FilterValues = objFilterValue
    
    maryScheme(curSchemeItem.Index) = curSchemeItem
    
    strSql = "Select 条件配置 From 影像查询特性 Where 用户ID=[1] And 查询方案ID=[2]"
    
    Set rsData = ExecuteSql(strSql, "查询方案条件配置", mlngUserId, curSchemeItem.IdKey)
    
    If rsData Is Nothing Then Exit Sub
    If rsData.RecordCount <= 0 Then Exit Sub
    
    If objFilterValue.LoadStorageString(NVL(rsData!条件配置)) Then
        maryScheme(curSchemeItem.Index) = curSchemeItem
    End If
    
End Sub

Public Sub SaveIndividuation()
'保存个性化
    Dim i As Long
    
    For i = 1 To UBound(maryScheme)
        Call SaveSchemeWhereConfig(maryScheme(i))
    Next i
End Sub

Private Sub SaveSchemeWhereConfig(curSchemeItem As TSchemeItem)
'保存方案对应的条件配置
    Dim strCfg As String
    
    If curSchemeItem.FilterValues Is Nothing Then Exit Sub
    If curSchemeItem.FilterValues.IsUpdate = False Then Exit Sub
    
    strCfg = curSchemeItem.FilterValues.GetStorageString()
    
    Call ExecuteCmd("zl_影像查询_条件配置(" & mlngUserId & "," & curSchemeItem.IdKey & ",'" & strCfg & "')", "更新方案条件配置")
End Sub

Private Function GetSqlWithMoved(ByVal strSource As String) As String
'获取转储后的sql语句
    Dim i As Long
    Dim lngIndex As Long
    Dim strResult As String
    
    strResult = strSource
    For i = 1 To UBound(maryBakTables)
        lngIndex = InStr(strResult, maryBakTables(i))
        If lngIndex > 0 Then
            strResult = Replace(strResult, maryBakTables(i), "H" & maryBakTables(i))
        End If
    Next i
    
    GetSqlWithMoved = IIf(Len(strResult) <> Len(strSource), strResult, "")
End Function

Public Function Execute(ByRef dtStartdate As Date, ByRef dtEndDate As Date, _
    Optional ByVal blnIsRefresh As Boolean = False) As ADODB.Recordset
'执行方案
On Error GoTo errHandle
    Dim objSchemeItem As TSchemeItem
    Dim objCurParse As clsSqlParse
    Dim strSql As String
    Dim strMovedSql As String
    
    Set Execute = Nothing
    
    
    Call DoQueryState(0, "读取方案配置.")
    objSchemeItem = ReadSchemeItemWithId(mlngCurSchemeId)
    
    If objSchemeItem.IdKey <= 0 Then
        Err.Raise -1, "clsPacsQuery.Execute", "未查询到对应的方案配置信息。"
    End If
    
    objSchemeItem.Startdate = dtStartdate
    objSchemeItem.EndDate = dtEndDate
    
    Call DoQueryState(0, "读取配置查询条件.")
    If objSchemeItem.FilterValues Is Nothing Then
    '载入方案默认对应的条件录入配置
        Call LoadSchemeWhereConfig(objSchemeItem)
    End If
    
    '配置查询日期范围
    objSchemeItem.FilterValues.ParData("系统.开始日期") = dtStartdate
    objSchemeItem.FilterValues.ParData("系统.结束日期") = dtEndDate
    
    Call DoQueryState(0, "解析查询语句.")
    Set objCurParse = objSchemeItem.SqlParse
    
    If objCurParse.IsInit = False Then
        strSql = objSchemeItem.SqlScheme.Query
        Call objCurParse.init(strSql)
    End If
    
    Set mobjCurFilterValue = objSchemeItem.FilterValues
    Set mobjSqlParseEvent = objCurParse
    
    Call DoQueryState(0, "组织查询SQL.")
    strSql = mobjSqlParseEvent.GetQuerySql(blnIsRefresh)
    
    mblnIsMoved = zlDatabase.DateMoved(dtEndDate) '  True '读取是否转储标志...
    If mblnIsMoved Then
        '针对转储数据的查询，可能需要特殊处理
        strMovedSql = GetSqlWithMoved(strSql)
        If strMovedSql <> "" Then
            strSql = "(" & strSql & ") " & vbCrLf & " Union All " & vbCrLf & " (" & strMovedSql & ")"
        End If
    End If
    
    Set mobjCurFilterValue = Nothing
    
    Call DoQueryState(0, "获取查询结果.")
    Set Execute = QuerySchemeData(strSql)
Exit Function
errHandle:
    Err.Raise -1, "clsPacsQuery.Execute", "[Execute]处理错误>>" & vbCrLf & "  方案ID为：" & mlngCurSchemeId & vbCrLf & Err.Description
    Resume
End Function

Private Sub HideBorder(ByVal hwnd As Long)
    '隐藏窗口的标题框
    Dim lngWindowStyle As Long
    
    lngWindowStyle = GetWindowLong(hwnd, GWL_STYLE)
    lngWindowStyle = lngWindowStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
    
    Call SetWindowLong(hwnd, GWL_STYLE, lngWindowStyle Or WS_CHILD)
End Sub

Public Sub EmbedFindWindow(objContainer As Object, Optional owner As Object = Nothing)
'嵌入查找窗口
    Dim objSchemeItem As TSchemeItem
    Dim objFilterWindow As frmQueryFilter
         
    objSchemeItem = ReadSchemeItemWithId(mlngCurSchemeId)

    If objSchemeItem.IdKey <= 0 Then
        Err.Raise -1, "clsPacsQuery.GetFindHwnd", "未查询到对应的方案配置信息。"
    End If
    
    If objSchemeItem.FilterWindow Is Nothing Then
        Set objFilterWindow = New frmQueryFilter
        
        objFilterWindow.IsEmbed = True
        objFilterWindow.Visible = False
        
        Call objFilterWindow.ShowFilter(objSchemeItem, mdblFontSize, owner)
        Call objFilterWindow.LoadFace
        
        Call HideBorder(objFilterWindow.hwnd)
        
        Set objSchemeItem.FilterWindow = objFilterWindow
        
        Set objSchemeItem.InitValues = objFilterWindow.UpdateFindCondition().FilterValues
        
        maryScheme(objSchemeItem.Index) = objSchemeItem
    Else
        Set objFilterWindow = objSchemeItem.FilterWindow
    End If
    
    SetParent objFilterWindow.hwnd, objContainer.hwnd

    Call ShowObject(objFilterWindow)
    
    Call MoveWindow(objFilterWindow.hwnd, 0, 0, _
                    objContainer.ScaleX(objContainer.Width, vbTwips, vbPixels), _
                    objContainer.ScaleY(objContainer.Height, vbTwips, vbPixels), _
                    1)

End Sub


Public Sub EmbedConditionRestore()
On Error GoTo errHandle
    Dim objSchemeItem As TSchemeItem
    Dim objFilterWindow As frmQueryFilter
    
    objSchemeItem = ReadSchemeItemWithId(mlngCurSchemeId)

    If objSchemeItem.IdKey <= 0 Then
        Err.Raise -1, "clsPacsQuery.GetFindHwnd", "未查询到对应的方案配置信息。"
    End If
    
    If objSchemeItem.FilterWindow Is Nothing Then
        Err.Raise -1, "clsPacsQuery.EmbedFilter", "未配置嵌入式查找窗口。"
    End If
    
    Set objFilterWindow = objSchemeItem.FilterWindow
    
    Call objFilterWindow.Restore(objSchemeItem.InitValues)
    
    Exit Sub
errHandle:
End Sub


Private Sub ShowObject(objShow As Object)
'显示object对象
On Error GoTo errHandle
    objShow.Visible = True
Exit Sub
errHandle:
    ShowWindow objShow.hwnd, SW_RESTORE
End Sub


Public Sub EmbedSize(objContainer As Object)
On Error GoTo errHandle
    Dim objSchemeItem As TSchemeItem
    Dim objFilterWindow As frmQueryFilter
    
    objSchemeItem = ReadSchemeItemWithId(mlngCurSchemeId)

    If objSchemeItem.IdKey <= 0 Then
        Err.Raise -1, "clsPacsQuery.GetFindHwnd", "未查询到对应的方案配置信息。"
    End If
    
    If objSchemeItem.FilterWindow Is Nothing Then
        Err.Raise -1, "clsPacsQuery.EmbedFilter", "未配置嵌入式查找窗口。"
    End If
    
    Set objFilterWindow = objSchemeItem.FilterWindow
    
    Call MoveWindow(objFilterWindow.hwnd, _
                    objContainer.ScaleX(0, vbTwips, vbPixels), _
                    objContainer.ScaleY(0, vbTwips, vbPixels), _
                    objContainer.ScaleX(objContainer.Width, vbTwips, vbPixels), _
                    objContainer.ScaleY(objContainer.Height, vbTwips, vbPixels), _
                    1)
    Exit Sub
errHandle:

End Sub

Public Sub EmbedFree()
    Dim i As Long
    
    For i = 1 To UBound(maryScheme)
        Set maryScheme(i).SqlScheme = Nothing
        Set maryScheme(i).SqlParse = Nothing
        Set maryScheme(i).FilterValues = Nothing
        
        If Not (maryScheme(i).FilterWindow Is Nothing) Then
            ShowWindow maryScheme(i).FilterWindow.hwnd, SW_HIDE
            SetParent maryScheme(i).FilterWindow.hwnd, 0
            
            Unload maryScheme(i).FilterWindow
            
            Set maryScheme(i).FilterWindow = Nothing
        End If
    Next i
End Sub

Public Function IsMoreEmbedInput() As Boolean
'是否有更多嵌入录入
On Error GoTo errHandle
    Dim objSchemeItem As TSchemeItem
    Dim objFilterWindow As frmQueryFilter
    
    IsMoreEmbedInput = False
    
    objSchemeItem = ReadSchemeItemWithId(mlngCurSchemeId)
    
    If objSchemeItem.IdKey <= 0 Then
        Err.Raise -1, "clsPacsQuery.EmbedFilter", "未查询到对应的方案配置信息。"
    End If
    
    If objSchemeItem.FilterWindow Is Nothing Then
        Err.Raise -1, "clsPacsQuery.EmbedFilter", "未配置嵌入式查找窗口。"
    End If
    
    Set objFilterWindow = objSchemeItem.FilterWindow
    
    IsMoreEmbedInput = objFilterWindow.IsMoreInput
    
Exit Function
errHandle:
    IsMoreEmbedInput = True
    Err.Clear
End Function



Public Function EmbedFilter(Optional owner As Object, Optional ByVal blnIsMore As Boolean = False) As ADODB.Recordset
'弹出filter窗口执行方案查询
On Error GoTo errHandle
    Dim objSchemeItem As TSchemeItem
    Dim objCurSchemeItem As TSchemeItem
    Dim objFilterWindow As frmQueryFilter
    Dim objCurParse As clsSqlParse
    Dim strSql As String
    Dim strMovedSql As String
    
    
    Set EmbedFilter = Nothing
    
    objCurSchemeItem = ReadSchemeItemWithId(mlngCurSchemeId)
    
    If objCurSchemeItem.IdKey <= 0 Then
        Err.Raise -1, "clsPacsQuery.EmbedFilter", "未查询到对应的方案配置信息。"
    End If
    
    If objCurSchemeItem.FilterWindow Is Nothing Then
        Err.Raise -1, "clsPacsQuery.EmbedFilter", "未配置嵌入式查找窗口。"
    End If
    
    Set objFilterWindow = objCurSchemeItem.FilterWindow
    
    objSchemeItem = objFilterWindow.UpdateFindCondition()
    Set objSchemeItem.FilterWindow = objFilterWindow
    Set objSchemeItem.InitValues = objCurSchemeItem.InitValues
    
    
    If blnIsMore Then
        '弹出独立窗口
        Set objFilterWindow = New frmQueryFilter
        If objFilterWindow.ShowFilter(objSchemeItem, mdblFontSize, owner) = False Then Exit Function
        Unload objFilterWindow
    End If
    
    Call DoQueryState(0, "解析查询语句.")
    Set objCurParse = objSchemeItem.SqlParse
    
    If objCurParse.IsInit = False Then
        strSql = objSchemeItem.SqlScheme.Query
        Call objCurParse.init(strSql)
    End If
    
    maryScheme(objSchemeItem.Index) = objSchemeItem
    
    Set mobjCurFilterValue = objSchemeItem.FilterValues
    Set mobjSqlParseEvent = objCurParse
    
    Call DoQueryState(0, "组织查询SQL.")
    strSql = mobjSqlParseEvent.GetQuerySql(False)
    
    Dim dtEnd As Variant
    dtEnd = mobjCurFilterValue.ParData("系统.结束日期")
    
    mblnIsMoved = IIf(IsNull(dtEnd) Or IsEmpty(dtEnd), True, zlDatabase.DateMoved(dtEnd)) 'True '读取是否转储标志
    
    If mblnIsMoved Then
        '针对转储数据的查询，可能需要特殊处理
        strMovedSql = GetSqlWithMoved(strSql)
        If strMovedSql <> "" Then
            strSql = "(" & strSql & ") " & vbCrLf & " Union All " & vbCrLf & " (" & strMovedSql & ")"
        End If
    End If
    
    Set mobjCurFilterValue = Nothing
    
    Call DoQueryState(0, "获取查询结果.")
    '执行查询
    Set EmbedFilter = QuerySchemeData(strSql)
Exit Function
errHandle:
    Err.Raise -1, "clsPacsQuery.EmbedFilter", "[EmbedFilter]处理错误>>" & vbCrLf & "  方案ID为：" & mlngCurSchemeId & vbCrLf & Err.Description
    Resume
End Function

Public Function ExecuteWithFilter(ByRef dtStartdate As Date, ByRef dtEndDate As Date, _
    Optional owner As Object = Nothing) As ADODB.Recordset
'弹出filter窗口执行方案查询
On Error GoTo errHandle
    Dim objSchemeItem As TSchemeItem
    Dim objCurParse As clsSqlParse
    Dim strSql As String
    Dim strMovedSql As String
    Dim objFilterWindow As frmQueryFilter
    
    Set ExecuteWithFilter = Nothing
    
    Call DoQueryState(0, "读取方案配置.")
    objSchemeItem = ReadSchemeItemWithId(mlngCurSchemeId)

    If objSchemeItem.IdKey <= 0 Then
        Err.Raise -1, "clsPacsQuery.Execute", "未查询到对应的方案配置信息。"
    End If
    
    If dtStartdate <> Empty Then objSchemeItem.Startdate = dtStartdate
    If dtEndDate <> Empty Then objSchemeItem.EndDate = dtEndDate
    
    Call DoQueryState(0, "读取配置查询条件.")
    If objSchemeItem.FilterValues Is Nothing Then
    '载入方案默认对应的条件录入配置
        Call LoadSchemeWhereConfig(objSchemeItem)
    End If
    
    '配置查询日期范围
    If objSchemeItem.FilterValues.HasName("系统.开始日期") = False _
        And dtStartdate <> Empty Then
        objSchemeItem.FilterValues.ParData("系统.开始日期") = dtStartdate
    End If
    
    If objSchemeItem.FilterValues.HasName("系统.结束日期") = False _
        And dtEndDate <> Empty Then
        objSchemeItem.FilterValues.ParData("系统.结束日期") = dtEndDate
    End If
    
    If objSchemeItem.FilterWindow Is Nothing Then
        Set objFilterWindow = New frmQueryFilter
        objFilterWindow.IsEmbed = False
        
        Set objSchemeItem.FilterWindow = objFilterWindow
        
        maryScheme(objSchemeItem.Index) = objSchemeItem
    Else
        Set objFilterWindow = objSchemeItem.FilterWindow
    End If
    
    '退出查询
    If objFilterWindow.ShowFilter(objSchemeItem, mdblFontSize, owner) = False Then
        Exit Function
    End If
    
    '更新日期范围
    dtStartdate = objSchemeItem.Startdate
    dtEndDate = objSchemeItem.EndDate
    
    Call DoQueryState(0, "解析查询语句.")
    Set objCurParse = objSchemeItem.SqlParse
    
    If objCurParse.IsInit = False Then
        strSql = objSchemeItem.SqlScheme.Query
        Call objCurParse.init(strSql)
    End If
    
    maryScheme(objSchemeItem.Index) = objSchemeItem
    
    Set mobjCurFilterValue = objSchemeItem.FilterValues
    Set mobjSqlParseEvent = objCurParse
    
    Call DoQueryState(0, "组织查询SQL.")
    strSql = mobjSqlParseEvent.GetQuerySql(False)
    
    mblnIsMoved = zlDatabase.DateMoved(dtEndDate) 'True '读取是否转储标志
    
    If mblnIsMoved Then
        '针对转储数据的查询，可能需要特殊处理
        strMovedSql = GetSqlWithMoved(strSql)
        If strMovedSql <> "" Then
            strSql = "(" & strSql & ") " & vbCrLf & " Union All " & vbCrLf & " (" & strMovedSql & ")"
        End If
    End If
    
    Set mobjCurFilterValue = Nothing
    
    Call DoQueryState(0, "获取查询结果.")
    '执行查询
    Set ExecuteWithFilter = QuerySchemeData(strSql)
Exit Function
errHandle:
    Err.Raise -1, "clsPacsQuery.ExecuteWithFilterEx", "[ExecuteWithFilter]处理错误>>" & vbCrLf & "  方案ID为：" & mlngCurSchemeId & vbCrLf & Err.Description
    Resume
End Function


Public Function ExecuteWithLink(ByVal strLinkSql As String) As ADODB.Recordset
'连接其他sql语句进行查询，连接的sql语句，必须要返回医嘱ID，否则会产生异常
On Error GoTo errHandle
    Dim objCurSchemeItem As TSchemeItem
    Dim objCurParse As clsSqlParse
    Dim strSql As String
    Dim strMovedSql As String
    
    Set ExecuteWithLink = Nothing
    
    Call DoQueryState(0, "读取方案配置.")
    objCurSchemeItem = ReadSchemeItemWithId(mlngCurSchemeId)
    
    If objCurSchemeItem.IdKey <= 0 Then
        Err.Raise -1, "clsPacsQuery.Execute", "未查询到对应的方案配置信息。"
    End If
    
    Call DoQueryState(0, "解析查询语句.")
    Set objCurParse = objCurSchemeItem.SqlParse
    
    If objCurParse.IsInit = False Then
        strSql = objCurSchemeItem.SqlScheme.Query
        Call objCurParse.init(strSql)
    End If
    
    Set mobjSqlParseEvent = objCurParse
    
    Call DoQueryState(0, "组织查询SQL.")
    strSql = mobjSqlParseEvent.GetQuerySql(False)
    
    mblnIsMoved = True  '如果是按收藏方式查询，则mblnismoved需要设置为true
'    If mblnIsMoved Then
'        strMovedSql = GetSqlWithMoved(strSql)
'        If strMovedSql <> "" Then
'            strSql = "(" & strSql & ") " & vbCrLf & _
'                    " Union All " & vbCrLf & _
'                    " (" & strMovedSql & ")"
'        End If
'    End If
    
    strSql = "Select QA.* From (" & strSql & ") QA, (" & strLinkSql & ") QB Where QA.医嘱ID=QB.医嘱ID"
    
    If mblnIsMoved Then
        '针对转储数据的查询，可能需要特殊处理
        strMovedSql = GetSqlWithMoved(strSql)
        If strMovedSql <> "" Then
            strSql = "(" & strSql & ") " & vbCrLf & " Union All " & vbCrLf & " (" & strMovedSql & ")"
        End If
    End If
    
    Call DoQueryState(0, "获取查询结果.")
    Set ExecuteWithLink = QuerySchemeData(strSql)
Exit Function
errHandle:
    Err.Raise -1, "clsPacsQuery.ExecuteWithLink", "[ExecuteWithLink]处理错误>>" & vbCrLf & "  方案ID为：" & mlngCurSchemeId & vbCrLf & Err.Description
    Resume
End Function


Public Function ExecuteWithAttach(ByVal strAttach As String, ByVal attachData As Variant) As ADODB.Recordset
'使用指定的条件数据执行方案
On Error GoTo errHandle
    Dim objCurSchemeItem As TSchemeItem
    Dim objCurParse As clsSqlParse
    Dim strSql As String
    Dim strMovedSql As String
    
    Set ExecuteWithAttach = Nothing
    
    Call DoQueryState(0, "读取方案配置.")
    objCurSchemeItem = ReadSchemeItemWithId(mlngCurSchemeId)
    
    If objCurSchemeItem.IdKey <= 0 Then
        Err.Raise -1, "clsPacsQuery.Execute", "未查询到对应的方案配置信息。"
    End If
    
    Call DoQueryState(0, "解析查询语句.")
    Set objCurParse = objCurSchemeItem.SqlParse
    
    If objCurParse.IsInit = False Then
        strSql = objCurSchemeItem.SqlScheme.Query
        Call objCurParse.init(strSql)
    End If
    
    Set mobjSqlParseEvent = objCurParse
    
    Call DoQueryState(0, "组织查询SQL.")
    strSql = mobjSqlParseEvent.GetAttachSql(strAttach, attachData)
    
    '转储sql语句处理
    If mblnIsMoved Then
        '针对转储数据的查询，可能需要特殊处理
        strMovedSql = GetSqlWithMoved(strSql)
        If strMovedSql <> "" Then
            strSql = "(" & strSql & ") " & vbCrLf & " Union All " & vbCrLf & " (" & strMovedSql & ")"
        End If
    End If
    
    Call DoQueryState(0, "获取查询结果.")
    Set ExecuteWithAttach = QuerySchemeData(strSql)
Exit Function
errHandle:
    Err.Raise -1, "clsPacsQuery.ExecuteWithAttach", "[ExecuteWithAttach]处理错误>>" & vbCrLf & "  方案ID为：" & mlngCurSchemeId & vbCrLf & Err.Description
    Resume
End Function


Public Function DataConvert(rsData As ADODB.Recordset, Optional ByVal lngSchemeId As Long = 0) As ADODB.Recordset
'数据转换
    Dim i As Long
    Dim lngStartIndex As Long
    
    Dim objSqlScheme As clsSqlScheme
    Dim rsConvert As ADODB.Recordset
    Dim objShowCfg As clsScShowCfg
    
    Dim lngConvertCfgCount As Long
    Dim strConvertColName As String
    Dim strConvertDataCfg As String
    Dim strFieldValue As String
    
    Dim aryConvertCols() As String
    Dim aryConvertDatas() As String
    
    Set DataConvert = Nothing
    If rsData Is Nothing Then Exit Function
    
    If rsData.RecordCount <= 0 Then
        Set DataConvert = CopyRecordSet(rsData)
        Exit Function
    End If
    
    Set objSqlScheme = GetSqlScheme(lngSchemeId)
    lngConvertCfgCount = objSqlScheme.RefreshConvertColCount(rsData)
    
    '没有转换数据则直接返回
    If lngConvertCfgCount <= 0 Then
        Set DataConvert = CopyRecordSet(rsData)
        Exit Function
    End If
    
    Set rsConvert = CopyRecordSet(rsData)
    
    aryConvertCols = objSqlScheme.ConvertColNames
    aryConvertDatas = objSqlScheme.ConvertColDatas
 
    While Not rsConvert.EOF
        For i = 1 To lngConvertCfgCount
            strConvertColName = aryConvertCols(i)
            strConvertDataCfg = aryConvertDatas(i)
            
            strFieldValue = ";" & NVL(rsConvert.Fields(strConvertColName).Value) & "-"
            
            lngStartIndex = InStr(strConvertDataCfg, strFieldValue)
            If lngStartIndex > 0 Then
                strFieldValue = Mid(strConvertDataCfg & ";", lngStartIndex + Len(strFieldValue), 100)
                rsConvert.Fields(strConvertColName).Value = Mid(strFieldValue, 1, InStr(strFieldValue, ";") - 1)
            End If
        Next i
    
        rsConvert.MoveNext
    Wend
    
    '定位到第一行数据
    rsConvert.MoveFirst
    
    Set DataConvert = rsConvert

End Function


Public Function GetSqlScheme(Optional ByVal lngSchemeId As Long = 0) As clsSqlScheme
'获取指定方案配置
    
    Dim SchemeItem As TSchemeItem
    Dim lngCurSchemeId As Long
    
    lngCurSchemeId = lngSchemeId
    If lngSchemeId <= 0 Then lngCurSchemeId = mlngCurSchemeId
    
    SchemeItem = ReadSchemeItemWithId(lngCurSchemeId)
    
    Set GetSqlScheme = SchemeItem.SqlScheme
End Function

'***********************************************************************************************************

Private Function ReadSchemeWithIndex(ByVal lngSchemeIndex As Long) As clsSqlScheme
    Dim objSchemeItem As TSchemeItem
    
    objSchemeItem = maryScheme(lngSchemeIndex)
    
    If objSchemeItem.Index <= 0 Then
        Set ReadSchemeWithIndex = Nothing
        Exit Function
    End If
    
    If objSchemeItem.IsLoaded = False Then
        Call ParseScheme(objSchemeItem, True)
    End If
    
    Set ReadSchemeWithIndex = objSchemeItem.SqlScheme
End Function

Private Function ReadNullSchemeItem() As TSchemeItem
    ReadNullSchemeItem.IdKey = -1
    ReadNullSchemeItem.BaseInfo.SchemeId = -1
    ReadNullSchemeItem.BaseInfo.Name = ""
    ReadNullSchemeItem.BaseInfo.Ver = -1
    ReadNullSchemeItem.BaseInfo.Memo = ""
    ReadNullSchemeItem.BaseInfo.IsDefault = False
    ReadNullSchemeItem.BaseInfo.IsOften = False
    
    ReadNullSchemeItem.IsLoaded = False
    
    Set ReadNullSchemeItem.SqlScheme = Nothing
    Set ReadNullSchemeItem.SqlParse = Nothing
    
    ReadNullSchemeItem.Index = -1
End Function

Private Function ReadSchemeItemWithId(ByVal lngSchemeId As Long) As TSchemeItem
'判断方案是否存在
'返回方案对应所在数组中的索引
On Error GoTo errHandle
    Dim i As Long
    Dim objSchemeItem As TSchemeItem
    
    objSchemeItem = ReadNullSchemeItem
    
    '查找对应的方案对象
    For i = 1 To SchemeCount
        If maryScheme(i).IdKey = lngSchemeId Then
            objSchemeItem = maryScheme(i)
            Exit For
        End If
    Next i
    
    If objSchemeItem.IdKey <= 0 Then
    '方案不存在的处理
        objSchemeItem = ReadNullSchemeItem()
        objSchemeItem.IdKey = lngSchemeId
        
        Call WriteSchemeItemArray(objSchemeItem)
    End If
    
    If objSchemeItem.IsLoaded = False Then
        '如果方案读取失败，则返回nothing
        If ParseScheme(objSchemeItem, True) = False Then
            objSchemeItem = ReadNullSchemeItem
        Else
            maryScheme(objSchemeItem.Index) = objSchemeItem
        End If
    End If
    
    ReadSchemeItemWithId = objSchemeItem
Exit Function
errHandle:
    Err.Raise -1, "clsPacsQuery.ReadSchemeItemWithId", "[ReadSchemeItemWithId]处理错误>>" & vbCrLf & "  方案ID为：" & lngSchemeId & vbCrLf & Err.Description
    Resume
End Function


Private Function WriteSchemeItemArray(ByRef objSchemeItem As TSchemeItem) As Long
'写入方案配置项
'返回方案对应所在数组中的索引
On Error GoTo errHandle
    Dim lngUbound As Long
    Dim SqlParse As clsSqlParse
    
    WriteSchemeItemArray = 0
    
    lngUbound = UBound(maryScheme) + 1
    
    objSchemeItem.Index = lngUbound
    
    Set objSchemeItem.SqlScheme = New clsSqlScheme
'    Set sqlParse = New clsSqlParse
    
'    sqlParse.NullParameterHint = False '默认不提示
'    sqlParse.NullWhereHint = False '默认不提示
    
    Set objSchemeItem.SqlParse = New clsSqlParse
    
    '保存方案到数组
    ReDim Preserve maryScheme(lngUbound)
    maryScheme(lngUbound) = objSchemeItem
    
    WriteSchemeItemArray = lngUbound
Exit Function
errHandle:
    Err.Raise -1, "clsPacsQuery.WriteSchemeItemArray", "[WriteSchemeItemArray]处理错误>>" & vbCrLf & "  方案ID为：" & objSchemeItem.IdKey & vbCrLf & Err.Description
    Resume
End Function


Private Function ParseScheme(ByRef objSchemeItem As TSchemeItem, ByVal blnIsForceRead As Boolean) As Boolean
'载入方案配置
'返回方案对应所在数组中的索引
On Error GoTo errHandle
    Dim strSchemeXml As String
    Dim strCacheSchemeFile As String
    Dim blnRead As Boolean
    Dim objSqlScheme As clsSqlScheme
    Dim strSchemeDes As String
    
    
    ParseScheme = False
    
    
    If objSchemeItem.IsLoaded = True And blnIsForceRead = False Then Exit Function
    
    Set objSqlScheme = objSchemeItem.SqlScheme
        
    '判断本地文件是否缓存
    strCacheSchemeFile = Replace(gstrCachePath & "\" & objSchemeItem.IdKey & ".XML", "\\", "\")
    

    blnRead = False
    
    If Dir(strCacheSchemeFile) <> "" Then
        Call objSqlScheme.OpenFile(strCacheSchemeFile)
        If objSqlScheme.SchemeId <= 0 Then objSqlScheme.SchemeId = objSchemeItem.IdKey
        
        If objSqlScheme.Ver > -2 And objSqlScheme.Ver < objSchemeItem.BaseInfo.Ver Then blnRead = True
    Else
        blnRead = True
    End If
    
    If blnRead Then
        strSchemeXml = ReadSchemeXml(objSchemeItem.IdKey, strSchemeDes)
        
        If Len(strSchemeXml) > 0 Then
            Call objSqlScheme.OpenScheme(strSchemeXml)
            objSqlScheme.Descript = strSchemeDes
            
            If objSqlScheme.SchemeId <= 0 Then objSqlScheme.SchemeId = objSchemeItem.IdKey
            
            Call objSqlScheme.SaveAsScheme(strCacheSchemeFile)
            
            '判断是否需要更新zlBakTables信息
            Call UpdateBakTablesCache
        End If
    End If
    
    objSchemeItem.IsLoaded = True
    
    
    ParseScheme = True
Exit Function
errHandle:
    Err.Raise -1, "clsPacsQuery.ParseScheme", "[ParseScheme]处理错误>>" & vbCrLf & "  方案ID为：" & objSchemeItem.IdKey & vbCrLf & Err.Description
    Resume
End Function


Private Function QuerySchemeData(ByVal strSql As String) As ADODB.Recordset
'查询方案数据
On Error GoTo errHandle
    Set QuerySchemeData = Nothing
     
    
    If Len(strSql) <= 0 Then
'        Err.Raise -1, "clsPacsQuery.QuerySchemeData", "解析后的SQL语句为空。"
        Exit Function
    End If
    
    Set QuerySchemeData = ExecuteCore(strSql, "执行查询", mobjSqlParseEvent.ParValues)
    
Exit Function
errHandle:
    Err.Raise -1, "clsPacsQuery.QuerySchemeData", "[QuerySchemeData]处理错误>>" & vbCrLf & "  查询语句为：" & strSql & vbCrLf & Err.Description
    Resume
End Function



Private Sub Class_Initialize()
    ReDim maryScheme(0)
    ReDim mstrBakTables(0)
    
    mstrStationName = UCase(StationName)
End Sub

Private Sub Class_Terminate()
    Dim i As Long
    
    Set mobjSqlParseEvent = Nothing
    
    For i = 1 To UBound(maryScheme)
        Set maryScheme(i).SqlScheme = Nothing
        Set maryScheme(i).SqlParse = Nothing
        Set maryScheme(i).FilterValues = Nothing
        
        If Not (maryScheme(i).FilterWindow Is Nothing) Then
'            ShowWindow maryScheme(i).FilterWindow.hWnd, SW_HIDE
'            SetParent maryScheme(i).FilterWindow.hWnd, 0
            
            Unload maryScheme(i).FilterWindow
            
            Set maryScheme(i).FilterWindow = Nothing
        End If
    Next i
End Sub

Private Sub DoQueryState(ByVal lngCurState As Long, ByVal strStateMsg As String)
On Error GoTo errHandle
    '输出查询状态消息
    RaiseEvent OnQueryState(lngCurState, strStateMsg)
Exit Sub
errHandle:
    Debug.Print "clsPacsQuery.DoQueryState Err:" & Err.Description
End Sub

Private Sub DoGetParameterValue(ByVal strParName As String, ByRef Value As Variant)
On Error GoTo errHandle
    '获取参数
    RaiseEvent OnGetParameterValue(strParName, Value)
Exit Sub
errHandle:
    Err.Raise -1, "clsPacsQuery.DoGetParameterValue", "[DoGetParameterValue]处理错误>>" & vbCrLf & "  参数名称：" & strParName & vbCrLf & Err.Description
    Resume
End Sub

Private Sub mobjSqlParseEvent_OnGetParameterValue(ByVal strParName As String, ByRef Value As Variant)
'获取参数值
On Error GoTo errHandle
    If Not (mobjCurFilterValue Is Nothing) Then
        Value = mobjCurFilterValue.ParData(strParName)
        If IsEmpty(Value) = False And IsNull(Value) = False And Value <> "" Then Exit Sub
    End If
    
    Call DoGetParameterValue(strParName, Value)
Exit Sub
errHandle:
    Err.Raise -1, "clsPacsQuery.OnGetParameterValue", "[OnGetParameterValue]处理错误>>" & vbCrLf & "  参数名称：" & strParName & vbCrLf & Err.Description
    Resume
End Sub
