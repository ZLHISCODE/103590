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

Private mobjRegExp As Object

Private mobjCurFilterValue As clsSqlFilterValue '用于参数取值
Private mblnIsNeedDateRange As Boolean
'Private mblnIsIgnoreSysPar As Boolean

Private mobjSchemeCfg As frmSchemeCfg
Attribute mobjSchemeCfg.VB_VarHelpID = -1

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

Property Get objSqlParse() As clsSqlParse
    Set objSqlParse = mobjSqlParseEvent
End Property


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

Public Sub ShowSchemeCfg(ByVal lngModuleNo As Long, ByVal strSysParDefines As String, ByVal strBaseParDefines As String, ByVal bytFontSize As Byte, owner As Object)
'显示方案配置...
'lngModuleNo:模块号
'strSysParDefines:系统参数定义，多个参数使用“逗号分割”，如“[系统.系统号],[系统.模块号]...”
'bytFontSize:0-小字体；1-中字体；2-大字体
'如果有调整返回true，没有调整返回false
On Error GoTo errHandle
    
    If mobjSchemeCfg Is Nothing Then
        Set mobjSchemeCfg = New frmSchemeCfg
    End If
    
    Call mobjSchemeCfg.ShowMe(lngModuleNo, strSysParDefines, strBaseParDefines, bytFontSize, owner, mlngCurSchemeId)
Exit Sub
errHandle:
    Set mobjSchemeCfg = Nothing
    Err.Raise -1, "clsPacsQuery.ShowSchemeCfg", "方案配置调用失败:" & Err.Description
End Sub

Public Sub RefreshCfgFontSize(ByVal bytFontSize As Byte)
'调整方案配置界面的字体
'bytFontSize:0-小字体；1-中字体；2-大字体
    Dim objSchemeItem As TSchemeItem
    Dim objFilterWindow As frmQueryFilter
    
    On Error GoTo errHandle
    
    If Not mobjSchemeCfg Is Nothing Then mobjSchemeCfg.SetFontSize bytFontSize

         
    objSchemeItem = ReadSchemeItemWithId(mlngCurSchemeId)

    If objSchemeItem.IdKey <= 0 Then Exit Sub
    
    If Not objSchemeItem.FilterWindow Is Nothing Then
        Set objFilterWindow = objSchemeItem.FilterWindow
        objFilterWindow.RefreshFontSize bytFontSize
    End If
    
    Exit Sub
errHandle:
'    Err.Raise -1, "clsPacsQuery.RefreshCfgFontSize", "方案配置界面字体设置失败:" & Err.Description
End Sub

Public Function ShowUserScheme(ByVal lngModuleNo As Long, ByVal lngDeptId As Long, owner As Object) As Boolean
'显示用户常用查询方案配置...
'如果有调整返回true，没有调整返回false
On Error GoTo errHandle
    Dim objQueryCfg As New frmUserQueryReleation
    
    If lngDeptId <= 0 Then
        Err.Raise -1, "clsPacsQuery.ShowUserScheme", "当前科室不允许进行调整。" & Err.Description
        Exit Function
    End If
    
    ShowUserScheme = objQueryCfg.ShowUserScheme(owner, lngModuleNo, mlngUserId, lngDeptId)
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

Public Sub InitUserInfo()
'功能：获取登陆用户信息
    Dim rsTmp As New ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
  
    If Not rsTmp.EOF Then
        glngUserId = Val(NVL(rsTmp!Id))
        gstrUserAccount = NVL(rsTmp!用户名)
        gstrUserName = NVL(rsTmp!姓名)
    End If
    
    Exit Sub
errH:

End Sub

Public Sub SyncSysPar(ByVal strDeptId As String)
    Call InitUserInfo
    
    gstrDeptId = strDeptId
End Sub

Public Sub LoadQueryScheme(ByVal lngSysNo As Long, ByVal lngModuleNo As Long, _
    ByVal lngDeptId As Long, ByVal lngUserId As Long)
'载入查询方案到sqlScheme和sqlParse对象中
On Error GoTo errHandle
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strDeptIds As String
    
    ReDim maryScheme(0)
    
    'lngDeptId如果为0表示加载该人员对应所有部门的查询方案
    strDeptIds = ""
    If lngDeptId = 0 Then
        strSql = "SELECT 人员ID,WM_CONCAT(部门ID) as 关联部门  FROM 部门人员  where 人员ID=[1] GROUP BY 人员ID"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询用户对应科室", lngUserId)
        If rsData.RecordCount > 0 Then
            strDeptIds = ",0," & NVL(rsData!关联部门) & ","
        End If
    Else
        strDeptIds = ",0," & lngDeptId & ","
    End If
    
    
    strSql = "Select  A.ID, A.方案名称, A.方案说明, A.所属科室,A.版本, Nvl(B.是否默认,A.是否默认) As 是否默认, Nvl(B.是否常用,A.是否常用) As 是否常用, B.默认加载站点, B.所属站点,  B.用户ID " & _
                " From 影像查询方案 A, 影像查询关联 B " & _
                " Where A.ID=B.查询方案ID(+) And A.使用状态=1 And A.所属模块=[1] And B.用户ID(+)=[2] And instr([3], ',' || nvl(A.所属科室, 0) || ',') >0 order by 是否默认 Desc, 是否常用 Desc, 方案序号"
    

    Set rsData = ExecuteSql(strSql, "载入所有方案", lngModuleNo, mlngUserId, strDeptIds)
    
    If rsData Is Nothing Then Exit Sub
    If rsData.RecordCount <= 0 Then Exit Sub
    
    glngSysNo = lngSysNo
    glngModuleNo = lngModuleNo
    gstrDeptId = lngDeptId
    
    Call InitUserInfo
    
    rsData.Filter = "用户ID=" & mlngUserId
    
    '用户ID小于0，说明没有根据用户过滤出对应的方案配置
    '如果大于0，说明用户已经重新配置了一套默认设置
    If rsData.RecordCount <= 0 Then
        rsData.Filter = ""
    End If
            
    '对方案进行筛选
    Call WriteScheme(rsData, True)
    

    '针对后续新增的，还没有进行用户常用查询配置的方案进行读取（用户配置常用方案关联后，后续新增的方案不会自动关联到用户）
    'rsData.Filter不为空表示已根据用户配置情况过滤
    If rsData.Filter <> "" And rsData.Filter <> 0 Then
        rsData.Filter = "用户ID = NULL"
        If rsData.RecordCount > 0 Then
            Call WriteScheme(rsData, False)
        End If

    End If
    

    If UBound(maryScheme) > 0 Then
        Call ChangeCurScheme(maryScheme(1).BaseInfo.SchemeId)
    End If
    
Exit Sub
errHandle:
    Err.Raise -1, "clsPacsQuery.LoadQueryScheme", "[LoadQueryScheme]处理错误>>" & vbCrLf & Err.Description
    Resume
End Sub

Private Sub WriteScheme(rsData As ADODB.Recordset, Optional ByVal blnIsReadPerson As Boolean = True)
'rsData方案关联信息
'blnIsReadPerson是否读取私有设置，注意rsdata中可能是已根据用户查找到的用户关联配置，也有可能直接是默认的方案配置数据
    Dim i As Long
    Dim objSchemeItem As TSchemeItem
    Dim objTmpScheme As TSchemeItem
    Dim strScName As String
    Dim strSchemeStationName As String
    Dim strDefaultLoadStation As String
    Dim blnIsSetStationDef As Boolean
    Dim blnIsSetUserDef As Boolean
    Dim blnIsSetDeptDef As Boolean
    Dim blnIsSetModuleDef As Boolean
    
    Dim blnIsDef As Boolean
    Dim blnIsContinue As Boolean
    
    blnIsSetStationDef = False '是否设置站点的默认加载方案
    blnIsSetUserDef = False
    blnIsSetDeptDef = False
    blnIsSetModuleDef = False
    
    While Not rsData.EOF
        strScName = NVL(rsData!方案名称)
        strSchemeStationName = NVL(rsData!所属站点)
        strDefaultLoadStation = NVL(rsData!默认加载站点)
        
        If strSchemeStationName = "" Or UCase(strSchemeStationName) = mstrStationName Then
            
            objSchemeItem.IdKey = Val(NVL(rsData!Id))
            objSchemeItem.BaseInfo.SchemeId = objSchemeItem.IdKey
            objSchemeItem.BaseInfo.Ver = Val(NVL(rsData!版本))
            objSchemeItem.BaseInfo.Name = NVL(rsData!方案名称)
            objSchemeItem.BaseInfo.Memo = NVL(rsData!方案说明)
            objSchemeItem.BaseInfo.IsDefault = False
            objSchemeItem.BaseInfo.IsOften = False
            
            'TODO:需要根据科室设置获取对应的默认方案
            If blnIsReadPerson Then
                objSchemeItem.BaseInfo.IsOften = IIf(Val(NVL(rsData!是否常用)) <> 0, True, False)
                
                blnIsContinue = False
                blnIsDef = False
                
                '判断是否设置站点默认
                If Trim(strDefaultLoadStation) <> "" Then
                    If InStr(strDefaultLoadStation, mstrStationName) > 0 Then
                        If blnIsSetStationDef = False Then
                            '只有读取第一次才允许设置
                            blnIsDef = True
                            blnIsSetStationDef = True
                            
                            blnIsSetUserDef = True
                            blnIsSetDeptDef = True
                            blnIsSetModuleDef = True
                        Else
                            blnIsDef = False
                        End If
                    End If
                End If
                
                '如果没有设置站点默认，则判断是否对用户默认，科室默认进行了设置
                If blnIsSetStationDef = False Then
                    blnIsDef = IIf(Val(NVL(rsData!是否默认)) <> 0, True, False)
                    
                    If blnIsDef Then
                        If blnIsSetModuleDef = False Then
                            blnIsSetModuleDef = True
                        Else
                            blnIsDef = False
                        End If
                        
                        If Val(NVL(rsData!所属科室)) <> 0 Then
                            If blnIsSetDeptDef = False Then
                                blnIsDef = True
                                
                                blnIsSetDeptDef = True
                                blnIsSetModuleDef = True
                            Else
                                blnIsDef = False
                            End If
                        End If
                        
                        If Val(NVL(rsData!用户ID)) <> 0 Then
                            If blnIsSetUserDef = False Then
                                blnIsDef = True
                                
                                blnIsSetUserDef = True
                                blnIsSetDeptDef = True
                                blnIsSetModuleDef = True
                            Else
                                blnIsDef = False
                            End If
                        End If
                    End If
                End If
                
                '允许站点默认设置
                If blnIsSetStationDef And blnIsDef Then blnIsContinue = True
                '允许用户默认设置
                If blnIsSetUserDef And blnIsDef Then blnIsContinue = True
                '允许科室默认设置
                If blnIsSetDeptDef And blnIsDef Then blnIsContinue = True
                '允许模块默认设置
                If blnIsSetModuleDef And blnIsDef Then blnIsContinue = True
                
                '如果已经配置了站点默认，则不读取后续的默认配置
                objSchemeItem.BaseInfo.IsDefault = blnIsDef
                    
                If blnIsContinue Then
             
                    For i = 1 To UBound(maryScheme)
                        maryScheme(i).BaseInfo.IsDefault = False
                    Next
                    
                    objSchemeItem.BaseInfo.IsDefault = True
          
                    
                    If UBound(maryScheme) >= 1 Then
                    
                        objTmpScheme = maryScheme(1)
                        
                        objSchemeItem.Index = 1
                        maryScheme(1) = objSchemeItem
                        
                        objSchemeItem = objTmpScheme
                    End If
                End If
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
On Error GoTo errH
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
    Exit Function
errH:
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

Public Function GetSqlWithMoved(ByVal strSource As String) As String
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
    
    mblnIsNeedDateRange = True
'    mblnIsIgnoreSysPar = False
    
    Set mobjCurFilterValue = objSchemeItem.FilterValues
    
    Set mobjSqlParseEvent = objCurParse
    
    Call DoQueryState(0, "组织查询SQL.")
    strSql = mobjSqlParseEvent.GetQuerySql(blnIsRefresh)
    
    mblnIsMoved = zlDatabase.DateMoved(dtStartdate)
    
    Set mobjCurFilterValue = Nothing
    
    Call DoQueryState(0, "获取查询结果.")
    Set Execute = DBQueryAnalyzer(objSchemeItem.SqlParse.SqlStruct, strSql, objSchemeItem.SqlScheme.HistoryDBHint, mblnIsMoved)
Exit Function
errHandle:
    Call Clipboard.SetText(Err.Description)
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

Public Function EmbedFindWindow(objContainer As Object, Optional owner As Object = Nothing) As Boolean
'嵌入查找窗口
    Dim objSchemeItem As TSchemeItem
    Dim objFilterWindow As frmQueryFilter
         
    EmbedFindWindow = False
    
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
        
        EmbedFindWindow = True
    Else
        Set objFilterWindow = objSchemeItem.FilterWindow
    End If
    
    SetParent objFilterWindow.hwnd, objContainer.hwnd

    Call ShowObject(objFilterWindow)
    
    Call MoveWindow(objFilterWindow.hwnd, 0, 0, _
                    objContainer.ScaleX(objContainer.Width, vbTwips, vbPixels), _
                    objContainer.ScaleY(objContainer.Height, vbTwips, vbPixels), _
                    1)

End Function


Public Sub EmbedConditionRestore(Optional ByVal isValidTime As Boolean = False, Optional ByVal dtstartTime As Date = Empty, Optional ByVal dtendTime As Date = Empty)
On Error GoTo errHandle
    Dim objSchemeItem As TSchemeItem
    Dim objFilterWindow As frmQueryFilter
    
    objSchemeItem = ReadSchemeItemWithId(mlngCurSchemeId)

    If objSchemeItem.IdKey <= 0 Then
        Err.Raise -1, "clsPacsQuery.GetFindHwnd", "未查询到对应的方案配置信息。"
    End If
    
    If objSchemeItem.FilterWindow Is Nothing Then
        Exit Sub
'        Err.Raise -1, "clsPacsQuery.EmbedFilter", "未配置嵌入式查找窗口。"
    End If
    
    Set objFilterWindow = objSchemeItem.FilterWindow
    
    Call objFilterWindow.Restore(objSchemeItem.InitValues, isValidTime, dtstartTime, dtendTime)
    
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
    Dim objPopupFilterWindow As frmQueryFilter
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
    
    If objSchemeItem.IdKey <= 0 Then Exit Function
    
    Set objSchemeItem.FilterWindow = objFilterWindow
    Set objSchemeItem.InitValues = objCurSchemeItem.InitValues
    
    
    If blnIsMore Then
        '弹出独立窗口
        Set objPopupFilterWindow = New frmQueryFilter
        If objPopupFilterWindow.ShowFilter(objSchemeItem, mdblFontSize, owner) = False Then Exit Function
        
        Call objFilterWindow.ReadUserInputConfig(objSchemeItem.FilterValues)
        
        Unload objPopupFilterWindow
    End If
    
    Call DoQueryState(0, "解析查询语句.")
    Set objCurParse = objSchemeItem.SqlParse
    
    If objCurParse.IsInit = False Then
        strSql = objSchemeItem.SqlScheme.Query
        Call objCurParse.init(strSql)
    End If
    
    maryScheme(objSchemeItem.Index) = objSchemeItem
    
    mblnIsNeedDateRange = True
'    mblnIsIgnoreSysPar = False
    
    Set mobjCurFilterValue = objSchemeItem.FilterValues
    Set mobjSqlParseEvent = objCurParse
    
    Call DoQueryState(0, "组织查询SQL.")
    strSql = mobjSqlParseEvent.GetQuerySql(False)
    
    Dim dtStart As Variant
    dtStart = mobjCurFilterValue.ParData("系统.开始日期")
 
    mblnIsMoved = IIf(IsNull(dtStart) Or IsEmpty(dtStart) Or dtStart = 0, True, zlDatabase.DateMoved(dtStart)) 'True '读取是否转储标志
 
    Set mobjCurFilterValue = Nothing
    
    Call DoQueryState(0, "获取查询结果.")
    '执行查询
    Set EmbedFilter = DBQueryAnalyzer(objSchemeItem.SqlParse.SqlStruct, strSql, objSchemeItem.SqlScheme.HistoryDBHint, mblnIsMoved)
    
    
Exit Function
errHandle:
    Call Clipboard.SetText(Err.Description)
    Err.Raise -1, "clsPacsQuery.EmbedFilter", "[EmbedFilter]处理错误>>" & vbCrLf & "  方案ID为：" & mlngCurSchemeId & vbCrLf & Err.Description
    Resume
End Function


Private Function DBQueryAnalyzer(objSqlStruct As clsSqlStruct, ByVal strSql As String, ByVal strHistoryDBHint As String, _
    ByVal blnIsMoved As Boolean) As ADODB.Recordset
    Dim strMovedSql As String
    Dim lngSelectIndex As Long
    Dim lngRuleIndex As Long
    Dim lngRuleEnd As Long
    Dim blnHasRule As Boolean
    Dim strRule As String
    Dim rsOnLine As ADODB.Recordset
    Dim rsHistory As ADODB.Recordset
    
    If blnIsMoved Then
        strMovedSql = GetSqlWithMoved(strSql)
        If strMovedSql <> "" Then
            If Trim(strHistoryDBHint) <> "" Then
                lngSelectIndex = InStr(strMovedSql, "SELECT")
                lngRuleIndex = InStr(strMovedSql, "/*+")
                 
                blnHasRule = False
                
                If lngRuleIndex > 0 Then
                    If Trim(Mid(strMovedSql, lngSelectIndex + 6, lngRuleIndex - (lngSelectIndex + 6))) = "" Then
                        blnHasRule = True
                    End If
                End If
                
                If blnHasRule Then
                    '存在规则的处理
                    lngRuleEnd = InStr(strMovedSql, "*/")
                    
                    If lngRuleEnd > 0 Then
                        strRule = Mid(strMovedSql, lngRuleIndex, lngRuleEnd - (lngRuleIndex) + 2)
                        strMovedSql = Replace(strMovedSql, strRule, strHistoryDBHint, 1, 1)
                    End If
                Else
                    strMovedSql = Replace(strMovedSql, "SELECT", "SELECT " & strHistoryDBHint, 1, 1)
                End If
                
                Set rsOnLine = QuerySchemeData(strSql)
                Set rsHistory = QuerySchemeData(strMovedSql)
                
                
                Set DBQueryAnalyzer = ConnectRecordset(rsOnLine, rsHistory)
                
            Else
                                
                strSql = "(" & strSql & ") " & vbCrLf & " Union All " & vbCrLf & " (" & strMovedSql & ")"
                
                '处理union all中分别包含orderby 报错问题
                If InStr(objSqlStruct.Others, "ORDER BY") > 0 Then
                    strSql = Replace(strSql, Trim(objSqlStruct.Others), "")
                    
                    strSql = strSql & RegReplace(objSqlStruct.Others, "\S[.]", "")
                End If
                
                Set DBQueryAnalyzer = QuerySchemeData(strSql)
            End If
            
            Exit Function
        End If
        
        Set DBQueryAnalyzer = QuerySchemeData(strSql)
    Else
        Set DBQueryAnalyzer = QuerySchemeData(strSql)
    End If
    
End Function

Private Function RegReplace(ByVal strSource As String, ByVal strFind As String, ByVal strReplace As String) As String
'使用正则表达式进行替换
    
    If mobjRegExp Is Nothing Then Set mobjRegExp = CreateObject("VBScript.RegExp")
    mobjRegExp.Pattern = strFind
    
    mobjRegExp.Global = True
    mobjRegExp.IgnoreCase = True
    mobjRegExp.MultiLine = True
    
    RegReplace = mobjRegExp.Replace(strSource, strReplace)
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
    
    mblnIsNeedDateRange = True
'    mblnIsIgnoreSysPar = False
    
    Set mobjCurFilterValue = objSchemeItem.FilterValues
    Set mobjSqlParseEvent = objCurParse
    
    Call DoQueryState(0, "组织查询SQL.")
    strSql = mobjSqlParseEvent.GetQuerySql(False)
    
    mblnIsMoved = IIf(IsNull(dtStartdate) Or IsEmpty(dtStartdate) Or dtStartdate = 0, True, zlDatabase.DateMoved(dtStartdate)) 'zlDatabase.DateMoved(dtStartdate) 'True '读取是否转储标志
    
    Set mobjCurFilterValue = Nothing
    
    Call DoQueryState(0, "获取查询结果.")
    '执行查询
    Set ExecuteWithFilter = DBQueryAnalyzer(objSchemeItem.SqlParse.SqlStruct, strSql, objSchemeItem.SqlScheme.HistoryDBHint, mblnIsMoved)
Exit Function
errHandle:
    Call Clipboard.SetText(Err.Description)
    Err.Raise -1, "clsPacsQuery.ExecuteWithFilterEx", "[ExecuteWithFilter]处理错误>>" & vbCrLf & "  方案ID为：" & mlngCurSchemeId & vbCrLf & Err.Description
    Resume
End Function


Public Function ExecuteWithLink(ByVal strLinkSql As String, _
    Optional ByVal blnHasPar As Boolean = False, Optional ByVal lngQueryMovedState As Long = 0) As ADODB.Recordset
'连接其他sql语句进行查询，连接的sql语句，必须要返回医嘱ID，否则会产生异常
'lngQueryMovedState: 0-直接使用mblnIsMoved的值，1-不使用转储查询，2-使用转储查询
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
    
    
'    mblnIsMoved = blnIsQueryMovedData 'True  '如果是按收藏方式查询，则mblnismoved需要设置为true
'    If mblnIsMoved Then
'        strMovedSql = GetSqlWithMoved(strSql)
'        If strMovedSql <> "" Then
'            strSql = "(" & strSql & ") " & vbCrLf & _
'                    " Union All " & vbCrLf & _
'                    " (" & strMovedSql & ")"
'        End If
'    End If
    
    Call DoQueryState(0, "解析查询语句.")
    
    If blnHasPar Then
        Set objCurParse = New clsSqlParse
        
        strSql = "Select QA.* From (" & objCurSchemeItem.SqlScheme.Query & ") QA, (" & strLinkSql & ") QB Where QA.医嘱ID=QB.医嘱ID"
        Call objCurParse.init(strSql)
        
        Set mobjSqlParseEvent = objCurParse
        
        Call DoQueryState(0, "组织查询SQL.")
        strSql = mobjSqlParseEvent.GetQuerySql(False)
        
    Else
        Set objCurParse = objCurSchemeItem.SqlParse
        
        If objCurParse.IsInit = False Then
            strSql = objCurSchemeItem.SqlScheme.Query
            Call objCurParse.init(strSql)
        End If
        
        Set mobjSqlParseEvent = objCurParse
        
        Call DoQueryState(0, "组织查询SQL.")
        strSql = mobjSqlParseEvent.GetQuerySql(False)
        
        
        strSql = "Select QA.* From (" & strSql & ") QA, (" & strLinkSql & ") QB Where QA.医嘱ID=QB.医嘱ID"
    End If
    
    
    Call DoQueryState(0, "获取查询结果.")
    
    If lngQueryMovedState = 0 Then
        Set ExecuteWithLink = DBQueryAnalyzer(objCurSchemeItem.SqlParse.SqlStruct, strSql, objCurSchemeItem.SqlScheme.HistoryDBHint, mblnIsMoved)
    Else
        Set ExecuteWithLink = DBQueryAnalyzer(objCurSchemeItem.SqlParse.SqlStruct, strSql, objCurSchemeItem.SqlScheme.HistoryDBHint, IIf(lngQueryMovedState = 1, False, True))
    End If
Exit Function
errHandle:
    Call Clipboard.SetText(Err.Description)
    Err.Raise -1, "clsPacsQuery.ExecuteWithLink", "[ExecuteWithLink]处理错误>>" & vbCrLf & "  方案ID为：" & mlngCurSchemeId & vbCrLf & Err.Description
    Resume
End Function


Public Function ExecuteWithAttach(ByVal strAttach As String, ByVal attachData As Variant, _
    Optional ByVal lngQueryMovedState As Long = 0) As ADODB.Recordset
'lngQueryMovedState: 0-直接使用mblnIsMoved的值，1-不使用转储查询，2-使用转储查询
'使用指定的条件数据执行方案
On Error GoTo errHandle
    Dim objCurSchemeItem As TSchemeItem
    Dim objCurParse As clsSqlParse
    Dim strSql As String
    Dim strMovedSql As String
    Dim strExtPro As String
    
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
       
    mblnIsNeedDateRange = IsNeedDataRange(strAttach, objCurSchemeItem.SqlScheme)
'    mblnIsIgnoreSysPar = IsIgnoreSysPar(strAttach, objCurSchemeItem.SqlScheme)
    
    If IsNull(attachData) = False And Len(attachData) > 0 Then
        If objCurSchemeItem.SqlScheme.GetSerachExtValue(strAttach, EXT_UPPERCONVERT) = "1" Then
        '判断是否需要转换为大写
            attachData = UCase(attachData)
        End If
        
        If objCurSchemeItem.SqlScheme.GetSerachExtValue(strAttach, EXT_NUMBERCONVERT) = "1" Then
        '判断是否需要转换为数字
            attachData = Val(attachData)
        End If
        
        If InStr(attachData, "%") <> 1 _
            And InStr(attachData, "%") <> Len(attachData) Then
            '判断是否进行模糊匹配处理
            strExtPro = objCurSchemeItem.SqlScheme.GetSerachExtValue(strAttach, EXT_LIKEWAY)
            
            Select Case strExtPro
                Case EXT_PRO_VALUE_LEFTWAY
                    attachData = attachData & "%"
                Case EXT_PRO_VALUE_RIGHTWAY
                    attachData = "%" & attachData
                Case EXT_PRO_VALUE_FULLWAY
                    attachData = "%" & attachData & "%"
            End Select
        End If
    End If
    
    Set mobjSqlParseEvent = objCurParse
    
    If mblnIsNeedDateRange Then
        Set mobjCurFilterValue = objCurSchemeItem.FilterValues
        Call mobjCurFilterValue.ClearParValue(True)
    End If
    
    Call DoQueryState(0, "组织查询SQL.")
    strSql = mobjSqlParseEvent.GetAttachSql(strAttach, attachData)
    
    Call DoQueryState(0, "查询规则识别.")
    strSql = RuleIdent(strSql, strAttach)
        
    Call DoQueryState(0, "获取查询结果.")
    
    If lngQueryMovedState = 0 Then
        If mblnIsNeedDateRange = False Then
        '当不需要日期范围条件时，需要将mblnIsMoved设置为true，从转储表中查询对应的数据
        '如果有日期范围，则直接使用之前的mblnIsMoved状态
            mblnIsMoved = True
        End If
        
        Set ExecuteWithAttach = DBQueryAnalyzer(objCurSchemeItem.SqlParse.SqlStruct, strSql, objCurSchemeItem.SqlScheme.HistoryDBHint, mblnIsMoved)
    Else
        Set ExecuteWithAttach = DBQueryAnalyzer(objCurSchemeItem.SqlParse.SqlStruct, strSql, objCurSchemeItem.SqlScheme.HistoryDBHint, IIf(lngQueryMovedState = 1, False, True))
    End If
Exit Function
errHandle:
    Call Clipboard.SetText(Err.Description)
    Err.Raise -1, "clsPacsQuery.ExecuteWithAttach", "[ExecuteWithAttach]处理错误>>" & vbCrLf & "  方案ID为：" & mlngCurSchemeId & vbCrLf & Err.Description
    Resume
End Function

Private Function RuleIdent(ByVal strSql As String, ByVal strIdentItem As String) As String
'规则识别符为：/*F=<RULE:姓名;门诊号;住院号;>*/ ，凡是<>中存在的项目，都会被自动替换为RULE
    Dim strRuleDes As String
    Dim strSqlRule As String
    Dim lngStartIndex As Long
    
    RuleIdent = strSql
    
    lngStartIndex = InStr(strSql, "/*F=<")
    If lngStartIndex <= 0 Then
        Exit Function
    End If
    
    strRuleDes = Mid(strSql, lngStartIndex, InStr(strSql, ">*/") - lngStartIndex + 3)
    If InStr(strRuleDes, strIdentItem & ";") > 0 Then
        strSqlRule = Mid(strRuleDes, 6, InStr(strRuleDes, ":") - 6)
        
        RuleIdent = Replace(strSql, strRuleDes, "/*+ " & strSqlRule & "*/")
    Else
        RuleIdent = Replace(strSql, strRuleDes, "")
    End If
End Function

Private Function IsNeedDataRange(ByVal strFindItem As String, objCurScheme As clsSqlScheme) As Boolean
'判断是否需要结合系统日期范围查找
    Dim i As Long
    
    IsNeedDataRange = True
    
    If strFindItem = "系统.医嘱ID" Or strFindItem = "[系统.医嘱ID]" Then
        IsNeedDataRange = False
        Exit Function
    End If
    
    If strFindItem = "系统.病人ID" Or strFindItem = "[系统.病人ID]" Then
        IsNeedDataRange = False
        Exit Function
    End If
    
    For i = 1 To objCurScheme.SerachCfgCount
        If objCurScheme.SerachCfg(i).Name = strFindItem Or "[" & objCurScheme.SerachCfg(i).Name & "]" = strFindItem Then
            If Val(GetExtPropertyValue(objCurScheme.SerachCfg(i).ExtProperty, EXT_DATERANGE)) = 0 Then
                IsNeedDataRange = False
            End If
            Exit Function
        End If
    Next
End Function


'Private Function IsIgnoreSysPar(ByVal strFindItem As String, objCurScheme As clsSqlScheme) As Boolean
''是否忽略查询相关的系统参数,如科室ID,用户账号等
''默认是不忽视系统参数，只有使用医嘱id或病人ID时，才忽视系统参数
'    Dim i As Long
'
'    IsIgnoreSysPar = False
'
'    If strFindItem = "系统.医嘱ID" Or strFindItem = "[系统.医嘱ID]" Then
'        IsIgnoreSysPar = True
'        Exit Function
'    End If
'
'    If strFindItem = "系统.病人ID" Or strFindItem = "[系统.病人ID]" Then
'        IsIgnoreSysPar = True
'        Exit Function
'    End If
'
'    For i = 1 To objCurScheme.SerachCfgCount
'        If objCurScheme.SerachCfg(i).Name = strFindItem Or "[" & objCurScheme.SerachCfg(i).Name & "]" = strFindItem Then
'            If Val(GetExtPropertyValue(objCurScheme.SerachCfg(i).ExtProperty, EXT_IGNORESYSPAR)) = 1 Then
'                IsIgnoreSysPar = True
'            End If
'
'            Exit Function
'        End If
'    Next
'End Function



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
    
    Set QuerySchemeData = ExecuteCoreEx(strSql, "执行查询", mobjSqlParseEvent)
    
Exit Function
errHandle:
    Err.Raise -1, "clsPacsQuery.QuerySchemeData", "[QuerySchemeData]处理错误>>" & vbCrLf & "  查询语句为：" & strSql & vbCrLf & Err.Description
    Resume
End Function



Private Sub Class_Initialize()
    ReDim maryScheme(0)
    ReDim mstrBakTables(0)
    
    mstrStationName = UCase(StationName)
    gblnTimeChanged = False
End Sub

Private Sub Class_Terminate()
    Dim i As Long
    
    Set mobjSqlParseEvent = Nothing
    Set mobjSchemeCfg = Nothing
    Set mobjCurFilterValue = Nothing
    Set mobjRegExp = Nothing
    
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
 
    If mblnIsNeedDateRange = False Then
        If (strParName = "系统.开始日期" Or strParName = "系统.结束日期") Then
            Value = Null
            Exit Sub
        End If
    End If
    
'    If mblnIsIgnoreSysPar Then
'    '忽略除[系统.开始日期]和[系统.结束日期]之外的系统参数
'
'    End If
    
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
