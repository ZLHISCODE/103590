VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueryShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mbytFontSize As Byte
Private mlngSchemeNo As Long '当前使用的方案号

Private WithEvents mfrmPacsQuery As frmPacsQuery  '正常工作窗体
Attribute mfrmPacsQuery.VB_VarHelpID = -1
Private WithEvents mfrmPacsQueryDisplay As frmPacsQuery  '方案配置时点击 测试 显示的窗口
Attribute mfrmPacsQueryDisplay.VB_VarHelpID = -1
Private WithEvents mfrmQueryDisplay As frmQueryDisplay '方案配置测试 中的 列表窗体
Attribute mfrmQueryDisplay.VB_VarHelpID = -1

Private mDataGrid As VSFlexGrid
Private mlngModule As Long '模块号
Private mlngUserId As Long '病人ID
Private mstrRoomID As String '科室ID
Private mcnOracle As ADODB.Connection
Private mstrCachePath As String
Private mlngSys As Long
Private mstrDBUser As String
Private mfrmParent As Object


Private mintShowType As Integer  '0-用于pacsMain   1-方案验证

Public Event OnSetParaForPacsMain( _
    ByRef lngModule As Long, _
    ByRef lngUserId As Long, _
    ByRef strRoomID As String, _
    ByRef cnOracle As ADODB.Connection, _
    ByRef strCachePath As String, _
    ByRef lngSys As Long, _
    ByRef strDBUser As String, _
    ByRef bytFontSize As Byte, _
    ByRef frmParent As Object)
    
Public Event OnSetParaForTest( _
    ByRef lngModule As Long, _
    ByRef lngUserId As Long, _
    ByRef strRoomID As String, _
    ByRef cnOracle As ADODB.Connection, _
    ByRef strCachePath As String, _
    ByRef lngSys As Long, _
    ByRef strDBUser As String)
    
Public Event OnListRowSelClear() '列表选中项目重置（例如经过快速过滤后导致列表不显示数据，此时需要同步更新一些状态）
Public Event OnColStatistics(ByVal strStatisticsInfo As String)   '进行列统计
Public Event OnDblClick() '双击
Public Event OnRefreshSelectTab(ByVal lngAdvideID As Long)
Public Event OnSelectScheme(ByVal strName As String)
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OnSelChange(ByVal blnPassive As Boolean)


Property Get ShowingRowCount() As Integer
'取得当前显示行的数量
    Dim i As Long
    Dim lngCount As Long
    
    lngCount = 0
    For i = 1 To mDataGrid.Rows - 1
        If Not mDataGrid.RowHidden(i) Then lngCount = lngCount + 1
    Next i
    
    ShowingRowCount = lngCount
End Property

Property Get rsDataShow() As ADODB.Recordset
    If Not mfrmPacsQuery Is Nothing Then Set rsDataShow = mfrmPacsQuery.rsDataShow
    If Not mfrmQueryDisplay Is Nothing Then Set rsDataShow = mfrmPacsQueryDisplay.rsDataShow
End Property

Property Get rsData() As ADODB.Recordset
    If Not mfrmPacsQuery Is Nothing Then Set rsData = mfrmPacsQuery.rsData
    If Not mfrmQueryDisplay Is Nothing Then Set rsData = mfrmPacsQueryDisplay.rsData
End Property

Property Get objSqlScheme() As clsSqlScheme
    If Not mfrmPacsQuery Is Nothing Then Set objSqlScheme = mfrmPacsQuery.objSqlScheme
End Property


Property Get objQuery() As clsPacsQuery
    If Not mfrmPacsQuery Is Nothing Then Set objQuery = mfrmPacsQuery.objQuery
End Property

Property Get SelectionRow() As Long
     SelectionRow = mDataGrid.RowSel
End Property

Property Get cboHistorys() As Object
    Set cboHistorys = mfrmPacsQuery.cboHistory
End Property

Property Get DataGrid() As Object
    Set DataGrid = mDataGrid
End Property

Property Get HwndOfFace() As Long
    HwndOfFace = -1
    If Not mfrmPacsQuery Is Nothing Then
        HwndOfFace = mfrmPacsQuery.hwnd

    ElseIf Not mfrmQueryDisplay Is Nothing Then
        HwndOfFace = mfrmQueryDisplay.hwnd

    End If
    
End Property

Property Get Width() As Long
    If Not mfrmPacsQuery Is Nothing Then Width = mfrmPacsQuery.Width
End Property

Property Get HaveFace() As Boolean
    HaveFace = (Not mfrmPacsQuery Is Nothing) Or (Not mfrmQueryDisplay Is Nothing)
End Property


Public Sub init(ByVal intCallType As Integer)
'初始化，根据类型不同，触发获取参数事件

    Set mcnOracle = New ADODB.Connection
    
    If intCallType = 0 Then
        Set mfrmPacsQuery = New frmPacsQuery
        RaiseEvent OnSetParaForPacsMain(mlngModule, mlngUserId, mstrRoomID, mcnOracle, mstrCachePath, mlngSys, mstrDBUser, mbytFontSize, mfrmParent)
        
'        Call mfrmPacsQuery.SetVars(mcnOracle, mlngModule, mlngUserId, mstrRoomID, -1, 0, mlngSys, mstrDBUser, mbytFontSize)
        Call mfrmPacsQuery.SetVars(varName_数据库连接, mcnOracle)
        Call mfrmPacsQuery.SetVars(varName_模块号, mlngModule)
        Call mfrmPacsQuery.SetVars(varName_用户ID, mlngUserId)
        Call mfrmPacsQuery.SetVars(varName_科室ID, mstrRoomID)
        Call mfrmPacsQuery.SetVars(varName_系统号, mlngSys)
        Call mfrmPacsQuery.SetVars(varName_数据库用户名, mstrDBUser)
        Call mfrmPacsQuery.SetVars(varName_字号, mbytFontSize)
        Call mfrmPacsQuery.SetVars(varName_查询方案ID, -1)
        Call mfrmPacsQuery.SetVars(varName_查询界面类型, 0)
        Call mfrmPacsQuery.SetVars(varName_列表关键字, "医嘱ID")
        Call mfrmPacsQuery.SetVars(varName_父窗体, mfrmParent)
        
        
        Call mfrmPacsQuery.init
        Set mDataGrid = mfrmPacsQuery.vsfList
    Else
        mintShowType = intCallType
        RaiseEvent OnSetParaForTest(mlngModule, mlngUserId, mstrRoomID, mcnOracle, mstrCachePath, mlngSys, mstrDBUser)
         
        If mfrmQueryDisplay Is Nothing Then Set mfrmQueryDisplay = New frmQueryDisplay
        If mfrmPacsQueryDisplay Is Nothing Then Set mfrmPacsQueryDisplay = New frmPacsQuery
         
        Call SetParent(mfrmPacsQueryDisplay.hwnd, mfrmQueryDisplay.hwnd)
        Call MoveWindow(mfrmPacsQueryDisplay.hwnd, 0, 0, _
                        mfrmPacsQueryDisplay.ScaleX(mfrmQueryDisplay.Width - 210, vbTwips, vbPixels), _
                        mfrmPacsQueryDisplay.ScaleY(mfrmQueryDisplay.Height, vbTwips, vbPixels), _
                        0)
                        
'        Call mfrmPacsQueryDisplay.SetVars(mcnOracle, mlngModule, mlngUserId, mstrRoomID, 0, 1, mlngSys, mstrDBUser)
        Call mfrmPacsQueryDisplay.SetVars(varName_数据库连接, mcnOracle)
        Call mfrmPacsQueryDisplay.SetVars(varName_模块号, mlngModule)
        Call mfrmPacsQueryDisplay.SetVars(varName_用户ID, mlngUserId)
        Call mfrmPacsQueryDisplay.SetVars(varName_科室ID, mstrRoomID)
        Call mfrmPacsQueryDisplay.SetVars(varName_系统号, mlngSys)
        Call mfrmPacsQueryDisplay.SetVars(varName_数据库用户名, mstrDBUser)
        Call mfrmPacsQueryDisplay.SetVars(varName_字号, mbytFontSize)
        Call mfrmPacsQueryDisplay.SetVars(varName_查询方案ID, 1)
        Call mfrmPacsQueryDisplay.SetVars(varName_查询界面类型, 1)
        Call mfrmPacsQueryDisplay.SetVars(varName_列表关键字, "医嘱ID")
        
        Call mfrmPacsQueryDisplay.init
        Set mDataGrid = mfrmPacsQueryDisplay.vsfList
                        
    End If
    
End Sub

'Public Function SetVars( _
'    ByVal cnOracle As ADODB.Connection, _
'    Optional ByVal lngModule As Long = 0, _
'    Optional ByVal lngUserId As Long = 0, _
'    Optional ByVal strCurRoom As String = "0", _
'    Optional ByVal lngSchemeId As Long = 0, _
'    Optional ByVal intShowType As Integer = 0, _
'    Optional ByVal lngSys As Long = 0, _
'    Optional ByVal strDBUser As String = 0)
''需要的变量和参数赋值
'
''    Set mcnOracle = cnOracle
''    mlngModule = lngModule
''    mlngUserID = lngUserID
'    mstrRoomID = strCurRoom
''    mlngSchemeID = lngSchemeId
''    mintShowType = intShowType
''    mstrListKeyCol = "医嘱ID"
'
'    If Not mfrmPacsQuery Is Nothing Then Call mfrmPacsQuery.SetVars(mcnOracle, mlngModule, mlngUserId, mstrRoomID, -1, 0, lngSys, strDBUser)
    
Public Sub SetVars(ByVal strVarName As String, ByVal Value As Variant)
    If Not mfrmPacsQuery Is Nothing Then Call mfrmPacsQuery.SetVars(strVarName, Value)
End Sub

Private Sub Class_Terminate()

    Set mfrmParent = Nothing

    If Not mfrmPacsQuery Is Nothing Then
        Unload mfrmPacsQuery
        Set mfrmPacsQuery = Nothing
    End If
    
    If Not mfrmQueryDisplay Is Nothing Then
        Unload mfrmQueryDisplay
        Set mfrmQueryDisplay = Nothing
    End If
    
    
    
End Sub

Private Sub mfrmPacsQuery_OnColStatistics(ByVal strStatisticsInfo As String)
    If Not mDataGrid Is Nothing Then RaiseEvent OnColStatistics(strStatisticsInfo)
End Sub

Private Sub mfrmPacsQuery_OnDblClick()
    RaiseEvent OnDblClick
End Sub

Private Sub mfrmPacsQuery_OnListRowSelClear()
    RaiseEvent OnListRowSelClear
End Sub

Private Sub mfrmPacsQuery_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
End Sub

Private Sub mfrmPacsQuery_OnRefreshSelectTab(ByVal lngAdviceID As Long)
    RaiseEvent OnRefreshSelectTab(lngAdviceID)
End Sub


Private Sub mfrmPacsQuery_OnSelChange()
    RaiseEvent OnSelChange(True)
End Sub

Private Sub mfrmPacsQuery_OnSelectScheme(ByVal strName As String)
    RaiseEvent OnSelectScheme(strName)
End Sub

Private Sub mfrmPacsQueryDisplay_OnSelectScheme(ByVal strName As String)
    RaiseEvent OnSelectScheme(strName)
End Sub

Private Sub mfrmQueryDisplay_OnFormQueryUnload()
    If Not mfrmPacsQueryDisplay Is Nothing Then
        Unload mfrmPacsQueryDisplay
        Set mfrmPacsQueryDisplay = Nothing
    End If
    
    If Not mfrmQueryDisplay Is Nothing Then
        Unload mfrmQueryDisplay
        Set mfrmQueryDisplay = Nothing
    End If
    
End Sub


Public Sub ShowTest(ByVal frmParent As Object)

    If mfrmPacsQueryDisplay Is Nothing Or mfrmQueryDisplay Is Nothing Then Exit Sub

    mfrmPacsQueryDisplay.Visible = True
    mfrmQueryDisplay.Visible = True
End Sub

Private Sub mfrmQueryDisplay_OnTestFormResize()
    If Not mfrmPacsQueryDisplay Is Nothing Then
        Call MoveWindow(mfrmPacsQueryDisplay.hwnd, 0, 0, _
                        mfrmPacsQueryDisplay.ScaleX(mfrmQueryDisplay.Width - 210, vbTwips, vbPixels), _
                        mfrmPacsQueryDisplay.ScaleY(mfrmQueryDisplay.Height, vbTwips, vbPixels), _
                        0)
    End If
End Sub


Public Sub LocateRow(Optional ByVal lngRowIndex As Long = -1)
'定位指定行，默认定位为最后一位
    Dim lngRow As Long
    Dim iCol As Long
    
    If mDataGrid.Rows <= 1 Then Exit Sub
    
    lngRow = lngRowIndex
    If lngRow < 0 Then
        lngRow = mDataGrid.Rows - 1
    End If
    
    '取得第一个未隐藏的列
    For iCol = 1 To mDataGrid.Cols - 1
        If Not mDataGrid.ColHidden(iCol) Then Exit For
    Next iCol
    
    Call mDataGrid.Select(lngRow, iCol)
    Call mDataGrid.ShowCell(lngRow, iCol)
End Sub

Public Function FindRowIndex(ByVal strFindValue As String, ByVal strColName As String, _
    Optional ByVal blnIsPrecise As Boolean = False) As Long
'查找指定值并返回所在行索引
    Dim i As Long
    Dim lngCol As Long

    FindRowIndex = -1
    If Trim(strFindValue) = "" Then Exit Function

    lngCol = mDataGrid.ColIndex(strColName)

    For i = 1 To mDataGrid.Rows - 1
        If Not mDataGrid.RowHidden(i) Then

            If UCase(mDataGrid.TextMatrix(i, lngCol)) Like IIf(blnIsPrecise, UCase(strFindValue), "*" & UCase(strFindValue) & "*") Then
                FindRowIndex = i
                Exit Function
            End If
        End If
    Next i
End Function

Public Function QueryConfig(ByVal lngModule As Long, ByVal frmParent As Object) As Boolean
    Dim strSys As String

    strSys = "[系统.系统号],[系统.模块号],[系统.科室ID],[系统.用户ID],[系统.患者ID],[系统.医嘱ID]"
    strSys = strSys & ",[系统.服务器日期],[系统.服务器时间],[系统.本地日期],[系统.本地时间]"
    strSys = strSys & "[,系统.开始日期],[系统.结束日期],[系统.起始行],[系统.结束行]"
    
    If mfrmPacsQuery.objQuery Is Nothing Then Exit Function
    
    With mfrmPacsQuery.objQuery
        Call mfrmPacsQuery.objQuery.ShowSchemeCfg(lngModule, strSys, frmParent)
    End With
End Function

Public Function UpdateRow(ByVal blIsAdd As Boolean, ByVal lngAdviceID As Long, ByVal blSelectThisRow As Boolean) As Boolean
'根据医嘱ID,重新查询一行数据并且刷新行
    UpdateRow = False
    If mfrmPacsQuery.UpdateRow(blIsAdd, lngAdviceID) = True Then UpdateRow = True
    If blSelectThisRow Then RaiseEvent OnSelChange(False)
End Function

Public Function ExecuteWithLink(ByVal strSql As String) As Boolean
    Call mfrmPacsQuery.ExecuteWithLink(strSql)
End Function

Public Sub RefreshListData()
'刷新检查列表,功能相当于点击刷新
    Call mfrmPacsQuery.ExecuteQuery("刷新")
    
End Sub

Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
    mbytFontSize = bytFontSize
    If Not mfrmPacsQuery Is Nothing Then Call mfrmPacsQuery.ReSetFormFontSize(bytFontSize)
End Sub

Public Function ExecuteQuery(ByVal strExecuteType As String, Optional ByVal LngSetRow As Long = 0) As Boolean
    If Not mfrmPacsQuery Is Nothing Then Call mfrmPacsQuery.ExecuteQuery(strExecuteType, LngSetRow)
End Function



