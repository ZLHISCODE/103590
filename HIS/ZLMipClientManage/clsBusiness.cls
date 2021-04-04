VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusiness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'变量定义

Private mstrSQL As String
Private mclsDataOracle As zlDataOracle.clsDataOracle
Private mstrTitle As String
Private mstrDbUser As String
Private mrsSQL As ADODB.Recordset
Private mstrOwner As String

'######################################################################################################################
'接口方法
Public Function InitBusiness(ByVal objDataOracle As zlDataOracle.clsDataOracle, ByVal strTitle As String, ByVal strOwner As String) As Boolean
    '******************************************************************************************************************
    '功能：初始化
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Set mclsDataOracle = objDataOracle
    mstrTitle = strTitle
    mstrOwner = strOwner & "."
'    mstrOwner = ""
    
    InitBusiness = True
    
    Exit Function
    
errHand:
    
End Function

Public Function CheckSetuped(ByVal strCode As String) As Byte
    '******************************************************************************************************************
    '功能：根据标识查询是否已安装该系统消息数据
    '参数：strCode 标识
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "Select 1 From zlmip_data_setup Where Data_Code=[1]"
    Set rsData = zlDataBase.OpenSQLRecord(strSQL, gstrSysName, strCode)
    If rsData.BOF = False Then
        CheckSetuped = 1
    Else
        CheckSetuped = 0
    End If
    
End Function

Public Function Get_Ver() As String
    '******************************************************************************************************************
    '功能：获取消息客户端的当前版本号
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select 内容 From zlRegInfo Where 项目='消息集成平台客户端'"
    Set rsData = zlDataBase.OpenSQLRecord(strSQL, gstrSysName)
    If rsData.BOF = False Then
        Get_Ver = rsData("内容").Value
    Else
        Get_Ver = ""
    End If
    
End Function

Public Function GetImageList(Optional ByVal intIconSize As Integer = 16) As Object
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    
    Set GetImageList = frmResource.GetImageList(intIconSize)
    
End Function

Public Function GetMaxCode(ByVal strTable As String, ByVal strField As String, Optional ByVal intCodeLength As Integer = 4) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim rsTemp As ADODB.Recordset
    Dim str最大编码 As String
    
    mstrSQL = "Select Max(" & strField & ") As 最大编码 From " & strTable
    
    Set rsTemp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
    If rsTemp.BOF = False Then
        str最大编码 = Trim(zlCommFun.NVL(rsTemp("最大编码").Value))
    End If
    
    If str最大编码 = "" Then
        GetMaxCode = Format(Val(str最大编码) + 1, String(intCodeLength, "0"))
    Else
        GetMaxCode = Format(Val(str最大编码) + 1, String(Len(str最大编码), "0"))
    End If
    
End Function

Public Function GetMaxUserTableCode(ByVal strPre As String, Optional ByVal intCodeLength As Integer = 3) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim rsTemp As ADODB.Recordset
    Dim str最大编码 As String
    
    mstrSQL = "Select Max(tab_code) As 最大编码 From zlmip_table Where tab_type=2"
    
    Set rsTemp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
    If rsTemp.BOF = False Then
        str最大编码 = Trim(zlCommFun.NVL(rsTemp("最大编码").Value))
        str最大编码 = Mid(str最大编码, InStr(str最大编码, strPre) + Len(strPre))
    End If
    
    If str最大编码 = "" Then
        GetMaxUserTableCode = Format(Val(str最大编码) + 1, String(intCodeLength, "0"))
    Else
        GetMaxUserTableCode = Format(Val(str最大编码) + 1, String(Len(str最大编码), "0"))
    End If
    
End Function

Public Function IsOwner() As Boolean
    '******************************************************************************************************************
    '功能：读取Infotable结构
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    
    mstrSQL = "SELECT 1 FROM ZLTOOLS.ZLSYSTEMS WHERE upper(所有者)=USER"
    Set rsTmp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
    IsOwner = (rsTmp.BOF = False)
            
End Function

Public Function IsDBAUser() As Boolean
    '******************************************************************************************************************
    '功能：读取Infotable结构
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    
    mstrSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTmp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
    IsDBAUser = (rsTmp.BOF = False)
            
End Function

Public Function GetSystemOwner(ByVal blnDBA As Boolean, ByVal strDBUser As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
        
    On Error GoTo errHand
    
    If blnDBA = True Then
        mstrSQL = "Select Distinct 所有者 From zlSystems"
    Else
        mstrSQL = "Select Distinct 所有者 From zlSystems Where Upper(所有者) = [1]"
    End If
    Set GetSystemOwner = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strDBUser)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetBusinessKind(ByVal blnDBA As Boolean, ByVal strDBUser As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
        
    On Error GoTo errHand
    

    If blnDBA = True Then
        mstrSQL = "Select rpad(a.data_code,5,' ')||' - '||a.data_title As 业务,a.data_source As 类型 From zltools.zlmip_data_setup a"
    Else
        mstrSQL = "Select rpad(a.data_code,5,' ')||' - '||a.data_title As 业务,a.data_source As 类型 From zltools.zlmip_data_setup a Where Upper(a.data_owner) = [1]"
    End If

    Set GetBusinessKind = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, UCase(strDBUser))
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function TableStruct() As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取Infotable结构
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mstrSQL = "Select * From zlmip_table Where 1=2"
    Set TableStruct = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function TableRelationStruct() As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取Infotable结构
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mstrSQL = "Select * From zlmip_tab_extend Where 1=2"
    Set TableRelationStruct = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function


Public Function TableRelationEdit(Optional ByVal strUpdateMode As String = "", Optional ByVal rsPara As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim strLine As String
    Dim aryLine As Variant
    
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)
    
    Select Case UCase(strUpdateMode)
    '------------------------------------------------------------------------------------------------------------------
    Case "INSERT", "UPDATE"
            
        '消息
        mstrSQL = "Tab_Extend_Update('" & zlCommFun.GetParameter(rsPara, "id") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "来源信息") & "'," & _
                                            Val(zlCommFun.GetParameter(rsPara, "关系类型")) & "," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "关系命名") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "目标信息") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
        
        '赋值
        mstrSQL = "TabExt_Condition_Delete('" & zlCommFun.GetParameter(rsPara, "id") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
        lngCount = Val(zlCommFun.GetParameter(rsPara, "赋值关系段数"))
        For lngLoop = 1 To lngCount
            strTemp = Trim(zlCommFun.GetParameter(rsPara, "赋值关系_" & lngLoop))
            If strTemp <> "" Then
                mstrSQL = "TabExt_Condition_Insert('" & zlCommFun.GetParameter(rsPara, "id") & "','" & strTemp & "')"
                mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
                Call SQLRecordAdd(mrsSQL, mstrSQL)
            End If
        Next

    '------------------------------------------------------------------------------------------------------------------
    Case "DELETE"
        mstrSQL = "Tab_Extend_Delete('" & zlCommFun.GetParameter(rsPara, "id") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    End Select
        
    TableRelationEdit = SQLRecordExecute(True)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
    
End Function

Public Function EventStruct() As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取Event结构
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mstrSQL = "Select * From m_Event Where 1=2"
    Set EventStruct = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function


Public Function ReadEventKind() As ADODB.Recordset
    
    On Error GoTo errHand
    
    mstrSQL = "Select Distinct kind From m_Event Where kind is not null"
    Set ReadEventKind = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function

Public Function EventRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strStyle As String
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        mstrSQL = "Select 'event' As 图标,a.* From m_Event a Order By a.code"
        Set EventRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("ID")
        mstrSQL = "Select 'event' As 图标,a.* From m_Event a Where id=[1]"
        Set EventRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "id"))
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("Code")
        mstrSQL = "Select a.ID,a.code As 编码,a.title As 名称,a.app As 程序,a.device As 设备,a.kind as 类型 From m_Event a Where code Like [1]"
        Set EventRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "code"))
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("Title")
        mstrSQL = "Select a.ID,a.code As 编码,a.title As 名称,a.app As 程序,a.device As 设备,a.kind as 类型 From m_Event a Where title Like [1]"
        Set EventRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "title"))
        
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("FilterData")
        
        strStyle = zlCommFun.GetCondition(rsCondition, "FilterStyle")
        strTemp = zlCommFun.GetCondition(rsCondition, "FilterText")
        
        Select Case strStyle
        Case ""
            mstrSQL = "Select a.ID,a.code As 编码,a.title As 名称,a.app As 程序,a.device As 设备,a.kind as 类型 From m_Event a Where code Like [1]  Or a.title Like [2]"
            Set EventRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strTemp & "%", "%" & strTemp & "%")
        Case "编码"
            mstrSQL = "Select a.ID,a.code As 编码,a.title As 名称,a.app As 程序,a.device As 设备,a.kind as 类型 From m_Event a Where code Like [1]"
            Set EventRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strTemp & "%")
        Case "名称"
            
            mstrSQL = "Select a.ID,a.code As 编码,a.title As 名称,a.app As 程序,a.device As 设备,a.kind as 类型 From m_Event a Where title Like [1]"
            Set EventRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, "%" & strTemp & "%")
        End Select
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

'Public Function EventEdit1(Optional ByVal strUpdateMode As String = "", Optional ByVal rsParameter As ADODB.Recordset) As Boolean
'    '******************************************************************************************************************
'    '功能：编辑Event数据
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'
'    On Error GoTo errHand
'
'    Call SQLRecord(mrsSQL)
'
'    Select Case UCase(strUpdateMode)
'    '------------------------------------------------------------------------------------------------------------------
'    Case "INSERT", "UPDATE"
'        mstrSQL = "Event_Update('" & zlCommFun.GetParameter(rsParameter, "id") & "'," & _
'                                            "'" & zlCommFun.GetParameter(rsParameter, "code") & "'," & _
'                                            "'" & zlCommFun.GetParameter(rsParameter, "title") & "'," & _
'                                            "'" & zlCommFun.GetParameter(rsParameter, "app") & "'," & _
'                                            "'" & zlCommFun.GetParameter(rsParameter, "device") & "'," & _
'                                            "'" & zlCommFun.GetParameter(rsParameter, "kind") & "'," & _
'                                            "'" & zlCommFun.GetParameter(rsParameter, "note") & "')"
'        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
'        Call SQLRecordAdd(mrsSQL, mstrSQL)
'    '------------------------------------------------------------------------------------------------------------------
'    Case "DELETE"
'        mstrSQL = "Event_Delete('" & zlCommFun.GetParameter(rsParameter, "id") & "')"
'        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
'        Call SQLRecordAdd(mrsSQL, mstrSQL)
'    End Select
'
'    EventEdit1 = SQLRecordExecute(False)
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'
'End Function

Public Function EventMsgRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        
        mstrSQL = "Select 'event' As 图标,a.id,a.event_id,a.msg_code as 消息,a.msg_ver as 版本,b.tab_title as 入口,a.tab_id,a.send_fail,a.note As 说明 " & _
                    "From m_Evt_Message a,zlmip_table b Where a.tab_id=b.id(+) And a.event_id=[1]"
                    
        Set EventMsgRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "业务事件id"))
    '------------------------------------------------------------------------------------------------------------------
    Case "事件消息"
    
        mstrSQL = "Select 'event' As 图标,a.id,a.event_id,a.msg_code as 消息,a.msg_ver as 版本,b.tab_title as 入口,a.tab_id,a.send_fail,a.note As 说明 " & _
                    "From m_Evt_Message a,zlmip_table b Where a.tab_id=b.id(+) And a.id=[1]"
                    
        Set EventMsgRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "事件消息id"))
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("Select")
        mstrSQL = _
            "Select ID, Null As 上级id, Null As 消息编号, Title As 名称, Null As 入口信息, 0 As 末级,Null As 说明" & vbNewLine & _
            "From m_Event" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select a.Id, a.Event_Id As 上级id, a.Msg_Code As 消息编号, Null As 名称, b.tab_Title As 入口信息, 1 As 末级,a.Note As 说明" & vbNewLine & _
            "From m_Evt_Message A, zlmip_table B" & vbNewLine & _
            "Where a.tab_id = b.Id(+)"
        Set EventMsgRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function EventMsgEdit(Optional ByVal strUpdateMode As String = "", Optional ByVal rsPara As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim strLine As String
    Dim aryLine As Variant
    
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)
    
    Select Case UCase(strUpdateMode)
    '------------------------------------------------------------------------------------------------------------------
    Case "INSERT", "UPDATE"
            
        '消息
        mstrSQL = "Evt_Message_Update('" & zlCommFun.GetParameter(rsPara, "id") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "业务事件id") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "消息标识") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "消息版本") & "'," & _
                                            IIf(zlCommFun.GetParameter(rsPara, "入口信息") = "", "Null", "'" & zlCommFun.GetParameter(rsPara, "入口信息") & "'") & "," & _
                                            Val(zlCommFun.GetParameter(rsPara, "失败处理")) & "," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "说明") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
        
        '配置
        mstrSQL = "EvtMsg_Config_Delete('" & zlCommFun.GetParameter(rsPara, "id") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
        lngCount = Val(zlCommFun.GetParameter(rsPara, "消息配置段数"))
        For lngLoop = 1 To lngCount
            strTemp = Trim(zlCommFun.GetParameter(rsPara, "消息配置_" & lngLoop))
            
            If strTemp <> "" Then
                mstrSQL = "EvtMsg_Config_Insert('" & zlCommFun.GetParameter(rsPara, "id") & "','" & strTemp & "')"
                mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
                Call SQLRecordAdd(mrsSQL, mstrSQL)
            End If
        Next
                
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("Copy")
            
        '消息
        mstrSQL = "Evt_Message_Copy('" & zlCommFun.GetParameter(rsPara, "业务事件id") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "参考消息id") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
                               
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("TargetConfig")
    
        '投递目标服务器
        mstrSQL = "EvtMsg_Target_Update('" & zlCommFun.GetParameter(rsPara, "id") & "','" & Trim(zlCommFun.GetParameter(rsPara, "投递目标")) & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("BatchTargetConfig")
        
        '投递目标服务器
        mstrSQL = "EvtMsg_Target_BatchUpdate('" & zlCommFun.GetParameter(rsPara, "TargetService") & "'," & Trim(zlCommFun.GetParameter(rsPara, "ConfigOption")) & ",'" & Trim(zlCommFun.GetParameter(rsPara, "ConfigString")) & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
    '------------------------------------------------------------------------------------------------------------------
    Case "DELETE"
        mstrSQL = "Evt_Message_Delete('" & zlCommFun.GetParameter(rsPara, "事件消息id") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    End Select
        
    EventMsgEdit = SQLRecordExecute(True)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
    
End Function

Public Function ItemConfigRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("item_id")
        
        
        mstrSQL = _
                "Select a.id,node_order," & vbNewLine & _
                "       a.parent_id," & vbNewLine & _
                "       a.node_title as 节点标题," & vbNewLine & _
                "       Decode(a.node_type, 1,'Complex', 2,'Data') as 节点类型," & vbNewLine & _
                "       a.data_type as 数据类型," & vbNewLine & _
                "       a.min_occurs as 最小次数," & vbNewLine & _
                "       a.max_occurs As 最大次数," & vbNewLine & _
                "       a.config_occurs as 数据重复," & vbNewLine & _
                "       a.config_express as 数据赋值," & vbNewLine & _
                "       a.config_occurs_key as 数据重复_Key," & vbNewLine & _
                "       a.config_express_key as 数据赋值_Key" & vbNewLine & _
                "  From zlmip_item_config a" & vbNewLine & _
                " Where a.item_id = [1]" & vbNewLine & _
                " Order by a.node_order"
                                
        'substr('                                  ', 1, 2*(level - 1)) || 节点类型 As 节点类型
        
        mstrSQL = _
                "select level,id,parent_id,'' as 图标," & vbNewLine & _
                "       节点类型," & vbNewLine & _
                "       节点标题," & vbNewLine & _
                "       最小次数 || ' ～ ' || 最大次数 As 重复频率," & vbNewLine & _
                "       数据类型," & vbNewLine & _
                "       数据重复," & vbNewLine & _
                "       数据赋值," & vbNewLine & _
                "       数据重复_Key," & vbNewLine & _
                "       数据赋值_Key" & vbNewLine & _
                "  From (" & mstrSQL & ") " & vbNewLine & _
                " start with parent_id is null" & vbNewLine & _
                "connect by prior id = parent_id Order by node_order"
        
        Set ItemConfigRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "item_id"))
        
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function

Public Function EventMsgServerRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取TargetServer数据
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""

        mstrSQL = "Select a.ID,a.title As 名称,Decode(a.interface_type,1,'Socket',2,'WebService')||':'||a.interface_para As 接口,a.app As 程序,a.device As 设备,a.note As 注释 " & _
                    "From m_Service a,m_EvtMsg_Target b " & _
                    "Where b.evt_msg_id=[1] And b.target_service_id=a.id And a.kind=1"
                    
        Set EventMsgServerRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "事件消息id"))

    '------------------------------------------------------------------------------------------------------------------
    Case "配置"
        If rsCondition Is Nothing Then
            mstrSQL = "Select a.ID,0 As 选择,Decode(a.interface_type,1,'Socket',2,'WebService')||':'||a.interface_para As 接口,a.app As 程序,a.title As 名称,a.device As 设备,a.Note As 注释 From m_Service a Where a.Kind=1"
            Set EventMsgServerRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
        Else
            mstrSQL = "Select a.ID,Decode(b.evt_msg_id,null,0,1) As 选择,Decode(a.interface_type,1,'Socket',2,'WebService')||':'||a.interface_para As 接口,a.title As 名称,a.app As 程序,a.device As 设备,a.note As 注释 " & _
                        "From m_Service a,m_EvtMsg_Target b " & _
                        "Where b.evt_msg_id(+)=[1] And b.target_service_id(+)=a.id And a.Kind=1"
                        
            Set EventMsgServerRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "事件消息id"))
        End If
        
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function ReadEventMsg(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取TargetServer数据
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case "所有"

        mstrSQL = "select 0 As 选择,b.id,a.title As 事件,b.msg_code As 消息 from m_Event a,m_Evt_Message b Where a.id=b.event_id Order by a.code,b.msg_code "
                    
        Set ReadEventMsg = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)

                
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function GetTableTree(ByVal strInfoDataKey As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    
    
    Set rsData = New ADODB.Recordset
    With rsData
        .Fields.Append "id", adVarChar, 100
        .Fields.Append "上级id", adVarChar, 100
        .Fields.Append "名称", adVarChar, 100
        .Fields.Append "末级", adTinyInt
        .Fields.Append "类型", adTinyInt
        .Fields.Append "关系", adTinyInt
        .Fields.Append "序号", adInteger
        
        .Open
    End With
    
    mstrSQL = "Select id,'' as parent_id,tab_title,'' as fld from zlmip_table where id=[1]"
    Set rs = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strInfoDataKey)
    
    If rs.BOF = False Then
        '取字段
        
        rsData.AddNew
        
        rsData("序号").Value = 1
        
        rsData("id").Value = "T_" & strInfoDataKey
        rsData("上级id").Value = ""
        rsData("名称").Value = rs("tab_title").Value
        rsData("末级").Value = 0
        rsData("类型").Value = 0
        rsData("关系").Value = 0
        
        
        If GetInfoTableField(rsData, "T_" & strInfoDataKey, strInfoDataKey) = False Then Exit Function
        If GetInfoChildTable(rsData, "T_" & strInfoDataKey, strInfoDataKey) = False Then Exit Function
        
    End If
    
    If rsData.RecordCount > 0 Then rsData.MoveFirst
    
    Set GetTableTree = rsData
End Function

Private Function GetInfoChildTable(ByRef rsData As ADODB.Recordset, ByVal strParentKey As String, ByVal strDataKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As ADODB.Recordset
            
    mstrSQL = "Select a.id,a.ext_order,a.ext_type,a.ext_title,b.tab_title,a.target_tab_id from zlmip_tab_extend a,zlmip_table b where a.source_tab_id=[1] And a.target_tab_id=b.id"
    Set rs = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strDataKey)
    If rs.BOF = False Then
        
        Do While Not rs.EOF
            
            rsData.AddNew
            rsData("序号").Value = rsData.RecordCount
            
            rsData("id").Value = "R_" & rs("id").Value & "_T_" & strDataKey
            rsData("上级id").Value = strParentKey
            

            rsData("名称").Value = rs("ext_title").Value

            
            rsData("末级").Value = 0
            rsData("关系").Value = zlCommFun.NVL(rs("ext_type").Value, 1)
            rsData("类型").Value = 0
                        
            '取字段
            Call GetInfoTableField(rsData, "R_" & rs("id").Value & "_T_" & strDataKey, rs("target_tab_id").Value)
                        
            '再取此表的下级关系表
            Call GetInfoChildTable(rsData, "R_" & rs("id").Value & "_T_" & strDataKey, rs("target_tab_id").Value)
            
            rs.MoveNext
        Loop
    End If

    GetInfoChildTable = True
    
End Function

Private Function GetInfoTableField(ByRef rsData As ADODB.Recordset, ByVal strParentKey As String, ByVal strTargetDataKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim rs As ADODB.Recordset
    
    mstrSQL = "Select fld_order,fld_title From zlmip_tab_field Where tab_id=[1]"
    
    Set rs = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strTargetDataKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            rsData.AddNew
            
            rsData("序号").Value = rsData.RecordCount
            
            rsData("id").Value = strParentKey & "_F_" & rs("fld_order").Value
            rsData("上级id").Value = strParentKey
            rsData("名称").Value = rs("fld_title").Value
            
            rsData("末级").Value = 0
            rsData("关系").Value = 0
            rsData("类型").Value = 1
            
            
            rs.MoveNext
        Loop
    End If
    
    GetInfoTableField = True
    
End Function

Public Function TableRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取Infotable数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTemp As String
    Dim strStyle As String
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        mstrSQL = "Select 'infotable' As 图标,a.* From zlmip_table a Where Nvl(data_code,'-')=[1] And tab_type=[2]"
        
        strTemp = zlCommFun.GetCondition(rsCondition, "FilterText")
        If strTemp <> "" Then
            strStyle = zlCommFun.GetCondition(rsCondition, "FilterStyle")
            Select Case strStyle
            Case "编码"
                mstrSQL = mstrSQL & " And a.tab_code Like [3]"
            Case "名称"
                mstrSQL = mstrSQL & " And a.tab_title Like [3]"
            End Select
            strTemp = "%" & UCase(strTemp) & "%"
        End If
        
        Set TableRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "data_code"), zlCommFun.GetCondition(rsCondition, "tab_type"), strTemp)
            
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("ID")
        mstrSQL = "Select 'infotable' As 图标,a.* From zlmip_table a Where a.id=[1]"
        Set TableRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "id"))
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("SelectData")
        mstrSQL = "Select a.ID,a.tab_code As 编码,a.tab_title As 名称,a.tab_note As 说明 From zlmip_table a Where Nvl(a.data_code,'-')=[1]"
        Set TableRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "data_code"))
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("FilterData")
        
        strStyle = zlCommFun.GetCondition(rsCondition, "FilterStyle")
        strTemp = zlCommFun.GetCondition(rsCondition, "FilterText")
        
        Select Case strStyle
        Case ""
            mstrSQL = "Select a.ID,a.tab_code As 编码,a.tab_title As 名称,a.tab_note As 说明 From zlmip_table a Where (a.tab_code Like [1] Or a.tab_title Like [2])"
            Set TableRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strTemp & "%", "%" & strTemp & "%")
        Case "编码"
            mstrSQL = "Select a.ID,a.tab_code As 编码,a.tab_title As 名称,a.tab_note As 说明 From zlmip_table a Where a.tab_code Like [1]"
            Set TableRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strTemp & "%")
        Case "名称"
            mstrSQL = "Select a.ID,a.tab_code As 编码,a.tab_title As 名称,a.tab_note As 说明 From zlmip_table a Where a.tab_title Like [1]"
            Set TableRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, "%" & strTemp & "%")
        End Select
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function TableFieldRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取Infotable数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        mstrSQL = "Select a.* From zlmip_tab_field a Where a.tab_id=[1]"
        Set TableFieldRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "业务信息id"))
    '------------------------------------------------------------------------------------------------------------------
    Case "字段"
        
        If zlCommFun.GetCondition(rsCondition, "字段前缀") = "" Then
            mstrSQL = "Select a.fld_title As 字段名称,a.fld_type As 字段类型 From zlmip_tab_field a Where a.tab_id=[1]"
        Else
            mstrSQL = "Select '" & zlCommFun.GetCondition(rsCondition, "字段前缀") & "'||'.'||a.fld_title As 字段名称,Decode(a.fld_type,1,'数值',2,'字符',3,'日期') As 字段类型 From zlmip_tab_field a Where a.tab_id=[1]"
        End If
        
        Set TableFieldRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "业务信息id"))
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function TableParameterRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取Infotable数据
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        mstrSQL = "Select 'infotable' As 图标,tab_id,para_order,para_title,decode(para_type,2,'字符',1,'数值',3,'日期','自定义') as para_type,para_default,para_note  From zlmip_tab_parameter"
        Set TableParameterRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("tab_id")
        mstrSQL = "Select 'infotable' As 图标,tab_id,para_order,para_title,decode(para_type,2,'字符',1,'数值',3,'日期','自定义') as para_type,para_default,para_note  From zlmip_tab_parameter Where tab_id=[1]"
        Set TableParameterRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "id"))
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function TableRelationParameterRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取Infotable数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        
        mstrSQL = _
            "select x.para_title As 参数名称," & vbNewLine & _
            "       Decode(x.para_type,1,'数值',2,'字符',3,'日期') As 参数类型," & vbNewLine & _
            "       '=' As 赋值关系," & vbNewLine & _
            "       Decode(y.source_fld, null, null, '原表.' || y.source_fld) As 参数赋值" & vbNewLine & _
            "  from zlmip_tab_parameter x," & vbNewLine & _
            "       (select b.target_fld,b.source_fld " & vbNewLine & _
            "          from zlmip_tab_extend a, zlmip_tabext_condition b" & vbNewLine & _
            "         where a.id = [1] " & vbNewLine & _
            "           and a.id = b.ext_id) y" & vbNewLine & _
            " where x.tab_id = [2] " & vbNewLine & _
            "   and x.para_title = y.target_fld(+)" & vbNewLine & _
            " order by x.para_order"

        Set TableRelationParameterRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "id"), zlCommFun.GetCondition(rsCondition, "目标信息表id"))
        
'    '------------------------------------------------------------------------------------------------------------------
'    Case "字段"
'
'        If zlCommFun.GetCondition(rsCondition, "字段前缀") = "" Then
'            mstrSQL = "Select a.fld_title As 字段名称,a.fld_type As 字段类型 From zlmip_tab_field a Where a.tab_id=[1]"
'        Else
'            mstrSQL = "Select '" & zlCommFun.GetCondition(rsCondition, "字段前缀") & "'||'.'||a.fld_title As 字段名称,Decode(a.fld_type,1,'数值',2,'字符',3,'日期') As 字段类型 From zlmip_tab_field a Where a.tab_id=[1]"
'        End If
'
'        Set TableRelationParameterRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "业务信息id"))
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function TableEdit(Optional ByVal strUpdateMode As String = "", Optional ByVal rsPara As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：编辑Infotable数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngCount As Long
    Dim lngLoop As Long
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)
    
    Select Case UCase(strUpdateMode)
    '------------------------------------------------------------------------------------------------------------------
    Case "INSERT", "UPDATE"
        
        '信息表
        mstrSQL = "Table_Update('" & zlCommFun.GetParameter(rsPara, "id") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "data_code") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "tab_code") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "tab_title") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "tab_sqltext") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "tab_note") & "')"
        mstrSQL = "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
                
        '信息表SQL
        '--------------------------------------------------------------------------------------------------------------
        mstrSQL = "Tab_Parameter_Delete('" & zlCommFun.GetParameter(rsPara, "id") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
        lngCount = Val(zlCommFun.GetParameter(rsPara, "SQL参数个数"))
        For lngLoop = 1 To lngCount
            strTemp = Trim(zlCommFun.GetParameter(rsPara, "SQL参数_" & lngLoop))
            
            If strTemp <> "" Then
                mstrSQL = "Tab_Parameter_Insert('" & zlCommFun.GetParameter(rsPara, "id") & "','" & strTemp & "')"
                mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
                Call SQLRecordAdd(mrsSQL, mstrSQL)
            End If
        Next
        
        '信息表字段
        '--------------------------------------------------------------------------------------------------------------
        mstrSQL = "Tab_Field_Delete('" & zlCommFun.GetParameter(rsPara, "id") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)

        lngCount = Val(zlCommFun.GetParameter(rsPara, "SQL字段个数"))
        For lngLoop = 1 To lngCount
            strTemp = Trim(zlCommFun.GetParameter(rsPara, "SQL字段_" & lngLoop))

            If strTemp <> "" Then
                mstrSQL = "Tab_Field_Insert('" & zlCommFun.GetParameter(rsPara, "id") & "','" & strTemp & "')"
                mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
                Call SQLRecordAdd(mrsSQL, mstrSQL)
            End If
        Next
                
    '------------------------------------------------------------------------------------------------------------------
    Case "DELETE"
        mstrSQL = "Table_Delete('" & zlCommFun.GetParameter(rsPara, "id") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    End Select
        
    TableEdit = SQLRecordExecute(True)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function TableParameterEdit(Optional ByVal strUpdateMode As String = "", Optional ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：编辑Infotable数据
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)
    
    Select Case UCase(strUpdateMode)
    '------------------------------------------------------------------------------------------------------------------
    Case "INSERT", "UPDATE"
        mstrSQL = "Tab_Parameter_Update('" & zlCommFun.GetParameter(rsParameter, "tab_id") & "'," & _
                                            zlCommFun.GetParameter(rsParameter, "para_order") & "," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "para_title") & "'," & _
                                            zlCommFun.GetParameter(rsParameter, "para_type") & "," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "para_defalt") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "para_note") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "DELETE"
        mstrSQL = "Tab_Parameter_Delete('" & zlCommFun.GetParameter(rsParameter, "tab_id") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    End Select
        
    TableParameterEdit = SQLRecordExecute(False)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetTreeInfoTable(ByRef rsData As ADODB.Recordset, ByVal strDataKey As String, Optional ByVal strParentKey As String) As Boolean
    Dim rs As New ADODB.Recordset
    
    mstrSQL = "Select a.id,Decode(a.ext_type, 1, '引用', 2, '组合') As 类型,a.target_tab_id,a.ext_title,b.tab_title,c.para " & _
                "From zlmip_tab_extend a,zlmip_table b, " & _
                "       (select x.ext_id," & vbNewLine & _
                "               ltrim(f_List2str(Cast(Collect(' '||x.target_fld || ' = ' || Decode(x.source_fld,null,null,'原表.'||x.source_fld)||' ') As t_Strlist))) As para" & vbNewLine & _
                "          from zlmip_tabext_condition x, zlmip_tab_extend y" & vbNewLine & _
                "         where y.source_tab_id = [1] " & vbNewLine & _
                "           and x.ext_id = y.id" & vbNewLine & _
                "         group by x.ext_id) c" & vbNewLine & _
                "Where a.Source_tab_id = [1] And a.target_tab_id=b.id And c.ext_id=a.id Order By a.ext_order"
    Set rs = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strDataKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            rsData.AddNew
            
            rsData("id").Value = rs("id").Value
            rsData("parent_id").Value = strParentKey
            rsData("图标").Value = "infotable"
            rsData("命名").Value = rs("ext_title").Value
            rsData("目标").Value = rs("title").Value
            rsData("目标id").Value = rs("target_tab_id").Value
            rsData("类型").Value = rs("类型").Value
            rsData("标志").Value = IIf(strParentKey = "", 0, 1)
            rsData("参数").Value = rs("para").Value
            
            Call GetTreeInfoTable(rsData, rs("target_tab_id").Value, rs("id").Value)
                        
            rs.MoveNext
        Loop
    End If
        
    GetTreeInfoTable = True
    
End Function

Public Function TableRelationRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取Infotable数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTemp As String
    Dim rsData As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        Set rsData = New ADODB.Recordset
        With rsData
            .Fields.Append "id", adVarChar, 32
            .Fields.Append "parent_id", adVarChar, 32
            .Fields.Append "图标", adVarChar, 10
            .Fields.Append "命名", adVarChar, 30
            .Fields.Append "目标", adVarChar, 30
            .Fields.Append "目标id", adVarChar, 32
            .Fields.Append "类型", adVarChar, 10
            .Fields.Append "标志", adTinyInt
            .Fields.Append "参数", adVarChar, 1000
            .Open
        End With
        
        If GetTreeInfoTable(rsData, zlCommFun.GetCondition(rsCondition, "id"), "") Then
            If rsData.RecordCount > 0 Then rsData.MoveFirst
            Set TableRelationRead = rsData
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("Condition")
        
        mstrSQL = _
            "select x.ext_id,x.target_fld||' = '||'原表.'||x.source_fld As 条件" & vbNewLine & _
            "          from zlmip_tabext_condition x" & vbNewLine & _
            "         where x.ext_id = [1]"
                
        Set TableRelationRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "id"))
        
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("id")
                
        mstrSQL = _
            "Select 'infotable' As 图标," & vbNewLine & _
            "       a.id," & vbNewLine & _
            "       Decode(a.ext_type, 1, '引用', 2, '组合') As 类型," & vbNewLine & _
            "       a.ext_title As 命名," & vbNewLine & _
            "       b.tab_title As 目标," & vbNewLine & _
            "       a.target_tab_id as 目标id," & vbNewLine & _
            "       c.para As 参数" & vbNewLine & _
            "  From zlmip_tab_extend a," & vbNewLine & _
            "       zlmip_table b," & vbNewLine & _
            "       (select x.ext_id," & vbNewLine & _
            "               ltrim(f_List2str(Cast(Collect(' '||x.target_fld || ' = ' || Decode(x.source_fld,null,null,'原表.'||x.source_fld)||' ') As t_Strlist))) As para" & vbNewLine & _
            "          from zlmip_tabext_condition x" & vbNewLine & _
            "         where x.ext_id = [1] " & vbNewLine & _
            "         group by x.ext_id) c" & vbNewLine & _
            " Where a.id = [1] " & vbNewLine & _
            "   And a.target_tab_id = b.id" & vbNewLine & _
            "   And a.id = c.ext_id(+)"
                
        Set TableRelationRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "id"))
                
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CheckSQLPara(ByVal strSQL As String) As Boolean
    '******************************************************************************************************************
    '功能：编辑Infotable数据
    '参数：
    '返回：
    '******************************************************************************************************************
    '功能：检查SQL语句中参数符"[]"是否配对,以及参数号是否正确(非数字,不连续)
    
    Dim intLeft As Integer, intRight As Integer
    Dim intMin As Integer, intMax As Integer
    Dim strTmp As String, StrPar As String, strPars As String
    Dim i As Long
    
    For i = 1 To Len(strSQL)
        If Mid(strSQL, i, 1) = "[" Then intLeft = intLeft + 1
        If Mid(strSQL, i, 1) = "]" Then intRight = intRight + 1
    Next
    If intLeft <> intRight Then Exit Function '"["与"]"不配对
    
    If intLeft = 0 And intRight = 0 Then
        CheckSQLPara = True
        Exit Function
    End If
    
    strTmp = strSQL
    intMin = 32767
    Do While InStr(strTmp, "[") > 0
        strTmp = Mid(strTmp, InStr(strTmp, "[") + 1)
        StrPar = Left(strTmp, InStr(strTmp, "]") - 1)
        If Trim(StrPar) = "" Then
            StrPar = 0
        ElseIf Not IsNumeric(StrPar) Then
            Exit Function '非数字编号
        End If
        If CInt(StrPar) < intMin Then intMin = CInt(StrPar)
        If CInt(StrPar) > intMax Then intMax = CInt(StrPar)
        If InStr(strPars, "," & CInt(StrPar)) = 0 Then strPars = strPars & "," & CInt(StrPar)
    Loop
    If intMin <> 0 Then Exit Function '不是从0开始编号
    If strPars <> "" Then strPars = Mid(strPars, 2)
    If UBound(Split(strPars, ",")) <> intMax Then Exit Function '不是连续编号
    
    CheckSQLPara = True
End Function

Public Function GetSQLPara(ByVal strSQL As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：编辑Infotable数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTemp As String
    Dim lngPos1 As Long
    Dim lngPos2 As Long
    Dim intNo As Integer
    Dim rsSQLPara As ADODB.Recordset
    
    Set rsSQLPara = New ADODB.Recordset
    With rsSQLPara
        .Fields.Append "序号", adTinyInt
        .Open
    End With
    
    
    strTemp = strSQL
    
    lngPos1 = InStr(strTemp, "[")
    Do While lngPos1 > 0
        
        lngPos2 = InStr(strTemp, "]")
        intNo = Mid(strTemp, lngPos1 + 1, lngPos2 - lngPos1 - 1)
        
        rsSQLPara.AddNew
        rsSQLPara("序号").Value = intNo
                        
        strTemp = Mid(strTemp, lngPos2 + 1)
        lngPos1 = InStr(strTemp, "[")
        
    Loop
    
    If rsSQLPara.RecordCount > 0 Then rsSQLPara.MoveFirst
    
    Set GetSQLPara = rsSQLPara
    
End Function

Public Function GetSQLParaCount(strSQL As String) As Integer
    '功能：返回SQL语句中参数的个数,以序号为准
    Dim strTmp As String, StrPar As String, strPars As String
    
    strTmp = strSQL
    
    Do While InStr(strTmp, "[") > 0
        strTmp = Mid(strTmp, InStr(strTmp, "[") + 1)
        StrPar = Left(strTmp, InStr(strTmp, "]") - 1)
        If Trim(StrPar) = "" Then StrPar = 0
        If InStr(strPars, "," & CInt(StrPar)) = 0 Then strPars = strPars & "," & CInt(StrPar)
    Loop
    If strPars = "" Then
        GetSQLParaCount = 0
    Else
        strPars = Mid(strPars, 2)
        GetSQLParaCount = UBound(Split(strPars, ",")) + 1
    End If
End Function

Public Function CheckSQL(ByVal strSQL As String, ByVal rsSQLPara As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：编辑Infotable数据
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim strLeft As String
    Dim strRight As String
    Dim strFlag As String
    Dim StrPar As String
    Dim strErr As String
    Dim rsData As New ADODB.Recordset
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim intNo As Integer
    
    On Error GoTo errHand

    strFlag = strSQL
    
    intPos1 = InStr(strFlag, "[")
    Do While intPos1 > 0
        
        intPos2 = InStr(strFlag, "]")
        intNo = Mid(strFlag, intPos1 + 1, intPos2 - intPos1 - 1)
        
        strLeft = Mid(strFlag, 1, intPos1 - 1)
        strRight = Mid(strFlag, intPos2 + 1)
                
        rsSQLPara.Filter = ""
        rsSQLPara.Filter = "序号=" & intNo
        If rsSQLPara.RecordCount > 0 Then
            Select Case rsSQLPara("类型").Value
            Case "数值"
                strFlag = strLeft & "0" & strRight
            Case "字符"
                strFlag = strLeft & "'a'" & strRight
            Case "日期"
                strFlag = strLeft & "To_Date('2013-01-01 12:00:00','yyyy-mm-dd hh24:mi:ss')" & strRight
            End Select
            
        Else
            '没有定义，退出
            MsgBox "SQL参数有变化，需要同步调整参数定义！", vbInformation, mstrTitle
            Exit Function
        End If
        
        intPos1 = InStr(strFlag, "[")
    Loop
    
    strSQL = "Select 1 From (" & strFlag & ") Where 1=2"
    Set rsData = mclsDataOracle.OpenSQLRecord(strSQL, mstrTitle)
    
    CheckSQL = True

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description, vbInformation
'    Resume
End Function

Public Function GetSQLField(ByVal strSQL As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：编辑Infotable数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As ADODB.Recordset
    Dim strLeft As String
    Dim strRight As String
    Dim strFlag As String
    Dim StrPar As String
    Dim strErr As String
    Dim rsData As New ADODB.Recordset
    Dim intLoop As Integer
    
    On Error GoTo errHand

    strFlag = strSQL

    Do While InStr(strFlag, "[") > 0
        strLeft = Mid(strFlag, 1, InStr(strFlag, "[") - 1)
        strRight = Mid(strFlag, InStr(strFlag, "]") + 1)
        strFlag = strLeft & "Null" & strRight
    Loop
    
    
    Set rsData = New ADODB.Recordset
    With rsData
        .Fields.Append "序号", adTinyInt
        .Fields.Append "名称", adVarChar, 60
        .Fields.Append "类型", adTinyInt
        .Open
    End With
    
    strSQL = "Select * From (" & strFlag & ") Where 1=2"
    Set rs = mclsDataOracle.OpenSQLRecord(strSQL, mstrTitle)
    
    For intLoop = 0 To rs.Fields.Count - 1
        
        rsData.AddNew
        
        rsData("序号").Value = intLoop + 1
        rsData("名称").Value = rs.Fields(intLoop).Name
            
        Select Case rs.Fields(intLoop).Type
        Case adChar, adLongVarChar, adLongVarWChar, adVarChar, adLongVarWChar, 139, 202, 130
            rsData("类型").Value = 2
        Case adInteger, adDecimal, adDouble, adNumeric, adTinyInt, adSingle
            rsData("类型").Value = 1
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            rsData("类型").Value = 3
        End Select
    Next
            
    If rsData.RecordCount > 0 Then rsData.MoveFirst

    Set GetSQLField = rsData

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
'    MsgBox Err.Description, vbInformation
'    Resume
End Function

'Public Function ItemClassStruct() As ADODB.Recordset
'    '******************************************************************************************************************
'    '功能：读取Infotable结构
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'
'    On Error GoTo errHand
'
'    mstrSQL = "Select * From zlmip_folder Where 1=2"
'    Set ItemClassStruct = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'
'End Function

'Public Function ItemClassRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
'    '******************************************************************************************************************
'    '功能：读取zlmip_Folder数据
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'
'    On Error GoTo errHand
'
'    Select Case UCase(strConditionMode)
'    '------------------------------------------------------------------------------------------------------------------
'    Case ""
'
'        mstrSQL = _
'            "Select ID, 上级id, 编码, 名称, 图标, 选中图标" & vbNewLine & _
'            "From (Select ID, parent_id As 上级id, folder_code As 编码, folder_name As 名称, 'folder' As 图标, 'folder_open' As 选中图标" & vbNewLine & _
'            "       From zlmip_folder" & vbNewLine & _
'            "       Where Nvl(data_code,'-')=[1]" & vbNewLine & _
'            "       Start With parent_id Is Null" & vbNewLine & _
'            "       Connect By Prior id = parent_id)" & vbNewLine & _
'            "Order By 编码"
'
'
'        Set ItemClassRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "data_code"))
'    '------------------------------------------------------------------------------------------------------------------
'    Case UCase("ID")
'        If zlCommFun.GetCondition(rsCondition, "id") = "" Then
'            mstrSQL = _
'                "Select ID, 上级id, 编码, 名称, 图标, 选中图标" & vbNewLine & _
'                "From (Select ID, parent_id As 上级id, folder_code As 编码, folder_name As 名称, 'folder' As 图标, 'folder_open' As 选中图标" & vbNewLine & _
'                "       From zlmip_folder" & vbNewLine & _
'                "       Where  Nvl(data_code,'-')=[2]" & vbNewLine & _
'                "       Start With parent_id Is Null" & vbNewLine & _
'                "       Connect By Prior id = parent_id)" & vbNewLine & _
'                "Order By 编码"
'        Else
'            mstrSQL = _
'                "Select A.ID, A.parent_id As 上级id, A.folder_code As 编码, A.folder_name As 名称, B.folder_code As 上级编码, B.folder_name As 上级名称, 'folder' As 图标," & vbNewLine & _
'                "       'folder_open' As 选中图标" & vbNewLine & _
'                "From zlmip_folder A, zlmip_folder B" & vbNewLine & _
'                "Where A.parent_id = B.id(+) And A.id = [1]"
'        End If
'        Set ItemClassRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "id"), zlCommFun.GetCondition(rsCondition, "data_code"))
'    '------------------------------------------------------------------------------------------------------------------
'    Case UCase("SelectData")
'        '
'    '------------------------------------------------------------------------------------------------------------------
'    Case UCase("FilterData")
'        '
'    End Select
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    If zlComLib.ErrCenter = 1 Then
'        Resume
'    End If
'End Function

'Public Function GetClassDefaultCode(ByVal 上级编码_In As String, ByVal 最大编码_In As String, ByVal 最大编码长度_In As Long, Optional ByRef blnMsg As Boolean = False) As ADODB.Recordset
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '说明：
'    '******************************************************************************************************************
'    Dim v_缺省编码 As String
'    Dim n_调整编码 As Integer
'    Dim n_允许调整 As Integer
'    Dim n_允许编码长度 As Long
'    Dim n_允许编码最大长度 As Long
'    Dim v_Error As String
'
'    On Error GoTo errHand
'    If 最大编码长度_In = Len(上级编码_In) Then
'        ShowSimpleMsg "最大编码和编码长度已经达到最大限制，无法递增编码。"
'        v_缺省编码 = ""
'        n_调整编码 = 0
'        n_允许调整 = 0
'        blnMsg = True
'        Exit Function
'    End If
'
'    If 最大编码_In = "" Then
'        n_调整编码 = 1
'        v_缺省编码 = "01"
'        n_允许编码长度 = 最大编码长度_In - Len(上级编码_In)
'        n_允许编码最大长度 = n_允许编码长度
'        n_允许调整 = 0
'    Else
'        n_调整编码 = 0
'        n_允许编码长度 = Len(最大编码_In) - Len(上级编码_In)
'        n_允许编码最大长度 = 最大编码长度_In - Len(上级编码_In)
'        n_允许调整 = 1
'
'        If Mid(最大编码_In, Len(上级编码_In) + 1, 30) = Lpad("9", n_允许编码长度, 9) Then
'            If n_允许编码长度 >= 最大编码长度_In Then
'                ShowSimpleMsg "最大编码和编码长度已经达到最大限制，无法递增编码。"
'                n_调整编码 = 0
'                n_允许调整 = 0
'                v_缺省编码 = Lpad(" ", n_允许编码长度, " ")
'                blnMsg = True
'                Exit Function
'            Else
'                ShowSimpleMsg "最大编码已经达到本级限制，你可以扩充编码长度以满足需要。"
'                n_调整编码 = 1
'                n_允许调整 = 1
'                v_缺省编码 = "1" & Lpad("0", n_允许编码长度, "0")
'                n_允许编码长度 = n_允许编码长度 + 1
'                n_允许编码最大长度 = 最大编码长度_In - Len(上级编码_In)
'                blnMsg = True
'                Exit Function
'            End If
'        Else
'            v_缺省编码 = Lpad(CStr(Val(Mid(最大编码_In, Len(上级编码_In) + 1, 30)) + 1), n_允许编码长度, "0")
'        End If
'
'    End If
'
'    mstrSQL = _
'        "Select [1] As 缺省编码,[2] As 上级编码,[3] As 允许编码长度,[4] As 允许编码最大长度,[5] As 调整编码,[6] As 允许调整 From Dual"
'
'    Set GetClassDefaultCode = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, v_缺省编码, 上级编码_In, n_允许编码长度, n_允许编码最大长度, n_调整编码, n_允许调整)
'
'    Exit Function
'errHand:
'    If zlComLib.ErrCenter = 1 Then
'        Resume
'    End If
'End Function
'
'Public Function ItemClassMaxCode(ByVal strUpKey As String) As ADODB.Recordset
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '说明：
'    '******************************************************************************************************************
'    On Error GoTo errHand
'
'    mstrSQL = "Select Max(folder_code) As 最大编码 From zlmip_folder Where Nvl(parent_id, '-') = Nvl([1], '-')"
'    Set ItemClassMaxCode = zlDataBase.OpenSQLRecord(mstrSQL, mstrTitle, strUpKey)
'
'    Exit Function
'errHand:
'    If zlComLib.ErrCenter = 1 Then
'        Resume
'    End If
'End Function
'
'
'Public Function ItemClassEdit(Optional ByVal strUpdateMode As String = "", Optional ByVal rsPara As ADODB.Recordset) As Boolean
'    '******************************************************************************************************************
'    '功能：编辑Infotable数据
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim lngCount As Long
'    Dim lngLoop As Long
'    Dim strTemp As String
'
'    On Error GoTo errHand
'
'    Call SQLRecord(mrsSQL)
'
'    Select Case UCase(strUpdateMode)
'    '------------------------------------------------------------------------------------------------------------------
'    Case "INSERT"
'
'        '信息表
'        mstrSQL = "ItemClass_Insert('" & zlCommFun.GetParameter(rsPara, "id") & "'," & _
'                                            "'" & zlCommFun.GetParameter(rsPara, "parent_id") & "'," & _
'                                            "'" & zlCommFun.GetParameter(rsPara, "data_code") & "'," & _
'                                            "'" & zlCommFun.GetParameter(rsPara, "folder_code") & "'," & _
'                                            "'" & zlCommFun.GetParameter(rsPara, "folder_name") & "'," & _
'                                            zlCommFun.GetParameter(rsPara, "adjustcodelength") & ")"
'        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
'        Call SQLRecordAdd(mrsSQL, mstrSQL)
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "UPDATE"
'
'        '信息表
'        mstrSQL = "ItemClass_Update('" & zlCommFun.GetParameter(rsPara, "id") & "'," & _
'                                            "'" & zlCommFun.GetParameter(rsPara, "parent_id") & "'," & _
'                                            "'" & zlCommFun.GetParameter(rsPara, "data_code") & "'," & _
'                                            "'" & zlCommFun.GetParameter(rsPara, "folder_code") & "'," & _
'                                            "'" & zlCommFun.GetParameter(rsPara, "folder_name") & "'," & _
'                                            zlCommFun.GetParameter(rsPara, "adjustcodelength") & ")"
'        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
'        Call SQLRecordAdd(mrsSQL, mstrSQL)
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "DELETE"
'        mstrSQL = "ItemClass_Delete('" & zlCommFun.GetParameter(rsPara, "id") & "')"
'        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
'        Call SQLRecordAdd(mrsSQL, mstrSQL)
'    End Select
'
'    ItemClassEdit = SQLRecordExecute(True)
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    If zlComLib.ErrCenter = 1 Then
'        Resume
'    End If
'End Function

Public Function ItemStruct() As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取Event结构
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mstrSQL = "Select * From zlmip_item Where 1=2"
    Set ItemStruct = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function ItemRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strStyle As String
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        mstrSQL = "Select 'item' As 图标,a.* From zlmip_item a Order By a.code"
        Set ItemRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("item_type")
        
        mstrSQL = _
            "Select 'item' As 图标, a.Id, a.Item_Code, a.Item_Flow, a.Item_Title, a.Start_Date, Decode(Trunc(a.Stop_Date),To_Date('3000-01-01','yyyy-mm-dd'),Null,a.Stop_Date) As Stop_Date, a.Item_Note," & vbNewLine & _
            "       Decode(a.Trigger_Type, 1, '模块触发', '轮询触发') As Trigger_Type," & vbNewLine & _
            "       Decode(a.Again_Policy, 0, '无须重发', 1, '一直重发', 2, '重发' || a.Again_Para || '次', 3, a.Again_Para || '小时内一直重发') As Again_Policy_Title" & vbNewLine & _
            "From zlmip_Item A" & vbNewLine & _
            "Where Nvl(a.data_code,'-')=[2] And a.Item_Type=[1]"
                
        strTemp = zlCommFun.GetCondition(rsCondition, "FilterText")
        If strTemp <> "" Then
            strStyle = zlCommFun.GetCondition(rsCondition, "FilterStyle")
            Select Case strStyle
            Case "项目标识"
                mstrSQL = mstrSQL & " And a.Item_Code Like [3]"
            Case "项目名称"
                mstrSQL = mstrSQL & " And a.Item_Title Like [3]"
            End Select
            strTemp = "%" & UCase(strTemp) & "%"
        End If
                        
        Set ItemRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "item_type"), zlCommFun.GetCondition(rsCondition, "data_code"), strTemp)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("ID")
        mstrSQL = _
            "Select 'item' As 图标, a.Id, a.Item_Code, a.Item_Flow, a.Item_Title,a.tab_id, a.again_policy, a.trigger_frequency, a.check_frequency, a.check_freq_internal, a.Start_Date, a.Stop_Date, a.Item_Note," & vbNewLine & _
            "       Decode(a.Again_Policy, 0, '无须重发', 1, '一直重发', 2, '重发' || a.Again_Para || '次', 3, a.Again_Para || '小时内一直重发') As Again_Policy_Title," & vbNewLine & _
            "       c.Tab_Title, a.Trigger_Condition" & vbNewLine & _
            "From zlmip_Item A, zlmip_Table C" & vbNewLine & _
            "Where a.Id = [1] And a.Tab_Id = c.Id(+)"
        
        Set ItemRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "id"))
        
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function ItemFieldRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strStyle As String
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("item_id")
        mstrSQL = "Select a.fld_order,a.fld_title,a.fld_type From zlmip_item_field a Where a.item_id=[1] Order By a.fld_order"
        Set ItemFieldRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "item_id"))
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ItemFrequencyRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strStyle As String
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("item_id")
        mstrSQL = "Select * From zlmip_item_frequency a Where a.item_id=[1] Order by a.freq_order"
        Set ItemFrequencyRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "item_id"))

    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

'Public Function ItemDefaultCode(ByVal strClassKey As String, ByVal strClassCode As String, intMaxLength As Integer) As String
'    '******************************************************************************************************************
'    '功能：产生缺省编码
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim rs As New ADODB.Recordset
'    Dim str最大编码 As String
'    Dim str编码 As String
'
'    mstrSQL = "Select Max(item_code) As 最大编码 From zlmip_item Where Nvl(folder_id, '-') = Nvl([1], '-')"
'    Set rs = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strClassKey)
'    If rs.BOF = False Then
'        str最大编码 = Trim(zlCommFun.NVL(rs("最大编码").Value))
'    End If
'    If str最大编码 = "" Then
'        str编码 = strClassCode & "01"
'    ElseIf Len(str最大编码) > Len(strClassCode) Then
'        str最大编码 = Mid(str最大编码, Len(strClassCode) + 1)
'        str编码 = strClassCode & Format(Val(str最大编码) + 1, String(Len(str最大编码), "0"))
'    Else
'        str编码 = strClassCode & "01"
'    End If
'
'
'    If Len(str编码) <= intMaxLength Then
'        ItemDefaultCode = str编码
'    End If
'
'End Function

Public Function ItemEdit(Optional ByVal strUpdateMode As String = "", Optional ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：编辑zlmip_item数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngCount As Long
    Dim lngLoop As Long
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)
       
    
    Select Case UCase(strUpdateMode)
    '------------------------------------------------------------------------------------------------------------------
    Case "INSERT", "UPDATE"
        mstrSQL = "Item_Update('" & zlCommFun.GetParameter(rsParameter, "id") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "data_code") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "item_code") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "item_title") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "item_flow") & "'," & _
                                            Val(zlCommFun.GetParameter(rsParameter, "check_frequency")) & "," & _
                                            Val(zlCommFun.GetParameter(rsParameter, "check_freq_internal")) & "," & _
                                            Val(zlCommFun.GetParameter(rsParameter, "again_policy")) & "," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "again_para") & "'," & _
                                            "to_date('" & zlCommFun.GetParameter(rsParameter, "start_date") & "','yyyy-mm-dd')," & _
                                            "to_date('" & zlCommFun.GetParameter(rsParameter, "stop_date") & "','yyyy-mm-dd')," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "freq_content") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "trigger_condition") & "'," & _
                                            Val(zlCommFun.GetParameter(rsParameter, "trigger_frequency")) & "," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "item_note") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
        '触发条件结果字段
        '--------------------------------------------------------------------------------------------------------------
        mstrSQL = "Item_Field_Delete('" & zlCommFun.GetParameter(rsParameter, "id") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)

        lngCount = Val(zlCommFun.GetParameter(rsParameter, "SQL字段个数"))
        For lngLoop = 1 To lngCount
            strTemp = Trim(zlCommFun.GetParameter(rsParameter, "SQL字段_" & lngLoop))

            If strTemp <> "" Then
                mstrSQL = "Item_Field_Insert('" & zlCommFun.GetParameter(rsParameter, "id") & "','" & strTemp & "')"
                mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
                Call SQLRecordAdd(mrsSQL, mstrSQL)
            End If
        Next
        
    '------------------------------------------------------------------------------------------------------------------
    Case "DELETE"
        mstrSQL = "Item_Delete('" & zlCommFun.GetParameter(rsParameter, "id") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    End Select
        
    ItemEdit = SQLRecordExecute(False)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function ItemConfigEdit(Optional ByVal strUpdateMode As String = "", Optional ByVal rsPara As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim strLine As String
    Dim aryLine As Variant
    
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)
    
    Select Case UCase(strUpdateMode)
    '------------------------------------------------------------------------------------------------------------------
    Case "INSERT", "UPDATE"
            
        '消息
        mstrSQL = "Item_UpdateTab('" & zlCommFun.GetParameter(rsPara, "item_id") & "'," & _
                                            IIf(zlCommFun.GetParameter(rsPara, "tab_id") = "", "Null", "'" & zlCommFun.GetParameter(rsPara, "tab_id") & "'") & ")"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
                
        '配置
        mstrSQL = "Item_Config_Delete('" & zlCommFun.GetParameter(rsPara, "item_id") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
        lngCount = Val(zlCommFun.GetParameter(rsPara, "消息配置段数"))
        For lngLoop = 1 To lngCount
            strTemp = Trim(zlCommFun.GetParameter(rsPara, "消息配置_" & lngLoop))
            
            If strTemp <> "" Then
                mstrSQL = "Item_Config_Insert('" & zlCommFun.GetParameter(rsPara, "item_id") & "','" & strTemp & "')"
                mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
                Call SQLRecordAdd(mrsSQL, mstrSQL)
            End If
        Next

    End Select
        
    ItemConfigEdit = SQLRecordExecute(True)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
    
End Function


Public Function PersonRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strStyle As String
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        mstrSQL = _
            "Select ID,上级id,名称" & vbNewLine & _
            "From (Select 'G' || a.Id As ID, Null As 上级id, a.名称" & vbNewLine & _
            "       From 部门表 A" & vbNewLine & _
            "       Where Exists (Select 1 From 部门人员 B Where b.缺省 = 1 And b.部门id = a.Id)" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 'H' || a.Id As ID, 'G' || b.部门id As 上级id, a.姓名||'('||a.编号||')' As 名称" & vbNewLine & _
            "       From 人员表 A, 部门人员 B" & vbNewLine & _
            "       Where a.Id = b.人员id And b.缺省 = 1)" & vbNewLine & _
            "Start With 上级id Is Null" & vbNewLine & _
            "Connect By Prior ID = 上级id"
        
        Set PersonRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function DeptRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strStyle As String
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        mstrSQL = "Select a.id,上级id,a.名称||'('||a.编码||')' As 名称,a.简码,a.位置 from 部门表 a Where Nvl(a.撤档时间,To_Date('3000-01-01','yyyy-mm-dd'))<>To_Date('3000-01-01','yyyy-mm-dd')  start with a.上级id is null connect by prior a.id=a.上级id"
        Set DeptRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function PersonPropertyRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strStyle As String
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        mstrSQL = "Select a.名称 As ID,a.名称,a.简码 from 人员性质分类 a Order By a.编码"
        Set PersonPropertyRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function RoleRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：权限角色
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        'Granted_Role
        
        mstrSQL = _
            "Select ID, 上级id, 名称" & vbNewLine & _
            "From (Select Distinct 'G' || Decode(组名, Null, '未分组', 组名) As ID, Null As 上级id, Decode(组名, Null, '未分组', 组名) As 名称" & vbNewLine & _
            "       From Zlrolegroups" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 'H' || Substr(x.Granted_Role, 4) As ID, 'G' || Decode(y.组名, Null, '未分组', y.组名) As 上级id, Substr(x.Granted_Role, 4) As 名称" & vbNewLine & _
            "       From (Select a.Granted_Role" & vbNewLine & _
            "              From User_Role_Privs A" & vbNewLine & _
            "              Where a.Granted_Role Like 'ZL_%'" & vbNewLine & _
            "              Union" & vbNewLine & _
            "              Select a.Granted_Role" & vbNewLine & _
            "              From Dba_Role_Privs A" & vbNewLine & _
            "              Where a.Granted_Role Like 'ZL_%' And a.Grantee = 'SYSTEM' And Not Exists" & vbNewLine & _
            "               (Select 1 From Dba_Role_Privs X Where x.Granted_Role = a.Granted_Role And Grantee = 'ZLHIS')) X," & vbNewLine & _
            "            Zlrolegroups Y" & vbNewLine & _
            "       Where x.Granted_Role = y.角色(+))" & vbNewLine & _
            "Start With 上级id Is Null" & vbNewLine & _
            "Connect By Prior ID = 上级id"

        
        Set RoleRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function StationRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：权限角色
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        mstrSQL = _
            "Select ID, 上级id, 名称" & vbNewLine & _
            "From (Select Distinct 'G' || Decode(a.部门, Null, '未知部门', a.部门) As ID, Null As 上级id, Decode(a.部门, Null, '未知部门', a.部门) As 名称" & vbNewLine & _
            "       From zlClients A" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 'H' || a.工作站 || '(' || a.IP || ')' As ID, 'G' || Decode(a.部门, Null, '未知部门', a.部门) As 上级id, a.工作站 || '(' || a.IP || ')' As 名称" & vbNewLine & _
            "       From zlClients A)" & vbNewLine & _
            "Start With 上级id Is Null" & vbNewLine & _
            "Connect By Prior ID = 上级id"
                
        Set StationRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ModuleRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：权限角色
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        mstrSQL = _
            "Select *" & vbNewLine & _
            "From (Select 'G'||编号 As ID, Null As 上级id, 编号||'-'||名称 As 名称" & vbNewLine & _
            "       From zlSystems" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 'H'||b.系统||'_'||b.序号||'_'||a.功能 As ID, 'G'||b.系统 As 上级id, b.序号||'-'||b.标题||Decode(a.功能, '基本', '', '('||a.功能||')') As 名称" & vbNewLine & _
            "       From zlProgFuncs A, zlPrograms B" & vbNewLine & _
            "       Where a.系统 = b.系统 And a.序号 = b.序号)" & vbNewLine & _
            "Start With 上级id Is Null" & vbNewLine & _
            "Connect By Prior ID = 上级id"
                        
        Set ModuleRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ItemDeliverStruct() As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取Event结构
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mstrSQL = "Select deliver_title From zlmip_item_deliver Where 1=2"
    Set ItemDeliverStruct = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function ItemDeliverRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strStyle As String
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        mstrSQL = "Select 'event' As 图标,a.* From zlmip_item a Order By a.code"
        Set ItemDeliverRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("item_id")
        mstrSQL = "Select 'event' As 图标,a.id,a.item_id,a.deliver_order,a.deliver_title,a.deliver_code,a.deliver_object.getStringVal() As deliver_object From zlmip_item_deliver a Where a.item_id=[1]"
        Set ItemDeliverRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "item_id"))
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("ID")
        mstrSQL = "Select 'event' As 图标,a.id,a.item_id,a.deliver_order,a.deliver_title,a.deliver_code,a.deliver_object.getStringVal() As deliver_object From zlmip_item_deliver a Where a.id=[1]"
        Set ItemDeliverRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "id"))

    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function ItemDeliverEdit(Optional ByVal strUpdateMode As String = "", Optional ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：编辑zlmip_item数据
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)
        
    Select Case UCase(strUpdateMode)
    '------------------------------------------------------------------------------------------------------------------
    Case "INSERT", "UPDATE"
        mstrSQL = "Item_Deliver_Update('" & zlCommFun.GetParameter(rsParameter, "id") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "item_id") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "deliver_title") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "deliver_code") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "deliver_object") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "DELETE"
        mstrSQL = "Item_Deliver_Delete('" & zlCommFun.GetParameter(rsParameter, "id") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    End Select
        
    ItemDeliverEdit = SQLRecordExecute(False)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ClientRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strStyle As String
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        mstrSQL = "Select 0 As 选择,'station' As 图标,工作站 As ID,a.工作站,a.IP,a.操作系统,a.部门,a.说明,Decode(b.mipuser,Null,'<默认>',b.mipuser) As 消息用户,b.mipuserpassword as 消息密码 From zlClients a,zlmip_clientuser b Where a.工作站=b.station(+) Order By a.工作站"
        Set ClientRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("FilterData")
        
        strStyle = zlCommFun.GetCondition(rsCondition, "FilterStyle")
        strTemp = zlCommFun.GetCondition(rsCondition, "FilterText")
        
        Select Case strStyle
        Case "工作站"
            mstrSQL = "Select 0 As 选择,'station' As 图标,工作站 As ID,a.工作站,a.IP,a.操作系统,a.部门,a.说明,Decode(b.mipuser,Null,'<默认>',b.mipuser) As 消息用户,b.mipuserpassword as 消息密码 From zlClients a,zlmip_clientuser b Where a.工作站=b.station(+) and a.工作站 Like [1] Order By a.工作站"
            Set ClientRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, "%" & strTemp & "%")
        Case "IP"
            mstrSQL = "Select 0 As 选择,'station' As 图标,工作站 As ID,a.工作站,a.IP,a.操作系统,a.部门,a.说明,Decode(b.mipuser,Null,'<默认>',b.mipuser) As 消息用户,b.mipuserpassword as 消息密码 From zlClients a,zlmip_clientuser b Where a.工作站=b.station(+) and a.IP Like [1] Order By a.工作站"
            Set ClientRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, "%" & strTemp & "%")
        Case "部门"
            mstrSQL = "Select 0 As 选择,'station' As 图标,工作站 As ID,a.工作站,a.IP,a.操作系统,a.部门,a.说明,Decode(b.mipuser,Null,'<默认>',b.mipuser) As 消息用户,b.mipuserpassword as 消息密码 From zlClients a,zlmip_clientuser b Where a.工作站=b.station(+) and a.部门 Like [1] Order By a.工作站"
            Set ClientRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, "%" & strTemp & "%")
        End Select
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function ClientsEdit(ByVal strEditMode As String, Optional ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：编辑zlmip_item数据
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)
            
    Select Case UCase(strEditMode)
    '------------------------------------------------------------------------------------------------------------------
    Case "UPDATE"
        mstrSQL = "Clients_Update('" & zlCommFun.GetParameter(rsParameter, "工作站") & "','" & zlCommFun.GetParameter(rsParameter, "消息用户") & "','" & zlCommFun.GetParameter(rsParameter, "消息密码") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "COPY"
        mstrSQL = "Clients_Copy('" & zlCommFun.GetParameter(rsParameter, "来源工作站") & "','" & zlCommFun.GetParameter(rsParameter, "目标工作站") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    End Select
        
    ClientsEdit = SQLRecordExecute(False)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ParameterRead(ByVal intParameter As Integer) As String
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    
    mstrSQL = "Select 参数值 From zlOptions WHERE 参数号=[1] "
    Set rsTmp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, intParameter)
    If rsTmp.BOF = False Then
        ParameterRead = zlCommFun.NVL(rsTmp("参数值").Value)
    End If
End Function

Public Function ParameterEdit(ByVal strEditMode As String, Optional ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：编辑zlmip_item数据
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)
            
    Select Case UCase(strEditMode)
    '------------------------------------------------------------------------------------------------------------------
    Case "INSERT", "UPDATE"
        mstrSQL = "Options_Update(" & zlCommFun.GetParameter(rsParameter, "参数号") & ",'" & zlCommFun.GetParameter(rsParameter, "参数名") & "','" & zlCommFun.GetParameter(rsParameter, "参数值") & "')"
        mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    End Select
        
    ParameterEdit = SQLRecordExecute(False)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SendLogRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strStyle As String
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("send_content")
        mstrSQL = "Select a.send_content.getstringval() As send_content From zlmip_send_log a Where a.id=[1]"
        Set SendLogRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "send_content"))
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("FilterData")
        
        mstrSQL = _
            "Select 0 As 选择, a.Id, a.Item_Identity, a.Send_Date, a.Send_Again_Date, a.send_mipuser,Decode(Nvl(a.send_again,0),0,'',1,'自动重发') As send_again," & vbNewLine & _
            "       a.Send_User || '(' || a.Send_Username || ')' As Send_User, a.Send_Station,a.Send_Module,Decode(a.Send_Again_Number,0,Null,a.Send_Again_Number) As Send_Again_Number,Decode(a.send_state,1,'发送成功',2,'发送失败') As send_state,a.send_note" & vbNewLine & _
            "From zlmip_Send_Log A Where a.Send_Date Between [1] And [2] "
                    
        strStyle = zlCommFun.GetCondition(rsCondition, "FilterStyle")
        strTemp = zlCommFun.GetCondition(rsCondition, "FilterText")
        Select Case strStyle
        Case "发送站点"
            mstrSQL = mstrSQL & " And a.Send_Station Like [3]"
        Case "发送模块"
            mstrSQL = mstrSQL & " And a.Send_Module Like [3]"
        Case "消息用户"
            mstrSQL = mstrSQL & " And a.send_mipuser Like [3]"
        Case "产品用户"
            mstrSQL = mstrSQL & " And (a.Send_User Like [3] Or a.Send_Username Like [3])"
        Case "发送消息"
            mstrSQL = mstrSQL & " And a.Item_Identity Like [3]"
        End Select
        
        mstrSQL = mstrSQL & " Order By a.Send_Date Desc"
        Set SendLogRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, CDate(zlCommFun.GetCondition(rsCondition, "Start_Date") & " 00:00:00"), CDate(zlCommFun.GetCondition(rsCondition, "End_Date") & " 23:59:59"), "%" & strTemp & "%")
        
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function SendLogEdit(ByVal strEditMode As String, Optional ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：编辑zlmip_item数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngCount As Long
    Dim lngLoop As Long
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)
            
    Select Case UCase(strEditMode)
    '------------------------------------------------------------------------------------------------------------------
    Case "DELETE"
    
        '配置
        
        lngCount = Val(zlCommFun.GetParameter(rsParameter, "SendLog_SegmentCount"))
        For lngLoop = 1 To lngCount
            strTemp = Trim(zlCommFun.GetParameter(rsParameter, "SendLog_" & lngLoop))
            
            If strTemp <> "" Then
                mstrSQL = "Send_Log_Delete('" & strTemp & "')"
                mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
                Call SQLRecordAdd(mrsSQL, mstrSQL)
            End If
        Next
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("Repeat")
        
        lngCount = Val(zlCommFun.GetParameter(rsParameter, "SendLog_SegmentCount"))
        For lngLoop = 1 To lngCount
            strTemp = Trim(zlCommFun.GetParameter(rsParameter, "SendLog_" & lngLoop))
            
            If strTemp <> "" Then
                mstrSQL = "Send_Log_Repeat('" & strTemp & "')"
                mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
                Call SQLRecordAdd(mrsSQL, mstrSQL)
            End If
        Next
        
    End Select
        
    SendLogEdit = SQLRecordExecute(False)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetSendLogContent(ByVal strKey As String) As String
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strStyle As String
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim intLoop As Integer
    
    On Error GoTo errHand
    
    mstrSQL = "Select length(a.send_content.getclobval()) As ContentLength From zlmip_send_log a Where a.id=[1]"
    Set rsTemp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strKey)
    If rsTemp.BOF = False Then
        If rsTemp("ContentLength").Value > 2000 Then
            For intLoop = 1 To rsTemp("ContentLength").Value / 2000 + 1
                
                mstrSQL = "select To_Char(substr(a.send_content.getclobval()," & CDbl(intLoop) * 2000 - 1999 & ",2000)) as send_content from zlmip_send_log a where a.id = [1] "
                Set rsTemp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strKey)
                If rsTemp.BOF = False Then
                    GetSendLogContent = GetSendLogContent & zlCommFun.NVL(rsTemp("send_content").Value)
                End If
            Next
        Else
            mstrSQL = "Select a.send_content.getstringval() As send_content From zlmip_send_log a Where a.id=[1]"
            Set rsTemp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strKey)
            If rsTemp.BOF = False Then
                GetSendLogContent = zlCommFun.NVL(rsTemp("send_content").Value)
            End If
        End If
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
'    MsgBox Err.Description
'    Resume
End Function

Public Function GetReceiveLogContent(ByVal strKey As String) As String
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strStyle As String
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim intLoop As Integer
    
    On Error GoTo errHand
    
    mstrSQL = "Select length(a.receive_content.getclobval()) As ContentLength From zlmip_receive_log a Where a.id=[1]"
    Set rsTemp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strKey)
    If rsTemp.BOF = False Then
        If rsTemp("ContentLength").Value > 2000 Then
            For intLoop = 1 To rsTemp("ContentLength").Value / 2000 + 1
                
                mstrSQL = "select To_Char(substr(a.receive_content.getclobval()," & CDbl(intLoop) * 2000 - 1999 & ",2000)) as receive_content from zlmip_receive_log a where a.id = [1] "
                Set rsTemp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strKey)
                If rsTemp.BOF = False Then
                    GetReceiveLogContent = GetReceiveLogContent & zlCommFun.NVL(rsTemp("receive_content").Value)
                End If
            Next
        Else
            mstrSQL = "Select a.receive_content.getstringval() As receive_content From zlmip_receive_log a Where a.id=[1]"
            Set rsTemp = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strKey)
            If rsTemp.BOF = False Then
                GetReceiveLogContent = zlCommFun.NVL(rsTemp("receive_content").Value)
            End If
        End If
    End If

    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
'    MsgBox Err.Description
'    Resume
End Function

Public Function ReceiveLogRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strStyle As String
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("id")
        mstrSQL = "Select a.receive_content.getstringval() As receive_content From zlmip_receive_log a Where a.id=[1]"
        Set ReceiveLogRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "id"))
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("FilterData")
        
        mstrSQL = _
            "Select 0 As 选择, a.Id, a.Item_Identity, a.receive_Date, a.receive_mipuser,a.receive_module," & vbNewLine & _
            "       a.receive_User || '(' || a.receive_Username || ')' As receive_User, a.receive_Station,a.receive_note" & vbNewLine & _
            "From zlmip_receive_log A Where a.receive_date Between [1] And [2] "
                    
        strStyle = zlCommFun.GetCondition(rsCondition, "FilterStyle")
        strTemp = zlCommFun.GetCondition(rsCondition, "FilterText")
        Select Case strStyle
        Case "接收站点"
            mstrSQL = mstrSQL & " And a.receive_station Like [3]"
        Case "接收模块"
            mstrSQL = mstrSQL & " And a.receive_module Like [3]"
        Case "消息用户"
            mstrSQL = mstrSQL & " And a.receive_mipuser Like [3]"
        Case "产品用户"
            mstrSQL = mstrSQL & " And (a.receive_user Like [3] Or a.receive_username Like [3])"
        Case "接收消息"
            mstrSQL = mstrSQL & " And a.Item_Identity Like [3]"
        End Select
        
        mstrSQL = mstrSQL & " Order By a.receive_date Desc"
        Set ReceiveLogRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, CDate(zlCommFun.GetCondition(rsCondition, "Start_Date") & " 00:00:00"), CDate(zlCommFun.GetCondition(rsCondition, "End_Date") & " 23:59:59"), "%" & strTemp & "%")
        
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function ReceiveLogEdit(ByVal strEditMode As String, Optional ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：编辑zlmip_item数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngCount As Long
    Dim lngLoop As Long
    Dim strTemp As String
    
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)
            
    Select Case UCase(strEditMode)
    '------------------------------------------------------------------------------------------------------------------
    Case "DELETE"
    
        '配置
        
        lngCount = Val(zlCommFun.GetParameter(rsParameter, "SendLog_SegmentCount"))
        For lngLoop = 1 To lngCount
            strTemp = Trim(zlCommFun.GetParameter(rsParameter, "SendLog_" & lngLoop))
            
            If strTemp <> "" Then
                mstrSQL = "Receive_Log_Delete('" & strTemp & "')"
                mstrSQL = mstrOwner & "zlMip_PkgManage." & mstrSQL
                Call SQLRecordAdd(mrsSQL, mstrSQL)
            End If
        Next
        
    End Select
        
    ReceiveLogEdit = SQLRecordExecute(False)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function AnalyseCompent() As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mstrSQL = "Select Distinct Upper(g.部件) As 部件" & vbCrLf & _
            " From zlPrograms g, zlRegFunc r" & vbCrLf & _
            " Where g.序号 = r.序号 And Trunc(g.系统 / 100) = r.系统 And Upper(g.部件) <> 'ZL9REPORT'"
    
    Set AnalyseCompent = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
End Function

Public Function GetGrantMenu(ByVal blnRunningInIDE As Boolean, ByVal gstrMenuSys As String, ByVal strSystems As String, ByVal strObjs As String, ByVal strSys As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    If blnRunningInIDE Then
        mstrSQL = "Select 层次, ID As 编号, Nvl(上级id, 0) As 上级, 标题, Decode(Nvl(短标题,'空'),'空',标题,短标题) as 短标题, 快键, 说明, Nvl(模块, 0) As 模块, Nvl(系统, 0) As 系统, " & _
                 "        Nvl(图标, 0) As 图标, 部件, Decode(Upper(RTrim(部件)), 'ZL9REPORT', 1, 0) As 报表 " & _
                 " From Table(Cast(ZLTOOLS.f_Reg_Menu('" & gstrMenuSys & "', " & strSystems & ", " & strObjs & ") As ZLTOOLS.t_Menu_Rowset)) " & _
                 " Union " & _
                 " Select A.层次, A.ID, Nvl(上级id, 0) As 上级, A.标题, Decode(Nvl(A.短标题,'空'),'空',A.标题,A.短标题) As 短标题, A.快键, A.说明, Nvl(A.模块, 0) As 模块, " & _
                 "        Nvl(A.系统, 0) As 系统, Nvl(图标, 0) As 图标, C.部件, Decode(C.部件, 'ZL9REPORT', 1, 0) As 报表 " & _
                 " From (Select Level As 层次, ID, 上级id, 标题, 短标题, 快键, 说明, 模块, 系统, 图标 " & _
                 "        From zlMenus " & _
                 "        Where 组别 = '" & gstrMenuSys & "' And Nvl(系统, 0) IN(" & strSys & ") " & _
                 "        Start With 上级id Is Null " & _
                 "        Connect By Prior ID = 上级id) A, " & _
                 "      (Select 系统, 模块 " & _
                 "        From zlMenus A " & _
                 "        Where 组别 = '" & gstrMenuSys & "' And Nvl(系统, 0) IN (" & strSys & ") " & _
                 "        Minus " & _
                 "        Select 系统 * 100, 序号 From Zlregfunc Where 系统 * 100 IN (" & strSys & ")) B," & _
                 "      (select Upper(RTrim(部件)) as 部件,序号 From zlPrograms ) C " & _
                 " Where A.系统 = B.系统 And Nvl(A.模块, '0') = Nvl(B.模块, '0') And Nvl(A.模块, 0) = C.序号(+) " & _
                 " Order By 层次, 报表, 系统, 模块, 编号"

    Else
        mstrSQL = "SELECT 层次, Id AS 编号, Nvl(上级id, 0) AS 上级, 标题, Decode(Nvl(短标题,'空'),'空',标题,短标题) As 短标题, 快键, 说明, Nvl(模块, 0) AS 模块, Nvl(系统, 0) AS 系统, " & _
                 "        Nvl(图标, 0) AS 图标, 部件, Decode(Upper(Rtrim(部件)), 'ZL9REPORT', 1, 0) AS 报表 " & _
                 " FROM TABLE(CAST(Zltools.f_Reg_Menu('" & gstrMenuSys & "', " & strSystems & ", " & strObjs & ") As " & _
                 " Zltools.t_Menu_Rowset)) " & _
                 " Order By 层次, 报表, 系统, 模块, 编号"
    End If
    
    Set GetGrantMenu = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
End Function

Public Function GetSystemInfo(ByVal strSystems As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    '读用户信息赋予公共，便于其他程序使用

    mstrSQL = "Select S.*" & _
            " From zlSystems S,(Select Distinct owner From All_Tables Where Table_Name='部门表') D" & _
            " Where Upper(S.所有者)=D.Owner And S.编号 In (" & strSystems & ") Order by S.编号"
    
    Set GetSystemInfo = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
End Function

Public Function GetUserInfo(ByVal strOwner As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mstrSQL = "Select R.*,D.编码 as 部门编码,D.名称 as 部门名称,P.编号,P.姓名,P.简码" & _
            " From " & strOwner & ".上机人员表 U," & strOwner & ".人员表 P," & strOwner & ".部门表 D," & strOwner & ".部门人员 R" & _
            " Where U.人员ID = P.ID And R.部门ID = D.ID And P.ID=R.人员ID and U.用户名=USER And (P.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.撤档时间 Is Null) and R.缺省=1"
    
    Set GetUserInfo = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
End Function

Public Function GetProgFuncsCount() As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mstrSQL = "   Select Count(*) Records From zlprogfuncs " & _
             "   Where 系统 Is Null" & _
             "   And 序号 in (Select 序号 From zlRoleGrant G,session_roles S Where G.角色=S.Role)"
                             
    Set GetProgFuncsCount = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
End Function

Public Function GetAccount(ByVal StrComponent As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mstrSQL = "SELECT Substr(Lpad(编号, 5, '0'), 4) 编号, 编号 系统, 名称 " & _
             " FROM Zlsystems " & _
             " WHERE 编号 IN " & _
             "       (SELECT DISTINCT p.系统 " & _
             "        FROM Zlprograms p, " & _
             "             (SELECT 系统, 序号 " & _
             "               FROM (SELECT DISTINCT p.系统, p.序号, r.功能 AS 授权 " & _
             "                      FROM Zlprogfuncs p, Zlregfunc r " & _
             "                      WHERE Trunc(p.系统 / 100) = r.系统(+) AND p.序号 = r.序号(+) AND p.功能 = r.功能(+) AND " & _
             "                            (EXISTS (SELECT 1 FROM Session_Roles WHERE Role = 'DBA') OR " & _
             "                             p.系统 IN (SELECT 编号 FROM Zlsystems WHERE Upper(所有者) = USER) OR " & _
             "                             p.系统 IN (SELECT 系统 FROM Zlrolegrant g, Session_Roles s WHERE g.角色 = s.Role)) " & _
             "                      MINUS " & _
             "                      SELECT DISTINCT s.系统, s.序号, r.功能 AS 授权 " & _
             "                      FROM Zlprogprivs s, Zlregfunc r " & _
             "                      WHERE Trunc(s.系统 / 100) = r.系统(+) AND s.序号 = r.序号(+) AND s.功能 = r.功能(+) AND " & _
             "                            (EXISTS (SELECT 1 FROM Session_Roles WHERE Role = 'DBA') OR " & _
             "                             s.系统 IN (SELECT 编号 FROM Zlsystems WHERE Upper(所有者) = USER) OR " & _
             "                             s.系统 IN (SELECT 系统 FROM Zlrolegrant g, Session_Roles s WHERE g.角色 = s.Role)) AND " & _
             "                            s.所有者 <> USER AND s.对象 IN (SELECT Object_Name " & _
             "                                                            FROM User_Objects " & _
             "                                                            WHERE Object_Type IN ('SEQUENCE', 'TABLE', 'VIEW', 'PROCEDURE', 'FUNCTION', 'PACAKEG'))) " & _
             "               WHERE 授权 IS NULL AND 系统 IS NULL OR 授权 IS NOT NULL) f"
    mstrSQL = mstrSQL & "       WHERE p.系统 = f.系统 AND p.序号 = f.序号 AND " & _
             "              Upper(p.部件) IN (" & StrComponent & ")) " & _
             " ORDER BY 名称, 编号"
        
    Set GetAccount = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function

Public Function GetDeliveryTree(ByVal strDeliveobject As String, ByRef rsDeliver As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：读取投递目标数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim objXML As New clsMessageXML
    Dim strTitle As String
    Dim strKey As String
    Dim strNodeName As String
    Dim strSys As String
    Dim strPrivilige As String
    
    On Error GoTo errHand
                                    
    Call objXML.LoadString(strDeliveobject)
    
    '产品(人员)
    '--------------------------------------------------------------------------------------------------------------
    Set rsTmp = objXML.GetNodeData("//persons/person")
    If rsTmp.BOF = False Then
        strTitle = ""
        strKey = ""
        Do While Not rsTmp.EOF
            strNodeName = zlCommFun.NVL(rsTmp("nodename").Value)
            If strNodeName = "person" Then
                If strTitle <> "" Then Call InsertDeliveryTree("人员", strKey, strTitle, rsDeliver)
                strTitle = ""
                strKey = ""
            ElseIf strNodeName = "title" Then
                strTitle = zlCommFun.NVL(rsTmp("nodetext").Value)
            ElseIf strNodeName = "key" Then
                strKey = "产品(人员)_" & zlCommFun.NVL(rsTmp("nodetext").Value)
            End If
            
            rsTmp.MoveNext
        Loop
        If strTitle <> "" Then Call InsertDeliveryTree("人员", strKey, strTitle, rsDeliver)
    End If
    
    '消息(人员)
    '--------------------------------------------------------------------------------------------------------------
    Set rsTmp = objXML.GetNodeData("//persons/message")
    If rsTmp.BOF = False Then
        strTitle = ""
        strKey = ""
        Do While Not rsTmp.EOF
            strNodeName = zlCommFun.NVL(rsTmp("nodename").Value)
            If strNodeName = "message" Then
                strTitle = zlCommFun.NVL(rsTmp("nodetext").Value)
                strKey = "消息(人员)_" & zlCommFun.NVL(rsTmp("nodetext").Value)
                Call InsertDeliveryTree("人员", strKey, strTitle, rsDeliver)
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    
    '产品(部门)
    '--------------------------------------------------------------------------------------------------------------
    Set rsTmp = objXML.GetNodeData("//depts/dept")
    If rsTmp.BOF = False Then
        strTitle = ""
        strKey = ""
        Do While Not rsTmp.EOF
            strNodeName = zlCommFun.NVL(rsTmp("nodename").Value)
            If strNodeName = "dept" Then
                If strTitle <> "" Then Call InsertDeliveryTree("部门", strKey, strTitle, rsDeliver)
                strTitle = ""
                strKey = ""
            ElseIf strNodeName = "title" Then
                strTitle = zlCommFun.NVL(rsTmp("nodetext").Value)
            ElseIf strNodeName = "key" Then
                strKey = "产品(部门)_" & zlCommFun.NVL(rsTmp("nodetext").Value)
            End If
            
            rsTmp.MoveNext
        Loop
        If strTitle <> "" Then Call InsertDeliveryTree("部门", strKey, strTitle, rsDeliver)
    End If
    
    '消息(部门)
    '--------------------------------------------------------------------------------------------------------------
    Set rsTmp = objXML.GetNodeData("//depts/message")
    If rsTmp.BOF = False Then
        strTitle = ""
        strKey = ""
        Do While Not rsTmp.EOF
            strNodeName = zlCommFun.NVL(rsTmp("nodename").Value)
            If strNodeName = "message" Then
                strTitle = zlCommFun.NVL(rsTmp("nodetext").Value)
                strKey = "消息(部门)_" & zlCommFun.NVL(rsTmp("nodetext").Value)
                Call InsertDeliveryTree("部门", strKey, strTitle, rsDeliver)
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    '产品(性质)
    '--------------------------------------------------------------------------------------------------------------
    Set rsTmp = objXML.GetNodeData("//personpropertys/personproperty")
    If rsTmp.BOF = False Then
        strTitle = ""
        strKey = ""
        Do While Not rsTmp.EOF
            strNodeName = zlCommFun.NVL(rsTmp("nodename").Value)
            If strNodeName = "personproperty" Then
                strTitle = zlCommFun.NVL(rsTmp("nodetext").Value)
                strKey = "产品(性质)_" & zlCommFun.NVL(rsTmp("nodetext").Value)
                Call InsertDeliveryTree("人员性质", strKey, strTitle, rsDeliver)
                
            End If
            
            rsTmp.MoveNext
        Loop
    End If
    
    
    '消息(性质)
    '--------------------------------------------------------------------------------------------------------------
    Set rsTmp = objXML.GetNodeData("//personpropertys/message")
    If rsTmp.BOF = False Then
        strTitle = ""
        strKey = ""
        Do While Not rsTmp.EOF
            strNodeName = zlCommFun.NVL(rsTmp("nodename").Value)
            If strNodeName = "message" Then
                strTitle = zlCommFun.NVL(rsTmp("nodetext").Value)
                strKey = "消息(性质)_" & zlCommFun.NVL(rsTmp("nodetext").Value)
                Call InsertDeliveryTree("人员性质", strKey, strTitle, rsDeliver)
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    '产品(角色)
    '--------------------------------------------------------------------------------------------------------------
    Set rsTmp = objXML.GetNodeData("//roles/role")
    If rsTmp.BOF = False Then
        strTitle = ""
        strKey = ""
        Do While Not rsTmp.EOF
            strNodeName = zlCommFun.NVL(rsTmp("nodename").Value)
            If strNodeName = "role" Then
                
                strTitle = zlCommFun.NVL(rsTmp("nodetext").Value)
                strKey = "产品(角色)_" & zlCommFun.NVL(rsTmp("nodetext").Value)
                Call InsertDeliveryTree("角色", strKey, strTitle, rsDeliver)
                
            End If
            
            rsTmp.MoveNext
        Loop
    End If
    
    '消息(角色)
    '--------------------------------------------------------------------------------------------------------------
    Set rsTmp = objXML.GetNodeData("//roles/message")
    If rsTmp.BOF = False Then
        strTitle = ""
        strKey = ""
        Do While Not rsTmp.EOF
            strNodeName = zlCommFun.NVL(rsTmp("nodename").Value)
            If strNodeName = "message" Then
                strTitle = zlCommFun.NVL(rsTmp("nodetext").Value)
                strKey = "消息(角色)_" & zlCommFun.NVL(rsTmp("nodetext").Value)
                Call InsertDeliveryTree("角色", strKey, strTitle, rsDeliver)
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    '产品(站点)
    '--------------------------------------------------------------------------------------------------------------
    Set rsTmp = objXML.GetNodeData("//stations/station")
    If rsTmp.BOF = False Then
        strTitle = ""
        strKey = ""
        Do While Not rsTmp.EOF
            strNodeName = zlCommFun.NVL(rsTmp("nodename").Value)
            If strNodeName = "station" Then
                
                strTitle = zlCommFun.NVL(rsTmp("nodetext").Value)
                strKey = "产品(站点)_" & zlCommFun.NVL(rsTmp("nodetext").Value)
                Call InsertDeliveryTree("站点", strKey, strTitle, rsDeliver)
                
            End If
            
            rsTmp.MoveNext
        Loop
    End If
    
    '消息(站点)
    '--------------------------------------------------------------------------------------------------------------
    Set rsTmp = objXML.GetNodeData("//stations/message")
    If rsTmp.BOF = False Then
        strTitle = ""
        strKey = ""
        Do While Not rsTmp.EOF
            strNodeName = zlCommFun.NVL(rsTmp("nodename").Value)
            If strNodeName = "message" Then
                strTitle = zlCommFun.NVL(rsTmp("nodetext").Value)
                strKey = "消息(站点)_" & zlCommFun.NVL(rsTmp("nodetext").Value)
                Call InsertDeliveryTree("站点", strKey, strTitle, rsDeliver)
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    '产品(模块)
    '--------------------------------------------------------------------------------------------------------------
    Set rsTmp = objXML.GetNodeData("//modules/module")
    If rsTmp.BOF = False Then
        strTitle = ""
        strKey = ""
        Do While Not rsTmp.EOF
            strNodeName = zlCommFun.NVL(rsTmp("nodename").Value)
            If strNodeName = "module" Then
                If strTitle <> "" Then Call InsertDeliveryTree("模块", "产品(模块)_" & strKey, strTitle, rsDeliver)
                strTitle = ""
                strKey = ""
            ElseIf strNodeName = "title" Then
                strTitle = zlCommFun.NVL(rsTmp("nodetext").Value)
            ElseIf strNodeName = "key" Then
                strKey = zlCommFun.NVL(rsTmp("nodetext").Value)
            End If
            
            rsTmp.MoveNext
        Loop
        If strTitle <> "" Then Call InsertDeliveryTree("模块", "产品(模块)_" & strKey, strTitle, rsDeliver)
    End If
        
    '消息(模块)
    '--------------------------------------------------------------------------------------------------------------
    Set rsTmp = objXML.GetNodeData("//modules/message")
    If rsTmp.BOF = False Then
        strTitle = ""
        strKey = ""
        Do While Not rsTmp.EOF
            strNodeName = zlCommFun.NVL(rsTmp("nodename").Value)
            If strNodeName = "message" Then
                strTitle = zlCommFun.NVL(rsTmp("nodetext").Value)
                strKey = "消息(模块)_" & zlCommFun.NVL(rsTmp("nodetext").Value)
                Call InsertDeliveryTree("模块", strKey, strTitle, rsDeliver)
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    '产品(用户)
    '--------------------------------------------------------------------------------------------------------------
    Set rsTmp = objXML.GetNodeData("//mipusers/mipuser")
    If rsTmp.BOF = False Then
        strTitle = ""
        strKey = ""
        Do While Not rsTmp.EOF
            strNodeName = zlCommFun.NVL(rsTmp("nodename").Value)
            If strNodeName = "mipuser" Then
                
                strTitle = zlCommFun.NVL(rsTmp("nodetext").Value)
                strKey = "产品(用户)_" & zlCommFun.NVL(rsTmp("nodetext").Value)
                Call InsertDeliveryTree("消息用户", strKey, strTitle, rsDeliver)
                
            End If
            
            rsTmp.MoveNext
        Loop
    End If
    
    '消息(用户)
    '--------------------------------------------------------------------------------------------------------------
    Set rsTmp = objXML.GetNodeData("//mipusers/message")
    If rsTmp.BOF = False Then
        strTitle = ""
        strKey = ""
        Do While Not rsTmp.EOF
            strNodeName = zlCommFun.NVL(rsTmp("nodename").Value)
            If strNodeName = "message" Then
                strTitle = zlCommFun.NVL(rsTmp("nodetext").Value)
                strKey = "消息(用户)_" & zlCommFun.NVL(rsTmp("nodetext").Value)
                Call InsertDeliveryTree("消息用户", strKey, strTitle, rsDeliver)
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    GetDeliveryTree = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function InsertDeliveryTree(ByVal strGroup As String, ByVal strTargetKey As String, ByVal strTargetContent As String, ByRef rsDeliver As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：填充数据到选中的目标表格
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intRow As Integer
    Dim intCount As Integer
                
    intCount = Switch(strGroup = "人员", 1, strGroup = "部门", 2, strGroup = "人员性质", 3, strGroup = "角色", 4, strGroup = "站点", 5, strGroup = "模块", 6, strGroup = "消息用户", 7)
                    
    rsDeliver.Filter = ""
    rsDeliver.Filter = "id='" & strGroup & "'"
    If rsDeliver.RecordCount = 0 Then
        rsDeliver.AddNew
        rsDeliver("id").Value = strGroup
        rsDeliver("上级id").Value = ""
        rsDeliver("状态").Value = 1
        rsDeliver("标题").Value = strGroup
        rsDeliver("排序1").Value = intCount * 100
        rsDeliver("排序2").Value = ""
    End If
    rsDeliver.Filter = ""
    
    rsDeliver.AddNew
    rsDeliver("id").Value = strTargetKey
    rsDeliver("上级id").Value = strGroup
    rsDeliver("状态").Value = 1
    rsDeliver("标题").Value = strTargetContent
    rsDeliver("排序1").Value = intCount * 100 + 1
    rsDeliver("排序2").Value = strTargetContent

    InsertDeliveryTree = True
    
End Function

Public Function UpgradeMipClient() As Boolean
    '******************************************************************************************************************
    '功能:升级消息集成平台客户端数据库
    '参数:
    '返回:
    '******************************************************************************************************************
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    '功能：进行服务器管理工具的升级
    '参数：strInstallFile   安装脚本的完整位置
'    Dim cnTools As New ADODB.Connection
    Dim rsTemp As New ADODB.Recordset
'    Dim dblVerDatabase As Double, dblVerFile As Double, varVer As Variant
'    Dim objSys As New Scripting.FileSystemObject, objText As Scripting.TextStream, strLine As String
    Dim strPath As String
'    Dim objFolder As Scripting.Folder, objFile As Scripting.File
'    Dim arrVer() As Double, i As Long, j As Long
    
    On Error GoTo errHand
    
'    strPath = Left(strInstallFile, Len(strInstallFile) - Len("zlSetup.ini"))
    
'    '1、判断管理工具的版本
    
'    Set rsTemp = mclsOracle.OpenCursor("", "ZLTOOLS", "B_Public.Get_Ver")
'    If rsTemp.EOF = False Then
'        '产生一个12位的数字
'        dblVerDatabase = GetVerDouble(rsTemp("内容"))
'    End If
'
'    '2、获得安装程序需要的版本
''    On Error Resume Next
'    Set objText = objSys.OpenTextFile(strInstallFile) '打开zlSetup.ini文件
'    If Err <> 0 Then
'        MsgBox "无法打开安装配置文件。" & vbCrLf & vbTab & strInstallFile, vbInformation, gstrSysName
'        Exit Function
'    End If
    
    
    
'    Do Until objText.AtEndOfStream
'        strLine = Trim(objText.ReadLine)
'        If Left(strLine, 9) = "[管理工具版本号]" Then
'            dblVerFile = GetVerDouble(Mid(strLine, 10))
'            Exit Do
'        End If
'    Loop
'    objText.Close
    
'    dblVerFile = GetVerDouble(mstrManageVersion)
'
'    '3、比较版本，是否需要升级
'    If dblVerDatabase >= dblVerFile Then
'        '满足要求，不需要升级
'        UpgradeManager = True
'        Exit Function
'    End If
    
'    '4、连接zltools
'    If gstrToolsPwd = "-" Or gstrToolsPwd = "" Then
'        gstrToolsPwd = frmToolsPwd.GetToolsPwd()
'        If gstrToolsPwd = "-" Then Exit Function
'    End If
    
'    cnTools.Open "Driver={Microsoft ODBC for Oracle};Server=" & gstrServer, "ZLTOOLS", gstrToolsPwd
'    If cnTools.State = adStateClosed Then
'        MsgBox "无法以ZLTOOLS用户连接，错误信息如下：" & vbCrLf & vbCrLf & Err.Description, vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    '5、按文件的版本逐个运行
'    '5.1 获得升级目录
'    strPath = Left(strInstallFile, Len(strInstallFile) - Len("zlSetup.ini")) & "..\..\Tools"
'    Set objFolder = objSys.GetFolder(strPath)
'    If Err <> 0 Then
'        MsgBox "打开管理脚本存放目录（[安装目录]\Tools）错误。", vbInformation, gstrSysName
'        Exit Function
'    End If
'    strPath = objFolder.Path & "\" '该地址已经规则了，去掉..
'
'    '5.2 列出目录下的所有升级文件
'    ReDim arrVer(0 To 0) '首先初始化
'    For Each objFile In objFolder.Files
'        If UCase(objFile.Name) Like "ZLUPGRADE*.SQL" Then
'            '升级文件
'            strLine = Mid(objFile.Name, 10)
'            strLine = Mid(strLine, 1, Len(strLine) - 4)
'
'            i = UBound(arrVer) + 1
'            ReDim Preserve arrVer(0 To i)
'            arrVer(i) = GetVerDouble(strLine)
'        End If
'    Next
'
'    '将版本从小到大地排列
'    For i = 1 To UBound(arrVer) - 1
'        For j = i + 1 To UBound(arrVer)
'            If arrVer(i) > arrVer(j) Then
'                varVer = arrVer(i)
'                arrVer(i) = arrVer(j)
'                arrVer(j) = varVer
'            End If
'        Next
'    Next
    
    
'    For i = 1 To UBound(arrVer)
'        If arrVer(i) > dblVerDatabase And arrVer(i) <= dblVerFile Then
'            '执行升级脚本
'            If RunUpgrade(cnTools, strPath & "ZLUPGRADE" & GetVerString(arrVer(i)) & ".SQL") = False Then
'                Exit Function
'            End If
'
'            '更新版本号
'            Call OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Update_Ver", GetVerString(arrVer(i)))
'        End If
'    Next
'
'    If arrVer(UBound(arrVer)) < dblVerFile Then
'        '还是没有完全升级成功
'        MsgBox "缺少管理工具的升级脚本：ZLUPGRADE" & GetVerString(dblVerFile) & ".SQL", vbInformation, gstrSysName
'        Exit Function
'    End If
    
    
'    Set objFolder = mobjFso.GetFolder(strInstallFile)
    
    UpgradeMipClient = True
    Exit Function
errHand:
    
End Function

Public Function UnloadMipClient() As Boolean
    '******************************************************************************************************************
    '功能:卸载消息集成平台客户端数据库
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim strSQL As String
    Dim intLoop As Integer
    Dim varObject As Variant
    Dim strObject As String
    Dim rsTemp As ADODB.Recordset
        
    On Error GoTo errHand
       
    '删除过程及同义词
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select Distinct Object_Name From All_Objects Where Owner = 'ZLTOOLS' And Object_Type = [1] And Object_Name Like 'ZLMIP_%'"
    Set rsTemp = mclsDataOracle.OpenSQLRecord(strSQL, mstrTitle, "PROCEDURE")
    If rsTemp.BOF = False Then
        Do While Not rsTemp.EOF
            If mclsDataOracle.ExecuteSQL("Drop Public Synonym " & rsTemp("Object_Name").Value, mstrTitle) = False Then Exit Function
            If mclsDataOracle.ExecuteSQL("Drop Procedure zltools." & rsTemp("Object_Name").Value, mstrTitle) = False Then Exit Function
            rsTemp.MoveNext
        Loop
    End If
    
    '删除包及同义词
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select Distinct Object_Name From All_Objects Where Owner = 'ZLTOOLS' And Object_Type = [1] And Object_Name Like 'ZLMIP_%'"
    Set rsTemp = mclsDataOracle.OpenSQLRecord(strSQL, mstrTitle, "PACKAGE")
    If rsTemp.BOF = False Then
        Do While Not rsTemp.EOF
            If mclsDataOracle.ExecuteSQL("Drop Public Synonym " & rsTemp("Object_Name").Value, mstrTitle) = False Then Exit Function
            If mclsDataOracle.ExecuteSQL("Drop Package zltools." & rsTemp("Object_Name").Value, mstrTitle) = False Then Exit Function
            rsTemp.MoveNext
        Loop
    End If
        
    '删除表及同义词
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select Table_Name,Constraint_Name From All_Constraints Where Owner = 'ZLTOOLS' And Table_Name Like 'ZLMIP_%' Order By Decode(Constraint_Type, 'R', 1, 'U', 2, 3)"
    Set rsTemp = mclsDataOracle.OpenSQLRecord(strSQL, mstrTitle, "TABLE")
    If rsTemp.BOF = False Then
        Do While Not rsTemp.EOF
            strSQL = "Alter Table zltools." & rsTemp("Table_Name").Value & " Drop Constraint " & rsTemp("Constraint_Name").Value
            If mclsDataOracle.ExecuteSQL(strSQL, mstrTitle) = False Then Exit Function
            rsTemp.MoveNext
        Loop
    End If
    
    strSQL = "Select Distinct Object_Name From All_Objects Where Owner = 'ZLTOOLS' And Object_Type = [1] And Object_Name Like 'ZLMIP_%'"
    Set rsTemp = mclsDataOracle.OpenSQLRecord(strSQL, mstrTitle, "TABLE")
    If rsTemp.BOF = False Then
        Do While Not rsTemp.EOF
            Call mclsDataOracle.ExecuteSQL("Drop Public Synonym " & rsTemp("Object_Name").Value, mstrTitle)
            If mclsDataOracle.ExecuteSQL("Drop Table zltools." & rsTemp("Object_Name").Value, mstrTitle) = False Then Exit Function
            rsTemp.MoveNext
        Loop
    End If
        
    '------------------------------------------------------------------------------------------------------------------
'    strSQL = "Delete From zlOptions Where 参数号 In (16,17,18)"
'    Call mclsDataOracle.ExecuteSQL(strSQL, mstrTitle)
    
    strSQL = "Delete From zlReginfo Where 项目='消息集成平台客户端'"
    Call mclsDataOracle.ExecuteSQL(strSQL, mstrTitle)
    
    UnloadMipClient = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox "卸载消息集成平台客户端失败，原因如下：" & vbCrLf & Err.Description, vbCritical, mstrTitle
'    Resume
End Function

'######################################################################################################################
'私有方法

Private Function SQLRecordExecute(Optional ByVal blnHaveTrans As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim strSQL As String

    On Error GoTo errHand

    If mrsSQL.RecordCount > 0 Then

        blnTran = True

        If blnHaveTrans Then mclsDataOracle.BeginTrans

        mrsSQL.MoveFirst

        For intLoop = 1 To mrsSQL.RecordCount

            strSQL = CStr(mrsSQL("SQL").Value)
            Call zlDataBase.ExecuteProcedure(strSQL, mstrTitle)

            mrsSQL.MoveNext
        Next

        If blnHaveTrans Then mclsDataOracle.CommitTrans
        blnTran = False
    End If

    SQLRecordExecute = True

    Exit Function
errHand:

    If blnTran And blnHaveTrans Then mclsDataOracle.RollbackTrans

    MsgBox Err.Description

'    If ErrCenter = 1 Then
'        Resume
'    End If

End Function




