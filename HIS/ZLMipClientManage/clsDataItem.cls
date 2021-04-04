VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
    mstrOwner = strOwner
    
    InitBusiness = True
    
    Exit Function
    
errHand:
    
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

Public Function TableStruct() As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取Infotable结构
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mstrSQL = "Select * From zlmsp_table Where 1=2"
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
    
    mstrSQL = "Select * From zlmsp_tab_extend Where 1=2"
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
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
        
        '赋值
        mstrSQL = "TabExt_Condition_Delete('" & zlCommFun.GetParameter(rsPara, "id") & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
        lngCount = Val(zlCommFun.GetParameter(rsPara, "赋值关系段数"))
        For lngLoop = 1 To lngCount
            strTemp = Trim(zlCommFun.GetParameter(rsPara, "赋值关系_" & lngLoop))
            If strTemp <> "" Then
                mstrSQL = "TabExt_Condition_Insert('" & zlCommFun.GetParameter(rsPara, "id") & "','" & strTemp & "')"
                mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
                Call SQLRecordAdd(mrsSQL, mstrSQL)
            End If
        Next

    '------------------------------------------------------------------------------------------------------------------
    Case "DELETE"
        mstrSQL = "Tab_Extend_Delete('" & zlCommFun.GetParameter(rsPara, "id") & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
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

Public Function EventEdit(Optional ByVal strUpdateMode As String = "", Optional ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：编辑Event数据
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)
    
    Select Case UCase(strUpdateMode)
    '------------------------------------------------------------------------------------------------------------------
    Case "INSERT", "UPDATE"
        mstrSQL = "Event_Update('" & zlCommFun.GetParameter(rsParameter, "id") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "code") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "title") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "app") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "device") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "kind") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "note") & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "DELETE"
        mstrSQL = "Event_Delete('" & zlCommFun.GetParameter(rsParameter, "id") & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    End Select
        
    EventEdit = SQLRecordExecute(False)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

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
                    "From m_Evt_Message a,zlmsp_table b Where a.tab_id=b.id(+) And a.event_id=[1]"
                    
        Set EventMsgRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "业务事件id"))
    '------------------------------------------------------------------------------------------------------------------
    Case "事件消息"
    
        mstrSQL = "Select 'event' As 图标,a.id,a.event_id,a.msg_code as 消息,a.msg_ver as 版本,b.tab_title as 入口,a.tab_id,a.send_fail,a.note As 说明 " & _
                    "From m_Evt_Message a,zlmsp_table b Where a.tab_id=b.id(+) And a.id=[1]"
                    
        Set EventMsgRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "事件消息id"))
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("Select")
        mstrSQL = _
            "Select ID, Null As 上级id, Null As 消息编号, Title As 名称, Null As 入口信息, 0 As 末级,Null As 说明" & vbNewLine & _
            "From m_Event" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select a.Id, a.Event_Id As 上级id, a.Msg_Code As 消息编号, Null As 名称, b.tab_Title As 入口信息, 1 As 末级,a.Note As 说明" & vbNewLine & _
            "From m_Evt_Message A, zlmsp_table B" & vbNewLine & _
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
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
        
        '配置
        mstrSQL = "EvtMsg_Config_Delete('" & zlCommFun.GetParameter(rsPara, "id") & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
        lngCount = Val(zlCommFun.GetParameter(rsPara, "消息配置段数"))
        For lngLoop = 1 To lngCount
            strTemp = Trim(zlCommFun.GetParameter(rsPara, "消息配置_" & lngLoop))
            
            If strTemp <> "" Then
                mstrSQL = "EvtMsg_Config_Insert('" & zlCommFun.GetParameter(rsPara, "id") & "','" & strTemp & "')"
                mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
                Call SQLRecordAdd(mrsSQL, mstrSQL)
            End If
        Next
                
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("Copy")
            
        '消息
        mstrSQL = "Evt_Message_Copy('" & zlCommFun.GetParameter(rsPara, "业务事件id") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "参考消息id") & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
                               
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("TargetConfig")
    
        '投递目标服务器
        mstrSQL = "EvtMsg_Target_Update('" & zlCommFun.GetParameter(rsPara, "id") & "','" & Trim(zlCommFun.GetParameter(rsPara, "投递目标")) & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("BatchTargetConfig")
        
        '投递目标服务器
        mstrSQL = "EvtMsg_Target_BatchUpdate('" & zlCommFun.GetParameter(rsPara, "TargetService") & "'," & Trim(zlCommFun.GetParameter(rsPara, "ConfigOption")) & ",'" & Trim(zlCommFun.GetParameter(rsPara, "ConfigString")) & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
    '------------------------------------------------------------------------------------------------------------------
    Case "DELETE"
        mstrSQL = "Evt_Message_Delete('" & zlCommFun.GetParameter(rsPara, "事件消息id") & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    End Select
        
    EventMsgEdit = SQLRecordExecute(True)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
    
End Function

Public Function EventMsgConfigRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取MsgEvent数据
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case "事件消息"
        
        
        mstrSQL = _
                "Select a.id,node_order," & vbNewLine & _
                "       a.parent_id," & vbNewLine & _
                "       a.node_title as 节点标题," & vbNewLine & _
                "       Decode(a.node_type, 1,'Segment', 2,'Composite', 3,'Data', 4,'Group') as 节点类型," & vbNewLine & _
                "       a.data_type as 数据类型," & vbNewLine & _
                "       a.note as 节点说明," & vbNewLine & _
                "       a.min_occurs as 最小次数," & vbNewLine & _
                "       a.max_occurs As 最大次数," & vbNewLine & _
                "       a.config_occurs as 数据重复," & vbNewLine & _
                "       a.config_express as 数据赋值," & vbNewLine & _
                "       a.config_occurs_key as 数据重复_Key," & vbNewLine & _
                "       a.config_express_key as 数据赋值_Key" & vbNewLine & _
                "  From m_Evtmsg_Config a" & vbNewLine & _
                " Where a.evt_msg_id = [1]" & vbNewLine & _
                " Order by a.node_order"
                                
        'substr('                                  ', 1, 2*(level - 1)) || 节点类型 As 节点类型
        
        mstrSQL = _
                "select level,id,parent_id,'' as 图标," & vbNewLine & _
                "       节点类型," & vbNewLine & _
                "       节点标题," & vbNewLine & _
                "       节点说明," & vbNewLine & _
                "       最小次数 || ' ～ ' || 最大次数 As 重复频率," & vbNewLine & _
                "       数据类型," & vbNewLine & _
                "       数据重复," & vbNewLine & _
                "       数据赋值," & vbNewLine & _
                "       数据重复_Key," & vbNewLine & _
                "       数据赋值_Key" & vbNewLine & _
                "  From (" & mstrSQL & ") " & vbNewLine & _
                " start with parent_id is null" & vbNewLine & _
                "connect by prior id = parent_id Order by node_order"
        
        Set EventMsgConfigRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "事件消息id"))
        
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
    
    mstrSQL = "Select id,'' as parent_id,tab_title,'' as fld from zlmsp_table where id=[1]"
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
            
    mstrSQL = "Select a.id,a.ext_order,a.ext_type,a.ext_title,b.tab_title,a.target_tab_id from zlmsp_tab_extend a,zlmsp_table b where a.source_tab_id=[1] And a.target_tab_id=b.id"
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
    
    mstrSQL = "Select fld_order,fld_title From zlmsp_tab_field Where tab_id=[1]"
    
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
        mstrSQL = "Select 'infotable' As 图标,a.* From zlmsp_table a"
        Set TableRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("ID")
        mstrSQL = "Select 'infotable' As 图标,a.* From zlmsp_table a Where id=[1]"
        Set TableRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "id"))
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("SelectData")
        mstrSQL = "Select a.ID,a.tab_code As 编码,a.tab_title As 名称,a.tab_note As 说明 From zlmsp_table a"
        Set TableRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("FilterData")
        
        strStyle = zlCommFun.GetCondition(rsCondition, "FilterStyle")
        strTemp = zlCommFun.GetCondition(rsCondition, "FilterText")
        
        Select Case strStyle
        Case ""
            mstrSQL = "Select a.ID,a.tab_code As 编码,a.tab_title As 名称,a.tab_note As 说明 From zlmsp_table a Where a.tab_code Like [1] Or a.tab_title Like [2]"
            Set TableRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strTemp & "%", "%" & strTemp & "%")
        Case "编码"
            mstrSQL = "Select a.ID,a.tab_code As 编码,a.tab_title As 名称,a.tab_note As 说明 From zlmsp_table a Where a.tab_code Like [1]"
            Set TableRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strTemp & "%")
        Case "名称"
            mstrSQL = "Select a.ID,a.tab_code As 编码,a.tab_title As 名称,a.tab_note As 说明 From zlmsp_table a Where a.tab_title Like [1]"
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
        mstrSQL = "Select a.* From zlmsp_tab_field a Where a.tab_id=[1]"
        Set TableFieldRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "业务信息id"))
    '------------------------------------------------------------------------------------------------------------------
    Case "字段"
        
        If zlCommFun.GetCondition(rsCondition, "字段前缀") = "" Then
            mstrSQL = "Select a.fld_title As 字段名称,a.fld_type As 字段类型 From zlmsp_tab_field a Where a.tab_id=[1]"
        Else
            mstrSQL = "Select '" & zlCommFun.GetCondition(rsCondition, "字段前缀") & "'||'.'||a.fld_title As 字段名称,Decode(a.fld_type,1,'数值',2,'字符',3,'日期') As 字段类型 From zlmsp_tab_field a Where a.tab_id=[1]"
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
        mstrSQL = "Select 'infotable' As 图标,tab_id,para_order,para_title,decode(para_type,2,'字符',1,'数值',3,'日期','自定义') as para_type,para_default,para_note  From zlmsp_tab_parameter"
        Set TableParameterRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("tab_id")
        mstrSQL = "Select 'infotable' As 图标,tab_id,para_order,para_title,decode(para_type,2,'字符',1,'数值',3,'日期','自定义') as para_type,para_default,para_note  From zlmsp_tab_parameter Where tab_id=[1]"
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
            "  from zlmsp_tab_parameter x," & vbNewLine & _
            "       (select b.target_fld,b.source_fld " & vbNewLine & _
            "          from zlmsp_tab_extend a, zlmsp_tabext_condition b" & vbNewLine & _
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
'            mstrSQL = "Select a.fld_title As 字段名称,a.fld_type As 字段类型 From zlmsp_tab_field a Where a.tab_id=[1]"
'        Else
'            mstrSQL = "Select '" & zlCommFun.GetCondition(rsCondition, "字段前缀") & "'||'.'||a.fld_title As 字段名称,Decode(a.fld_type,1,'数值',2,'字符',3,'日期') As 字段类型 From zlmsp_tab_field a Where a.tab_id=[1]"
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
                                            "'" & zlCommFun.GetParameter(rsPara, "tab_code") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "tab_title") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "tab_sqltext") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "tab_note") & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
        
        '信息表SQL
        '--------------------------------------------------------------------------------------------------------------
        mstrSQL = "Tab_Parameter_Delete('" & zlCommFun.GetParameter(rsPara, "id") & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
        lngCount = Val(zlCommFun.GetParameter(rsPara, "SQL参数个数"))
        For lngLoop = 1 To lngCount
            strTemp = Trim(zlCommFun.GetParameter(rsPara, "SQL参数_" & lngLoop))
            
            If strTemp <> "" Then
                mstrSQL = "Tab_Parameter_Insert('" & zlCommFun.GetParameter(rsPara, "id") & "','" & strTemp & "')"
                mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
                Call SQLRecordAdd(mrsSQL, mstrSQL)
            End If
        Next
        
        '信息表字段
        '--------------------------------------------------------------------------------------------------------------
        mstrSQL = "Tab_Field_Delete('" & zlCommFun.GetParameter(rsPara, "id") & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)

        lngCount = Val(zlCommFun.GetParameter(rsPara, "SQL字段个数"))
        For lngLoop = 1 To lngCount
            strTemp = Trim(zlCommFun.GetParameter(rsPara, "SQL字段_" & lngLoop))

            If strTemp <> "" Then
                mstrSQL = "Tab_Field_Insert('" & zlCommFun.GetParameter(rsPara, "id") & "','" & strTemp & "')"
                mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
                Call SQLRecordAdd(mrsSQL, mstrSQL)
            End If
        Next
                
    '------------------------------------------------------------------------------------------------------------------
    Case "DELETE"
        mstrSQL = "Table_Delete('" & zlCommFun.GetParameter(rsPara, "id") & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
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
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "DELETE"
        mstrSQL = "Tab_Parameter_Delete('" & zlCommFun.GetParameter(rsParameter, "tab_id") & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
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
                "From zlmsp_tab_extend a,zlmsp_table b, " & _
                "       (select x.ext_id," & vbNewLine & _
                "               ltrim(f_List2str(Cast(Collect(' '||x.target_fld || ' = ' || Decode(x.source_fld,null,null,'原表.'||x.source_fld)||' ') As t_Strlist))) As para" & vbNewLine & _
                "          from zlmsp_tabext_condition x, zlmsp_tab_extend y" & vbNewLine & _
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
            "          from zlmsp_tabext_condition x" & vbNewLine & _
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
            "  From zlmsp_tab_extend a," & vbNewLine & _
            "       zlmsp_table b," & vbNewLine & _
            "       (select x.ext_id," & vbNewLine & _
            "               ltrim(f_List2str(Cast(Collect(' '||x.target_fld || ' = ' || Decode(x.source_fld,null,null,'原表.'||x.source_fld)||' ') As t_Strlist))) As para" & vbNewLine & _
            "          from zlmsp_tabext_condition x" & vbNewLine & _
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
    MsgBox Err.Description, vbInformation
'    Resume
End Function

Public Function ItemClassStruct() As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取Infotable结构
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mstrSQL = "Select * From zlmsp_folder Where 1=2"
    Set ItemClassStruct = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function ItemClassRead(Optional ByVal strConditionMode As String = "", Optional ByVal rsCondition As ADODB.Recordset) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取zlMsp_Folder数据
    '参数：
    '返回：
    '******************************************************************************************************************
        
    On Error GoTo errHand
    
    Select Case UCase(strConditionMode)
    '------------------------------------------------------------------------------------------------------------------
    Case ""
        mstrSQL = _
            "Select ID, 上级id, 编码, 名称, 图标, 选中图标" & vbNewLine & _
            "From (Select ID, parent_id As 上级id, folder_code As 编码, folder_name As 名称, 'folder' As 图标, 'folder_open' As 选中图标" & vbNewLine & _
            "       From zlmsp_folder" & vbNewLine & _
            "       Start With parent_id Is Null" & vbNewLine & _
            "       Connect By Prior id = parent_id)" & vbNewLine & _
            "Order By 编码"
        Set ItemClassRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("ID")
        If zlCommFun.GetCondition(rsCondition, "id") = "" Then
            mstrSQL = _
                "Select ID, 上级id, 编码, 名称, 图标, 选中图标" & vbNewLine & _
                "From (Select ID, parent_id As 上级id, folder_code As 编码, folder_name As 名称, 'folder' As 图标, 'folder_open' As 选中图标" & vbNewLine & _
                "       From zlmsp_folder" & vbNewLine & _
                "       Start With parent_id Is Null" & vbNewLine & _
                "       Connect By Prior id = parent_id)" & vbNewLine & _
                "Order By 编码"
        Else
            mstrSQL = _
                "Select A.ID, A.parent_id As 上级id, A.folder_code As 编码, A.folder_name As 名称, B.folder_code As 上级编码, B.folder_name As 上级名称, 'folder' As 图标," & vbNewLine & _
                "       'folder_open' As 选中图标" & vbNewLine & _
                "From zlmsp_folder A, zlmsp_folder B" & vbNewLine & _
                "Where A.parent_id = B.id(+) And A.id = [1]"
        End If
        Set ItemClassRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "id"))
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("SelectData")
        '
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("FilterData")
        '
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetClassDefaultCode(ByVal 上级编码_In As String, ByVal 最大编码_In As String, ByVal 最大编码长度_In As Long, Optional ByRef blnMsg As Boolean = False) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************
    Dim v_缺省编码 As String
    Dim n_调整编码 As Integer
    Dim n_允许调整 As Integer
    Dim n_允许编码长度 As Long
    Dim n_允许编码最大长度 As Long
    Dim v_Error As String
    
    On Error GoTo errHand
    If 最大编码长度_In = Len(上级编码_In) Then
        ShowSimpleMsg "最大编码和编码长度已经达到最大限制，无法递增编码。"
        v_缺省编码 = ""
        n_调整编码 = 0
        n_允许调整 = 0
        blnMsg = True
        Exit Function
    End If
    
    If 最大编码_In = "" Then
        n_调整编码 = 1
        v_缺省编码 = "01"
        n_允许编码长度 = 最大编码长度_In - Len(上级编码_In)
        n_允许编码最大长度 = n_允许编码长度
        n_允许调整 = 0
    Else
        n_调整编码 = 0
        n_允许编码长度 = Len(最大编码_In) - Len(上级编码_In)
        n_允许编码最大长度 = 最大编码长度_In - Len(上级编码_In)
        n_允许调整 = 1
        
        If Mid(最大编码_In, Len(上级编码_In) + 1, 30) = Lpad("9", n_允许编码长度, 9) Then
            If n_允许编码长度 >= 最大编码长度_In Then
                ShowSimpleMsg "最大编码和编码长度已经达到最大限制，无法递增编码。"
                n_调整编码 = 0
                n_允许调整 = 0
                v_缺省编码 = Lpad(" ", n_允许编码长度, " ")
                blnMsg = True
                Exit Function
            Else
                ShowSimpleMsg "最大编码已经达到本级限制，你可以扩充编码长度以满足需要。"
                n_调整编码 = 1
                n_允许调整 = 1
                v_缺省编码 = "1" & Lpad("0", n_允许编码长度, "0")
                n_允许编码长度 = n_允许编码长度 + 1
                n_允许编码最大长度 = 最大编码长度_In - Len(上级编码_In)
                blnMsg = True
                Exit Function
            End If
        Else
            v_缺省编码 = Lpad(CStr(Val(Mid(最大编码_In, Len(上级编码_In) + 1, 30)) + 1), n_允许编码长度, "0")
        End If
        
    End If
    
    mstrSQL = _
        "Select [1] As 缺省编码,[2] As 上级编码,[3] As 允许编码长度,[4] As 允许编码最大长度,[5] As 调整编码,[6] As 允许调整 From Dual"
    
    Set GetClassDefaultCode = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, v_缺省编码, 上级编码_In, n_允许编码长度, n_允许编码最大长度, n_调整编码, n_允许调整)
    
    Exit Function
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ItemClassMaxCode(ByVal strUpKey As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************
    On Error GoTo errHand

    mstrSQL = "Select Max(folder_code) As 最大编码 From zlmsp_folder Where Nvl(parent_id, '-') = Nvl([1], '-')"
    Set ItemClassMaxCode = zlDataBase.OpenSQLRecord(mstrSQL, mstrTitle, strUpKey)
    
    Exit Function
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function ItemClassEdit(Optional ByVal strUpdateMode As String = "", Optional ByVal rsPara As ADODB.Recordset) As Boolean
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
    Case "INSERT"
        
        '信息表
        mstrSQL = "ItemClass_Insert('" & zlCommFun.GetParameter(rsPara, "id") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "parent_id") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "folder_code") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "folder_name") & "'," & _
                                            zlCommFun.GetParameter(rsPara, "adjustcodelength") & ")"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
                        
    '------------------------------------------------------------------------------------------------------------------
    Case "UPDATE"
        
        '信息表
        mstrSQL = "ItemClass_Update('" & zlCommFun.GetParameter(rsPara, "id") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "parent_id") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "folder_code") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsPara, "folder_name") & "'," & _
                                            zlCommFun.GetParameter(rsPara, "adjustcodelength") & ")"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
        
    '------------------------------------------------------------------------------------------------------------------
    Case "DELETE"
        mstrSQL = "ItemClass_Delete('" & zlCommFun.GetParameter(rsPara, "id") & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    End Select
        
    ItemClassEdit = SQLRecordExecute(True)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ItemStruct() As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取Event结构
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mstrSQL = "Select * From zlmsp_item Where 1=2"
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
        mstrSQL = "Select 'event' As 图标,a.* From zlmsp_item a Order By a.code"
        Set ItemRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle)
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("folder_id")
        mstrSQL = "Select a.folder_code,a.folder_name From zlmsp_folder a Where id=[1]"
        Set ItemRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "folder_id"))
        
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("ID")
        mstrSQL = "Select 'event' As 图标,a.* From m_Event a Where id=[1]"
        Set ItemRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "id"))
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("Code")
        mstrSQL = "Select a.ID,a.code As 编码,a.title As 名称,a.app As 程序,a.device As 设备,a.kind as 类型 From m_Event a Where code Like [1]"
        Set ItemRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "code"))
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("Title")
        mstrSQL = "Select a.ID,a.code As 编码,a.title As 名称,a.app As 程序,a.device As 设备,a.kind as 类型 From m_Event a Where title Like [1]"
        Set ItemRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, zlCommFun.GetCondition(rsCondition, "title"))
        
    '------------------------------------------------------------------------------------------------------------------
    Case UCase("FilterData")
        
        strStyle = zlCommFun.GetCondition(rsCondition, "FilterStyle")
        strTemp = zlCommFun.GetCondition(rsCondition, "FilterText")
        
        Select Case strStyle
        Case ""
            mstrSQL = "Select a.ID,a.code As 编码,a.title As 名称,a.app As 程序,a.device As 设备,a.kind as 类型 From m_Event a Where code Like [1]  Or a.title Like [2]"
            Set ItemRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strTemp & "%", "%" & strTemp & "%")
        Case "编码"
            mstrSQL = "Select a.ID,a.code As 编码,a.title As 名称,a.app As 程序,a.device As 设备,a.kind as 类型 From m_Event a Where code Like [1]"
            Set ItemRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strTemp & "%")
        Case "名称"
            
            mstrSQL = "Select a.ID,a.code As 编码,a.title As 名称,a.app As 程序,a.device As 设备,a.kind as 类型 From m_Event a Where title Like [1]"
            Set ItemRead = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, "%" & strTemp & "%")
        End Select
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function ItemDefaultCode(ByVal strClassKey As String, ByVal strClassCode As String, intMaxLength As Integer) As String
    '******************************************************************************************************************
    '功能：产生缺省编码
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim str最大编码 As String
    Dim str编码 As String

    mstrSQL = "Select Max(item_code) As 最大编码 From zlmsp_item Where Nvl(folder_id, '-') = Nvl([1], '-')"
    Set rs = mclsDataOracle.OpenSQLRecord(mstrSQL, mstrTitle, strClassKey)
    If rs.BOF = False Then
        str最大编码 = Trim(zlCommFun.NVL(rs("最大编码").Value))
    End If
    If str最大编码 = "" Then
        str编码 = strClassCode & "01"
    ElseIf Len(str最大编码) > Len(strClassCode) Then
        str最大编码 = Mid(str最大编码, Len(strClassCode) + 1)
        str编码 = strClassCode & Format(Val(str最大编码) + 1, String(Len(str最大编码), "0"))
    Else
        str编码 = strClassCode & "01"
    End If
    
    
    If Len(str编码) <= intMaxLength Then
        ItemDefaultCode = str编码
    End If
    
End Function

Public Function EventEdit(Optional ByVal strUpdateMode As String = "", Optional ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：编辑Event数据
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Call SQLRecord(mrsSQL)
    
    Select Case UCase(strUpdateMode)
    '------------------------------------------------------------------------------------------------------------------
    Case "INSERT", "UPDATE"
        mstrSQL = "Event_Update('" & zlCommFun.GetParameter(rsParameter, "id") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "code") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "title") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "app") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "device") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "kind") & "'," & _
                                            "'" & zlCommFun.GetParameter(rsParameter, "note") & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "DELETE"
        mstrSQL = "Event_Delete('" & zlCommFun.GetParameter(rsParameter, "id") & "')"
        mstrSQL = mstrOwner & ".zl_PkgMspManage." & mstrSQL
        Call SQLRecordAdd(mrsSQL, mstrSQL)
    End Select
        
    EventEdit = SQLRecordExecute(False)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
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






