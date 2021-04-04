VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrivilege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'变量定义
Private mstrRole As String
Private mstrOwner() As String
Private mstrSQL As String
Private mrsModule As ADODB.Recordset                '模块授予情况记录
Private mrsTable  As ADODB.Recordset                '编码表授予情况记录
Private mrsFunction  As ADODB.Recordset             '函数授予情况记录
Private mrsPrivs As ADODB.Recordset
Private mrsGranded As ADODB.Recordset            '保存以前授权的模块
Private mrsFuncPars As ADODB.Recordset           '函数参数所访问的表
Private mcnOracle As ADODB.Connection
Private mobjCollectionMoudal As Collection       '
Private mobjCollectionGranded As Collection      '
Private mobjColGrantObject As Collection         '具有的权限
Private mobjColRevokeObject As Collection        '不具有的权限
Private mstrFailInfo As String
Private mblnHaveBase As Boolean                  '是否具有字典管理工具模块权限

'事件定义
Public Event BeforeProgress(ByVal Title As String, ByVal Max As Long)
Public Event Progressing(ByVal Progress As Long)
Public Event AfterProgress()

'######################################################################################################################

Public Function InitOracle(ByVal cnOracle As ADODB.Connection)
    Set mcnOracle = cnOracle
    InitOracle = True
End Function

Public Function InitPrivilege(strOwner() As String, ByVal strRole As String, ByVal rsModule As ADODB.Recordset, ByVal rsTable As ADODB.Recordset, ByVal rsFunction As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：接口函数，初始化权限处理数据
    '参数：
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    Dim strKey As String
    
    On Error GoTo errHand
    
    mstrFailInfo = ""
    mstrRole = strRole
    mstrOwner = strOwner
    Set mrsModule = rsModule
    Set mrsTable = rsTable
    Set mrsFunction = rsFunction
    
    '------------------------------------------------------------------------------------------------------------------
    Set mrsGranded = OpenCursor(mcnOracle, "ZLTOOLS.B_Popedom.Get_Role_Grant", mstrRole)
    Set mrsPrivs = mrsGranded.NextRecordset
    Set mrsFuncPars = mrsGranded.NextRecordset
    '判断是否具有字典管理工具权限
    mrsModule.Filter = "序号=11 And 系统=0 And 授权否=1"
    mblnHaveBase = Not mrsModule.EOF
    
    mrsModule.Filter = ""
    If mrsModule.BOF = False Then
        mrsModule.MoveFirst
        Do While Not mrsModule.EOF
            strKey = "K" & mrsModule("系统").value & "'" & mrsModule("序号").value & "'" & mrsModule("功能").value & "'" & mrsModule("授权否").value
            mobjCollectionMoudal.Add "1", strKey
            mrsModule.MoveNext
        Loop
    End If
        
    If mrsGranded.BOF = False Then
        Do While Not mrsGranded.EOF
            strKey = "K" & mrsGranded("系统").value & "'" & mrsGranded("序号").value & "'" & mrsGranded("功能").value
            mobjCollectionGranded.Add "1", strKey
            mrsGranded.MoveNext
        Loop
    End If
    
    InitPrivilege = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
End Function

Public Function RevokePrivilege() As Boolean
    '******************************************************************************************************************
    '功能：接口函数，清除已有的对象权限
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim blnReturn As Boolean
    
    On Error GoTo errHand
    
    '1  取消以前授予的对数据表的访问权限
    '------------------------------------------------------------------------------------------------------------------
    blnReturn = GetDictRevokeObjectData(mrsTable)

    '2 取消以前授予的对函数的访问权限
    '------------------------------------------------------------------------------------------------------------------
    If blnReturn Then blnReturn = GetFunctionRevokeObjectData(mrsFunction, mrsFuncPars)
    
    '3  取消以前授予的对模块的访问权限
    '------------------------------------------------------------------------------------------------------------------
    If blnReturn Then blnReturn = GetMoudalRevokeObjectData(mrsGranded, mrsPrivs)
           
    '4  开始执行清除权限
    '------------------------------------------------------------------------------------------------------------------
    If blnReturn Then blnReturn = DeleteGrant(mrsGranded, mrsPrivs)
    If blnReturn Then blnReturn = ExecuteRevokeObject(mobjColRevokeObject)
    
    RevokePrivilege = blnReturn

    Exit Function
    
errHand:
    MsgBox err.Description
End Function

Public Property Get FailInfo() As String
    FailInfo = mstrFailInfo
End Property

Public Function ReGrantPrivilege(ByVal strRole As String, strOwner() As String) As Boolean
    '******************************************************************************************************************
    '功能：接口，重新授权
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTemp As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim lngCount As Long
    Dim objCollection As New Collection
    Dim strKey As String
    
    
    mstrFailInfo = ""
    
    mstrSQL = "select Distinct P.对象,P.所有者,P.权限 " & _
              "  from  zlProgPrivs p,zlRolegrant R " & _
              "  Where nvl(p.系统, 0) = nvl(r.系统, 0) And p.序号 = r.序号 And p.功能 = r.功能 And R.角色='" & strRole & "'"

    rsTemp.Open mstrSQL, mcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.BOF = False Then
        mstrSQL = "select Upper(Owner) As 所有者,Upper(table_name) As 对象,Upper(Privilege) As 权限 from role_tab_privs where role='" & strRole & "' And table_name Not Like 'BIN$%'"
        rs.Open mstrSQL, mcnOracle
        If rs.BOF = False Then
            Do While Not rs.EOF
                objCollection.Add "1", "K" & rs("所有者").value & "'" & rs("对象").value & "'" & rs("权限").value
                rs.MoveNext
            Loop
        End If
                
        Do While Not rsTemp.EOF
            strKey = "K" & UCase(rsTemp("所有者").value) & "'" & UCase(rsTemp("对象").value) & "'" & UCase(rsTemp("权限").value)
            
            If IsCollectionElement(objCollection, strKey) = False Then
                mstrSQL = "Grant " & rsTemp("权限").value & " on " & rsTemp("所有者").value & "." & rsTemp("对象").value & " to " & strRole
                Call ExecuteSQL(mstrSQL)
            End If
            
            rsTemp.MoveNext
        Loop

        '授予一些特殊表的权限
        '------------------------------------------------------------------------------------------------------------------
        Call GrantSpecialPrivilege(strRole, strOwner)
    End If
    
    ReGrantPrivilege = True
    
End Function

Public Function GrantPrivilege() As Boolean
    '******************************************************************************************************************
    '功能：接口函数，授予对象权限
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim blnReturn As Boolean
    On Error GoTo errHand
    
    '1 对数据表进行重新授权
    '------------------------------------------------------------------------------------------------------------------
    blnReturn = GetDictGrantObjectData(mrsTable)
    
    '2 对函数进行重新授权
    '------------------------------------------------------------------------------------------------------------------
    If blnReturn Then blnReturn = GetFunctionGrantObjectData(mrsFunction, mrsFuncPars)
        
    '3 对模块进行重新授权
    '------------------------------------------------------------------------------------------------------------------
    If blnReturn Then blnReturn = GetMoudalGrantObjectData(mrsGranded, mrsPrivs)

    '4 执行授权
    '------------------------------------------------------------------------------------------------------------------
    If blnReturn Then blnReturn = SaveGrant(mrsGranded, mrsPrivs)
    If blnReturn Then blnReturn = ExecuteGrantObject(mobjColGrantObject)
    If blnReturn Then blnReturn = GrantSpecialPrivilege(mstrRole, mstrOwner)
    
    GrantPrivilege = blnReturn
    
    Exit Function
    
errHand:
    MsgBox err.Description
End Function

Private Function GrantSpecialPrivilege(ByVal strRole As String, strOwner() As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    GrantSpecialPrivilege = GrantSpecialToRole(mcnOracle, strRole, mblnHaveBase, strOwner)

    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function GetDictRevokeObjectData(ByVal rsTable As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：读取函数的相关对象权限数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Long
    Dim intCount As Long
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    On Error GoTo errHand
        
    rsTable.Filter = "授权否 = 0"
    intCount = rsTable.RecordCount
    If intCount > 0 Then
    
        RaiseEvent BeforeProgress("收集字典取消权限", 100)
        
        intLoop = 0
        Do Until rsTable.EOF
            
            If rsTable("最初授权") = 1 Then
                
                Call AddPrivilegeData(rsTable("所有者") & "'" & rsTable("表名") & "'select", 2)
                Call AddPrivilegeData(rsTable("所有者") & "'" & rsTable("表名") & "'insert", 2)
                Call AddPrivilegeData(rsTable("所有者") & "'" & rsTable("表名") & "'update", 2)
                Call AddPrivilegeData(rsTable("所有者") & "'" & rsTable("表名") & "'delete", 2)
                        
            End If
            
            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
                    
            rsTable.MoveNext
        Loop
        
        RaiseEvent AfterProgress
    End If
    
    GetDictRevokeObjectData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Function GetDictGrantObjectData(ByVal rsTable As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：读取函数的相关对象权限数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Long
    Dim intCount As Long
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    On Error GoTo errHand

    rsTable.Filter = "授权否 = 1"
    intCount = rsTable.RecordCount
    If intCount > 0 Then
        
        RaiseEvent BeforeProgress("正在收集字典权限", 100)
        
        intLoop = 0
        Do Until rsTable.EOF
            Call AddPrivilegeData(rsTable("所有者") & "'" & rsTable("表名") & "'select", 1)
            Call AddPrivilegeData(rsTable("所有者") & "'" & rsTable("表名") & "'insert", 1)
            Call AddPrivilegeData(rsTable("所有者") & "'" & rsTable("表名") & "'update", 1)
            Call AddPrivilegeData(rsTable("所有者") & "'" & rsTable("表名") & "'delete", 1)

            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
                    
            rsTable.MoveNext
        Loop
        
        RaiseEvent AfterProgress
    End If
    
    GetDictGrantObjectData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Function GetFunctionRevokeObjectData(ByVal rsFunction As ADODB.Recordset, ByVal rsFuncPars As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：读取函数的相关对象权限数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strFail As String
    Dim blnAllow As Boolean
    Dim lngMoudal As Long
    Dim strFunc As String
    Dim intLoop As Long
    Dim intCount As Long
    Dim var对象 As Variant
    Dim strKey As String
    Dim i As Long
    Dim varTemp As Variant
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    On Error GoTo errHand
    
    rsFunction.Filter = "授权否 = 0"
    
    intCount = rsFunction.RecordCount
    If intCount > 0 Then
        RaiseEvent BeforeProgress("收集函数取消权限", 100)
        intLoop = 0
        
        Do Until rsFunction.EOF
            If rsFunction("最初授权") = 1 Then
                '只有以前是授过该数据表的角色，才需要执行取消功能
                Call AddPrivilegeData(rsFunction("所有者") & "'" & UCase(rsFunction("函数名")) & "'execute", 2)
                
                '读取参数中使用对象列表
                rsFuncPars.Filter = "系统 = " & mrsFunction("系统") & " and 函数名 = '" & mrsFunction("函数名") & "'"
                Do Until rsFuncPars.EOF
                    var对象 = Split(Replace(rsFuncPars("对象"), "|", ","), ",")
                    For i = LBound(var对象) To UBound(var对象)
                        If var对象(i) <> "" Then
                            varTemp = Split(var对象(i), ".")
                            Call AddPrivilegeData(CStr(varTemp(0)) & "'" & CStr(varTemp(1)) & "'select", 2)
                        End If
                    Next
                    rsFuncPars.MoveNext
                Loop
            End If
            
            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
                        
            rsFunction.MoveNext
        Loop
        RaiseEvent AfterProgress
    End If
        
    GetFunctionRevokeObjectData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function GetFunctionGrantObjectData(ByVal rsFunction As ADODB.Recordset, ByVal rsFuncPars As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：读取函数的相关对象权限数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strFail As String
    Dim blnAllow As Boolean
    Dim lngMoudal As Long
    Dim strFunc As String
    Dim intLoop As Long
    Dim intCount As Long
    Dim var对象 As Variant
    Dim strKey As String
    Dim i As Long
    Dim varTemp As Variant
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    On Error GoTo errHand
    
    rsFunction.Filter = "授权否 = 1"
    
    intCount = rsFunction.RecordCount
    If intCount > 0 Then
        RaiseEvent BeforeProgress("正在收集函数权限", 100)
        intLoop = 0
        Do Until mrsFunction.EOF
'            '只有以前是授过该数据表的角色，才需要执行取消功能
            Call AddPrivilegeData(rsFunction("所有者") & "'" & UCase(rsFunction("函数名")) & "'execute", 1)
            
            '读取参数中使用对象列表
            rsFuncPars.Filter = "系统 = " & rsFunction("系统") & " and 函数名 = '" & rsFunction("函数名") & "'"
            Do Until rsFuncPars.EOF
                var对象 = Split(Replace(rsFuncPars("对象"), "|", ","), ",")
                For i = LBound(var对象) To UBound(var对象)
                    If var对象(i) <> "" Then
                        varTemp = Split(var对象(i), ".")
                        Call AddPrivilegeData(CStr(varTemp(0)) & "'" & CStr(varTemp(1)) & "'select", 1)
                    End If
                Next
                rsFuncPars.MoveNext
            Loop

            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
                        
            rsFunction.MoveNext
        Loop
        RaiseEvent AfterProgress
    End If
    
    GetFunctionGrantObjectData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Function GetMoudalRevokeObjectData(ByVal rsGranded As ADODB.Recordset, ByVal rsPrivs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：读取模块的相关对象权限数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim blnAllow As Boolean
    Dim lngMoudal As Long
    Dim strFunc As String
    Dim intLoop As Long
    Dim intCount As Long
    Dim strKey As String
    Dim lngSys As Long
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    On Error GoTo errHand
    
    rsPrivs.Sort = "系统,序号,功能"
    intCount = rsPrivs.RecordCount
    If intCount > 0 Then
        RaiseEvent BeforeProgress("收集模块取消权限", 100)
    
        Do While Not rsPrivs.EOF
            If lngSys <> rsPrivs("系统").value Or lngMoudal <> rsPrivs("序号").value Or strFunc <> rsPrivs("功能").value Then
                lngSys = rsPrivs("系统").value
                lngMoudal = rsPrivs("序号").value
                strFunc = IIf(IsNull(rsPrivs("功能").value), "", rsPrivs("功能").value)
                
                blnAllow = False
                
                strKey = "K" & lngSys & "'" & lngMoudal & "'" & strFunc & "'0"
                If IsCollectionElement(mobjCollectionMoudal, strKey) = True Then
                    strKey = "K" & lngSys & "'" & lngMoudal & "'" & strFunc
                    blnAllow = IsCollectionElement(mobjCollectionGranded, strKey)
                End If
                
            End If
            
            If blnAllow = True Then
                Call AddPrivilegeData(rsPrivs("所有者").value & "'" & rsPrivs("对象").value & "'" & rsPrivs("权限").value, 2)
            End If
            
            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
                
            rsPrivs.MoveNext
        Loop
        
        RaiseEvent AfterProgress
        
    End If
    GetMoudalRevokeObjectData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
End Function

Private Function GetMoudalGrantObjectData(ByVal rsGranded As ADODB.Recordset, ByVal rsPrivs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：读取模块的相关对象权限数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim blnAllow As Boolean
    Dim lngMoudal As Long
    Dim strFunc As String
    Dim intLoop As Long
    Dim intCount As Long
    Dim strKey As String
    Dim lngSys As Long
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    On Error GoTo errHand
    
    rsPrivs.Sort = "系统,序号,功能"
    intCount = rsPrivs.RecordCount
    If intCount > 0 Then
        RaiseEvent BeforeProgress("正在收集模块权限", 100)
                
        Do While Not rsPrivs.EOF
            If lngSys <> rsPrivs("系统").value Or lngMoudal <> rsPrivs("序号").value Or strFunc <> rsPrivs("功能").value Then
                
                lngSys = rsPrivs("系统").value
                lngMoudal = rsPrivs("序号").value
                strFunc = IIf(IsNull(rsPrivs("功能").value), "", rsPrivs("功能").value)

                strKey = "K" & lngSys & "'" & lngMoudal & "'" & strFunc & "'1"
                blnAllow = IsCollectionElement(mobjCollectionMoudal, strKey)

            End If
            
            If blnAllow = True Then
                Call AddPrivilegeData(rsPrivs("所有者").value & "'" & rsPrivs("对象").value & "'" & rsPrivs("权限").value, 1)
            End If
            
            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
            
            rsPrivs.MoveNext
        Loop
        
        RaiseEvent AfterProgress
    
    End If
    
    GetMoudalGrantObjectData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
End Function

Private Function ExecuteRevokeObject(ByVal objColGrangObject As Collection) As Boolean
    '******************************************************************************************************************
    '功能：执行授权
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Long
    Dim intCount As Long
    Dim strTemp As String
    Dim varTemp As Variant
    Dim rs As New ADODB.Recordset
    Dim strFail As String
    Dim objCollection As New Collection
    Dim strKey As String
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    On Error GoTo errHand
    
    intCount = objColGrangObject.Count
    If intCount > 0 Then
        
        RaiseEvent BeforeProgress("正在取消角色权限", 100)
        
        mstrSQL = "select Upper(Owner) As 所有者,Upper(table_name) As 对象,Upper(Privilege) As 权限 from role_tab_privs where role='" & mstrRole & "' And table_name Not Like 'BIN$%'"
        rs.Open mstrSQL, mcnOracle
        If rs.BOF = False Then
            Do While Not rs.EOF
                objCollection.Add "1", "K" & rs("所有者").value & "'" & rs("对象").value & "'" & rs("权限").value
                rs.MoveNext
            Loop
        End If
        
        For intLoop = 1 To intCount
            strTemp = objColGrangObject.Item(intLoop)
            varTemp = Split(strTemp, "'")
            
            strKey = "K" & UCase(CStr(varTemp(0))) & "'" & UCase(CStr(varTemp(1))) & "'" & UCase(CStr(varTemp(2)))

            If IsCollectionElement(objCollection, strKey) = True Then
                Call ExecuteSQL("revoke " & CStr(varTemp(2)) & " on " & CStr(varTemp(0)) & "." & CStr(varTemp(1)) & " from " & mstrRole)
                
                If err <> 0 Then
                    strFail = strFail & vbCrLf & CStr(varTemp(2)) & " on " & CStr(varTemp(0)) & "." & CStr(varTemp(1))
                    err.Clear
                End If
            End If
            
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
            
        Next
        
        RaiseEvent AfterProgress
   End If
    
    ExecuteRevokeObject = True
    
    Exit Function
   '------------------------------------------------------------------------------------------------------------------
errHand:
   MsgBox err.Description
End Function

Private Function ExecuteGrantObject(ByVal objColGrangObject As Collection)
    '******************************************************************************************************************
    '功能：执行授权
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Long
    Dim intCount As Long
    Dim strTemp As String
    Dim varTemp As Variant
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strKey As String
    Dim blnAllow As Boolean
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    Dim objCollection As New Collection
    
    On Error GoTo errHand
    
    intCount = objColGrangObject.Count
    If intCount > 0 Then
        
        RaiseEvent BeforeProgress("正在授予角色权限", 100)
        
        mstrSQL = "select Upper(Owner) As 所有者,Upper(table_name) As 对象,Upper(Privilege) As 权限 from role_tab_privs where role='" & mstrRole & "' And table_name Not Like 'BIN$%'"
        rs.Open mstrSQL, mcnOracle
        If rs.BOF = False Then
            Do While Not rs.EOF
                objCollection.Add "1", "K" & rs("所有者").value & "'" & rs("对象").value & "'" & rs("权限").value
                rs.MoveNext
            Loop
        End If
        
        For intLoop = 1 To intCount
            strTemp = objColGrangObject.Item(intLoop)
            varTemp = Split(strTemp, "'")
            strKey = "K" & UCase(CStr(varTemp(0))) & "'" & UCase(CStr(varTemp(1))) & "'" & UCase(CStr(varTemp(2)))
            
            If IsCollectionElement(objCollection, strKey) = False Then
                
                strSQL = "grant " & CStr(varTemp(2)) & " on " & CStr(varTemp(0)) & "." & CStr(varTemp(1)) & " to " & mstrRole
                
                On Error Resume Next
                err = 0
                Call ExecuteSQL("Grant " & CStr(varTemp(2)) & " on " & CStr(varTemp(0)) & "." & CStr(varTemp(1)) & " to " & mstrRole)
                
                
                If err <> 0 Then
                    mstrFailInfo = mstrFailInfo & vbCrLf & CStr(varTemp(2)) & " on " & CStr(varTemp(0)) & "." & CStr(varTemp(1))
                    err.Clear
                End If
                
                On Error GoTo errHand
            End If
            
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
        Next
        
        RaiseEvent AfterProgress
    End If
    
    ExecuteGrantObject = True
    
    Exit Function
   '------------------------------------------------------------------------------------------------------------------
errHand:
   MsgBox err.Description
'   Resume
End Function

Private Function IsCollectionElement(ByVal objCollection As Collection, ByVal strKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTemp As String
    
    On Error GoTo errHand
    err = 0
    strTemp = objCollection.Item(strKey)
    IsCollectionElement = (err = 0)
    
    Exit Function
    
errHand:
    
End Function

Private Function DeleteGrant(ByVal rsGranded As ADODB.Recordset, ByVal rsPrivs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：取消以前授予的对模块的访问权限
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Long
    Dim intCount As Long
    Dim strKey As String
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    Dim strTemp As String
    
    On Error GoTo errHand
    
    mrsModule.Filter = "授权否 = 0 "
    intCount = mrsModule.RecordCount
    
    If intCount > 0 Then
        
        RaiseEvent BeforeProgress("正在删除角色权限", 100)
        
        Do Until mrsModule.EOF

            strKey = "K" & mrsModule("系统").value & "'" & mrsModule("序号").value & "'" & mrsModule("功能").value
            If IsCollectionElement(mobjCollectionGranded, strKey) = True Then
                '以前有授过的，清除
                If strTemp <> "" Then strTemp = strTemp & "''"
                strTemp = strTemp & mrsModule("系统").value & "''" & mrsModule("序号").value & "''" & mrsModule("功能").value

                If LenB(strTemp) > 1800 Then
                    mstrSQL = "zl_zlRoleGrant_BatchDelete('" & mstrRole & "','" & strTemp & "')"
                    Call ExecuteProcedure(mstrSQL, "授权")
                    strTemp = ""
                End If
                
            End If
            
            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
                        
            mrsModule.MoveNext
        Loop
        
        If strTemp <> "" Then
            mstrSQL = "zl_zlRoleGrant_BatchDelete('" & mstrRole & "','" & strTemp & "')"
            Call ExecuteProcedure(mstrSQL, "授权")
        End If
        
        RaiseEvent AfterProgress
        
    End If
    
    DeleteGrant = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Function SaveGrant(ByVal rsGranded As ADODB.Recordset, ByVal rsPrivs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Long
    Dim intCount As Long
    Dim strKey As String
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    Dim strTemp As String
    
    On Error GoTo errHand
    
    mrsModule.Filter = "授权否 = 1 "
    intCount = mrsModule.RecordCount
    If intCount > 0 Then
        
        RaiseEvent BeforeProgress("正在保存角色权限", 100)
        
        Do Until mrsModule.EOF

            strKey = "K" & mrsModule("系统").value & "'" & mrsModule("序号").value & "'" & mrsModule("功能").value
            If IsCollectionElement(mobjCollectionGranded, strKey) = False Then
                '以前没授过的，新授
                
                If strTemp <> "" Then strTemp = strTemp & "''"
                strTemp = strTemp & IIf(mrsModule("系统") = 0, "null", mrsModule("系统")) & "''" & mrsModule("序号").value & "''" & mrsModule("功能").value
                
                If LenB(strTemp) > 1800 Then
                    mstrSQL = "zl_zlRoleGrant_BatchInsert('" & mstrRole & "','" & strTemp & "')"
                    Call ExecuteProcedure(mstrSQL, "授权")
                    strTemp = ""
                End If

'                '清空错误，因为可能插入语句不能正常执行(由于唯一约束)
'                Err.Clear
            End If
            
            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
                        
            mrsModule.MoveNext
        Loop
        
        If strTemp <> "" Then
            mstrSQL = "zl_zlRoleGrant_BatchInsert('" & mstrRole & "','" & strTemp & "')"
            Call ExecuteProcedure(mstrSQL, "授权")
        End If
                
        RaiseEvent AfterProgress
        
    End If
    
    SaveGrant = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Sub AddPrivilegeData(ByVal strPrivilege As String, Optional ByVal bytMode As Byte = 1)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error Resume Next
    
    Select Case bytMode
    Case 1
        mobjColGrantObject.Add strPrivilege, "K" & strPrivilege
    Case 2
        mobjColRevokeObject.Add strPrivilege, "K" & strPrivilege
    End Select
End Sub

Private Sub ExecuteSQL(ByVal strSQL As String)
    mcnOracle.Execute strSQL
End Sub

Private Sub Class_Initialize()
    Set mobjCollectionMoudal = New Collection
    Set mobjCollectionGranded = New Collection
    Set mobjColGrantObject = New Collection
    Set mobjColRevokeObject = New Collection
    
    Set mrsPrivs = New ADODB.Recordset
    Set mrsGranded = New ADODB.Recordset           '保存以前授权的模块
    Set mrsFuncPars = New ADODB.Recordset          '函数参数所访问的表
End Sub

Private Sub Class_Terminate()
    Set mobjCollectionMoudal = Nothing
    Set mobjCollectionGranded = Nothing
    Set mobjColGrantObject = Nothing
    Set mobjColRevokeObject = Nothing
    
    Set mrsPrivs = Nothing
    Set mrsGranded = Nothing
    Set mrsFuncPars = Nothing
    Set mrsModule = Nothing
    Set mrsTable = Nothing
    Set mrsFunction = Nothing
    Set mcnOracle = Nothing
End Sub
