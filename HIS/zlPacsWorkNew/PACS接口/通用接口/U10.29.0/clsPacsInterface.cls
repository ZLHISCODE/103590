VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacsInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'######################################################################################################################


'自定义数据表
Public Type TCusTable
    strDatas() As String
    strColumns() As String
End Type


'错误显示类型
Public Enum TErrorShowType
    estNoDisplay = 1
    estShowMsg = 2
End Enum


'查询病人信息的条件类型
Public Enum TPatientWhereType
    pwtPatientId = 1   '病人ID
    pwtInHospital = 2  '住院号
    pwtOutPatient = 3  '门诊号
    pwtSickCard = 4    '就诊卡号
    pwtIdCard = 5      '身份证号
    pwtHealthNum = 6   '体检号
    pwtPatientName = 7 '病人姓名
End Enum


'申请单查询条件类型
Public Enum TRequestWhereType
    rwtPatientId = 1   '病人ID
    rwtInHospital = 2  '住院号
    rwtOutPatient = 3  '门诊号
    rwtSickCard = 4    '就诊卡号
    rwtIdCard = 5      '身份证号
    rwtHealthNum = 6   '体检号
    rwtPatientName = 7 '病人姓名
    rwtAdviceId = 8    '医嘱ID
End Enum




Private mblnTrans As Boolean
Private mstrOwner As String '系统所有者
Private mstrError As String '错误休息
Private mlngSys As Long '系统号
Private mstrNullValue As String  '空的替换值
Private mstrSplitChar As String    '分割符
Private mCurTable As TCusTable '当前表对象
Private mErrShowType As TErrorShowType  '错误消息显示方式




'######################################################################################################################


'错误消息的显示类型...
'Property Get ErrShowType() As TErrorShowType
'    ErrShowType = mErrShowType
'End Property
'
'Property Let ErrShowType(value As TErrorShowType)
'    mErrShowType = value
'End Property

Public Function GetErrShowType() As TErrorShowType
    GetErrShowType = mErrShowType
End Function


'返回值分隔符----------------------
'Property Get SplitChar() As String
'    SplitChar = mstrSplitChar
'End Property
'
'Property Let SplitChar(value As String)
'    mstrSplitChar = value
'End Property

Public Function GetSplitChar() As String
    GetSplitChar = mstrSplitChar
End Function


'系统号-----------------------
'Property Get SysNo() As Long
'    SysNo = mlngSys
'End Property
'
'Property Let SysNo(value As Long)
'    mlngSys = value
'End Property

Public Function GetSysNo() As Long
    GetSysNo = mlngSys
End Function


'系统所有者---------------------
'Property Get SysOwner() As String
'    SysOwner = mstrOwner
'End Property
'
'Property Let SysOwner(value As String)
'    mstrOwner = value
'End Function

Public Function GetSysOwner() As String
    GetSysOwner = mstrOwner
End Function


'查询的数据列-------------------
Property Get Tables() As TCusTable
    Tables = mCurTable
End Property


'空的替换值----------------------
'Property Get NullValue() As String
'    NullValue = mstrNullValue
'End Property
'
'Property Let NullValue(value As String)
'    mstrNullValue = value
'End Property

Public Function GetNullValue() As String
    GetNullValue = mstrNullValue
End Function


Public Function InitInterface(ByVal strServerName As String, _
    ByVal strUserName As String, _
    ByVal strUserPwd As String, _
    ByVal SysNo As Long, _
    ByVal SysOwner As String, _
    ByVal NullValue As String, _
    ByVal SplitChar As String, _
    ByVal errType As TErrorShowType) As Boolean
    '******************************************************************************************************************
    '功能：初始化接口对象
    '参数： strServerName -- ORACLE服务名
    '       strUserName -- HIS中的用户名
    '       strUserPwd -- HIS中的密码
    '       SysNo -- 系统号，HIS的默认系统号,一般是100
    '       SysOwner -- HIS数据库的所有者，一般是ZLHIS
    '       NullValue -- 空值传
    '       SplitChar -- 分隔符，用来分隔查询返回的数据，因此分隔符必须是查询返回数据中所不包含的字符，否则查询返回的表中，字段会出现错乱。一般可用“~”
    '       errType --  是否显示错误提示对话框
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mlngSys = SysNo
    mstrOwner = SysOwner
    mstrNullValue = NullValue
    mstrSplitChar = SplitChar
    mErrShowType = errType
    
    mstrError = ""
    
    InitInterface = OraDataOpen(strServerName, strUserName, strUserPwd)
    
    Exit Function
errHand:
    mstrError = Err.Description
        
End Function


Private Function CheckConnect() As Boolean
    '******************************************************************************************************************
    '功能：检查是否成功连接数据库
    '参数：
    '返回：
    '说明：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    mstrError = ""
    CheckConnect = False
        
    If Not (gcnOracle Is Nothing) Then
        If gcnOracle.State = adStateOpen Then
            CheckConnect = True
            Exit Function
        End If
    End If
    
    mstrError = "数据库连接未连接成功或丢失！"
    
    Exit Function
errHand:
    mstrError = Err.Description

End Function


Public Function GetLastError() As String
    '******************************************************************************************************************
    '功能：取得错误消息
    '参数：
    '返回：
    '说明：
    '******************************************************************************************************************
    GetLastError = mstrError
    
    If mErrShowType = estShowMsg Then
        MsgBox mstrError, vbCritical, gstrSysName
    End If
    
End Function


Public Function BeginTrans() As Boolean
    '******************************************************************************************************************
    '功能：开始事务
    '参数：
    '返回：
    '说明：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mstrError = ""
    BeginTrans = False
    
    If CheckConnect = False Then Exit Function
     
    If mblnTrans = False Then
        mblnTrans = True
        gcnOracle.BeginTrans
    End If
    
    BeginTrans = True
    
    Exit Function
errHand:
    mstrError = Err.Description
End Function


Public Function CommitTrans() As Boolean
    '******************************************************************************************************************
    '功能：提交事务
    '参数：
    '返回：
    '说明：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    CommitTrans = False
    mstrError = ""
    
    If CheckConnect = False Then Exit Function
        
    If mblnTrans Then
        gcnOracle.CommitTrans
        mblnTrans = False
    End If
    
    CommitTrans = True
    
    Exit Function
errHand:
    mstrError = Err.Description
End Function


Public Function RollbackTrans() As Boolean
    '******************************************************************************************************************
    '功能：回滚事务
    '参数：
    '返回：
    '说明：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    RollbackTrans = False
    mstrError = ""
    
    If CheckConnect = False Then Exit Function
    
    If mblnTrans Then
        gcnOracle.RollbackTrans
        mblnTrans = False
    End If
    
    RollbackTrans = True
    
    Exit Function
errHand:
    mstrError = Err.Description
End Function


Public Function ExecutePacsProcedure(ByVal strProcedureName As String) As Boolean
    '******************************************************************************************************************
    '功能：执行存储过程，不返回数据
    '参数： strProcedureName -- 需要执行的存储过程名称
    '返回：
    '说明：
    '******************************************************************************************************************
    On Error GoTo errHand
    
        Dim I As Integer
        Dim strSql As String
        
        mstrError = ""
        ExecutePacsProcedure = False
        
        If CheckConnect = False Then Exit Function
        
        strSql = IIf(mstrOwner <> "", mstrOwner & ".", "") & "b_PacsInterface." & strProcedureName
        Call ExecuteProcedure(strSql, gstrSysName)
        
        ExecutePacsProcedure = True
    
    Exit Function
errHand:
    mstrError = Err.Description

End Function


Public Function GetPacsCursor(ByVal strProcedureName As String, _
    Optional ByVal strFilterValue As String, _
    Optional blnIsNoParamerer As Boolean = False) As Boolean
    '******************************************************************************************************************
    '功能：获取PACS数据的入口
    '参数： strProcedureName -- 存储过程名称
    '       strFilterValue -- 过滤条件值
    '       blnIsNoParamerer -- 是否有附加参数
    '返回：
    '说明：
    '******************************************************************************************************************
    
    On Error GoTo errHand
        
        Dim I As Integer
        Dim strSql As String
        Dim rsTmp As New ADODB.Recordset
        
        Dim strResultItem As String
        Dim aryFilterValue(0 To 20) As Variant
       
        aryFilterValue(0) = Null: aryFilterValue(1) = Null: aryFilterValue(2) = Null: aryFilterValue(3) = Null: aryFilterValue(4) = Null
        aryFilterValue(5) = Null: aryFilterValue(6) = Null: aryFilterValue(7) = Null: aryFilterValue(8) = Null: aryFilterValue(9) = Null
        aryFilterValue(10) = Null: aryFilterValue(1) = Null: aryFilterValue(12) = Null: aryFilterValue(13) = Null: aryFilterValue(14) = Null
        aryFilterValue(15) = Null: aryFilterValue(16) = Null: aryFilterValue(17) = Null: aryFilterValue(18) = Null: aryFilterValue(19) = Null: aryFilterValue(20) = Null:
        
        mstrError = ""
        GetPacsCursor = False
        
        Erase mCurTable.strColumns()
        Erase mCurTable.strDatas()
        
        If CheckConnect = False Then Exit Function
               
        '设置查询条件
        If strFilterValue <> "" Then
            For I = 0 To UBound(Split(strFilterValue, mstrSplitChar))
                aryFilterValue(I) = Split(strFilterValue, mstrSplitChar)(I)
                If I = 20 Then I = UBound(Split(strFilterValue, mstrSplitChar))
            Next I
        End If
        
        '查询数据
        strSql = "b_PacsInterface." & strProcedureName
        
        If strFilterValue <> "" Then
            Set rsTmp = OpenCursor(gstrSysName, mstrOwner, strSql, aryFilterValue(0), aryFilterValue(1), aryFilterValue(2), aryFilterValue(3), aryFilterValue(4), aryFilterValue(5), _
                                aryFilterValue(6), aryFilterValue(7), aryFilterValue(8), aryFilterValue(9), aryFilterValue(10), aryFilterValue(11), aryFilterValue(12), aryFilterValue(13), aryFilterValue(14), _
                                aryFilterValue(15), aryFilterValue(16), aryFilterValue(17), aryFilterValue(18), aryFilterValue(19), aryFilterValue(20))
        Else
            If blnIsNoParamerer Then
                Set rsTmp = OpenCursor(gstrSysName, mstrOwner, strSql, Null)
            Else
                Set rsTmp = OpenCursor(gstrSysName, mstrOwner, strSql, "")
            End If
        End If
        
        '返回列名称
        ReDim Preserve mCurTable.strColumns(0)
        For I = 0 To rsTmp.Fields.Count - 1
            If mCurTable.strColumns(0) <> "" Then
                ReDim Preserve mCurTable.strColumns(UBound(mCurTable.strColumns) + 1)
            End If
            
            mCurTable.strColumns(UBound(mCurTable.strColumns)) = rsTmp.Fields(I).Name
        Next I
        
        If rsTmp.RecordCount <= 0 Then
            GetPacsCursor = True
            Exit Function
        End If
        
        '返回数据
        ReDim Preserve mCurTable.strDatas(0)
        Do While Not rsTmp.EOF
            
            strResultItem = Nvl(rsTmp.Fields(0).Value, mstrNullValue)
            For I = 1 To rsTmp.Fields.Count - 1
                strResultItem = strResultItem & mstrSplitChar & Nvl(rsTmp.Fields(I).Value, mstrNullValue)
            Next I
            
            If mCurTable.strDatas(0) <> "" Then
                ReDim Preserve mCurTable.strDatas(UBound(mCurTable.strDatas()) + 1)
            End If
            
            mCurTable.strDatas(UBound(mCurTable.strDatas())) = strResultItem
            
            rsTmp.MoveNext
        Loop
    
        
        GetPacsCursor = True
    
    Exit Function
errHand:
    mstrError = Err.Description

End Function




Public Function GetAdoData(ByVal strProcedureName As String, _
    ByVal strFilterValue As String, _
    Optional blnIsNoParamerer As Boolean = False) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：获取ado数据
    '参数： strProcedureName -- 存储过程名称
    '       strFilterValue -- 过滤参数值
    '       blnIsNoParamerer -- 是否有附加参数
    '返回：
    '说明：
    '******************************************************************************************************************
    
    On Error GoTo errHand
        
        Dim I As Long
        Dim strSql As String
        Dim aryFilterValue(0 To 20) As String
        
        mstrError = ""
        Set GetAdoData = Nothing
        
        If CheckConnect = False Then Exit Function
               
        If strFilterValue <> "" Then
            For I = 0 To UBound(Split(strFilterValue, mstrSplitChar))
                aryFilterValue(I) = Split(strFilterValue, mstrSplitChar)(I)
                If I = 20 Then I = UBound(Split(strFilterValue, mstrSplitChar))
            Next I
        End If
        
        '查询数据
        strSql = "b_PacsInterface." & strProcedureName
        If strFilterValue <> "" Then
            Set GetAdoData = OpenCursor(gstrSysName, mstrOwner, strSql, aryFilterValue(0), aryFilterValue(1), aryFilterValue(2), aryFilterValue(3), aryFilterValue(4), aryFilterValue(5), _
                                aryFilterValue(6), aryFilterValue(7), aryFilterValue(8), aryFilterValue(9), aryFilterValue(10), aryFilterValue(11), aryFilterValue(12), aryFilterValue(13), aryFilterValue(14), _
                                aryFilterValue(15), aryFilterValue(16), aryFilterValue(17), aryFilterValue(18), aryFilterValue(19), aryFilterValue(20))
        Else
            If blnIsNoParamerer Then
                Set GetAdoData = OpenCursor(gstrSysName, mstrOwner, strSql, Null)
            Else
                Set GetAdoData = OpenCursor(gstrSysName, mstrOwner, strSql, "")
            End If
        End If
        

    Exit Function
errHand:
    mstrError = Err.Description
End Function



Public Function GetRecordValueByColumnName(strDatas() As String, _
    strColumns() As String, _
    ByVal lngRecordIndex As Long, _
    ByVal strCurColumn As String) As String
    '******************************************************************************************************************
    '功能：获取记录值
    '参数： strDatas() -- 表数据集合
    '       strColumns() -- 列名称集合
    '       lngRecordIndex -- 记录行数索引
    '       strCurColumn -- 列名称
    '******************************************************************************************************************
    On Error GoTo errHand
        mstrError = ""
          
        GetRecordValueByColumnName = Split(strDatas(lngRecordIndex), mstrSplitChar)(GetColumnIndex(strColumns(), strCurColumn))
    
    Exit Function
errHand:
    mstrError = Err.Description
    
End Function


Public Function GetRecordValueByColumnIndex(strDatas() As String, _
    ByVal lngRecordIndex As Long, _
    ByVal lngColumnIndex As Long) As String
    '******************************************************************************************************************
    '功能：获取记录值
    '参数： strDatas() -- 表数据集合
    '       lngRecordIndex -- 记录行数索引
    '       lngColumnIndex -- 列索引
    '******************************************************************************************************************
    On Error GoTo errHand
        mstrError = ""
        
        GetRecordValueByColumnIndex = Split(strDatas(lngRecordIndex), mstrSplitChar)(lngColumnIndex)
    
    Exit Function
errHand:
    mstrError = Err.Description
    
End Function


Public Function GetCurValueByColumnName(ByVal lngRecordIndex As Long, ByVal strCurColumn As String) As String
    '******************************************************************************************************************
    '功能：从当前记录集中获取记录值
    '参数： lngRecordIndex -- 记录行数索引
    '       strCurColumn -- 列名称
    '******************************************************************************************************************
    On Error GoTo errHand
        mstrError = ""
          
        GetCurValueByColumnName = Split(mCurTable.strDatas(lngRecordIndex), mstrSplitChar)(GetColumnIndex(mCurTable.strColumns(), strCurColumn))
    
    Exit Function
errHand:
    mstrError = Err.Description
End Function


Public Function GetCurRecordCount() As Long
    '******************************************************************************************************************
    '功能：取得当前记录数量
    '******************************************************************************************************************
    On Error GoTo errHand
        mstrError = ""
        GetCurRecordCount = UBound(mCurTable.strDatas) + 1
        
        Exit Function
errHand:
    mstrError = Err.Description
End Function


Public Function GetCurColumnCount() As Long
    '******************************************************************************************************************
    '功能：取得当前数据列数量
    '******************************************************************************************************************
    On Error GoTo errHand
        mstrError = ""
        GetCurColumnCount = UBound(mCurTable.strColumns) + 1
        
        Exit Function
errHand:
    mstrError = Err.Description
End Function


Public Function GetCurColumnIndex(ByVal strCurColumn As String) As Long
    '******************************************************************************************************************
    '功能：从当前数据表中获取指定字段索引
    '参数： strCurColumn -- 列名称
    '******************************************************************************************************************
    On Error GoTo errHand
          
        Dim I As Long
        Dim lngFieldCount As Long
        
        mstrError = ""
        GetCurColumnIndex = -1
        
        '获取字段数量
        lngFieldCount = UBound(mCurTable.strColumns())
            
        '返回字段所在索引
        For I = 0 To lngFieldCount
            If UCase(mCurTable.strColumns(I)) = UCase(strCurColumn) Then
                GetCurColumnIndex = I
                Exit Function
            End If
        Next I
        
    Exit Function
errHand:
    mstrError = Err.Description
End Function


Public Function GetCurValueByColumnIndex(ByVal lngRecordIndex As Long, ByVal lngColumnIndex As Long) As String
    '******************************************************************************************************************
    '功能：从当前记录集中获取记录值
    '参数： lngRecordIndex -- 记录行数索引
    '       lngColumnIndex -- 列索引
    '******************************************************************************************************************
    On Error GoTo errHand
        mstrError = ""
        
        GetCurValueByColumnIndex = Split(mCurTable.strDatas(lngRecordIndex), mstrSplitChar)(lngColumnIndex)
    
    Exit Function
errHand:
    mstrError = Err.Description
End Function


Public Function GetCurRecordData(ByVal lngRecordIndex As Long) As String
    '******************************************************************************************************************
    '功能：获取当前行的数据
    '参数： lngRecordIndex -- 记录行数索引
    '******************************************************************************************************************
    On Error GoTo errHand
        mstrError = ""
        GetCurRecordData = mCurTable.strDatas(lngRecordIndex)
    
    Exit Function
errHand:
    mstrError = Err.Description
End Function


Public Function GetRecordData(strDatas() As String, ByVal lngRecordIndex As Long) As String
    '******************************************************************************************************************
    '功能：获取行的数据
    '参数： strDatas() -- 记录表数据集
    '       lngRecordIndex -- 记录行数索引
    '******************************************************************************************************************
    On Error GoTo errHand
        mstrError = ""
        GetRecordData = strDatas(lngRecordIndex)
    
    Exit Function
errHand:
    mstrError = Err.Description
End Function


Public Function GetRecordCount(strDatas() As String) As Long
    '******************************************************************************************************************
    '功能：获取返回的记录数量
    '参数： strDatas() -- 记录表数据集
    '******************************************************************************************************************
    On Error GoTo errHand
        mstrError = ""
        GetRecordCount = UBound(strDatas()) + 1
        
        Exit Function
errHand:
    mstrError = Err.Description
End Function



Public Function GetColumnCount(strColumns() As String) As Long
    '******************************************************************************************************************
    '功能：获取返回的字段数量
    '参数： strColumns() -- 列名称数据集
    '******************************************************************************************************************
    On Error GoTo errHand
        mstrError = ""
        GetColumnCount = UBound(strColumns()) + 1
        
        Exit Function
errHand:
    mstrError = Err.Description
End Function



Public Function GetColumnIndex(strColumns() As String, ByVal strCurColumn As String) As Integer
    '******************************************************************************************************************
    '功能：获取指定字段索引
    '参数： strColumns() -- 列名称数据集
    '       strCurColumn -- 列名称
    '******************************************************************************************************************
    On Error GoTo errHand
          
        Dim I As Long
        Dim lngFieldCount As Long
        
        mstrError = ""
        GetColumnIndex = -1
        
        '获取字段数量
        lngFieldCount = UBound(strColumns())
            
        '返回字段所在索引
        For I = 0 To lngFieldCount
            If UCase(strColumns(I)) = UCase(strCurColumn) Then
                GetColumnIndex = I
                Exit Function
            End If
        Next I
        
    Exit Function
errHand:
    mstrError = Err.Description
End Function


Public Function GetColumnName(strColumns() As String, ByVal columnIndex As Integer) As String
    '******************************************************************************************************************
    '功能：获取列名称
    '参数： strColumns() -- 列名称数据集
    '       columnIndex -- 列索引
    '******************************************************************************************************************
    On Error GoTo errHand
        mstrError = ""
        
        GetColumnName = strColumns(columnIndex)
        
    Exit Function
errHand:
    mstrError = Err.Description
End Function


Public Function GetCurColumnName(ByVal columnIndex As Integer) As String
    '******************************************************************************************************************
    '功能：获取列名称
    '参数： columnIndex -- 列索引
    '******************************************************************************************************************
    On Error GoTo errHand
        mstrError = ""
        
        GetCurColumnName = mCurTable.strColumns(columnIndex)
        
    Exit Function
errHand:
    mstrError = Err.Description
End Function


'基础数据
'######################################################################################################################

Public Function GetDeptItems(Optional ByVal strFilter As String) As Boolean
    '******************************************************************************************************************
    '功能：科室清单
    '参数： strFilter -- 过滤条件
    '返回：
    '说明：
    '******************************************************************************************************************
    
    GetDeptItems = GetPacsCursor("GetDeptItems", strFilter)
End Function


Public Function GetChargeTypes(Optional ByVal strFilter As String) As Boolean
    '******************************************************************************************************************
    '功能：检验/检查项目
    '参数： strFilter -- 过滤条件
    '返回：
    '说明：
    '******************************************************************************************************************
    
    GetChargeTypes = GetPacsCursor("GetChargeTypes", strFilter)
End Function


Public Function GetPacsItems(Optional ByVal strFilter As String) As Boolean
    '******************************************************************************************************************
    '功能：检验/检查项目
    '参数： strFilter -- 过滤条件
    '返回：
    '说明：
    '******************************************************************************************************************
    
    GetPacsItems = GetPacsCursor("GetPacsItems", strFilter)
End Function



Public Function GetAdviceItems(ByVal lngAdviceKey As Long) As Boolean
    '******************************************************************************************************************
    '功能：获取检查的详细方法
    '参数： lngAdviceKey -- 医嘱ID
    '返回：
    '说明：
    '******************************************************************************************************************
    GetAdviceItems = GetPacsCursor("GetAdviceItems", lngAdviceKey)
End Function

Public Function GetAdviceFees(ByVal lngAdviceKey As Long) As Boolean
    '******************************************************************************************************************
    '功能：获取检查的费用
    '参数： lngAdviceKey -- 医嘱ID
    '返回：
    '说明：
    '******************************************************************************************************************
    GetAdviceFees = GetPacsCursor("GetAdviceFees", lngAdviceKey)
End Function


Public Function GetPatientInfo(ByVal strQueryKey As String, ByVal lngWhereType As TPatientWhereType) As Boolean
    '******************************************************************************************************************
    '功能：根据指定的条件类型查询病人信息
    '参数： strQueryKey -- 查询的值
    '       lngWhereType -- 查询类型
    '返回：
    '说明：
    '******************************************************************************************************************
    
    GetPatientInfo = GetPacsCursor("GetPatient", lngWhereType & mstrSplitChar & strQueryKey)
End Function



Public Function GetRequestInfo(ByVal strQueryKey As String, ByVal lngWhereType As TRequestWhereType) As Boolean
    '******************************************************************************************************************
    '功能：根据指定条件获取申请单信息
    '参数： strQueryKey -- 查询的值
    '       lngWhereType -- 查询类型
    '返回：
    '说明：
    '******************************************************************************************************************
    
    GetRequestInfo = GetPacsCursor("GetRequestInfo", lngWhereType & mstrSplitChar & strQueryKey)
End Function


Public Function GetRequestInfo1(ByVal strStartDate As String, ByVal strEndDate As String, ByVal strStudyType As String) As Boolean
    '******************************************************************************************************************
    '功能：根据指定条件获取申请单信息
    '参数： strQueryKey -- 查询的值
    '       lngWhereType -- 查询类型
    '返回：
    '说明：
    '******************************************************************************************************************
    
    GetRequestInfo1 = GetPacsCursor("GetRequestInfo1", strStartDate & mstrSplitChar & strEndDate & mstrSplitChar & strStudyType)
End Function



Public Function GetRequestExecuteStatus(ByVal lngAdviceKey As Long) As Long
    '******************************************************************************************************************
    '功能：获取医嘱执行状态
    '参数： lngAdviceKey -- 医嘱ID
    '返回：
    '说明：
    '******************************************************************************************************************
    On Error GoTo errHand

    
    mstrError = ""
    
    GetRequestExecuteStatus = -10000
    If GetPacsCursor("GetRequestStatus", lngAdviceKey) Then
        
        If UBound(mCurTable.strDatas()) < 0 Then
            '当没有数据时，则触发异常
            'Call Err.Raise(-1, "GetRequestAdviceStatus", "没有查询到相关数据。")
            mstrError = "没有查询到相关数据。"
            Exit Function
        End If
        
        GetRequestExecuteStatus = VarNvl(GetRecordValueByColumnName(mCurTable.strDatas(), mCurTable.strColumns(), 0, "执行状态"), 0)
    End If
    
    
    Exit Function
errHand:
    mstrError = Err.Description
    
End Function


Public Function GetRequestAdviceStatus(ByVal lngAdviceKey As Long) As Long
    '******************************************************************************************************************
    '功能：获取医嘱状态
    '参数： lngAdviceKey -- 医嘱ID
    '返回：
    '说明：
    '******************************************************************************************************************
    On Error GoTo errHand

    
    mstrError = ""
    
    GetRequestAdviceStatus = -10000
    If GetPacsCursor("GetRequestStatus", lngAdviceKey) Then
        
        If UBound(mCurTable.strDatas()) < 0 Then
            '当没有数据时，则触发异常
            'Call Err.Raise(-1, "GetRequestAdviceStatus", "没有查询到相关数据。")
            mstrError = "没有查询到相关数据。"
            Exit Function
        End If
        
        GetRequestAdviceStatus = VarNvl(GetRecordValueByColumnName(mCurTable.strDatas(), mCurTable.strColumns(), 0, "医嘱状态"), 0)
    End If
        
    Exit Function
errHand:
    mstrError = Err.Description
End Function


Public Function GetRequestExeProcedureStatus(ByVal lngAdviceKey As Long) As Long
    '******************************************************************************************************************
    '功能：获取医嘱执行过程
    '参数： lngAdviceKey -- 医嘱ID
    '返回：
    '说明：
    '******************************************************************************************************************
    On Error GoTo errHand

    
    mstrError = ""
    
    GetRequestExeProcedureStatus = -10000
    If GetPacsCursor("GetRequestStatus", lngAdviceKey) Then
        
        If UBound(mCurTable.strDatas()) < 0 Then
            '当没有数据时，则触发异常
            'Call Err.Raise(-1, "GetRequestAdviceStatus", "没有查询到相关数据。")
            mstrError = "没有查询到相关数据。"
            Exit Function
        End If
        
        GetRequestExeProcedureStatus = VarNvl(GetRecordValueByColumnName(mCurTable.strDatas(), mCurTable.strColumns(), 0, "执行过程"), 0)
    End If
    
    
    Exit Function
errHand:
    mstrError = Err.Description
End Function


Public Function CancelRequest(ByVal lngAdviceKey As Long, Optional ByVal lngExecOne As Long) As Boolean
    '******************************************************************************************************************
    '功能：取消申请单的执行
    '参数： lngAdviceKey--医嘱ID，成套执行时是主医嘱ID，单独执行时是每一个部位医嘱或者主医嘱的ID
    '       lngExecOne -- 单独执行，0-成套执行；1-单独执行
    '返回：
    '说明：
    '******************************************************************************************************************

    CancelRequest = ExecutePacsProcedure("CancelRequest(" & lngAdviceKey & "," & lngExecOne & ")")
End Function


Public Function RecevieRequest( _
    ByVal lngAdviceKey As Long, _
    Optional ByVal strExeRoom As String, _
    Optional ByVal lngStudyNo As Long, _
    Optional ByVal strDevice As String, _
    Optional ByVal lngHeight As Long, _
    Optional ByVal lngWeight As Long, _
    Optional ByVal strStudyDoc As String, _
    Optional ByVal StrExeDate As Date, _
    Optional ByVal strExeDes As String, _
    Optional ByVal lngExecOne As Long _
) As Boolean
    '******************************************************************************************************************
    '功能：接收医嘱申请
    '参数： lngAdviceKey--医嘱ID，成套执行时是主医嘱ID，单独执行时是每一个部位医嘱或者主医嘱的ID
    '       strExeRoom -- 执行间
    '       lngStudyNo -- 检查号
    '       strDevice -- 检查设备
    '       lngHeight -- 身高
    '       lngWeight -- 体重
    '       strStudyDoc -- 检查技师
    '       StrExeDate -- 执行时间
    '       strExeDes -- 执行说明
    '       lngExecOne -- 单独执行，0-成套执行；1-单独执行
    '返回：
    '说明：
    '******************************************************************************************************************
    
    
    RecevieRequest = ExecutePacsProcedure("RecevieRequest(" & lngAdviceKey & ",'" & strExeRoom & "'," & lngStudyNo & ",'" & _
                                            strDevice & "'," & lngHeight & "," & lngWeight & ",'" & strStudyDoc & "'," & _
                                            To_Date(StrExeDate) & ",'" & strExeDes & "'," & lngExecOne & ")")
End Function



Public Function ModifyRequest( _
    ByVal lngAdviceKey As Long, _
    Optional ByVal strExeRoom As String, _
    Optional ByVal lngStudyNo As Long, _
    Optional ByVal strDevice As String, _
    Optional ByVal lngHeight As Long, _
    Optional ByVal lngWeight As Long, _
    Optional ByVal strStudyDoc As String, _
    Optional ByVal StrExeDate As Date, _
    Optional ByVal strExeDes As String, _
    Optional ByVal lngExecOne As Long _
) As Boolean
    '******************************************************************************************************************
    '功能：修改医嘱申请
    '参数： lngAdviceKey--医嘱ID，成套执行时是主医嘱ID，单独执行时是每一个部位医嘱或者主医嘱的ID
    '   strExeRoom --执行间
    '       lngStudyNo -- 检查号
    '       strDevice -- 检查设备
    '       lngHeight -- 身高
    '       lngWeight -- 体重
    '       strStudyDoc -- 检查技师
    '       StrExeDate -- 执行时间
    '       strExeDes -- 执行说明
    '       lngExecOne -- 单独执行，0-成套执行；1-单独执行
    '返回：
    '说明：
    '******************************************************************************************************************
    
    
    ModifyRequest = ExecutePacsProcedure("RecevieRequest(" & lngAdviceKey & ",'" & strExeRoom & "'," & lngStudyNo & ",'" & _
                                            strDevice & "'," & lngHeight & "," & lngWeight & ",'" & strStudyDoc & "'," & _
                                            To_Date(StrExeDate) & ",'" & strExeDes & "'," & lngExecOne & ")")
End Function


Public Function DeleteReport(ByVal lngAdviceKey As Long) As Boolean
    '******************************************************************************************************************
    '功能：删除报告
    '参数： lngAdviceKey -- 医嘱ID
    '返回：
    '说明：
    '******************************************************************************************************************
    
    DeleteReport = ExecutePacsProcedure("DeleteReport(" & lngAdviceKey & ")")
End Function


Public Function DeleteElectrocardioReport(ByVal lngAdviceKey As Long) As Boolean
    '******************************************************************************************************************
    '功能：删除心电报告
    '参数： lngAdviceKey -- 医嘱ID
    '返回：
    '说明：
    '******************************************************************************************************************
    
    DeleteElectrocardioReport = ExecutePacsProcedure("DeleteElectrocardioReport(" & lngAdviceKey & ")")
End Function


Public Function SendReport( _
    ByVal lngAdviceKey As Long, _
    ByVal strReportView As String, _
    ByVal strReportAdvice As String, _
    ByVal strReportDoctor As String, _
    Optional ByVal strAuditingDoctor As String _
) As Boolean
    '******************************************************************************************************************
    '功能：发送报告数据
    '参数： lngAdviceKey -- 医嘱ID
    '       strReportView -- 报告所见
    '       strReportAdvice -- 报告建议
    '       strReportDoctor -- 报告医生
    '       strAuditingDoctor -- 审核医生
    '返回：
    '说明：
    '******************************************************************************************************************
    
    If Trim(strAuditingDoctor) <> "" Then
        SendReport = ExecutePacsProcedure("SendReport(" & lngAdviceKey & ",'" & strReportView & "','" & strReportAdvice & "','" & _
                                            strReportDoctor & "','" & strAuditingDoctor & "')")
    Else
        SendReport = ExecutePacsProcedure("SendReport(" & lngAdviceKey & ",'" & strReportView & "','" & strReportAdvice & "','" & _
                                            strReportDoctor & "')")
    End If
End Function




Public Function SendElectrocardioReport(ByVal lngAdviceKey As Long, _
    ByVal strReportTitle As String, _
    ByVal strReportImgFiles As String, _
    ByVal strReportResult As String, _
    ByVal strReportAdvice As String, _
    ByVal strReportDoctor As String, _
    Optional ByVal strAuditingDoctor As String _
    ) As Boolean
    '******************************************************************************************************************
    '功能：发送心电报告数据
    '参数： lngAdviceKey -- 医嘱ID
    '       strReportTitle -- 报告标题
    '       strReportImgFiles -- 报告图像
    '       strReportResult -- 诊断结果
    '       strReportAdvice -- 诊断建议
    '       strReportDoctor -- 报告医生
    '       strAuditingDoctor -- 审核医生
    '返回：
    '说明：
    '******************************************************************************************************************
On Error GoTo errHand

    Dim aryFiles() As String
    Dim arySql() As String
    
    Dim rsTemp As ADODB.Recordset
    
    Dim I As Long
    Dim lngFileCount As Long
    Dim strSql As String
    
    
    
    '保存报告文本部分
    If Trim(strAuditingDoctor) <> "" Then
        SendElectrocardioReport = ExecutePacsProcedure("SendElectrocardioReport(" & lngAdviceKey & ",'" & _
                                                                    strReportTitle & "','" & _
                                                                    strReportResult & "','" & _
                                                                    strReportAdvice & "','" & _
                                                                    strReportDoctor & "','" & _
                                                                    strAuditingDoctor & "')")
    Else
        SendElectrocardioReport = ExecutePacsProcedure("SendElectrocardioReport(" & lngAdviceKey & ",'" & _
                                                                    strReportTitle & "','" & _
                                                                    strReportResult & "','" & _
                                                                    strReportAdvice & "','" & _
                                                                    strReportDoctor & "')")
    End If
    
    
    '保存报告图像部分
    aryFiles() = Split(strReportImgFiles, mstrSplitChar)
     
    On Error Resume Next
        lngFileCount = UBound(aryFiles())
        
        
    '如果没有文件将产生异常
    If lngFileCount < 0 Then Exit Function


    ReDim Preserve arySql(0)
    For I = LBound(aryFiles()) To lngFileCount
        If Trim(aryFiles(I)) <> "" Then
            strSql = "select " & IIf(mstrOwner <> "", mstrOwner & ".", "") & "b_PacsInterface.AddElectrocardioReportImage([1]) as 图像内容ID from dual"
            Set rsTemp = OpenSQLRecord(strSql, "添加图像记录", lngAdviceKey)
            
            Call zlBlobSave(6, rsTemp!图像内容ID, aryFiles(I), arySql())
        End If
    Next I
    
        
    If arySql(0) = "" Then
        mstrError = "不能取得图像流数据。"
        Exit Function
    End If
    
    
    Err = 0: On Error GoTo errRollback
    
    Call BeginTrans
    

    For I = LBound(arySql) To UBound(arySql)
        If Trim(arySql(I)) <> "" Then Call ExecuteProcedure(IIf(mstrOwner <> "", mstrOwner & ".", "") & arySql(I), gstrSysName)
    Next
    
    Call CommitTrans
    
    SendElectrocardioReport = True
    Exit Function
    
errRollback:
        Call RollbackTrans
        mstrError = "报告图像保存事务执行失败。"
    
    Exit Function
errHand:
    SendElectrocardioReport = False
    mstrError = Err.Description
End Function



'################################################################################################################
'## 功能：  将指定的文件保存到指定记录的LOB字段中
'##
'## 参数：  Action      :操作类型（用以区别是操作哪个表）
'##         KeyWord     :确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'##         strFile     :用户指定存放的文件名；不指定时，取当前路径产生文件名
'##
'## 返回：  成功返回True，失败返回False
'##
'## 说明：  Action取值说明：
'##         0-病历标记图形；1-病历文件格式；2-病历文件图形；3-病历范文格式；4-病历范文图形；5-电子病历格式；6-电子病历图形；
'################################################################################################################
Private Function zlBlobSave(ByVal Action As Long, _
    ByVal KeyWord As String, _
    ByVal strFile As String, _
    ByRef arySql() As String) As Boolean
    
    Dim conChunkSize As Integer
    Dim strSql As String
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    Err = 0: On Error GoTo errHand
    
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        strText = Join(aryHex, "")
        
        If arySql(0) <> "" Then
            ReDim Preserve arySql(UBound(arySql()) + 1)
        End If
        
        arySql(UBound(arySql())) = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "','" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
    Next
    Close lngFileNum
    zlBlobSave = True
    Exit Function

errHand:
    Close lngFileNum
    zlBlobSave = False
    
    mstrError = Err.Description
End Function


Private Function GetReportId(ByVal lngAdviceKey As Long) As Long
    '******************************************************************************************************************
    '功能：获取病历ID
    '参数： lngAdviceKey -- 医嘱ID
    '返回：
    '说明：
    '******************************************************************************************************************
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    GetReportId = -1
    
    '获取病历ID
    strSql = "select 病历id from 病人医嘱报告 where 医嘱ID=[1]"
    Set rsTemp = OpenSQLRecord(strSql, gstrSysName, lngAdviceKey)
    
    If rsTemp.RecordCount <= 0 Then
        mstrError = "报告尚未创建。"
        Exit Function
    End If
    
    GetReportId = rsTemp!病历id
End Function


Public Function SendReportImages(ByVal lngAdviceKey As Long, ByVal strImgFiles As String) As Boolean
    '******************************************************************************************************************
    '功能：保存报告图像
    '参数： lngAdviceKey -- 医嘱ID
    '       strImgFiles -- 报告图文件
    '返回：
    '说明：
    '******************************************************************************************************************
    On Error GoTo errHand
        Dim I As Integer
        Dim strSql As String
        Dim lngReportId As Long
        Dim lngImgTableId As Long
        Dim lngSqlCound As Long
        
        Dim aryFiles() As String
        Dim arySql() As String
        Dim rsTemp As ADODB.Recordset
        
        SendReportImages = False
        
        aryFiles() = Split(strImgFiles, mstrSplitChar)
        
        '如果没有文件将产生异常
        If UBound(aryFiles()) < 0 Then Exit Function
        
        '获取病历ID
        lngReportId = GetReportId(lngAdviceKey)
        
        '获取插入报告图的图表ID
        strSql = "Select Id As 表格Id From 电子病历内容" & vbNewLine & _
                 " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
                 " Order By 对象序号"
        Set rsTemp = OpenSQLRecord(strSql, gstrSysName, lngReportId)
        
        If rsTemp.RecordCount <= 0 Then
            mstrError = "报告图尚未构造，请在电子病历设计中构造报告图像。"
            Exit Function
        End If
        
        lngImgTableId = rsTemp!表格Id
        
        '创建报告图结构
        strSql = IIf(mstrOwner <> "", mstrOwner & ".", "") & "ZL_影像报告图像_保存(" & lngImgTableId & "," & UBound(aryFiles()) + 1 & ")"
        Call ExecuteProcedure(strSql, gstrSysName)
        
        strSql = "Select Id,内容行次 From 电子病历内容 Where 父ID= [1] And substr(对象属性,1,1)='2' Order By 内容行次 "
        Set rsTemp = OpenSQLRecord(strSql, gstrSysName, lngImgTableId)
        
        '插入报告图像
        I = 0
        ReDim Preserve arySql(0)
        
        Do While Not rsTemp.EOF
            Call zlBlobSave(6, rsTemp!Id, aryFiles(I), arySql())
            
            I = I + 1
            rsTemp.MoveNext
        Loop
        
        If arySql(0) = "" Then
            mstrError = "不能取得图像流数据。"
            Exit Function
        End If
        
        
        Err = 0: On Error GoTo errRollback
        
        Call BeginTrans
        

        For lngSqlCound = LBound(arySql) To UBound(arySql)
            Call ExecuteProcedure(IIf(mstrOwner <> "", mstrOwner & ".", "") & arySql(lngSqlCound), gstrSysName)
        Next
        
        Call CommitTrans
        
        SendReportImages = True
        
        Exit Function
errRollback:
        Call RollbackTrans
        mstrError = "报告图像保存事务执行失败。"
    
    Exit Function
errHand:
    SendReportImages = False
    mstrError = Err.Description
    
End Function




Public Function SendReportAffix(ByVal lngAdviceKey As Long, ByVal strAffixFiles As String) As Boolean
    '******************************************************************************************************************
    '功能：保存报告附件
    '参数： lngAdviceKey -- 医嘱ID
    '       strAffixFiles -- 医嘱附件文件
    '返回：
    '说明：
    '******************************************************************************************************************
    On Error GoTo errHand
        Dim strSql As String
        Dim lngReportId As Long
        Dim lngSqlCound As Long
        
        Dim arySql() As String
        
        Dim objFile As File
        Dim objFSO As New Scripting.FileSystemObject
        Dim rsTemp As ADODB.Recordset
    
        SendReportAffix = False
        
        If Trim(strAffixFiles) = "" Then Exit Function
        
        '获取病历ID
        strSql = "select 病历id from 病人医嘱报告 where 医嘱ID=[1]"
        Set rsTemp = OpenSQLRecord(strSql, gstrSysName, lngAdviceKey)
        
        If rsTemp.RecordCount <= 0 Then
            mstrError = "报告尚未创建。"
            Exit Function
        End If
        
        lngReportId = rsTemp!病历id
        
        Set objFile = objFSO.GetFile(strAffixFiles)
        
        ReDim Preserve arySql(0)
        If Not zlBlobSave(8, lngReportId & ",10000", strAffixFiles, arySql()) Then
            mstrError = "附件保存失败。"
            Exit Function
        End If
        
        If arySql(0) = "" Then
            mstrError = "不能取得附件流数据。"
            Exit Function
        End If
        
        Err = 0: On Error GoTo errRollback
        Call BeginTrans
        
        '清除使用该接口添加的附件
        strSql = IIf(mstrOwner <> "", mstrOwner & ".", "") & "b_PacsInterface.ClearReportAffix(" & lngReportId & ",'PacsInterface')"
        Call ExecuteProcedure(strSql, gstrSysName)
        
        '添加报告附件记录
        strSql = IIf(mstrOwner <> "", mstrOwner & ".", "") & "b_PacsInterface.AddReportAffix(" & lngReportId & ",'" & objFile.Name & "'," & Format(objFile.Size / 1024, "0.00") & ",'PacsInterface')"
        Call ExecuteProcedure(strSql, gstrSysName)
        
        '添加附件
        For lngSqlCound = LBound(arySql) To UBound(arySql)
            Call ExecuteProcedure(IIf(mstrOwner <> "", mstrOwner & ".", "") & arySql(lngSqlCound), gstrSysName)
        Next
        
        Call CommitTrans
    
        SendReportAffix = True
        
        Exit Function
errRollback:
        Call RollbackTrans
        mstrError = "附件保存事务执行失败。"
        
    Exit Function
errHand:
    SendReportAffix = False
    mstrError = Err.Description

End Function





'######################################################################################################################

Private Sub Class_Initialize()
    On Error Resume Next
    
    If mblnTrans Then gcnOracle.RollbackTrans
    
    If gcnOracle.State = adStateOpen Then gcnOracle.Close
    Set gcnOracle = Nothing
    
    gstrSysName = "中联软件(PACS接口)"
    mstrOwner = "ZLHIS"
    mlngSys = 100
    
    mstrSplitChar = ","
    mstrNullValue = ""
    mErrShowType = estNoDisplay
End Sub


Private Sub Class_Terminate()
    If gcnOracle.State = adStateOpen Then gcnOracle.Close
    
    Set gcnOracle = Nothing
End Sub
