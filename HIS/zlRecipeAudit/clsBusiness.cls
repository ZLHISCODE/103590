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
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------
'说明: 为其他业务系统开放的接口类
'--------------------------------------------------------------------------------
Option Explicit

Public Enum RA_Side
    s_外部 = 1
    s_内部 = 2
End Enum

Public Enum RA_AutoAuditReturn
    aar_异常 = 0
    aar_未开启审方 = 1
    aar_未开启提醒 = 2
    aar_调整药品 = 3
    aar_忽略提醒 = 4
    aar_全部合格 = 5
    aar_处方审查条件未设置 = 6
End Enum

Private mblnInit As Boolean                                 'True已调用初始化方法（Init）；False未调用
Private mblnOutPatient As Boolean                           'True门诊病人；False住院病人；由外部程序调用Init方法传入
Private mrsRAItems As ADODB.Recordset                       '
Private mrsAuditSend As ADODB.Recordset                     '
Private mstrAutoAuditID As String                           '

Public Sub Init(ByVal cnMain As ADODB.Connection, ByVal blnOutPatient As Boolean, Optional ByVal lngSys As Long = 100)
'功能：初始化类模块
'参数：
'  cnMain：Oracle数据库连接对象
'  blnOutPatient：True门诊；False住院
    
    Dim blnParaBasis As Boolean
    Dim strSQL As String
    
    If cnMain Is Nothing Then
        MsgBox "数据库连接对象未创建！", vbInformation, gstrSysName
        Exit Sub
    End If
    If cnMain.State <> adStateOpen Then
        MsgBox "数据库连接对象未打开！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'clsBusiness类与clsRecipeAudit类应该是同一个数据库的连接，但用不同的连接对象
    'Set gcnBusiness = cnMain
    
    glngSys = lngSys
    mblnOutPatient = blnOutPatient
    If gcnOracle Is Nothing Then Set gcnOracle = cnMain
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '临床公共方法
    On Error Resume Next
    Set gobjPubAdvice = New zlPublicAdvice.clsPublicAdvice
    If Not gobjPubAdvice Is Nothing Then
        Call gobjPubAdvice.InitCommon(gcnOracle, glngSys)
    End If
    Err.Clear: On Error GoTo 0
    
    '审方项目依据
    blnParaBasis = (Val(zlDatabase.GetPara("处方审查依据", glngSys)) = 1)    '1-依据《处方点评管理规范》28项；2-依据《处方管理办法》7项
    
    '初始化mrsRAItems
    On Error GoTo errHandle
    strSQL = "Select ID, 类别, 编码, 简称, PASS结果 " & vbCr & _
             "From 处方审查项目 " & vbCr & _
             "Where (作废时间 Is Null Or To_Char(作废时间, 'yyyy') = '3000') And 类别 In ([1], 3, 4) " & vbCr & _
             "  And " & IIf(blnOutPatient, " 是否门诊启用 = 1 ", " 是否住院启用 = 1 ")
    Set mrsRAItems = zlDatabase.OpenSQLRecord(strSQL, IIf(blnOutPatient, "获取门诊审查项目", "获取住院审查项目"), IIf(blnParaBasis, 2, 1))
    
    '成功初始化
    mblnInit = True
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume Next
End Sub

Public Function AutoAudit( _
    ByVal frmOwner As Object, _
    ByVal bytMode As RA_Side, _
    ByVal strSubmitID As String, _
    ByVal lngClinicID As Long, _
    ByVal bytPatientMode As Byte, _
    ByVal lngPatientID As Long, _
    ByVal lngBillID As Long, _
    Optional ByRef strErr As String _
) As RA_AutoAuditReturn
'功能：自动审查
'参数：
'  frmOwner：宿主窗体对象
'  'bytClass：调用的业务；1-门诊业务；2-住院业务
'  bytMode：1-外部调用；2-内部调用；控制是否显示提示对话框和参数检查待处理
'  strSubmitID：待提交审查的给药途径医嘱ID（格式：给药途径医嘱ID[,给药途径医嘱ID[...]]）
'  lngClinicID：临床科室ID
'  bytPatientMode：病人来源；0-门诊；1-住院
'  lngPatientID：病人ID
'  lngBillID：主页ID或挂号ID
'  strErr（可选实参）：异常信息
'返回：
'  0-异常，外部程序终止流程或继续流程
'  1-未开启审方，外部程序继续流程
'  2-开启审方，但未开启不合格提醒，外部程序继续流程
'  3-开启审方，开启不合格提醒，医生操作“调整药品”，外部程序终止流程
'  4-开启审方，开启不合格提醒，医生操作“忽略提醒”，外部程序继续流程
'  5-开启审方，开启不合格提醒，审查没有不合格的项目，外部程序继续流程
'  6-门诊开启审方，开启不合格提醒，但“处方审查条件”未有任何设置，外部程序继续流程

    Dim frmWarning As frmRAWarning
    Dim blnParaRA As Boolean, blnParaWarning As Boolean
    Dim lngTmp As Long
    Dim strTmp As String, strIDs As String, strNG As String
    Dim strMedicalID As String
    Dim bytResult As Byte
    Dim arrSubmit As Variant, arrID As Variant
    
    gstrErrInfo = ""
    mstrAutoAuditID = ""
    
    '更新病人类型的变量值（由于临床用的全局对象变量）
    mblnOutPatient = (bytPatientMode = 0)
    
    If mblnInit = False Then
        strErr = "请先调用接口初始化的方法"
        AutoAudit = RA_AutoAuditReturn.aar_异常
        Exit Function
    End If
    
    '检查处方审查参数；0-门诊和住院都不启用；1-门诊启用，住院不启用；2-门诊不启用，住院启用；3-门诊和住院都启用
    lngTmp = Val(zlDatabase.GetPara("处方审查", glngSys))
    If mblnOutPatient Then
        '门诊业务
        blnParaRA = (lngTmp = 1 Or lngTmp = 3)
        strTmp = "未开启门诊处方审查，无需调用本接口！"
    Else
        '住院业务
        blnParaRA = (lngTmp = 2 Or lngTmp = 3)
        strTmp = "未开启住院处方审查，无需调用本接口！"
    End If
    If blnParaRA = False Then
        strErr = strTmp
        AutoAudit = RA_AutoAuditReturn.aar_未开启审方
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    '自动审查不合格提醒参数
    If mblnOutPatient Then
        '门诊
        blnParaWarning = (Val(zlDatabase.GetPara("提醒门诊医生不合格医嘱", glngSys)) = 1)
    Else
        '住院
        blnParaWarning = (Val(zlDatabase.GetPara("提醒住院医生不合格医嘱", glngSys)) = 1)
    End If
    
'    '取提交的医嘱ID
'    arrSubmit = Split(strSubmitID, ",")
'    For lngTmp = LBound(arrSubmit) To UBound(arrSubmit)
'        strIDs = strIDs & Val(arrSubmit(lngTmp)) & IIf(lngTmp >= UBound(arrSubmit), "", ",")
'    Next
    
    If strSubmitID = "" Then
        strErr = "未传入医嘱ID信息"
        AutoAudit = RA_AutoAuditReturn.aar_异常
        Exit Function
    End If
    
    '准备记录集和数据结构
    SetRecordsetStructure 1, mrsAuditSend
    
    If mblnOutPatient Then
        '检查门诊类处方审查条件
        lngTmp = IsOutpatientControl(strSubmitID)
        If lngTmp = 0 Or lngTmp = 2 Then
            AutoAudit = RA_AutoAuditReturn.aar_处方审查条件未设置
            Exit Function
        End If
    End If
    
    '审查项目
    With mrsRAItems
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            '审查处理
            bytResult = AuditDrug(.Fields, lngPatientID, IIf(mblnOutPatient, 1, 2), lngBillID, strSubmitID, strMedicalID)
            
            '组装数据
            If strMedicalID <> "" Then
                '有医嘱ID
                arrID = Split(strMedicalID, ";")
                For lngTmp = LBound(arrID) To UBound(arrID)
                    mrsAuditSend.AddNew
                    mrsAuditSend!审查项目ID = !ID
                    mrsAuditSend!编码 = !编码
                    mrsAuditSend!简称 = !简称
                    mrsAuditSend!审查结果 = bytResult
                    mrsAuditSend!医嘱ID = Val(arrID(lngTmp))
                    mrsAuditSend.Update
                Next
            Else
                '无医嘱ID，表示整批医嘱
                mrsAuditSend.AddNew
                mrsAuditSend!审查项目ID = !ID
                mrsAuditSend!编码 = !编码
                mrsAuditSend!简称 = !简称
                mrsAuditSend!审查结果 = bytResult
                mrsAuditSend!医嘱ID = 0
                mrsAuditSend.Update
            End If

            .MoveNext
        Loop
    End With
    
    '保存审查的给药途径医嘱ID
    mstrAutoAuditID = strSubmitID
    
    '审查不合格提醒
    If blnParaWarning And bytMode = RA_Side.s_外部 Then
        '开启提醒，并且外部调用
        With mrsAuditSend
            strNG = ""
            If .RecordCount > 0 Then .MoveFirst
            Do While .EOF = False
                If !审查结果 = 2 Then       '只处理不合格
                    strNG = strNG & zlstr.FormatString("[1],[2]|", !审查项目ID, NVL(!医嘱ID, "0"))
                End If
                .MoveNext
            Loop
            If strNG <> "" Then strNG = Left(strNG, Len(strNG) - 1)
        End With
        
        If strNG <> "" Then
            '显示不合格信息
            Set frmWarning = New frmRAWarning
            If frmWarning.ShowMe(strNG, frmOwner) Then
                AutoAudit = RA_AutoAuditReturn.aar_调整药品
            Else
                AutoAudit = RA_AutoAuditReturn.aar_忽略提醒
            End If
        Else
            AutoAudit = RA_AutoAuditReturn.aar_全部合格       '审查没有不合格的项目
        End If
    Else
        '关闭提醒或内部调用有审方数据
        AutoAudit = RA_AutoAuditReturn.aar_未开启提醒
    End If
    
    Exit Function
    
errHandle:
    strErr = gstrErrInfo
    AutoAudit = 0
End Function

Public Function BuildData(ByVal strSubmitID As String, ByVal lngClinicID As Long, _
    ByVal bytPatientMode As Byte, ByVal lngPatientID As Long, ByVal lngBillID As Long, _
    ByRef strNotSendIDs As String, _
    Optional ByRef strErr As String _
) As Boolean
'功能：临床工作站调用本方法，生成待审查的（门诊/住院）数据
'参数：
'  strSubmitID：待提交审查的给药途径医嘱ID（格式：给药途径医嘱ID[,给药途径医嘱ID[...]]）
'  lngClinicID：临床科室ID
'  bytPatientMode：0-门诊；1-住院
'  lngPatientID：病人ID
'  lngBillID：主页ID或挂号ID
'  strNotSendIDs（实参）：返回不能“发送”的给药途径药嘱ID（格式：给药途径医嘱ID,...）
'  strErr（实参）：异常信息
'返回：True成功；False失败

    Dim typAudit As RA_AutoAuditReturn
    Dim strSQL As String, strIDs As String
    Dim strTmp As String
    Dim lngRelationID As Long
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim lngStoreID As Long, lngAuditID As Long
    Dim arrID As Variant
    Dim colSQL As New Collection
    Dim strCannotSend As String, strSubmitNot As String
    Dim intPara As Integer
    
    '更新病人类型的变量值（由于临床用的全局对象变量）
    mblnOutPatient = (bytPatientMode = 0)
    
    If mblnInit = False Then
        strErr = "请先调用接口初始化的方法"
        BuildData = False
        MsgBox strErr
        Exit Function
    End If
    
    If Trim(strSubmitID) = "" Then
        strErr = "提交的相关ID为空"
        BuildData = False
        MsgBox strErr
        Exit Function
    End If
    
    '处方审查参数
    intPara = Val(zlDatabase.GetPara("处方审查", glngSys, , "0"))
    If mblnOutPatient Then
        If Not (intPara = 1 Or intPara = 3) Then
            '未开启门诊审方
            BuildData = True
'            MsgBox "未开启门诊审方"
            Exit Function
        End If
    Else
        If Not (intPara = 2 Or intPara = 3) Then
            '未开启住院审方
            BuildData = True
'            MsgBox "未开启住院审方"
            Exit Function
        End If
    End If
    
    '门诊的处方审查时机参数；1-处方发送前；2-处方配发药前
    intPara = Val(zlDatabase.GetPara("门诊审方时机", glngSys, , "1"))
    
    '门诊：处方发送前，调整长时间未审查的审方记录状态；药房配发药前，不调整
    If mblnOutPatient And intPara = 1 Then
        Call AuditSend(lngClinicID, strSubmitID, strTmp)
    End If
    
    On Error GoTo hErr
    
    '先分离已提交和未提交的相关ID
    strSQL = "Select Distinct d.Column_Value 相关id, c.状态, c.审查结果, c.Id 审方id " & vbNewLine & _
             "From 病人医嘱记录 A, 处方审查明细 B, 处方审查记录 C, Table(f_Num2list([1], ',')) D " & vbNewLine & _
             "Where d.Column_Value = a.相关id(+) And a.Id = b.医嘱id(+) And b.最后提交(+) = 1 And b.审方id = c.Id(+) " & vbNewLine & _
             "    And a.诊疗类别 In ('5', '6', '7') "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "分离相关ID", strSubmitID)
    With rsTemp
        Do While .EOF = False
            If IsNull(!审方id) Then
                '未提交的相关id
                strSubmitNot = strSubmitNot & "," & CStr(!相关id)
            Else
'                '门诊-发送前审方。
'                If intPara = 1 And mblnOutPatient Then
                    Select Case NVL(!状态, 0)
                    Case 0          '待审
                        strCannotSend = strCannotSend & "," & CStr(!相关id)
                    Case 1          '已审
                        If NVL(!审查结果, 0) = 2 Then   '2-不合格
                            strCannotSend = strCannotSend & "," & CStr(!相关id)
                        End If
                    Case 2, 3       '超时免审
                    Case 11         '已审被撤销
                        strCannotSend = strCannotSend & "," & CStr(!相关id)
                    End Select
'                End If
            End If
            
            .MoveNext
        Loop
        .Close
    End With
    '未提交的相关ID
    If strSubmitNot <> "" Then strSubmitNot = Mid(strSubmitNot, 2)
    '已提交，但未审查或审查不合格的相关ID
    If strCannotSend <> "" Then strCannotSend = Mid(strCannotSend, 2)
    
    '处理未提交的医嘱
    If strSubmitNot <> "" Then
        '审方工作站是否开启审方事务
        If mblnOutPatient Then
            '门诊
'            strSQL = "Select Count(1) Rec " & vbNewLine & _
'                     "From 处方审查参数 " & vbNewLine & _
'                     "Where Nvl(服务对象, 0) = 0 And ',' || 来源科室 || ',' Like '%,' || [1] || '%,' And 是否开启审方 = 1 "
            strSQL = "Select Count(1) Rec " & vbNewLine & _
                     "From 处方审查参数 " & vbNewLine & _
                     "Where Nvl(服务对象, 0) = 0 And 是否开启审方 = 1 "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取处方审查工作站开启审方事务") ', lngClinicID)
            If rsTemp!Rec <= 0 Then
                '审方工作站未开启事务，不生成处方审查记录
                rsTemp.Close
                strSubmitNot = ""   '未提交的医嘱不返回给上层业务
                GoTo makOne
            End If
            rsTemp.Close
        Else
            '住院不检查是否开启审方事务
        End If
        
        typAudit = AutoAudit(Nothing, s_内部, strSubmitNot, lngClinicID, bytPatientMode, lngPatientID, lngBillID, strErr)
        '内部调用返回“未开启提醒”值
        If typAudit <> aar_未开启提醒 Then
            '无需生成审查记录
            strErr = "无需提交审查"
            strSubmitNot = ""       '未提交的医嘱不返回给上层业务
            MsgBox strErr
            GoTo makOne
        End If
        
        '检查数据
        If mrsAuditSend Is Nothing Then
            strErr = "审查数据未生成"
            strSubmitNot = ""       '未提交的医嘱不返回给上层业务
            MsgBox strErr
            GoTo makOne
        End If
        
        If mrsAuditSend.State <> adStateOpen Then mrsAuditSend.Open
        
        If mrsAuditSend.RecordCount <= 0 Then
            strErr = "审查数据无"
            strSubmitNot = ""       '未提交的医嘱不返回给上层业务
            MsgBox strErr
            GoTo makOne
        End If
        
        '按药房ID分别提交待审查记录
    '    strSQL = "Select a.执行科室id, a.Id, a.相关id, b.医嘱id, c.审查结果, " & vbNewLine & _
    '             "    Max(c.状态) Over(Partition By a.相关id) 状态, " & vbNewLine & _
    '             "    Count(a.相关id) Over(Partition By a.相关id) as 药嘱数, " & vbNewLine & _
    '             "    Sum(Decode(nvl(b.医嘱id,0), 0, 0, 1)) Over(Partition By a.相关id) as 提交数 " & vbNewLine & _
    '             "From 病人医嘱记录 A, 处方审查明细 B, 处方审查记录 C, Table(f_Num2list([1], ',')) D " & vbNewLine & _
    '             "Where a.相关id = d.Column_Value And a.Id = b.医嘱id(+) And b.审方id = c.Id(+) " & _
    '             "    And a.诊疗类别 In ('5', '6', '7') And b.最后提交(+) = 1 " & vbNewLine & _
    '             "Order By a.执行科室id, a.相关id, a.ID "
        
        strSQL = "Select a.执行科室id, a.Id " & vbNewLine & _
                 "From 病人医嘱记录 A, Table(f_Num2list([1], ',')) B " & vbNewLine & _
                 "Where a.相关id = b.Column_Value " & vbNewLine & _
                 "Order By a.执行科室ID, a.相关Id, a.ID "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取医嘱信息", strSubmitNot)
        With rsTemp
            If .EOF = False Then lngStoreID = !执行科室ID
            Do While .EOF = False
                '检查是否可以提交的医嘱ID
                    
                ''准备执行过程
                If lngStoreID = !执行科室ID Then
                    strIDs = strIDs & zlstr.FormatString("[1],", !ID)
                Else
makSQL:
                    '按药房ID分别提交审查记录
                    lngAuditID = zlDatabase.GetNextId("处方审查记录")
                    If lngAuditID <= 0 Then
                        strErr = "处方审查记录的序列对象不存在，无法提交待审查记录"
                        strSubmitNot = ""
                        MsgBox strErr, vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If strIDs <> "" Then strIDs = Left(strIDs, Len(strIDs) - 1)
                    
                    '待药师审查医嘱明细
                    strSQL = zlstr.FormatString("ZL_处方审查_INSERT( [1],[2],[3],[4],[5],'[6]',[7],'[8]' )", _
                                lngAuditID, _
                                lngPatientID, _
                                IIf(mblnOutPatient = False, "Null", lngBillID), _
                                IIf(mblnOutPatient = False, lngBillID, "Null"), _
                                lngClinicID, _
                                UserInfo.姓名, _
                                lngStoreID, _
                                strIDs)
                    
                    'SQL加入集合对象
                    AddArray colSQL, strSQL
                    
                    '自动审查结果
                    ''审查结果无医嘱ID，按药房ID分别提交
                    strTmp = ""
                    mrsAuditSend.Filter = "(审查结果=2 and 医嘱ID<=0) Or (审查结果=1 and 医嘱ID<=0)"
                    'If mrsAuditSend.RecordCount > 0 Then mrsAuditSend.MoveFirst
                    Do While mrsAuditSend.EOF = False
                        strTmp = strTmp & zlstr.FormatString("[1],0|", "" & mrsAuditSend!审查结果 & "_" & mrsAuditSend!审查项目ID)
                        mrsAuditSend.MoveNext
                    Loop
                    If strTmp <> "" Then
                        strTmp = Left(strTmp, Len(strTmp) - 1)
                        strSQL = zlstr.FormatString("ZL_处方审查_AUTO( [1], '[2]' )", _
                                    lngAuditID, _
                                    strTmp)
                        AddArray colSQL, strSQL
                    End If
                    
                    ''审查结果有医嘱ID，只提交一次
                    strTmp = ""
                    arrID = Split(strIDs, ",")
                    For i = LBound(arrID) To UBound(arrID)
                        strTmp = strTmp & zlstr.FormatString("(审查结果=2 and 医嘱id=[1]) or ", Val(arrID(i)))
                    Next
                    If strTmp <> "" Then
                        mrsAuditSend.Filter = Left(strTmp, Len(strTmp) - 4)     '过滤出有不合格的医嘱ID
                        strTmp = ""
                        If mrsAuditSend.RecordCount > 0 Then mrsAuditSend.MoveFirst
                        Do While mrsAuditSend.EOF = False
                            strTmp = strTmp & zlstr.FormatString("2_[1],[2]|", mrsAuditSend!审查项目ID, mrsAuditSend!医嘱ID)
                            mrsAuditSend.MoveNext
                        Loop
                        If strTmp <> "" Then
                            strTmp = Left(strTmp, Len(strTmp) - 1)
                            strSQL = zlstr.FormatString("ZL_处方审查_AUTO( [1], '[2]' )", _
                                        lngAuditID, _
                                        strTmp)
                            AddArray colSQL, strSQL
                        End If
                    End If
                    
                    If .EOF = False Then
                        '重新存储当前记录的ID
                        strIDs = zlstr.FormatString("[1],", !ID)
                    End If
                End If
                
                If .EOF = False Then
                    lngStoreID = !执行科室ID
                    .MoveNext
                    If .EOF Then GoTo makSQL
                End If
                    
            Loop
        End With
        
        '执行存储过程
        Err = 0: On Error GoTo errHandle
        ExecuteProcedureArray colSQL, "提交待审查记录"
        
    End If
    
makOne:
    
    On Error GoTo hErr
    
    '处理已提交的医嘱是否允许发送
    If strCannotSend <> "" Then
        strSQL = "Select f_List2Str(Cast(Collect(Cast(a.相关id as Varchar2(20))) as t_StrList), ',') 相关id " & vbNewLine & _
                 "From 病人医嘱记录 A, 处方审查明细 B, " & vbNewLine & _
                 "     (Select /*+ RULE*/ b.审方id " & vbNewLine & _
                 "      From 病人医嘱记录 A, 处方审查明细 B, 处方审查记录 C, Table(f_Num2list([1], ',')) D " & vbNewLine & _
                 "      Where d.Column_Value = a.相关id And a.Id = b.医嘱id And b.最后提交 = 1 And b.审方id = c.Id) C " & vbNewLine & _
                 "Where a.Id = b.医嘱id And b.审方id = c.审方id And a.相关id Is Not Null "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取不能提交的相关ID", strCannotSend)
        If rsTemp.EOF = False Then
            strNotSendIDs = NVL(rsTemp!相关id)
        End If
    End If

    If intPara = 1 And mblnOutPatient Then
        '再将刚提交审核的相关ID返回
        If strSubmitNot <> "" Then
            If strNotSendIDs <> "" Then
                strNotSendIDs = strNotSendIDs & "," & strSubmitNot
            Else
                strNotSendIDs = strSubmitNot
            End If
        End If
    End If
    
    BuildData = True
    Exit Function
    
hErr:
    If zl9ComLib.ErrCenter() = 1 Then
        Resume
    Else
        strNotSendIDs = strCannotSend   '如果有异常，未提交的医嘱不提交审查（门诊、住院）
    End If
    Exit Function
    
errHandle:
    gcnOracle.RollbackTrans
    strErr = Err.Description
    Call zl9ComLib.ErrCenter
    strNotSendIDs = strCannotSend       '如果有异常，未提交的医嘱不提交审查（门诊、住院）
    BuildData = False
End Function

Private Function GetGoodID(ByVal lngAuditID As Long) As String
'功能：获取合格或免审的审方ID对应的医嘱ID
'参数：
'  lngAuditID：审方ID
'返回：医嘱ID字符串（格式：医嘱ID,...）

    Dim strSQL As String, strReturn As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select a.医嘱ID From 处方审查明细 A, 处方审查记录 B " & vbCr & _
             "Where a.审方ID = b.ID And (b.审查结果 = 1 or b.审查结果 is Null And Nvl(b.状态, 0) between 1 and 3) And a.审方ID = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取审方ID对应的医嘱ID", lngAuditID)
    Do While rsTemp.EOF = False
        strReturn = strReturn & zlstr.FormatString(",[1]", rsTemp!医嘱ID)
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    If strReturn <> "" Then strReturn = Mid(strReturn, 2)
    GetGoodID = strReturn
    
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Function AuditSend(ByVal lngClinicDeptID As Long, ByVal strSubmitID As String, _
    ByRef strSendID As String, _
    Optional ByRef strErr As String) As Boolean
'功能：检查门诊/住院发送的医嘱ID是否审查合格
'参数：
'  lngClinicDeptID：临床开嘱科室ID
'  strSubmitID：检查的给药途径医嘱ID（格式：给药途径医嘱ID,...）
'  strSendID：可发送的医嘱ID
'  strErr：失败的出错信息
'返回：True成功；False失败

    Dim strSQL As String, strReturn As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim blnTrans As Boolean
    
    If mblnInit = False Then
        strErr = "请先调用接口初始化的方法"
        AuditSend = False
        Exit Function
    End If
    
    On Error GoTo errHandle

    '检查审查状态
    strSQL = "Select Distinct b.Id, b.审查人, b.审查结果, b.状态, b.提交时间 " & vbNewLine & _
             "From 处方审查明细 A, 处方审查记录 B, 病人医嘱记录 C, Table(f_Num2list([1], ',')) D " & vbNewLine & _
             "Where a.审方id = b.Id And a.医嘱id = c.Id And c.相关Id = d.Column_Value " & vbNewLine & _
             "    And a.最后提交 = 1 And c.诊疗类别  in ('5','6','7') "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取审查记录的状态", strSubmitID)
    If rsTemp.RecordCount <= 0 Then
        rsTemp.Close
        strErr = "处方审查记录不存在"
        Exit Function
    End If
    
    '遍历审查ID
    Do While rsTemp.EOF = False
        If mblnOutPatient Then
            '门诊业务
            If Val(zlCommFun.NVL(rsTemp!审查结果)) = 0 And zlCommFun.NVL(rsTemp!状态) = 0 Then
                '未审查
                '对超时长未审查的记录状态调整（存储过程中判断是否超时）
                If blnTrans = False Then
                    gcnOracle.BeginTrans
                    blnTrans = True
                End If
                
                strSQL = zlstr.FormatString("ZL_处方审查_UPDATE(1, [1], [2])", lngClinicDeptID, rsTemp!ID)
                Call zlDatabase.ExecuteProcedure(strSQL, "调整处方审查记录状态")
            End If
        End If
            
        '已审查或免审
        strTemp = GetGoodID(rsTemp!ID)
        If strTemp <> "" Then
            strReturn = strReturn & strTemp & ","
        End If
        
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    If strReturn <> "" Then strReturn = Left(strReturn, Len(strReturn) - 1)
    
    If blnTrans Then gcnOracle.CommitTrans
    
    AuditSend = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
End Function

Public Function Auditting(ByVal strSubmitID As String, ByRef strAudittingID As String) As Boolean
'功能：检查医嘱是否正在审查中
'参数：
'  strSubmitID：检查的给药途径医嘱ID（格式：给药途径医嘱ID[,给药途径医嘱ID...]）
'  strAudittingID：正在审查的医嘱ID
'返回：True成功；False失败

    Dim strSQL As String, strReturn As String
    Dim rsTemp As ADODB.Recordset
    
    If mblnInit = False Then
        Auditting = False
        Exit Function
    End If
    
    If strSubmitID = "" Then Exit Function
    
    On Error GoTo errHandle
    
'    strSQL = "Select A2.医嘱id " & vbNewLine & _
'             "From 处方审查明细 A1, 处方审查明细 A2, 处方审查记录 B, Table(f_Num2list([1], ',')) C " & vbNewLine & _
'             "Where A1.审方id = b.Id And A1.医嘱id = c.Column_Value And b.Id = A2.审方id And A1.最后提交 = 1 " & vbNewLine & _
'             "  And (b.锁定用户 Is Not Null Or b.锁定用户 <> '') " & vbNewLine & _
'             "Order By a2.医嘱id "
    strSQL = "Select b.医嘱id " & vbNewLine & _
             "From 病人医嘱记录 A, 处方审查明细 B, 处方审查记录 C, Table(f_Num2list([1], ',')) D " & vbNewLine & _
             "Where a.相关ID = d.Column_Value And a.Id = b.医嘱ID And b.审方Id = c.ID And b.最后提交 = 1 " & vbNewLine & _
             "  And (c.锁定用户 Is Not Null Or c.锁定用户 <> '') " & vbNewLine & _
             "Order By b.医嘱id "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取正在审查的医嘱ID", strSubmitID)
    Do While rsTemp.EOF = False
        strReturn = strReturn & zlstr.FormatString("[1],", rsTemp!医嘱ID)
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    If strReturn <> "" Then strReturn = Left(strReturn, Len(strReturn) - 1)
    strAudittingID = strReturn
    
    Auditting = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Sub ShowResult(ByVal lngMedicalID As Long, ByVal frmOwner As Object)
'功能：查看某医嘱ID对应审方的审查结果
'参数：
'  lngSubmitID：给药途径医嘱ID
'  frmOwner：宿主窗体对象

    Dim frmResult As frmRAResult
    
    If mblnInit = False Then Exit Sub
    
    Set frmResult = New frmRAResult
    frmResult.ShowMe lngMedicalID, frmOwner
    Set frmResult = Nothing
End Sub

Public Function CancelData(ByVal strSubmitID As String, ByRef strLockID As String, Optional ByRef strErr As String) As Boolean
'功能：1.对未审查、未锁定审查的待审查数据进行撤销；2.对已审查，但未发送的待审查数据进行撤销
'参数：
'  strSubmitID：待撤销审查的给药途径医嘱ID（格式：给药途径医嘱ID,...）
'  strLockID：被锁定的医嘱ID
'  strErr：失败的出错信息
'返回：True成功；False失败

    Dim cmdSQL As New ADODB.Command
    
    If mblnInit = False Then
        strErr = "请先调用接口初始化的方法"
        CancelData = False
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    '过程有出参，因些未用公共方法调用存储过程
    With cmdSQL
        .ActiveConnection = gcnOracle
        .CommandText = "ZL_处方审查_CANCEL"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("P1", adVarChar, adParamInput, 4000, strSubmitID)
        .Parameters.Append .CreateParameter("P2", adVarChar, adParamOutput, 4000)
        .Execute
        strLockID = .Parameters("P2").Value
    End With
    
    CancelData = True
    Exit Function
    
errHandle:
    strErr = Err.Description
End Function

Public Function CloseWindows() As Boolean
'功能：关闭部件里所有的窗体
'返回：True成功；False失败

    Dim frmThis As Form
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    mblnInit = False
    gintHoursRecipe = 24
End Sub

Private Sub Class_Terminate()
    Call CloseWindows
    If Not gobjRecipeAuditEx Is Nothing Then
        gobjRecipeAuditEx.CloseWindows
        Set gobjRecipeAuditEx = Nothing
    End If
    If Not gobjPubAdvice Is Nothing Then Set gobjPubAdvice = Nothing
End Sub

Private Function IsOutpatientControl(ByVal strMedicalIDs As String) As Byte
'功能：检查门诊处方药嘱是否为“处方审查条件”的监控范围
'参数：
'  strMedicalIDs：给药途径药嘱ID字符串
'返回：0-未设置门诊处方审查条件，即无监管审查的处方；1-需要审查的处方；2-无需审查的处方；

    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If strMedicalIDs = "" Then Exit Function

    '首先查检“处方审查条件”是否有记录
    strSQL = "Select Count(1) Rec From 处方审查条件 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查处方审查条件的记录数")
    If rsTemp!Rec <= 0 Then
        rsTemp.Close
        Exit Function
    End If
    
    '1-临床科室
    strSQL = "Select Count(1) Rec " & vbNewLine & _
             "From 病人医嘱记录 A, Table(f_Num2list([1])) B, 处方审查条件 C " & vbNewLine & _
             "Where a.Id = b.Column_Value And a.开嘱科室id = c.科室id And c.类别 = [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "处方审查条件-临床科室", strMedicalIDs, 1)
    If rsTemp!Rec > 0 Then
        rsTemp.Close
        IsOutpatientControl = 1
        Exit Function
    End If
    
    '2-医生
    strSQL = "Select Count(1) Rec " & vbNewLine & _
             "From 病人医嘱记录 A, Table(f_Num2list([1])) B, 人员表 C, 处方审查条件 D " & vbNewLine & _
             "Where a.Id = b.Column_Value And a.开嘱医生 = c.姓名 And c.Id = d.医生id And d.类别 = [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "处方审查条件-医生", strMedicalIDs, 2)
    If rsTemp!Rec > 0 Then
        rsTemp.Close
        IsOutpatientControl = 1
        Exit Function
    End If
    
    '说明：病人诊断记录.医嘱id 是 病人医嘱记录.相关id
    '3-诊断
    strSQL = "Select Count(1) Rec" & vbNewLine & _
             "From (Select 1 " & vbNewLine & _
             "      From 病人诊断记录 A, Table(f_Num2list([1])) B, 处方审查条件 C" & vbNewLine & _
             "      Where a.医嘱id = b.Column_Value And a.诊断id = c.诊断id And c.类别 = [2] And Rownum < 2" & vbNewLine & _
             "      Union All" & vbNewLine & _
             "      Select 1" & vbNewLine & _
             "      From 病人诊断记录 A, 病人诊断医嘱 B, Table(f_Num2list([1])) C, 处方审查条件 D" & vbNewLine & _
             "      Where a.Id = b.诊断id And b.医嘱id = c.Column_Value And a.诊断id = d.诊断id And d.类别 = [2] And Rownum < 2" & vbNewLine & _
             ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "处方审查条件-诊断", strMedicalIDs, 3)
    If rsTemp!Rec > 0 Then
        rsTemp.Close
        IsOutpatientControl = 1
        Exit Function
    End If
    
    '4-疾病
    strSQL = "Select Count(1) Rec" & vbNewLine & _
             "From (Select 1 " & vbNewLine & _
             "      From 病人诊断记录 A, Table(f_Num2list([1])) B, 处方审查条件 C" & vbNewLine & _
             "      Where a.医嘱id = b.Column_Value And a.疾病id = c.疾病id And c.类别 = [2] And Rownum < 2" & vbNewLine & _
             "      Union All" & vbNewLine & _
             "      Select 1" & vbNewLine & _
             "      From 病人诊断记录 A, 病人诊断医嘱 B, Table(f_Num2list([1])) C, 处方审查条件 D" & vbNewLine & _
             "      Where a.Id = b.诊断id And b.医嘱id = c.Column_Value And a.疾病id = d.疾病id And d.类别 = [2] And Rownum < 2" & vbNewLine & _
             ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "处方审查条件-疾病", strMedicalIDs, 4)
    If rsTemp!Rec > 0 Then
        rsTemp.Close
        IsOutpatientControl = 1
        Exit Function
    End If
    
    '5-药品
    strSQL = "Select Count(1) Rec " & vbNewLine & _
             "From 病人医嘱记录 A, Table(f_Num2list([1])) B, 处方审查条件 C " & vbNewLine & _
             "Where a.相关Id = b.Column_Value And a.诊疗项目id = c.药名id And c.类别 = [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "处方审查条件-药名", strMedicalIDs, 5)
    If rsTemp!Rec > 0 Then
        rsTemp.Close
        IsOutpatientControl = 1
        Exit Function
    End If
    
    rsTemp.Close
    
    IsOutpatientControl = 2
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function
